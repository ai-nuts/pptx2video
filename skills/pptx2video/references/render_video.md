# `python -m pptx2video.render_video` internals and debugging

## Stage diagram

```
svg_final/*.svg ──Playwright/Chrome──▶ slide-01.png ... slide-NN.png
      │
      └─ fallback: PPTX ──libreoffice──▶ PDF ──pdftoppm──▶ slide-01.png ...
                                          │
                                          │ pair by script JSON order
                                          │ (fallback: manifest/sorted names)
                                          ▼
   audio/<id>.mp3 ─ffprobe─▶ duration ───┘
                                          │
   visual_cues.json ──────────────────────┤ optional box/cursor attention overlays
                                          │
                                          ▼
                              ffmpeg per-slide segment
                              (libx264 + aac, padded with 0.3s silence)
                                          │
                                          ▼
                        ffmpeg concat demuxer (stream copy)
                                          │
                                          ▼
                              <project>/exports/<name>.mp4
```

## Why these specific tool choices

- **Browser SVG rasterization first**: ppt-master authors and previews the deck
  as SVG, and `svg_final` contains the expanded icon/vector content used for
  export. Rendering those SVGs with Chrome avoids LibreOffice reflowing text or
  geometry into frames that do not match the deck the user inspected.
- **LibreOffice for PPTX→PDF as fallback**: still available via
  `--frame-source pptx` or when no SVG deck exists, but it is no longer the
  preferred path for ppt-master projects.
- **pdftoppm over `convert`**: ImageMagick's `convert` works but defaults to
  Ghostscript under the hood, which is roughly 3× slower per page and
  occasionally rasterizes embedded fonts as pixelated bitmaps at low DPI.
- **ffmpeg concat *demuxer* over concat *filter***: every per-slide segment
  is encoded with identical codecs and dimensions, so the demuxer can
  stream-copy them with no re-encoding pass. Concat filter would re-encode
  everything (slow, lossy).
- **imageio-ffmpeg before system ffmpeg**: when `imageio_ffmpeg` is installed,
  the renderer can use its portable modern binary instead of an outdated
  system ffmpeg. Set `PPTX2VIDEO_FFMPEG` / `PPTX2VIDEO_FFPROBE` to override it.
  The older `PAPER2VIDEO_*` names remain compatibility aliases.
- **`-pix_fmt yuv420p`**: required for QuickTime/Safari/older Chrome
  playback. Without it you get a video that plays everywhere except the
  most popular consumer environment.

## Per-slide segment recipe

```
ffmpeg -y \
  -loop 1 -framerate 30 -i slide-05.png \
  -i 05-results.mp3 \
  -vf "scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2:color=black,format=yuv420p" \
  -af "aresample=44100,aformat=channel_layouts=stereo:sample_rates=44100,apad" \
  -c:v libx264 -preset medium -crf 20 \
  -c:a aac -b:a 192k \
  -strict -2 \
  -pix_fmt yuv420p -r 30 -t <audio_dur + pad> \
  -movflags +faststart \
  seg_0005.mp4
```

Notes:
- `-loop 1` on the image keeps the same frame on screen for the whole segment.
- Bare `apad` pads silence onto the end of the audio so the cut isn't
  jarring. We avoid `apad=pad_dur=N` because older ffmpeg builds don't support
  `pad_dur`; the segment-level `-t <audio_dur + pad>` still clips the padded
  stream to the desired length. We don't use `-shortest` because that ends the
  segment when audio ends, which can clip the final word in some MP3 encodings.
- `-strict -2` keeps older ffmpeg builds working with their experimental AAC
  encoder. Modern ffmpeg accepts it harmlessly.
- `-t` pins total duration so an over-padded audio doesn't extend the
  segment past the user's pad-tail expectation.

## Concat list format

The list file given to ffmpeg's concat demuxer:

```
file '/abs/path/to/seg_0001.mp4'
file '/abs/path/to/seg_0002.mp4'
file '/abs/path/to/seg_0003.mp4'
```

Paths are absolute and quoted (via `shlex.quote`) so spaces/Unicode in the
project path don't break the parser. The demuxer requires consistent codecs
and resolution. That invariant is why we encode each segment with the same
parameters.

## Audio ordering

Call `python -m pptx2video.render_video` with `--script-json <script.json>`
whenever possible.
The compositor then orders `audio/<id>.mp3` by the script's `sections` array,
which is required for paper2assets ids such as `problem`, `method`, and
`key-result`.

If `--script-json` is omitted, the script auto-detects
`<audio-dir>/script.json`, then `<project>/narration.json`, then
`<audio-dir>/manifest.json`. Only when none of those exist does it fall back to
alphabetical `*.mp3` order, which is safe for numeric ppt-master stems like
`01-intro.mp3` but unsafe for semantic ids.

## Duration reports

`python -m pptx2video.render_video --target-minutes N` does not change narration length at render
time. It writes a final duration report after real audio durations are known:

```json
{
  "schema_version": "paper2video_duration_report.v1",
  "status": "within_tolerance",
  "target_seconds": 180.0,
  "actual_seconds": 198.2,
  "target_delta_seconds": 18.2,
  "slides": [
    {"index": 1, "audio": "01_intro.mp3", "audio_seconds": 18.4, "visual_cues": 2}
  ]
}
```

Shape the source narration before TTS to change its duration. A higher-level
orchestrator may provide a narration planner; `--target-minutes` itself only
writes the real post-audio duration check.

## Final package QA

After rendering, run `python -m pptx2video.check_video_package --strict`.
Prefer passing
`--frames-dir <video_out>/assets/slides/frames`, the directory written by
`python -m pptx2video.render_video --frames-out`, so QA checks the exact frames used in
the MP4.
If no frames directory is provided, the checker falls back to rendering the
PPTX through LibreOffice for legacy packages. It probes audio/MP4 streams,
verifies slide and script counts, checks PPTX text-box overflow/overlap risks,
flags undersized visuals, and runs pixel-level checks for blank/sparse rendered
slides.

For videos with attention overlays, also pass `--strict-attention`,
`--require-visual-cues`, `--require-cue-plan`, and `--require-word-timings`.
That gate fails when cue coverage is too low, cue chunks are skipped, timings
fall back to proportional estimates, accepted cues are below confidence
threshold, accepted targets point at low-value slide chrome such as captions
or headers, or the rendered `geometry_box` no longer matches the cue `box`.

## Attention overlays

`--attention-mode highlight --visual-cues visual_cues.json` adds
per-slide ffmpeg overlays after slide scaling/padding and before final
`format=yuv420p`. Coordinates in `visual_cues.json` are normalized to the final
frame size, so a cue works across 720p/1080p/4K renders.

The current renderer supports:

- `highlight`: applies `type: "highlight"` cues. Box cues render as the
  selected normalized region; point-only legacy cues render as the older soft
  dot fallback.
- `cursor`: mouse-pointer overlay centered at a normalized point.
- `both`: applies both `highlight` and `cursor` cues when the cue file contains
  both types. Use this only when cue targets are precise enough.

`--highlight-style` controls how highlight cues look:

| Style | Use |
|---|---|
| `box` | Low-opacity slate fill and border around the selected box |
| `cursor` | Mouse pointer only at the cue point |
| `box_cursor` | Box plus mouse pointer, useful for geometry review |
| `spotlight` | Feathered dim-out around the selected box |
| `spotlight_cursor` | Feathered dim-out plus mouse pointer |
| `laser` | Red laser-pointer dot only at the cue point |
| `box_laser` | Focus box plus red laser-pointer dot |
| `spotlight_laser` | Default. Feathered dim-out plus red laser-pointer dot |

The default `spotlight_laser` style is the production path because it combines
semantic box geometry with a tolerant spotlight and a presenter-like laser dot.
Spotlight styles generate one full-frame transparent alpha mask per cue, keep
the accepted box at original brightness, and dim the outside area with a
continuous feathered falloff. At 1080p the default feather is about 56 px. They
are visually more tolerant than a hard box, but can make long full-video
renders slower than plain `box`.

Cursor styles use a generated transparent pointer overlay. Within each slide,
the pointer eases from one cue point to the next shortly before the next cue
starts, so it behaves like a presenter moving a mouse rather than jumping
instantly between regions.

Laser styles use the same eased cue-to-cue movement, but render a small red
dot with a soft halo instead of the mouse pointer.

See [visual_cues.md](visual_cues.md) for the JSON shape.

## Common failures and fixes

### `[concat @ 0x...] DTS X < Y out of order`

Usually means one of the segments has a different stream layout. Re-check
that every segment was produced by `encode_segment` (not pre-existing files
from a previous run with different settings). The script wipes
`<project>/.video_work/` on each invocation precisely to avoid this. If
you passed `--keep-temp` from a previous run, clear it manually.

### Output MP4 plays video but no audio

Usually a missing AAC encoder. Run `ffmpeg -encoders | grep aac`. If
nothing prints, your ffmpeg was built without AAC support. Install a
mainstream package, not a stripped-down one.

### Output is tiny (e.g. 50KB) and `ffprobe` shows zero duration

Concat list points at empty/corrupt segments. Re-run with `--keep-temp` and
inspect `<project>/.video_work/segments/`. Each `seg_NNNN.mp4` should be
several MB and play standalone in `ffplay`/VLC.

### LibreOffice "source file could not be loaded"

`.pptx` path has unicode or relative components. The script resolves to
absolute paths but if you assemble the command yourself, prefer
`Path(...).resolve()` first.

### LibreOffice silently produces no PDF

A logged-in GUI session is holding a lock on `~/.config/libreoffice`. The
script sidesteps this by passing `-env:UserInstallation=file://<tmpdir>` so
each render gets a fresh user profile. If you see this happen, you're
running an older or locally modified `pptx2video.render_video` module.

## Performance notes

| Stage | Cost driver | Typical 20-slide deck |
|-------|-------------|----------------------|
| SVG → PNG | Browser screenshots at output resolution | 3–10 s |
| PPTX → PDF fallback | One LibreOffice startup | 5–15 s |
| PDF → PNG fallback | DPI × page count | 5–10 s @ 150 DPI |
| Encode segments | `crf 20 medium` × audio length | ~0.5× real-time × total audio |
| Concat | Stream copy | <2 s |

A 5-minute narrated deck on a modern laptop takes about 3 minutes to render
end-to-end. Most of that time is in the per-segment x264 pass. Drop to
`-preset veryfast` (edit `encode_segment`) if you want to halve render time
at the cost of ~15% larger output files.

## Extending

If you want to add slide-to-slide crossfade transitions, encode each
segment with a 0.5 s overlap region and use ffmpeg's `xfade` filter at
concat time instead of the demuxer. That requires a re-encoding concat
pass (slow), which is why the default keeps hard cuts. Ship the simple
version first and only add transitions when a user asks.
