# Visual cues JSON - highlight boxes and cursor overlays

`python -m pptx2video.render_video` can burn simple attention cues into each
static slide segment.
Use this when a narrated video feels too inert but you do not want to animate
the slide deck itself.

## CLI

```bash
python -m pptx2video.render_video <project_path> \
  --pptx <project_path>/exports/<name>.pptx \
  --audio-dir <project_path>/audio \
  --script-json <project_path>/audio/script.json \
  --attention-mode highlight \
  --highlight-style spotlight_laser \
  --visual-cues <project_path>/visual_cues.json \
  --out <project_path>/exports/<name>.mp4
```

`--attention-mode` controls which cues are applied:

| Mode | Behavior |
|---|---|
| `none` | Ignore cues; use only for an approved no-highlight render |
| `highlight` | Apply only `type: "highlight"` cues; default final-delivery mode |
| `cursor` | Apply only `type: "cursor"` cues |
| `both` | Apply both cue types when the cue file contains both |

## Schema

Production highlight cues are box-first. Coordinates are normalized to the final video
frame:

- `[0, 0]` is top-left.
- `[1, 1]` is bottom-right.
- A highlight `box` is `[x, y, w, h]` and is the geometry actually rendered.
  The renderer expands it by about one border width, then draws a low-opacity
  slate fill and soft border around that target.
- A highlight should also include `point` as the center for compatibility and
  audit tooling. Point-only highlight cues remain valid as degraded/debug
  fallbacks, but strict QA expects boxes.
- Automatic cues should include `semantic_*` and `geometry_*` fields. The
  semantic fields explain what narration target was selected; the geometry
  fields explain which PPTX or PPTX-cluster box was used for the visible
  highlight. When PPTX geometry cannot be matched with enough confidence,
  `geometry_matched` is false and the cue falls back to the semantic box.
- Timing should come from `edge_word_alignment`, which aligns each narration
  chunk back to the real word-boundary timeline. `duration_proportional` is
  scaffolding/debug timing only and should fail strict final QA when word
  timings are required.
- A cursor `point` is `[x, y]`.

```json
{
  "schema_version": "paper2video_visual_cues.v3",
  "cue_shape": "semantic_box",
  "slides": [
    {
      "id": "07_fineweb_accuracy_lift",
      "cues": [
        {
          "start": 3.2,
          "end": 8.5,
          "type": "highlight",
          "box": [0.12, 0.28, 0.14, 0.21],
          "point": [0.19, 0.39],
          "target": "cue_s07_c1_accuracy_lift",
          "target_role": "result",
          "semantic_target": "svg:g:cues07-result-card",
          "semantic_source": "svg",
          "semantic_box": [0.11, 0.27, 0.16, 0.22],
          "geometry_target": "TextBox 18",
          "geometry_source": "pptx",
          "geometry_box": [0.12, 0.28, 0.14, 0.21],
          "geometry_matched": true,
          "geometry_match_score": 5.72,
          "confidence": 0.82,
          "color": "#64748B",
          "opacity": 0.18,
          "size": 56
        },
        {
          "start": 9.0,
          "duration": 4.0,
          "type": "cursor",
          "point": [0.52, 0.62],
          "color": "#ff6b00",
          "opacity": 0.95,
          "size": 32
        }
      ]
    }
  ]
}
```

## Slide matching

Prefer `id`, matching the narration/audio stem:

```json
{"id": "07_fineweb_accuracy_lift", "cues": []}
```

Use `index` only when ids are unavailable:

```json
{"index": 7, "cues": []}
```

## Timing

Cue times are relative to the start of that slide's own video segment, not the
global MP4 timeline. They should align with the narration text for that slide.

Accepted timing forms:

```json
{"start": 3.2, "end": 8.5}
{"at": 9.0, "duration": 4.0}
```

Times are clipped to the slide segment duration, including `--pad-tail`.

## Rendering

`highlight` cues are box-first, and the default presentation style is
`spotlight_laser`: a feathered spotlight around the accepted box plus a small
red laser-pointer dot at the cue center. Existing point-only cue files remain
valid and render as cursor/point fallbacks, but final strict attention QA
requires highlight boxes.

`python -m pptx2video.render_video --highlight-style` controls the presentation of accepted
highlight cues:

| Style | Behavior |
|---|---|
| `box` | Subtle filled frame around the selected box |
| `cursor` | Mouse pointer only at the cue point |
| `box_cursor` | Box plus mouse pointer for debugging or reviewer comparisons |
| `spotlight` | Feathered dim-out around the selected box |
| `spotlight_cursor` | Feathered dim-out plus mouse pointer |
| `laser` | Red laser-pointer dot only at the cue point |
| `box_laser` | Focus box plus red laser-pointer dot |
| `spotlight_laser` | Default delivery style: feathered dim-out plus red laser-pointer dot |

The spotlight styles generate a full-frame transparent alpha mask for each cue:
the accepted box remains at original brightness while the surrounding slide
fades out with a continuous feather. At 1080p the default feather is about
56 px. They are visually tolerant, but full-video encoding can be slower
because each cue adds an extra overlay mask.

Cursor styles render a generated transparent mouse pointer. The renderer keeps
the same cue point semantics, but eases the visible pointer between consecutive
cue points on a slide shortly before each next cue starts.

Laser styles use the same cue point semantics and eased movement, but render a
small red dot with a soft halo instead of the mouse pointer.
