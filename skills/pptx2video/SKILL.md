---
name: pptx2video
description: Render any local PowerPoint PPTX into a fresh narrated video bundle through the public cross-platform pptx2video CLI, with Animation-Pane-authoritative object animations, unified Notes and Alt Text narration blocks, optional user script overrides, spotlight cues, block-level TTS caching, bottom-band subtitles, editable source delivery, timeline metadata, and strict QA. Use when a user invokes /pptx2video, asks for PPTX-to-video conversion, edits a PPTX and wants a rerender, needs animation-order conflict handling, or needs the Author Notes and Alt Text protocol.
---

# Convert an edited PPTX to video

This is a standalone, general-purpose PowerPoint-to-video package. It has no
runtime dependency on `paper2video`, ResearchStudio-Reel, or `ppt-master`, and
the complete CLI and renderer live in the installable `pptx2video` Python
package. It is not limited to research papers or to a previously delivered
`video.pptx`. Use any local PPTX as the visual source. Prefer canonical Author
Notes for precise narration and marker timing. When canonical Notes are absent,
read the same block grammar from Shape Alt Text. Never reuse rendered bundle
artifacts. Reusing a pristine block-level TTS cache entry is safe only when its
complete synthesis identity matches.

## Entry point

Treat `/pptx2video <input.pptx>` as a direct invocation of this skill. Use the
installed public CLI for all ordinary rendering:

```bash
pptx2video render <input.pptx> <new-output-directory> --resolution 1080p
```

Add `--animation-pointer cursor|laser` to follow native animation targets; it is
independent of Spotlight and defaults to `none`.

Do not import or execute private Paper2Video or ResearchStudio scripts. If the
CLI is unavailable, install the runtime from the public repository before
continuing:

```bash
python -m pip install \
  'pptx2video[svg] @ git+https://github.com/ai-nuts/pptx2video.git'
python -m playwright install chromium
pptx2video doctor --svg
```

## Authoring contract

For precise control, write a Notes block whose handle matches a top-level shape
or group:

```text
## [latency-card] latency-card
A new latency card appears. [[Spotlight] It reports lower latency.]
```

Alt Text uses the same single-block grammar, including both leading `#`
characters. The semantic text after `]` is optional and is not spoken:

```text
## [latency-card] latency-card
A new latency card appears. [[Spotlight] It reports lower latency.]
```

The baseline hash, generated ordering provenance, and semantic profiles live
inside the PPTX shape OOXML. The visible semantic after `]` is the concise
profile. PPTX generation or protocol writeback also persists a detailed
profile. Conflict display reads those stored values and uses deterministic
fallbacks for legacy decks, so it does not call an LLM. Native effect, target,
trigger, and delay data stay in PowerPoint's `p:timing` tree. They do not appear
in Notes or Alt Text. Older `[handle]` plus `Script:`, `[ID] handle`, and verbose
`[Paper2Video]` blocks remain readable and migrate to this unified form on the
next writeback. Keep Author Notes clean. Read
[references/editable_pptx.md](references/editable_pptx.md) for the
hidden provenance contract.

The stable handle resolves the target directly. By default, successfully
resolved canonical Notes blocks keep their current relative order as anchors.
Only narrated targets missing from Notes are inserted around those anchors by
current canvas geometry. The `--narration-order-policy author-notes` option is
retained as an explicit compatibility alias and produces the same
anchor-plus-insertion narration sequence. Animation Pane order independently
controls native animation row order and local effect timing. In the rendered
video, each Pane
phase is placed in output-only pre-roll immediately before its narrated block,
and the next phase waits for the preceding block to finish speaking. The only
canonical script marker is `Spotlight`; its
position determines an Edge word-boundary time or span. Legacy entrance markers
remain readable for migration, but they never replace, create, or reorder a
native animation.

Every delivered PPTX defaults to human-editable numbered click groups on every
slide. Make each non-`With Previous` Pane row the `On Click` leader of its own
top-level `mainSeq` group. Keep contiguous `With Previous` rows with that
leader so they share its canvas badge. PowerPoint derives the badge from this
topology; it is not a stored animation ID. Accept the rewrite only when effect
targets, presets, durations, delays, Pane start/end times, row order, and
simultaneous groups remain identical. Fail closed when the schedule cannot be
preserved. Use `--click-group-policy preserve` only when the source's automatic
click-group topology must remain unchanged.

For a PPT-native workflow without canonical Notes, put an explicit script in
the shape's Alt Text using the same grammar, and use the native Animation Pane
for effects:

```text
## [latency-card] latency-card
A new latency card appears and reports lower latency.
```

Read [references/animations.md](references/animations.md) before changing
effect names or resolving a protocol conflict.

## Authority and conflict rules

Apply this precedence:

1. An explicitly selected user `script.json` owns narration for that render.
2. Canonical Author Notes own handles. For narration, compare Notes and Alt Text
   with the last system-synchronized script hash stored in shape OOXML. A change
   on only one surface wins; if both changed differently, Notes wins and the
   authority report records the conflict.
3. A canonical Alt Text block provides narration when Notes are absent or when
   its narration alone differs from the stored baseline.
   A new animated target with only plain pre-protocol Alt Text uses that text
   once as its initial script and is normalized to the unified block grammar.
4. Narration order defaults to incremental `geometry` insertion. Every
   successfully resolved canonical Notes block is an anchor and keeps its
   relative Notes order. Only narrated targets absent from Notes are inserted
   around those anchors. At least 35 percent vertical overlap forms horizontal
   bands read top-to-bottom. Within each band, at least 50 percent horizontal
   overlap forms vertical lanes read left-to-right, and every lane is completed
   top-to-bottom before advancing right. Moving an existing anchored shape does
   not reorder established narration; move its Notes block when that order must
   change. Silent targets without geometry remain deterministic at the end.
   The explicit `author-notes` policy remains available for compatibility and
   uses the same Notes-anchor plus geometry-insertion sequence. Both policies
   fail closed when a newly inserted narrated block has no usable top-level
   geometry.
   After resolving Notes, Alt Text, an explicit `script.json`, or requested
   changed-element regeneration, but before TTS or output-bundle creation,
   compute two deterministic entrance sequences: a pure full-slide band/lane
   geometry order and native Animation Pane order. Number every eligible
   element again in each sequence; do not present the Notes-anchor narration
   sequence as geometry order. A conflict exists only when a relative inversion
   involves a newly inserted, Notes-missing narrated entrance target. This
   candidate filter does not geometrically reorder established Notes anchors.
   Ignore silent decoration, emphasis-only targets, and simultaneous groups.
   Emit `animation_order_authority.json` containing only slides with such
   conflicts, including exact Pane-row move guidance. Slides without conflicts
   remain silent and require no decision.
5. Animation Pane is the sole authority for ordinary native animation type,
   row order, `On Click`, `With Previous`, `After Previous`, duration, and
   delay. The video scheduler preserves those local relationships while adding
   narration release gates between Pane phases. A script can supplement the
   native effects only with `Spotlight`.
6. Shape OOXML stores `orderSource` and canonical `orderIndex` beside the script
   hash. A no-conflict render writes `author_notes` for established Notes
   anchors and `geometry` for newly inserted blocks, then rewrites contiguous
   indices for the effective sequence. A confirmed Pane choice writes
   `animation_pane` for that selected conflict slide's effective entrance
   sequence.
7. The PowerPoint canvas owns all visible pixels and geometry.

All 22 registered native entrance presets have explicit MP4 strategies. Complex
presets use named deterministic approximations and never silently become Fade.
Unknown entrance preset tuples and native animation classes outside the
supported entrance/emphasis contract fail explicitly.

When Notes and Alt Text both changed differently, Notes wins and the delivered
PPTX Alt Text is refreshed. When a legacy script entrance marker conflicts with
a native row, the native row wins and the conflict report records the ignored
marker. Without a native row, a legacy entrance marker is ignored rather than
inventing an animation. If a shape was deleted, remove its stale Notes block
and record it in `removed_stale_notes`; never rebind it by position or by a
remaining Animation Pane row. Ambiguous live handles still fail closed.

## One-command render

Install the public runtime and verify both core and SVG dependencies before
rendering:

```bash
python -m pip install \
  'pptx2video[svg] @ git+https://github.com/ai-nuts/pptx2video.git'
python -m playwright install chromium
pptx2video doctor --svg
```

Then invoke the local CLI:

```bash
pptx2video render \
  path/to/edited.pptx \
  path/to/new_video_bundle \
  --resolution 1080p
```

`--narration-order-policy geometry` is the default incremental policy. It
preserves the current canonical Notes block sequence and uses geometry only to
place narrated targets missing from Notes. `--narration-order-policy
author-notes` remains an explicit compatibility selection.

The default `--click-group-policy normalize` processes every slide, not only
the first slide. It produces one numbered badge per sequential phase while
keeping `With Previous` rows in the same numbered group. `preserve` is the
explicit opt-out.

The default `--animation-order-policy auto` is interactive only when stdin is
a TTY. It prompts only for conflict slides, in ascending slide index, and waits
for each decision before showing the next conflict slide. For example, resolve
slide 2 before displaying slide 4. In non-interactive or API use, a detected
conflict writes the report and exits with code 3 before creating the output
directory.

When Codex or another skill presents that report in conversation, show exactly
one conflict slide at a time and wait for the user's choice before showing the
next. Use the report's deterministic text or ASCII canvas, draw each element's
OOXML bounding box rather than only its center, and show three vertically
stacked views that approximate the PPTX canvas's actual aspect ratio after
accounting for monospace character-cell proportions. Fit all dimensions to the
detected terminal width rather than clipping a canvas. Use identical outlines
in all three views. Put only the geometry rank number inside each box in
`Geometry Order`, only the Pane rank number inside the same box in `Animation
Pane Order`, and only its stable identity letter such as `A`, `B`, or `C` in
`Element Identity Map`. Do not use combined labels such as `G1/P2`. Identity
letters follow natural sort of stable `shape_id`, independently of both order
sequences, so they carry no ordering meaning. Give every canvas its own
key-to-semantic legend titled exactly `LEGEND`. Beside `Geometry Order`, map each geometry rank number
to that element's semantic description. Beside `Animation Pane Order`, map
each Pane rank number to that element's semantic description. Beside `Element
Identity Map`, map each identity letter to that element's semantic description,
so a user can state an arbitrary preference such as `C -> A -> D -> B`. On a
wide terminal, place each legend to the right of its corresponding canvas. On
a narrower terminal, place it immediately below that canvas before presenting
the next view. The same selected profile applies to all three legends: default
to `concise`, or use the stored `detailed` profile through
`--semantic-profile detailed`. Below 26 columns, omit the bbox plots but retain
three labeled textual views, each with its own numeric-rank or identity-letter
semantic mapping. Thin
one-line elements must remain visibly wide and thin. Do not use an LLM to infer
either sequence, the identity map, the layout, or the semantic text at
conflict-display time, and do not list all conflict slides at once. A Pane
choice applies
only to the selected conflict slide; it serializes that slide's eligible
narrated entrance blocks in Pane order and synchronizes delivered Notes and
OOXML provenance. Other conflict slides keep their own decisions, and
no-conflict slides retain Notes-anchor plus inserted-geometry provenance. If
the user chooses `reading-order` on any slide, stop and present the exact Pane
move guidance from the report; do not silently render a video whose schedule
differs from the delivered PPTX.

Parse a user-selected Identity order with deterministic `identity_order.v1`.
For a conflict page whose Identity Map is `A` through `G`, `CABGDEF` means
exactly `C -> A -> B -> G -> D -> E -> F`. Treat input case-insensitively and
canonicalize it to uppercase. Require a complete permutation of the current
page's displayed identities, rejecting missing, duplicate, unknown, empty, or
invalid labels. Compact undelimited input is valid while every label is one
letter. If a page extends past `Z` to `AA`, require separators such as
`A,B,...,AA`. In a TTY, accept the order directly for the displayed page. For
API or non-TTY use, accept `--animation-order-sequence CABGDEF` for one conflict
page, or repeat `--animation-order-sequence SLIDE=ORDER` for every conflict
page. If the sequence exactly matches the current Pane, record it as explicit
Pane confirmation and render. If it differs, persist its Identity and shape-ID
orders in the decision report and stop for a source-PPTX Pane revision. Never
render an order that disagrees with the delivered PPTX.

Read [references/cli.md](references/cli.md) for installation and arbitrary-path
usage. Add `--script-json path/to/edited-script.json` only when a user-edited
external script should override PPTX narration. This command must:

1. Normalize unified Alt Text blocks from authoritative Notes when present.
2. Generate handles and unified Alt Text blocks for new animation targets.
3. Backfill canonical Notes when the source has only Alt Text or native rows.
4. Extract narration using user script, Notes, then Alt Text precedence, and
   apply the selected narration-order policy independently of script authority.
5. Normalize numbered click groups across every slide after change detection
   and protocol writeback. Verify that the renderer's Pane clock is unchanged
   and write `assets/meta/reports/click_group_normalization.json`.
6. Materialize Edge TTS and word timings per narration block, reusing only exact
   cache hits. During output assembly, prepend the corresponding Animation Pane
   phase as silence before each narrated block, shift every Edge word boundary
   by the same amount, then concatenate the blocks into slide audio. Generate
   deterministic silent audio for a native-only silent slide. Preserve each
   logical section ID in metadata while resolving its portable physical MP3
   through `manifest.json`.
7. Build the narration-serialized schedule: a narrated entrance completes before
   its first spoken word, and the following Pane phase starts only after the
   preceding narrated block ends. Preserve native simultaneous groups.
8. Align every subtitle cue to its actual first and last Edge TTS word
   boundaries. Never use proportional timing in a final render.
9. Render PPTX pixels, animations, audio, spotlight, and bottom-band subtitles.
   Read [references/render_video.md](references/render_video.md) for frame-source,
   attention-overlay, ffmpeg, performance, and debugging details.
10. Write `timeline.json`, subtitle-alignment evidence, mapping reports, and
   strict QA evidence.

Do not pass `--prebuilt-audio-dir` for a final render. Do not pass `--no-qa`.

The persistent TTS cache is content addressed per block. Its invalidation key
is the SHA-256 of a canonical JSON identity containing normalized narration
SHA-256, voice, rate, pitch, provider, provider version, and cache schema
version. A timing-aware render accepts a hit only when audio, metadata, and
word-boundary evidence all match their hashes and cache key. Editing one block
therefore regenerates that block only; changing any synthesis setting or
provider version invalidates every affected key. Animation pre-roll and tail
padding happen after cache materialization and never contaminate the pristine
cache entry.

For a change-aware rerender, ordinary users choose whether to keep existing
narration or regenerate only changed elements. The latter requires a previous
PPTX baseline and an OpenAI API key:

```bash
pptx2video render edited.pptx new-bundle \
  --baseline-pptx previous-video.pptx \
  --narration-mode regenerate
```

Regenerated scripts are written back into both Author Notes and unified Alt
Text blocks in the delivered PPTX. `script.json` remains an Advanced override, not the
normal editing surface.

## Ordinary animated PPTX

If the deck does not yet contain canonical Notes and Alt Text, bootstrap a copy
from an existing narration script, then edit that copy:

```bash
pptx2video bootstrap path/to/animated-source.pptx \
  --script-json path/to/script.json \
  --output path/to/editable-video.pptx \
  --report path/to/bootstrap-report.json
```

## Completion gate

Require these deliverables:

```text
new_video_bundle/
  video.mp4
  video_no_subtitles.mp4
  video.pptx
  manifest.json
  assets/audio/
  assets/captions/
  assets/meta/timeline.json
  assets/meta/reports/author_notes_authority.json
  assets/meta/reports/animation_order_authority.json
  assets/meta/reports/click_group_normalization.json
  assets/meta/reports/subtitle_timing_alignment.json
  assets/meta/reports/video_qa_report.json
```

Confirm the subtitle timing report has `status: word_aligned`, then confirm
`video_qa_report.json` has `passed: true`, `error: 0`, and `warning: 0`.
Do not claim completion until the strict renderer exits 0.

## macOS local use

Read [references/macos.md](references/macos.md). Install the same Python package,
install Playwright Chromium, verify native and SVG dependencies with
`pptx2video doctor --svg`, and invoke the standalone CLI from Terminal or an
external launcher.
