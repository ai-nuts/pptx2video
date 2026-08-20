# Author Notes animations in pptx2video

`pptx2video` uses PowerPoint animation metadata as an authoring contract. It
renders the current PPTX itself and does not record a PowerPoint playback
window.

## Ownership and timing

| Layer | Owner | What it stores |
|---|---|---|
| Editable object animation | PowerPoint Animation Pane / PPTX OOXML | Shape target, native effect, pane order, duration, trigger, delay, grouping |
| Narration mapping | PowerPoint Author Notes and Shape Alt Text | Unified stable-handle block, narration, optional `Spotlight` marker |
| Video timing | `animation_manifest.json` | Stable handle, PPTX shape id, and Edge word-aligned start/duration |
| Render evidence | `animation_render_report.json` | Strategy, global MP4 time, layer bbox, pixel sample times |
| Subtitle timing | SRT/VTT | Caption cues on the same audio clock |

The Author Notes bridge supplies narration and the video-only attention cue.
Notes and Alt Text use the same block grammar:

```text
## [result-card] optional semantic
Narration with optional [[Spotlight] spoken words].
```

The renderer reconciles these blocks with Animation Pane targets, then writes
the same single-block grammar to Alt Text. The accepted script hash and
generated ordering provenance live in the shape's `p2v:scriptBaseline` OOXML
extension, while native animation
target, effect, trigger, delay, and grouping remain in PowerPoint's `p:timing`
tree. Subtitles, spotlight/laser cues, and animations therefore share the same
audio clock without exposing generated metadata in Alt Text or borrowing
timestamps from one another.

Spotlight has two marker forms:

| Form | Spoken/captioned text | Video duration |
|---|---|---|
| `[[Spotlight]]` | Text after the marker | Native emphasis duration or 2.4 s default |
| `[[Spotlight] spoken phrase]` | The enclosed phrase | First enclosed Edge word start through last enclosed Edge word end |

The spoken-span form is preferred when a human editor wants direct duration
control without editing JSON. Its enclosed phrase remains ordinary narration
and subtitle text. It is valid only for `Spotlight`; empty scopes, partial-word
boundaries, and sequence-gated scopes that end before their resolved start fail
closed. The manifest and cue plan record `duration_source: script_scope`, the
scope text, and its resolved word range. Point markers remain backward
compatible.

For an ordinary user-added animated target that has only plain pre-protocol Alt
Text, normalization promotes that sentence to the target's initial narration
and writes it back in the unified block grammar plus canonical Notes.

Native animation semantics come directly from Animation Pane row order,
`On Click`, `With Previous`, `After Previous`, duration, and delay. For video,
the renderer divides that Pane clock into phases ending at narrated entrance
groups. It places each phase in output-only audio pre-roll before that block's
first Edge word, then releases the next phase only after the preceding block's
last Edge word. Silent connectors and Pane delays remain inside the following
phase, and explicit `With Previous` groups remain simultaneous. Trailing
native-only effects use output-only audio post-roll.

Before the final protocol extraction, the editable rerender route defaults to
`--click-group-policy normalize`. It processes every slide and rewrites the
`mainSeq` so every non-`With Previous` row is the `On Click` leader of an
independent top-level group. Contiguous `With Previous` rows stay in the same
group and therefore share the leader's PowerPoint canvas badge. Wrapper delays
are rebased to the new group origin, while effect-local delays, targets,
presets, durations, and row order remain unchanged. The transform is accepted
only when `_native_effects()` reports identical Pane start/end times and
simultaneous groups before and after. Unsupported or overlapping schedules
fail closed. Use `--click-group-policy preserve` to opt out.

The badge number is not a persistent animation ID. It is inferred from the
one-based position of each top-level click group in `p:cTn[@nodeType="mainSeq"]`.
An autoplay `onBegin` group is why PowerPoint can show `0`. The normalization
report records every slide's derived badge-to-shape mapping.

By default, every successfully resolved canonical Notes block keeps its current
relative order as an anchor. Deterministic page geometry is used only for
narrated targets missing from Notes, not to reshuffle established Notes blocks
or to copy Animation Pane position. Substantial vertical overlap forms
horizontal bands read top-to-bottom. Within a band, substantial horizontal
overlap forms vertical lanes read left-to-right, and each lane is completed
top-to-bottom before moving right. Moving an established anchored object alone
does not reorder it; geometry is recomputed for targets being inserted.

The `--narration-order-policy author-notes` option remains an explicit
compatibility selection. It produces the same Notes-anchor plus geometry
insertion narration sequence. Both policies write the effective sequence back to
canonical Notes. Established anchors receive `author_notes` provenance and
inserted blocks receive `geometry` provenance. Both policies fail closed when a
newly inserted narrated block lacks usable top-level geometry.

Before rendering, reconciliation deterministically computes a pure full-slide
band/lane geometry sequence and the native Animation Pane sequence. It assigns
fresh ranks to every eligible element in both sequences; the Notes-anchor
narration sequence is not labeled as geometry order. Only a relative inversion
involving a newly inserted, Notes-missing narrated entrance target is a
conflict, so this comparison does not geometrically reshuffle established
Notes anchors. Silent decoration, emphasis-only targets, and simultaneous
groups are excluded. The authority report contains only conflict slides, so
no-conflict slides remain silent. A TTY presents three vertically stacked,
terminal-width-responsive, aspect-ratio-aware text/ASCII canvases one conflict
slide at a time in ascending slide index: numeric Geometry order, numeric
Animation Pane order, and an order-neutral letter Identity map. Every canvas
has its own semantic legend titled exactly `LEGEND` on the right when space permits, or immediately
below it on a narrower terminal. A
Pane choice applies only to that selected slide. A reading-order choice emits
exact Pane row move guidance and stops. Non-interactive use writes the report
and exits before rendering.

The deterministic `identity_order.v1` protocol accepts a complete Identity
permutation such as `CABGDEF` for the displayed conflict page. It maps those
letters to stable shape IDs and rejects missing, duplicate, unknown, empty, or
invalid labels. Compact input is used through `Z`; labels from `AA` onward must
be separated. A sequence equal to the current Pane confirms Pane authority.
Any different sequence is recorded and stops for a matching PPTX Pane revision.

Author Notes are authoritative for handles, narration text, and `Spotlight`
intent. The selected narration-order policy owns block sequence. Per-block
script authority is reconciled with Alt Text through the hidden baseline hash.
The Animation Pane is authoritative for ordinary animation target, type, row
order, trigger, duration, delay, and grouping.
Legacy script entrance markers are migration input only: a native row wins a
conflict, and a marker without a native row is ignored. Deleting a shape removes
its stale Notes block and narration; the block is never rebound by position or
by a remaining Pane row.

## Support matrix

The package recognizes 22 native entrance effects. Every recognized effect has
an explicit MP4 strategy:

| Effect key | PowerPoint name | PPTX preset | MP4 strategy | Default duration |
|---|---|---|---|---:|
| `appear` | `Appear` | `1 / 0` | `appear` | 0.12 s |
| `fade` | `Fade In` | `10 / 0` | `alpha_fade` | 0.48 s |
| `fly` | `Fly In` | `2 / 4` | `fly_from_left` | 0.56 s |
| `cut` | `Cut In` | `42 / 8` | `cut_instant` | 0.12 s |
| `zoom` | `Zoom In` | `23 / 0` | `zoom_in` | 0.48 s |
| `wipe` | `Wipe In` | `22 / 1` | `wipe_from_left` | 0.52 s |
| `split` | `Split In` | `16 / 21` | `split_center_approx` | 0.52 s |
| `blinds` | `Blinds In` | `3 / 10` | `blinds_vertical_approx` | 0.52 s |
| `checkerboard` | `Checkerboard In` | `5 / 6` | `checkerboard_tiles_approx` | 0.56 s |
| `dissolve` | `Dissolve In` | `9 / 0` | `dissolve_alpha_approx` | 0.48 s |
| `random_bars` | `Random Bars In` | `14 / 10` | `random_bars_deterministic_approx` | 0.56 s |
| `peek` | `Peek In` | `12 / 4` | `peek_from_left_approx` | 0.52 s |
| `wheel` | `Wheel In` | `21 / 0` | `wheel_circle_approx` | 0.56 s |
| `box` | `Box In` | `4 / 0` | `box_reveal` | 0.52 s |
| `circle` | `Circle In` | `6 / 0` | `circle_reveal` | 0.52 s |
| `diamond` | `Diamond In` | `8 / 0` | `diamond_reveal` | 0.52 s |
| `plus` | `Plus In` | `13 / 0` | `plus_reveal` | 0.52 s |
| `strips` | `Strips In` | `18 / 12` | `diagonal_strips_approx` | 0.56 s |
| `wedge` | `Wedge In` | `20 / 0` | `diagonal_wedge_approx` | 0.56 s |
| `stretch` | `Stretch In` | `17 / 0` | `stretch_horizontal` | 0.52 s |
| `expand` | `Expand In` | `50 / 0` | `expand_from_center` | 0.52 s |
| `swivel` | `Swivel In` | `19 / 0` | `swivel_horizontal_approx` | 0.56 s |

The manifest and render report preserve the explicit strategy name. Complex
approximations use an `_approx` strategy identifier in machine-readable
evidence, so they are never presented as exact PowerPoint pixel playback and
never silently downgraded to Fade. An unknown entrance preset tuple fails during
extraction. A native animation class outside the supported entrance/emphasis
contract also fails explicitly rather than being omitted.
PowerPoint page transitions are a separate slide-level layer and are not part
of the object-animation manifest.

## Build and render

After a user changes the deck, invoke the installed CLI with a fresh output
directory:

```bash
pptx2video render edited.pptx new-video-bundle \
  --narration-order-policy geometry
```

The default is incremental `geometry` insertion around established Notes
anchors. `author-notes` remains an explicit compatibility selection and
produces the same anchor-plus-insertion sequence.

The renderer creates cumulative PPTX reveal states with LibreOffice and derives
each animation layer from adjacent pixel states. Text, color, position, image,
style, addition, and deletion edits therefore come from the current deck, not
from an earlier SVG export. The manifest records the PPTX SHA-256; rendering
and strict QA fail if the deck changes after the mapping is built.

The CLI builds the manifest, renders the video, and runs strict QA as one
transaction. It succeeds only when exact slide/order/locator/name coverage,
Edge timing provenance, strategy mapping, valid layer bboxes, transition pixel
changes, and PPTX source hashes all pass with zero errors and zero warnings.

## Is an upstream slide generator required?

No. Any tool may create the initial PPTX. After a native PPTX exists, a user can
edit it and run `pptx2video render` locally with no LLM or upstream generator.
See `editable_pptx.md` for the exact mutation contract.
