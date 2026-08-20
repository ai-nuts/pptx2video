# Editable PPTX local rerender contract

The editable route uses the delivered PowerPoint as the local authoring file.
The canvas owns visible pixels. Canonical Author Notes own handles. Narration
and marker positions are reconciled against the last system-synchronized script
hash, so a user edit made only in Notes or only in Alt Text is retained. Without
that baseline, Notes remain authoritative except for the documented plain Alt
Text migration. Narration order defaults to preserving all successfully
resolved canonical Notes blocks as anchors and inserting only Notes-missing
targets by deterministic horizontal-band and vertical-lane geometry. The
`--narration-order-policy author-notes` option remains available as an explicit
compatibility selection.
Animation Pane is authoritative for native effect type, row order, trigger,
duration, delay, and grouping. Rerendering requires no LLM or ppt-master
checkout.

Delivered editable decks default to numbered click phases on every slide.
After user-change detection and narration writeback, the renderer splits each
non-`With Previous` Pane row into its own top-level `On Click` group and keeps
contiguous `With Previous` rows with their leader. It rebases structural
wrapper delays and requires the parsed Pane start/end schedule to remain
identical. The report is
`assets/meta/reports/click_group_normalization.json`. Select
`--click-group-policy preserve` only to retain the source `mainSeq` topology.
PowerPoint derives canvas badge numbers from these click groups; the badges are
not stored shape or animation IDs.

For compatibility with ordinary PowerPoint editing, a new animated target may
start with one plain Alt Text sentence instead of protocol metadata. When that
target has no Notes block or canonical Alt Text block, normalization promotes
the plain text to its initial narration, then writes the unified grammar to
Notes and Alt Text.

## One-time bootstrap

An ordinary native animated deck can be converted to the protocol without an
LLM or a package-specific script:

```bash
pptx2video bootstrap <animated-source.pptx> \
  --script-json <narration-script.json> \
  --output <editable-video.pptx> \
  --report <bootstrap-report.json>
```

The bootstrap assigns stable handles to Animation Pane targets, writes concise
canonical Notes, and writes matching Alt Text with the same grammar. It splits
each slide narration deterministically across targets in band-and-lane reading
order, validates the result, and preserves the original native timing
tree.

## One-command rerender

```bash
pptx2video render <edited.pptx> <output_bundle> \
  --ids-from-script <previous_bundle>/assets/audio/script.json
```

The default is `--narration-order-policy geometry`, meaning incremental
geometry insertion around the current canonical Notes sequence. The
`--narration-order-policy author-notes` option is retained for explicit
compatibility.

The command performs these deterministic stages:

1. Parse slides in `p:sldIdLst` presentation order.
2. Read native animation targets, row order, trigger, duration, delay, and
   grouping from each slide's `p:timing` tree.
3. Resolve Notes directly by stable shape handle. Remove a Notes block whose
   managed target was deleted; never bind it to another shape by position or
   by a remaining Animation Pane row.
4. Compare Notes and Alt Text with the stored baseline hash, accept the edited
   side, and refresh both surfaces from the accepted script.
5. Optionally apply an explicitly selected user-edited `script.json`.
6. Materialize each narration block through the content-addressed TTS cache.
   Add output-only Animation Pane pre-roll before each narrated block, shift its
   word boundaries by exactly that pre-roll, and concatenate the sequenced
   blocks into slide audio.
7. Build a word-aligned animation manifest from normalized `video.pptx`. Each
   narrated entrance phase completes before that block starts speaking, and the
   next phase waits for the preceding narration to end.
8. Render cumulative reveal states from that exact PPTX with LibreOffice.
9. Align each subtitle cue to the actual first and last Edge word boundary.
10. Encode animation, optional spotlight, audio, and bottom-band subtitles.
11. Write measured duration, subtitle alignment evidence, and the media timeline.
12. Run strict media, timeline, protocol, source-hash, subtitle, and animation QA.

`--prebuilt-audio-dir` is an offline/test option. Its MP3 names and
`word_timings.json` must exactly match the newly extracted Notes script or the
manifest stage fails.

## Author Notes and Alt Text

Keep the human-authored Notes block concise:

```text
## [result-card] Main result card
Accuracy rises by [[Spotlight] twelve points].
```

The renderer writes the same one-block grammar to the corresponding target's
Alt Text. Keep both leading `#` characters; semantic text is optional and is
not spoken:

```text
## [result-card] Main result card
Accuracy rises by [[Spotlight] twelve points].
```

Treat the optional semantic text in the visible header as the concise semantic
profile. PPTX generation and protocol writeback preserve two non-spoken values
per managed shape: `semantic_concise` and `semantic_detailed`. The concise form
is the default conflict legend. The detailed form gives enough context when a
short label is ambiguous. Conflict presentation reads these stored values and
does not invoke an LLM.

Authority and validation rules:

- Notes handles are authoritative. Exact handles target shapes without relying
  on Animation Pane position.
- By default, all successfully resolved canonical Notes blocks keep their
  relative order as anchors. Only narrated targets absent from Notes are
  ordered and inserted by geometry. At least 35 percent vertical overlap forms
  a horizontal band, and bands are read top-to-bottom. Inside a band, at least
  50 percent horizontal overlap forms a vertical lane; lanes are read
  left-to-right and each lane is completed top-to-bottom before advancing
  right. Moving an existing anchored object alone does not reorder it. Reorder
  its Notes block to change the established spoken sequence. A newly inserted
  narrated block without usable top-level geometry fails closed; silent targets
  without geometry remain stable at the end. The explicit `author-notes`
  compatibility policy produces the same anchor-plus-insertion narration
  sequence and applies the same fail-closed geometry requirement.
- After the final Notes, Alt Text, `script.json`, or regeneration authority is
  resolved in a temporary PPTX, reconciliation deterministically computes a
  pure full-slide band/lane geometry order and Animation Pane order. Every
  eligible element receives a fresh rank in both sequences. The comparison
  does not substitute the Notes-anchor narration sequence for geometry and does
  not reshuffle established anchors. A conflict exists only when a relative
  inversion involves a newly inserted, Notes-missing narrated entrance target.
  The report includes the inserted block's visible canvas text, narration,
  geometry rank, native row, Pane start, every crossed target, and exact
  predecessor/successor Pane rows. It contains only conflict slides. Silent
  decoration, emphasis-only targets, simultaneous groups, and no-conflict
  slides do not trigger a decision. Preflight stops before TTS and before
  creating the fresh bundle.
- Interactive CLI output renders three aspect-ratio-aware ASCII canvases
  stacked vertically: `GEOMETRY READING ORDER`, `ANIMATION PANE ORDER`, and
  `IDENTITY MAP (letters identify elements, not order)`. Their dimensions come
  from the PPTX's actual slide size and an explicit monospace character-cell
  proportion, then shrink together to the detected terminal width. All three
  use identical scaled OOXML bounding-box frames. The first two frames contain
  only their plain numeric ranks. The third contains only stable identity
  letters. Never put letters into the two order views or combine labels as
  `G1/P2`. Assign identity letters by natural sort of stable `shape_id`,
  independently of both order sequences, and continue after `Z` with `AA`,
  `AB`, and so on. Identity letters therefore have no ordering meaning. Attach
  a separate semantic legend titled exactly `LEGEND` to every view. The Geometry legend keys the
  selected descriptions by geometry rank, the Animation Pane legend keys them
  by Pane rank, and the Identity legend keys them by the order-neutral letters.
  On a wide terminal, place each legend to the right of its own canvas. On a
  narrower terminal, place each legend immediately below its own canvas before
  the next canvas. Apply the selected `concise` or `detailed` semantic profile
  consistently to all three legends. The Identity legend lets the user specify
  an arbitrary preference such as `C -> A -> D -> B`. Below 26 columns, omit
  all bbox plots but retain three labeled textual views, each with its own
  numeric-rank or identity-letter semantic mapping. No LLM participates
  in rank calculation, identity assignment, canvas layout, or semantic lookup.
  Prompt for conflict slides one at a time in ascending slide index, waiting for
  each choice before showing the next slide. Conversational agents must use the
  same one-page-at-a-time flow and must not display every conflict slide in one
  message. Non-interactive runs write the conflict-only report and exit without
  rendering.
- Custom input uses deterministic `identity_order.v1`. A compact permutation
  such as `CABGDEF` means `C`, then `A`, then `B`, `G`, `D`, `E`, and `F`.
  Validate it against the current page's complete Identity Map, rejecting
  missing, duplicate, unknown, empty, or invalid labels. Canonicalize case to
  uppercase. Compact form is allowed only while all labels are one character;
  require separators after `Z`. A sequence matching Pane confirms Pane
  authority. Any different sequence is persisted with its stable shape-ID order
  and stops for a PPTX Pane revision. Parsing and mapping are code-only.
- `--animation-order-policy animation-pane` records an explicit confirmation
  for the selected conflict slide only and continues to the next conflict page.
  It serializes that slide's eligible narrated entrance blocks in Pane order and
  writes the effective order to delivered Notes and OOXML provenance before
  TTS. Other conflict pages keep their own choices. A no-conflict default
  retains `author_notes` provenance for established Notes anchors and `geometry`
  for inserted blocks. `reading-order` outputs exact predecessor/successor Pane
  move guidance and stops. The renderer does not silently override native
  timing because that would make the video disagree with the delivered
  PowerPoint.
- Animation Pane is authoritative for ordinary animation type, row order,
  trigger, duration, delay, and grouping. The video scheduler preserves each
  Pane phase and simultaneous group, but inserts narration release gates so an
  element animates before its script and the next phase waits for that script.
  The only canonical script marker is `Spotlight`, which supplements the native
  schedule with a video attention cue.
- Legacy script entrance markers remain readable during migration. They cannot
  replace, create, or reorder native rows. A conflicting native row wins; an
  entrance marker without a native row is ignored and reported.
- The last system-synchronized script hash is stored only in shape OOXML. A
  Notes-only edit or Alt-Text-only edit wins. If both
  changed to the same value, accept it. If both changed differently, Notes wins
  and `author_notes_authority.json` reports `conflict: true` with both hashes.
- The authoritative provenance node is
  `ppt/slides/slideN.xml` →
  `p:cNvPr/a:extLst/a:ext/p2v:scriptBaseline`. Besides `sha256`, it stores
  `orderSource` and canonical `orderIndex`. Alt Text does not expose these
  generated fields. Default no-conflict renders record Notes-owned anchors as
  `author_notes` and newly inserted blocks as `geometry`, with contiguous
  `orderIndex` values for the effective sequence. A confirmed Pane choice
  records that selected conflict slide's effective entrance sequence as
  `animation_pane`.
- Semantic profiles use a separate sibling `a:ext` whose URI is
  `https://github.com/microsoft/ResearchStudio/paper2video/semantic-provenance/2026`.
  Its payload is
  `<p2vs:blockSemantics schemaVersion="1"><p2vs:concise>...</p2vs:concise><p2vs:detailed>...</p2vs:detailed></p2vs:blockSemantics>`.
  The report exposes the same values as `semantic_concise` and
  `semantic_detailed`; legacy `semantic` remains a concise alias. For a legacy
  deck without this extension, derive concise deterministically from the visible
  header semantic, first visible text, shape name, then handle. Derive detailed
  from the accepted narration, joined visible text, then concise. Persist the
  resolved profiles on the next protocol writeback so later conflict display
  remains LLM-free.
- For a legacy PPTX with no baseline, a plain Alt Text replacement on a
  Notes-owned animated shape is promoted once as the user-edited script. Other
  ambiguous legacy Notes/managed-Alt differences remain Notes-first and are
  reported rather than silently discarded.
- Marker position determines its Edge word-boundary start time. The preferred
  `[[Spotlight] spoken phrase]` form keeps the phrase in speech and subtitles,
  then uses the first and last enclosed Edge word boundaries as the exact cue
  interval. Legacy `[[Spotlight]]` remains a point marker with native or
  default duration.
- Native emphasis effects map to the deterministic `Spotlight` video cue.
- Unsupported Notes marker names, ambiguous handles, nested targets, and empty
  slide narration fail closed.
- Only top-level PowerPoint elements are supported as editable animation
  targets. Group related primitives first, then animate and identify the group.
- Older `[handle]` plus `Script:`, `[ID] handle`, and verbose `[Paper2Video]`
  blocks remain readable. Every writeback migrates Alt Text to exactly one
  `## [handle] optional semantic` block with its narration below. All generated
  details remain in OOXML or native `p:timing`.
- All 22 registered native entrance presets have explicit MP4 strategies. Some
  complex effects use a deterministic approximation whose strategy name ends in
  `_approx`; no recognized effect silently becomes Fade. Unsupported entrance
  tuples and native animation classes outside the supported entrance/emphasis
  contract fail explicitly.
- Final subtitle cues must use `edge_word_boundary` timing. A cue starts at its
  first spoken word and ends at its last spoken word. Punctuation attaches to
  those words. Missing or mismatched boundaries fail closed instead of falling
  back to character-proportional estimates.

## Add, delete, and modify

Modify an element:

1. Change text, image, color, style, size, or position on the PowerPoint canvas.
2. Edit the Notes transcript if the spoken narration should change.
3. Keep the Notes handle stable unless intentionally renaming it. If renamed,
   the next render refreshes the Alt Text handle and script.
4. Rerun the command. Both static and animated pixels come from the edited
   PPTX, so no SVG regeneration is needed.

Add an element:

1. Add a top-level shape or group.
2. Give it one Alt Text block beginning `## [new-result] optional semantic`.
3. Optionally give it a native entrance effect.
4. To choose an exact narration position, add `## [new-result] ...` in Notes at
   that position. To use automatic insertion, leave the target absent from
   Notes; normalization inserts its Alt Text narration by geometry and writes
   the resulting canonical Notes sequence.
5. Add `[[Spotlight]]` or `[[Spotlight] spoken words]` only if the video needs an
   attention cue. Choose ordinary animation in PowerPoint's Animation Pane.

Delete an element:

1. Delete the shape from the slide. PowerPoint removes its native animation.
2. Rerun. The renderer removes the matching stale Notes block automatically,
   omits its narration, and records the action in `removed_stale_notes`.
3. A new shape at the same position does not inherit the deleted block. Stale
   Notes are never rebound by position or by leftover Animation Pane order.

Reordering is also explicit. Reorder canonical Notes blocks to change the
spoken sequence of established elements. Canvas geometry determines only where
targets absent from Notes are inserted; moving an established anchored shape
alone does not reorder it. Reorder Animation Pane rows to change native playback
order and trigger relationships. When
preflight reports a conflict, make that preference explicit instead of relying
on the new animation row that PowerPoint commonly appends at the end.

## Block-level TTS cache

Each narration block is an independent cache unit. The cache key is SHA-256 of
a canonical JSON identity containing:

- cache schema version;
- SHA-256 of NFKC-normalized, whitespace-collapsed narration;
- Edge voice, rate, and pitch;
- provider name and provider version.

A word-timed render accepts a hit only if metadata, pristine MP3, and word
boundaries exist; the identity matches; the MP3 hash matches metadata; and the
timing record repeats both the cache key and audio hash. Editing one block
invalidates only that block. Changing voice, rate, pitch, provider, provider
version, or schema invalidates the affected identities. Blocks are sequenced and
concatenated after retrieval. Animation pre-roll and minimum-duration tail
padding are applied only to the copied slide output and never stored back into
the pristine cache. The audio manifest records `hit`, `miss`, `partial_hit`,
`disabled`, or `not_applicable` plus each block's cache identity.

## Reproducibility evidence

`animation_manifest.json` records:

- `source_kind: "pptx"`;
- the exact PPTX SHA-256;
- stable slide IDs, section IDs, shape IDs, Alt Text handles, native order,
  Animation-Pane-selected effect names, ignored legacy marker conflicts, and
  Edge-aligned Spotlight times.

`animation_render_report.json` repeats the source hash and records every layer
bbox and MP4 sample window. Strict QA verifies the delivered PPTX hash and
checks an early/late pixel pair for every mapped effect.

The automated regression suite edits a synthetic PowerPoint in sequence:

- red card to blue card, proving modification reaches encoded MP4 pixels;
- add a green card, proving manifest and video reveal count increase;
- delete the original card, proving its mapping and pixels disappear;
- stale an Alt Text handle, proving Notes wins and compact Alt Text is refreshed;
- omit Notes for a native target, proving silent native effects remain valid;
- add scripted targets absent from Notes across multiple bands and staggered
  vertical lanes, proving default geometry inserts them without reordering
  established Notes anchors;
- select `author-notes`, proving the explicit compatibility path preserves the
  same Notes-anchor sequence;
- use Alt Text scripts without Notes, proving spatial generation and writeback;
- add a legacy entrance marker with no native row, proving it is ignored;
- add a native emphasis plus `[[Spotlight]]`, proving local attention mapping.

Previously delivered decks using `[ID] result-card` or `[result-card]` plus
`Script:` remain readable for local rerenders. Bootstrap and all newly rendered
decks write `## [result-card] optional semantic` followed by narration in both
Alt Text and Notes.
