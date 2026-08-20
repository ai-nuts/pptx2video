# Installable pptx2video CLI

The CLI is part of the standalone `pptx2video` Python package. After
installation it does not depend on a ResearchStudio, Paper2Video, or ppt-master
checkout. User PPTX and output paths are always command arguments.

## Install

Install the agent skill first:

```bash
npx skills add ai-nuts/pptx2video --skill pptx2video
```

Then install the public CLI runtime. Python 3.11 or newer is required:

```bash
python -m pip install \
  'pptx2video[svg] @ git+https://github.com/ai-nuts/pptx2video.git'
```

Install Playwright's Chromium after the runtime. The Python extra
installs the Playwright module, while this separate command installs its
browser executable:

```bash
python -m playwright install chromium
```

Maintainers can install a locally built wheel without preserving the
source-tree location:

```bash
python -m build
python -m pip install './dist/pptx2video-0.5.0-py3-none-any.whl[svg]'
```

Install the optional OpenAI dependency only when changed-element narration
regeneration is required:

```bash
python -m pip install \
  'pptx2video[svg,regenerate] @ git+https://github.com/ai-nuts/pptx2video.git'
```

Native dependencies remain platform-managed:

```text
FFmpeg with the required filters and encoders
FFprobe
LibreOffice
Poppler (pdftoppm)
Chrome/Chromium-family browser for SVG rendering
```

Use Homebrew on macOS, `apt-get` or `dnf` on Linux, and `winget` or Chocolatey
on Windows. Edge TTS requires network access while rendering. OpenAI access is
needed only when the user explicitly selects changed-element narration
regeneration. The package and CLI layout is cross-platform. Linux is currently
the strict-QA reference environment, macOS has a native local workflow, and
Windows remains preview support until it passes the same end-to-end gate.

## Check dependencies

```bash
pptx2video doctor --svg
```

The doctor checks the exact FFmpeg selected for the runtime, including the
required `ass`, animation, and audio filters plus `libx264`, `aac`, and
`libmp3lame` encoders. Rendering, subtitle burn-in, and audio processing are
pinned to that checked FFmpeg pair for the CLI process. With `--svg`, it also
checks that the Playwright module is importable and that either a system
Chrome/Chromium-family executable or Playwright's bundled Chromium executable
is available. It verifies paths without launching a browser GUI. Plain
`pptx2video doctor` remains a lightweight core-only check.

## Render

```bash
pptx2video render \
  "/Users/me/Documents/My edited deck.pptx" \
  "/Users/me/Movies/my-new-video-bundle" \
  --resolution 1080p
```

The PPTX can be anywhere, and paths with spaces are supported. Relative shell
paths are resolved against the current terminal directory. The output bundle
must not already exist, which prevents reuse of old audio, video, or caches.

Every render defaults to numbered PowerPoint click phases across all slides:

```bash
pptx2video render edited.pptx new-bundle \
  --click-group-policy normalize

# Keep the source mainSeq topology unchanged.
pptx2video render edited.pptx new-bundle \
  --click-group-policy preserve
```

`normalize` makes each non-`With Previous` Pane row the `On Click` leader of a
new top-level group. Contiguous `With Previous` rows stay with that leader and
share its badge. Structural wrapper delays are rebased, and the command fails
closed unless parsed Pane start/end times and simultaneous groups remain
identical. The per-slide evidence is written to
`assets/meta/reports/click_group_normalization.json`. Canvas badges are derived
click indices, not persistent animation IDs.

Narration order defaults to incremental geometry insertion:

```bash
pptx2video render edited.pptx new-bundle \
  --narration-order-policy geometry

# Explicit compatibility selection for Notes-anchor ordering.
pptx2video render edited.pptx new-bundle \
  --narration-order-policy author-notes
```

`geometry` preserves every successfully resolved canonical Notes block as an
anchor in its current relative order. It applies deterministic band-and-lane
ordering only to narrated targets missing from Notes, then inserts those blocks
around the anchors. Vertical overlap of at least 35 percent forms horizontal
bands read top-to-bottom. Horizontal overlap of at least 50 percent inside a
band forms vertical lanes read left-to-right, and each lane is completed
top-to-bottom before moving right. Moving an existing anchored shape alone does
not reorder it; reorder its Notes block when the established spoken sequence
must change. `author-notes` is retained as an explicit compatibility selection
and produces the same anchor-plus-insertion narration sequence. Both policies
fail closed when a newly inserted narrated block has no usable top-level
geometry. Neither selection changes handles, scripts, or native Animation Pane
timing.

Before creating that directory, the CLI resolves final narration authority and
deterministically calculates both a pure full-slide band/lane geometry sequence
and the native Animation Pane sequence. Every eligible element receives a fresh
rank in each sequence. This comparison view is separate from Notes-anchor
narration writeback, which still inserts only new blocks and does not reshuffle
established anchors. A conflict exists only when a relative inversion involves
a newly inserted, Notes-missing narrated entrance target. Silent decoration,
emphasis-only targets, and simultaneous groups are excluded. The report
contains only conflict slides; no-conflict slides produce no prompt or other UI.

In a TTY, `auto` shows one conflict slide at a time in ascending slide index and
waits for a choice before continuing. Its text/ASCII view uses three vertically
stacked canvases sized from the PPTX's actual slide width and height plus an
explicit monospace character-cell proportion. All three use identical scaled
OOXML bounding-box frames. `GEOMETRY READING ORDER` puts only geometry rank
numbers inside the frames. `ANIMATION PANE ORDER` puts only Pane rank numbers
inside the corresponding frames. `IDENTITY MAP (letters identify elements, not
order)` puts only stable letters such as `A`, `B`, and `C` inside them. Combined
labels such as `G1/P2` are not used, and the first two canvases never contain
letters.

Each canvas has its own keyed semantic legend titled exactly `LEGEND`. The Geometry legend maps each
geometry rank number to its element's semantic description. The Animation Pane
legend maps each Pane rank number to its element's semantic description. The
Identity legend maps each stable letter to its element's semantic description.
On a wide terminal, each legend appears to the right of its corresponding
canvas. On a narrower terminal, that legend moves immediately below its canvas,
before the next view. The identity map thereby provides neutral references for
a free-form choice such as `C -> A -> D -> B`. Letters are assigned by natural
sort of stable `shape_id`, independently of both Geometry and Animation Pane
order. After `Z`, labels continue as `AA`, `AB`, and so on.
`--semantic-profile concise` applies to all three legends by default. Select
the stored detailed descriptions with:

```bash
pptx2video render edited.pptx new-bundle \
  --semantic-profile detailed
```

Both profiles are stored in shape OOXML during PPTX generation or protocol
writeback. The conflict report exposes them as `semantic_concise` and
`semantic_detailed`. Display reads those stored values, with deterministic
legacy fallbacks, so no LLM is required for ordering, layout, identity mapping,
or semantic lookup. A proof line therefore appears as a wide, thin frame rather
than as a center-positioned block. The CLI detects the terminal column count,
shrinks width and height together, and honors `COLUMNS` in redirected or test
environments. A wide-terminal example is:

```text
Slide 2
GEOMETRY READING ORDER
+--------------------------+  LEGEND
| [1]              [2]     |  1: Any PPTX
+--------------------------+  2: Result

ANIMATION PANE ORDER
+--------------------------+  LEGEND
| [2]              [1]     |  1: Result
+--------------------------+  2: Any PPTX

IDENTITY MAP (letters identify elements, not order)
+--------------------------+  LEGEND
| [A]              [B]     |  A: Any PPTX
+--------------------------+  B: Result
```

When the canvas and its legend do not fit side by side, render the same legend
directly below that canvas. Do not move all three legends into one detached
section. Below 26 columns, omit the unreadable bbox plots and emit three labeled
textual views, each retaining its own numeric-rank or identity-letter semantic
mapping.

The deterministic `identity_order.v1` protocol accepts the Identity letters
directly:

```bash
# One conflict page: C, then A, then B, G, D, E, F.
pptx2video render edited.pptx new-bundle \
  --animation-order-sequence CABGDEF

# Multiple conflict pages: scope every order by slide and repeat the option.
pptx2video render edited.pptx new-bundle \
  --animation-order-sequence 2=CABGDEF \
  --animation-order-sequence 4=ABCDEFG
```

Input is case-insensitive and canonicalized to uppercase. Every displayed
Identity label must appear exactly once. Missing, duplicate, unknown, empty, or
invalid labels fail closed before bundle creation. Compact input is unambiguous
through `Z`; when labels include `AA` or later, separate every token, for
example `A,B,...,AA`. The parser maps labels to stable shape IDs by code only.
It does not use semantic text or an LLM.

An Identity sequence equal to the current Pane order is explicit Pane
confirmation and may render. Any different sequence is recorded in the
decision report and exits before rendering so the source Animation Pane can be
updated. This keeps the delivered PPTX and MP4 behavior consistent.

The CLI resolves slide 2 before it displays slide 4. In an API or other
non-interactive run, it writes
`<output>.animation-order-authority.json`, exits with code 3, and does not create
the output directory. A custom report path must remain outside the fresh output
directory:

```bash
# Confirm that the current Animation Pane is intentional.
pptx2video render edited.pptx new-bundle \
  --animation-order-policy animation-pane

# Select reading order and receive exact Pane move guidance without rendering.
pptx2video render edited.pptx new-bundle \
  --animation-order-policy reading-order
```

When a real conflict exists, an `animation-pane` choice applies only to the
selected conflict slide. It makes that slide's eligible narrated entrance
blocks follow Pane order and writes the effective order back to delivered Notes
and OOXML provenance. Other conflict slides keep their independently selected
policy. The ordinary no-conflict default retains `author_notes` provenance for
established anchors and `geometry` provenance for inserted blocks.

`reading-order` deliberately outputs the exact predecessor/successor Pane move
guidance for that slide and stops before rendering. Change the reported rows in
PowerPoint and rerun with `auto`; automatically changing only the MP4 schedule
would make the video disagree with the delivered PPTX's native timing tree.

The default narration text priority is canonical Author Notes, then explicit
Shape Alt Text blocks using the same `## [handle] optional semantic` grammar.
This text authority is independent from the default incremental insertion
policy. Legacy `Script:` fields remain readable and migrate on writeback. Native-only
silent slides are valid. To keep the existing narration while recording a
baseline comparison:

```bash
pptx2video render edited.pptx new-bundle \
  --baseline-pptx previous-video.pptx
```

To ask OpenAI to regenerate only added or modified elements:

```bash
export OPENAI_API_KEY="..."
pptx2video render edited.pptx new-bundle \
  --baseline-pptx previous-video.pptx \
  --narration-mode regenerate
```

The default model is `gpt-5.6-sol`; override it with
`--regeneration-model` when required. To use an edited `script.json` instead:

```bash
pptx2video render edited.pptx new-bundle \
  --script-json edited-script.json
```

A standard section `text` replaces the whole slide narration. For precise
per-element timing, use handle-addressed elements:

```json
{
  "sections": [
    {
      "id": "slide-one",
      "elements": [
        {
          "handle": "latency-card",
          "script": "The card appears. [[Spotlight]] Notice the lower latency."
        }
      ]
    }
  ]
}
```

The CLI succeeds only after the renderer exits successfully and
`video_qa_report.json` reports `passed: true`, `error: 0`, and `warning: 0`.
The delivered `video.pptx` contains synchronized Author Notes and unified Alt
Text blocks, including automatically registered new animation targets.

Fresh bundles may reuse pristine block-level TTS artifacts when their complete
content-addressed synthesis identity matches. The platform user cache is used by
default. Choose an explicit cache location or disable reuse with:

```bash
pptx2video render edited.pptx new-bundle \
  --tts-cache-dir /path/to/pptx2video-tts-cache

pptx2video render edited.pptx another-new-bundle --no-tts-cache
```

## Bootstrap an ordinary animated PPTX

```bash
pptx2video bootstrap ordinary.pptx \
  --script-json narration.json \
  --output editable-video.pptx
```
