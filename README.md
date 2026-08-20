# pptx2video

`pptx2video` is a self-contained Python package for turning an edited `.pptx`
file into a narrated MP4 bundle. It reads visible slide content, native
Animation Pane timing, Author Notes, Shape Alt Text, and hidden OOXML
provenance directly from the presentation.

The package has no runtime dependency on `ResearchStudio-Reel`, `paper2video`,
or `ppt-master`. The complete renderer, protocol parser, TTS pipeline,
subtitle alignment, timeline builder, and strict QA gate live under
`src/pptx2video/`.

## Install the skill

Install the agent skill with the same cross-agent installer used by
`ppt-master`:

```bash
npx skills add ai-nuts/pptx2video --skill pptx2video
```

After installation, invoke it directly from a supporting agent host:

```text
/pptx2video /absolute/path/to/edited.pptx
```

The skill delegates rendering to the public `pptx2video` CLI. It does not call
private ResearchStudio or Paper2Video scripts.

## Install the CLI runtime

Python 3.11 or newer is required:

```bash
python -m pip install \
  'pptx2video[svg] @ git+https://github.com/ai-nuts/pptx2video.git'
python -m playwright install chromium
pptx2video doctor --svg
```

For changed-element narration regeneration:

```bash
python -m pip install \
  'pptx2video[svg,regenerate] @ git+https://github.com/ai-nuts/pptx2video.git'
```

Native dependencies are FFmpeg, Poppler, and LibreOffice. `pptx2video doctor`
reports the lightweight core status without probing Playwright. The explicit
`pptx2video doctor --svg` check also requires the Playwright module and an
executable system Chrome/Chromium-family browser or Playwright's installed
Chromium. It validates browser availability without launching a GUI.

## Render

Always choose a new output directory:

```bash
pptx2video render edited.pptx new-video-bundle --resolution 1080p
```

The command writes `video.mp4`, `video_no_subtitles.mp4`, an editable
`video.pptx`, exact Edge word-boundary captions, timeline metadata, protocol
reports, and a strict QA report. A successful final render requires zero QA
errors and zero warnings.

Animation Pane remains authoritative for native effect order and timing.
Geometry is used to place newly inserted narrated targets around established
Notes anchors. When those authorities conflict, the CLI emits deterministic
Geometry, Animation Pane, and Identity Map views and waits for a page-scoped
decision. A custom identity sequence that differs from the source Pane is
recorded and stops before rendering, so the delivered PPTX and MP4 cannot
diverge.

## Repository layout

```text
pptx2video/
  pyproject.toml
  src/pptx2video/        # complete standalone runtime and CLI
  skills/pptx2video/     # directly installable agent skill and references
```

Machine-readable schema names beginning with `paper2video_` and the existing
OOXML extension URIs are retained for backward compatibility with previously
authored decks. They are serialized compatibility identifiers, not runtime
dependencies.

See `skills/pptx2video/SKILL.md` for the agent workflow and authoring contract.
