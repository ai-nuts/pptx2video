# macOS local use

The standalone `pptx2video` package runs from Terminal after a native PPTX
exists. It does not require an LLM, `paper2video`, ResearchStudio-Reel, or
`ppt-master` for an ordinary render.

## Requirements

- Apple Silicon or Intel macOS with Homebrew
- Python 3.11 or newer
- LibreOffice for PPTX-to-PDF rendering
- Poppler for `pdftoppm`
- FFmpeg and FFprobe
- Playwright and Chromium for SVG slide rendering
- Network access for fresh Edge TTS

Install the native tools with Homebrew:

```bash
brew install ffmpeg poppler
brew install --cask libreoffice
```

Install the agent skill and public Python runtime:

```bash
npx skills add ai-nuts/pptx2video --skill pptx2video
python3 -m venv .venv
source .venv/bin/activate
python -m pip install \
  'pptx2video[svg] @ git+https://github.com/ai-nuts/pptx2video.git'
python -m playwright install chromium
```

Install the `regenerate` extra as well only when the user explicitly requests
OpenAI-assisted narration regeneration.

## Verify dependencies

Run the doctor before rendering:

```bash
pptx2video doctor --svg
```

It verifies the exact FFmpeg filters and encoders used by the renderer, plus
FFprobe, LibreOffice, Poppler, Python packages, the installed runtime,
Playwright, and an executable system or Playwright Chromium browser. It does
not launch a browser GUI. Fix every missing required dependency before
continuing. Plain `pptx2video doctor` remains available as a lightweight
core-only check.

## Render

Choose any input PPTX and a new output directory:

```bash
pptx2video render \
  "/Users/me/Documents/My edited deck.pptx" \
  "/Users/me/Movies/my-new-video-bundle" \
  --resolution 1080p
```

Paths containing spaces are supported. The output directory must not already
exist, which prevents stale media or report reuse. The command reports success
only after the renderer and strict QA finish with zero errors and zero warnings.

The delivered `video.pptx` is the preferred input for later edits. Another PPTX
may use canonical Notes for precise narration, or the same
`## [handle] optional semantic` block grammar in Shape Alt Text plus native
PowerPoint animations. Legacy `Script:` fields remain readable and migrate on
writeback.

## Optional launcher integration

A Finder shortcut, Automator workflow, or independently maintained signed app
may invoke the installed `pptx2video` command. Keep such launchers outside this
package and pass the selected PPTX and output directory as ordinary CLI
arguments. The package does not depend on a repository-relative `.app` path.
