from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from pptx2video.cli import _parser, _render
from pptx2video.render_video import (
    AnimationEffect,
    AnimationLayer,
    AnimationSlideAssets,
    SlidePair,
    VisualCue,
    _animation_pointer_cues,
    encode_segment,
)


class AnimationPointerTests(unittest.TestCase):
    @staticmethod
    def _layer(
        path: Path,
        *,
        order: int,
        start: float,
        x: int,
        y: int,
        width: int,
        height: int,
    ) -> AnimationLayer:
        return AnimationLayer(
            effect=AnimationEffect(
                order=order,
                locator=f"shape-{order}",
                name="Fade In",
                start=start,
                duration=0.4,
                timing_source="animation_pane",
                shape_id=str(order),
            ),
            path=path,
            x=x,
            y=y,
            width=width,
            height=height,
        )

    def test_pointer_uses_first_rendered_layer_in_animation_group(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            assets = AnimationSlideAssets(
                index=1,
                base_frame=root / "base.png",
                layers=[
                    self._layer(
                        root / "one.png", order=1, start=0.5,
                        x=10, y=20, width=40, height=20,
                    ),
                    self._layer(
                        root / "two.png", order=2, start=0.5,
                        x=70, y=40, width=20, height=20,
                    ),
                    self._layer(
                        root / "three.png", order=3, start=1.5,
                        x=80, y=60, width=20, height=20,
                    ),
                ],
            )

            self.assertEqual(
                _animation_pointer_cues(
                    assets, style="none", width=100, height=80, segment_end=3.0,
                ),
                [],
            )
            cues = _animation_pointer_cues(
                assets, style="laser", width=100, height=80, segment_end=3.0,
            )

            self.assertEqual(len(cues), 2)
            self.assertEqual(cues[0].point, (0.3, 0.375))
            self.assertEqual((cues[0].start, cues[0].end), (0.5, 3.0))
            self.assertEqual(cues[0].style, "laser")
            self.assertEqual(cues[1].point, (0.9, 0.875))

    def test_public_cli_forwards_animation_pointer(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            pptx = root / "deck.pptx"
            pptx.write_bytes(b"pptx")
            output = root / "bundle"
            args = _parser().parse_args(
                ["render", str(pptx), str(output), "--animation-pointer", "cursor"]
            )

            def fake_run(_module: str, command: list[str]) -> None:
                self.assertEqual(
                    command[command.index("--animation-pointer") + 1],
                    "cursor",
                )
                report = output / "assets" / "meta" / "reports" / "video_qa_report.json"
                report.parent.mkdir(parents=True)
                report.write_text(
                    json.dumps({"passed": True, "counts": {"error": 0, "warning": 0}}),
                    encoding="utf-8",
                )

            with mock.patch("pptx2video.cli._run_runtime", side_effect=fake_run):
                self.assertEqual(_render(args), 0)

    def test_animation_spotlight_and_pointer_share_the_filter_graph(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            base = root / "base.png"
            layer_path = root / "layer.png"
            audio = root / "audio.mp3"
            for path in (base, layer_path, audio):
                path.write_bytes(b"fixture")
            assets = AnimationSlideAssets(
                index=1,
                base_frame=base,
                layers=[
                    self._layer(
                        layer_path, order=1, start=0.25,
                        x=40, y=30, width=80, height=40,
                    )
                ],
            )
            pair = SlidePair(index=1, frame=base, audio=audio, duration=2.0)

            with mock.patch(
                "pptx2video.render_video.subprocess.run",
                return_value=SimpleNamespace(returncode=0, stderr=""),
            ) as run:
                encode_segment(
                    pair,
                    root / "segment.mp4",
                    width=320,
                    height=180,
                    fps=5,
                    pad_tail=0.5,
                    ffmpeg="ffmpeg",
                    visual_cues=[
                        VisualCue(
                            cue_type="highlight",
                            start=0.4,
                            end=1.4,
                            box=(0.1, 0.1, 0.3, 0.3),
                            point=(0.25, 0.25),
                            style="spotlight",
                            no_ink_tighten=True,
                        )
                    ],
                    animation_assets=assets,
                    animation_pointer="cursor",
                )

            command = run.call_args.args[0]
            graph = command[command.index("-filter_complex") + 1]
            self.assertIn("[animsrc0]", graph)
            self.assertIn("[animbase0][spotmask0]overlay", graph)
            self.assertIn("[cursor]", graph)
            self.assertIn("[spotbase0][cursor]overlay", graph)


if __name__ == "__main__":
    unittest.main()
