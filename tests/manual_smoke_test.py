import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.ppt.exporter import ExportConfig, export_lyrics_to_pptx


SAMPLE_LYRICS = """Verse 1
Your love, O Lord, reaches to the heavens
Your faithfulness stretches to the skies

Chorus
I will lift my voice and sing
You are good, forever King

Verse 2
In every season You are near
You calm my heart and cast out fear
"""


def run() -> None:
    out = Path("output") / "smoke_test.pptx"
    result = export_lyrics_to_pptx(
        SAMPLE_LYRICS,
        out,
        ExportConfig(use_local_ai=False, max_stanzas_per_slide=3),
    )
    print(f"Created: {result.resolve()}")
    print("Open in PowerPoint slideshow mode and verify click-to-transition behavior.")


if __name__ == "__main__":
    run()
