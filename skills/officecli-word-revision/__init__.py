"""Packaged entry point for the existing Word revision skill."""
from pathlib import Path
import runpy

__version__ = (Path(__file__).parent / "VERSION").read_text().strip()


def main():
    runpy.run_path(
        str(Path(__file__).parent / "scripts" / "word_revision_pipeline.py"),
        run_name="__main__",
    )
