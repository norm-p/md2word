"""
pandoc.py — Create a new .docx from Markdown via pypandoc (create mode).

Uses a reference .docx template (--ref-doc) when provided to inherit fonts,
colors, and named styles. Without one, pandoc's default styling is applied.
"""

from __future__ import annotations

import os
import subprocess
from pathlib import Path

import click

from .. import output as out


def _wsl_to_windows(p: Path) -> str:
    """Convert a Linux path to a Windows path when running on WSL with a .exe pandoc."""
    try:
        result = subprocess.run(
            ["wslpath", "-w", str(p.resolve())],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return str(p)


def _ensure_pandoc() -> str:
    """Resolve the pandoc binary path, downloading if necessary.

    Resolution order:
    1. PYPANDOC_PANDOC env var (from .env or shell) — explicit path to binary
    2. System pandoc on PATH
    3. Auto-download via pypandoc into its default platform location
    """
    import shutil

    try:
        from md2word.ai.client import load_env
        load_env()
    except ImportError:
        pass

    # 1. Explicit path from env
    pandoc_env = os.environ.get("PYPANDOC_PANDOC")
    if pandoc_env:
        pandoc_path = Path(pandoc_env)
        if not pandoc_path.is_absolute():
            pandoc_path = Path(__file__).resolve().parents[3] / pandoc_path
        if pandoc_path.exists():
            os.environ["PYPANDOC_PANDOC"] = str(pandoc_path)
            return str(pandoc_path)

    # 2. System pandoc on PATH
    if shutil.which("pandoc"):
        return "pandoc"

    # 3. Auto-download
    out.info("Pandoc not found — downloading automatically (one-time setup)")
    import pypandoc
    pypandoc.download_pandoc()
    out.detail("Pandoc installed successfully")
    return "pandoc"


def run(
    input_path: Path,
    output_path: Path,
    *,
    ref_doc: Path | None,
    toc: bool,
    verbose: bool,
) -> None:
    """Convert input_path (.md) to output_path (.docx) using pandoc."""
    pandoc_bin = _ensure_pandoc()

    extra_args: list[str] = []
    if ref_doc is not None:
        if not ref_doc.exists():
            raise FileNotFoundError(f"Reference doc not found: {ref_doc}")
        extra_args.append(f"--reference-doc={ref_doc}")
    if toc:
        extra_args.append("--toc")

    if verbose:
        out.verbose(f"Running pandoc: {input_path} -> {output_path}", True)
        if extra_args:
            out.verbose(f"extra_args: {extra_args}", True)

    # When pandoc is a Windows .exe on WSL, pypandoc can't pass Linux paths
    # (it validates os.path.exists before calling pandoc, and pandoc.exe
    # can't resolve /mnt/c/... paths). Call pandoc directly via subprocess.
    if pandoc_bin.endswith(".exe"):
        cmd = [
            pandoc_bin,
            _wsl_to_windows(input_path),
            "-o", _wsl_to_windows(output_path),
        ]
        if ref_doc is not None:
            cmd.append(f"--reference-doc={_wsl_to_windows(ref_doc)}")
        if toc:
            cmd.append("--toc")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode != 0:
            raise RuntimeError(
                f"Pandoc exited with code {result.returncode}: {result.stderr}"
            )
    else:
        import pypandoc
        pypandoc.convert_file(
            str(input_path),
            "docx",
            outputfile=str(output_path),
            extra_args=extra_args,
        )
