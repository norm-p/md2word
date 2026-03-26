"""
md2word CLI entry point.

Pipeline:
  1. interpret  (deterministic) — create / update based on --base
  2. approach select  (deterministic)
  3. run approach     (AI for map/edit_plan/summary, deterministic for rest)
  4. validate         (deterministic)
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import click

from .ai.interpret import interpret
from .validate import extract_docx_text, validate_output

_VERSION_RE = re.compile(r"^(.+?)(?:_v(\d{3,}))$")


def _resolve_args(
    positional: tuple[Path, ...],
    input_opt: Path | None,
    base_opt: Path | None,
) -> tuple[Path, Path | None]:
    """Merge positional args with -i/-b switches. Returns (input_md, base)."""
    input_md = input_opt
    base = base_opt

    if len(positional) >= 1 and input_md is None:
        input_md = positional[0]
    elif len(positional) >= 1 and input_md is not None:
        if base is None:
            base = positional[0]

    if len(positional) >= 2 and base is None:
        base = positional[1]

    if len(positional) > 2:
        raise click.UsageError("Too many positional arguments (expected at most 2: INPUT.md [BASE.docx]).")

    if input_md is None:
        raise click.UsageError("Missing input: provide INPUT.md as first argument or via -i/--input.")

    if not input_md.exists():
        raise click.UsageError(f"Input file not found: {input_md}")

    return input_md, base


def _next_versioned_path(base: Path) -> Path:
    """Generate the next versioned output path from a base .docx path.

    report.docx       → report_v001.docx
    report_v001.docx  → report_v002.docx
    report_v012.docx  → report_v013.docx

    If the computed path already exists, keep incrementing.
    """
    stem = base.stem
    suffix = base.suffix
    parent = base.parent

    m = _VERSION_RE.match(stem)
    if m:
        name_part = m.group(1)
        next_num = int(m.group(2)) + 1
    else:
        name_part = stem
        next_num = 1

    while True:
        candidate = parent / f"{name_part}_v{next_num:03d}{suffix}"
        if not candidate.exists():
            return candidate
        next_num += 1


def _detect_approach(mode: str, output: Path) -> str:
    if output.suffix.lower() == ".xlsx":
        return "excel"
    if mode == "update":
        return "xml_edit"
    return "pandoc"


@click.command()
@click.argument("args", nargs=-1, type=click.Path(path_type=Path))
@click.option("-i", "--input", "input_opt", type=click.Path(exists=True, path_type=Path), default=None,
              help="Markdown input file.")
@click.option("-b", "--base", type=click.Path(path_type=Path), default=None,
              help="Existing .docx to update (triggers update mode).")
@click.option("-o", "--output", type=click.Path(path_type=Path), default=None,
              help="Output file path (default: versioned from base, or input stem + .docx).")
@click.option("--ref-doc", type=click.Path(path_type=Path), default=None,
              help="Reference .docx for pandoc styling (create mode only).")
@click.option("--toc", is_flag=True, default=False,
              help="Inject TOC placeholder (create mode, pandoc only).")
@click.option("--accept-changes", is_flag=True, default=False,
              help="Skip change summary and confirmation (update mode).")
@click.option("-v", "--verbose", is_flag=True, default=False,
              help="Show step-by-step progress including AI reasoning.")
def main(
    args: tuple[Path, ...],
    input_opt: Path | None,
    base: Path | None,
    output: Path | None,
    ref_doc: Path | None,
    toc: bool,
    accept_changes: bool,
    verbose: bool,
) -> None:
    """Convert a Markdown file to a Word document.

    \b
    Positional usage:
      md2word INPUT.md                    # create mode
      md2word INPUT.md BASE.docx          # update, writes BASE_v001.docx
      md2word INPUT.md BASE.docx -o OUT   # update, writes OUT

    \b
    Switch usage:
      md2word -i INPUT.md -b BASE.docx -o OUT.docx
    """
    # Resolve positional args + switches
    input_md, base = _resolve_args(args, input_opt, base)

    # 1. Interpret (deterministic)
    result = interpret(input_md.read_text(encoding="utf-8"), has_base=base is not None)

    if verbose:
        click.echo(f"Mode: {result.mode} — {result.rationale}")

    # 2. Resolve output path
    if output is None:
        if base is not None:
            output = _next_versioned_path(base)
        else:
            output = input_md.with_suffix(".docx")

    if verbose:
        click.echo(f"Output: {output}")

    # 3. Read base text (for verbose only now, xml_edit reads it internally)
    if base is not None:
        if not base.exists():
            click.echo(f"Error: base file not found: {base}", err=True)
            sys.exit(1)
        if verbose:
            click.echo(f"Reading base document: {base}")

    # 4. Select approach
    approach = _detect_approach(result.mode, output)
    if verbose:
        click.echo(f"Approach: {approach}")

    # 5. Dispatch
    _dispatch(
        approach=approach,
        input_md=input_md,
        output=output,
        base=base,
        ref_doc=ref_doc,
        toc=toc,
        accept_changes=accept_changes,
        verbose=verbose,
    )

    # 6. Validate
    validate_output(output)
    click.echo(f"Done: {output}")


def _dispatch(
    approach: str,
    input_md: Path,
    output: Path,
    base: Path | None,
    ref_doc: Path | None,
    toc: bool,
    accept_changes: bool,
    verbose: bool,
) -> None:
    if approach == "pandoc":
        from .approaches.pandoc import run
        run(input_md, output, ref_doc=ref_doc, toc=toc, verbose=verbose)
    elif approach == "xml_edit":
        from .approaches.xml_edit import run
        run(input_md, output, target=base, accept_changes=accept_changes, verbose=verbose)
    elif approach == "excel":
        from .approaches.excel import run  # type: ignore[import]
        run(input_md, output, verbose=verbose)
    else:
        raise RuntimeError(f"Unknown approach: {approach}")
