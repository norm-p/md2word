"""
output.py — Standardized console output for md2word CLI.

Provides a consistent output convention:
  [N/M]     Top-level pipeline step (always shown)
    detail  Sub-info under current step (always shown)
    verbose Sub-detail (shown only with -v)
  Warning:  Non-fatal degradation (yellow)
  Error:    Fatal problem (red, stderr)
  Done:     Final success (green)
"""

from __future__ import annotations

import click

_step = 0


def init_steps() -> None:
    """Reset step counter for a new pipeline run."""
    global _step
    _step = 0


def step(msg: str) -> None:
    """Top-level pipeline phase. Always shown."""
    global _step
    _step += 1
    click.echo(f"[step {_step}] {msg}")


def detail(msg: str) -> None:
    """Detail under the current step. Always shown."""
    click.echo(f"  {msg}")


def verbose(msg: str, is_verbose: bool) -> None:
    """Sub-detail shown only with -v."""
    if is_verbose:
        click.echo(f"    {msg}")


def warn(msg: str) -> None:
    """Non-fatal warning. Always shown."""
    click.echo(click.style(f"  Warning: {msg}", fg="yellow"))


def error(msg: str) -> None:
    """Fatal error. Always shown, to stderr."""
    click.echo(click.style(f"  Error: {msg}", fg="red"), err=True)


def success(msg: str) -> None:
    """Final completion message."""
    click.echo(click.style(msg, fg="green"))


def info(msg: str) -> None:
    """Informational notice (non-critical, always shown)."""
    click.echo(f"  Info: {msg}")
