"""PyInstaller entry point for md2word.exe."""

from freeze_hook import patch_pandoc_path
patch_pandoc_path()

from md2word.cli import main  # noqa: E402

if __name__ == "__main__":
    main()
