# /// script
# requires-python = ">=3.13"
# dependencies = [
#     "oletools",
#     "typer",
# ]
# ///

# Based on: https://www.xltrail.com/blog/auto-export-vba-commit-hook
import os
import shutil
from pathlib import Path

import typer
from oletools.olevba3 import VBA_Parser

OFFICE_FILE_EXTENSIONS = ("xlsb", "xls", "xlsm", "xla", "xlt", "xlam", "potm", "pptm")
KEEP_NAME = False  # Set this to True if you would like to keep "Attribute VB_Name"


def dump_vba(
    document_path: str | Path,
    keep_names: bool = False,
    out_path: str | Path = "src.vba",
):
    """Dump VBA macros from an Office document to a directory.

    Parameters
    ----------
    document_path : str | Path
        Path to the Office document.
    keep_names : bool, optional
        Whether to keep "Attribute VB_Name" lines in the output, by default False.
    out_path : str | Path, optional
        Output directory path, by default "src.vba".
    """
    vba_path = Path(out_path)
    vba_parser = VBA_Parser(document_path)
    vba_modules = (
        vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []
    )

    for _, _, filename, content in vba_modules:
        _content = []
        for line in content.splitlines():
            if line.startswith("Attribute") and "VB_" in line:
                if "VB_Name" in line and keep_names:
                    _content.append(line)
            else:
                _content.append(line)
        non_empty_lines_of_code = len([c for c in _content if c])
        if non_empty_lines_of_code == 0:
            continue
        if not vba_path.exists():
            vba_path.mkdir(parents=True)
        with open(vba_path / filename, "w", encoding="utf-8") as f:
            f.write("\n".join(_content).strip())


def main():
    for root, dirs, files in os.walk("."):
        for f in dirs:
            if f.endswith(".vba"):
                shutil.rmtree(os.path.join(root, f))

        for f in files:
            # Skip temporary office files
            if f.startswith("~$"):
                continue
            if f.endswith(OFFICE_FILE_EXTENSIONS):
                try:
                    dump_vba(os.path.join(root, f))
                except PermissionError as e:
                    typer.echo(e, err=True)
                    raise typer.Exit(e.errno if e.errno is not None else 1)


if __name__ == "__main__":
    main()
