from pathlib import Path
import typer
from rich import print
from .converter import export_workbook_sheets_to_pdf

app = typer.Typer(help="Excel → per-sheet PDFs via xlwings (requires Excel)")

@app.command()
def convert(
    input: Path,
    output: Path = typer.Option(None, "-o", "--output", help="Output directory"),
    include_hidden: bool = typer.Option(False, help="Include hidden sheets"),
    portrait: bool = typer.Option(False, help="Force portrait orientation"),
    landscape: bool = typer.Option(False, help="Force landscape orientation"),
    no_fit: bool = typer.Option(False, help="Disable 'fit to 1 page wide'"),
):
    """Convert a single Excel workbook into per-sheet PDFs."""
    if portrait and landscape:
        print("[red]Choose either --portrait or --landscape, not both.[/red]")
        raise typer.Exit(code=2)

    orientation = True if landscape else False if portrait else None
    pdfs = export_workbook_sheets_to_pdf(
        xls_path=input,
        out_dir=output,
        include_hidden=include_hidden,
        landscape=orientation,
        fit_to_page=not no_fit,
    )
    for p in pdfs:
        print(f"[green]Created:[/green] {p}")

@app.command(name="convert-folder")
def convert_folder(
    folder: Path,
    output_root: Path = typer.Option(None, "-o", "--output-root", help="Root for outputs (mirror structure)"),
    recursive: bool = typer.Option(True, help="Recurse into subfolders"),
    include_hidden: bool = typer.Option(False, help="Include hidden sheets"),
    portrait: bool = typer.Option(False, help="Force portrait orientation"),
    landscape: bool = typer.Option(False, help="Force landscape orientation"),
    no_fit: bool = typer.Option(False, help="Disable 'fit to 1 page wide'"),
):
    """Convert ALL Excel files in a folder (one PDF per sheet)."""
    if portrait and landscape:
        print("[red]Choose either --portrait or --landscape, not both.[/red]")
        raise typer.Exit(code=2)

    exts = {".xls", ".xlsx", ".xlsm", ".xlsb"}
    files = folder.rglob("*") if recursive else folder.iterdir()

    orientation = True if landscape else False if portrait else None
    total = 0
    for f in files:
        if f.suffix.lower() not in exts or not f.is_file():
            continue
        out_dir = (output_root / f.parent.relative_to(folder)) if output_root else None
        pdfs = export_workbook_sheets_to_pdf(
            xls_path=f,
            out_dir=out_dir,
            include_hidden=include_hidden,
            landscape=orientation,
            fit_to_page=not no_fit,
        )
        for p in pdfs:
            print(f"[green]Created:[/green] {p}")
            total += 1
    print(f"[bold]Done.[/bold] Created {total} PDF(s).")
