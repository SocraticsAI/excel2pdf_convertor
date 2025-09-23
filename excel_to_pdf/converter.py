from pathlib import Path
import re
import sys
import xlwings as xw

ILLEGAL = r'[<>:"/\\|?*]'
ILLEGAL_RE = re.compile(ILLEGAL)

def safe_name(name: str, replacement: str = "_") -> str:
    return ILLEGAL_RE.sub(replacement, name)

def _apply_page_setup_cross_platform(sht, *, fit_to_page: bool, landscape: bool | None):
    """Best-effort page setup on both Mac and Windows."""
    is_mac = sys.platform == "darwin"

    if is_mac:
        # AppleScript: Excel exposes a command, not a PageSetup object
        kwargs = {}
        if fit_to_page:
            kwargs["zoom"] = False
            kwargs["fit_to_pages_wide"] = 1
            kwargs["fit_to_pages_tall"] = False
        if landscape is not None:
            kwargs["orientation"] = 2 if landscape else 1  # 2=landscape, 1=portrait
        if kwargs:
            try:
                sht.api.page_setup(**kwargs)
            except Exception:
                pass
    else:
        # Windows COM
        try:
            ps = sht.api.PageSetup
            if fit_to_page:
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = False
            if landscape is not None:
                ps.Orientation = 2 if landscape else 1
        except Exception:
            pass

def export_workbook_sheets_to_pdf(
    xls_path: Path,
    out_dir: Path | None = None,
    include_hidden: bool = False,
    fit_to_page: bool = True,
    landscape: bool | None = None,
) -> list[Path]:
    """
    Export each worksheet/chartsheet to its own PDF.
    Uses xlwings' Book.to_pdf(include=[sheetname]) for cross-platform reliability.
    """
    xls_path = Path(xls_path)
    if not xls_path.exists():
        raise FileNotFoundError(xls_path)

    if out_dir is None:
        out_dir = xls_path.parent / f"{xls_path.stem}_PDFs"
    out_dir.mkdir(parents=True, exist_ok=True)

    created: list[Path] = []

    app = xw.App(visible=False, add_book=False)
    try:
        book = app.books.open(str(xls_path), update_links=False, read_only=True)

        for sht in book.sheets:
            # Skip hidden unless requested
            try:
                visible = bool(sht.api.Visible)
            except Exception:
                visible = True
            if not include_hidden and not visible:
                continue

            # Page setup (best-effort)
            _apply_page_setup_cross_platform(
                sht, fit_to_page=fit_to_page, landscape=landscape
            )

            # Export via xlwings helper (cross-platform)
            pdf_name = f"{safe_name(xls_path.stem)} - {safe_name(sht.name)}.pdf"
            pdf_path = out_dir / pdf_name

            try:
                # book-level export but only including this sheet
                book.to_pdf(path=str(pdf_path), include=[sht.name])
                created.append(pdf_path)
            except Exception as e:
                print(f"Warning: failed to export '{sht.name}': {e}")

        book.close()
    finally:
        app.quit()

    return created
