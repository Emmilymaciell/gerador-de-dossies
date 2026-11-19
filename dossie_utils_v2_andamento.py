from __future__ import annotations

import hashlib
import json
import os
import math
import re
import shutil
import subprocess
import tempfile
import time
import unicodedata
from copy import deepcopy
from dataclasses import asdict, dataclass
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable, Iterable, Optional

import fitz  # PyMuPDF
from PIL import Image, ImageChops, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pypdf import PdfReader

SCRIPT_HOME = Path(__file__).resolve().parent

# Pasta-raiz com uma subpasta por mês ("01. Janeiro 2026", ...). Na versão pública
# aponta para os dados fictícios gerados por examples/gerar_exemplo.py; pode ser
# sobrescrita pela variável de ambiente DOSSIE_MONTHS_DIR.
MONTHS_BASE_DIR = Path(os.environ.get("DOSSIE_MONTHS_DIR", SCRIPT_HOME / "examples" / "produtos"))
DEFAULT_MONTH_ROOTS = [
    str(MONTHS_BASE_DIR / "01. Janeiro 2026"),
]

DEFAULT_OUTPUT_ROOT = Path(os.environ.get("DOSSIE_OUTPUT_DIR", SCRIPT_HOME / "saida"))

# Modelos Word de capa (opcionais). Sem eles, as capas são desenhadas pelo próprio código.
INTERNAL_CAPAS_DIR = SCRIPT_HOME / "CAPAS"

USE_WORD_TEMPLATE_INITIAL_COVER = True
USE_WORD_TEMPLATES_FOR_SECTION_COVERS = False
USE_TEMPLATE_BASE_IMAGES_FOR_SECTION_COVERS = True
SHOW_WORD_WINDOW = False
TOC_TC_IDENTIFIER = "DSS"
DOCX_ATTACHMENT_RENDER_SCALE = 2.4
PDF_BUILD_PREVIEW_KEEP_PREAMBLE = True
PDF_OUTLINE_TITLE_MAX = 240
COVER_RENDER_SCALE = 2.4
COVER_EDGE_TRIM_PX = 3
PDF_TOC_PAGE_WIDTH = 595.0
PDF_TOC_PAGE_HEIGHT = 842.0

GROUP_TITLES = {
    "G1": "Coordenação Executiva",
    "G2": "Engenharia",
    "G3": "Qualidade e Inspeção",
    "G4": "Administrativo-Financeiro",
    "G5": "Jurídico",
}

# Títulos genéricos (fictícios) usados quando o relatório do grupo não lista os produtos.
FALLBACK_PRODUCT_TITLES = {
    "P1.1": "Relatório de Coordenação",
    "P1.2": "Relatório de Comunicação",
    "P1.3": "Relatório de Tecnologia da Informação",
    "P2.1": "Relatório de Engenharia",
    "P2.2": "Relatório de Análise de Projetos",
    "P2.3": "Relatório de Acompanhamento de Obras",
    "P3.1": "Relatório de Qualidade",
    "P3.2": "Planos de Inspeção",
    "P3.3": "Relatório de Inspeção",
    "P4.1": "Relatório Financeiro",
    "P4.2": "Relatório de Riscos",
    "P5.1": "Relatório Jurídico",
    "P5.2": "Relatório de Gestão Contratual",
}

MONTHS_PT = [
    "janeiro", "fevereiro", "marco", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

PRODUCT_CODE_RE = re.compile(r"(?:produto\s*|\bP\s*)([1-5])\s*[._-]?\s*([0-9]{1,2})", re.I)
PRODUCT_INLINE_RE = re.compile(r"\b([1-5])\s*[._-]\s*([0-9]{1,2})\b")
REPORT_RE = re.compile(r"relat.*grupo\s*([1-5])", re.I)
EXCLUDE_PARTS = {"~$", ".tmp", ".bak", "desktop.ini", ".ds_store"}
NON_PDF_EXTENSIONS = {".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".ppt", ".pptx", ".csv", ".msg", ".zip"}
DELIVERABLE_EXTENSIONS = {".pdf"} | NON_PDF_EXTENSIONS
CONTROL_KEYWORDS = ("controle", "entrega", "produto")
REVIEW_TABLE_HEADERS = ("ID", "Tipo", "Código", "Título")


@dataclass
class SourceFile:
    path: str
    ext: str
    title: str
    original_name: str
    product_code: str
    group_code: str
    is_pdf: bool
    page_count: int = 0
    inferred_from: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class Entry:
    id: str
    kind: str
    title: str
    group_code: str
    product_code: str = ""
    status: str = ""
    notes: str = ""
    source_path: str = ""
    copied_path: str = ""
    ext: str = ""
    page_count: int = 0
    base_page: int = 0
    in_control_excel: bool = False
    source_kind: str = ""
    review_source_path: str = ""
    original_name: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class CoverRender:
    image_path: Path
    width_pt: float
    height_pt: float


# ---------------------------------------------------------------------------
# Normalização e helpers gerais
# ---------------------------------------------------------------------------

def strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "")
    return "".join(ch for ch in text if not unicodedata.combining(ch))

def normalize_text(text: str) -> str:
    text = strip_accents(text or "")
    text = text.replace("#U00f3", "o").replace("#U00e7", "c").replace("#U00e3", "a")
    text = text.replace("%C2%B0", " ").replace("N°", "N ").replace("Nº", "N ")
    text = re.sub(r"[_\-]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text

def slugify(text: str, max_len: int = 120) -> str:
    text = normalize_text(text)
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    if len(text) > max_len:
        text = text[:max_len].rstrip("-")
    return text or "arquivo"

def pretty_month_label(path: Path) -> str:
    name = re.sub(r"^\d+\.\s*", "", path.name).strip()
    return name or path.name

def short_month_key(path: Path) -> str:
    return slugify(pretty_month_label(path), 40)

def month_cover_label(month_label: str) -> str:
    month, year = month_year_parts(month_label)
    if month and year:
        return f"{month.capitalize()} - {year}"
    return month_label

def parse_product_code(value: str) -> Optional[str]:
    if not value:
        return None
    m = PRODUCT_CODE_RE.search(value) or PRODUCT_INLINE_RE.search(value)
    if not m:
        return None
    return f"P{int(m.group(1))}.{int(m.group(2))}"

def short_group_code(group: int | str) -> str:
    if isinstance(group, int):
        return f"G{group}"
    group = str(group)
    if group.upper().startswith("G"):
        digits = re.sub(r"[^0-9]", "", group) or "0"
        return f"G{int(digits)}"
    return f"G{int(group)}"

def product_sort_key(product_code: str) -> tuple[int, int]:
    m = re.search(r"P([1-5])\.([0-9]+)", product_code or "")
    if not m:
        return (99, 99)
    return (int(m.group(1)), int(m.group(2)))

def parse_group_from_name(name: str) -> Optional[str]:
    norm = normalize_text(name)
    m = REPORT_RE.search(norm)
    if m:
        return short_group_code(int(m.group(1)))
    m = re.search(r"grupo\s*([1-5])", norm)
    if m:
        return short_group_code(int(m.group(1)))
    m = re.search(r"\bg\s*([1-5])\b", norm)
    if m:
        return short_group_code(int(m.group(1)))
    return None

def clean_display_title(name: str, product_code: str = "") -> str:
    title = Path(str(name)).stem if "." in str(name) else str(name)
    title = title.replace("%C2%B0", " ").replace("#U00b0", " ")
    title = title.replace("_", " ")
    title = re.sub(r"\s+", " ", title).strip()

    if product_code:
        title = re.sub(rf"^\s*{re.escape(product_code)}\s*[-–—: ]*", "", title, flags=re.I)
        title = re.sub(rf"^\s*{re.escape(product_code.replace('P', ''))}\s*[-–—: ]*", "", title, flags=re.I)
        title = re.sub(rf"^\s*produto\s*{re.escape(product_code.replace('P', ''))}\s*[-–—: ]*", "", title, flags=re.I)

    title = re.sub(r"^\s*[1-5]\.[0-9]{1,2}\s*[-–—: ]*", "", title)
    title = re.sub(r"^\s*P\s*[1-5]\.[0-9]{1,2}\s*[-–—: ]*", "", title, flags=re.I)
    title = re.sub(r"\s+", " ", title).strip(" -_–—:")
    return title or "Documento"

def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize_text(a), normalize_text(b)).ratio()

def month_year_parts(month_label: str) -> tuple[str, str]:
    norm = normalize_text(month_label)
    month = next((m for m in MONTHS_PT if m in norm), "")
    year_match = re.search(r"(20\d{2})", month_label)
    return month, year_match.group(1) if year_match else ""

def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path

def safe_unlink(path: Path) -> None:
    if path.exists():
        if path.is_file() or path.is_symlink():
            try: path.unlink()
            except: pass
        else:
            shutil.rmtree(path, ignore_errors=True)

def month_output_paths(month_root: Path, output_root: Path) -> dict[str, Path]:
    label = pretty_month_label(month_root)
    num_prefix = month_root.name.split(".")[0].strip() if "." in month_root.name else short_month_key(month_root)
    folder_name = f"{num_prefix} - {label}".strip()
    base = output_root / folder_name
    annex_review = base / "Anexos - Revisão"
    return {
        "month_dir": base,
        "annex_review_dir": annex_review,
        "manifest": base / "Manifesto.json",
        "base_docx": annex_review / f"Dossiê Mensal - {label} - Revisão.docx",
        "base_pdf": base / f"Dossiê Mensal - {label} - Revisão.pdf",
        "final_docx": base / f"Dossiê Mensal - {label} - Final.docx",
        "final_pdf": base / f"Dossiê Mensal - {label} - Final.pdf",
        "levantamento_xlsx": base / f"Conferência do Dossiê - {label}.xlsx",
    }

def safe_pdf_page_count(path: Path) -> int:
    try:
        return len(PdfReader(str(path)).pages)
    except Exception:
        return 0

def read_pdf_text(path: Path, pages: int | None = None) -> str:
    try:
        reader = PdfReader(str(path))
        chunks: list[str] = []
        page_list = reader.pages if pages is None else reader.pages[:pages]
        for pg in page_list:
            try:
                chunks.append(pg.extract_text() or "")
            except Exception:
                pass
        return "\n".join(chunks)
    except Exception:
        return ""

def extract_pdf_title_guess(path: Path) -> str:
    text = read_pdf_text(path, pages=1)
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines()]
    lines = [ln for ln in lines if 6 <= len(ln) <= 180]
    bad_tokens = ("consorcio", "programa", "pagina", "relatorio mensal", "grupo")
    for line in lines[:18]:
        norm = normalize_text(line)
        if any(tok in norm for tok in bad_tokens):
            continue
        if re.search(r"(nota tecnica|parecer|relatorio|plano|oficio|carta|ata|memoria|checklist|analise|manual|controle|edital|minuta)", norm):
            return line.strip(" -")
    return ""


# ---------------------------------------------------------------------------
# Descoberta de estrutura e arquivos
# ---------------------------------------------------------------------------

def discover_report_pdfs(month_root: Path) -> dict[str, Path]:
    reports: dict[str, Path] = {}
    for path in month_root.rglob("*.pdf"):
        if any(ex in normalize_text(path.name) for ex in ("dossie mensal", "dossie sem anexos", "dossie final")):
            continue
        group_code = parse_group_from_name(path.name)
        if not group_code:
            text = read_pdf_text(path, pages=2)
            m = re.search(r"grupo\s*([1-5])", normalize_text(text))
            if m:
                group_code = short_group_code(int(m.group(1)))
        if group_code and group_code not in reports:
            reports[group_code] = path
    return reports

def parse_report_structure(report_pdf: Path) -> tuple[str, dict[str, str]]:
    text = read_pdf_text(report_pdf, pages=8)
    group_code = parse_group_from_name(report_pdf.name)
    if not group_code:
        m = re.search(r"grupo\s*([1-5])", normalize_text(text))
        if m:
            group_code = short_group_code(int(m.group(1)))
    group_title = GROUP_TITLES.get(group_code or "", f"Grupo {group_code or ''}".strip())

    products: dict[str, str] = {}
    patterns = [
        r"Produto\s+([1-5]\.[0-9]{1,2})\s*[–\-]\s*(.+)",
        r"\bP([1-5]\.[0-9]{1,2})\s*[–\-]\s*(.+)",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text):
            code = m.group(1)
            code = code if code.startswith("P") else f"P{code}"
            title = re.sub(r"\.{2,}.*", "", m.group(2)).strip(" .-–—")
            title = clean_display_title(title, code)
            if title:
                products[code] = title

    if not products and group_code:
        for code, title in FALLBACK_PRODUCT_TITLES.items():
            if code.startswith(f"P{group_code[1]}."):
                products[code] = title

    return group_title, dict(sorted(products.items(), key=lambda kv: product_sort_key(kv[0])))

def should_ignore_file(path: Path) -> bool:
    lower = path.name.lower()
    return any(tok in lower for tok in EXCLUDE_PARTS) or path.name.startswith("~$")

def guess_product_from_path(path: Path, known_products: list[str]) -> tuple[Optional[str], str]:
    candidates = [path.name] + [p.name for p in path.parents[:5]]
    for token in candidates:
        code = parse_product_code(token)
        if code:
            return code, "nome ou pasta"
    if path.suffix.lower() == ".pdf":
        code = parse_product_code(read_pdf_text(path, pages=1))
        if code:
            return code, "conteúdo do PDF"
    return None, ""

def detect_folder_context(path: Path, month_root: Path) -> tuple[Optional[str], Optional[str], str]:
    parents: list[Path] = []
    current = path.parent
    month_root_resolved = month_root.resolve()
    while True:
        parents.append(current)
        try:
            if current.resolve() == month_root_resolved:
                break
        except Exception:
            pass
        if current.parent == current:
            break
        current = current.parent

    nearest_group_code: Optional[str] = None
    nearest_group_idx: Optional[int] = None
    nearest_product_code: Optional[str] = None
    nearest_product_idx: Optional[int] = None

    for idx, folder in enumerate(parents):
        folder_name = folder.name
        if nearest_group_code is None:
            gcode = parse_group_from_name(folder_name)
            if gcode:
                nearest_group_code = gcode
                nearest_group_idx = idx
        if nearest_product_code is None:
            pcode = parse_product_code(folder_name)
            if pcode:
                nearest_product_code = pcode
                nearest_product_idx = idx

    if nearest_group_code and nearest_product_code:
        product_group = short_group_code(nearest_product_code[1])
        if product_group != nearest_group_code:
            return nearest_group_code, None, "pasta do grupo"

    if nearest_group_code and (nearest_product_code is None or (nearest_group_idx is not None and nearest_product_idx is not None and nearest_group_idx < nearest_product_idx)):
        return nearest_group_code, None, "pasta do grupo"

    if nearest_product_code:
        return short_group_code(nearest_product_code[1]), nearest_product_code, "pasta do produto"

    if nearest_group_code:
        return nearest_group_code, None, "pasta do grupo"

    return None, None, ""

def guess_group_from_path(path: Path) -> tuple[Optional[str], str]:
    candidates = [path.name] + [p.name for p in path.parents[:6]]
    for token in candidates:
        group_code = parse_group_from_name(token)
        if group_code:
            return group_code, "nome ou pasta"
    if path.suffix.lower() == ".pdf":
        text = read_pdf_text(path, pages=1)
        m = re.search(r"grupo\s*([1-5])", normalize_text(text))
        if m:
            return short_group_code(int(m.group(1))), "conteúdo do PDF"
    return None, ""


def discover_deliverable_files(month_root: Path, known_products: list[str]) -> list[SourceFile]:
    reports = set(discover_report_pdfs(month_root).values())
    files: list[SourceFile] = []
    seen: set[tuple[str, str, str]] = set()

    for path in month_root.rglob("*"):
        if not path.is_file() or should_ignore_file(path):
            continue
        if path in reports:
            continue

        ext = path.suffix.lower()
        if ext not in DELIVERABLE_EXTENSIONS:
            continue

        folder_group_code, folder_product_code, folder_inferred = detect_folder_context(path, month_root)

        code: Optional[str] = None
        group_code: Optional[str] = None
        inferred = ""

        if folder_group_code and not folder_product_code:
            group_code = folder_group_code
            inferred = folder_inferred
        elif folder_product_code:
            code = folder_product_code
            group_code = short_group_code(code[1])
            inferred = folder_inferred
        else:
            code, inferred = guess_product_from_path(path, known_products)
            if code:
                group_code = short_group_code(code[1])
            else:
                group_code, inferred = guess_group_from_path(path)
                if not group_code:
                    continue

        if folder_group_code and code and short_group_code(code[1]) != folder_group_code:
            code = None
            group_code = folder_group_code
            inferred = "pasta do grupo"

        filename_title = clean_display_title(path.name, product_code=code or "")
        title = filename_title

        if ext == ".pdf":
            pdf_title = extract_pdf_title_guess(path)
            if pdf_title:
                pdf_title = clean_display_title(pdf_title, product_code=code or "")
                generic_names = {"documento", "arquivo", "scan", "digitalizacao", "digitalização"}
                if normalize_text(filename_title) in generic_names:
                    title = pdf_title or filename_title

        key_code = code or group_code or ""
        key = (str(path), key_code, normalize_text(title))
        if key in seen:
            continue
        seen.add(key)

        files.append(
            SourceFile(
                path=str(path),
                ext=ext,
                title=title,
                original_name=path.name,
                product_code=code or "",
                group_code=group_code or "",
                is_pdf=(ext == ".pdf"),
                page_count=safe_pdf_page_count(path) if ext == ".pdf" else 0,
                inferred_from=inferred or "nome do arquivo",
            )
        )

    return files


def looks_like_control_excel(path: Path) -> bool:
    norm = normalize_text(str(path))
    return path.suffix.lower() in {".xlsx", ".xlsm", ".xls"} and all(k in norm for k in CONTROL_KEYWORDS)

def find_control_excels(month_root: Path) -> list[Path]:
    excels = [p for p in month_root.rglob("*.xls*") if looks_like_control_excel(p)]
    if excels:
        return sorted(excels)
    return sorted(month_root.rglob("*.xls*"))

def detect_header_row(ws) -> Optional[int]:
    wanted = {"produto", "documento", "arquivo", "descr", "status", "tipo", "empresa"}
    for row_idx in range(1, min(ws.max_row, 30) + 1):
        hits = 0
        for col_idx in range(1, min(ws.max_column, 30) + 1):
            value = ws.cell(row_idx, col_idx).value
            txt = normalize_text(str(value)) if value is not None else ""
            if any(w in txt for w in wanted):
                hits += 1
        if hits >= 3:
            return row_idx
    return None

def parse_control_excel(path: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return rows

    for ws in wb.worksheets:
        header_row = detect_header_row(ws)
        if not header_row:
            continue

        headers: dict[int, str] = {}
        for c in range(1, min(ws.max_column, 40) + 1):
            val = ws.cell(header_row, c).value
            if val is not None:
                headers[c] = normalize_text(str(val))

        col_product = next((c for c, h in headers.items() if "produto" in h), None)
        col_doc = next((c for c, h in headers.items() if "documento" in h or "arquivo" in h), None)
        col_desc = next((c for c, h in headers.items() if "descr" in h), None)
        col_status = next((c for c, h in headers.items() if "status" in h), None)
        col_type = next((c for c, h in headers.items() if "tipo" in h), None)

        for r in range(header_row + 1, min(ws.max_row, 5000) + 1):
            raw_product = ws.cell(r, col_product).value if col_product else None
            code = parse_product_code(str(raw_product or ""))
            if not code:
                continue

            raw_title = str(ws.cell(r, col_doc).value or "").strip() if col_doc else ""
            raw_desc = str(ws.cell(r, col_desc).value or "").strip() if col_desc else ""
            raw_status = ws.cell(r, col_status).value if col_status else None
            raw_type = ws.cell(r, col_type).value if col_type else None

            candidates = [raw_title, raw_desc]
            candidates = [
                c for c in candidates
                if c and normalize_text(c) not in {normalize_text(code), normalize_text(code.replace("P", ""))}
            ]
            if not candidates:
                continue

            title = max(candidates, key=lambda x: len(normalize_text(x)))
            title = clean_display_title(title, code)
            if not title or re.fullmatch(r"[0-9.\- ]+", title) or len(normalize_text(title)) < 3:
                continue

            rows.append({
                "product_code": code,
                "group_code": short_group_code(code[1]),
                "title": title,
                "status_text": str(raw_status or "").strip(),
                "type_text": str(raw_type or "").strip(),
                "excel_path": str(path),
                "worksheet": ws.title,
            })

    return rows


# ---------------------------------------------------------------------------
# Pastas e cópia de anexos
# ---------------------------------------------------------------------------

def _safe_name(name: str) -> str:
    name = re.sub(r"[<>:\|?*]", " ", name).strip()
    name = re.sub(r"\s+", " ", name)
    return name or "Sem nome"

def _clip_name(name: str, max_len: int) -> str:
    name = _safe_name(name)
    if len(name) <= max_len:
        return name
    stem = Path(name).stem
    suffix = Path(name).suffix
    keep = max(8, max_len - len(suffix) - 1)
    stem = stem[:keep].rstrip()
    return f"{stem}{suffix}" if suffix else stem

def _short_folder_label(code: str, title: str, max_len: int = 72) -> str:
    base = _safe_name(f"{code} - {title}")
    if len(base) <= max_len:
        return base
    room = max(12, max_len - len(code) - 3)
    return _safe_name(f"{code} - {_clip_name(title, room)}")

def _to_windows_extended_path(path: Path) -> str:
    p = str(Path(path).resolve())
    if os.name != "nt":
        return p
    prefix = chr(92) * 2 + "?" + chr(92)
    if p.startswith(prefix):
        return p
    if p.startswith(chr(92) * 2):
        return prefix + "UNC" + chr(92) + p.lstrip(chr(92))
    return prefix + p

def unique_preserved_path(dest_dir: Path, original_name: str) -> Path:
    ensure_dir(dest_dir)
    original_name = _safe_name(original_name)
    candidate = dest_dir / original_name
    if not candidate.exists():
        return candidate

    stem = Path(original_name).stem
    suffix = Path(original_name).suffix
    idx = 2
    while True:
        candidate = dest_dir / f"{stem} ({idx}){suffix}"
        if not candidate.exists():
            return candidate
        idx += 1

def group_folder_name(group_code: str, group_title: str) -> str:
    return _short_folder_label(group_code, group_title, max_len=48)

def product_folder_name(product_code: str, product_title: str) -> str:
    return _short_folder_label(product_code, product_title, max_len=72)

def product_dest_dir(annex_review_dir: Path, group_code: str, group_title: str, product_code: str, product_title: str) -> Path:
    return ensure_dir(
        annex_review_dir / group_folder_name(group_code, group_title) / product_folder_name(product_code, product_title)
    )

def copy_file_to_product_folder(src: Path, dest_dir: Path) -> Path:
    src = Path(src)
    dest_dir = ensure_dir(Path(dest_dir))
    if not src.exists():
        raise FileNotFoundError(f"Arquivo de origem não encontrado: {src}")

    dest = unique_preserved_path(dest_dir, src.name)
    try:
        shutil.copy2(_to_windows_extended_path(src), _to_windows_extended_path(dest))
    except FileNotFoundError as exc:
        raise FileNotFoundError(
            "Falha ao copiar arquivo para Anexos - Revisão. "
            f"Origem: {src} | Destino: {dest}"
        ) from exc
    return dest


# ---------------------------------------------------------------------------
# Manifesto
# ---------------------------------------------------------------------------

def build_initial_structure(month_root: Path) -> dict[str, Any]:
    reports = discover_report_pdfs(month_root)
    groups: list[dict[str, Any]] = []
    known_products: dict[str, str] = {}
    for idx in range(1, 6):
        gcode = short_group_code(idx)
        report_pdf = reports.get(gcode)
        gtitle = GROUP_TITLES[gcode]
        products = {k: v for k, v in FALLBACK_PRODUCT_TITLES.items() if k.startswith(f"P{idx}.")}
        if report_pdf:
            pg_title, parsed = parse_report_structure(report_pdf)
            if pg_title:
                gtitle = pg_title
            if parsed:
                products.update(parsed)

        groups.append({
            "group_code": gcode,
            "group_title": gtitle,
            "report_pdf": str(report_pdf) if report_pdf else "",
            "documents": [],
            "products": [],
        })
        known_products.update(products)
    return {"groups": groups, "known_products": known_products}


def _next_manifest_entry_id(manifest: dict[str, Any], prefix: str = "D") -> str:
    nums = []
    for entry in manifest.get("entries", []):
        entry_id = str(entry.get("id", ""))
        if entry_id.startswith(prefix):
            digits = re.sub(r"[^0-9]", "", entry_id[len(prefix):] or entry_id)
            if digits:
                nums.append(int(digits))
    next_num = (max(nums) + 1) if nums else 1
    return f"{prefix}{next_num:04d}"

def rebuild_manifest_entry_order(manifest: dict[str, Any]) -> dict[str, Any]:
    entry_by_id = {entry["id"]: entry for entry in manifest.get("entries", [])}
    ordered: list[dict[str, Any]] = []
    for group in manifest.get("groups", []):
        group["documents"] = sorted(
            group.get("documents", []),
            key=lambda d: (d.get("status") != "pdf_localizado", normalize_text(d.get("title", ""))),
        )
        group["products"] = sorted(group.get("products", []), key=lambda p: product_sort_key(p["product_code"]))
        ordered.append(entry_by_id.get(group["id"], group))
        for doc_item in group.get("documents", []):
            ordered.append(entry_by_id.get(doc_item["id"], doc_item))
        for product in group.get("products", []):
            product["documents"] = sorted(
                product.get("documents", []),
                key=lambda d: (d.get("status") != "pdf_localizado", normalize_text(d.get("title", ""))),
            )
            ordered.append(entry_by_id.get(product["id"], product))
            for doc_item in product.get("documents", []):
                ordered.append(entry_by_id.get(doc_item["id"], doc_item))
    manifest["entries"] = ordered
    return manifest


def build_month_manifest(month_root: Path, output_root: Path, clean_month_output: bool = True) -> tuple[dict[str, Any], dict[str, Path]]:
    month_root = Path(month_root)
    paths = month_output_paths(month_root, output_root)

    if clean_month_output:
        safe_unlink(paths["month_dir"])

    ensure_dir(paths["month_dir"])
    ensure_dir(paths["annex_review_dir"])

    base = build_initial_structure(month_root)
    groups = base["groups"]
    known_products = base["known_products"]

    source_files = discover_deliverable_files(month_root, list(known_products.keys()))
    entries: list[dict[str, Any]] = []
    next_id = 1

    def new_id(prefix: str) -> str:
        nonlocal next_id
        value = f"{prefix}{next_id:04d}"
        next_id += 1
        return value

    group_map = {g["group_code"]: g for g in groups}
    product_map = {p["product_code"]: p for g in groups for p in g["products"]}

    for group in groups:
        gid = new_id("G")
        group["id"] = gid
        group.setdefault("documents", [])
        group["folder_path"] = str(ensure_dir(paths["annex_review_dir"] / group_folder_name(group["group_code"], group["group_title"])))
        entries.append(Entry(id=gid, kind="group", title=f"{group['group_code']} – {group['group_title']}", group_code=group["group_code"]).to_dict())
        for product in group["products"]:
            pid = new_id("P")
            product["id"] = pid
            product["folder_path"] = str(product_dest_dir(paths["annex_review_dir"], group["group_code"], group["group_title"], product["product_code"], product["product_title"]))
            entries.append(Entry(
                id=pid,
                kind="product",
                title=f"{product['product_code']} – {product['product_title']}",
                group_code=group["group_code"],
                product_code=product["product_code"],
            ).to_dict())

    missing_rows: list[dict[str, Any]] = []
    seen_doc_keys: set[tuple[str, str]] = set()

    for sf in source_files:
        if not sf.group_code:
            continue

        group = group_map.setdefault(
            sf.group_code,
            {
                "group_code": sf.group_code,
                "group_title": GROUP_TITLES.get(sf.group_code, sf.group_code),
                "documents": [],
                "products": [],
                "report_pdf": "",
            },
        )
        group.setdefault("documents", [])
        if "id" not in group:
            group["id"] = new_id("G")
            group["folder_path"] = str(ensure_dir(paths["annex_review_dir"] / group_folder_name(group["group_code"], group["group_title"])))
            entries.append(Entry(
                id=group["id"],
                kind="group",
                title=f"{group['group_code']} – {group['group_title']}",
                group_code=group["group_code"],
            ).to_dict())

        if sf.product_code:
            prod = product_map.get(sf.product_code)
            if not prod:
                prod = {
                    "product_code": sf.product_code,
                    "product_title": known_products.get(sf.product_code, FALLBACK_PRODUCT_TITLES.get(sf.product_code, sf.product_code)),
                    "documents": [],
                    "id": new_id("P"),
                    "folder_path": str(product_dest_dir(paths["annex_review_dir"], group["group_code"], group["group_title"], sf.product_code, known_products.get(sf.product_code, FALLBACK_PRODUCT_TITLES.get(sf.product_code, sf.product_code)))),
                }
                group["products"].append(prod)
                product_map[sf.product_code] = prod
                entries.append(Entry(
                    id=prod["id"],
                    kind="product",
                    title=f"{prod['product_code']} – {prod['product_title']}",
                    group_code=group["group_code"],
                    product_code=sf.product_code,
                ).to_dict())

            key = (sf.product_code, normalize_text(sf.title))
            if key in seen_doc_keys:
                continue
            seen_doc_keys.add(key)

            copied = ""
            status = "pdf_localizado" if sf.is_pdf else "arquivo_nao_pdf"
            if sf.is_pdf:
                dest_dir = Path(prod["folder_path"])
                copied = str(copy_file_to_product_folder(Path(sf.path), dest_dir))

            entry = Entry(
                id=new_id("D"),
                kind="document",
                title=f"{sf.product_code} – {sf.title}",
                group_code=sf.group_code,
                product_code=sf.product_code,
                status=status,
                notes="",
                source_path=sf.path,
                copied_path=copied,
                ext=sf.ext,
                page_count=sf.page_count,
                source_kind="folder_file",
                review_source_path=copied,
                original_name=sf.original_name,
            ).to_dict()
            prod.setdefault("documents", []).append(entry)
            entries.append(entry)

            if not sf.is_pdf:
                missing_rows.append({
                    "Grupo": sf.group_code,
                    "Produto": sf.product_code,
                    "Documento": sf.title,
                    "Situação": "Arquivo encontrado sem PDF",
                    "Origem": "Pasta do mês",
                    "Caminho": sf.path,
                    "Extensão": sf.ext,
                    "Página base": "",
                    "Observações": "Converter para PDF e colocar dentro da subpasta do produto em Anexos - Revisão.",
                })
        else:
            key = (sf.group_code, normalize_text(sf.title))
            if key in seen_doc_keys:
                continue
            seen_doc_keys.add(key)

            copied = ""
            status = "pdf_localizado" if sf.is_pdf else "arquivo_nao_pdf"
            if sf.is_pdf:
                dest_dir = Path(group["folder_path"])
                copied = str(copy_file_to_product_folder(Path(sf.path), dest_dir))

            entry = Entry(
                id=new_id("D"),
                kind="document",
                title=sf.title,
                group_code=sf.group_code,
                product_code="",
                status=status,
                notes="",
                source_path=sf.path,
                copied_path=copied,
                ext=sf.ext,
                page_count=sf.page_count,
                source_kind="group_folder_file",
                review_source_path=copied,
                original_name=sf.original_name,
            ).to_dict()
            group.setdefault("documents", []).append(entry)
            entries.append(entry)

            if not sf.is_pdf:
                missing_rows.append({
                    "Grupo": sf.group_code,
                    "Produto": "",
                    "Documento": sf.title,
                    "Situação": "Arquivo encontrado sem PDF",
                    "Origem": "Pasta do mês",
                    "Caminho": sf.path,
                    "Extensão": sf.ext,
                    "Página base": "",
                    "Observações": "Converter para PDF e colocar diretamente dentro da pasta do grupo em Anexos - Revisão.",
                })

    control_rows: list[dict[str, Any]] = []
    for excel in find_control_excels(month_root):
        control_rows.extend(parse_control_excel(excel))

    for row in control_rows:
        prod = product_map.get(row["product_code"])
        if not prod:
            gcode = row["group_code"]
            group = group_map[gcode]
            prod = {
                "product_code": row["product_code"],
                "product_title": known_products.get(row["product_code"], FALLBACK_PRODUCT_TITLES.get(row["product_code"], row["product_code"])),
                "documents": [],
                "id": new_id("P"),
                "folder_path": str(product_dest_dir(paths["annex_review_dir"], group["group_code"], group["group_title"], row["product_code"], known_products.get(row["product_code"], FALLBACK_PRODUCT_TITLES.get(row["product_code"], row["product_code"])))),
            }
            group["products"].append(prod)
            product_map[row["product_code"]] = prod
            entries.append(Entry(
                id=prod["id"],
                kind="product",
                title=f"{prod['product_code']} – {prod['product_title']}",
                group_code=gcode,
                product_code=row["product_code"],
            ).to_dict())

        matched = None
        best = 0.0
        for d in prod.get("documents", []):
            score = similarity(row["title"], d["title"])
            if score > best:
                best = score
                matched = d

        if matched and best >= 0.58:
            matched["in_control_excel"] = True
            continue

        entry = Entry(
            id=new_id("D"),
            kind="document",
            title=f"{row['product_code']} – {row['title']}",
            group_code=row["group_code"],
            product_code=row["product_code"],
            status="controle_sem_pdf_localizado",
            notes="",
            ext="",
            source_kind="control_excel",
            in_control_excel=True,
            original_name="",
        ).to_dict()
        prod.setdefault("documents", []).append(entry)
        entries.append(entry)

        missing_rows.append({
            "Grupo": row["group_code"],
            "Produto": row["product_code"],
            "Documento": row["title"],
            "Situação": "Registro do controle de entregas sem PDF localizado",
            "Origem": Path(row["excel_path"]).name,
            "Caminho": row["excel_path"],
            "Extensão": "",
            "Página base": "",
            "Observações": f"Consta no controle, mas não foi localizado PDF. Aba: {row['worksheet']}. Colocar o PDF na subpasta do produto em Anexos - Revisão.",
        })

    manifest = {
        "version": 5,
        "month_root": str(month_root),
        "month_label": pretty_month_label(month_root),
        "month_cover_label": month_cover_label(pretty_month_label(month_root)),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "paths": {k: str(v) for k, v in paths.items()},
        "groups": groups,
        "entries": entries,
        "missing_rows": missing_rows,
        "initial_cover_template": str(find_cover_template("initial") or ""),
    }
    manifest = rebuild_manifest_entry_order(manifest)
    return manifest, paths


# ---------------------------------------------------------------------------
# Estilos DOCX
# ---------------------------------------------------------------------------

def create_style(document: Document, name: str, style_type, base: str | None = None):
    styles = document.styles
    if name in styles:
        return styles[name]
    style = styles.add_style(name, style_type)
    if base and base in styles:
        style.base_style = styles[base]
    return style

def configure_docx_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Aptos"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
    normal.font.size = Pt(10.5)

    s = create_style(doc, "DossieMainTitle", WD_STYLE_TYPE.PARAGRAPH, "Normal")
    s.font.name = "Aptos Display"
    s.font.bold = True
    s.font.size = Pt(21)
    s.font.color.rgb = RGBColor(31, 78, 95)
    s.paragraph_format.space_after = Pt(6)

    s = create_style(doc, "DossieSubtitle", WD_STYLE_TYPE.PARAGRAPH, "Normal")
    s.font.size = Pt(11)
    s.font.color.rgb = RGBColor(86, 94, 102)
    s.paragraph_format.space_after = Pt(6)

    for level, size, left in [(1, 11.0, 0), (2, 10.0, 0.7), (3, 9.2, 1.4)]:
        s = create_style(doc, f"DossieTOC{level}", WD_STYLE_TYPE.PARAGRAPH, "Normal")
        s.font.size = Pt(size)
        s.font.bold = level == 1
        s.font.color.rgb = RGBColor(34, 48, 58) if level == 1 else RGBColor(59, 69, 80)
        s.paragraph_format.left_indent = Cm(left)
        s.paragraph_format.space_after = Pt(1.6 if level == 1 else 1.0)
        s.paragraph_format.keep_together = True
        s.paragraph_format.keep_with_next = False

        builtin_name = f"TOC {level}"
        if builtin_name in doc.styles:
            toc_style = doc.styles[builtin_name]
            toc_style.font.name = "Aptos"
            toc_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
            toc_style.font.size = Pt(size)
            toc_style.font.bold = level == 1
            toc_style.font.color.rgb = RGBColor(34, 48, 58) if level == 1 else RGBColor(59, 69, 80)
            toc_style.paragraph_format.left_indent = Cm(left)
            toc_style.paragraph_format.space_after = Pt(1.6 if level == 1 else 1.0)

    s = create_style(doc, "DossieSectionTitle", WD_STYLE_TYPE.PARAGRAPH, "Normal")
    s.font.name = "Aptos Display"
    s.font.bold = True
    s.font.size = Pt(16)
    s.font.color.rgb = RGBColor(31, 78, 95)
    s.paragraph_format.space_after = Pt(8)

    s = create_style(doc, "DossieHint", WD_STYLE_TYPE.PARAGRAPH, "Normal")
    s.font.size = Pt(10)
    s.font.italic = True
    s.font.color.rgb = RGBColor(86, 94, 102)
    s.paragraph_format.space_after = Pt(6)

def configure_standard_section(section) -> None:
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

def set_section_vertical_alignment(section, mode: str = "top") -> None:
    sectPr = section._sectPr
    v_align = sectPr.find(qn("w:vAlign"))
    if v_align is None:
        v_align = OxmlElement("w:vAlign")
        sectPr.append(v_align)
    v_align.set(qn("w:val"), mode)

def configure_zero_margin_section(section, width_pt: float, height_pt: float) -> None:
    section.page_width = Pt(width_pt)
    section.page_height = Pt(height_pt)
    section.left_margin = Mm(0)
    section.right_margin = Mm(0)
    section.top_margin = Mm(0)
    section.bottom_margin = Mm(0)
    set_section_vertical_alignment(section, "top")

def _set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


# ---------------------------------------------------------------------------
# Word templates de capa - AJUSTADO PARA GARANTIR LEITURA
# ---------------------------------------------------------------------------

TEMPLATE_FILE_OVERRIDES = {
    "initial": "CAPA DOSSIÊ.docx",
    "group": "CAPA GRUPO.docx",
    "product": "capa_subsecao_produtos.docx",
    "document": "capa_documento_tecnica.docx",
}

TEMPLATE_HINTS = {
    "initial": ["CAPA DOSSIÊ", "capa_dossie", "capa_inicial", "dossie", "inicial"],
    "group": ["CAPA GRUPO", "capa_grupo", "grupo"],
    "product": ["capa_subsecao_produtos", "CAPA PRODUTO", "capa_produto", "produto", "subsecao"],
    "document": ["capa_documento_tecnica", "CAPA DOCUMENTO", "capa_documento", "documento", "tecnica"],
}

PLACEHOLDER_MAP = {
    "initial": [
        ("Outubro - 2025", lambda value: value),
        ("MÊS - ANO", lambda value: value),
        ("MES - ANO", lambda value: value),
        ("[MES]", lambda value: value),
        ("NOME DO MES", lambda value: value),
    ],
    "group": [
        ("G1 – Coordenação Executiva", lambda value: value),
        ("G1 - Coordenação Executiva", lambda value: value),
        ("NOME DO GRUPO", lambda value: value),
        ("TITULO DO GRUPO", lambda value: value),
        ("[GRUPO]", lambda value: value),
    ],
    "product": [
        ("Título do Produto", lambda value: value),
        ("Titulo do Produto", lambda value: value),
        ("NOME DO PRODUTO", lambda value: value),
        ("[PRODUTO]", lambda value: value),
    ],
    "document": [
        ("Descrição do Documento", lambda value: value),
        ("Descricao do Documento", lambda value: value),
        ("NOME DO DOCUMENTO", lambda value: value),
        ("TITULO DO DOCUMENTO", lambda value: value),
        ("[DOCUMENTO]", lambda value: value),
    ],
}

SECTION_TEMPLATE_TEXT_STYLES = {
    "group": {
        "fallback_box": (0.0820, 0.5000, 0.6350, 0.6070),
        "color": (255, 255, 255),
        "bold": True,
        "align": "left",
        "v_align": "top",
        "max_size": 50,
        "min_size": 20,
        "line_spacing": 1.08,
        "expand": (0.025, 0.020, 0.030, 0.050),
        "top_padding": 6,
        "vertical_bias_px": 0,
        "use_detected_box": True,
    },
    "product": {
        "fallback_box": (0.0950, 0.3720, 0.6650, 0.5400),
        "color": (33, 95, 109),
        "bold": True,
        "align": "left",
        "v_align": "top",
        "max_size": 38,
        "min_size": 18,
        "line_spacing": 1.05,
        "expand": (0.000, 0.000, 0.000, 0.000),
        "top_padding": 0,
        "vertical_bias_px": 0,
        "use_detected_box": False,
    },
    "document": {
        "fallback_box": (0.1550, 0.4100, 0.8450, 0.5750),
        "color": (255, 255, 255),
        "bold": False,
        "align": "center",
        "v_align": "center",
        "max_size": 34,
        "min_size": 18,
        "line_spacing": 1.05,
        "expand": (0.000, 0.000, 0.000, 0.000),
        "top_padding": 0,
        "vertical_bias_px": 18,
        "use_detected_box": False,
    },
}


def _iter_all_paragraphs(doc: Document):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for section in doc.sections:
        for container in (section.header, section.footer):
            for p in container.paragraphs:
                yield p
            for table in container.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            yield p

def replace_text_across_runs(paragraph, target: str, replacement: str) -> int:
    if not target or not paragraph.runs:
        return 0

    replacements = 0
    while True:
        runs = list(paragraph.runs)
        if not runs:
            break
        full_text = "".join(run.text for run in runs)
        start = full_text.find(target)
        if start < 0:
            break
        end = start + len(target)

        cursor = 0
        start_idx = None
        end_idx = None
        start_off = 0
        end_off = 0
        for idx, run in enumerate(runs):
            run_text = run.text
            next_cursor = cursor + len(run_text)
            if start_idx is None and start < next_cursor:
                start_idx = idx
                start_off = start - cursor
            if end_idx is None and end <= next_cursor:
                end_idx = idx
                end_off = end - cursor
                break
            cursor = next_cursor

        if start_idx is None or end_idx is None:
            break

        for idx in range(start_idx, end_idx + 1):
            run = runs[idx]
            text = run.text
            if idx == start_idx == end_idx:
                run.text = text[:start_off] + replacement + text[end_off:]
            elif idx == start_idx:
                run.text = text[:start_off] + replacement
            elif idx == end_idx:
                run.text = text[end_off:]
            else:
                run.text = ""

        replacements += 1

    return replacements

def apply_template_replacements(doc: Document, template_kind: str, value: str) -> int:
    total = 0
    for paragraph in _iter_all_paragraphs(doc):
        for target, resolver in PLACEHOLDER_MAP.get(template_kind, []):
            total += replace_text_across_runs(paragraph, target, resolver(value))
    return total

def find_cover_template(kind: str) -> Optional[Path]:
    if not INTERNAL_CAPAS_DIR.exists():
        return None

    candidates = [p for p in INTERNAL_CAPAS_DIR.iterdir() if p.is_file() and p.suffix.lower() in {".docx", ".doc"}]

    override_name = TEMPLATE_FILE_OVERRIDES.get(kind, "").strip()
    if override_name:
        exact = INTERNAL_CAPAS_DIR / override_name
        if exact.exists():
            return exact

        override_norm = normalize_text(Path(override_name).name)
        for path in candidates:
            if normalize_text(path.name) == override_norm or normalize_text(path.stem) == normalize_text(Path(override_name).stem):
                return path

    hints = TEMPLATE_HINTS.get(kind, [kind])

    for hint in hints:
        for path in candidates:
            if hint.lower() == path.stem.lower():
                return path

    hint_norms = [normalize_text(h) for h in hints]
    for path in candidates:
        norm = normalize_text(path.name)
        if any(hn in norm for hn in hint_norms):
            return path
    return None

def _resolve_cover_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = []
    if bold:
        candidates.extend([
            r"C:\Windows\Fonts\arialbd.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ])
    candidates.extend([
        r"C:\Windows\Fonts\arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ])
    for candidate in candidates:
        try:
            if Path(candidate).exists():
                return ImageFont.truetype(candidate, size=size)
        except Exception:
            pass
    return ImageFont.load_default()

def _wrap_cover_text(draw: ImageDraw.ImageDraw, text: str, font, max_width: int) -> list[str]:
    words = (text or '').split()
    if not words:
        return ['']
    lines = [words[0]]
    for word in words[1:]:
        trial = lines[-1] + ' ' + word
        width = draw.textbbox((0, 0), trial, font=font)[2]
        if width <= max_width:
            lines[-1] = trial
        else:
            lines.append(word)
    return lines

def _fast_cover_palette(template_kind: str) -> tuple[tuple[int, int, int], tuple[int, int, int], tuple[int, int, int]]:
    palettes = {
        'initial': ((22, 34, 52), (246, 248, 250), (214, 221, 227)),
        'group': ((21, 35, 56), (245, 247, 250), (92, 141, 181)),
        'product': ((17, 56, 91), (245, 247, 250), (76, 164, 160)),
        'document': ((34, 34, 38), (250, 250, 250), (120, 120, 128)),
    }
    return palettes.get(template_kind, palettes['document'])

def render_fast_cover_image(template_kind: str, value: str, temp_dir: Path, cache: dict[str, CoverRender]) -> CoverRender:
    cache_key = hashlib.sha1(f'fast::{template_kind}::{value}'.encode('utf-8')).hexdigest()
    if cache_key in cache:
        return cache[cache_key]

    ensure_dir(temp_dir)
    out_png = temp_dir / f'fast_{template_kind}_{cache_key}.png'

    width_px, height_px = 1240, 1754
    width_pt, height_pt = 595.0, 842.0
    bg, fg, accent = _fast_cover_palette(template_kind)

    image = Image.new('RGB', (width_px, height_px), bg)
    draw = ImageDraw.Draw(image)

    title_font = _resolve_cover_font(34, bold=False)
    text_font = _resolve_cover_font(66 if template_kind != 'document' else 58, bold=True)
    if template_kind == 'initial':
        eyebrow = 'Dossiê Mensal'
    elif template_kind == 'group':
        eyebrow = 'Grupo'
    elif template_kind == 'product':
        eyebrow = 'Produto'
    else:
        eyebrow = 'Documento'

    draw.rounded_rectangle((90, 120, width_px - 90, 142), radius=10, fill=accent)
    draw.text((100, 165), eyebrow, font=title_font, fill=fg)

    max_width = width_px - 180
    lines = _wrap_cover_text(draw, value, text_font, max_width)
    line_boxes = [draw.textbbox((0, 0), line, font=text_font) for line in lines]
    line_height = max((box[3] - box[1] for box in line_boxes), default=70) + 18
    total_h = line_height * len(lines)
    y = max(280, (height_px - total_h) // 2)
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=text_font)
        line_w = bbox[2] - bbox[0]
        x = (width_px - line_w) // 2
        draw.text((x, y), line, font=text_font, fill=fg)
        y += line_height

    footer = 'Programa'
    footer_font = _resolve_cover_font(28, bold=False)
    fb = draw.textbbox((0,0), footer, font=footer_font)
    draw.text(((width_px - (fb[2]-fb[0])) // 2, height_px - 120), footer, font=footer_font, fill=(220,220,225))

    image.save(out_png)
    _trim_cover_image_edges(out_png)
    result = CoverRender(image_path=out_png, width_pt=width_pt, height_pt=height_pt)
    cache[cache_key] = result
    return result


def _trim_cover_image_edges(image_path: Path, trim_px: int = COVER_EDGE_TRIM_PX) -> None:
    try:
        if trim_px <= 0:
            return
        with Image.open(image_path) as img:
            width, height = img.size
            if width <= trim_px * 2 + 10 or height <= trim_px * 2 + 10:
                return
            cropped = img.crop((trim_px, trim_px, width - trim_px, height - trim_px))
            cropped.save(image_path)
    except Exception:
        pass


def _resolve_relative_box(box: tuple[float, float, float, float], image_size: tuple[int, int]) -> tuple[int, int, int, int]:
    width, height = image_size
    x0, y0, x1, y1 = box
    return (
        int(width * x0),
        int(height * y0),
        int(width * x1),
        int(height * y1),
    )


def _expand_relative_box(
    box: tuple[int, int, int, int],
    image_size: tuple[int, int],
    expand: tuple[float, float, float, float],
) -> tuple[int, int, int, int]:
    width, height = image_size
    x0, y0, x1, y1 = box
    left, top, right, bottom = expand
    x0 = max(0, int(round(x0 - width * left)))
    y0 = max(0, int(round(y0 - height * top)))
    x1 = min(width - 1, int(round(x1 + width * right)))
    y1 = min(height - 1, int(round(y1 + height * bottom)))
    return (x0, y0, x1, y1)


def _detect_changed_bbox(
    original_image: Image.Image,
    blank_image: Image.Image,
    threshold: int = 22,
) -> Optional[tuple[int, int, int, int]]:
    if original_image.size != blank_image.size:
        return None

    diff = ImageChops.difference(original_image.convert("RGB"), blank_image.convert("RGB"))
    gray = diff.convert("L")
    mask = gray.point(lambda px: 255 if px >= threshold else 0)
    bbox = mask.getbbox()
    if not bbox:
        return None

    x0, y0, x1, y1 = bbox
    width, height = original_image.size
    area = max(1, (x1 - x0) * (y1 - y0))
    page_area = max(1, width * height)
    ratio = area / page_area
    if ratio > 0.16:
        return None
    return (x0, y0, x1, y1)


def _resolve_template_text_box(
    template_kind: str,
    original_render: Optional[CoverRender],
    base_render: CoverRender,
    style: dict[str, Any],
) -> tuple[int, int, int, int]:
    with Image.open(base_render.image_path) as img:
        fallback_box = _resolve_relative_box(style["fallback_box"], img.size)

    if not style.get("use_detected_box", True):
        return fallback_box

    if original_render is None:
        return fallback_box

    try:
        with Image.open(original_render.image_path).convert("RGB") as original_image, Image.open(base_render.image_path).convert("RGB") as blank_image:
            detected = _detect_changed_bbox(original_image, blank_image)
            if detected is None:
                return fallback_box
            return _expand_relative_box(detected, original_image.size, style.get("expand", (0, 0, 0, 0)))
    except Exception:
        return fallback_box


def _draw_fitted_cover_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    box: tuple[int, int, int, int],
    *,
    color: tuple[int, int, int],
    bold: bool,
    align: str,
    v_align: str,
    max_size: int,
    min_size: int,
    line_spacing: float = 1.10,
    top_padding: int = 0,
    vertical_bias_px: int = 0,
) -> None:
    x0, y0, x1, y1 = box
    max_width = max(20, x1 - x0)
    max_height = max(20, y1 - y0)
    text = (text or "").strip()
    if not text:
        return

    chosen = None
    for size in range(max_size, min_size - 1, -2):
        font = _resolve_cover_font(size, bold=bold)
        lines = _wrap_cover_text(draw, text, font, max_width)
        bboxes = [draw.textbbox((0, 0), line, font=font) for line in lines]
        heights = [bbox[3] - bbox[1] for bbox in bboxes] or [size]
        line_height = max(heights) + max(2, int(size * max(0.0, line_spacing - 1.0)))
        total_height = line_height * len(lines)
        if total_height <= max_height:
            chosen = (font, lines, bboxes, line_height, total_height)
            break

    if chosen is None:
        font = _resolve_cover_font(min_size, bold=bold)
        lines = _wrap_cover_text(draw, text, font, max_width)
        bboxes = [draw.textbbox((0, 0), line, font=font) for line in lines]
        heights = [bbox[3] - bbox[1] for bbox in bboxes] or [min_size]
        line_height = max(heights) + 4
        total_height = line_height * len(lines)
        chosen = (font, lines, bboxes, line_height, total_height)

    font, lines, bboxes, line_height, total_height = chosen
    if v_align == "center":
        y = y0 + max(0, (max_height - total_height) // 2) + int(vertical_bias_px)
    else:
        y = y0 + max(0, top_padding) + int(vertical_bias_px)

    max_start_y = max(y0, y1 - total_height)
    if y > max_start_y:
        y = max_start_y

    for line, bbox in zip(lines, bboxes):
        line_width = bbox[2] - bbox[0]
        if align == "center":
            x = x0 + max(0, (max_width - line_width) // 2)
        elif align == "right":
            x = x1 - line_width
        else:
            x = x0
        draw.text((x, y), line, font=font, fill=color)
        y += line_height


def render_cover_template_original_image(
    template_kind: str,
    temp_dir: Path,
    cache: dict[str, CoverRender],
    word_session: Optional[WordPdfSession] = None,
    progress: Optional[Callable[[str], None]] = None,
) -> Optional[CoverRender]:
    template_path = find_cover_template(template_kind)
    if not template_path:
        return None

    cache_key = hashlib.sha1(f"orig::{template_path.resolve()}::{template_kind}".encode("utf-8")).hexdigest()
    if cache_key in cache:
        return cache[cache_key]

    ensure_dir(temp_dir)
    temp_docx = temp_dir / f"orig_{template_kind}_{cache_key}.docx"
    temp_pdf = temp_dir / f"orig_{template_kind}_{cache_key}.pdf"
    temp_png = temp_dir / f"orig_{template_kind}_{cache_key}.png"

    try:
        if temp_docx.exists():
            safe_unlink(temp_docx)
        shutil.copy2(template_path, temp_docx)
    except Exception:
        return None

    if not try_convert_docx_to_pdf(temp_docx, temp_pdf, word_session=word_session):
        return None

    try:
        pdf = fitz.open(str(temp_pdf))
        page = pdf[0]
        rect = page.rect
        pix = page.get_pixmap(matrix=fitz.Matrix(COVER_RENDER_SCALE, COVER_RENDER_SCALE), alpha=False)
        pix.save(str(temp_png))
        pdf.close()
        _trim_cover_image_edges(temp_png)
        result = CoverRender(image_path=temp_png, width_pt=float(rect.width), height_pt=float(rect.height))
        cache[cache_key] = result
        return result
    except Exception:
        return None


def render_cover_template_base_image(
    template_kind: str,
    temp_dir: Path,
    cache: dict[str, CoverRender],
    word_session: Optional[WordPdfSession] = None,
    progress: Optional[Callable[[str], None]] = None,
) -> Optional[CoverRender]:
    template_path = find_cover_template(template_kind)
    if not template_path:
        return None

    cache_key = hashlib.sha1(f"base::{template_path.resolve()}::{template_kind}".encode("utf-8")).hexdigest()
    if cache_key in cache:
        return cache[cache_key]

    ensure_dir(temp_dir)
    temp_docx = temp_dir / f"base_{template_kind}_{cache_key}.docx"
    temp_pdf = temp_dir / f"base_{template_kind}_{cache_key}.pdf"
    temp_png = temp_dir / f"base_{template_kind}_{cache_key}.png"

    try:
        if temp_docx.exists():
            safe_unlink(temp_docx)
        shutil.copy2(template_path, temp_docx)
        doc = Document(str(temp_docx))
        apply_template_replacements(doc, template_kind, " ")
        doc.save(str(temp_docx))
    except Exception:
        return None

    if progress:
        progress(f"Preparando base visual da capa '{template_path.name}' para {template_kind}...")

    if not try_convert_docx_to_pdf(temp_docx, temp_pdf, word_session=word_session):
        return None

    try:
        pdf = fitz.open(str(temp_pdf))
        page = pdf[0]
        rect = page.rect
        pix = page.get_pixmap(matrix=fitz.Matrix(COVER_RENDER_SCALE, COVER_RENDER_SCALE), alpha=False)
        pix.save(str(temp_png))
        pdf.close()
        _trim_cover_image_edges(temp_png)
        result = CoverRender(image_path=temp_png, width_pt=float(rect.width), height_pt=float(rect.height))
        cache[cache_key] = result
        return result
    except Exception:
        return None


def render_cover_from_template_base_image(
    template_kind: str,
    value: str,
    temp_dir: Path,
    cache: dict[str, CoverRender],
    word_session: Optional[WordPdfSession] = None,
    progress: Optional[Callable[[str], None]] = None,
) -> Optional[CoverRender]:
    style = SECTION_TEMPLATE_TEXT_STYLES.get(template_kind)
    if not style:
        return None

    template_path = find_cover_template(template_kind)
    if not template_path:
        return None

    base_render = render_cover_template_base_image(
        template_kind,
        temp_dir,
        cache,
        word_session=word_session,
        progress=progress,
    )
    if base_render is None:
        return None

    cache_key = hashlib.sha1(
        f"styled::{template_path.resolve()}::{template_kind}::{value}".encode("utf-8")
    ).hexdigest()
    if cache_key in cache:
        return cache[cache_key]

    out_png = temp_dir / f"styled_{template_kind}_{cache_key}.png"
    try:
        image = Image.open(base_render.image_path).convert("RGB")
        draw = ImageDraw.Draw(image)
        original_render = render_cover_template_original_image(
            template_kind,
            temp_dir,
            cache,
            word_session=word_session,
            progress=progress,
        )
        box = _resolve_template_text_box(template_kind, original_render, base_render, style)
        _draw_fitted_cover_text(
            draw,
            value,
            box,
            color=style["color"],
            bold=style["bold"],
            align=style["align"],
            v_align=style.get("v_align", "center"),
            max_size=style["max_size"],
            min_size=style["min_size"],
            line_spacing=style.get("line_spacing", 1.10),
            top_padding=style.get("top_padding", 0),
            vertical_bias_px=style.get("vertical_bias_px", 0),
        )
        image.save(out_png)
        _trim_cover_image_edges(out_png)
        result = CoverRender(
            image_path=out_png,
            width_pt=base_render.width_pt,
            height_pt=base_render.height_pt,
        )
        cache[cache_key] = result
        return result
    except Exception:
        return None


def render_cover_image(
    template_kind: str,
    value: str,
    temp_dir: Path,
    cache: dict[str, CoverRender],
    word_session: Optional[WordPdfSession] = None,
    progress: Optional[Callable[[str], None]] = None,
) -> CoverRender:
    if template_kind == "initial" and USE_WORD_TEMPLATE_INITIAL_COVER:
        cover = render_cover_from_word_template(
            template_kind,
            value,
            temp_dir,
            cache,
            word_session=word_session,
            progress=progress,
        )
        if cover is not None:
            return cover

    if template_kind in {"group", "product", "document"} and USE_TEMPLATE_BASE_IMAGES_FOR_SECTION_COVERS:
        cover = render_cover_from_template_base_image(
            template_kind,
            value,
            temp_dir,
            cache,
            word_session=word_session,
            progress=progress,
        )
        if cover is not None:
            return cover

    if template_kind in {"group", "product", "document"} and USE_WORD_TEMPLATES_FOR_SECTION_COVERS:
        cover = render_cover_from_word_template(
            template_kind,
            value,
            temp_dir,
            cache,
            word_session=word_session,
            progress=progress,
        )
        if cover is not None:
            return cover

    return render_fast_cover_image(template_kind, value, temp_dir, cache)

def render_cover_from_word_template(
    template_kind: str,
    value: str,
    temp_dir: Path,
    cache: dict[str, CoverRender],
    word_session: Optional[WordPdfSession] = None,
    progress: Optional[Callable[[str], None]] = None,
) -> Optional[CoverRender]:

    template_path = find_cover_template(template_kind)
    if not template_path:
        return render_fast_cover_image(template_kind, value, temp_dir, cache)

    cache_key = hashlib.sha1(f"{template_path.resolve()}::{value}".encode("utf-8")).hexdigest()
    if cache_key in cache:
        return cache[cache_key]

    ensure_dir(temp_dir)
    temp_docx = temp_dir / f"{template_kind}_{cache_key}.docx"
    temp_pdf = temp_dir / f"{template_kind}_{cache_key}.pdf"
    temp_png = temp_dir / f"{template_kind}_{cache_key}.png"

    try:
        if temp_docx.exists():
            safe_unlink(temp_docx)
        shutil.copy2(template_path, temp_docx)
        doc = Document(str(temp_docx))
        replaced = apply_template_replacements(doc, template_kind, value)
        doc.save(str(temp_docx))
        if progress:
            template_name = template_path.name
            if replaced > 0:
                progress(f"Usando capa Word '{template_name}' para {template_kind}: {value[:70]}")
            else:
                progress(f"Capa Word '{template_name}' carregada para {template_kind}, sem placeholder substituído automaticamente.")
    except Exception:
        return render_fast_cover_image(template_kind, value, temp_dir, cache)

    if not try_convert_docx_to_pdf(temp_docx, temp_pdf, word_session=word_session):
        return render_fast_cover_image(template_kind, value, temp_dir, cache)

    try:
        pdf = fitz.open(str(temp_pdf))
        page = pdf[0]
        rect = page.rect
        pix = page.get_pixmap(matrix=fitz.Matrix(COVER_RENDER_SCALE, COVER_RENDER_SCALE), alpha=False)
        pix.save(str(temp_png))
        pdf.close()
        _trim_cover_image_edges(temp_png)
        result = CoverRender(image_path=temp_png, width_pt=float(rect.width), height_pt=float(rect.height))
        cache[cache_key] = result
        return result
    except Exception:
        return render_fast_cover_image(template_kind, value, temp_dir, cache)


# ---------------------------------------------------------------------------
# TOC / paginação lógica
# ---------------------------------------------------------------------------


def build_toc_lines(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    lines: list[dict[str, Any]] = []
    for entry in manifest["entries"]:
        if entry["kind"] == "group":
            lines.append({"id": entry["id"], "level": 1, "label": entry["title"]})
        elif entry["kind"] == "product":
            lines.append({"id": entry["id"], "level": 2, "label": entry["title"]})
        elif entry["kind"] == "document":
            level = 2 if not str(entry.get("product_code", "")).strip() else 3
            lines.append({"id": entry["id"], "level": level, "label": entry["title"]})
    return lines


def paginate_toc_entries_docx(entries: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    first_capacity = 32
    next_capacity = 42
    capacities = [first_capacity]
    pages: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    used = 0
    cap_idx = 0

    def line_weight(item: dict[str, Any]) -> int:
        txt = item["label"]
        per_line = 80 if item["level"] == 1 else 92 if item["level"] == 2 else 108
        return max(1, math.ceil(len(txt) / per_line))

    for item in entries:
        capacity = capacities[min(cap_idx, len(capacities) - 1)]
        weight = line_weight(item)
        if used + weight > capacity and current:
            pages.append(current)
            current = []
            used = 0
            cap_idx += 1
            if len(capacities) == 1:
                capacities.append(next_capacity)
            capacity = capacities[min(cap_idx, len(capacities) - 1)]
        current.append(item)
        used += weight

    if current:
        pages.append(current)
    return pages or [[]]

def calculate_page_map(manifest: dict[str, Any], include_initial_cover: bool = True) -> tuple[int, dict[str, int]]:
    toc_pages = len(paginate_toc_entries_docx(build_toc_lines(manifest)))
    page = (1 if include_initial_cover else 0) + toc_pages + 1
    page_map: dict[str, int] = {}
    for entry in manifest["entries"]:
        page_map[entry["id"]] = page
        page += 1
        if entry["kind"] == "document" and entry.get("status") == "pdf_localizado":
            page += int(entry.get("page_count") or 0)
    return toc_pages, page_map


def assign_base_pages(manifest: dict[str, Any], include_initial_cover: bool = True) -> None:
    _, page_map = calculate_page_map(manifest, include_initial_cover=include_initial_cover)
    for entry in manifest["entries"]:
        entry["base_page"] = page_map.get(entry["id"], 0)

    for group in manifest["groups"]:
        g_entry = next(e for e in manifest["entries"] if e["id"] == group["id"])
        group["base_page"] = g_entry["base_page"]
        for d in group.get("documents", []):
            d_entry = next(e for e in manifest["entries"] if e["id"] == d["id"])
            d["base_page"] = d_entry["base_page"]
        for product in group["products"]:
            p_entry = next(e for e in manifest["entries"] if e["id"] == product["id"])
            product["base_page"] = p_entry["base_page"]
            for d in product["documents"]:
                d_entry = next(e for e in manifest["entries"] if e["id"] == d["id"])
                d["base_page"] = d_entry["base_page"]

    update_missing_rows_base_pages(manifest)



def update_missing_rows_base_pages(manifest: dict[str, Any]) -> None:
    for row in manifest.get("missing_rows", []):
        target = None
        row_product = str(row.get("Produto", "")).strip()
        row_group = str(row.get("Grupo", "")).strip()
        row_doc = str(row.get("Documento", "")).strip()
        for entry in manifest["entries"]:
            if entry["kind"] != "document":
                continue
            doc_only = entry["title"].split("–", 1)[-1].strip() if "–" in entry["title"] else entry["title"]
            if row_product:
                if entry["product_code"] == row_product and similarity(doc_only, row_doc) >= 0.60:
                    target = entry
                    break
            else:
                if entry["group_code"] == row_group and not str(entry.get("product_code", "")).strip() and similarity(doc_only, row_doc) >= 0.60:
                    target = entry
                    break
        if target:
            row["Página base"] = target.get("base_page", "")


# ---------------------------------------------------------------------------
# Review DOCX
# ---------------------------------------------------------------------------

def insert_cover_page(doc: Document, cover: CoverRender, first_page: bool = False) -> None:
    if first_page:
        section = doc.sections[0]
    else:
        section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    configure_zero_margin_section(section, cover.width_pt, cover.height_pt)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    run.add_picture(str(cover.image_path), width=Pt(cover.width_pt), height=Pt(cover.height_pt))

def _escape_word_field_text(value: str) -> str:
    return (value or "").replace("\\", "\\\\").replace(chr(34), r"\"")


def _append_field_char(run, field_type: str) -> None:
    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), field_type)
    run._r.append(fld)


def _append_instr_text(run, instr: str) -> None:
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    run._r.append(instr_el)


def add_tc_field(paragraph, entry_text: str, level: int, identifier: str = TOC_TC_IDENTIFIER) -> None:
    hidden = paragraph.add_run()
    hidden.font.hidden = True
    _append_field_char(hidden, "begin")

    instr = paragraph.add_run()
    instr.font.hidden = True
    _append_instr_text(instr, f'TC "{_escape_word_field_text(entry_text)}" \\f {identifier} \\l {level}')

    end = paragraph.add_run()
    end.font.hidden = True
    _append_field_char(end, "end")


def bookmark_name_for_entry(entry_id: str) -> str:
    base = re.sub(r"[^A-Za-z0-9_]+", "_", str(entry_id or "").strip())
    return f"dss_{base or 'item'}"


def _next_bookmark_id() -> int:
    current = getattr(_next_bookmark_id, "_value", 0) + 1
    _next_bookmark_id._value = current
    return current


def add_bookmark(paragraph, bookmark_name: str) -> None:
    bookmark_id = str(_next_bookmark_id())
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), bookmark_id)
    start.set(qn("w:name"), bookmark_name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), bookmark_id)
    paragraph._p.append(start)
    paragraph._p.append(end)


def add_internal_hyperlink(
    paragraph,
    text: str,
    anchor: str,
    *,
    color: str = "22303A",
    underline: bool = False,
) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    hyperlink.set(qn("w:history"), "1")

    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")

    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), color)
    rpr.append(color_el)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single" if underline else "none")
    rpr.append(u)

    run.append(rpr)
    text_el = OxmlElement("w:t")
    text_el.set(qn("xml:space"), "preserve")
    text_el.text = text
    run.append(text_el)

    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def write_manual_toc(doc: Document, manifest: dict[str, Any], page_map: dict[str, int]) -> None:
    toc_pages = paginate_toc_entries_docx(build_toc_lines(manifest))
    first = True
    for page_idx, page_lines in enumerate(toc_pages):
        if not first:
            p = doc.add_paragraph()
            p.add_run().add_break(WD_BREAK.PAGE)
        if page_idx == 0:
            p = doc.add_paragraph(style="DossieMainTitle")
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.add_run("Dossiê Mensal")

            p = doc.add_paragraph(style="DossieSubtitle")
            p.add_run("Sumário estruturado por grupo, produto e documento")

            p = doc.add_paragraph(style="DossieSubtitle")
            p.add_run(manifest["month_label"])

        for line in page_lines:
            p = doc.add_paragraph(style=f"DossieTOC{line['level']}")
            p.paragraph_format.tab_stops.clear_all()
            p.paragraph_format.tab_stops.add_tab_stop(Cm(16.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            p.add_run(line["label"])
            p.add_run("	")
            p.add_run(str(page_map.get(line["id"], "")))
        first = False


def write_clickable_manual_toc(doc: Document, manifest: dict[str, Any], page_map: dict[str, int]) -> None:
    toc_pages = paginate_toc_entries_docx(build_toc_lines(manifest))
    entry_by_id = {entry["id"]: entry for entry in manifest["entries"]}
    first = True
    for page_idx, page_lines in enumerate(toc_pages):
        if not first:
            p = doc.add_paragraph()
            p.add_run().add_break(WD_BREAK.PAGE)
        if page_idx == 0:
            p = doc.add_paragraph(style="DossieMainTitle")
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.add_run("Dossiê Mensal")

            p = doc.add_paragraph(style="DossieSubtitle")
            p.add_run("Sumário estruturado por grupo, produto e documento")

            p = doc.add_paragraph(style="DossieSubtitle")
            p.add_run(manifest["month_label"])

        for line in page_lines:
            level = int(line["level"])
            color = "22303A" if level == 1 else "3B4550"
            p = doc.add_paragraph(style=f"DossieTOC{level}")
            p.paragraph_format.tab_stops.clear_all()
            p.paragraph_format.tab_stops.add_tab_stop(Cm(16.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            add_internal_hyperlink(p, line["label"], bookmark_name_for_entry(line["id"]), color=color, underline=False)
            p.add_run("\t")
            p.add_run(str(page_map.get(line["id"], "")))
        first = False


def write_native_toc(doc: Document, manifest: dict[str, Any]) -> None:
    p = doc.add_paragraph(style="DossieMainTitle")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.add_run("Dossiê Mensal")

    p = doc.add_paragraph(style="DossieSubtitle")
    p.add_run("Sumário estruturado por grupo, produto e documento")

    p = doc.add_paragraph(style="DossieSubtitle")
    p.add_run(manifest["month_label"])

    p = doc.add_paragraph()
    begin = p.add_run()
    _append_field_char(begin, "begin")
    instr = p.add_run()
    _append_instr_text(instr, f'TOC \\f {TOC_TC_IDENTIFIER} \\h \\z')
    separate = p.add_run()
    _append_field_char(separate, "separate")
    p.add_run("Atualizando sumário...")
    end = p.add_run()
    _append_field_char(end, "end")

    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


def build_review_docx(manifest: dict[str, Any], docx_path: Path) -> None:
    doc = Document()
    configure_docx_styles(doc)
    configure_standard_section(doc.sections[0])
    set_section_vertical_alignment(doc.sections[0], "top")

    include_initial_cover = bool(find_cover_template("initial"))

    with tempfile.TemporaryDirectory(prefix="prog_review_cover_") as tmp:
        temp_dir = Path(tmp)
        cache: dict[str, CoverRender] = {}
        if include_initial_cover:
            cover = render_cover_image("initial", manifest["month_cover_label"], temp_dir, cache)
            if cover:
                insert_cover_page(doc, cover, first_page=True)
                section = doc.add_section(WD_SECTION_START.NEW_PAGE)
                configure_standard_section(section)
                set_section_vertical_alignment(section, "top")

        _, page_map = calculate_page_map(manifest, include_initial_cover=include_initial_cover)
        write_manual_toc(doc, manifest, page_map)

        p = doc.add_paragraph()
        p.add_run().add_break(WD_BREAK.PAGE)

        p = doc.add_paragraph(style="DossieSectionTitle")
        p.add_run("Títulos para revisão manual")

        p = doc.add_paragraph(style="DossieHint")
        p.add_run(
            "O Word de revisão é apenas visual. A revisão que manda no dossiê final é a planilha Excel de conferência. Para excluir itens do dossiê final, apague as linhas correspondentes na aba 'Conferencia' do Excel. O script final vai reconstruir o dossiê completo usando apenas as linhas que permanecerem no Excel e vai absorver PDFs novos colocados nas pastas de revisão."
        )

        table = doc.add_table(rows=1, cols=4)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        for idx, name in enumerate(REVIEW_TABLE_HEADERS):
            hdr[idx].text = name
            _set_cell_shading(hdr[idx], "1F4E5F")
            for p_hdr in hdr[idx].paragraphs:
                for run in p_hdr.runs:
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.font.bold = True

        for entry in manifest["entries"]:
            row = table.add_row().cells
            row[0].text = entry["id"]
            row[1].text = entry["kind"]
            row[2].text = entry["product_code"] or entry["group_code"]
            row[3].text = entry["title"]

        widths = [Cm(2.0), Cm(2.5), Cm(3.2), Cm(11.5)]
        for row in table.rows:
            for idx, width in enumerate(widths):
                row.cells[idx].width = width

    ensure_dir(docx_path.parent)
    safe_unlink(docx_path)
    doc.save(str(docx_path))


def read_review_table_rows(doc: Document) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    expected_headers = [normalize_text(h) for h in REVIEW_TABLE_HEADERS]
    for table in doc.tables:
        if not table.rows:
            continue
        headers = [normalize_text(cell.text) for cell in table.rows[0].cells[:4]]
        if len(headers) >= 4 and headers[:4] == expected_headers:
            for row in table.rows[1:]:
                cells = row.cells
                if len(cells) < 4:
                    continue
                entry_id = cells[0].text.strip()
                kind = cells[1].text.strip().lower()
                code = cells[2].text.strip()
                title = cells[3].text.strip()
                if entry_id:
                    rows.append({
                        "id": entry_id,
                        "kind": kind,
                        "code": code,
                        "title": title,
                    })
            break
    return rows


def try_read_review_table(doc: Document) -> dict[str, str]:
    return {
        row["id"]: row["title"]
        for row in read_review_table_rows(doc)
        if row.get("title", "").strip()
    }


def _entry_display_suffix(title: str) -> str:
    return title.split("–", 1)[-1].strip() if "–" in title else title.strip()



def _filter_missing_rows_by_manifest(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    valid_products = {
        product["product_code"]
        for group in manifest["groups"]
        for product in group["products"]
    }
    valid_groups = {group["group_code"] for group in manifest["groups"]}
    kept_docs = [
        entry for entry in manifest["entries"]
        if entry.get("kind") == "document"
    ]

    filtered: list[dict[str, Any]] = []
    for row in manifest.get("missing_rows", []):
        product_code = str(row.get("Produto", "")).strip()
        group_code = str(row.get("Grupo", "")).strip()

        if product_code:
            if product_code not in valid_products:
                continue
        else:
            if group_code not in valid_groups:
                continue

        row_doc = str(row.get("Documento", "")).strip()
        if not row_doc:
            filtered.append(row)
            continue

        matched = False
        for entry in kept_docs:
            if product_code:
                if entry.get("product_code") != product_code:
                    continue
            else:
                if str(entry.get("group_code", "")).strip() != group_code or str(entry.get("product_code", "")).strip():
                    continue
            doc_only = _entry_display_suffix(str(entry.get("title", "")))
            if similarity(doc_only, row_doc) >= 0.60:
                matched = True
                break

        if matched:
            filtered.append(row)

    return filtered


def update_manifest_titles_from_docx(docx_path: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    manifest = deepcopy(manifest)
    if not docx_path.exists():
        return manifest

    doc = Document(str(docx_path))
    review_rows = read_review_table_rows(doc)
    if not review_rows:
        return manifest

    active_ids = {row["id"] for row in review_rows}
    title_updates = {
        row["id"]: row["title"].strip()
        for row in review_rows
        if row.get("title", "").strip()
    }

    original_entries = {entry["id"]: deepcopy(entry) for entry in manifest["entries"]}
    for entry_id, new_title in title_updates.items():
        if entry_id in original_entries:
            original_entries[entry_id]["title"] = new_title

    new_groups: list[dict[str, Any]] = []
    new_entries: list[dict[str, Any]] = []

    for group in manifest["groups"]:
        if group["id"] not in active_ids:
            continue

        group_copy = deepcopy(group)
        group_entry = deepcopy(original_entries[group["id"]])
        group_copy["group_title"] = _entry_display_suffix(group_entry["title"])
        new_entries.append(group_entry)

        kept_products: list[dict[str, Any]] = []
        for product in group["products"]:
            if product["id"] not in active_ids:
                continue

            product_copy = deepcopy(product)
            product_entry = deepcopy(original_entries[product["id"]])
            product_copy["product_title"] = _entry_display_suffix(product_entry["title"])
            new_entries.append(product_entry)

            kept_documents: list[dict[str, Any]] = []
            for doc_item in product.get("documents", []):
                if doc_item["id"] not in active_ids:
                    continue
                doc_entry = deepcopy(original_entries[doc_item["id"]])
                doc_copy = deepcopy(doc_item)
                doc_copy.update(doc_entry)
                kept_documents.append(doc_copy)
                new_entries.append(doc_entry)

            product_copy["documents"] = kept_documents
            kept_products.append(product_copy)

        group_copy["products"] = kept_products
        new_groups.append(group_copy)

    manifest["groups"] = new_groups
    manifest["entries"] = new_entries
    manifest["missing_rows"] = _filter_missing_rows_by_manifest(manifest)
    return manifest



def read_review_excel_rows(xlsx_path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    if not Path(xlsx_path).exists():
        return rows

    try:
        wb = load_workbook(xlsx_path, data_only=True, read_only=True)
    except Exception:
        return rows

    ws = wb["Conferencia"] if "Conferencia" in wb.sheetnames else wb[wb.sheetnames[0]]
    header_row = None
    headers: dict[int, str] = {}
    expected = {normalize_text(h) for h in ["Grupo", "Produto", "Documento", "Situação"]}

    for row_idx in range(1, min(ws.max_row, 10) + 1):
        current = {}
        found = set()
        for col_idx in range(1, min(ws.max_column, 20) + 1):
            value = ws.cell(row_idx, col_idx).value
            label = normalize_text(str(value)) if value is not None else ""
            if label:
                current[col_idx] = label
                if label in expected:
                    found.add(label)
        if len(found) >= 4:
            header_row = row_idx
            headers = current
            break

    if not header_row:
        return rows

    col_group = next((c for c, h in headers.items() if h == normalize_text("Grupo")), None)
    col_product = next((c for c, h in headers.items() if h == normalize_text("Produto")), None)
    col_document = next((c for c, h in headers.items() if h == normalize_text("Documento")), None)
    col_status = next((c for c, h in headers.items() if h == normalize_text("Situação")), None)

    for row_idx in range(header_row + 1, ws.max_row + 1):
        group_text = str(ws.cell(row_idx, col_group).value or "").strip() if col_group else ""
        product_value = ws.cell(row_idx, col_product).value if col_product else None
        product_text = str(product_value or "").strip()
        document_text = str(ws.cell(row_idx, col_document).value or "").strip() if col_document else ""
        status_text = str(ws.cell(row_idx, col_status).value or "").strip() if col_status else ""

        if not group_text and not product_text and not document_text:
            continue

        rows.append({
            "group_code": short_group_code(group_text) if group_text else "",
            "product_code": parse_product_code(product_text) or product_text,
            "document_title": document_text,
            "status": status_text,
        })

    return rows



def update_manifest_from_review_excel(xlsx_path: Path, manifest: dict[str, Any]) -> dict[str, Any]:
    manifest = deepcopy(manifest)
    review_rows = read_review_excel_rows(xlsx_path)
    if not review_rows:
        return manifest

    rows_by_product: dict[str, list[dict[str, str]]] = {}
    rows_by_group_docs: dict[str, list[dict[str, str]]] = {}
    for row in review_rows:
        product_code = parse_product_code(str(row.get("product_code", ""))) or str(row.get("product_code", "")).strip()
        group_code = str(row.get("group_code", "")).strip()
        row["product_code"] = product_code
        row["group_code"] = group_code
        if product_code:
            rows_by_product.setdefault(product_code, []).append(row)
        elif group_code:
            rows_by_group_docs.setdefault(group_code, []).append(row)

    original_entries = {entry["id"]: deepcopy(entry) for entry in manifest["entries"]}
    new_groups: list[dict[str, Any]] = []
    new_entries: list[dict[str, Any]] = []

    for group in manifest["groups"]:
        kept_products: list[tuple[dict[str, Any], dict[str, Any], list[dict[str, Any]]]] = []
        kept_group_docs: list[dict[str, Any]] = []
        used_group_doc_ids: set[str] = set()

        for row in [r for r in rows_by_group_docs.get(group["group_code"], []) if str(r.get("document_title", "")).strip()]:
            row_doc_title = str(row.get("document_title", "")).strip()
            best_doc = None
            best_score = 0.0
            for doc_item in group.get("documents", []):
                if doc_item["id"] in used_group_doc_ids:
                    continue
                doc_only = _entry_display_suffix(str(doc_item.get("title", "")))
                score = similarity(doc_only, row_doc_title)
                if normalize_text(doc_only) == normalize_text(row_doc_title):
                    score = 1.0
                if score > best_score:
                    best_score = score
                    best_doc = doc_item
            if best_doc is not None and best_score >= 0.45:
                doc_copy = deepcopy(original_entries[best_doc["id"]])
                doc_copy["title"] = row_doc_title
                doc_copy["product_code"] = ""
                kept_group_docs.append(doc_copy)
                used_group_doc_ids.add(best_doc["id"])

        for product in group["products"]:
            product_code = product["product_code"]
            product_rows = rows_by_product.get(product_code, [])
            if not product_rows:
                continue

            matched_docs: list[dict[str, Any]] = []
            used_doc_ids: set[str] = set()
            nonempty_rows = [row for row in product_rows if str(row.get("document_title", "")).strip()]

            for row in nonempty_rows:
                row_doc_title = str(row.get("document_title", "")).strip()
                best_doc = None
                best_score = 0.0
                for doc_item in product.get("documents", []):
                    if doc_item["id"] in used_doc_ids:
                        continue
                    doc_only = _entry_display_suffix(str(doc_item.get("title", "")))
                    score = similarity(doc_only, row_doc_title)
                    if normalize_text(doc_only) == normalize_text(row_doc_title):
                        score = 1.0
                    if score > best_score:
                        best_score = score
                        best_doc = doc_item

                if best_doc is not None and best_score >= 0.45:
                    doc_copy = deepcopy(original_entries[best_doc["id"]])
                    doc_copy["title"] = f"{product_code} – {row_doc_title}"
                    matched_docs.append(doc_copy)
                    used_doc_ids.add(best_doc["id"])

            keep_empty_product = any(not str(row.get("document_title", "")).strip() for row in product_rows)
            if not matched_docs and not keep_empty_product:
                continue

            product_copy = deepcopy(product)
            product_entry = deepcopy(original_entries[product["id"]])
            product_copy["documents"] = matched_docs
            kept_products.append((product_copy, product_entry, matched_docs))

        if not kept_products and not kept_group_docs:
            continue

        group_copy = deepcopy(group)
        group_entry = deepcopy(original_entries[group["id"]])
        group_copy["products"] = []
        group_copy["documents"] = kept_group_docs
        new_groups.append(group_copy)
        new_entries.append(group_entry)

        for doc_entry in kept_group_docs:
            new_entries.append(doc_entry)

        for product_copy, product_entry, matched_docs in kept_products:
            group_copy["products"].append(product_copy)
            new_entries.append(product_entry)
            for doc_entry in matched_docs:
                new_entries.append(doc_entry)

    manifest["groups"] = new_groups
    manifest["entries"] = new_entries
    manifest["missing_rows"] = _filter_missing_rows_by_manifest(manifest)
    manifest = rebuild_manifest_entry_order(manifest)
    return manifest


def select_review_docx(paths: dict[str, str]) -> Path:
    annex_review_dir = Path(paths["annex_review_dir"])
    preferred = Path(paths["base_docx"])
    docx_files = sorted(
        [p for p in annex_review_dir.glob("*.docx") if normalize_text(p.name) != normalize_text(Path(paths["final_docx"]).name)],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if docx_files:
        return docx_files[0]
    return preferred


# ---------------------------------------------------------------------------
# Revisão manual de PDFs na pasta Anexos - Revisão
# ---------------------------------------------------------------------------


def _all_group_dirs(manifest: dict[str, Any]) -> list[tuple[dict[str, Any], Path]]:
    result: list[tuple[dict[str, Any], Path]] = []
    for group in manifest["groups"]:
        folder = Path(group.get("folder_path") or (Path(manifest["paths"]["annex_review_dir"]) / group_folder_name(group["group_code"], group["group_title"])))
        ensure_dir(folder)
        group["folder_path"] = str(folder)
        result.append((group, folder))
    return result

def match_pdf_to_group_document(pdf_path: Path, group: dict[str, Any]) -> Optional[dict[str, Any]]:
    candidates = [d for d in group.get("documents", [])]
    if not candidates:
        return None

    pdf_title = clean_display_title(pdf_path.name, "")
    best = None
    best_score = 0.0
    for entry in candidates:
        doc_only = _entry_display_suffix(str(entry.get("title", "")))
        score = similarity(pdf_title, doc_only)
        if entry.get("status") != "pdf_localizado":
            score += 0.10
        if score > best_score:
            best = entry
            best_score = score

    if best_score >= 0.36:
        return best
    if len(candidates) == 1:
        return candidates[0]
    return None

def _all_product_dirs(manifest: dict[str, Any]) -> list[tuple[dict[str, Any], dict[str, Any], Path]]:
    result: list[tuple[dict[str, Any], dict[str, Any], Path]] = []
    for group in manifest["groups"]:
        for product in group["products"]:
            folder = Path(product.get("folder_path") or product_dest_dir(Path(manifest["paths"]["annex_review_dir"]), group["group_code"], group["group_title"], product["product_code"], product["product_title"]))
            ensure_dir(folder)
            product["folder_path"] = str(folder)
            result.append((group, product, folder))
    return result

def _find_product_by_folder(manifest: dict[str, Any], folder: Path) -> tuple[Optional[dict[str, Any]], Optional[dict[str, Any]]]:
    folder_resolved = folder.resolve()
    for group, product, prod_dir in _all_product_dirs(manifest):
        try:
            if prod_dir.resolve() == folder_resolved:
                return group, product
        except Exception:
            continue
    return None, None

def match_pdf_to_document(pdf_path: Path, product: dict[str, Any]) -> Optional[dict[str, Any]]:
    candidates = [d for d in product.get("documents", [])]
    if not candidates:
        return None

    normalized_name = clean_display_title(pdf_path.name, product.get("product_code", ""))
    pdf_title = normalized_name or clean_display_title(pdf_path.stem, product.get("product_code", ""))

    best = None
    best_score = 0.0
    for entry in candidates:
        doc_only = entry["title"].split("–", 1)[-1].strip() if "–" in entry["title"] else entry["title"]
        score = similarity(pdf_title, doc_only)
        if entry.get("status") != "pdf_localizado":
            score += 0.10
        if score > best_score:
            best = entry
            best_score = score

    if best_score >= 0.36:
        return best
    if len(candidates) == 1:
        return candidates[0]
    return None


def apply_review_pdfs(manifest: dict[str, Any]) -> dict[str, Any]:
    manifest = deepcopy(manifest)
    annex_review_dir = Path(manifest["paths"]["annex_review_dir"])
    _all_group_dirs(manifest)
    _all_product_dirs(manifest)

    ignored_names = {
        normalize_text(Path(manifest["paths"]["final_pdf"]).name),
        normalize_text(Path(manifest["paths"]["base_pdf"]).name),
    }

    for group, group_dir in _all_group_dirs(manifest):
        direct_pdfs = [p for p in group_dir.glob("*.pdf") if normalize_text(p.name) not in ignored_names]
        matched_entries: set[str] = set()

        for pdf_path in direct_pdfs:
            entry = match_pdf_to_group_document(pdf_path, group)
            if entry is None:
                entry = Entry(
                    id=_next_manifest_entry_id(manifest, "D"),
                    kind="document",
                    title=clean_display_title(pdf_path.name, ""),
                    group_code=group["group_code"],
                    product_code="",
                    status="pdf_localizado",
                    notes="",
                    source_path=str(pdf_path),
                    copied_path=str(pdf_path),
                    ext=".pdf",
                    page_count=safe_pdf_page_count(pdf_path),
                    source_kind="review_group_pdf",
                    review_source_path=str(pdf_path),
                    original_name=pdf_path.name,
                ).to_dict()
                group.setdefault("documents", []).append(entry)
                manifest.setdefault("entries", []).append(entry)
            else:
                entry["source_path"] = entry.get("source_path") or str(pdf_path)
                entry["copied_path"] = str(pdf_path)
                entry["review_source_path"] = str(pdf_path)
                entry["ext"] = ".pdf"
                entry["page_count"] = safe_pdf_page_count(pdf_path)
                entry["status"] = "pdf_localizado"
                entry["source_kind"] = "review_group_pdf"
                entry["notes"] = ""
                entry["original_name"] = pdf_path.name
            matched_entries.add(entry["id"])

        existing_files_norm = {normalize_text(p.name) for p in direct_pdfs}
        for entry in group.get("documents", []):
            if entry["id"] not in matched_entries and entry.get("copied_path"):
                current = Path(entry["copied_path"])
                if current.exists() and normalize_text(current.name) in existing_files_norm:
                    entry["page_count"] = safe_pdf_page_count(current)
                    entry["status"] = "pdf_localizado"

    for group, product, prod_dir in _all_product_dirs(manifest):
        pdfs = [p for p in prod_dir.glob("*.pdf") if normalize_text(p.name) not in ignored_names]
        matched_entries: set[str] = set()
        for pdf_path in pdfs:
            entry = match_pdf_to_document(pdf_path, product)
            if not entry:
                continue
            entry["source_path"] = entry.get("source_path") or str(pdf_path)
            entry["copied_path"] = str(pdf_path)
            entry["review_source_path"] = str(pdf_path)
            entry["ext"] = ".pdf"
            entry["page_count"] = safe_pdf_page_count(pdf_path)
            entry["status"] = "pdf_localizado"
            entry["source_kind"] = "review_pdf" if str(pdf_path) != entry.get("source_path", "") else entry.get("source_kind", "review_pdf")
            entry["notes"] = ""
            entry["original_name"] = pdf_path.name
            matched_entries.add(entry["id"])

        existing_files_norm = {normalize_text(p.name) for p in pdfs}
        for entry in product.get("documents", []):
            if entry["id"] not in matched_entries and entry.get("copied_path"):
                current = Path(entry["copied_path"])
                if current.exists() and normalize_text(current.name) in existing_files_norm:
                    entry["page_count"] = safe_pdf_page_count(current)
                    entry["status"] = "pdf_localizado"

    for group in manifest["groups"]:
        for doc_item in group.get("documents", []):
            for entry in manifest["entries"]:
                if entry["id"] == doc_item["id"]:
                    entry.update(doc_item)
                    break
        for product in group["products"]:
            for doc_item in product["documents"]:
                for entry in manifest["entries"]:
                    if entry["id"] == doc_item["id"]:
                        entry.update(doc_item)
                        break

    manifest = rebuild_manifest_entry_order(manifest)
    return manifest


# ---------------------------------------------------------------------------
# Renderização de PDF e composição final
# ---------------------------------------------------------------------------

def render_pdf_pages_to_images(
    pdf_path: Path,
    out_dir: Path,
    prefix: str = "page",
    scale: float = DOCX_ATTACHMENT_RENDER_SCALE,
) -> list[tuple[Path, float, float]]:
    ensure_dir(out_dir)
    images: list[tuple[Path, float, float]] = []
    try:
        doc = fitz.open(str(pdf_path))
        for idx in range(doc.page_count):
            page = doc[idx]
            rect = page.rect
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            out = out_dir / f"{prefix}_{idx + 1:04d}.png"
            pix.save(str(out))
            images.append((out, float(rect.width), float(rect.height)))
        doc.close()
    except Exception:
        return []
    return images

def add_fullpage_image(
    doc: Document,
    image_path: Path,
    width_pt: float,
    height_pt: float,
    first_page: bool = False,
    bookmark_name: str | None = None,
) -> None:
    if first_page:
        section = doc.sections[0]
    else:
        section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    configure_zero_margin_section(section, width_pt, height_pt)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run()
    run.add_picture(str(image_path), width=Pt(width_pt), height=Pt(height_pt))
    if bookmark_name:
        add_bookmark(p, bookmark_name)


def _entry_cover_value(entry: dict[str, Any]) -> str:
    if entry["kind"] == "document":
        return entry["title"].split("–", 1)[-1].strip() if "–" in entry["title"] else entry["title"]
    return entry["title"]

def _cover_page_indices_from_manifest(manifest: dict[str, Any]) -> list[int]:
    indices: list[int] = []
    for entry in manifest["entries"]:
        base_page = int(entry.get("base_page") or 0)
        if base_page > 0:
            indices.append(base_page - 1)
    return indices


def _safe_outline_title(value: str) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    if len(text) > PDF_OUTLINE_TITLE_MAX:
        text = text[:PDF_OUTLINE_TITLE_MAX].rstrip()
    return text or "Seção"


def _append_cover_page_from_preview(
    out_doc: fitz.Document,
    preview_doc: fitz.Document,
    page_index: int,
) -> None:
    if 0 <= page_index < preview_doc.page_count:
        out_doc.insert_pdf(preview_doc, from_page=page_index, to_page=page_index)


def _append_original_pdf(
    out_doc: fitz.Document,
    pdf_path: Path,
) -> None:
    src_doc = fitz.open(str(pdf_path))
    try:
        out_doc.insert_pdf(src_doc)
    finally:
        src_doc.close()


def _set_pdf_outline_from_manifest(out_doc: fitz.Document, manifest: dict[str, Any]) -> None:
    toc: list[list[int | str]] = []
    for entry in manifest.get("entries", []):
        kind = entry.get("kind")
        if kind == "group":
            level = 1
        elif kind == "product":
            level = 2
        elif kind == "document":
            level = 2 if not str(entry.get("product_code", "")).strip() else 3
        else:
            level = None
        if level is None:
            continue
        page_num = int(entry.get("base_page") or 0)
        if page_num <= 0:
            continue
        toc.append([level, _safe_outline_title(entry.get("title", "")), page_num])
    if toc:
        try:
            out_doc.set_toc(toc)
        except Exception:
            pass


def _append_fullpage_pdf_image(
    out_doc: fitz.Document,
    image_path: Path,
    width_pt: float,
    height_pt: float,
) -> None:
    page = out_doc.new_page(width=float(width_pt), height=float(height_pt))
    page.insert_image(page.rect, filename=str(image_path), keep_proportion=False, overlay=True)


def _wrap_text_for_pdf_toc(text: str, max_chars: int) -> list[str]:
    words = re.sub(r"\s+", " ", str(text or "")).strip().split()
    if not words:
        return [""]
    lines = [words[0]]
    for word in words[1:]:
        trial = lines[-1] + " " + word
        if len(trial) <= max_chars:
            lines[-1] = trial
        else:
            lines.append(word)
    return lines


def _append_clickable_toc_pages_to_pdf(
    out_doc: fitz.Document,
    manifest: dict[str, Any],
    page_map: dict[str, int],
) -> None:
    toc_pages = paginate_toc_entries_docx(build_toc_lines(manifest))
    for page_idx, page_lines in enumerate(toc_pages):
        page = out_doc.new_page(width=PDF_TOC_PAGE_WIDTH, height=PDF_TOC_PAGE_HEIGHT)
        if page_idx == 0:
            page.insert_text((62, 88), "Dossiê Mensal", fontsize=21, fontname="helvB", color=(31/255, 78/255, 95/255))
            page.insert_text((62, 118), "Sumário estruturado por grupo, produto e documento", fontsize=11, fontname="helv", color=(86/255, 94/255, 102/255))
            page.insert_text((62, 142), str(manifest.get("month_label", "")), fontsize=11, fontname="helv", color=(86/255, 94/255, 102/255))
            y = 182.0
        else:
            y = 78.0

        for line in page_lines:
            level = int(line["level"])
            label = str(line["label"])
            target_page = max(0, int(page_map.get(line["id"], 1)) - 1)

            font_size = 11.0 if level == 1 else 10.0 if level == 2 else 9.2
            font_name = "helvB" if level == 1 else "helv"
            color = (34/255, 48/255, 58/255) if level == 1 else (59/255, 69/255, 80/255)
            indent = 0.0 if level == 1 else 20.0 if level == 2 else 40.0
            max_chars = 80 if level == 1 else 92 if level == 2 else 108
            wrapped = _wrap_text_for_pdf_toc(label, max_chars)
            line_height = font_size * 1.48
            block_h = max(line_height, line_height * len(wrapped))

            label_rect = fitz.Rect(62 + indent, y, 482, y + block_h)
            page.insert_textbox(label_rect, "\n".join(wrapped), fontsize=font_size, fontname=font_name, color=color, align=0)

            page_num_rect = fitz.Rect(500, y, 540, y + block_h)
            page.insert_textbox(page_num_rect, str(page_map.get(line["id"], "")), fontsize=font_size, fontname="helv", color=color, align=2)

            leader_y = y + min(block_h - 3, font_size + 1.5)
            page.draw_line(fitz.Point(330, leader_y), fitz.Point(495, leader_y), color=(0.72, 0.74, 0.76), width=0.5, dashes="[1 2]")

            page.insert_link({
                "kind": fitz.LINK_GOTO,
                "from": fitz.Rect(62 + indent, y, 540, y + block_h),
                "page": target_page,
            })

            y += block_h + (6.0 if level == 1 else 4.0)


def build_final_pdf_high_quality(
    manifest: dict[str, Any],
    final_docx: Path,
    final_pdf: Path,
    progress: Optional[Callable[[str], None]] = None,
    word_session: Optional[WordPdfSession] = None,
) -> Path:
    include_initial_cover = bool(find_cover_template("initial"))
    _, page_map = calculate_page_map(manifest, include_initial_cover=include_initial_cover)

    with tempfile.TemporaryDirectory(prefix="prog_pdf_hq_") as tmp:
        temp_dir = Path(tmp)
        cover_cache: dict[str, CoverRender] = {}
        out_doc = fitz.open()
        try:
            if include_initial_cover:
                _emit_progress(progress, "Inserindo capa inicial em alta qualidade...")
                initial_cover = render_cover_image(
                    "initial",
                    manifest["month_cover_label"],
                    temp_dir / "covers",
                    cover_cache,
                    word_session=word_session,
                    progress=progress,
                )
                _append_fullpage_pdf_image(out_doc, initial_cover.image_path, initial_cover.width_pt, initial_cover.height_pt)

            _emit_progress(progress, "Gerando páginas de sumário clicável no PDF final...")
            _append_clickable_toc_pages_to_pdf(out_doc, manifest, page_map)

            _emit_progress(progress, "Montando PDF final em alta qualidade com capas sem borda e anexos originais...")
            for entry in manifest.get("entries", []):
                template_kind = "document" if entry.get("kind") == "document" else entry.get("kind")
                cover = render_cover_image(
                    str(template_kind),
                    _entry_cover_value(entry),
                    temp_dir / "covers",
                    cover_cache,
                    word_session=word_session,
                    progress=progress,
                )
                _append_fullpage_pdf_image(out_doc, cover.image_path, cover.width_pt, cover.height_pt)

                if entry.get("kind") == "document" and entry.get("status") == "pdf_localizado":
                    src = entry.get("copied_path") or entry.get("review_source_path") or entry.get("source_path")
                    if src and Path(src).exists():
                        _append_original_pdf(out_doc, Path(src))

            _set_pdf_outline_from_manifest(out_doc, manifest)
            safe_unlink(final_pdf)
            ensure_dir(final_pdf.parent)
            out_doc.save(str(final_pdf), garbage=4, deflate=True)
        finally:
            out_doc.close()

    return final_pdf


def build_final_docx(
    manifest: dict[str, Any],
    final_docx: Path,
    progress: Optional[Callable[[str], None]] = None,
    word_session: Optional[WordPdfSession] = None,
) -> Path:
    started = time.perf_counter()
    doc = Document()
    configure_docx_styles(doc)
    configure_standard_section(doc.sections[0])
    set_section_vertical_alignment(doc.sections[0], "top")

    include_initial_cover = bool(find_cover_template("initial"))
    _, page_map = calculate_page_map(manifest, include_initial_cover=include_initial_cover)
    total_entries = len(manifest["entries"])

    _emit_progress(progress, f"Montando DOCX final ({total_entries} seções base)...")

    with tempfile.TemporaryDirectory(prefix="prog_final_build_") as tmp:
        temp_dir = Path(tmp)
        cover_cache: dict[str, CoverRender] = {}

        if include_initial_cover:
            _emit_progress(progress, "Renderizando capa inicial...")
            cover = render_cover_image(
                "initial",
                manifest["month_cover_label"],
                temp_dir / "covers",
                cover_cache,
                word_session=word_session,
                progress=progress,
            )
            if cover:
                add_fullpage_image(doc, cover.image_path, cover.width_pt, cover.height_pt, first_page=True)
                section = doc.add_section(WD_SECTION_START.NEW_PAGE)
                configure_standard_section(section)
                set_section_vertical_alignment(section, "top")

        _emit_progress(progress, "Escrevendo sumário clicável no DOCX final...")
        write_clickable_manual_toc(doc, manifest, page_map)

        for idx, entry in enumerate(manifest["entries"], start=1):
            _emit_progress(progress, f"[{idx}/{total_entries}] Processando {entry['kind']}: {entry['title'][:90]}")

            kind = entry["kind"]
            template_kind = "document" if kind == "document" else kind
            cover = render_cover_image(
                template_kind,
                _entry_cover_value(entry),
                temp_dir / "covers",
                cover_cache,
                word_session=word_session,
                progress=progress,
            )
            if cover:
                add_fullpage_image(
                    doc,
                    cover.image_path,
                    cover.width_pt,
                    cover.height_pt,
                    first_page=False,
                    bookmark_name=bookmark_name_for_entry(entry["id"]),
                )
            else:
                fallback_sec = doc.add_section(WD_SECTION_START.NEW_PAGE)
                configure_standard_section(fallback_sec)
                set_section_vertical_alignment(fallback_sec, "top")
                p = doc.add_paragraph(style="DossieSectionTitle")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run(entry["title"])
                add_bookmark(p, bookmark_name_for_entry(entry["id"]))

            if entry["kind"] == "document" and entry.get("status") == "pdf_localizado":
                src = entry.get("copied_path") or entry.get("review_source_path") or entry.get("source_path")
                if src and Path(src).exists():
                    page_count = int(entry.get("page_count") or 0)
                    if page_count:
                        pass # Silenciado para não poluir o terminal, ele faz rapidinho
                    rendered_pages = render_pdf_pages_to_images(
                        Path(src),
                        temp_dir / "pdf_pages" / entry["id"],
                        prefix=entry["id"],
                        scale=DOCX_ATTACHMENT_RENDER_SCALE,
                    )
                    for img_path, width_pt, height_pt in rendered_pages:
                        add_fullpage_image(doc, img_path, width_pt, height_pt, first_page=False)

        ensure_dir(final_docx.parent)
        safe_unlink(final_docx)
        _emit_progress(progress, "Salvando DOCX final...")
        doc.save(str(final_docx))

    elapsed = time.perf_counter() - started
    _emit_progress(progress, f"DOCX final concluído em {elapsed:.1f}s.")
    return final_docx

# ---------------------------------------------------------------------------
# Conversão DOCX -> PDF A PROVA DE FALHAS (WORD VISÍVEL)
# ---------------------------------------------------------------------------

def _emit_progress(progress: Optional[Callable[[str], None]], message: str) -> None:
    if progress:
        progress(message)


class WordPdfSession:
    def __init__(self):
        self.word = None
        self._pythoncom = None

    def __enter__(self):
        try:
            import pythoncom
            import win32com.client

            self._pythoncom = pythoncom
            pythoncom.CoInitialize()

            self.word = win32com.client.DispatchEx("Word.Application")
            self.word.Visible = SHOW_WORD_WINDOW
            self.word.DisplayAlerts = 0

            try:
                self.word.Options.SaveNormalPrompt = False
            except Exception:
                pass
        except Exception:
            self.word = None
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.word is not None:
            try:
                self.word.Quit(False)
            except Exception:
                pass
        self.word = None

        if self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
        self._pythoncom = None

    def convert(self, docx_path: Path, pdf_path: Path, update_fields: bool = False) -> bool:
        if self.word is None:
            return False

        doc = None
        try:
            docx_resolved = str(docx_path.resolve())
            pdf_resolved = str(pdf_path.resolve())

            doc = self.word.Documents.Open(
                docx_resolved,
                ReadOnly=not update_fields,
                AddToRecentFiles=False,
                ConfirmConversions=False,
                OpenAndRepair=True,
                NoEncodingDialog=True,
                Visible=False,
            )
            if update_fields:
                try:
                    doc.Repaginate()
                except Exception:
                    pass
                try:
                    for toc in doc.TablesOfContents:
                        toc.Update()
                except Exception:
                    pass
                try:
                    doc.Fields.Update()
                except Exception:
                    pass
                try:
                    doc.Save()
                except Exception:
                    pass

            doc.ExportAsFixedFormat(pdf_resolved, 17)
            doc.Close(False)

            return pdf_path.exists() and pdf_path.stat().st_size > 0
        except Exception:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            return False


def try_convert_docx_to_pdf(docx_path: Path, pdf_path: Path, word_session: Optional[WordPdfSession] = None, update_fields: bool = False) -> bool:
    docx_path = Path(docx_path)
    pdf_path = Path(pdf_path)

    safe_unlink(pdf_path)
    ensure_dir(pdf_path.parent)

    if word_session is not None and word_session.convert(docx_path, pdf_path, update_fields=update_fields):
        return True

    # Segunda tentativa com uma sessão nova do Word, isolada da sessão que montou as capas.
    try:
        with WordPdfSession() as fresh_word_session:
            if fresh_word_session.word is not None and fresh_word_session.convert(docx_path, pdf_path, update_fields=update_fields):
                return True
    except Exception:
        pass

    # Rota secundária: LibreOffice, com timeout maior porque o DOCX final pode ser grande.
    candidates = [
        shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if not candidate or not Path(candidate).exists():
            continue
        try:
            outdir = str(pdf_path.parent)
            subprocess.run(
                [str(candidate), "--headless", "--convert-to", "pdf", "--outdir", outdir, str(docx_path)],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=300,
            )
            generated = pdf_path.parent / f"{docx_path.stem}.pdf"
            if generated.exists() and generated.stat().st_size > 0:
                if generated != pdf_path:
                    safe_unlink(pdf_path)
                    generated.replace(pdf_path)
                return True
        except Exception:
            continue

    return False


def _unused_convert_docx_to_pdf_legacy(docx_path: Path, pdf_path: Path, word_session: Optional[WordPdfSession] = None) -> bool:
    docx_path = Path(docx_path)
    pdf_path = Path(pdf_path)

    safe_unlink(pdf_path)
    ensure_dir(pdf_path.parent)

    # Rota primária: Usa o Word Visível e super rápido que abrimos
    if word_session is not None and word_session.convert(docx_path, pdf_path):
        return True

    # Rota Secundária caso a pessoa nem tenha o Word instalado: LibreOffice
    candidates = [
        shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if not candidate or not Path(candidate).exists():
            continue
        try:
            outdir = str(pdf_path.parent)
            subprocess.run(
                [str(candidate), "--headless", "--convert-to", "pdf", "--outdir", outdir, str(docx_path)],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=15 # Se o libreoffice travar, desiste rápido
            )
            generated = pdf_path.parent / f"{docx_path.stem}.pdf"
            if generated.exists() and generated.stat().st_size > 0:
                if generated != pdf_path:
                    safe_unlink(pdf_path)
                    generated.replace(pdf_path)
                return True
        except Exception:
            continue

    return False

def add_page_numbers_transparent(pdf_path: Path, skip_first_n: int = 1) -> None:
    doc = fitz.open(str(pdf_path))
    for idx in range(doc.page_count):
        if idx < skip_first_n:
            continue
        page = doc[idx]
        rect = page.rect
        x0 = rect.x1 / 2 - 20
        x1 = rect.x1 / 2 + 20
        y0 = rect.y1 - 24
        y1 = rect.y1 - 8
        page.insert_textbox(
            fitz.Rect(x0, y0, x1, y1),
            str(idx + 1),
            fontsize=9,
            fontname="helv",
            color=(0.42, 0.42, 0.42),
            align=1,
            overlay=True,
        )
    doc.saveIncr()
    doc.close()


# ---------------------------------------------------------------------------
# Conferência XLSX
# ---------------------------------------------------------------------------


def build_conferencia_rows(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for entry in manifest["entries"]:
        if entry["kind"] == "document":
            status_label = {
                "pdf_localizado": "PDF localizado",
                "arquivo_nao_pdf": "Arquivo encontrado sem PDF",
                "controle_sem_pdf_localizado": "Registro do controle de entregas sem arquivo localizado",
                "report_only_sem_arquivo": "Item mantido no sumário sem arquivo localizado",
            }.get(entry.get("status", ""), entry.get("status", ""))
            rows.append({
                "Grupo": entry["group_code"],
                "Produto": entry["product_code"],
                "Documento": entry["title"].split("–", 1)[-1].strip() if "–" in entry["title"] else entry["title"],
                "Situação": status_label,
                "Origem": entry.get("source_kind", ""),
                "Caminho": entry.get("copied_path", "") or entry.get("source_path", ""),
                "Extensão": entry.get("ext", ""),
                "Página base": entry.get("base_page", ""),
                "Observações": entry.get("notes", ""),
            })

    for group in manifest["groups"]:
        if not group.get("documents") and not group.get("products"):
            rows.append({
                "Grupo": group["group_code"],
                "Produto": "",
                "Documento": "",
                "Situação": "Subseção sem nenhum documento localizado",
                "Origem": "sumário",
                "Caminho": "",
                "Extensão": "",
                "Página base": group.get("base_page", ""),
                "Observações": "Nenhum documento localizado diretamente para este grupo.",
            })
        for product in group["products"]:
            docs = product.get("documents", [])
            if not docs:
                rows.append({
                    "Grupo": group["group_code"],
                    "Produto": product["product_code"],
                    "Documento": "",
                    "Situação": "Subseção sem nenhum documento localizado",
                    "Origem": "sumário",
                    "Caminho": "",
                    "Extensão": "",
                    "Página base": product.get("base_page", ""),
                    "Observações": "Nenhum documento localizado para esta subseção.",
                })

    known = {(r["Grupo"], r["Produto"], normalize_text(r["Documento"]), r["Situação"]) for r in rows}
    for row in manifest.get("missing_rows", []):
        key = (row.get("Grupo", ""), row.get("Produto", ""), normalize_text(row.get("Documento", "")), row.get("Situação", ""))
        if key not in known:
            rows.append(row)
            known.add(key)

    rows.sort(key=lambda r: (r["Grupo"], product_sort_key(r["Produto"] or "P9.99"), normalize_text(r["Documento"] or "")))
    return rows


def write_levantamento_workbook(manifest: dict[str, Any], xlsx_path: Path) -> None:
    rows = build_conferencia_rows(manifest)

    wb = Workbook()
    ws = wb.active
    ws.title = "Conferencia"
    headers = ["Grupo", "Produto", "Documento", "Situação", "Origem", "Caminho", "Extensão", "Página base", "Observações"]
    ws.append(headers)

    fill_header = PatternFill("solid", fgColor="1F4E5F")
    font_header = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D7DEE3")
    fills = {
        "PDF localizado": PatternFill("solid", fgColor="E8F3EA"),
        "Arquivo encontrado sem PDF": PatternFill("solid", fgColor="FFF4E5"),
        "Registro do controle de entregas sem arquivo localizado": PatternFill("solid", fgColor="FDECEC"),
        "Item mantido no sumário sem arquivo localizado": PatternFill("solid", fgColor="FDECEC"),
        "Subseção sem nenhum documento localizado": PatternFill("solid", fgColor="FDECEC"),
    }

    for col_idx, _ in enumerate(headers, start=1):
        cell = ws.cell(1, col_idx)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)

    for row in rows:
        ws.append([row.get(h, "") for h in headers])

    for row in ws.iter_rows(min_row=2):
        status = row[3].value or ""
        fill = fills.get(status)
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin)
            if fill:
                cell.fill = fill

    widths = {"A": 10, "B": 11, "C": 46, "D": 38, "E": 20, "F": 60, "G": 10, "H": 11, "I": 58}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    resumo = wb.create_sheet("Resumo")
    resumo["A1"] = "Conferência do dossiê mensal"
    resumo["A1"].font = Font(size=14, bold=True, color="1F4E5F")
    resumo["A3"] = "Mês"
    resumo["B3"] = manifest["month_label"]
    resumo["A4"] = "Gerado em"
    resumo["B4"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    resumo["A6"] = "Totais por situação"
    rowi = 7
    counts: dict[str, int] = {}
    for row in rows:
        counts[row["Situação"]] = counts.get(row["Situação"], 0) + 1
    for status, count in counts.items():
        resumo[f"A{rowi}"] = status
        resumo[f"B{rowi}"] = count
        rowi += 1
    resumo.column_dimensions["A"].width = 65
    resumo.column_dimensions["B"].width = 15
    resumo.sheet_view.showGridLines = False

    ensure_dir(xlsx_path.parent)
    safe_unlink(xlsx_path)
    wb.save(str(xlsx_path))


# ---------------------------------------------------------------------------
# Persistência
# ---------------------------------------------------------------------------

def save_manifest(manifest: dict[str, Any], path: Path) -> None:
    ensure_dir(path.parent)
    path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

def load_manifest(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


# ---------------------------------------------------------------------------
# Orquestração
# ---------------------------------------------------------------------------

def prepare_month_base(
    month_root: Path,
    output_root: Path,
    clean_month_output: bool = True,
    progress: Optional[Callable[[str], None]] = None,
) -> tuple[dict[str, Any], dict[str, Path]]:
    _emit_progress(progress, "Lendo estrutura do mês e preparando Anexos - Revisão...")
    manifest, paths = build_month_manifest(month_root, output_root, clean_month_output=clean_month_output)
    assign_base_pages(manifest, include_initial_cover=bool(find_cover_template("initial")))
    _emit_progress(progress, "Gerando Word de revisão...")
    build_review_docx(manifest, paths["base_docx"])
    
    # Processamento Visível!
    with WordPdfSession() as word_session:
        _emit_progress(progress, "Convertendo Word de revisão para PDF...")
        try_convert_docx_to_pdf(paths["base_docx"], paths["base_pdf"], word_session=word_session)
        
    save_manifest(manifest, paths["manifest"])
    write_levantamento_workbook(manifest, paths["levantamento_xlsx"])
    _emit_progress(progress, "Fase 1 concluída.")
    return manifest, paths

def finalize_month(
    manifest_path: Path,
    use_review_folder: bool = False,
    progress: Optional[Callable[[str], None]] = None,
) -> tuple[dict[str, Any], Path]:
    started = time.perf_counter()
    _emit_progress(progress, f"Abrindo manifesto: {manifest_path}")
    manifest = load_manifest(Path(manifest_path))

    review_xlsx = Path(manifest["paths"]["levantamento_xlsx"])
    _emit_progress(progress, f"Lendo Excel de revisão: {review_xlsx.name}")
    manifest = update_manifest_from_review_excel(review_xlsx, manifest)

    if use_review_folder:
        _emit_progress(progress, "Varrendo PDFs adicionados manualmente em Anexos - Revisão...")
        manifest = apply_review_pdfs(manifest)

    _emit_progress(progress, "Recalculando páginas e atualizando conferência...")
    assign_base_pages(manifest, include_initial_cover=bool(find_cover_template("initial")))
    write_levantamento_workbook(manifest, Path(manifest["paths"]["levantamento_xlsx"]))
    save_manifest(manifest, Path(manifest["paths"]["manifest"]))

    final_docx = Path(manifest["paths"]["final_docx"])
    final_pdf = Path(manifest["paths"]["final_pdf"])

    # Primeiro: monta o Word final com melhor resolução para anexos.
    with WordPdfSession() as word_session:
        build_final_docx(manifest, final_docx, progress=progress, word_session=word_session)

    # Depois: gera o PDF final em alta qualidade, preservando os PDFs originais dos anexos.
    _emit_progress(progress, "Gerando PDF final em alta qualidade, preservando os anexos originais...")
    with WordPdfSession() as pdf_session:
        try:
            build_final_pdf_high_quality(manifest, final_docx, final_pdf, progress=progress, word_session=pdf_session)
        except Exception as exc:
            _emit_progress(progress, "Falha no modo de alta qualidade; tentando conversão padrão do DOCX final...")
            if not try_convert_docx_to_pdf(final_docx, final_pdf, word_session=pdf_session, update_fields=False):
                raise RuntimeError(
                    "Não foi possível gerar o PDF final em alta qualidade nem converter o DOCX final automaticamente. "
                    "Verifique se o Microsoft Word abre normalmente esse DOCX e se o LibreOffice está instalado."
                ) from exc

    _emit_progress(progress, "Aplicando numeração transparente no PDF final...")
    add_page_numbers_transparent(final_pdf, skip_first_n=1 if find_cover_template("initial") else 0)

    elapsed = time.perf_counter() - started
    _emit_progress(progress, f"Fase final concluída em {elapsed:.1f}s.")
    return manifest, final_pdf
