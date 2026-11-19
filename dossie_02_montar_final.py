from __future__ import annotations

import argparse
import re
from pathlib import Path

from dossie_utils_v2_andamento import DEFAULT_MONTH_ROOTS, DEFAULT_OUTPUT_ROOT, finalize_month, month_output_paths


def _strip_accents(text: str) -> str:
    import unicodedata
    text = unicodedata.normalize("NFKD", text or "")
    return "".join(ch for ch in text if not unicodedata.combining(ch))



def _norm(text: str) -> str:
    text = _strip_accents(text or "").lower()
    text = text.replace("_", " ").replace("-", " ")
    text = " ".join(text.split())
    return text



def _month_display(path: Path) -> str:
    return re.sub(r"^\d+\.\s*", "", path.name).strip()



def _month_sort_key(path: Path):
    m = re.match(r"^\s*(\d+)", path.name)
    num = int(m.group(1)) if m else 999
    return (num, _norm(_month_display(path)))



def _should_skip_month_dir(path: Path) -> bool:
    norm = _norm(_month_display(path))
    skip_tokens = ("modelo", "template", "capa", "old", "backup")
    return any(tok in norm for tok in skip_tokens)



def _discover_month_dirs() -> list[Path]:
    bases = []
    for root in DEFAULT_MONTH_ROOTS:
        p = Path(root)
        if p.parent.exists():
            bases.append(p.parent)

    month_dirs = []
    seen = set()
    for base in bases:
        for child in sorted(base.iterdir(), key=_month_sort_key):
            if not child.is_dir() or _should_skip_month_dir(child):
                continue
            key = _norm(_month_display(child))
            if not key or key in seen:
                continue
            if any(bad in key for bad in ('modelo', 'template', 'templates', 'capa', 'capas')):
                continue
            seen.add(key)
            month_dirs.append(child)

    return sorted(month_dirs, key=_month_sort_key)



def _prompt_months() -> list[Path]:
    month_dirs = _discover_month_dirs()
    if not month_dirs:
        raise SystemExit(
            "Nenhuma pasta de mês foi encontrada automaticamente. "
            "Use --month-root para informar o caminho do mês."
        )

    while True:
        print()
        print("Meses encontrados:")
        for idx, month_dir in enumerate(month_dirs, start=1):
            print(f"  {idx:>2}) {_month_display(month_dir)}")
        print()

        raw = input(
            "Qual mês você quer processar? Digite o número da lista ou o nome do mês/ano: "
        ).strip()
        if not raw:
            continue

        raw_norm = _norm(raw)
        if raw.isdigit():
            idx = int(raw)
            if 1 <= idx <= len(month_dirs):
                return [month_dirs[idx - 1]]

        exact = [m for m in month_dirs if _norm(_month_display(m)) == raw_norm or _norm(m.name) == raw_norm]
        if len(exact) == 1:
            return exact

        partial = [m for m in month_dirs if raw_norm in _norm(_month_display(m)) or raw_norm in _norm(m.name)]
        if len(partial) == 1:
            return partial

        print("Não consegui identificar um único mês com esse texto. Tente de novo.")



def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Fase 2 do dossiê mensal do Programa: monta o DOCX e o PDF finais direto, "
            "sem absorver novos PDFs da pasta Anexos - Revisão. Usa os títulos atuais do Word de revisão."
        )
    )
    parser.add_argument("--manifest", action="append", help="Manifesto JSON específico. Pode repetir.")
    parser.add_argument("--month-root", action="append", help="Pasta do mês para localizar o manifesto automaticamente.")
    parser.add_argument("--output-root", default=str(DEFAULT_OUTPUT_ROOT), help="Pasta principal onde os meses foram gerados.")
    args = parser.parse_args()

    output_root = Path(args.output_root)
    if args.manifest:
        manifests = [Path(p) for p in args.manifest]
    else:
        month_roots = [Path(p) for p in args.month_root] if args.month_root else _prompt_months()
        manifests = [month_output_paths(m, output_root)["manifest"] for m in month_roots]

    print("\nFase 2: montar final direto")
    for manifest_path in manifests:
        if not manifest_path.exists():
            print(f"[AVISO] Manifesto não encontrado: {manifest_path}")
            continue

        manifest, final_pdf = finalize_month(manifest_path, use_review_folder=False, progress=lambda msg: print(f"[ANDAMENTO] {msg}", flush=True))
        print(f"[OK] Mês processado: {manifest.get('month_label', manifest_path.stem)}")
        print(f"[OK] Word final gerado: {manifest['paths']['final_docx']}")
        print(f"[OK] PDF final gerado: {final_pdf}")
        print(f"[OK] Conferência atualizada: {manifest['paths']['levantamento_xlsx']}")
        print(f"[OK] Manifesto atualizado: {manifest_path}")
        print("-" * 80)


if __name__ == "__main__":
    main()