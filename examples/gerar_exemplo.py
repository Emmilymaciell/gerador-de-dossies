"""Gera uma pasta de mês FICTÍCIA para demonstrar o Gerador de Dossiês.

Uso (na raiz do repositório):

    python examples/gerar_exemplo.py

Cria `examples/produtos/01. Janeiro 2026/` com:
- um relatório por grupo (PDF) que lista os produtos do grupo;
- subpastas de produto com documentos PDF;
- um .docx sem PDF correspondente (aparece na planilha de conferência);
- uma planilha de controle de entregas com um item sem arquivo localizado.

Todos os nomes, códigos e textos são inventados. Os arquivos gerados ficam
fora do Git (ver .gitignore).
"""
from __future__ import annotations

from pathlib import Path

import fitz  # PyMuPDF
from docx import Document
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent / "produtos" / "01. Janeiro 2026"

GRUPOS = {
    1: ("Coordenação Executiva", {"1.1": "Relatório de Coordenação", "1.2": "Relatório de Comunicação"}),
    2: ("Engenharia", {"2.1": "Relatório de Engenharia", "2.2": "Relatório de Análise de Projetos"}),
    3: ("Qualidade e Inspeção", {"3.1": "Relatório de Qualidade", "3.3": "Relatório de Inspeção"}),
}

DOCUMENTOS = {
    "1.1": ["Nota Técnica 001-2026 - Plano de trabalho.pdf", "Ata de reunião 002-2026.pdf"],
    "1.2": ["Relatório de comunicação - janeiro.pdf"],
    "2.1": ["Parecer técnico 010-2026 - Pavimento.pdf"],
    "2.2": ["Análise de projeto 011-2026 - Drenagem.pdf"],
    "3.1": ["Checklist de qualidade 020-2026.pdf"],
    "3.3": ["Relatório de inspeção 021-2026 - Trecho A.pdf"],
}


def _pdf(path: Path, linhas: list[str], paginas: int = 1) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open()
    for n in range(paginas):
        page = doc.new_page(width=595, height=842)
        y = 80
        for linha in linhas:
            page.insert_text((60, y), linha, fontsize=13)
            y += 22
        page.insert_text((60, 800), f"Documento fictício - página {n + 1}", fontsize=9)
    doc.save(path)
    doc.close()


def main() -> None:
    for g, (titulo, produtos) in GRUPOS.items():
        linhas = [f"Relatório Mensal - Grupo {g}", titulo, ""]
        linhas += [f"Produto {code} - {nome}" for code, nome in produtos.items()]
        _pdf(ROOT / f"Relatório Grupo {g}.pdf", linhas, paginas=2)

        for code, nome in produtos.items():
            pasta = ROOT / f"Grupo {g}" / f"P{code} - {nome}"
            for arquivo in DOCUMENTOS.get(code, []):
                _pdf(pasta / arquivo, [Path(arquivo).stem, "Conteúdo fictício para demonstração."])

    # Documento sem PDF -> entra no levantamento de pendências
    docx_path = ROOT / "Grupo 2" / "P2.1 - Relatório de Engenharia" / "Memória de cálculo 012-2026.docx"
    docx_path.parent.mkdir(parents=True, exist_ok=True)
    d = Document()
    d.add_paragraph("Memória de cálculo fictícia (sem PDF de propósito).")
    d.save(docx_path)

    # Planilha de controle de entregas (nome contém "controle", "entrega" e "produto")
    wb = Workbook()
    ws = wb.active
    ws.title = "Janeiro"
    ws.append(["Produto", "Documento", "Descrição", "Status", "Tipo"])
    ws.append(["P1.1", "Nota Técnica 001-2026", "Plano de trabalho", "Entregue", "Nota técnica"])
    ws.append(["P2.1", "Parecer técnico 010-2026", "Pavimento", "Entregue", "Parecer"])
    ws.append(["P3.3", "Relatório de inspeção 022-2026", "Trecho B", "Em elaboração", "Relatório"])
    wb.save(ROOT / "Controle de Entregas por Produto.xlsx")

    print(f"Pasta de exemplo criada em: {ROOT}")


if __name__ == "__main__":
    main()
