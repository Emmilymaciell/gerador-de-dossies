# STATE — Gerador de Dossiês (versão pública sanitizada)

> Resumo sanitizado do estado canônico do projeto privado (criado em 2026-10-11).
> Detalhes de infraestrutura interna, caminhos e identificação do cliente foram removidos.

## O que é

GUI (customtkinter) que monta o dossiê mensal de um programa de gerenciamento de
obras de infraestrutura; empacotável com PyInstaller. Trabalho feito no contexto
de um consórcio de infraestrutura rodoviária (vínculo Dynatest).

## Histórico (resumido)

- Primeiro commit do projeto privado: 2026-04-15.
- Antecessor: pasta de scripts soltos (três scripts de linha de comando); a GUI
  (`app.py`) foi adicionada depois, reaproveitando os scripts via `main()`.
- Uma versão anterior ("V005", 2026-04-14) gerava capas com borda; foi substituída
  pela versão atual.

## Notas técnicas

- **Portabilidade:** no original, as pastas de entrada e de modelos de capa eram
  caminhos absolutos da máquina da autora. Na versão pública foram parametrizados
  (`DOSSIE_MONTHS_DIR`, `DOSSIE_OUTPUT_DIR`, `CAPAS/` relativo ao script).
- **Cuidado:** `dossie_01_preparar_base.py` APAGA a pasta do mês de saída antes de
  regenerar, a menos que receba `--no-clean`.
- A conversão DOCX → PDF usa o Microsoft Word via COM (`pywin32`): requer Windows
  com Word instalado.
- Fora do Git: `build/`, `dist/` (executável de ~107 MB) e saídas geradas.

## Verificação

Smoke seguro: compilar os módulos sem executar (`python -m py_compile *.py`).
Não há suíte de testes automatizados neste projeto.

## Próximo passo

Nenhum trabalho de código pendente; parametrização completa dos caminhos e
decisão futura sobre evolução do produto.
