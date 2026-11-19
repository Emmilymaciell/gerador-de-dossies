# Dossiê mensal — versão revisada dos scripts (README original, sanitizado)

> README de trabalho da autora, mantido como registro. Caminhos locais e
> identificação do cliente foram substituídos; nomes de saída atualizados
> conforme `month_output_paths()` no código atual.

Arquivos principais:

- `dossie_01_preparar_base.py`
- `dossie_02_montar_final.py`
- `dossie_03_aplicar_revisao.py`
- `dossie_utils_v2_andamento.py`

## Lógica nova

### Script 1 — preparar base
Gera, para cada mês:
- `Dossiê Mensal - <mês> - Revisão.docx` (dentro de `Anexos - Revisão`)
- `Manifesto.json`
- `Conferência do Dossiê - <mês>.xlsx` (levantamento de documentos sem PDF)
- pasta `Anexos - Revisão`, com subpastas por grupo e produto

Ele **não gera o PDF final**.

### Script 2 — montar final
Usa o DOCX base já gerado e os PDFs que já estiverem em `Anexos - Revisão`.

### Script 3 — aplicar revisão
Lê a pasta `Anexos - Revisão` do mês:
- usa o DOCX revisado que você editar manualmente
- absorve os PDFs faltantes colocados ali
- gera o DOCX e o PDF finais

## Estrutura de saída
Cada mês fica em uma pasta própria dentro da pasta de saída (`saida/` por padrão
na versão pública; configurável por `--output-root` ou `DOSSIE_OUTPUT_DIR`).

## Como rodar

### 1) Gerar a base
```bat
cd /d "C:\caminho\exemplo\gerador-de-dossies"
py dossie_01_preparar_base.py
```

### 2) Gerar final direto, sem revisão manual
```bat
cd /d "C:\caminho\exemplo\gerador-de-dossies"
py dossie_02_montar_final.py
```

### 3) Aplicar revisão manual e gerar final
```bat
cd /d "C:\caminho\exemplo\gerador-de-dossies"
py dossie_03_aplicar_revisao.py
```

## Dependências

```bat
py -m pip install -r requirements.txt
```

## Observações importantes

- O base sai em **Word editável**.
- O final sai em **PDF com sumário clicável**.
- A numeração é aplicada **uma única vez no final**, para evitar sobreposição.
- O levantamento inclui:
  - arquivos encontrados sem PDF
  - registros do controle de entregas sem PDF localizado
  - itens mantidos no sumário sem arquivo final localizado
- O levantamento traz também a **página base** do item no dossiê sem anexos.
