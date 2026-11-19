# Sanitização desta versão pública

Esta é uma cópia sanitizada de um sistema em uso real num contrato de infraestrutura. Nenhum dado do cliente está aqui; tudo o que aparece como exemplo é inventado.

## O que foi removido

- Modelos de capa `.docx` com identidade visual do cliente; ícone, atalho `.bat` e `.spec` com o nome do cliente.
- Mascote original (silhueta derivada do logotipo do cliente) — trocado por mascote genérico gerado com Pillow.
- Arquivos de proveniência/governança interna (`PROVENANCE.md`, `OWNERSHIP.md`, `project.yaml`).

## O que foi trocado

- Nome do programa/contrato → `Programa` (identificadores `prog_*`).
- Títulos reais dos 5 grupos e 20 produtos → títulos genéricos (os mesmos nos repositórios irmãos).
- Caminhos locais → caminhos relativos (`examples/produtos`, `saida/`, `CAPAS/`) e variáveis `DOSSIE_MONTHS_DIR` / `DOSSIE_OUTPUT_DIR`.

## O que é fictício

- `examples/gerar_exemplo.py` gera uma pasta de mês fictícia (PDFs, um `.docx`, planilha de controle); a saída fica fora do Git.
- Imagens em `IMAGENS/` são genéricas, sem metadados.

## Como verificar

Sem suíte automatizada (o projeto nunca teve). Verificação feita por smoke com dados fictícios:

```bat
py -m pip install -r requirements.txt
py examples\gerar_exemplo.py
py app.py
```

A geração de DOCX/PDF exige Windows com Microsoft Word (COM).

Varreduras da versão pública (2026-11-19): gitleaks `dir` sem achados; busca por caminho local, e-mail, nomes do cliente e links privados sem ocorrências.
