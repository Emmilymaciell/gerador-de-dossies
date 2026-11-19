# CAPAS (opcional)

Coloque aqui modelos Word de capa (`.docx`) se quiser capas personalizadas:

| Tipo | Nome esperado |
|---|---|
| Capa inicial do dossiê | `CAPA DOSSIÊ.docx` (texto `MÊS - ANO` é substituído pelo mês) |
| Capa de grupo | `CAPA GRUPO.docx` |
| Capa de produto | `capa_subsecao_produtos.docx` |
| Capa de documento | `capa_documento_tecnica.docx` |

Na versão pública os modelos originais foram removidos (identidade visual do cliente).
Sem modelos, o gerador desenha capas simples com Pillow (`render_fast_cover_image`).
