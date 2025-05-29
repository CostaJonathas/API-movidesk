# Exportador de Tickets Movidesk para Excel

Este script Python consome a API do Movidesk, extrai informações de tickets ativos e exporta os dados organizados em abas separadas de um arquivo Excel.

## Funcionalidades
- Consulta múltiplas URLs da API do Movidesk.
- Converte datas UTC para o fuso horário de Brasília.
- Extrai o nome da empresa a partir de dados aninhados.
- Exporta os resultados para um arquivo Excel com várias abas.

## Pré-requisitos

- Python 3.x
- Bibliotecas:
  - requests
  - pandas
  - openpyxl
  - pytz

Instale com:

```bash
pip install requests pandas openpyxl pytz
