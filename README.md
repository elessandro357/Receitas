# Comparativo de Receitas — Resumo 1º Semestre (Streamlit)

Dashboard simples e direto para comparar as receitas **2024 x 2025** a partir da **aba Resumo** do seu Excel.
Funciona com a planilha enviada (`Resumo_1S`) e com variações (detecta colunas de 2024 e 2025 automaticamente).

## Como rodar

1. Instale as dependências (ideal usar um virtualenv):
   ```bash
   pip install -r requirements.txt
   ```

2. Coloque seu arquivo `.xlsx` em `data/` (já incluí um exemplo).
   Por padrão, o app procura `data/comparativo_receitas_2024_2025_1S_COM_ICMS.xlsx`.
   Se quiser usar outro nome/caminho, defina a variável de ambiente:
   ```bash
   export RECEITAS_FILE="data/seu_arquivo.xlsx"
   ```

3. Rode o app:
   ```bash
   streamlit run app.py
   ```

4. No dashboard, você pode:
   - **Selecionar/limpar** os *segmentos* (multiselect + botões rápidos).
   - **Ordenar** por 2025, 2024, diferença ou ordem alfabética.
   - **Definir Top N** para focar nos principais itens.
   - **Clicar na legenda** dos gráficos para ocultar/mostrar séries.
   - **Baixar** a tabela filtrada em CSV/Excel.

## Estrutura

```
receitas_dashboard/
├─ app.py
├─ requirements.txt
├─ README.md
└─ data/
   └─ comparativo_receitas_2024_2025_1S_COM_ICMS.xlsx  # exemplo (substitua pelo seu)
```

## Notas técnicas

- O app tenta localizar automaticamente a aba de resumo (qualquer nome contendo "Resumo"), com fallback para `Resumo_1S`.
- As colunas dos anos são detectadas por *substring* (`"2024"` e `"2025"`). Se houver múltiplas colunas por ano, a que contiver indícios de "Jan/Jun/1S" é priorizada.
- As diferenças são recomputadas internamente: `diff_abs = 2025 - 2024` e `diff_% = diff_abs / 2024`.
- Variáveis e funções em inglês onde faz sentido para manter o padrão do seu código.

## Deploy rápido (opcional)

- **GitHub:** crie um repositório e suba os arquivos.
- **Streamlit Cloud:** conecte ao repositório e defina `app.py` como *main file*.
- **Env var opcional:** `RECEITAS_FILE` apontando para seu `.xlsx` em `data/`.