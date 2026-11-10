# Ferramenta do painel

`build.py` le o Relatorio de Vendas LPA (convertido para .xlsx) e gera o `index.html` com os dados criptografados por PIN.
Os PINs NAO ficam aqui (estao guardados na skill metas-es do Claude).

    soffice --headless --convert-to xlsx Relatorio.xlsb
    python3 -I ferramenta/build.py Relatorio.xlsx 09/11/2026 pins.json index.html   # gera index.html + pasta dados/
