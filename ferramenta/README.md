# Ferramenta do painel

`build.py` lê o Relatório de Vendas LPA (convertido para .xlsx) e gera o `index.html` e a pasta `dados/`
(um arquivo criptografado por acesso). As senhas NÃO ficam aqui: o `pins.json` fica fora do repositório.

    soffice --headless --convert-to xlsx Relatorio.xlsb
    python3 -I ferramenta/build.py Relatorio.xlsx 09/11/2026 /caminho/pins.json index.html

No final o programa imprime a conferência (meta e venda por supervisão) para comparar com o Excel
e avisa se criou senha nova ou se apareceu supervisão sem acesso.
