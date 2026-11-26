"""Extrai de um FECHAMENTO mensal (Relatório de Vendas LPA, .xlsb ou .xlsx) só o que o histórico de trocas usa.

Uso: python3 -I historico.py <fechamento.xlsb> <AAAA-MM> <saida.json.gz>

Saída: JSON compactado com uma linha por supervisão/vendedor/loja/SKU da REGIONAL ES:
venda R$ e troca R$ (sem SAC) e troca SAC R$. Guardar FORA do repositório (pasta do HD),
porque não é criptografado. O build.py junta esses arquivos com --hist <pasta>.
"""
import sys, re, json, gzip
from python_calamine import CalamineWorkbook

if len(sys.argv) != 4 or not re.fullmatch(r'\d{4}-\d{2}', sys.argv[2]):
    sys.exit(__doc__)
SRC, MES, OUT = sys.argv[1:4]
REGIONAL = 'REGIONAL ES'
COLS = ['REGIONAL', 'SUPERVISÃO', 'VENDEDOR', 'REDE', 'CODIGO', 'LOJA', 'FAMILIA', 'CATEGORIA',
        'VENDA R$', 'TROCA R$', 'TROCA SAC R$']   # SKU é opcional (o fechamento de julho veio sem produto)

wb = CalamineWorkbook.from_path(SRC)
if 'Plan1' not in wb.sheet_names:
    sys.exit(f'ERRO: o fechamento não tem a aba "Plan1". Abas: {wb.sheet_names}')
rows = wb.get_sheet_by_name('Plan1').to_python()
hdr = [str(h).strip() for h in rows[0]]
ix = {h: i for i, h in enumerate(hdr)}
faltam = [c for c in COLS if c not in ix]
if faltam:
    sys.exit(f'ERRO: a aba Plan1 mudou; colunas não encontradas: {faltam}')


def num(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def txt(v, vazio='(SEM NOME)'):
    s = re.sub(r'\s+', ' ', str(v if v is not None else '').replace('\xa0', ' ')).strip()
    return s or vazio


def cod_loja(v):
    try:
        return str(int(float(v)))
    except (TypeError, ValueError):
        return txt(v, '?')


agg = {}
for r in rows[1:]:
    if txt(r[ix['REGIONAL']], '') != REGIONAL:
        continue
    vals = [num(r[ix['VENDA R$']]), num(r[ix['TROCA R$']]), num(r[ix['TROCA SAC R$']])]
    if not any(vals):
        continue
    sku = str(r[ix['SKU']] or '') if 'SKU' in ix else ''
    cod, _, nome = sku.partition('|')
    if not sku:
        cod, nome = '', '(sem detalhe de produto neste mês)'
    k = (txt(r[ix['SUPERVISÃO']]), txt(r[ix['VENDEDOR']]), txt(r[ix['REDE']]), cod_loja(r[ix['CODIGO']]),
         txt(r[ix['LOJA']]), txt(r[ix['CATEGORIA']]), txt(r[ix['FAMILIA']], 'OUTROS'),
         txt(cod, '') if nome else '', txt(nome or cod))
    a = agg.setdefault(k, [0.0, 0.0, 0.0])
    for i in range(3):
        a[i] += vals[i]
if not agg:
    sys.exit(f'ERRO: nenhuma linha da {REGIONAL} na aba Plan1')

campos = ['sup', 'v', 'rede', 'lcod', 'lnome', 'cat', 'fam', 'scod', 'sku']
out = [dict(zip(campos, k), vr=round(a[0], 2), tr=round(a[1], 2), tsac=round(a[2], 2)) for k, a in agg.items()]
with gzip.open(OUT, 'wt', encoding='utf-8') as f:
    json.dump(dict(mes=MES, rows=out), f, ensure_ascii=False, separators=(',', ':'))
tv = sum(x['vr'] for x in out); tt = sum(x['tr'] for x in out)
print(('AVISO  fechamento sem coluna SKU: produtos aparecem como "sem detalhe"\n' if 'SKU' not in ix else '') + f'OK  {MES}: {len(out)} linhas · venda R$ {tv:,.0f} · troca R$ {tt:,.0f} ({tt / tv * 100 if tv else 0:.1f}%)')
