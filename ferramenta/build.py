"""Gera o Painel de Metas da Regional ES a partir do Relatório de Vendas LPA.

Uso: python3 -I build.py <relatorio.xlsx> <data dd/mm/aaaa> <pins.json> <saida index.html>

Saída: <saida index.html> + pasta dados/ ao lado, com um arquivo criptografado por acesso.
O pins.json guarda as senhas e NUNCA deve ir para o repositório.
"""
import sys, zlib, json, os, re, base64, hashlib, secrets
import openpyxl
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

if len(sys.argv) != 5:
    sys.exit(__doc__)
SRC, DATA_REL, PINS_PATH, OUT = sys.argv[1:5]
if not re.fullmatch(r'\d{2}/\d{2}/\d{4}', DATA_REL):
    sys.exit(f'ERRO: data "{DATA_REL}" fora do formato dd/mm/aaaa')

HERE = os.path.dirname(os.path.abspath(__file__))
REGIONAL = 'REGIONAL ES'
SUP = 'ALEX PORTES'                                   # linhas deste "vendedor" são CDs / carga fechada
GERAL = ['ALEX PORTES', 'MARCELO FIALHO']             # acesso à regional inteira
SUPS_ACESSO = ['BIANCA RICCI', 'LEONARDO JUNCA', 'MOYSES BADARO']  # cada um vê só a própria supervisão
ITER = 250000

COLS = ['REGIONAL', 'SUPERVISÃO', 'VENDEDOR', 'REDE', 'CODIGO', 'LOJA', 'FAMILIA', 'CATEGORIA', 'SKU',
        'META CAIXA', 'VENDA FAT CAIXA', 'CARTEIRA - CX',
        'META R$', 'VENDA R$', 'DEV R$', 'DEV SAC R$', 'TROCA R$', 'TROCA SAC R$', 'PEDIDOS ABERTOS R$',
        'META KG', 'VENDA KG', 'DEV KG', 'DEV SAC KG', 'TROCA KG', 'TROCA SAC KG', 'PEDIDOS ABERTOS KG']

wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
for aba in ('FAT+VOL', 'Plan1'):
    if aba not in wb.sheetnames:
        sys.exit(f'ERRO: o relatório não tem a aba "{aba}". Abas: {wb.sheetnames}')

r1 = next(wb['FAT+VOL'].iter_rows(max_row=1, values_only=True))
try:
    dia, dias = int(r1[4]), int(r1[5])
    assert 0 <= dia <= dias and 28 <= dias <= 31
except Exception:
    sys.exit(f'ERRO: não achei dia/dias do mês na aba FAT+VOL (linha 1, colunas E e F): {r1[:8]}')

it = wb['Plan1'].iter_rows(values_only=True)
hdr = [str(h).strip() if h is not None else '' for h in next(it)]
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
    """Limpa tabs, espaços duplos e espaços especiais que vêm do sistema."""
    s = re.sub(r'\s+', ' ', str(v if v is not None else '').replace('\xa0', ' ')).strip()
    return s or vazio


def cod_loja(v):
    try:
        return str(int(float(v)))
    except (TypeError, ValueError):
        return txt(v, '?')


rows = []
for r in it:
    if txt(r[ix['REGIONAL']], '') != REGIONAL:
        continue
    g = lambda c: num(r[ix[c]])
    cx = [g('META CAIXA'), g('VENDA FAT CAIXA'), g('CARTEIRA - CX')]
    val = [g('META R$'), g('VENDA R$'), g('DEV R$') + g('DEV SAC R$'), g('TROCA R$'), g('PEDIDOS ABERTOS R$'),
           g('META KG'), g('VENDA KG'), g('DEV KG') + g('DEV SAC KG') + g('TROCA KG') + g('TROCA SAC KG'),
           g('PEDIDOS ABERTOS KG'), g('TROCA SAC R$')]
    if not any(cx) and not any(val):
        continue
    vend = txt(r[ix['VENDEDOR']])
    sku = str(r[ix['SKU']] or '')
    cod, _, nome = sku.partition('|')
    rows.append(dict(
        v=vend if vend != SUP else 'ALEX PORTES (CDs / carga fechada)',
        sup=txt(r[ix['SUPERVISÃO']]), rede=txt(r[ix['REDE']]),
        lcod=cod_loja(r[ix['CODIGO']]), lnome=txt(r[ix['LOJA']]),
        cat=txt(r[ix['CATEGORIA']]), fam=txt(r[ix['FAMILIA']], 'OUTROS'),
        scod=txt(cod, '') if nome else '', sku=txt(nome or cod), cx=cx, val=val))
if not rows:
    sys.exit(f'ERRO: nenhuma linha da {REGIONAL} na aba Plan1')

# Lojas diferentes com o mesmo nome (códigos diferentes) NÃO podem ser somadas:
# a loja é identificada pelo código; o nome ganha o código quando é repetido.
nomes_por_cod = {}
for x in rows:
    nomes_por_cod.setdefault(x['lcod'], x['lnome'])
cods_por_nome = {}
for c, n in nomes_por_cod.items():
    cods_por_nome.setdefault(n, set()).add(c)
for x in rows:
    n = nomes_por_cod[x['lcod']]
    x['loja'] = f"{n} (cód. {x['lcod']})" if len(cods_por_nome[n]) > 1 else n
    x['skey'] = x['scod'] or x['sku']       # produto identificado pelo código


def pack(rs, with_vend):
    lojas = sorted({x['loja'] for x in rs})
    rede_da = {x['loja']: x['rede'] for x in rs}
    sk = {x['skey']: (x['sku'], x['scod']) for x in rs}
    skeys = sorted(sk, key=lambda k: (sk[k][0], k))
    fams = sorted({x['fam'] for x in rs}); cats = sorted({x['cat'] for x in rs})
    vends = sorted({x['v'] for x in rs}); sups = sorted({x['sup'] for x in rs})
    L, S, F, C, V, P = ({k: i for i, k in enumerate(a)} for a in (lojas, skeys, fams, cats, vends, sups))
    # linha: 0 loja, 1 categoria, 2 produto, 3 meta cx, 4 venda cx, 5 carteira cx, 6 família,
    #        7 meta R$, 8 venda R$, 9 devolução R$, 10 troca R$ (sem SAC), 11 carteira R$,
    #        12 meta KG, 13 venda KG, 14 dev+troca KG, 15 carteira KG, 16 troca SAC R$, [17 vendedor, 18 supervisão]
    d = [[L[x['loja']], C[x['cat']], S[x['skey']]] + [round(v, 2) for v in x['cx']] + [F[x['fam']]]
         + [round(v, 2) for v in x['val']]
         + ([V[x['v']], P[x['sup']]] if with_vend else []) for x in rs]
    return dict(fams=fams, lojas=lojas, redes=[rede_da[l] for l in lojas],
                skus=[sk[k][0] for k in skeys], cods=[sk[k][1] for k in skeys], cats=cats,
                vends=vends if with_vend else None, sups=sups if with_vend else None, d=d)


pins = json.load(open(PINS_PATH, encoding='utf-8')) if os.path.exists(PINS_PATH) else {}
vendedores = sorted({x['v'] for x in rows if not x['v'].startswith(SUP) and not x['v'].startswith('INATIVO')})
novos = [v for v in vendedores + GERAL + SUPS_ACESSO if v not in pins]
for v in novos:
    pins[v] = f"{secrets.randbelow(900000) + 100000}"
if novos:
    json.dump(pins, open(PINS_PATH, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


def enc(obj, senha):
    salt, iv = secrets.token_bytes(16), secrets.token_bytes(12)
    key = hashlib.pbkdf2_hmac('sha256', senha.encode(), salt, ITER, 32)
    comp = zlib.compress(json.dumps(obj, ensure_ascii=False, separators=(',', ':')).encode(), 9)
    b = lambda x: base64.b64encode(x).decode()
    return dict(s=b(salt), i=b(iv), c=b(AESGCM(key).encrypt(iv, comp, None)))


blobs = {v: enc(pack([x for x in rows if x['v'] == v], False), pins[v]) for v in vendedores}
blobs['ALEX PORTES (supervisor)'] = enc(pack(rows, True), pins['ALEX PORTES'])
blobs['MARCELO FIALHO (gerente)'] = enc(pack(rows, True), pins['MARCELO FIALHO'])
for sp in SUPS_ACESSO:
    blobs[sp + ' (supervisor)'] = enc(pack([x for x in rows if x['sup'] == sp], True), pins[sp])

ddir = os.path.join(os.path.dirname(os.path.abspath(OUT)), 'dados')
os.makedirs(ddir, exist_ok=True)
for f in os.listdir(ddir):
    os.remove(os.path.join(ddir, f))
files = {}
for i, (nome, b) in enumerate(blobs.items()):
    files[nome] = f'dados/d{i:02d}.json'
    json.dump(b, open(os.path.join(ddir, f'd{i:02d}.json'), 'w'))
ordem = sorted(blobs, key=lambda n: (('(' in n), n))   # vendedores primeiro, depois gestão
meta_pub = dict(data=DATA_REL, dia=dia, dias=dias, iter=ITER, nomes=ordem, files=files, ver=secrets.token_hex(4))
html = open(os.path.join(HERE, 'template.html'), encoding='utf-8').read()
if '/*__META__*/null' not in html:
    sys.exit('ERRO: template.html sem o marcador de dados')
open(OUT, 'w', encoding='utf-8').write(html.replace('/*__META__*/null', json.dumps(meta_pub, ensure_ascii=False)))

# ---- conferência (comparar com o Excel antes de publicar) ----
print(f'OK  relatório {DATA_REL} · dia {dia}/{dias} · {len(rows)} linhas · {len(vendedores)} vendedores · '
      f'{len(blobs)} acessos · {len({x["loja"] for x in rows})} lojas')
tot = {}
for x in rows:
    t = tot.setdefault(x['sup'], [0, 0, 0]); t[0] += x['val'][0]; t[1] += x['val'][1]; t[2] += x['cx'][0]
for s in sorted(tot):
    print(f'    {s:<16} meta R$ {tot[s][0]:>13,.0f} · venda R$ {tot[s][1]:>12,.0f} · meta cx {tot[s][2]:>9,.0f}')
print(f'    {"REGIONAL":<16} meta R$ {sum(t[0] for t in tot.values()):>13,.0f}')
if novos:
    print('ATENÇÃO  senhas novas criadas (avisar o Alex e guardar na skill):', {v: pins[v] for v in novos})
sem_acesso = sorted({x['sup'] for x in rows} - set(GERAL) - set(SUPS_ACESSO))
if sem_acesso:
    print('ATENÇÃO  supervisão nova sem acesso próprio:', sem_acesso)
