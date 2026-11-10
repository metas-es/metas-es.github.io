"""Gera o Painel de Metas (Regional ES - Alex Portes) a partir do Relatório de Vendas LPA.
Uso: python3 -I build.py <relatorio.xlsx> <data dd/mm/aaaa> <pins.json> <saida index.html>
"""
import sys, zlib, json, os, re, math, base64, hashlib, secrets
import openpyxl
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

SRC, DATA_REL, PINS_PATH, OUT = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
HERE = os.path.dirname(os.path.abspath(__file__))
SUP = 'ALEX PORTES'
REGIONAL = 'REGIONAL ES'
GERAL = ['ALEX PORTES', 'MARCELO FIALHO']          # acesso à regional inteira
SUPS_ACESSO = ['BIANCA RICCI', 'LEONARDO JUNCA', 'MOYSES BADARO']  # cada um vê só a própria supervisão
ITER = 250000

wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
r1 = next(wb['FAT+VOL'].iter_rows(max_row=1, values_only=True))
dia, dias = int(r1[4]), int(r1[5])

ws = wb['Plan1']
it = ws.iter_rows(values_only=True)
hdr = list(next(it))
ix = {h: i for i, h in enumerate(hdr)}

def num(v):
    try: return float(v or 0)
    except: return 0.0

def nome_sku(s):
    s = str(s or '')
    return s.split('|', 1)[1].strip() if '|' in s else s.strip()

raw = list(it)
# vendedores da equipe = quem aparece na supervisão do Alex; pega TODAS as lojas/redes deles (qualquer supervisão)
rows = []
for r in raw:
    if r[ix['REGIONAL']] != REGIONAL: continue
    vend = r[ix['VENDEDOR']]
    meta, venda, cart = num(r[ix['META CAIXA']]), num(r[ix['VENDA FAT CAIXA']]), num(r[ix['CARTEIRA - CX']])
    val = [num(r[ix['META R$']]), num(r[ix['VENDA R$']]),
           num(r[ix['DEV R$']]) + num(r[ix['DEV SAC R$']]), num(r[ix['TROCA R$']]),
           num(r[ix['PEDIDOS ABERTOS R$']]),
           num(r[ix['META KG']]), num(r[ix['VENDA KG']]),
           num(r[ix['DEV KG']]) + num(r[ix['DEV SAC KG']]) + num(r[ix['TROCA KG']]) + num(r[ix['TROCA SAC KG']]),
           num(r[ix['PEDIDOS ABERTOS KG']]),
           num(r[ix['TROCA SAC R$']])]
    if meta == 0 and venda == 0 and cart == 0 and not any(val): continue
    rows.append(dict(v=vend if vend != SUP else 'ALEX PORTES (CDs / carga fechada)', sup=r[ix['SUPERVISÃO']], rede=r[ix['REDE']], loja=r[ix['LOJA']], cat=r[ix['CATEGORIA']],
                     sku=nome_sku(r[ix['SKU']]), meta=meta, venda=venda, cart=cart,
                     fam=r[ix['FAMILIA']] or 'OUTROS', val=val))

def pack(rs, with_vend):
    lojas = sorted({x['loja'] for x in rs}); skus = sorted({x['sku'] for x in rs})
    fams = sorted({x['fam'] for x in rs}); F = {k: i for i, k in enumerate(fams)}
    cats = sorted({x['cat'] for x in rs}); vends = sorted({x['v'] for x in rs}); sups = sorted({x['sup'] for x in rs})
    redes = {x['loja']: x['rede'] for x in rs}
    L = {k: i for i, k in enumerate(lojas)}; S = {k: i for i, k in enumerate(skus)}
    C = {k: i for i, k in enumerate(cats)}; V = {k: i for i, k in enumerate(vends)}; P = {k: i for i, k in enumerate(sups)}
    # linha: 0 loja,1 cat,2 sku,3 meta cx,4 venda cx,5 cart cx,6 familia,
    # 7 meta R$,8 venda R$,9 dev R$,10 troca R$ (sem SAC),11 carteira R$,12 meta KG,13 venda KG,14 dev+troca KG,15 carteira KG,16 troca SAC R$,[17 vend,18 sup]
    d = [[L[x['loja']], C[x['cat']], S[x['sku']], round(x['meta'], 2), round(x['venda'], 2), round(x['cart'], 2), F[x['fam']]]
         + [round(v, 2) for v in x["val"]]
         + ([V[x['v']], P[x['sup']]] if with_vend else []) for x in rs]
    return dict(fams=fams, lojas=lojas, redes=[redes[l] for l in lojas], skus=skus, cats=cats,
                vends=vends if with_vend else None, sups=sups if with_vend else None, d=d)

pins_path = PINS_PATH
pins = json.load(open(pins_path)) if os.path.exists(pins_path) else {}
vendedores = sorted({x['v'] for x in rows if not x['v'].startswith(SUP) and not str(x['v']).startswith('INATIVO')})
for v in vendedores + GERAL + SUPS_ACESSO:
    if v not in pins: pins[v] = f"{secrets.randbelow(900000) + 100000}"
json.dump(pins, open(pins_path, 'w'), ensure_ascii=False, indent=1)

def enc(obj, pin):
    salt, iv = secrets.token_bytes(16), secrets.token_bytes(12)
    key = hashlib.pbkdf2_hmac('sha256', pin.encode(), salt, ITER, 32)
    raw = zlib.compress(json.dumps(obj, ensure_ascii=False, separators=(',', ':')).encode(), 9)
    ct = AESGCM(key).encrypt(iv, raw, None)
    b = lambda x: base64.b64encode(x).decode()
    return dict(s=b(salt), i=b(iv), c=b(ct))

blobs = {}
for v in vendedores:
    blobs[v] = enc(pack([x for x in rows if x['v'] == v], False), pins[v])
blobs['ALEX PORTES (supervisor)'] = enc(pack(rows, True), pins['ALEX PORTES'])
blobs['MARCELO FIALHO (gerente)'] = enc(pack(rows, True), pins['MARCELO FIALHO'])
for sp in SUPS_ACESSO:
    blobs[sp + ' (supervisor)'] = enc(pack([x for x in rows if x['sup'] == sp], True), pins[sp])

outdir = os.path.dirname(os.path.abspath(OUT)); ddir = os.path.join(outdir, 'dados')
os.makedirs(ddir, exist_ok=True)
for f in os.listdir(ddir): os.remove(os.path.join(ddir, f))
files = {}
for i, (nome, b) in enumerate(blobs.items()):
    fn = f'd{i:02d}.json'; files[nome] = 'dados/' + fn
    json.dump(b, open(os.path.join(ddir, fn), 'w'))
ordem = sorted(blobs, key=lambda n: (('(' in n), n))   # vendedores primeiro, depois gestão
meta_pub = dict(data=DATA_REL, dia=dia, dias=dias, iter=ITER, nomes=ordem, files=files, ver=secrets.token_hex(4))
html = open(os.path.join(HERE, 'template.html'), encoding='utf-8').read()
html = html.replace('/*__META__*/null', json.dumps(meta_pub, ensure_ascii=False)) \
           
open(OUT, 'w', encoding='utf-8').write(html)
print('linhas', len(rows), 'vendedores', len(vendedores), 'acessos', len(blobs), 'dia', dia, '/', dias)
