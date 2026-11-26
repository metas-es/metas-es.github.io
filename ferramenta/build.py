"""Gera o Painel de Metas da Regional ES a partir do Relatório de Vendas LPA.

Uso: python3 -I build.py <relatorio.xlsx> <data dd/mm/aaaa> <pins.json> <saida index.html> [--hist <pasta>]

--hist <pasta>: pasta com os fechamentos já extraídos por historico.py (AAAA-MM.json.gz), para o
filtro de período da aba Trocas. Essa pasta fica FORA do repositório (não é criptografada).

Saída: <saida index.html> + pasta dados/ ao lado, com um arquivo criptografado por acesso.
O pins.json guarda as senhas e NUNCA deve ir para o repositório.
"""
import sys, zlib, json, os, re, base64, hashlib, secrets, gzip, glob
import openpyxl
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

args = sys.argv[1:]
HIST_DIR = None
if '--hist' in args:
    i = args.index('--hist')
    if i + 1 >= len(args):
        sys.exit(__doc__)
    HIST_DIR = args[i + 1]
    del args[i:i + 2]
if len(args) != 4:
    sys.exit(__doc__)
SRC, DATA_REL, PINS_PATH, OUT = args
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

# ---- histórico de trocas (fechamentos mensais) ----
MES_ATUAL = f'{DATA_REL[6:10]}-{DATA_REL[3:5]}'
hist = []
if HIST_DIR:
    arqs = sorted(glob.glob(os.path.join(HIST_DIR, '*.json.gz')))
    if not arqs:
        sys.exit(f'ERRO: nenhum fechamento (AAAA-MM.json.gz) na pasta {HIST_DIR}')
    for a in arqs:
        h = json.load(gzip.open(a, 'rt', encoding='utf-8'))
        com_troca = {}
        for x in h['rows']:
            x['mes'] = h['mes']
            if x['v'] == SUP:
                x['v'] = 'ALEX PORTES (CDs / carga fechada)'
            if x['tr'] or x['tsac']:
                hist.append(x)
            else:   # produtos sem troca: só a venda importa, somada por loja/categoria (arquivo menor)
                k = (x['sup'], x['v'], x['rede'], x['lcod'], x['cat'], x['fam'])
                o = com_troca.setdefault(k, dict(x, scod='', sku='(demais produtos)', vr=0.0, tr=0.0, tsac=0.0))
                o['vr'] += x['vr']
        hist.extend(com_troca.values())
MESES = sorted({x['mes'] for x in hist} | {MES_ATUAL})
ATUAL_NO_HIST = any(x['mes'] == MES_ATUAL for x in hist)

# Lojas diferentes com o mesmo nome (códigos diferentes) NÃO podem ser somadas:
# a loja é identificada pelo código; o nome ganha o código quando é repetido.
nomes_por_cod = {}
for x in rows + sorted(hist, key=lambda x: x['mes'], reverse=True):   # nome atual primeiro
    nomes_por_cod.setdefault(x['lcod'], x['lnome'])
cods_por_nome = {}
for c, n in nomes_por_cod.items():
    cods_por_nome.setdefault(n, set()).add(c)
for x in rows + hist:
    n = nomes_por_cod[x['lcod']]
    x['loja'] = f"{n} (cód. {x['lcod']})" if len(cods_por_nome[n]) > 1 else n
    x['skey'] = x['scod'] or x['sku']       # produto identificado pelo código


def pack(rs, hs, with_vend):
    tudo = rs + hs
    lojas = sorted({x['loja'] for x in tudo})
    rede_da = {x['loja']: x['rede'] for x in hs}
    rede_da.update({x['loja']: x['rede'] for x in rs})          # rede atual vale mais
    sk = {x['skey']: (x['sku'], x['scod']) for x in tudo}
    skeys = sorted(sk, key=lambda k: (sk[k][0], k))
    fams = sorted({x['fam'] for x in tudo}); cats = sorted({x['cat'] for x in tudo})
    vends = sorted({x['v'] for x in tudo}); sups = sorted({x['sup'] for x in tudo})
    M = {m: i for i, m in enumerate(MESES)}
    L, S, F, C, V, P = ({k: i for i, k in enumerate(a)} for a in (lojas, skeys, fams, cats, vends, sups))
    # linha: 0 loja, 1 categoria, 2 produto, 3 meta cx, 4 venda cx, 5 carteira cx, 6 família,
    #        7 meta R$, 8 venda R$, 9 devolução R$, 10 troca R$ (sem SAC), 11 carteira R$,
    #        12 meta KG, 13 venda KG, 14 dev+troca KG, 15 carteira KG, 16 troca SAC R$, [17 vendedor, 18 supervisão]
    d = [[L[x['loja']], C[x['cat']], S[x['skey']]] + [round(v, 2) for v in x['cx']] + [F[x['fam']]]
         + [round(v, 2) for v in x['val']]
         + ([V[x['v']], P[x['sup']]] if with_vend else []) for x in rs]
    # histórico: 0 mês, 1 loja, 2 categoria, 3 produto, 4 família, 5 venda R$, 6 troca R$, 7 troca SAC R$,
    #            [8 vendedor, 9 supervisão]
    h = [[M[x['mes']], L[x['loja']], C[x['cat']], S[x['skey']], F[x['fam']],
          round(x['vr'], 2), round(x['tr'], 2), round(x['tsac'], 2)]
         + ([V[x['v']], P[x['sup']]] if with_vend else []) for x in hs]
    return dict(fams=fams, lojas=lojas, redes=[rede_da[l] for l in lojas],
                skus=[sk[k][0] for k in skeys], cods=[sk[k][1] for k in skeys], cats=cats,
                vends=vends if with_vend else None, sups=sups if with_vend else None, d=d,
                meses=MESES, mesAtual=M[MES_ATUAL], atualNoHist=ATUAL_NO_HIST, h=h)


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


blobs = {v: enc(pack([x for x in rows if x['v'] == v], [x for x in hist if x['v'] == v], False), pins[v])
         for v in vendedores}
blobs['ALEX PORTES (supervisor)'] = enc(pack(rows, hist, True), pins['ALEX PORTES'])
blobs['MARCELO FIALHO (gerente)'] = enc(pack(rows, hist, True), pins['MARCELO FIALHO'])
for sp in SUPS_ACESSO:
    blobs[sp + ' (supervisor)'] = enc(pack([x for x in rows if x['sup'] == sp], [x for x in hist if x['sup'] == sp], True),
                                      pins[sp])

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
if hist:
    print(f'    histórico de trocas: {", ".join(sorted({x["mes"] for x in hist}))} · {len(hist)} linhas'
          + (' · o mês do relatório já tem fechamento (a aba Trocas usa o fechamento)' if ATUAL_NO_HIST else ''))
    ht = {}
    for x in hist:
        t = ht.setdefault((x['mes'], x['sup']), [0, 0]); t[0] += x['vr']; t[1] += x['tr']
    for (m, sp), t in sorted(ht.items()):
        if sp == SUP:
            print(f'    {m} {sp:<16} venda R$ {t[0]:>12,.0f} · troca R$ {t[1]:>10,.0f}')
if novos:
    print('ATENÇÃO  senhas novas criadas (avisar o Alex e guardar na skill):', {v: pins[v] for v in novos})
sem_acesso = sorted({x['sup'] for x in rows} - set(GERAL) - set(SUPS_ACESSO))
if sem_acesso:
    print('ATENÇÃO  supervisão nova sem acesso próprio:', sem_acesso)
