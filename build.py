#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""指定校推薦リストから出願候補 README.md と Excel を生成する。

使い方:
    python3 build.py                # 今日を基準日にして再生成
    python3 build.py 2026-10-15     # 基準日を指定

学校から新しいリストが来たら data/ の xlsx を差し替えて再実行するだけ。
本人の条件（出席率・JLPT 点数）は下の ME を書き換える。
"""
import sys, re, json, glob, datetime as dt
from collections import Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 本人の条件 ───────────────────────────────────────────
ME = {'attendance': 100, 'jlpt': 'N2', 'jlpt_score': 111, 'eju': None}

SRC   = sorted(glob.glob('data/指定校推薦リスト*.xlsx'))[-1]
SHEET = '大学（2026年度）  '
TODAY = dt.date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else dt.date.today()

# ── 判定ルール（本人の条件に対する不可・要確認）───────────────
NG = {
 '中日本自動車短期大学（岐阜）': '短期大学（本人の除外指定）',
 '大阪キリスト教短期大学':      '短期大学（本人の除外指定）',
 '駒沢女子大学・短期大学':      '女子大学（本人の除外指定）',
 '大阪女学院大学':             '女子大学（本人の除外指定）',
 '平安女学院大学':             '女子大学（本人の除外指定）',
 '近畿大学':      'EJU必須（日本語300点＋総合科目100点＋総得点400点）。JLPT不可',
 '日本工業大学':   'EJU必須（日本語230点、数学コース2も受験要）。JLPT不可',
 '神戸学院大学':   'EJU必須（日本語240点）。JLPT不可',
 '大阪国際大学':   'EJU必須（日本語200点）。JLPT不可',
 '上海大学':      '中国国籍以外が条件。該当しない',
 '太成学院大学':   f'N2は112点以上が必要。{ME["jlpt_score"]}点で1点不足（EJU220/JPT525/J.TEST C級600でも可）',
}
WARN = {
 '追手門学院大学（パートナー校制度）':
    '日本語(N3以上)は満たすが、英語スコア(TOEFL iBT74 / IELTS5.5 / TOEIC800 / Duolingo100)も別途必須',
 '園田学園大学':   '「日本語教育の参照枠 B2 以上」が条件。N2 111点が B2 判定になるか要確認',
 '城西国際大学':   'リスト「編集中」。条件・日程とも未記載',
 '名古屋経済大学': '条件は N2 で可。ただし出願期間欄が 2025 年の日付のまま＝更新漏れの可能性',
 '京都外国語大学': '日本語学科は N1 110点以上で不可。グローバルスタディーズ／グローバル観光学科なら N2 100点以上で可',
}

S = lambda x: '' if x is None else str(x).strip()

def parse():
    ws = openpyxl.load_workbook(SRC, data_only=True)[SHEET]
    R = [[c.value for c in r] for r in ws.iter_rows()]
    out, i = [], 0
    while i < len(R):
        if S(R[i][5]) == '人数' and S(R[i][6]).startswith('推薦条件'):
            d = R[i + 1]
            rec = {'status': S(d[0]), 'name': S(d[1]), 'quota': S(d[5]), 'cond': S(d[6])}
            j = i + 2
            while j < len(R) and not (S(R[j][5]) == '人数' and S(R[j][6]).startswith('推薦条件')):
                if S(R[j][1]) == '出願期間':
                    n = R[j + 1]
                    for c, k in ((1, '出願期間'), (2, '試験日'), (3, '発表'), (4, '手続締切'), (5, '試験科目')):
                        rec[k] = S(n[c])
                    break
                j += 1
            if rec['name']:
                out.append(rec)
            i = j
            continue
        i += 1
    seen, U = set(), []
    for x in out:
        k = x['name'].split('\n')[0]
        if k in seen:
            continue
        seen.add(k)
        x['short'] = k
        x['gakubu'] = '／'.join(x['name'].split('\n')[1:])
        U.append(x)
    return U

def windows(t):
    """出願期間セルから (開始, 締切) を全部拾い、締切が未来のものだけ返す。"""
    out = []
    for m in re.finditer(r'(\d{1,2})[/月](\d{1,2})日?\s*[（(]?[月火水木金土日]?[）)]?\s*'
                         r'[～~\-–]\s*(?:\d{4}年)?(\d{1,2})[/月](\d{1,2})', t.replace('　', ' ')):
        a, b, c, d = (int(x) for x in m.groups())
        y1 = 2026 if a >= 8 else 2027
        y2 = 2026 if c >= 8 else 2027
        if c < a:
            y2 = y1 + 1
        try:
            o, cl = dt.date(y1, a, b), dt.date(y2, c, d)
            if cl >= TODAY:
                out.append((o, cl))
        except ValueError:
            pass
    return sorted(out)

def quota(q):
    """Excel が数値化した枠（'1.0'）を '1名' に戻す。"""
    q = re.sub(r'^(\d+)\.0$', r'\1名', q.strip())
    return q.replace('\n', '／')

def jcond(c):
    m = re.search(r'(N\s*[12１２][^\n]{0,80})', c)
    return m.group(1)[:80] if m else ('日本語要件の記載なし' if c else '')

def attreq(c):
    m = re.findall(r'出席[率状況][^\n]{0,12}?(\d{2})\s*[％%]', c)
    return int(max(m)) if m else None

def classify(x):
    n = x['short']
    if n in NG:
        return '✕', NG[n]
    if n in WARN:
        return '⚠', WARN[n]
    a = attreq(x['cond'])
    if a is not None and ME['attendance'] < a:
        return '✕', f'出席率 {a}% 以上が必要（本人 {ME["attendance"]}%）'
    return '◎', ''

def build():
    U = parse()
    rows = []
    for x in U:
        v, why = classify(x)
        ws_ = windows(x.get('出願期間', ''))
        rows.append({
            'v': v, 'why': why, 'n': x['short'], 'g': x['gakubu'],
            'q': quota(x['quota']),
            'sen': any(k in x['cond'] for k in ('専願', '単願', '必ず入学', '入学意思')),
            'j': jcond(x['cond']), 'a': attreq(x['cond']),
            'sub': x.get('試験科目', ''), 'p': x.get('出願期間', ''),
            't': x.get('試験日', ''), 'r': x.get('発表', ''), 'h': x.get('手続締切', ''),
            'w': ws_, 'cond': x['cond'],
        })
    o = {'◎': 0, '⚠': 1, '✕': 2}
    rows.sort(key=lambda r: (o[r['v']], r['w'][0][1] if r['w'] else dt.date(2099, 1, 1)))
    return rows

# ── Excel ────────────────────────────────────────────────
def excel(rows, path='指定校推薦_出願候補.xlsx'):
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title = '出願候補'
    H = ['判定', '締切まで', '学校名', '学部・学科', '枠', '専願', '日本語条件', '出席率条件',
         '出願期間', '試験日', '合格発表', '入学手続き締切', '試験科目', '判定理由・注意点']
    W = [6, 9, 22, 34, 16, 7, 34, 11, 30, 26, 22, 26, 26, 46]
    thin = Side(style='thin', color='D0CEC7')
    BD = Border(left=thin, right=thin, top=thin, bottom=thin)
    FILL = {'◎': 'E8F0E6', '⚠': 'FBF2DC', '✕': 'F2E6E3'}
    ws.append(H)
    for j, (h, w) in enumerate(zip(H, W), 1):
        c = ws.cell(1, j)
        c.font = Font(bold=True, size=10, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='2C5070')
        c.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[1].height = 30
    for r in rows:
        left = (r['w'][0][1] - TODAY).days if r['w'] else None
        ws.append([r['v'], f'{left}日' if left is not None else '—', r['n'], r['g'], r['q'],
                   '専願' if r['sen'] else '', r['j'],
                   f"{r['a']}%以上" if r['a'] else '記載なし',
                   r['p'], r['t'], r['r'], r['h'], r['sub'], r['why']])
        i = ws.max_row
        for j in range(1, len(H) + 1):
            c = ws.cell(i, j)
            c.border = BD; c.font = Font(size=10)
            c.alignment = Alignment(vertical='top', wrap_text=True)
            c.fill = PatternFill('solid', fgColor=FILL[r['v']])
        ws.cell(i, 1).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(i, 1).font = Font(size=13, bold=True)
        ws.cell(i, 2).alignment = Alignment(horizontal='center', vertical='center')
        if left is not None and left <= 30:
            ws.cell(i, 2).font = Font(size=11, bold=True, color='A63A2B')
        if r['sen']:
            ws.cell(i, 6).font = Font(size=10, bold=True, color='A63A2B')
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = f'A1:N{ws.max_row}'

    w2 = wb.create_sheet('条件全文')
    w2.append(['判定', '学校名', '枠', '推薦条件（原文）'])
    for j, w in enumerate([6, 22, 18, 110], 1):
        c = w2.cell(1, j)
        c.font = Font(bold=True, size=10, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='2C5070')
        c.alignment = Alignment(horizontal='center', vertical='center')
        w2.column_dimensions[get_column_letter(j)].width = w
    for r in rows:
        w2.append([r['v'], r['n'], r['q'], r['cond']])
        i = w2.max_row
        for j in range(1, 5):
            c = w2.cell(i, j)
            c.border = BD; c.font = Font(size=9)
            c.alignment = Alignment(vertical='top', wrap_text=True)
            c.fill = PatternFill('solid', fgColor=FILL[r['v']])
    w2.freeze_panes = 'B2'
    wb.save(path)
    return path

# ── README ───────────────────────────────────────────────
def md(rows):
    C = Counter(r['v'] for r in rows)
    ok = [r for r in rows if r['v'] == '◎']
    wn = [r for r in rows if r['v'] == '⚠']
    ng = [r for r in rows if r['v'] == '✕']
    esc = lambda s: (s or '').replace('|', '／').replace('\n', ' ')

    L = []
    L.append('# 指定校推薦 出願候補（2027年4月入学）\n')
    L.append(f'**更新 {TODAY}** ｜ 出席率 {ME["attendance"]}% ｜ JLPT {ME["jlpt"]} {ME["jlpt_score"]}点\n')
    L.append('## 校数\n')
    L.append('| 区分 | 校数 |')
    L.append('|---|---:|')
    L.append(f'| 指定校リスト掲載（大学・2026年度） | **{len(rows)}** |')
    L.append(f'| ◎ この条件で出願可 | **{C["◎"]}** |')
    L.append(f'| ⚠ 追加確認が必要 | {C["⚠"]} |')
    L.append(f'| ✕ 出願不可 | {C["✕"]} |')
    L.append(f'| うち専願（合格＝入学義務） | {sum(1 for r in ok if r["sen"])} / {C["◎"]} |')
    L.append('\n> 別途、専門学校の指定校が 32 校ある（このリストは大学のみ）。\n')

    nxt = [r for r in ok if r['w']]
    if nxt:
        n0 = min(nxt, key=lambda r: r['w'][0][1])
        d0 = (n0['w'][0][1] - TODAY).days
        L.append(f'## 次の締切\n\n**{n0["n"]} — {n0["w"][0][1]:%-m/%-d}（あと {d0} 日）**\n')

    L.append('## ◎ 出願可\n')
    L.append('| 締切 | 残 | 学校 | 枠 | 専願 |')
    L.append('|---|---:|---|---|:-:|')
    for r in ok:
        if r['w']:
            o, c = r['w'][0]
            left = (c - TODAY).days
            dl = f'{c:%-m/%-d}'
            if left <= 14:
                dl = f'**{dl}** 🔴'
            elif left <= 30:
                dl = f'**{dl}** 🟡'
            leftn = str(left)
        else:
            dl, leftn = '—', '—'
        L.append(f'| {dl} | {leftn} | {esc(r["n"])} | {esc(r["q"])} | {"専" if r["sen"] else ""} |')

    L.append('\n## ⚠ 要確認\n')
    for r in wn:
        L.append(f'- **{esc(r["n"])}** — {esc(r["why"])}')
    L.append('\n## ✕ 出願不可\n')
    for r in ng:
        L.append(f'- **{esc(r["n"])}** — {esc(r["why"])}')

    L.append('\n## 各校の詳細\n')
    for r in ok + wn:
        L.append(f'<details><summary><b>{r["v"]} {esc(r["n"])}</b>'
                 + (f' — {esc(r["q"])}' if r['q'] else '')
                 + (' ｜<b>専願</b>' if r['sen'] else '') + '</summary>\n')
        L.append(f'**学部・学科**：{esc(r["g"]) or "—"}  ')
        L.append(f'**日本語条件**：{esc(r["j"]) or "—"}  ')
        L.append(f'**出席率**：{str(r["a"]) + "% 以上" if r["a"] else "記載なし"}  ')
        if r['why']:
            L.append(f'**注意**：{esc(r["why"])}  ')
        L.append('')
        for k, lab in (('p', '出願期間'), ('t', '試験日'), ('r', '発表'),
                       ('h', '入学手続き締切'), ('sub', '試験科目')):
            if r[k]:
                L.append(f'**{lab}**')
                L.append('```')
                L.append(r[k])
                L.append('```')
        L.append('\n</details>\n')

    L.append('---\n')
    L.append('### 前提\n')
    L.append('- 入学は全て **2027年4月**。表中の 2026年10〜12月 は出願・試験日であって入学月ではない。')
    L.append(f'- 出席率 {ME["attendance"]}% のため、最も厳しい 95% 要件も含め全校クリア。この項目は候補を絞る条件にならない。')
    L.append(f'- N2 {ME["jlpt_score"]}点。数値指定があるのは 太成学院 112点（不足）と 京都外国語 100点（クリア）の 2 校のみ。')
    L.append('- **専願＝合格したら入学義務**。併願できないため、複数校への同時出願は原則不可。')
    L.append('- 日付・条件は元シートのセルから機械抽出したもの。**出願前に必ず募集要項の原本で照合すること。**')
    L.append(f'\n出典：`{SRC}`「{SHEET.strip()}」\n')
    L.append('### 更新のしかた\n')
    L.append('学校から新しいリストが来たら `data/` の xlsx を差し替えて:\n')
    L.append('```bash\npython3 build.py && git commit -am "update" && git push\n```')
    return '\n'.join(L) + '\n'

if __name__ == '__main__':
    rows = build()
    open('README.md', 'w', encoding='utf-8').write(md(rows))
    p = excel(rows)
    C = Counter(r['v'] for r in rows)
    print(f'基準日 {TODAY} ｜ 全{len(rows)}校  ◎{C["◎"]}  ⚠{C["⚠"]}  ✕{C["✕"]}')
    print(f'  → README.md, {p}')
