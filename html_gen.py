#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""出願候補データから docs/index.html（スマホ優先）を生成する。"""
import json, re, datetime as dt, html as H
from collections import Counter

CSS = """
:root{
  --bg:#f1f2ef; --surf:#fbfbf9; --sunk:#e7e8e3; --line:#dcddd6; --line2:#c5c7bd;
  --ink:#191c1f; --ink2:#464b51; --dim:#767b82;
  --acc:#1f4e6b; --acc-bg:#e3ecf2; --on-acc:#fbfbf9;
  --hot:#ad3b2b; --hot-bg:#f8e7e3;
  --warm:#8a6519; --warm-bg:#f7efdb;
  --good:#356b46;
  --sh:0 1px 2px rgba(25,28,31,.06),0 8px 20px -14px rgba(25,28,31,.3);
}
@media (prefers-color-scheme:dark){:root:not([data-t="light"]){
  --bg:#0f1215; --surf:#181c20; --sunk:#1f2429; --line:#282e34; --line2:#394047;
  --ink:#e8e7e3; --ink2:#b6bbc1; --dim:#8b9199;
  --acc:#7fb3d5; --acc-bg:#16242e; --on-acc:#0f1215;
  --hot:#e28a7b; --hot-bg:#2b1c19;
  --warm:#d2a860; --warm-bg:#282017;
  --good:#7bb992;
  --sh:0 1px 2px rgba(0,0,0,.4),0 8px 20px -14px rgba(0,0,0,.7);
}}
:root[data-t="dark"]{
  --bg:#0f1215; --surf:#181c20; --sunk:#1f2429; --line:#282e34; --line2:#394047;
  --ink:#e8e7e3; --ink2:#b6bbc1; --dim:#8b9199;
  --acc:#7fb3d5; --acc-bg:#16242e; --on-acc:#0f1215;
  --hot:#e28a7b; --hot-bg:#2b1c19;
  --warm:#d2a860; --warm-bg:#282017;
  --good:#7bb992;
  --sh:0 1px 2px rgba(0,0,0,.4),0 8px 20px -14px rgba(0,0,0,.7);
}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--ink);
  font:400 15px/1.6 -apple-system,BlinkMacSystemFont,"Hiragino Sans","Noto Sans JP","Yu Gothic",system-ui,sans-serif;
  -webkit-font-smoothing:antialiased;padding-bottom:56px}
.mono{font-family:ui-monospace,"SF Mono",Menlo,monospace;font-variant-numeric:tabular-nums}
.wrap{max-width:820px;margin-inline:auto;padding:0 14px}

/* header */
header{background:var(--surf);border-bottom:1px solid var(--line);padding:18px 0 0}
.ttl{font:600 21px/1.25 "Hiragino Mincho ProN","Yu Mincho",serif;letter-spacing:.01em;margin:0}
.sub{color:var(--dim);font-size:12.5px;margin-top:5px;display:flex;flex-wrap:wrap;gap:4px 10px}
.golink{color:#C4405F;font-weight:700;text-decoration:none;border-bottom:1px solid currentColor}
.golink:hover{opacity:.75}
.sub b{color:var(--ink2);font-weight:600}
.counts{display:flex;gap:7px;margin:14px 0 0;overflow-x:auto;padding-bottom:2px;
  scrollbar-width:none}
.counts::-webkit-scrollbar{display:none}
.cnt{flex:none;background:var(--sunk);border:1px solid var(--line);border-radius:8px;
  padding:8px 11px;min-width:74px}
.cnt .n{font:600 20px/1 ui-monospace,monospace;font-variant-numeric:tabular-nums}
.cnt .l{font-size:10.5px;color:var(--dim);margin-top:3px;letter-spacing:.03em}
.cnt.a .n{color:var(--acc)} .cnt.h .n{color:var(--hot)}

/* filters */
.bar{position:sticky;top:0;z-index:20;background:var(--surf);
  border-bottom:1px solid var(--line);margin-top:14px}
.facets{max-width:820px;margin-inline:auto;padding:8px 0 9px;
  display:flex;flex-direction:column;gap:5px}
.fl{display:flex;gap:6px;overflow-x:auto;padding:2px 14px;scrollbar-width:none;
  align-items:center;max-width:820px;margin-inline:auto}
.fl::-webkit-scrollbar{display:none}
.fl>.lab{flex:none;font:600 10.5px/1 inherit;letter-spacing:.1em;color:var(--dim);
  width:34px;text-align:right;margin-right:2px}
.fl>.sublab{flex:none;font:500 10.5px/1 inherit;color:var(--dim);
  white-space:nowrap;margin-right:4px;width:64px;text-align:right}
.facets{padding-bottom:10px}
.fl.sub .f{font-size:12px;padding:6px 10px;border-style:dashed}
.fl.sub .f[aria-pressed="true"]{border-style:solid}
.f .n{opacity:.6;font-size:11px;margin-left:3px;font-variant-numeric:tabular-nums}
.f[hidden]{display:none}
.reset{flex:none;font:500 11.5px/1 inherit;color:var(--dim);background:none;border:0;
  cursor:pointer;text-decoration:underline;padding:6px 4px}
details.g{border-top:1px solid var(--line)}
details.g:first-child{border-top:0}
details.g>summary{list-style:none;cursor:pointer;padding:9px 14px;
  font:600 13px/1 inherit;color:var(--ink);display:flex;align-items:center;gap:7px;
  max-width:820px;margin-inline:auto}
details.g>summary::-webkit-details-marker{display:none}
details.g>summary::before{content:"▸";color:var(--dim);font-size:11px;
  transition:transform .15s;display:inline-block}
details.g[open]>summary::before{transform:rotate(90deg)}
details.g>summary:focus-visible{outline:2px solid var(--acc);outline-offset:-2px}
.badge{font:600 10.5px/1 ui-monospace,monospace;color:var(--surf);background:var(--acc);
  border-radius:99px;padding:3px 6px;display:none}
.badge.on{display:inline-block}
details.g .fl{padding:3px 14px 3px 26px}
details.g .fl:last-child{padding-bottom:9px}
.n1box{background:var(--warm-bg);border:1px solid var(--warm);border-radius:7px;
  padding:9px 11px;margin-top:10px;font-size:12.5px;line-height:1.6;color:var(--ink2)}
.n1box b{color:var(--warm)}
.ch.n1{border-color:var(--warm);color:var(--warm);background:var(--warm-bg)}
.hit{font-size:12px;color:var(--dim);padding:2px 14px;max-width:820px;margin-inline:auto}
.hit{display:flex;align-items:center;gap:6px;padding-bottom:8px}
.hit b{color:var(--acc);font-weight:600}
.reset{margin-left:auto}
.f{flex:none;font:500 13px/1 inherit;padding:8px 13px;border-radius:99px;cursor:pointer;
  border:1px solid var(--line2);background:transparent;color:var(--ink2);
  transition:background .12s,color .12s,border-color .12s}
.f[aria-pressed="true"]{background:var(--acc);border-color:var(--acc);color:var(--on-acc);font-weight:600}
.f:focus-visible{outline:2px solid var(--acc);outline-offset:2px}

/* next-up */
.next{margin:16px 0 4px;background:var(--hot-bg);border:1px solid var(--hot);
  border-left-width:4px;border-radius:9px;padding:13px 15px;display:flex;gap:14px;align-items:center}
.next .d{font:600 30px/1 ui-monospace,monospace;color:var(--hot);font-variant-numeric:tabular-nums}
.next .d small{display:block;font:600 9.5px/1.3 inherit;letter-spacing:.1em;margin-top:4px}
.next .x{min-width:0}
.next .n{font-weight:600;font-size:15px}
.next .s{font-size:12.5px;color:var(--ink2);margin-top:2px}

/* section head */
h2{font:600 13px/1 inherit;letter-spacing:.09em;color:var(--dim);
  margin:26px 0 10px;text-transform:uppercase;display:flex;align-items:center;gap:9px}
h2::after{content:"";flex:1;height:1px;background:var(--line)}

/* card */
.list{display:flex;flex-direction:column;gap:9px}
.c{background:var(--surf);border:1px solid var(--line);border-left:3px solid var(--line2);
  border-radius:9px;box-shadow:var(--sh);overflow:hidden}
.c.hot{border-left-color:var(--hot)} .c.warm{border-left-color:var(--warm)}
.c.cool{border-left-color:var(--acc)} .c.ng{border-left-color:var(--line2);opacity:.72}
.c[hidden]{display:none}
.hd{display:grid;grid-template-columns:42px 1fr auto;gap:10px;align-items:start;
  padding:12px 13px;cursor:pointer;width:100%;background:none;border:0;
  font:inherit;color:inherit;text-align:left}
.hd:focus-visible{outline:2px solid var(--acc);outline-offset:-2px}
.cd{text-align:center;padding-top:2px}
.cd .n{font:600 19px/1 ui-monospace,monospace;font-variant-numeric:tabular-nums;display:block}
.cd .u{font-size:9.5px;color:var(--dim);letter-spacing:.06em;display:block;margin-top:3px}
.c.hot .cd .n{color:var(--hot)} .c.warm .cd .n{color:var(--warm)} .c.cool .cd .n{color:var(--acc)}
.nm{font-weight:600;font-size:15px;line-height:1.35;min-width:0}
.mt{font-size:11.5px;color:var(--dim);margin-top:3px;line-height:1.45;
  display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.chips{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px}
.cost{padding-top:1px}
.ch{font:500 10.5px/1 inherit;padding:3.5px 6px;border-radius:5px;border:1px solid var(--line2);
  color:var(--ink2);background:var(--sunk);white-space:nowrap}
.ch.sen{border-color:var(--hot);color:var(--hot);background:var(--hot-bg)}
.ch.q{border-color:var(--acc);color:var(--acc);background:var(--acc-bg)}
.ch.good{border-color:var(--good);color:var(--good);background:transparent}
.ch.w{border-color:var(--warm);color:var(--warm);background:var(--warm-bg)}
.ar{color:var(--dim);font-size:13px;transition:transform .18s;flex:none}
.c.open .ar{transform:rotate(90deg)}
.bd{display:none;padding:0 13px 14px;border-top:1px solid var(--line);margin-top:2px}
.c.open .bd{display:block}
.bd dl{margin:0;display:flex;flex-direction:column;gap:11px;padding-top:12px}
.bd dt{font:600 10.5px/1 inherit;letter-spacing:.09em;color:var(--dim);text-transform:uppercase}
.bd dd{margin:5px 0 0;font-size:13.5px;line-height:1.65;white-space:pre-wrap;color:var(--ink2)}
.bd dd.pre{font-family:ui-monospace,"SF Mono",Menlo,monospace;font-size:12.5px;
  background:var(--sunk);border:1px solid var(--line);border-radius:6px;padding:9px 11px;
  overflow-x:auto;color:var(--ink)}
.warnbox{background:var(--warm-bg);border:1px solid var(--warm);border-radius:7px;
  padding:10px 12px;font-size:13px;color:var(--ink2);margin-top:12px}
.c.ng .warnbox{background:var(--hot-bg);border-color:var(--hot)}

/* rarity */
.c.t-gold{border-top:3px solid #c9962a}
.c.t-purple{border-top:3px solid #8b5cf6}
.c.t-blue{border-top:3px solid #3b82f6}
.ch.tg{background:linear-gradient(135deg,#f5d06f,#c9962a);color:#3a2c05;border-color:#c9962a;font-weight:700}
.ch.tp{background:#8b5cf6;color:#fff;border-color:#8b5cf6;font-weight:700}
.ch.tb{background:#3b82f6;color:#fff;border-color:#3b82f6;font-weight:700}
.ch.tw{background:var(--sunk);color:var(--dim);border-color:var(--line2)}
.tierbox{border-radius:7px;padding:9px 11px;margin-top:10px;font-size:12.5px;line-height:1.6;
  border:1px solid var(--line2);color:var(--ink2);background:var(--sunk)}
.binsec{margin-top:26px;border-top:1px solid var(--line);padding-top:8px}
.binsec summary{cursor:pointer;font:600 13px/1.6 inherit;color:var(--dim);list-style:none;padding:8px 0}
.binsec summary::before{content:"▸ ";font-size:11px}
.binsec[open] summary::before{content:"▾ "}
.binrow{display:flex;gap:10px;padding:7px 4px;font-size:13px;color:var(--dim);
  border-bottom:1px dashed var(--line)}
.binrow b{color:var(--ink2);font-weight:500}

/* cost */
.cost{text-align:right;flex:none;display:flex;flex-direction:column;gap:1px;align-items:flex-end}
.cost .v{font:600 16px/1 ui-monospace,monospace;font-variant-numeric:tabular-nums;white-space:nowrap}
.cost .v.net{color:var(--good)}
.cost .v.na{color:var(--dim);font-size:12.5px;font-weight:500}
.cost .k{font-size:9.5px;color:var(--dim);letter-spacing:.05em}
.cost .was{font-size:10.5px;color:var(--dim);text-decoration:line-through;
  font-family:ui-monospace,monospace}
.loc{display:inline-flex;align-items:center;gap:3px}
.feebox{background:var(--acc-bg);border:1px solid var(--acc);border-radius:7px;
  padding:10px 12px;margin-top:12px;font-size:13px;line-height:1.6;color:var(--ink2)}
.feebox b{color:var(--ink)}
.feebox .big{font:600 19px/1.2 ui-monospace,monospace;color:var(--good);display:block;margin:3px 0 6px}
.sortbar{display:flex;gap:6px;align-items:center;margin:18px 0 -4px;font-size:12.5px;color:var(--dim)}
.sortbar button{font:500 12.5px/1 inherit;padding:6px 11px;border-radius:99px;cursor:pointer;
  border:1px solid var(--line2);background:transparent;color:var(--ink2)}
.sortbar button[aria-pressed="true"]{background:var(--sunk);border-color:var(--acc);
  color:var(--acc);font-weight:600}
.sortbar button:focus-visible{outline:2px solid var(--acc);outline-offset:2px}

/* notes */
.notes{margin-top:30px;border-top:1px solid var(--line);padding-top:16px;
  font-size:12.5px;line-height:1.75;color:var(--dim)}
.notes b{color:var(--ink2)} .notes ul{margin:8px 0;padding-left:18px}
.notes li{margin:5px 0}
.notes a{color:var(--acc)}
.empty{text-align:center;color:var(--dim);font-size:13.5px;padding:34px 0}
@media(min-width:560px){.ttl{font-size:25px}.hd{grid-template-columns:56px 1fr auto;gap:13px}}
@media(prefers-reduced-motion:reduce){*{transition:none!important}}
"""

JS = """
const $=s=>document.querySelectorAll(s);
const DAY=864e5, T=new Date(); T.setHours(0,0,0,0);
// 倒计时はビルド時ではなく閲覧時に再計算する
$('.c[data-dl]').forEach(c=>{
  const dl=new Date(c.dataset.dl+'T00:00:00');
  const d=Math.round((dl-T)/DAY);
  c.dataset.left=d;
  const n=c.querySelector('.cd .n'), u=c.querySelector('.cd .u');
  if(!n) return;
  if(d<0){ n.textContent='—'; u.textContent='終了'; c.classList.add('ng'); c.dataset.over='1'; }
  else { n.textContent=d; u.textContent=d===0?'本日':'日'; }
  c.classList.remove('hot','warm','cool');
  c.classList.add(d<0?'ng':d<=14?'hot':d<=30?'warm':'cool');
});
// トグル
$('.hd').forEach(b=>b.onclick=()=>{
  const c=b.closest('.c'); c.classList.toggle('open');
  b.setAttribute('aria-expanded', c.classList.contains('open'));
});
// ---- 多維ファセット絞り込み（各軸は AND、複数値フィールドは包含判定）----
const KEY='shingaku.facets.v2';
const MULTI={chiho:1,prefs:1,wt:1};            // 空白区切りの複数値フィールド
const GROUP={req:['reqtop','req','ab','form','exam','qb'],
             loc:['chiho','prefs','zone'],
             fee:['band','wv','wt'],
             etc:['status','dl','n1']};
const F={};
$('.f').forEach(b=>{ if(!(b.dataset.facet in F)) F[b.dataset.facet]='all'; });

function dlBand(L){ return L<0?null : L<=14?'d14' : L<=30?'d30' : L<=60?'d60' : 'd90'; }

function match(c){
  for(const k in F){
    const want=F[k]; if(want==='all') continue;
    if(k==='dl'){
      const L=+c.dataset.left, b=dlBand(L);
      // 「30天内」は 14 天内も含む、という累積で扱う
      const ord=['d14','d30','d60','d90'];
      if(b===null || ord.indexOf(b)>ord.indexOf(want)) return false;
      continue;
    }
    const got=c.dataset[k]||'';
    if(MULTI[k]){ if(!got.split(' ').includes(want)) return false; }
    else if(got!==want) return false;
  }
  return true;
}

function apply(){
  let n=0;
  $('.c').forEach(c=>{ const ok=match(c); c.hidden=!ok; if(ok)n++; });
  $('.sec').forEach(s=>{ s.hidden=![...s.querySelectorAll('.c')].some(c=>!c.hidden); });
  document.getElementById('empty').hidden=n>0;
  document.getElementById('nhit').textContent=n;

  // 押下状態
  $('.f').forEach(b=>b.setAttribute('aria-pressed', F[b.dataset.facet]===b.dataset.v));

  // 親が選ばれている時は、子チップを親に属するものだけに絞る
  $('.f[data-parent]').forEach(b=>{
    const pf=b.dataset.facet==='req'?'reqtop':'chiho';
    const pv=F[pf];
    b.hidden = pv!=='all' && b.dataset.parent!==pv;
    if(b.hidden && F[b.dataset.facet]===b.dataset.v){ F[b.dataset.facet]='all'; }
  });

  // グループごとの有効フィルタ数バッジ
  for(const g in GROUP){
    const n2=GROUP[g].filter(k=>F[k]&&F[k]!=='all').length;
    const el=document.querySelector(`.badge[data-g="${g}"]`);
    if(el){ el.textContent=n2; el.classList.toggle('on', n2>0); }
  }
  try{ localStorage.setItem(KEY, JSON.stringify(F)) }catch(e){}
}

$('.f').forEach(b=>b.onclick=()=>{
  const k=b.dataset.facet;
  F[k] = (F[k]===b.dataset.v) ? 'all' : b.dataset.v;   // 同じものを再クリックで解除
  apply(); apply();                                    // 2回目で親子の整合を取る
});
document.getElementById('reset').onclick=()=>{ for(const k in F) F[k]='all'; apply(); };
try{ const sv=JSON.parse(localStorage.getItem(KEY)||'{}');
     for(const k in sv) if(k in F) F[k]=sv[k]; }catch(e){}
apply(); apply();

// 並べ替え（締切順 / 学費順）
const SKEY='shingaku.sort';
function sortBy(mode){
  document.querySelectorAll('.sortbar button')
    .forEach(b=>b.setAttribute('aria-pressed', b.dataset.s===mode));
  document.querySelectorAll('.sec .list').forEach(list=>{
    const cs=[...list.children];
    cs.sort((a,b)=>{
      if(mode==='fee'){
        const fa=+a.dataset.fee||9e9, fb=+b.dataset.fee||9e9;
        if(fa!==fb) return fa-fb;
      }
      if(mode==='tier'){
        const R={gold:0,purple:1,blue:2,white:3};
        const ta=R[a.dataset.tier]??9, tb=R[b.dataset.tier]??9;
        if(ta!==tb) return ta-tb;
        const fa=+a.dataset.fee||9e9, fb=+b.dataset.fee||9e9;
        if(fa!==fb) return fa-fb;
      }
      return (+a.dataset.left) - (+b.dataset.left);
    });
    cs.forEach(c=>list.appendChild(c));
  });
  try{ localStorage.setItem(SKEY,mode) }catch(e){}
}
document.querySelectorAll('.sortbar button').forEach(b=>b.onclick=()=>sortBy(b.dataset.s));
let s0='date'; try{ s0=localStorage.getItem(SKEY)||'date' }catch(e){}
sortBy(s0);
"""


def quota_short(q):
    """「【姫路キャンパス】／２名／／【大阪天王寺キャンパス】／1名」→「共3名」"""
    if not q:
        return ''
    z = str.maketrans('０１２３４５６７８９', '0123456789')
    q2 = q.translate(z)
    ns = [int(x) for x in re.findall(r'(\d+)\s*名', q2)]
    if ns:
        return f'共{sum(ns)}名' if len(ns) > 1 else f'{ns[0]}名'
    if '若干' in q2:
        return '若干名'
    if '記載なし' in q2 or not q2.strip():
        return '未注明'
    return q2[:12]


def man(n):
    """1148000 → '114.8万' / 800000 → '80万'"""
    v = n / 10000
    return (f'{v:.0f}' if abs(v - round(v)) < .05 else f'{v:.1f}') + '万'


def render(rows, TODAY, ME, src, sheet):
    e = lambda s: H.escape(s or '')
    binned = [r for r in rows if r.get('bin')]
    rows = [r for r in rows if not r.get('bin')]
    C = Counter(r['v'] for r in rows)
    ok = [r for r in rows if r['v'] == '◎']
    sen_n = sum(1 for r in ok if r['sen'])
    nxt = min([r for r in ok if r['w']], key=lambda r: r['w'][0][1], default=None)

    def card(r):
        v = {'◎': 'ok', '◇': 'cond', '⚠': 'warn', '✕': 'ng'}[r['v']]
        dl = r['w'][0][1] if r['w'] else None
        left = (dl - TODAY).days if dl else None
        sev = 'ng' if v == 'ng' else 'cool' if v == 'cond' else ('hot' if left is not None and left <= 14
              else 'warm' if left is not None and left <= 30 else 'cool')
        sev += ' t-' + r['tier']
        num = str(left) if left is not None else '—'
        unit = '天' if left is not None else ('不可' if v == 'ng' else '未定')
        loc, fee = r['loc'], r['fee']
        pref = loc.get('pref', '')
        camp = loc.get('campus', [])

        # --- 学費セル ---
        net, lst = fee.get('net'), fee.get('list')
        if net:
            cost = (f'<span class="was">{man(lst)}</span>' if lst and lst != net else '') + \
                   f'<span class="v net">{man(net)}</span><span class="k">第一年实付</span>'
        else:
            cost = ''

        chips = [f'<span class="ch t{ {"gold":"g","purple":"p","blue":"b","white":"w"}[r["tier"]] }">'
                 f'{ {"gold":"金","purple":"紫","blue":"蓝","white":"白"}[r["tier"]] }</span>']
        if r['q']:
            chips.append(f'<span class="ch q">名额 {e(quota_short(r["q"]))}</span>')
        if r['sen']:
            chips.append('<span class="ch sen">专愿</span>')
        if fee.get('hit'):
            chips.append('<span class="ch good">自动减免已计入</span>')
        elif fee.get('upside'):
            chips.append('<span class="ch w">有可争取减免</span>')
        if r['exam'] == 'oral':
            chips.append('<span class="ch">不考笔试</span>')
        if r['n1gain']:
            chips.append('<span class="ch n1">N1 加成</span>')
        if r['v'] == '⚠':
            chips.append('<span class="ch w">待确认</span>')

        meta = []
        if pref:
            meta.append(f'<span class="loc">{e(pref)}'
                        + (f' · {e(camp[0][1])}' if camp else '') + '</span>')
        if r['g']:
            meta.append(e(r['g'][:26]))

        det = []
        if r['q']:
            det.append(f'<dt>名额（原文）</dt><dd>{e(r["q"])}</dd>')
        if camp:
            cl = '<br>'.join(f'{e(a)} — {e(b)}' for a, b in camp)
            det.append(f'<dt>地址</dt><dd>{e(pref)}<br>{cl}</dd>')
        if loc.get('memo'):
            det.append(f'<dt>校区备注</dt><dd>{e(loc["memo"])}</dd>')
        for k, lab in (('g', '学部・学科'), ('j', '日语要求')):
            if r[k]:
                det.append(f'<dt>{lab}</dt><dd>{e(r[k])}</dd>')
        for k, lab in (('p', '报名期间'), ('t', '考试日'), ('r', '放榜'),
                       ('h', '入学手续截止'), ('sub', '考试科目')):
            if r[k]:
                det.append(f'<dt>{lab}</dt><dd class="pre">{e(r[k])}</dd>')

        fb = ''
        if fee:
            fb = (f'<div class="feebox"><span class="big">第一年实付 {man(net)}日元（{net:,}）</span>'
                  + (f'定价 <b>{man(lst)}日元</b>，' if lst and lst != net else '')
                  + f'<b>已计入（自动适用）</b>：{e(fee.get("auto") or "无")}<br>'
                  + (f'<b>可争取（不计入实付）</b>：{e(fee["upside"])}<br>' if fee.get('upside') else '')
                  + (f'{e(fee["note"])}<br>' if fee.get('note') else '')
                  + f'<span style="font-size:11.5px">来源：{e(fee.get("src",""))}，2026-09-01 核对'
                  + (('<br>' + '　'.join(f'<a href="{H.escape(u)}" target="_blank" rel="noopener">🔗 {e(lab)}</a>'
                                          for lab, u in fee['src_url'])) if fee.get('src_url') else '')
                  + (f'<br><a href="{H.escape(fee["evidence"])}" target="_blank" rel="noopener">📷 官网页面存证截图（点开看原图）</a>'
                     if fee.get('evidence') else '')
                  + '</span></div>')

        why = f'<div class="warnbox">{e(r["why"])}</div>' if r['why'] else ''
        if r['datefix']:
            why += f'<div class="warnbox">{e(r["datefix"])}</div>'
        if r['n1gain']:
            why += f'<div class="n1box"><b>考出 N1 的话</b>：{e(r["n1gain"])}</div>'
        why += (f'<div class="tierbox"><b>{ {"gold":"🟡 金卡","purple":"🟣 紫卡","blue":"🔵 蓝卡","white":"⚪ 白卡"}[r["tier"]] }</b>'
                f'：{e(r["tier_why"])}</div>')
        return (f'<article class="c {sev}" data-status="{v}" '
                f'data-form="{"sen" if r["sen"] else "free"}" '
                f'data-reqtop="{r["reqtop"]}" data-req="{r["req"]}" data-ab="{r["ab"]}" '
                f'data-exam="{r["exam"]}" data-qb="{r["qb"]}" data-zone="{e(r["zone"])}" '
                f'data-chiho="{e(" ".join(r["chiho"]))}" data-prefs="{e(" ".join(r["prefs"]))}" '
                f'data-band="{r["band"]}" data-wv="{r["wv"]}" data-wt="{" ".join(r["wt"])}" '
                f'data-n1="{"y" if r["n1gain"] else "n"}" data-tier="{r["tier"]}" '
                f'data-fee="{net or lst or 0}" '
                f'data-left="{left if left is not None else 9999}"'
                + (f' data-dl="{dl.isoformat()}"' if dl else '') + '>'
                f'<button class="hd" aria-expanded="false">'
                f'<span class="cd"><span class="n mono">{num}</span><span class="u">{unit}</span></span>'
                f'<span class="nm">{e(r["n"])}'
                + (f'<span class="mt">{" · ".join(meta)}</span>' if meta else '')
                + (f'<span class="chips">{"".join(chips)}</span>' if chips else '')
                + f'</span><span class="cost">{cost}</span></button>'
                f'<div class="bd">{why}{fb}<dl>{"".join(det)}</dl></div></article>')

    def sec(title, items):
        if not items:
            return ''
        return (f'<section class="sec"><h2>{title}</h2>'
                f'<div class="list">{"".join(card(r) for r in items)}</div></section>')

    # ---- ファセットのチップを組み立てる ----
    def chips(facet, items, allcnt):
        o = [f'<button class="f" data-facet="{facet}" data-v="all" '
             f'aria-pressed="true">全部<span class="n">{allcnt}</span></button>']
        for val, lab, cnt, par in items:
            o.append(f'<button class="f" data-facet="{facet}" data-v="{e(val)}" '
                     + (f'data-parent="{e(par)}" ' if par else '')
                     + f'aria-pressed="false">{e(lab)}<span class="n">{cnt}</span></button>')
        return ''.join(o)

    def byfield(field, order, labels, parent=None):
        c = Counter()
        par = {}
        for r in rows:
            vs = r[field] if isinstance(r[field], list) else [r[field]]
            for i, v in enumerate(vs):
                c[v] += 1
                if parent:
                    pv = r[parent]
                    par[v] = pv[i] if isinstance(pv, list) and i < len(pv) else (
                        pv[0] if isinstance(pv, list) else pv)
        keys = [k for k in order if c[k]] if order else \
               [k for k, _ in sorted(c.items(), key=lambda t: -t[1])]
        return [(k, labels.get(k, k), c[k], par.get(k, '')) for k in keys]

    N = len(rows)

    # --- 申请条件 ---
    F_REQTOP = chips('reqtop', byfield('reqtop', ['jlpt', 'eju', 'eng', 'unknown'],
        {'jlpt': 'JLPT 可', 'eju': '只认 EJU', 'eng': '另需英语', 'unknown': '条件未记载'}), N)
    F_REQ = chips('req', byfield('req', ['n3', 'n2', 'n2hi', 'split', 'cefr'],
        {'n3': 'N3 即可', 'n2': 'N2 可', 'n2hi': 'N2 需 112 分',
         'split': '按学科不同', 'cefr': 'CEFR 判定'}, parent='reqtop'), N)
    F_ATT = chips('ab', byfield('ab', ['a80', 'a85', 'a90', 'a95', 'ax'],
        {'a80': '80%↑', 'a85': '85%↑', 'a90': '90%↑', 'a95': '95%↑', 'ax': '未记载'}), N)
    TL = {'gold': '金', 'purple': '紫', 'blue': '蓝', 'white': '白'}
    F_TIER = chips('tier', byfield('tier', ['gold', 'purple', 'blue', 'white'], TL), N)
    F_FORM = chips('form', [('sen', '专愿', sum(1 for r in rows if r['sen']), ''),
                            ('free', '可兼报', sum(1 for r in rows if not r['sen']), '')], N)
    F_EXAM = chips('exam', byfield('exam', ['oral', 'written', 'unknown'],
        {'oral': '只有书类＋面试', 'written': '有笔试／小论文', 'unknown': '未记载'}), N)
    F_QB = chips('qb', byfield('qb', ['q1', 'q3', 'q5', 'qs', 'qx'],
        {'q1': '1 名', 'q3': '2〜3 名', 'q5': '4 名以上', 'qs': '若干名', 'qx': '未注明'}), N)

    # --- 学校位置 ---
    F_CHIHO = chips('chiho', byfield('chiho', ['近畿', '関東', '中部', '北海道', '海外'],
        {'近畿': '近畿', '関東': '关东', '中部': '中部', '北海道': '北海道', '海外': '海外'}), N)
    F_PREF = chips('prefs', byfield('prefs', None, {}, parent='chiho'), N)
    F_ZONE = chips('zone', byfield('zone',
        ['大阪市内', '大阪府下', '神戸・阪神', '播磨', '京都', '奈良', 'kansai_other', 'far', 'overseas'],
        {'kansai_other': '関西・其他', 'far': '要搬家（关西外）', 'overseas': '海外'}), N)

    # --- 学费 ---
    F_BAND = chips('band', byfield('band', ['a', 'b', 'c'],
        {'a': '～100 万', 'b': '100〜120 万', 'c': '120 万～'}), N)
    F_WV = chips('wv', byfield('wv', ['auto', 'try', 'no'],
        {'auto': '自动减免·已计入', 'try': '可争取（评选制）', 'no': '首年无减免'}), N)
    F_WT = chips('wt', byfield('wt', ['w_adm', 'w_tui', 'w_att', 'w_gpa', 'w_inc'],
        {'w_adm': '免入学金', 'w_tui': '学费打折', 'w_att': '按出勤率',
         'w_gpa': '按成绩', 'w_inc': '按家庭收入'}), N)

    # --- 时间・其他 ---
    cv = Counter(r['v'] for r in rows)
    F_STATUS = chips('status', [(k, lab, cv[m], '') for k, lab, m in
                                [('ok', '可报', '◎'), ('cond', '条件卡', '◇'),
                                 ('ng', '不可报', '✕')]
                                if cv[m]], N)
    F_DL = chips('dl', [('d14', '14 天内', 0, ''), ('d30', '30 天内', 0, ''),
                        ('d60', '60 天内', 0, ''), ('d90', '60 天以上', 0, '')], N)
    n1n = sum(1 for r in rows if r['n1gain'])
    F_N1 = chips('n1', [('y', '考出 N1 会变好', n1n, ''),
                        ('n', '与 N1 无关', N - n1n, '')], N)

    nb = ''
    if nxt:
        d = (nxt['w'][0][1] - TODAY).days
        nb = (f'<div class="next"><div class="d mono">{d}<small>天后</small></div>'
              f'<div class="x"><div class="n">{e(nxt["n"])}</div>'
              f'<div class="s">{nxt["w"][0][1]:%-m月%-d日} 截止'
              + ('・专愿' if nxt['sen'] else '・可兼报') + '</div></div></div>')

    return f"""<!doctype html>
<html lang="zh-CN"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#1f4e6b">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="指定校">
<title>指定校推荐 报考清单</title>
<style>{CSS}</style>
</head><body>

<header>
  <div class="wrap">
    <h1 class="ttl">指定校推荐 报考清单</h1>
    <div class="sub">
      <span>2027年4月入学</span><span>·</span>
      <span>出勤率 <b>{ME['attendance']}%</b></span><span>·</span>
      <span>JLPT <b>{ME['jlpt']} {ME['jlpt_score']}点</b></span><span>·</span>
      <span>更新 <b>{TODAY}</b></span>
      <span>·</span>
      <a class="golink" href="momoyama_docs.html">🍑 桃山推薦 出願材料チェック →</a>
      <span>·</span>
      <a class="golink" href="det_drill.html">🔤 DET 补全训练 →</a>
    </div>
    <div class="counts">
      <div class="cnt"><div class="n">{len(rows)}</div><div class="l">纳入考虑</div></div>
      <div class="cnt a"><div class="n">{C['◎']}</div><div class="l">可报</div></div>
      <div class="cnt"><div class="n">{C['◇']}</div><div class="l">条件卡</div></div>
      <div class="cnt"><div class="n">{C['✕']}</div><div class="l">不可报</div></div>
      <div class="cnt h"><div class="n">{sen_n}</div><div class="l">专愿</div></div>
      <div class="cnt"><div class="n">32</div><div class="l">专门学校</div></div>
    </div>
  </div>
  <div class="bar">
    <div class="facets">
      <details class="g" open><summary>申请条件<span class="badge" data-g="req"></span></summary>
        <div class="fl"><span class="sublab">稀有度</span>{F_TIER}</div>
        <div class="fl"><span class="sublab">证书种类</span>{F_REQTOP}</div>
        <div class="fl"><span class="sublab">日语等级</span>{F_REQ}</div>
        <div class="fl"><span class="sublab">出勤率</span>{F_ATT}</div>
        <div class="fl"><span class="sublab">报考形式</span>{F_FORM}</div>
        <div class="fl"><span class="sublab">选拔方式</span>{F_EXAM}</div>
        <div class="fl"><span class="sublab">名额</span>{F_QB}</div>
      </details>
      <details class="g"><summary>学校位置<span class="badge" data-g="loc"></span></summary>
        <div class="fl"><span class="sublab">地方</span>{F_CHIHO}</div>
        <div class="fl"><span class="sublab">都道府県</span>{F_PREF}</div>
        <div class="fl"><span class="sublab">通学圈</span>{F_ZONE}</div>
      </details>
      <details class="g"><summary>学费<span class="badge" data-g="fee"></span></summary>
        <div class="fl"><span class="sublab">实付档</span>{F_BAND}</div>
        <div class="fl"><span class="sublab">减免有无</span>{F_WV}</div>
        <div class="fl"><span class="sublab">减免类型</span>{F_WT}</div>
      </details>
      <details class="g"><summary>时间・其他<span class="badge" data-g="etc"></span></summary>
        <div class="fl"><span class="sublab">判定</span>{F_STATUS}</div>
        <div class="fl"><span class="sublab">截止</span>{F_DL}</div>
        <div class="fl"><span class="sublab">N1 加成</span>{F_N1}</div>
      </details>
    </div>
    <div class="hit"><b id="nhit">—</b> 所符合<button class="reset" id="reset">清空全部</button></div>
  </div>
</header>

<main class="wrap">
  {nb}
  <div class="sortbar"><span>排序</span>
    <button data-s="date" aria-pressed="true">按截止日</button>
    <button data-s="fee"  aria-pressed="false">按学费</button>
    <button data-s="tier" aria-pressed="false">按稀有度</button>
  </div>
  {sec('◎ 可报', ok)}
  {sec('◇ 条件卡 — 拿到英语成绩即入池', [r for r in rows if r['v'] == '◇'])}
  {sec('✕ 不可报', [r for r in rows if r['v'] == '✕'])}
  <p class="empty" id="empty" hidden>没有符合条件的学校。</p>

  <details class="binsec"><summary>🗑 已丢弃 {len(binned)} 所（短大・女大，按你的指令不予考虑，不参与筛选）</summary>
    {''.join(f'<div class="binrow"><b>{e(r["n"])}</b><span>{e(r["why"])}</span></div>' for r in binned)}
  </details>

  <div class="notes">
    <b>怎么看这张表</b>
    <ul>
      <li>入学时期全部是 <b>2027年4月</b>。表里 2026年10〜12月 的日期是报名和考试日，不是入学月。</li>
      <li>出勤率 {ME['attendance']}%，连最严的 95% 要求也过，所以<b>这一项筛不掉任何学校</b>。</li>
      <li>N2 {ME['jlpt_score']} 分。写了具体分数线的只有两所：<b>太成学院 112 分（差1分不够）</b>、<b>京都外国語 100 分（够）</b>。</li>
      <li><b>专愿＝考上就必须去</b>，不能同时报别的学校。所以这不是撒网，是押一注。</li>
      <li>金额是<b>第一年</b>的。「实付」＝定价减去你能拿到的减免；划掉的是定价。第二年起的金额多数会变（有的减免只有第一年，有的跟成绩挂钩）。</li>
      <li>只是大学名单。另外还有<b>专门学校的指定校 32 所</b>没做进来。</li>
      <li>学校名、学部学科名、以及报名期间/考试日那些引用原文<b>保留日文</b>——这些是你去官网和募集要項里搜索、核对时要用的原字符串。</li>
      <li>日期和条件是从原表机械抽取的，学费是官网公开值。<b>报名前必须用募集要項原本核对一遍。</b></li>
    </ul>
    来源：<code>{e(src)}</code>「{e(sheet.strip())}」＋ 各大学官网<br>
    剩余天数是按你打开页面当天在浏览器里算的，所以不重新构建也是准的。<br>
    学校发新版名单时，把 <code>data/</code> 里的 xlsx 换掉，然后 <code>python3 build.py</code> → push。
  </div>
</main>

<script>{JS}</script>
</body></html>
"""
