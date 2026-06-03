"""사군자 회비 대시보드  |  python dashboard.py  →  http://localhost:8888"""
import json, os
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

EXCEL_PATH = os.path.join(os.path.dirname(__file__), '..', 'data', '사군자.xlsx')

def eval_formula(expr):
    if not isinstance(expr, str): return expr
    if expr.startswith('='): expr = expr[1:]
    try: return float(eval(expr.replace('INT(','int('), {'int':int,'abs':abs,'__builtins__':{}}))
    except: return None

def ym_sort_key(ym):
    try:
        p = ym.replace('년 ','-').replace('월','').strip().split('-')
        return (int(p[0]), int(p[1]))
    except: return (99,99)

def load_member_status():
    """입금현황 2행에서 유지/탈퇴 읽기. 미표기는 유지로 간주. 탈퇴만 제외."""
    wb = load_workbook(EXCEL_PATH, data_only=True)
    ws = wb['입금현황']
    rows = list(ws.iter_rows(min_row=1, max_row=2, values_only=True))
    headers, statuses = rows[0], rows[1]
    active = []
    for name, status in zip(headers, statuses):
        if name in (None, 'No', '탈퇴멤버', '계'): continue
        if status == '탈퇴': continue   # 탈퇴만 제외, 미표기는 유지로 간주
        active.append(name)
    return active

def load_income_by_month():
    # 회비 입금은 탈퇴멤버 포함 전체 합산 (실제 계좌 입금액 기준)
    df = pd.read_excel(EXCEL_PATH, sheet_name='입금현황')
    members = [c for c in df.columns if c not in ['No','계'] and not str(c).startswith('Unnamed')]
    data = df[df['No'].apply(lambda x: isinstance(x,str) and '년' in str(x))].copy()
    result = {}
    for _, row in data.iterrows():
        total = sum(pd.to_numeric(row[m],errors='coerce') for m in members if pd.notna(row[m]))
        result[row['No']] = int(total)
    return result

def load_all_payment_status():
    """모든 월의 납부 현황 반환 (유지 멤버만)"""
    active = load_member_status()
    df = pd.read_excel(EXCEL_PATH, sheet_name='입금현황')
    members = [c for c in df.columns if c in active]
    data = df[df['No'].apply(lambda x: isinstance(x,str) and '년' in str(x))]
    real = data[data[members].apply(lambda r: r.notna().any(), axis=1)]
    all_status = {}
    for ym in real['No'].tolist():
        row = df[df['No']==ym].iloc[0]
        all_status[ym] = {m: (pd.notna(row[m]) and float(row[m])>0) for m in members}
    months = list(all_status.keys())
    latest = months[-1] if months else ''
    return all_status, latest, members

def load_wonjang_rows():
    wb = load_workbook(EXCEL_PATH)
    ws = wb['원장']
    rows = []
    for row in ws.iter_rows(min_row=3, values_only=False):
        일자_v=row[2].value; 분류1=row[3].value; 분류2=row[4].value
        차변_r=row[5].value; 대변_r=row[6].value; 비고=row[8].value
        if 일자_v is None or 일자_v=='입력': continue
        try: 일자=pd.to_datetime(일자_v)
        except: continue
        연월=f"{str(일자.year)[2:]}년 {일자.month}월"
        차변=float(eval_formula(차변_r)) if 차변_r is not None and eval_formula(차변_r) is not None else 0.0
        대변=float(eval_formula(대변_r)) if 대변_r is not None and eval_formula(대변_r) is not None else 0.0
        rows.append({'일자':일자,'연월':연월,'분류1':분류1,'분류2':분류2,'차변':차변,'대변':대변,'순액':대변-차변,'비고':str(비고) if 비고 else ''})
    return rows

def load_member_totals():
    # 탈퇴 포함 전체 (잔액 계산용)
    df = pd.read_excel(EXCEL_PATH, sheet_name='입금현황')
    members = [c for c in df.columns if c not in ['No','탈퇴멤버','계'] and not str(c).startswith('Unnamed')]
    data = df[df['No'].apply(lambda x: isinstance(x,str) and '년' in str(x))].copy()
    return [{'name':m,'paid':int(pd.to_numeric(data[m],errors='coerce').fillna(0).sum())} for m in members]

def build_data():
    rows = load_wonjang_rows()
    income_by_month = load_income_by_month()
    all_pay_status, latest_ym, all_members = load_all_payment_status()
    extra_income={}
    for r in rows:
        if r['분류1']=='입금' and r['분류2'] in ('지각비','이자'):
            extra_income[r['연월']]=extra_income.get(r['연월'],0)+r['대변']
    expense_by_cat={}; expense_by_month={}
    for r in rows:
        if r['분류1']=='출금' and r['차변']>0:
            net = r['차변'] - max(r['대변'], 0)  # 대변(환불)이 있으면 차감
            expense_by_cat[r['분류2']]=expense_by_cat.get(r['분류2'],0)+net
            expense_by_month[r['연월']]=expense_by_month.get(r['연월'],0)+net
    total_in=sum(income_by_month.values())+sum(extra_income.values())
    total_out=sum(expense_by_cat.values())
    balance=total_in-total_out
    all_months=sorted(set(list(income_by_month)+list(expense_by_month)),key=ym_sort_key)
    monthly=[]; cumulative=0
    for ym in all_months:
        inc=income_by_month.get(ym,0)+extra_income.get(ym,0)
        exp=expense_by_month.get(ym,0)
        cumulative+=inc-exp
        monthly.append({'month':ym,'income':int(inc),'expense':int(exp),'cumulative':int(cumulative)})
    mwi=[m for m in monthly if m['income']>0]
    def avg_n(n):
        sub=mwi[-n:] if len(mwi)>=n else mwi
        return int(sum(m['income'] for m in sub)/len(sub)) if sub else 0
    cat_data=[{'name':k,'value':int(v)} for k,v in sorted(expense_by_cat.items(),key=lambda x:-x[1]) if v>0]
    member_totals=load_member_totals()
    valid=[r for r in rows if r['차변']>0 or r['대변']>0]
    for r in valid:
        if r['분류2']=='회비' and r['분류1']=='입금' and r['대변']==0:
            r['대변']=income_by_month.get(r['연월'],0); r['순액']=r['대변']
    recent=sorted(valid,key=lambda r:r['일자'],reverse=True)[:10]
    recent_list=[{'date':r['일자'].strftime('%y년 %m월 %d일'),'type':r['분류2'],'amount':int(r['순액']),'memo':r['비고'],'is_income':r['분류1']=='입금'} for r in recent]
    member_names=[mt['name'] for mt in member_totals]
    payment_by_month={
        ym: [{'name':m,'paid':st.get(m,False)} for m in all_members if m in member_names]
        for ym,st in all_pay_status.items()
    }
    pay_months=list(payment_by_month.keys())
    return {
        'summary':{'total_in':total_in,'total_out':total_out,'balance':balance,'avg3':avg_n(3),'avg6':avg_n(6),'avg12':avg_n(12),'updated':datetime.now().strftime('%Y-%m-%d %H:%M'),'latest_ym':latest_ym},
        'monthly':monthly,'categories':cat_data,'members':member_totals,
        'payment_by_month':payment_by_month,'pay_months':pay_months,'recent':recent_list,
    }

HTML = '''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>사군자 회비</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{
  --bg:#080808;
  --s1:#111111;
  --s2:#161616;
  --s3:#1c1c1c;
  --line:rgba(255,255,255,0.08);
  --line2:rgba(255,255,255,0.14);
  --p:#00D4A0;
  --p-dim:rgba(0,212,160,0.10);
  --p-dim2:rgba(0,212,160,0.18);
  --w:#ffffff;
  --w2:rgba(255,255,255,0.7);
  --w3:rgba(255,255,255,0.35);
  --w4:rgba(255,255,255,0.15);
  --neg:#ff5e5e;
  --neg-dim:rgba(255,94,94,0.10);
}
*{box-sizing:border-box;margin:0;padding:0;}
html{background:var(--bg);}
body{background:var(--bg);color:var(--w);font-family:'DM Sans',sans-serif;font-weight:400;min-height:100vh;-webkit-font-smoothing:antialiased;}

/* NAV */
nav{
  position:sticky;top:0;z-index:50;
  background:rgba(8,8,8,0.8);
  backdrop-filter:blur(24px);
  border-bottom:1px solid var(--line);
  padding:0 28px;
  height:52px;
  display:flex;align-items:center;justify-content:space-between;
}
.nav-logo{font-family:'DM Serif Display',serif;font-size:17px;color:var(--w);letter-spacing:0.3px;}
.nav-logo em{color:var(--p);font-style:normal;}
.nav-right{display:flex;align-items:center;gap:14px;}
.nav-time{font-size:11px;color:var(--w3);letter-spacing:0.5px;}
.nav-btn{background:transparent;border:1px solid var(--line2);color:var(--w3);padding:5px 14px;border-radius:6px;cursor:pointer;font-family:inherit;font-size:11px;letter-spacing:0.8px;transition:border-color .2s,color .2s;}
.nav-btn:hover{border-color:var(--p);color:var(--p);}

.page{max-width:760px;margin:0 auto;padding:32px 20px 80px;}

/* HERO BLOCK */
.hero{margin-bottom:20px;}
.hero-eyebrow{font-size:10px;letter-spacing:2.5px;color:var(--w3);text-transform:uppercase;margin-bottom:14px;}
.hero-main{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:20px;}
.hero-num{font-family:'DM Serif Display',serif;font-size:58px;line-height:1;letter-spacing:-2px;color:var(--w);}
.hero-num sup{font-size:22px;letter-spacing:0;vertical-align:super;color:var(--w2);}
.hero-meta{text-align:right;}
.hero-meta-period{font-size:11px;color:var(--w3);margin-bottom:6px;}
.hero-meta-net{font-size:13px;color:var(--p);font-weight:500;}
.hero-divider{height:1px;background:var(--line);margin-bottom:20px;}
.hero-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:0;}
.hstat{padding:0 20px;}
.hstat:first-child{padding-left:0;}
.hstat:last-child{padding-right:0;}
.hstat+.hstat{border-left:1px solid var(--line);}
.hstat-lbl{font-size:10px;letter-spacing:1.5px;color:var(--w3);text-transform:uppercase;margin-bottom:6px;}
.hstat-val{font-size:22px;font-weight:300;color:var(--w);}
.hstat-sub{font-size:10px;color:var(--w3);margin-top:3px;}

/* SECTION */
.sec{background:var(--s1);border:1px solid var(--line);border-radius:14px;padding:22px 24px;margin-bottom:14px;}
.sec-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;}
.sec-title{font-size:10px;letter-spacing:2px;color:var(--w3);text-transform:uppercase;}

/* TABS */
.tabs{display:flex;gap:2px;background:var(--s3);padding:3px;border-radius:8px;}
.tab{font-size:11px;padding:4px 12px;border-radius:6px;border:none;background:transparent;color:var(--w3);cursor:pointer;font-family:inherit;letter-spacing:0.3px;transition:all .15s;}
.tab.on{background:var(--s2);color:var(--w);border:1px solid var(--line2);}

/* CHART */
.chart-box{position:relative;width:100%;height:200px;}
.chart-box-lg{position:relative;width:100%;height:230px;}

/* LEGEND */
.leg-row{display:flex;gap:16px;margin-bottom:12px;}
.leg{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--w3);}
.leg-line{width:18px;height:2px;border-radius:1px;}

/* FORECAST */
.sc-pills{display:flex;gap:6px;margin-bottom:16px;flex-wrap:wrap;}
.sc-pill{font-size:11px;padding:5px 13px;border-radius:6px;border:1px solid var(--line2);background:transparent;color:var(--w3);cursor:pointer;font-family:inherit;letter-spacing:0.3px;transition:all .15s;}
.sc-pill.on{background:var(--p-dim2);color:var(--p);border-color:rgba(0,212,160,0.3);}
.fc-row{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin-top:16px;}
.fcc{border-radius:10px;padding:14px 16px;}
.fcc-mo{font-size:10px;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:6px;}
.fcc-val{font-size:18px;font-weight:300;margin-bottom:4px;}
.fcc-chg{font-size:11px;}
.fcc.pos{background:var(--p-dim);border:1px solid rgba(0,212,160,0.15);}
.fcc.pos .fcc-mo{color:rgba(0,212,160,0.5);}
.fcc.pos .fcc-val{color:var(--p);}
.fcc.pos .fcc-chg{color:rgba(0,212,160,0.6);}
.fcc.neg{background:var(--neg-dim);border:1px solid rgba(255,94,94,0.15);}
.fcc.neg .fcc-mo{color:rgba(255,94,94,0.5);}
.fcc.neg .fcc-val{color:var(--neg);}
.fcc.neg .fcc-chg{color:rgba(255,94,94,0.6);}

/* 2COL */
.two{display:grid;grid-template-columns:1fr 1fr;gap:14px;}

/* MEMBER GRID */
.pay-header{display:flex;align-items:baseline;justify-content:space-between;margin-bottom:14px;}
.pay-ym{font-size:12px;color:var(--p);font-weight:500;letter-spacing:0.3px;}
.pay-count{font-size:11px;color:var(--w3);}
.pay-count span{color:var(--w2);}
.mgrid{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;}
.mc{display:flex;flex-direction:column;align-items:center;gap:5px;padding:11px 4px;border-radius:10px;border:1px solid var(--line);}
.mc.paid{background:var(--p-dim);border-color:rgba(0,212,160,0.2);}
.mc.unpaid{background:var(--s3);}
.mc-av{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:500;}
.mc.paid .mc-av{background:rgba(0,212,160,0.2);color:var(--p);}
.mc.unpaid .mc-av{background:rgba(255,255,255,0.06);color:var(--w3);}
.mc-name{font-size:10px;letter-spacing:-0.2px;}
.mc.paid .mc-name{color:var(--w);}
.mc.unpaid .mc-name{color:var(--w3);}
.mc-tag{font-size:9px;padding:2px 7px;border-radius:4px;font-weight:500;letter-spacing:0.3px;}
.mc.paid .mc-tag{background:rgba(0,212,160,0.15);color:var(--p);}
.mc.unpaid .mc-tag{background:rgba(255,255,255,0.05);color:var(--w4);}

/* TIMELINE */
.tl{display:flex;flex-direction:column;}
.tl-item{display:flex;gap:12px;padding:10px 0;border-bottom:1px solid var(--line);}
.tl-item:last-child{border-bottom:none;}
.tl-icon{width:28px;height:28px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0;margin-top:1px;}
.tl-icon.in{background:var(--p-dim2);}
.tl-icon.out{background:var(--neg-dim);}
.tl-body{flex:1;min-width:0;}
.tl-row1{display:flex;justify-content:space-between;align-items:baseline;gap:8px;margin-bottom:3px;}
.tl-memo{font-size:12px;color:var(--w);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;}
.tl-amt{font-size:13px;font-weight:500;flex-shrink:0;}
.tl-amt.in{color:var(--p);}
.tl-amt.out{color:var(--neg);}
.tl-row2{display:flex;align-items:center;gap:7px;}
.tl-date{font-size:10px;color:var(--w3);}
.tl-tag{font-size:9px;padding:2px 7px;border-radius:4px;letter-spacing:0.2px;}
.tag-모임비{background:rgba(255,255,255,0.06);color:var(--w3);}
.tag-회비{background:var(--p-dim);color:var(--p);}
.tag-경조사비{background:rgba(248,190,100,0.1);color:#f8be64;}
.tag-강연비{background:rgba(168,139,250,0.1);color:#a88bfa;}
.tag-경품비{background:var(--neg-dim);color:var(--neg);}
.tag-지각비{background:rgba(0,212,160,0.07);color:rgba(0,212,160,0.6);}
.tag-이자{background:rgba(255,255,255,0.05);color:var(--w3);}
.tag-기타{background:rgba(255,255,255,0.05);color:var(--w3);}

@media(max-width:580px){
  .two{grid-template-columns:1fr;}
  .hero-num{font-size:44px;}
  .hstat-val{font-size:18px;}
  .mgrid{grid-template-columns:repeat(3,1fr);}
  .fc-row{grid-template-columns:repeat(3,1fr);}
}
</style>
</head>
<body>

<nav>
  <span class="nav-logo">사<em>군</em>자</span>
  <div class="nav-right">
    <span class="nav-time" id="nav-time">—</span>
    <button class="nav-btn" onclick="init()">REFRESH</button>
  </div>
</nav>

<div class="page">

  <!-- HERO -->
  <div class="hero">
    <div class="hero-eyebrow">Fund Balance</div>
    <div class="hero-main">
      <div class="hero-num"><sup>₩</sup><span id="h-bal">—</span></div>
      <div class="hero-meta">
        <div class="hero-meta-period" id="h-period">—</div>
        <div class="hero-meta-net" id="h-net">—</div>
      </div>
    </div>
    <div class="hero-divider"></div>
    <div class="hero-stats">
      <div class="hstat">
        <div class="hstat-lbl">Total In</div>
        <div class="hstat-val" id="h-in">—</div>
        <div class="hstat-sub">회비 + 지각비 + 이자</div>
      </div>
      <div class="hstat">
        <div class="hstat-lbl">Total Out</div>
        <div class="hstat-val" id="h-out">—</div>
        <div class="hstat-sub">모임비 + 경조사비 외</div>
      </div>
      <div class="hstat">
        <div class="hstat-lbl">월 평균 입금</div>
        <div class="hstat-val" id="h-avg">—</div>
        <div class="hstat-sub">최근 6개월 기준</div>
      </div>
    </div>
  </div>

  <!-- 월별 차트 -->
  <div class="sec">
    <div class="sec-head">
      <span class="sec-title">Monthly Flow</span>
      <div class="tabs">
        <button class="tab on" onclick="setMode('bar',this)">막대</button>
        <button class="tab" onclick="setMode('line',this)">추세</button>
        <button class="tab" onclick="setMode('cum',this)">누적</button>
      </div>
    </div>
    <div class="leg-row">
      <span class="leg"><span class="leg-line" style="background:#00D4A0"></span>입금</span>
      <span class="leg"><span class="leg-line" style="background:#ff5e5e"></span>출금</span>
      <span class="leg"><span class="leg-line" style="background:rgba(255,255,255,0.35);border-top:1px dashed rgba(255,255,255,0.35);height:0"></span>누적잔액</span>
    </div>
    <div class="chart-box"><canvas id="mainC"></canvas></div>
  </div>

  <!-- 전망 -->
  <div class="sec">
    <div class="sec-head"><span class="sec-title">Balance Forecast</span></div>
    <div id="fc-desc" style="font-size:10px;color:rgba(255,255,255,0.28);margin-bottom:14px;letter-spacing:0.2px;line-height:1.6;"></div>
    <div class="chart-box-lg"><canvas id="fcC"></canvas></div>
    <div class="fc-row" id="fc-cards"></div>
  </div>

  <!-- 멤버 + 타임라인 -->
  <div class="two">
    <div class="sec" style="margin-bottom:0">
      <div class="sec-head">
        <span class="sec-title">Monthly Payment</span>
        <select id="pay-select" onchange="onPaySelect(this.value)" style="background:#1c1c1c;border:1px solid rgba(255,255,255,0.14);color:rgba(255,255,255,0.7);padding:4px 10px;border-radius:6px;font-size:11px;font-family:inherit;cursor:pointer;outline:none;"></select>
      </div>
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;">
        <span class="pay-count" id="pay-count">—</span>
        <div onclick="navigator.clipboard.writeText('79790864002')" title="클릭하면 계좌번호 복사" style="cursor:pointer;background:rgba(0,212,160,0.08);border:1px solid rgba(0,212,160,0.2);border-radius:8px;padding:5px 11px;text-align:right;">
          <div style="font-size:9px;color:rgba(0,212,160,0.6);letter-spacing:0.8px;margin-bottom:2px;">카카오뱅크 · 조현민</div>
          <div style="font-size:12px;font-weight:500;color:#00D4A0;letter-spacing:1px;">7979-0864-002</div>
        </div>
      </div>
      <div class="mgrid" id="mgrid"></div>
    </div>
    <div class="sec" style="margin-bottom:0">
      <div class="sec-head"><span class="sec-title">Recent Transactions</span></div>
      <div class="tl" id="tl"></div>
    </div>
  </div>

</div>

<script>
const fmtFull = n => Math.round(Math.abs(n)).toLocaleString('ko-KR');
const fmtK    = n => { const a=Math.abs(Math.round(n)); return a>=1000000?(a/1000000).toFixed(1)+'M':a>=10000?(a/10000).toFixed(0)+'만':a.toLocaleString(); };

let D=null, MC=null, FC=null, mode='bar';

async function init(){
  document.getElementById('nav-time').textContent='…';
  const r=await fetch('/api/data'); D=await r.json(); render();
}

function render(){
  const s=D.summary;
  document.getElementById('nav-time').textContent=s.updated;
  document.getElementById('h-bal').textContent=fmtFull(s.balance);
  const real=D.monthly.filter(m=>m.income>0||m.expense>0);
  document.getElementById('h-period').textContent=real.length?`${real[0].month} — ${real[real.length-1].month}`:'';
  const net=s.total_in-s.total_out;
  document.getElementById('h-net').textContent=(net>=0?'▲ +':'▼ ')+fmtK(net)+'원 순증';
  document.getElementById('h-in').textContent=fmtK(s.total_in)+'원';
  document.getElementById('h-out').textContent=fmtK(s.total_out)+'원';
  document.getElementById('h-avg').textContent=fmtK(s.avg6)+'원';
  buildMain(); buildFc(); buildPaySelect(); buildMembers(); buildTL();
}

function buildMain(){
  if(MC) MC.destroy();
  const real=D.monthly.filter(m=>m.income>0||m.expense>0);
  const labels=real.map(m=>m.month);
  let ds;
  if(mode==='cum'){
    ds=[{type:'line',data:real.map(m=>m.cumulative),borderColor:'rgba(255,255,255,0.5)',backgroundColor:'rgba(255,255,255,0.03)',fill:true,tension:0.4,pointRadius:0,borderWidth:1.5,borderDash:[4,3]}];
  } else {
    ds=[
      {type:mode,label:'입금',data:real.map(m=>m.income),backgroundColor:mode==='bar'?'rgba(0,212,160,0.65)':'transparent',borderColor:'#00D4A0',borderRadius:mode==='bar'?3:0,borderWidth:1.5,tension:0.4,pointRadius:0,fill:false},
      {type:mode,label:'출금',data:real.map(m=>m.expense),backgroundColor:mode==='bar'?'rgba(255,94,94,0.55)':'transparent',borderColor:'#ff5e5e',borderRadius:mode==='bar'?3:0,borderWidth:1.5,tension:0.4,pointRadius:0,fill:false},
    ];
  }
  MC=new Chart(document.getElementById('mainC'),{
    data:{labels,datasets:ds},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmtK(c.raw)+'원'}}},
      scales:{
        x:{ticks:{color:'rgba(255,255,255,0.25)',font:{size:10},maxRotation:45,autoSkip:true,maxTicksLimit:10},grid:{display:false},border:{display:false}},
        y:{ticks:{color:'rgba(255,255,255,0.25)',font:{size:10},callback:v=>fmtK(v)},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}}
      }}
  });
}

function iqrFilter(arr){
  const sorted=[...arr].sort((a,b)=>a-b);
  const q1=sorted[Math.floor(sorted.length*0.25)];
  const q3=sorted[Math.floor(sorted.length*0.75)];
  const iqr=q3-q1;
  const upper=q3+1.5*iqr;
  return arr.filter(v=>v<=upper);
}

function buildFc(){
  const real=D.monthly.filter(m=>m.income>0||m.expense>0);
  const li=real.filter(m=>m.income>0).slice(-6);
  const ai=Math.round(li.reduce((a,m)=>a+m.income,0)/Math.max(li.length,1));
  const lo=real.filter(m=>m.expense>0).slice(-14);
  const loVals=lo.map(m=>m.expense);
  const loFiltered=iqrFilter(loVals);
  const ao=Math.round(loFiltered.reduce((a,v)=>a+v,0)/Math.max(loFiltered.length,1));
  const excluded=loVals.length-loFiltered.length;
  const cb=D.summary.balance;

  // 5개 시나리오 — 에메랄드(절약) ↔ 레드(지출과다) 그라데이션
  const scenarios=[
    {label:'지출 -20%', in:ai, out:Math.round(ao*0.8), color:'rgba(0,212,160,1.0)',  dim:'rgba(0,212,160,0.07)',  dash:[]},
    {label:'지출 -10%', in:ai, out:Math.round(ao*0.9), color:'rgba(0,212,160,0.45)', dim:'rgba(0,212,160,0.03)',  dash:[4,3]},
    {label:'현재 추세',  in:ai, out:ao,                 color:'rgba(255,255,255,0.55)',dim:'rgba(255,255,255,0.02)',dash:[5,4]},
    {label:'지출 +10%', in:ai, out:Math.round(ao*1.1), color:'rgba(255,94,94,0.45)', dim:'rgba(255,94,94,0.03)',  dash:[4,3]},
    {label:'지출 +20%', in:ai, out:Math.round(ao*1.2), color:'rgba(255,94,94,1.0)',  dim:'rgba(255,94,94,0.07)',  dash:[]},
  ];

  // 각 시나리오별 12개월 포인트 계산
  const hist=real.slice(-5);
  const lastM=hist[hist.length-1]?.month||'';
  const lm=lastM.match(/(\d+)년 (\d+)월/);
  let baseYY=lm?parseInt(lm[1]):26, baseMM=lm?parseInt(lm[2]):5;
  const fcLabels=[];
  for(let i=1;i<=12;i++){
    let mm=baseMM+i,yy=baseYY+Math.floor((mm-1)/12);
    mm=((mm-1)%12)+1;
    fcLabels.push(`${yy%100}년 ${mm}월`);
  }
  const allL=[...hist.map(m=>m.month),...fcLabels];
  const histConnect=hist[hist.length-1].cumulative;

  const datasets=[
    // 실제 이력
    {label:'실제',data:[...hist.map(m=>m.cumulative),...new Array(12).fill(null)],
     borderColor:'rgba(255,255,255,0.4)',backgroundColor:'rgba(255,255,255,0.02)',
     fill:true,tension:0.4,pointRadius:0,borderWidth:2}
  ];

  scenarios.forEach(sc=>{
    const pts=[]; let bal=cb;
    for(let i=0;i<12;i++){ bal+=sc.in-sc.out; pts.push(Math.round(bal)); }
    datasets.push({
      label:sc.label,
      data:[...new Array(4).fill(null), histConnect, ...pts],
      borderColor:sc.color, backgroundColor:sc.dim,
      fill:false, tension:0.4, pointRadius:0, borderWidth:1.5,
      borderDash:sc.dash
    });
  });

  if(FC)FC.destroy();
  FC=new Chart(document.getElementById('fcC'),{
    type:'line',
    data:{labels:allL,datasets},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{
          mode:'index',intersect:false,
          callbacks:{
            label:c=>c.raw!=null?` ${c.dataset.label}: ${fmtK(c.raw)}원`:null,
            filter:c=>c.raw!=null
          }
        }
      },
      scales:{
        x:{ticks:{color:'rgba(255,255,255,0.25)',font:{size:10},maxRotation:45},grid:{display:false},border:{display:false}},
        y:{ticks:{color:'rgba(255,255,255,0.25)',font:{size:10},callback:v=>fmtK(v)},grid:{color:'rgba(255,255,255,0.04)'},border:{display:false}}
      }}
  });

  // 설명 텍스트
  const desc=document.getElementById('fc-desc');
  if(desc) desc.innerHTML=
    `입금 기준: 최근 6개월 평균 <b style="color:rgba(255,255,255,0.7)">${fmtK(ai)}원</b> &nbsp;·&nbsp; `+
    `지출 기준: 최근 12개월 IQR 평균 <b style="color:rgba(255,255,255,0.7)">${fmtK(ao)}원</b> (비경상 ${excluded}개월 제외)`;

  // 카드: 5개 시나리오 × 12개월 후 잔액
  const fcc=document.getElementById('fc-cards');
  fcc.innerHTML='';
  // 카드: 에메랄드↔레드 그라데이션 배경
  const cardStyles=[
    {bg:'rgba(0,212,160,0.18)', border:'rgba(0,212,160,0.35)', mo:'rgba(0,212,160,0.55)', val:'#00D4A0', chg:'rgba(0,212,160,0.7)'},
    {bg:'rgba(0,212,160,0.07)', border:'rgba(0,212,160,0.18)', mo:'rgba(0,212,160,0.4)',  val:'rgba(0,212,160,0.8)', chg:'rgba(0,212,160,0.5)'},
    {bg:'rgba(255,255,255,0.04)',border:'rgba(255,255,255,0.12)',mo:'rgba(255,255,255,0.3)',val:'rgba(255,255,255,0.75)',chg:'rgba(255,255,255,0.4)'},
    {bg:'rgba(255,94,94,0.07)',  border:'rgba(255,94,94,0.18)', mo:'rgba(255,94,94,0.4)',  val:'rgba(255,94,94,0.8)', chg:'rgba(255,94,94,0.5)'},
    {bg:'rgba(255,94,94,0.18)',  border:'rgba(255,94,94,0.35)', mo:'rgba(255,94,94,0.55)', val:'#ff5e5e', chg:'rgba(255,94,94,0.7)'},
  ];
  scenarios.forEach((sc,i)=>{
    const bal12=Math.round(cb+(sc.in-sc.out)*12);
    const chg=bal12-cb;
    const isP=chg>=0;
    const netMo=sc.in-sc.out;
    const cs=cardStyles[i];
    fcc.innerHTML+=`<div class="fcc" style="background:${cs.bg};border:1px solid ${cs.border};">
      <div class="fcc-mo" style="color:${cs.mo}">${sc.label}</div>
      <div class="fcc-val" style="color:${cs.val}">${fmtK(bal12)}원</div>
      <div class="fcc-chg" style="color:${cs.chg}">${isP?'▲':'▼'} ${fmtK(Math.abs(chg))}</div>
      <div style="font-size:10px;margin-top:4px;color:${cs.chg};opacity:0.8;">월 ${isP?'+':''}${fmtK(netMo)}</div>
    </div>`;
  });
}

function buildPaySelect(){
  const sel=document.getElementById('pay-select');
  sel.innerHTML='';
  // 최신 월이 맨 위로
  [...D.pay_months].reverse().forEach(ym=>{
    const opt=document.createElement('option');
    opt.value=ym; opt.textContent=ym;
    if(ym===D.summary.latest_ym) opt.selected=true;
    sel.appendChild(opt);
  });
}

function buildMembers(ym){
  ym=ym||D.summary.latest_ym;
  const ms=D.payment_by_month[ym]||[];
  const pc=ms.filter(m=>m.paid).length;
  document.getElementById('pay-count').innerHTML=`<span style="color:#00D4A0;font-weight:500">${pc}명</span> / ${ms.length}명 납부완료`;
  const g=document.getElementById('mgrid'); g.innerHTML='';
  // 납부자 먼저, 미납자 뒤에
  const sorted=[...ms].sort((a,b)=>b.paid-a.paid);
  sorted.forEach(m=>{
    const c=m.paid?'paid':'unpaid';
    g.innerHTML+=`<div class="mc ${c}">
      <div class="mc-av">${m.name[0]}</div>
      <div class="mc-name">${m.name}</div>
      <span class="mc-tag">${m.paid?'납부':'미납'}</span>
    </div>`;
  });
}

function onPaySelect(ym){ buildMembers(ym); }

function buildTL(){
  const t=document.getElementById('tl'); t.innerHTML='';
  D.recent.forEach(r=>{
    const cls=r.is_income?'in':'out',sign=r.is_income?'+':'-';
    const icon=r.is_income?'↑':'↓';
    t.innerHTML+=`<div class="tl-item">
      <div class="tl-icon ${cls}">${icon}</div>
      <div class="tl-body">
        <div class="tl-row1">
          <span class="tl-memo">${r.memo||'—'}</span>
          <span class="tl-amt ${cls}">${sign}${fmtFull(r.amount)}원</span>
        </div>
        <div class="tl-row2">
          <span class="tl-date">${r.date}</span>
          <span class="tl-tag tag-${r.type}">${r.type}</span>
        </div>
      </div>
    </div>`;
  });
}

function setMode(m,b){mode=m;document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));b.classList.add('on');buildMain();}


init();
</script>
</body>
</html>'''

class Handler(BaseHTTPRequestHandler):
    def log_message(self,f,*a): pass
    def do_HEAD(self):
        p=urlparse(self.path).path
        if p=='/api/data':
            self.send_response(200);self.send_header('Content-Type','application/json;charset=utf-8');self.end_headers()
        else:
            self.send_response(200);self.send_header('Content-Type','text/html;charset=utf-8');self.end_headers()
    def do_GET(self):
        p=urlparse(self.path).path
        if p=='/api/data':
            try:
                body=json.dumps(build_data(),ensure_ascii=False).encode('utf-8')
                self.send_response(200);self.send_header('Content-Type','application/json;charset=utf-8');self.send_header('Content-Length',len(body));self.end_headers();self.wfile.write(body)
            except Exception as e:
                import traceback
                body=json.dumps({'error':str(e),'trace':traceback.format_exc()}).encode()
                self.send_response(500);self.send_header('Content-Type','application/json');self.end_headers();self.wfile.write(body)
        else:
            body=HTML.encode('utf-8')
            self.send_response(200);self.send_header('Content-Type','text/html;charset=utf-8');self.send_header('Content-Length',len(body));self.end_headers();self.wfile.write(body)

if __name__=='__main__':
    import socket, webbrowser, threading, time, sys

    # Render는 PORT 환경변수로 포트를 지정함
    env_port = os.environ.get('PORT')
    if env_port:
        port = int(env_port)
        server = HTTPServer(('0.0.0.0', port), Handler)
        print(f'사군자 대시보드 실행 중: http://0.0.0.0:{port}')
        try: server.serve_forever()
        except KeyboardInterrupt: print('\n서버 종료.')
    else:
        # 로컬 실행: 빈 포트 자동 탐색 + 브라우저 자동 열기
        def find_free_port(start=8000):
            for p in range(start, start+10):
                try:
                    s=socket.socket(); s.bind(('',p)); s.close(); return p
                except: continue
            return start

        start_port = 8000
        for i, arg in enumerate(sys.argv):
            if arg == '--port' and i+1 < len(sys.argv):
                start_port = int(sys.argv[i+1])

        port = find_free_port(start_port)
        url  = f'http://localhost:{port}'

        try:
            server = HTTPServer(('0.0.0.0', port), Handler)
        except OSError:
            print(f'포트 {port} 사용 중. 브라우저에서 {url} 을 열어보세요.')
            webbrowser.open(url); exit()

        threading.Thread(target=lambda: (time.sleep(1.2), webbrowser.open(url)), daemon=True).start()

        print(f"""
  ╔══════════════════════════════════════╗
  ║   사군자 회비 대시보드 실행 중        ║
  ╠══════════════════════════════════════╣
  ║  주소:  {url:<28} ║
  ║  종료:  Ctrl+C  또는 창 닫기         ║
  ╚══════════════════════════════════════╝
""")
        try: server.serve_forever()
        except KeyboardInterrupt: print('\n  서버 종료.')
