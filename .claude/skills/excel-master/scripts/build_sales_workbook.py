"""MBE 영업 대시보드 엑셀 빌더 (템플릿 겸 예제).

입력: ytd.json(대시보드 마스터 DB) 형식 -> 7개 시트 워크북
  python build_sales_workbook.py --input ytd.json --out MBE_영업분석.xlsx
다른 데이터(재고 등)는 이 파일의 시트 함수를 복사해 열 구성만 바꿔 쓰세요.
2단계 빌드(수식 값 미리 저장)까지 자동으로 합니다.
"""
import argparse, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlkit import Kit, P, NF, rc, cn
import verify_xlsx

CH_ORDER = ['백화점', '아울렛', '대리점', '직영점', '온라인', '기타']
AUTHOR = '영업부 정병식 과장'


def stat(D):
    """요약 문구용 수치(파이썬 계산, 정적 문구)."""
    S = D['stores']
    s = lambda k, f=lambda x: True: sum(x[k] for x in S if f(x))
    ytd, py = s('ytd'), s('ytd_py')
    tg = s('ytd_tgt'); yt = sum(x['ytd'] for x in S if x['ytd_tgt'])
    mt = s('mtd_tgt'); mm = sum(x['mtd'] for x in S if x['mtd_tgt'])
    chs = {}
    for c in CH_ORDER:
        rs = [x for x in S if x['channel'] == c]
        t = sum(x['ytd_tgt'] for x in rs)
        if t:
            chs[c] = sum(x['ytd'] for x in rs if x['ytd_tgt']) / t
    bad = [x for x in S if x['active'] == '운영' and any(k in (x.get('flags') or '') for k in ('미만', '이하'))]
    return dict(ytd=ytd, py=py, yoy=ytd / py - 1, ach=yt / tg if tg else None, mtd=s('mtd'), mpy=s('mtd_py'),
                myoy=s('mtd') / s('mtd_py') - 1 if s('mtd_py') else None, mach=mm / mt if mt else None, chs=chs, bad=bad)


def build(path, D, cached=None, author=AUTHOR):
    K = Kit(path, cached)
    nm, day = D['month'], D['month_day']
    S = D['stores']
    ops = sorted([x for x in S if x['active'] == '운영'], key=lambda x: -x['ytd'])
    cls = sorted([x for x in S if x['active'] != '운영'], key=lambda x: -x['ytd'])
    rows = ops + cls
    n, n_op, N = len(rows), len(ops), len(rows) + 1
    asof = f"{D['asof'][:4]}년 {int(D['asof'][5:7])}월 {int(D['asof'][8:10])}일 마감 ({D['period']})"
    chans = [c for c in CH_ORDER if any(x['channel'] == c for x in S)]
    ST = stat(D)

    # ---- 원데이터 열 지도 ----
    M0, P0, T0 = 5, 5 + nm, 5 + 2 * nm
    X = T0 + nm
    C = dict(code=0, name=1, ch=2, vendor=3, act=4, ytd_tgt=X, mtd_tgt=X + 1, tag=X + 2, net=X + 3, stock=X + 4, flags=X + 5,
             ytd=X + 6, ytd_py=X + 7, yoy=X + 8, ach=X + 9, mtd=X + 10, mtd_py=X + 11, mtd_yoy=X + 12, mtd_ach=X + 13,
             disc=X + 14, st=X + 15)
    for j in range(nm):
        C[f'm{j}'], C[f'p{j}'], C[f't{j}'] = M0 + j, P0 + j, T0 + j

    def R(k):                                   # 전체 범위
        return f"'원데이터'!${cn(C[k])}$2:${cn(C[k])}${N}"

    def Rop(k):                                 # 운영 매장 범위(앞쪽 n_op행)
        return f"'원데이터'!${cn(C[k])}$2:${cn(C[k])}${n_op + 1}"

    def cell(k, i):                             # i번째 데이터 행의 셀
        return f"'원데이터'!${cn(C[k])}${i + 2}"

    # ==================== 시트 생성(탭 순서) ====================
    dash = K.sheet('대시보드', tab=P['navy'], widths=[(0, 0, 1.5), (1, 12, 12.6)])
    wch = K.sheet('유통별', tab=P['navy2'], widths=[(0, 0, 1.5), (1, 1, 12), (2, 9, 12.5), (10, 10, 20)])
    wmo = K.sheet('월별추이', tab=P['navy2'], widths=[(0, 0, 1.5), (1, 1, 14), (2, nm + 2, 9.5)])
    wst = K.sheet('매장별', tab=P['steel'], widths=[(0, 0, 1.5), (1, 1, 6), (2, 2, 26), (3, 4, 9), (5, 11, 11), (12, 12, 18), (13, 13, 34)])
    wan = K.sheet('분석', tab=P['steel'], widths=[(0, 0, 1.5), (1, 12, 12.6)])
    wr = K.sheet('원데이터', tab=P['grey'], zoom=90, landscape=True, hide_grid=False)
    wh = K.sheet('사용설명', tab=P['gold'], widths=[(0, 0, 1.5), (1, 1, 18), (2, 2, 90)])

    # ==================== 원데이터 ====================
    hin = K.fmt(bold=True, align='center', bg_color=P['grey'], border=1, border_color=P['line'], text_wrap=True)
    hcal = K.fmt(bold=True, align='center', bg_color=P['sky'], border=1, border_color=P['navy2'], text_wrap=True)
    heads = (['코드', '매장명', '유통', '거래처', '운영'] + [f'26년 {i + 1}월' for i in range(nm)] + [f'25년 {i + 1}월' for i in range(nm)]
             + [f'목표 {i + 1}월' for i in range(nm)] + ['누계목표(진도)', '당월목표(진도)', 'TAG금액', '실판매', '재고(현판가)', '특이사항']
             + ['누계', '전년 누계', '누계 신장율', '누계 달성률', '당월', '전년 같은 기간', '당월 신장율', '당월 진도달성', '할인율', '판매율'])
    wr.set_row(0, 34)
    for j, h in enumerate(heads):
        wr.write(0, j, h, hin if j < C['ytd'] else hcal)
    f_txt = K.fmt(); f_m1 = K.fmt(num_format=NF['m1']); f_yoy = K.fmt(num_format=NF['yoy']); f_pct = K.fmt(num_format=NF['pct'])
    f_calc = K.fmt(num_format=NF['m1'], bg_color=P['mist'])
    f_cyoy = K.fmt(num_format=NF['yoy'], bg_color=P['mist']); f_cpct = K.fmt(num_format=NF['pct'], bg_color=P['mist'])
    for i, x in enumerate(rows):
        r = i + 1; e = i + 2
        for k, v in (('code', x['code']), ('name', x['name']), ('ch', x['channel']), ('vendor', x['vendor'] or ''), ('act', x['active'])):
            wr.write_string(r, C[k], v, f_txt)
        for j in range(nm):
            wr.write_number(r, C[f'm{j}'], x['m'][j], f_m1)
            wr.write_number(r, C[f'p{j}'], x['p'][j], f_m1)
            wr.write_number(r, C[f't{j}'], x['t'][j], f_m1)
        for k, key in (('ytd_tgt', 'ytd_tgt'), ('mtd_tgt', 'mtd_tgt'), ('tag', 'tag_amt'), ('net', 'net_amt'), ('stock', 'stock_amt')):
            wr.write_number(r, C[k], x[key] or 0, f_m1)
        wr.write_string(r, C['flags'], x.get('flags') or '', f_txt)
        L = lambda k: f"{cn(C[k])}{e}"
        K.fx(wr, r, C['ytd'], f"=SUM({L('m0')}:{L(f'm{nm - 1}')})", f_calc)
        K.fx(wr, r, C['ytd_py'], f"=SUM({L('p0')}:{L(f'p{nm - 1}')})", f_calc)
        K.fx(wr, r, C['yoy'], f'=IF({L("ytd_py")}>0,{L("ytd")}/{L("ytd_py")}-1,"")', f_cyoy)
        K.fx(wr, r, C['ach'], f'=IF({L("ytd_tgt")}>0,{L("ytd")}/{L("ytd_tgt")},"")', f_cpct)
        K.fx(wr, r, C['mtd'], f"={L(f'm{nm - 1}')}", f_calc)
        K.fx(wr, r, C['mtd_py'], f"={L(f'p{nm - 1}')}", f_calc)
        K.fx(wr, r, C['mtd_yoy'], f'=IF({L("mtd_py")}>0,{L("mtd")}/{L("mtd_py")}-1,"")', f_cyoy)
        K.fx(wr, r, C['mtd_ach'], f'=IF({L("mtd_tgt")}>0,{L("mtd")}/{L("mtd_tgt")},"")', f_cpct)
        K.fx(wr, r, C['disc'], f'=IF({L("tag")}>0,1-{L("net")}/{L("tag")},"")', f_cpct)
        K.fx(wr, r, C['st'], f'=IF({L("tag")}+{L("stock")}>0,{L("tag")}/({L("tag")}+{L("stock")}),"")', f_cpct)
    wr.freeze_panes(1, 2); wr.autofilter(0, 0, n, len(heads) - 1)
    wr.set_column(0, 0, 8); wr.set_column(1, 1, 26); wr.set_column(2, 4, 9); wr.set_column(5, X - 1, 8.5)
    wr.set_column(X, X + 4, 11); wr.set_column(C['flags'], C['flags'], 30); wr.set_column(C['ytd'], C['st'], 10.5)
    wr.repeat_rows(0)

    # ==================== 유통별 ====================
    K.title(wch, 1, 1, 10, '유통별 실적 현황', asof)
    K.section(wch, 3, 1, 10, '1. 유통별 실적', '[단위:백만원, %]')
    hdr = ['유통', '운영 매장', f'{nm}월 매출', f'{nm}월 진도달성', f'{nm}월 전년비', '누계 매출', '누계 달성', '누계 전년비', '구성비', '월별 추이']
    nfs = [None, NF['qty'], NF['m0'], NF['pct'], NF['yoy'], NF['m0'], NF['pct'], NF['yoy'], NF['pct'], None]
    r0 = 5; hr = r0  # 표 머리글 행(0부터)
    data = []
    mtg, ytg = R('mtd_tgt'), R('ytd_tgt')
    for i, c in enumerate(chans):
        e = hr + 2 + i
        B = f'$B{e}'
        sif = lambda k, extra='': f"SUMIFS({R(k)},{R('ch')},{B}{extra})"
        data.append([c, f"=COUNTIFS({R('ch')},{B},{R('act')},\"운영\")",
                     f"={sif('mtd')}",
                     f"=IFERROR({sif('mtd', f',{mtg},\">0\"')}/{sif('mtd_tgt')},\"\")",
                     f"=IFERROR({sif('mtd')}/{sif('mtd_py')}-1,\"\")", f"={sif('ytd')}",
                     f"=IFERROR({sif('ytd', f',{ytg},\">0\"')}/{sif('ytd_tgt')},\"\")",
                     f"=IFERROR({sif('ytd')}/{sif('ytd_py')}-1,\"\")", f"=G{e}/G${hr + 2 + len(chans)}", None])
    et = hr + 2 + len(chans)
    data.append(['합  계', f"=SUM(C{hr + 2}:C{et - 1})", f"=SUM({R('mtd')})",
                 f"=SUMIFS({R('mtd')},{R('mtd_tgt')},\">0\")/SUM({R('mtd_tgt')})", f"=SUM({R('mtd')})/SUM({R('mtd_py')})-1",
                 f"=SUM({R('ytd')})", f"=SUMIFS({R('ytd')},{R('ytd_tgt')},\">0\")/SUM({R('ytd_tgt')})",
                 f"=SUM({R('ytd')})/SUM({R('ytd_py')})-1", f"=SUM(J{hr + 2}:J{et - 1})", None])
    last = K.table(wch, hr, 1, hdr, data, nfs, ['center'] + ['right'] * 8 + ['center'], header='navy', total_last=True, row_h=24, bold_cols=(0,))
    K.databar(wch, f'J{hr + 2}:J{et - 1}'); K.threshold(wch, f'E{hr + 2}:E{et}'); K.threshold(wch, f'H{hr + 2}:H{et}')
    # 월별 보조 데이터(숨김): 유통별 월 매출 -> 스파크라인
    hc = 12
    for j in range(nm):
        wch.write(hr, hc + j, f'{j + 1}월', K.fmt(font_size=8, font_color=P['sub']))
    for i, c in enumerate(chans):
        for j in range(nm):
            K.fx(wch, hr + 1 + i, hc + j, f"=SUMIFS({R(f'm{j}')},{R('ch')},$B{hr + 2 + i})", K.fmt(font_size=8, num_format='#,##0', font_color=P['sub']))
        K.spark(wch, f'K{hr + 2 + i}', f"'유통별'!{rc(hr + 1 + i, hc)}:{rc(hr + 1 + i, hc + nm - 1)}")
    for j in range(nm):
        K.fx(wch, hr + 1 + len(chans), hc + j, f"=SUM({R(f'm{j}')})", K.fmt(font_size=8, num_format='#,##0', font_color=P['sub']))
    K.spark(wch, f'K{et + 1}', f"'유통별'!{rc(hr + 1 + len(chans), hc)}:{rc(hr + 1 + len(chans), hc + nm - 1)}")
    wch.set_column(hc, hc + nm - 1, 6, None, {'hidden': True})
    nr = K.notes(wch, last + 2, 1, 10, [
        f'{nm}월은 1일~{day}일 실적이며 전년은 같은 기간 기준입니다.',
        '진도달성 = 실적 ÷ (월 목표 × 일 목표 진도). 목표가 없는 매장은 달성률 계산에서 제외합니다.',
        '달성률: 100% 이상 파랑 굵게, 90% 미만 빨강. 마이너스는 빨간 ▲.'], bullet='※')
    K.footer(wch, nr + 1, 10, author, 1)

    # ==================== 월별추이 ====================
    K.title(wmo, 1, 1, nm + 2, 'MB 월별 실적 트렌드', asof)
    K.section(wmo, 3, 1, nm + 2, '1. 월별 매출 (26년 · 25년 · 신장율)', '[단위:백만원, %]')
    mh = ['구  분'] + [f'{i + 1}월' for i in range(nm)] + ['누  계']
    hrM = 5
    hf = K.fmt(bold=True, align='center', bg_color=P['navy'], font_color=P['white'], border=1, border_color=P['line'])
    hfy = K.fmt(bold=True, align='center', bg_color=P['yellow'], border=1, border_color=P['line'])
    wmo.set_row(hrM, 26)
    for j, h in enumerate(mh):
        wmo.write(hrM, 1 + j, h, hfy if j == nm else hf)
    lab = K.fmt(bold=True, align='center', bg_color=P['grey2'], border=1, border_color=P['line'])
    def mrow(rr, name, fn, nf, tot):
        wmo.write(rr, 1, name, lab); wmo.set_row(rr, 22)
        for j in range(nm):
            fm = K.fmt(num_format=nf, align='right', border=1, border_color=P['line'], bold=(j == nm - 1 and nf != NF['yoy']),
                       **({'bg_color': '#FFF2CC'} if j == nm - 1 else {}))
            K.fx(wmo, rr, 2 + j, fn(j), fm)
        K.fx(wmo, rr, 2 + nm, tot, K.fmt(num_format=nf, align='right', border=1, border_color=P['line'], bold=True, bg_color=P['sky']))
    mrow(hrM + 1, '26년', lambda j: f"=SUM({R(f'm{j}')})", NF['m0'], f"=SUM({cn(2)}{hrM + 2}:{cn(1 + nm)}{hrM + 2})")
    mrow(hrM + 2, '25년', lambda j: f"=SUM({R(f'p{j}')})", NF['m0'], f"=SUM({cn(2)}{hrM + 3}:{cn(1 + nm)}{hrM + 3})")
    mrow(hrM + 3, '신장율', lambda j: f'=IF({cn(2 + j)}{hrM + 3}>0,{cn(2 + j)}{hrM + 2}/{cn(2 + j)}{hrM + 3}-1,"")', NF['yoy'],
         f'=IF({cn(2 + nm)}{hrM + 3}>0,{cn(2 + nm)}{hrM + 2}/{cn(2 + nm)}{hrM + 3}-1,"")')
    mrow(hrM + 4, '월 목표', lambda j: f"=SUM({R(f't{j}')})", NF['m0'], f"=SUM({cn(2)}{hrM + 5}:{cn(1 + nm)}{hrM + 5})")
    wmo.write(hrM + 5, 1, '추  이', lab)
    wmo.merge_range(hrM + 5, 2, hrM + 5, 1 + nm, '', K.fmt(border=1, border_color=P['line']))
    K.spark(wmo, f'C{hrM + 6}', f"'월별추이'!C{hrM + 2}:{cn(1 + nm)}{hrM + 2}", 'column')
    wmo.set_row(hrM + 5, 34)
    wmo.conditional_format(hrM + 3, 2, hrM + 3, 1 + nm, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': K.fmt(font_color=P['red'])})
    K.section(wmo, hrM + 8, 1, nm + 2, '2. 유통별 전년비 히트맵', '[단위:%]')
    h2 = hrM + 10
    wmo.set_row(h2, 24)
    for j, h in enumerate(['유  통'] + [f'{i + 1}월' for i in range(nm)] + ['누  계']):
        wmo.write(h2, 1 + j, h, hf)
    for i, c in enumerate(chans):
        rr = h2 + 1 + i; e = rr + 1
        wmo.write(rr, 1, c, lab); wmo.set_row(rr, 22)
        cf = K.fmt(num_format=NF['yoy'], align='center', border=1, border_color=P['line'])
        for j in range(nm):
            K.fx(wmo, rr, 2 + j, f"=IFERROR(SUMIFS({R(f'm{j}')},{R('ch')},$B{e})/SUMIFS({R(f'p{j}')},{R('ch')},$B{e})-1,\"\")", cf)
        K.fx(wmo, rr, 2 + nm, f"=IFERROR(SUMIFS({R('ytd')},{R('ch')},$B{e})/SUMIFS({R('ytd_py')},{R('ch')},$B{e})-1,\"\")",
             K.fmt(num_format=NF['yoy'], align='center', border=1, border_color=P['line'], bold=True))
    K.heat(wmo, f"C{h2 + 2}:{cn(2 + nm)}{h2 + 1 + len(chans)}", 'div')
    nr = K.notes(wmo, h2 + len(chans) + 2, 1, nm + 2, [
        f'{nm}월(노랑)은 1일~{day}일 진행 중이며 전년은 같은 기간입니다.',
        '히트맵: 파랑은 전년 대비 증가, 빨강은 감소(±30% 기준 색 농도).'], bullet='※')
    K.footer(wmo, nr + 1, nm + 2, author, 1)

    # ==================== 매장별 ====================
    K.title(wst, 1, 1, 13, '매장별 실적 (운영 매장)', asof)
    K.section(wst, 3, 1, 13, f'1. 매장별 누계 실적 — {n_op}개 매장, 누계 매출 순', '[단위:백만원, %]')
    sh = ['순위', '매장명', '유통', '거래처', f'{nm}월 매출', f'{nm}월 진도달성', '누계 매출', '누계 달성', '누계 전년비', '할인율', '판매율', '월별 추이', '특이사항']
    snf = [NF['qty'], None, None, None, NF['m1'], NF['pct'], NF['m1'], NF['pct'], NF['yoy'], NF['pct'], NF['pct'], None, None]
    sal = ['center', 'left', 'center', 'center', 'right', 'right', 'right', 'right', 'right', 'right', 'right', 'center', 'left']
    hrS = 5; srows = []
    for i in range(n_op):
        e = hrS + 2 + i
        srows.append([f'=RANK(H{e},$H${hrS + 2}:$H${hrS + 1 + n_op})', f"={cell('name', i)}&\"\"", f"={cell('ch', i)}&\"\"", f"={cell('vendor', i)}&\"\"",
                      f"={cell('mtd', i)}", f"={cell('mtd_ach', i)}", f"={cell('ytd', i)}", f"={cell('ach', i)}", f"={cell('yoy', i)}",
                      f"={cell('disc', i)}", f"={cell('st', i)}", None, f"={cell('flags', i)}&\"\""])
    lastS = K.table(wst, hrS, 1, sh, srows, snf, sal, header='navy', row_h=21, bold_cols=(1,))
    a, b = hrS + 2, hrS + 1 + n_op
    K.databar(wst, f'H{a}:H{b}'); K.threshold(wst, f'G{a}:G{b}'); K.threshold(wst, f'I{a}:I{b}')
    K.heat(wst, f'J{a}:J{b}', 'div'); K.heat(wst, f'L{a}:L{b}')
    for i in range(n_op):
        K.spark(wst, f'M{hrS + 2 + i}', f"'원데이터'!{rc(i + 1, M0)}:{rc(i + 1, M0 + nm - 1)}")
    wst.freeze_panes(hrS + 1, 3); wst.autofilter(hrS, 1, lastS, 13); wst.repeat_rows(hrS)
    nr = K.notes(wst, lastS + 2, 1, 13, ['순위는 작성 시점 누계 매출 기준 정렬이며 숫자는 원데이터와 수식으로 연결됩니다.',
                                        '머리글의 필터 화살표로 유통·거래처별로 골라 볼 수 있습니다.'], bullet='※')
    K.footer(wst, nr + 1, 13, author, 1)

    # ==================== 분석(재고·할인) ====================
    K.title(wan, 1, 1, 12, '할인율·판매율 분석', asof)
    K.section(wan, 3, 1, 6, '1. 매장별 할인율 × 판매율', '[단위:%]')
    K.section(wan, 3, 7, 12, '2. 확인 필요 매장 (자동 판정)', '')
    K.scatter_chart(wan, 'B6', '원데이터', (1, C['disc'], n_op, C['disc']), (1, C['st'], n_op, C['st']),
                    '할인율 (낮을수록 좋음)', '판매율 (높을수록 좋음)', w=560, h=330)
    bad = sorted([x for x in ST['bad'] if x['mtd_tgt']], key=lambda x: (x['mtd'] / x['mtd_tgt']))[:12]
    brow = []
    for x in bad:
        fl = [f for f in (x.get('flags') or '').split(' · ') if any(k in f for k in ('미만', '이하'))]
        brow.append([x['name'][:14], x['channel'], x['mtd'] / x['mtd_tgt'], x['ytd_yoy'], (fl[0] if fl else '')[:18]])
    lb = K.table(wan, 5, 7, ['매장', '유통', f'{nm}월 진도달성', '누계 전년비', '사유'], [r + [None] * 0 for r in brow],
                 [None, None, NF['pct'], NF['yoy'], None], ['left', 'center', 'right', 'right', 'left'], header='navy', row_h=21)
    wan.set_column(7, 7, 18); wan.set_column(11, 12, 14)
    K.threshold(wan, f'J7:J{lb + 1}')
    nr = K.notes(wan, 26, 1, 12, [
        '오른쪽 아래(할인율 높고 판매율 낮음)에 몰린 매장이 재고·할인 점검 1순위입니다.',
        '확인 필요 매장은 대시보드 마스터 DB의 판정 문구(미만·이하)가 붙은 운영 매장을 당월 진도달성이 낮은 순으로 보여줍니다(작성 시점 값).'], bullet='※')
    K.footer(wan, nr + 1, 12, author, 1)

    # ==================== 대시보드 ====================
    dash.set_portrait(); dash.fit_to_pages(1, 1)
    K.title(dash, 1, 1, 12, f"MBE 영업 대시보드 · {D['month']}월", asof)
    K.section(dash, 4, 1, 12, 'Ⅰ. 핵심 지표', '[단위:백만원, %]')
    yfm = '"전년 대비 "+0.0%;[Red]"전년 대비 "-0.0%;"전년 대비 -"'
    cards = [
        dict(label=f'{nm}월 매출 (1~{day}일)', value=f"=SUM({R('mtd')})", nf='#,##0', sub=f"=SUM({R('mtd')})/SUM({R('mtd_py')})-1", sub_nf=yfm, accent=P['gold']),
        dict(label=f'{nm}월 진도달성', value=f"=SUMIFS({R('mtd')},{R('mtd_tgt')},\">0\")/SUM({R('mtd_tgt')})", nf='0.0%',
             sub=f"=SUM({R('mtd_tgt')})", sub_nf='"진도 목표 "#,##0', accent=P['gold']),
        dict(label='누계 매출', value=f"=SUM({R('ytd')})", nf='#,##0', sub=f"=SUM({R('ytd')})/SUM({R('ytd_py')})-1", sub_nf=yfm, accent=P['navy']),
        dict(label='누계 달성', value=f"=SUMIFS({R('ytd')},{R('ytd_tgt')},\">0\")/SUM({R('ytd_tgt')})", nf='0.0%',
             sub=f"=SUM({R('ytd_tgt')})", sub_nf='"진도 목표 "#,##0', accent=P['navy']),
        dict(label='할인율', value=f"=1-SUM({R('net')})/SUM({R('tag')})", nf='0.0%', sub=f"=SUM({R('tag')})", sub_nf='"TAG "#,##0', accent=P['steel']),
        dict(label='판매율', value=f"=SUM({R('tag')})/(SUM({R('tag')})+SUM({R('stock')}))", nf='0.0%', sub=f"=SUM({R('stock')})",
             sub_nf='"재고(현판가) "#,##0', accent=P['steel']),
    ]
    K.kpi_cards(dash, 5, 1, cards, 2)
    K.section(dash, 9, 1, 6, 'Ⅱ. 월별 매출 추이 (올해·전년·목표)', '[백만원]')
    K.section(dash, 9, 7, 12, f'Ⅲ. 유통별 누계 달성률', '[%]')
    cats = (hrM, 2, hrM, 1 + nm)
    K.col_line_chart(dash, 'B11', '월별추이', cats, [('26년', (hrM + 1, 2, hrM + 1, 1 + nm))], lines=[('25년(전년)', (hrM + 2, 2, hrM + 2, 1 + nm))],
                     markers=[('월 목표', (hrM + 4, 2, hrM + 4, 1 + nm))], w=548, h=340)
    n_bar = sum(1 for c in chans if any(x['channel'] == c and x['ytd_tgt'] for x in S))
    K.bar_chart(dash, 'H11', '유통별', (hr + 1, 1, hr + n_bar, 1), (hr + 1, 7, hr + n_bar, 7), name='누계 달성', w=548, h=340,
                nf='0%', max_=1.2)
    K.section(dash, 29, 1, 6, 'Ⅳ. 누계 매출 TOP 10 매장', '[백만원, %]')
    K.section(dash, 29, 7, 12, 'Ⅴ. 확인 필요 매장', '[%]')
    t10 = []
    for i in range(10):
        t10.append([i + 1, f"={cell('name', i)}", None, f"={cell('ytd', i)}", f"={cell('ach', i)}", f"={cell('yoy', i)}"])
    th = K.fmt(bold=True, align='center', bg_color=P['navy'], font_color=P['white'], border=1, border_color=P['line'])
    dash.set_row(30, 24)
    for j, h in [(1, '순위'), (4, '누계 매출'), (5, '누계 달성'), (6, '전년비')]:
        dash.write(30, j, h, th)
    dash.merge_range(30, 2, 30, 3, '매장', th)
    for i, rw in enumerate(t10):
        rr = 31 + i; dash.set_row(rr, 21)
        z = lambda **k: K.fmt(border=1, border_color=P['line'], **({'bg_color': P['mist']} if i % 2 else {}), **k)
        dash.write(rr, 1, rw[0], z(align='center', bold=True))
        dash.merge_range(rr, 2, rr, 3, '', z(align='left', bold=True)); K.put(dash, rr, 2, rw[1], z(align='left', bold=True))
        K.put(dash, rr, 4, rw[3], z(num_format=NF['m0'], align='right'))
        K.put(dash, rr, 5, rw[4], z(num_format=NF['pct'], align='right'))
        K.put(dash, rr, 6, rw[5], z(num_format=NF['yoy'], align='right'))
    K.databar(dash, 'E32:E41'); K.threshold(dash, 'F32:F41')
    dash.set_row(30, 24)
    dash.merge_range(30, 7, 30, 8, '매장', th); dash.write(30, 9, '유통', th); dash.write(30, 10, f'{nm}월 진도', th)
    dash.merge_range(30, 11, 30, 12, '사유', th)
    for i in range(10):
        rr = 31 + i
        z = lambda **k: K.fmt(border=1, border_color=P['line'], **({'bg_color': P['mist']} if i % 2 else {}), **k)
        if i < len(brow):
            b = brow[i]
            dash.merge_range(rr, 7, rr, 8, b[0], z(align='left'))
            dash.write(rr, 9, b[1], z(align='center')); dash.write(rr, 10, b[2], z(num_format=NF['pct'], align='right'))
            dash.merge_range(rr, 11, rr, 12, b[4], z(align='left', font_size=9))
        else:
            dash.merge_range(rr, 7, rr, 8, '', z()); dash.write_blank(rr, 9, None, z()); dash.write_blank(rr, 10, None, z())
            dash.merge_range(rr, 11, rr, 12, '', z())
    K.threshold(dash, 'K32:K41')
    K.section(dash, 42, 1, 12, 'Ⅵ. 핵심 해석', '')
    lowc = min(ST['chs'], key=ST['chs'].get); highc = max(ST['chs'], key=ST['chs'].get)
    lines = [
        f"누계 매출 {ST['ytd']:,.0f}백만원, 전년 대비 {ST['yoy'] * 100:+.1f}%, 누계 목표(진도) 대비 {ST['ach'] * 100:.1f}% 입니다.",
        f"{nm}월(1~{day}일) 매출 {ST['mtd']:,.0f}백만원으로 진도 대비 {ST['mach'] * 100:.1f}%, 전년 같은 기간 대비 {ST['myoy'] * 100:+.1f}% 입니다.",
        f"유통별 누계 달성률은 {highc} {ST['chs'][highc] * 100:.1f}%가 가장 높고 {lowc} {ST['chs'][lowc] * 100:.1f}%가 가장 낮으며, 확인 필요 매장은 {len(ST['bad'])}개입니다.",
    ]
    nr = K.notes(dash, 43, 1, 12, lines, '▶', height=24)
    mbb = next((b for b in D.get('brands', []) if b.get('brand') == 'MB'), None)
    n1 = '※ 핵심 해석과 확인 필요 매장은 작성 시점 값이며, 지표·표·차트는 원데이터와 수식으로 연결돼 있습니다.'
    dash.write(nr, 1, n1, K.fmt(font_size=9, font_color=P['sub']))
    if mbb:
        ds = ST['ytd'] - mbb['ytd']; dp = ST['py'] - mbb['ytd_py']
        dash.write(nr + 1, 1, f"※ 대조: 매장 행 합계 기준입니다. 마스터 DB 브랜드(MB) 합계는 누계 {mbb['ytd']:,.1f} / 전년 {mbb['ytd_py']:,.1f}백만원으로 "
                              f"매장 합계와 {ds:+,.1f} / {dp:+,.1f}백만원 차이가 있습니다(매장에 배분되지 않은 금액).",
                   K.fmt(font_size=9, font_color=P['sub']))
    K.footer(dash, nr + 3, 12, author, 1)

    # ==================== 사용설명 ====================
    K.title(wh, 1, 1, 2, '사용 설명', None)
    wh.set_row(3, 6)
    tbl = [('대시보드', '핵심 지표 6개, 월별 추이, 유통별 달성률, TOP 10, 확인 필요 매장, 핵심 해석'),
           ('유통별', '유통별 당월·누계 실적, 달성률, 구성비, 월별 추이 미니 차트'),
           ('월별추이', '26년·25년·신장율·월 목표 3행 블록과 유통×월 전년비 히트맵'),
           ('매장별', '운영 매장 전체 순위표. 막대·색조·미니 차트, 필터로 골라 보기'),
           ('분석', '할인율×판매율 분포와 확인 필요 매장'),
           ('원데이터', '모든 숫자의 출처. 회색 머리글은 입력값, 하늘색 머리글은 수식 계산값')]
    wh.write(4, 1, '시트', K.fmt(bold=True, bg_color=P['navy'], font_color=P['white'], align='center'))
    wh.write(4, 2, '내용', K.fmt(bold=True, bg_color=P['navy'], font_color=P['white']))
    for i, (s, t) in enumerate(tbl):
        wh.write_url(5 + i, 1, f"internal:'{s}'!A1", K.fmt(font_color=P['blue'], underline=1, align='center', border=1, border_color=P['line']), s)
        wh.write(5 + i, 2, t, K.fmt(border=1, border_color=P['line'], indent=1)); wh.set_row(5 + i, 22)
    K.section(wh, 12, 1, 2, '표기 규칙', '')
    rules = ['마이너스(감소)는 빨간 ▲로 표시합니다. 예) ▲8.2%',
             '달성률 100% 이상은 파랑 굵게, 90% 미만은 빨강입니다.',
             '막대 모양(데이터 바)은 값의 크기, 파랑-빨강 색조는 전년 대비 증감입니다.',
             '노랑 머리글·셀은 현재 진행 중인 달입니다.',
             '금액 단위는 표마다 오른쪽 위 끝에 적었습니다(기본 백만원).']
    K.notes(wh, 13, 1, 2, rules, '•', 20)
    K.section(wh, 19, 1, 2, '데이터 갱신 방법', '')
    K.notes(wh, 20, 1, 2, ['원데이터 시트의 회색 머리글 열(월별 매출·전년·목표, TAG·실판매·재고)에 새 값을 덮어쓰면 모든 시트가 자동 계산됩니다.',
                          '매장 수가 바뀌면 원데이터 행을 추가한 뒤 수식 범위를 넓히거나, 이 파일을 다시 생성하세요.',
                          '핵심 해석·확인 필요 매장 문구는 재생성 시 갱신됩니다.'], '•', 36)
    dash.activate()
    K.close()
    return K


def load_ytd(path):
    return json.load(open(path, encoding='utf8'))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True); ap.add_argument('--out', required=True); ap.add_argument('--author', default=AUTHOR)
    a = ap.parse_args()
    D = load_ytd(a.input)
    tmp = a.out + '.tmp.xlsx'
    k1 = build(tmp, D, None, a.author)                       # 1차: 수식만
    vals = verify_xlsx.recalc_values(tmp, k1.fcells)         # LibreOffice 재계산 값 수집
    bad = {k: v for k, v in vals.items() if isinstance(v, str) and v.startswith('#')}
    if bad:
        print('수식 오류:', list(bad.items())[:10])
    clean = {k: ('' if v is None else v) for k, v in vals.items()}
    build(a.out, D, clean, a.author)                         # 2차: 값 주입
    os.remove(tmp)
    print('완료:', a.out, '| 수식', len(k1.fcells), '개 | 오류', len(bad))


if __name__ == '__main__':
    main()
