"""엑셀 기법 카탈로그: 영업 엑셀에 쓰는 시각화·서식·수식 기법을 한 파일에 모았습니다(각 시트 = 복사해 쓰는 레시피).
  python recipes_demo.py 기법카탈로그.xlsx
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlkit import Kit, P, NF, rc, cn

MONTHS = [f'{i}월' for i in range(1, 13)]
SALES = [2588, 2803, 5619, 5810, 5678, 4144, 3401, 3193, 4748, 6500, 7200, 6800]
LY = [2715, 2940, 5184, 5981, 5646, 4492, 3718, 3952, 5390, 7466, 7000, 7100]
TGT = [3000, 3300, 5900, 6800, 6500, 5100, 4200, 4500, 6400, 8500, 7700, 7400]
STORES = [('롯데잠실', '백화점', 795, .951, .055), ('현대판교', '백화점', 870, .782, -.178), ('신세계강남', '백화점', 640, 1.08, .12),
          ('김해', '아울렛', 934, .908, -.094), ('기흥', '아울렛', 914, .836, -.092), ('여주', '아울렛', 560, .70, -.046),
          ('대구', '대리점', 322, .72, -.13), ('전주', '대리점', 289, .66, -.21), ('온라인', '온라인', 8502, .699, -.239)]


def build(path):
    K = Kit(path)
    # ---------- 1. 서식 ----------
    ws = K.sheet('1.서식·조건부서식', tab=P['navy'], widths=[(0, 0, 1.5), (1, 1, 14), (2, 9, 13)])
    K.title(ws, 1, 1, 9, '조건부 서식 · 미니 차트 레시피', '데이터 막대 / 색조 / 아이콘 / 스파크라인 / 임계값')
    K.section(ws, 3, 1, 9, '1. 매장 성과표', '[단위:백만원, %]')
    hdr = ['매장', '유통', '누계 매출', '달성률', '전년비', '월별 추이', '아이콘(달성)', '막대(매출)', '색조(전년비)']
    rows = [[s[0], s[1], s[2], s[3], s[4], None, s[3], s[2], s[4]] for s in STORES]
    last = K.table(ws, 5, 1, hdr, rows, [None, None, NF['m0'], NF['pct'], NF['yoy'], None, NF['pct'], NF['m0'], NF['yoy']],
                   ['left', 'center', 'right', 'right', 'right', 'center', 'center', 'right', 'right'], row_h=22)
    a, b = 7, 6 + len(STORES)
    K.threshold(ws, f'E{a}:E{b}')                                   # 100%↑ 파랑 / 90%↓ 빨강 (글자색)
    K.databar(ws, f'I{a}:I{b}')                                     # 데이터 막대
    K.heat(ws, f'J{a}:J{b}', 'div')                                 # 증감 색조(빨강-흰색-파랑)
    ws.conditional_format(f'H{a}:H{b}', {'type': 'icon_set', 'icon_style': '3_traffic_lights', 'icons': [
        {'criteria': '>=', 'type': 'number', 'value': 1}, {'criteria': '>=', 'type': 'number', 'value': .9}]})
    for i in range(len(STORES)):                                    # 스파크라인: 행마다 12개월
        for j in range(12):
            ws.write(30 + i, 12 + j, SALES[j] * (0.6 + 0.08 * i), K.fmt(font_size=8, font_color=P['sub']))
        K.spark(ws, f'G{7 + i}', f"'1.서식·조건부서식'!{rc(30 + i, 12)}:{rc(30 + i, 23)}")
    ws.set_column(12, 23, 5, None, {'hidden': True})
    K.notes(ws, last + 2, 1, 9, ['달성률 칸: 글자색 규칙(회사 관례). 아이콘 칸: 신호등 아이콘(임팩트가 필요할 때만).',
                                '데이터 막대는 크기, 색조는 증감, 스파크라인은 추이. 한 표에 3가지 이상 겹치면 오히려 안 보입니다.'], '※')
    # ---------- 2. 차트 ----------
    SH = '2.차트 5종'
    wc = K.sheet(SH, tab=P['navy2'], widths=[(0, 0, 1.5), (1, 16, 10)])
    K.title(wc, 1, 1, 16, '차트 선택 가이드 · 5종 레시피', '추이=막대+선 / 비교·목표=가로 막대 / 구성=도넛 / 증감=폭포 / 집중=파레토')
    # 차트 데이터는 아래(60행~)에 모아 두고, 차트는 위쪽 격자에 배치
    D0 = 60
    K.section(wc, D0 - 2, 1, 16, '차트 데이터 (수정하면 위 차트가 바뀝니다)', '')
    hf = K.fmt(bold=True, font_size=9, bg_color=P['grey'], align='center'); cf = K.fmt(font_size=9, num_format='#,##0')
    for j, h in enumerate(['월'] + MONTHS):
        wc.write(D0, 1 + j, h, hf)
    for r, (nm, arr) in enumerate((('26년', SALES), ('25년', LY), ('목표', TGT))):
        wc.write(D0 + 1 + r, 1, nm, K.fmt(font_size=9, bold=True))
        for j, v in enumerate(arr):
            wc.write(D0 + 1 + r, 2 + j, v, cf)
    K.col_line_chart(wc, 'B4', SH, (D0, 2, D0, 13), [('26년', (D0 + 1, 2, D0 + 1, 13))], [('25년', (D0 + 2, 2, D0 + 2, 13))],
                     [('목표', (D0 + 3, 2, D0 + 3, 13))], w=560, h=300)
    # 가로 막대(달성률) + 도넛(구성비) 데이터
    B0 = D0 + 6
    for j, h in enumerate(['유통', '달성률', '구성비']):
        wc.write(B0, 1 + j, h, hf)
    for i, (c, v, sh) in enumerate((('백화점', .79, .24), ('아울렛', .89, .36), ('대리점', .72, .17), ('직영점', .69, .02), ('온라인', .70, .19))):
        wc.write(B0 + 1 + i, 1, c, K.fmt(font_size=9)); wc.write(B0 + 1 + i, 2, v, K.fmt(font_size=9, num_format='0%'))
        wc.write(B0 + 1 + i, 3, sh, K.fmt(font_size=9, num_format='0%'))
    K.bar_chart(wc, 'K4', SH, (B0 + 1, 1, B0 + 5, 1), (B0 + 1, 2, B0 + 5, 2), name='달성률', w=440, h=300, max_=1.2)
    K.doughnut_chart(wc, 'B21', SH, (B0 + 1, 1, B0 + 5, 1), (B0 + 1, 3, B0 + 5, 3), w=400, h=290)
    # 폭포(증감): 누적막대 + 투명 기준 막대
    W0 = B0 + 8
    steps = [('전년 누계', 49034), ('백화점', -450), ('아울렛', -72), ('대리점', -680), ('직영점', -150), ('온라인', -2700), ('올해 누계', 45488)]
    base, up, dn, tot, run = [], [], [], [], 0
    for i, (n, v) in enumerate(steps):
        if i in (0, len(steps) - 1):
            base.append(0); up.append(0); dn.append(0); tot.append(v); run = v
        else:
            run += v
            base.append(min(run, run - v)); up.append(max(v, 0)); dn.append(max(-v, 0)); tot.append(0)
    for j, h in enumerate(['단계', '기준', '증가', '감소', '합계']):
        wc.write(W0, 1 + j, h, hf)
    for i, (n, _) in enumerate(steps):
        for j, v in enumerate((n, base[i], up[i], dn[i], tot[i])):
            wc.write(W0 + 1 + i, 1 + j, v, K.fmt(font_size=9) if j == 0 else cf)
    wf = K.wb.add_chart({'type': 'column', 'subtype': 'stacked'})
    for col, color, nm in ((2, None, '기준'), (3, P['navy2'], '증가'), (4, P['red'], '감소'), (5, P['navy'], '합계')):
        wf.add_series({'name': nm, 'categories': [SH, W0 + 1, 1, W0 + 7, 1], 'values': [SH, W0 + 1, col, W0 + 7, col],
                       'fill': {'none': True} if color is None else {'color': color}, 'gap': 40,
                       'data_labels': {'value': color is not None, 'num_format': '#,##0;;', 'font': {'size': 8, 'name': '맑은 고딕'}}})
    K._style(wf, legend=None, w=620, h=290); K._axes(wf)
    wf.set_y_axis({'min': 40000, 'num_format': '#,##0', 'num_font': {'name': '맑은 고딕', 'size': 9, 'color': P['sub']},
                   'line': {'none': True}, 'major_gridlines': {'visible': True, 'line': {'color': '#E7E6E6'}}})
    wc.insert_chart('H21', wf)
    # 파레토(집중도): 내림차순 막대 + 누적% 선(보조축)
    P0 = W0 + 10
    pv = sorted([x[2] for x in STORES], reverse=True)
    for j, h in enumerate(['순위', '매출', '누적%']):
        wc.write(P0, 1 + j, h, hf)
    for i, v in enumerate(pv):
        wc.write(P0 + 1 + i, 1, f'{i + 1}위', K.fmt(font_size=9)); wc.write(P0 + 1 + i, 2, v, cf)
        K.fx(wc, P0 + 1 + i, 3, f'=SUM($C${P0 + 2}:C{P0 + 2 + i})/SUM($C${P0 + 2}:$C${P0 + 1 + len(pv)})', K.fmt(font_size=9, num_format='0%'))
    pc = K.wb.add_chart({'type': 'column'})
    pc.add_series({'name': '매출', 'categories': [SH, P0 + 1, 1, P0 + len(pv), 1], 'values': [SH, P0 + 1, 2, P0 + len(pv), 2],
                   'fill': {'color': P['navy']}, 'gap': 50})
    pl = K.wb.add_chart({'type': 'line'})
    pl.add_series({'name': '누적%', 'categories': [SH, P0 + 1, 1, P0 + len(pv), 1], 'values': [SH, P0 + 1, 3, P0 + len(pv), 3], 'y2_axis': True,
                   'line': {'color': P['gold'], 'width': 2.25}, 'marker': {'type': 'circle', 'size': 5, 'fill': {'color': P['gold']}}})
    pc.combine(pl); K._style(pc, w=560, h=290); K._axes(pc)
    pl.set_y2_axis({'num_format': '0%', 'max': 1, 'min': 0, 'num_font': {'name': '맑은 고딕', 'size': 9}})
    wc.insert_chart('B37', pc)
    K.notes(wc, 54, 1, 16, ['추이는 막대+선(전년·목표), 크기·목표 비교는 가로 막대(축 최대 120%), 구성은 도넛(5개 이하), 증감 분해는 폭포(누적막대+투명 기준, 축 최소값 조정),',
                          '집중도는 파레토(막대+누적% 선). 3D·원형 7조각 이상·무지개색은 쓰지 않습니다. 분포는 산점도(분석 시트 참고).'], '※', 24)
    # ---------- 3. 수식·선택형 ----------
    wf3 = K.sheet('3.수식·선택형', tab=P['steel'], widths=[(0, 0, 1.5), (1, 8, 14)])
    K.title(wf3, 1, 1, 8, '수식 패턴 · 드롭다운으로 바뀌는 대시보드', '선택 유통을 바꾸면 아래 숫자가 즉시 바뀝니다')
    K.section(wf3, 3, 1, 8, '1. 유통 선택', '')
    wf3.write(5, 1, '선택 ▶', K.fmt(bold=True, align='right'))
    wf3.write(5, 2, '백화점', K.fmt(bold=True, bg_color=P['yellow'], border=1, align='center'))
    wf3.data_validation(5, 2, 5, 2, {'validate': 'list', 'source': ['백화점', '아울렛', '대리점', '온라인'], 'input_title': '유통', 'input_message': '목록에서 고르세요'})
    base0 = 9
    K.table(wf3, base0, 1, ['매장', '유통', '누계 매출', '달성률', '전년비'], [[s[0], s[1], s[2], s[3], s[4]] for s in STORES],
            [None, None, NF['m0'], NF['pct'], NF['yoy']], ['left', 'center', 'right', 'right', 'right'], header='grey')
    e = base0 + 1 + len(STORES)
    rg = lambda col: f'${col}${base0 + 2}:${col}${e}'
    cards = [dict(label='선택 유통 매출', value=f'=SUMIFS({rg("D")},{rg("C")},$C$6)', nf='#,##0', accent=P['gold']),
             dict(label='매장 수', value=f'=COUNTIFS({rg("C")},$C$6)', nf='0', accent=P['navy']),
             dict(label='최고 매출 매장', value=f'=IFERROR(INDEX({rg("B")},MATCH(SUMPRODUCT(MAX(({rg("C")}=$C$6)*{rg("D")})),{rg("D")},0)),"-")',
                  nf='@', accent=P['steel'])]
    K.kpi_cards(wf3, base0 + len(STORES) + 3, 1, cards, 2)
    K.notes(wf3, base0 + len(STORES) + 7, 1, 8, ['SUMIFS/COUNTIFS로 피벗 없이 집계, INDEX+MATCH로 조회(XLOOKUP·FILTER는 호환 문제로 쓰지 않음).',
                                                 '드롭다운(데이터 유효성 검사)은 시트 하나로 여러 보고서를 대신합니다.'], '※')
    K.close()


if __name__ == '__main__':
    build(sys.argv[1] if len(sys.argv) > 1 else '기법카탈로그.xlsx')
    print('완료')
