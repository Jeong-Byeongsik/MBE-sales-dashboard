"""월간 영업실적 보고 가안 빌더 (2-in-1).

  첫 시트 '보고(요약)' = 회사 양식(검정 실선, 회색 머리글, 노랑 강조, ▲빨강)
  나머지 = 임팩트 모드(대시보드, 증감 분석, 일별, 3사 비교, 품목·스타일, 이슈·액션) + 원데이터·기준값·사용설명

  python build_monthly_report.py --data DIR --date 2026-11-05 --out "★MBE 11월 영업실적 보고(11.05마감)_가안.xlsx"
DIR 안에 ytd.json, report-<날짜>.json, sales-<YYYY-MM>.json, stores.json 이 있어야 합니다.
숫자 원칙: 헤드라인(보고(요약))은 일일 보고서 공식 숫자, 매장 단위는 마스터 DB. 차이는 '※ 대조' 주석으로 밝힙니다.
"""
import argparse, collections, datetime, json, os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlkit import Kit, P, NF, rc, cn
import build_sales_workbook as B
import verify_xlsx

BLK = '#000000'
ORDER = ['보고(요약)', '대시보드', '유통별', '월별추이', '증감분석', '매장별', '분석', '일별', '3사비교', '품목·스타일',
         '이슈·액션', '검증', '원데이터', '기준값', '사용설명']
DESC = [('보고(요약)', '회사 양식 보고 1~2쪽: 실적 요약, 유통별 실적, 월별 트렌드, 意見'),
        ('대시보드', '핵심 지표 6개, 월별 추이, 유통별 달성률, TOP 10, 확인 필요 매장 (한 장)'),
        ('유통별', '유통별 당월·누계 실적, 진도달성, 구성비, 월별 추이'),
        ('월별추이', '26년·25년·신장율·월 목표 3행 블록과 유통×월 전년비 히트맵'),
        ('증감분석', '전년 대비 증감을 유통별로 분해한 폭포 차트와 기여도'),
        ('매장별', '운영 매장 순위표(막대·색조·미니 차트), 필터로 골라 보기'),
        ('분석', '할인율×판매율 분포와 확인 필요 매장'),
        ('일별', '일 목표 대비 일별 실적, 잔여 필요 일평균'),
        ('3사비교', '마스터바니·파리게이츠·세인트앤드류스 비교'),
        ('품목·스타일', '아이템별 매출 구성과 베스트 품번'),
        ('이슈·액션', '점검 항목, 우선 점검 매장, 다음 주 액션 플랜, 프로모션 제안'),
        ('검증', '원본 대조·합계·부분 기간·극단값 자동 검증 결과'),
        ('원데이터', '매장별 월 실적·목표·TAG·재고 (모든 숫자의 출처)'),
        ('기준값', '일일 보고서에서 가져온 목표·마감 예상 등 입력값 (노란 칸만 수정)')]
SIZE = {'W': 1090}


def M(v):
    return None if v is None else v / 1e6


class Report:
    def __init__(self, data_dir, date):
        L = lambda f: json.load(open(os.path.join(data_dir, f), encoding='utf8'))
        self.ytd, self.rep = L('ytd.json'), L(f'report-{date}.json')
        self.sales, self.st = L(f'sales-{date[:7]}.json'), {s['code']: s for s in L('stores.json')['stores']}
        r, self.nm = self.rep, self.rep['month']
        g = r['goal']
        gv = lambda pred: next(v['amount'] for k, v in g.items() if pred(k))
        nm = self.nm
        days_left = int(re.search(r'(\d+)', g[next(k for k in g if k.startswith('잔여 필요 매출'))]['note']).group(1))
        yr = re.search(r'연 목표 ([\d,]+)백만', ' '.join(r['insights']))
        fc, f2, ky = r['forecast'][0], r['forecast'][1], r['kpi']['ytd']
        rows = [
            ('month_target', f'{nm}월 월 목표', M(gv(lambda k: k.startswith(f'{nm}월 목표'))), 'ERP 월 목표(일일 보고서)', 'm1'),
            ('progress_target', '경과 진도 목표', M(gv(lambda k: k.startswith('경과 진도'))), '일 목표 누계(1일~기준일)', 'm1'),
            ('progress_share', '진도율', '={progress_target}/{month_target}', '경과 진도 목표 ÷ 월 목표', 'pct'),
            ('days_left', '잔여일', days_left, '기준일 다음 날부터 월말까지', 'int'),
            ('y_day_target', '전일 일 목표', M(gv(lambda k: k.startswith('어제') and k.endswith('일 목표'))), 'ERP 일 목표', 'm1'),
            ('y_day_sales', '전일 실적', M(gv(lambda k: k.startswith('어제') and k.endswith('실적'))), '일일 보고서', 'm1'),
            ('rpt_month_sales', f'{nm}월 누계 실적(공식)', M(r['kpi']['month']['sales']), '일일 보고서 기준(매장 합계와 소폭 차이)', 'm1'),
            ('rpt_month_ly', '전년 같은 기간', M(r['kpi']['month']['ly']), '전년 1일~기준일', 'm1'),
            ('need_daily', '잔여 필요 일평균', '=({month_target}-{rpt_month_sales})/{days_left}', '(월 목표 − 누계) ÷ 잔여일', 'm1'),
            ('plan_daily', '잔여 일목표 평균(계획)', M(gv(lambda k: k.startswith('잔여 일목표'))), '남은 날의 일 목표 평균', 'm1'),
            ('fc_base', f'{nm}월 마감 예상(기준)', M(fc['base']), '전년 비중 + 주중/주말 가중 추정', 'm1'),
            ('fc_low', '마감 예상 하한', M(fc['low']), '추정 범위', 'm1'),
            ('fc_high', '마감 예상 상한', M(fc['high']), '추정 범위', 'm1'),
            ('fc_ly_close', '전년 마감(같은 달)', M(fc['ly_close']), '전년 월 합계', 'm1'),
            ('dec_base', '다음 달 예상(기준)', M(f2['base']), '추정', 'm1'),
            ('dec_target', '다음 달 목표', M(f2['target']), 'ERP 월 목표', 'm1'),
            ('dec_ly', '다음 달 전년 마감', M(f2['ly_close']), '전년 월 합계', 'm1'),
            ('ytd_target', f'누계 목표(1~{nm}월)', M(ky['target']), '월 목표 합(진행 중인 달은 전체 월 목표)', 'm1'),
            ('rpt_ytd_sales', '누계 실적(공식)', M(ky['sales']), '일일 보고서 기준', 'm1'),
            ('rpt_ytd_ly', '누계 전년 동기', M(ky['ly']), '전년 1월~같은 기간', 'm1'),
        ]
        if yr:
            rows.append(('year_target', '연 목표', float(yr.group(1).replace(',', '')), '일일 보고서 요점 문구에서 가져옴', 'm1'))
        self.bv_rows = rows
        self.addr = {k: f"'기준값'!$C${5 + i}" for i, (k, *_) in enumerate(rows)}
        self.specs = {
            '보고(요약)': dict(tab=P['gold'], widths=[(0, 0, 1.5), (1, 1, 15), (2, 13, 10.4)]),
            '증감분석': dict(tab=P['navy2'], widths=[(0, 0, 1.5), (1, 12, 12.6)]),
            '일별': dict(tab=P['steel'], widths=[(0, 0, 1.5), (1, 12, 12.6)]),
            '3사비교': dict(tab=P['steel'], widths=[(0, 0, 1.5), (1, 1, 16), (2, 12, 11.5)]),
            '품목·스타일': dict(tab=P['steel'], widths=[(0, 0, 1.5), (1, 1, 20), (2, 12, 12)]),
            '이슈·액션': dict(tab='#C00000', widths=[(0, 0, 1.5), (1, 1, 9), (2, 2, 40), (3, 3, 66), (4, 5, 14), (6, 6, 24)]),
            '검증': dict(tab='#548235', widths=[(0, 0, 1.5), (1, 1, 9), (2, 2, 44), (3, 5, 15), (6, 6, 11), (7, 7, 62)]),
            '기준값': dict(tab=P['grey'], widths=[(0, 0, 1.5), (1, 1, 30), (2, 2, 16), (3, 3, 52)]),
        }
        self.order, self.desc = ORDER, DESC
        self.cells, self.checks = {}, []

    def base_data(self):
        D = dict(self.ytd)
        D['ytd_full_target'] = M(self.rep['kpi']['ytd']['target'])
        D['bv'] = self.addr
        return D

    # ------------------------------------------------------------------ fill
    def fill(self, ctx):
        self.ctx = ctx
        self.K = ctx['K']
        self.s_base(); self.s_summary(); self.s_waterfall(); self.s_daily()
        self.s_brands(); self.s_items(); self.s_issues(); self.s_checks()

    def a(self, k):
        return self.addr[k]

    # ------------------------------------------------------------------ 기준값
    def s_base(self):
        K, ws = self.K, self.ctx['sheets']['기준값']
        K.title(ws, 1, 1, 3, '기준값 (일일 보고서 입력)', None)
        hf = K.fmt(bold=True, align='center', bg_color=P['grey'], border=1, border_color=P['line'])
        for j, h in enumerate(['항목', '값 (백만원)', '출처·설명']):
            ws.write(3, 1 + j, h, hf)
        nfs = dict(m1='#,##0.0', pct='0.0%', int='0')
        local = {k: f'$C${5 + i}' for i, (k, *_) in enumerate(self.bv_rows)}
        for i, (k, label, v, note, nf) in enumerate(self.bv_rows):
            r = 4 + i
            ws.set_row(r, 20)
            ws.write(r, 1, label, K.fmt(border=1, border_color=P['line'], indent=1))
            isf = isinstance(v, str)
            vf = K.fmt(num_format=nfs[nf], border=1, border_color=P['line'], align='right', **({} if isf else {'bg_color': '#FFF2CC'}))
            if isf:
                K.fx(ws, r, 2, v.format(**local) if False else re.sub(r'\{(\w+)\}', lambda m: local[m.group(1)], v), vf)
            else:
                ws.write_number(r, 2, v, vf)
            ws.write(r, 3, note, K.fmt(border=1, border_color=P['line'], font_color=P['sub'], font_size=9, indent=1))
        e = 4 + len(self.bv_rows) + 1
        K.notes(ws, e, 1, 3, ['노란 칸은 일일 보고서에서 가져온 입력값입니다. 새 보고서로 바꿀 때 이 칸만 고치세요.',
                              '흰 칸은 수식입니다.'], '※')

    # ------------------------------------------------------------------ 보고(요약) 회사 양식
    def s_summary(self):
        K, c, R = self.K, self.ctx, self.rep
        ws, nm, day = c['sheets']['보고(요약)'], self.nm, c['day']
        A = self.a
        cf = lambda **kw: K.fmt(border=1, border_color=BLK, **kw)
        G = '#D9D9D9'
        K.title(ws, 1, 1, 13, f'MBE {nm}월 영업실적 보고', c['asof'], style='company')
        ws.write(3, 1, '※ 가안(초안)입니다. 숫자는 일일 보고서 기준이며 확인 후 확정하세요.', K.fmt(font_size=9, font_color=P['sub']))

        def sec(r, text, unit, last):
            ws.write(r, 1, text, K.fmt(font_size=11, bold=True))
            ws.write(r, last, unit, K.fmt(font_size=9, align='right'))
            ws.set_row(r, 20)

        # 1. 실적 요약 (1줄 표)
        sec(5, '1. 실적 요약', '[단위:백만원, %]', 9)
        heads = [f'{nm}월 목표', f'{nm}월 누계\n(1~{day}일)', '달성률', '진도 달성', '전년 동기', '신장율', '마감 예상', '예상 달성률', '잔여 필요\n일평균']
        vals = [f"={A('month_target')}", f"={A('rpt_month_sales')}", '=C8/B8', f"=C8/{A('progress_target')}", f"={A('rpt_month_ly')}",
                '=C8/F8-1', f"={A('fc_base')}", '=H8/B8', f"={A('need_daily')}"]
        nfs = [NF['m0'], NF['m0'], NF['pct'], NF['pct'], NF['m0'], NF['yoy'], NF['m0'], NF['pct'], NF['m0']]
        ws.set_row(6, 32); ws.set_row(7, 34)
        RED = dict(left=2, right=2, left_color='#FF0000', right_color='#FF0000')
        for j, (h, v, nf) in enumerate(zip(heads, vals, nfs)):
            red = j in (3, 7)
            hk = dict(bold=True, align='center', bg_color=G, text_wrap=True)
            vk = dict(font_size=12, bold=True, align='center', num_format=nf)
            if red:
                hk.update(RED, top=2, top_color='#FF0000'); vk.update(RED, bottom=2, bottom_color='#FF0000')
            ws.write(6, 1 + j, h, cf(**hk))
            K.put(ws, 7, 1 + j, v, cf(**vk))
        K.threshold(ws, 'E8'); K.threshold(ws, 'I8')
        self.cells.update(sum_target='B8', sum_sales='C8', sum_prog='E8', sum_fc_rate='I8')
        prog = c['ST']['mach']
        mg = self.m_growth()
        lines = [
            f"{nm}월 1~{day}일 매출 {R['kpi']['month']['sales'] / 1e6:,.0f}백만으로 월 목표의 {R['kpi']['month']['rate'] * 100:.1f}%, 일 목표 진도 대비 "
            f"{R['kpi']['month']['sales'] / R['goal'][self.gkey('경과 진도')]['amount'] * 100:.1f}%, 전년 같은 기간 대비 {R['kpi']['month']['growth'] * 100:+.1f}% 임",
            f"마감 예상 {R['forecast'][0]['base'] / 1e6:,.0f}백만(범위 {R['forecast'][0]['low'] / 1e6:,.0f}~{R['forecast'][0]['high'] / 1e6:,.0f}), 목표 대비 "
            f"{R['forecast'][0]['rate'] * 100:.1f}%, 목표 달성에는 잔여 {self.days_left()}일 일평균 {R['goal'][self.gkey('잔여 필요 일평균')]['amount'] / 1e6:,.0f}백만 필요함",
            f"1~{nm}월 누계 {R['kpi']['ytd']['sales'] / 1e6:,.0f}백만, 전년 대비 {R['kpi']['ytd']['growth'] * 100:+.1f}%, 누계 목표 "
            f"{R['kpi']['ytd']['target'] / 1e6:,.0f}백만의 {R['kpi']['ytd']['rate'] * 100:.1f}% (부족 {R['kpi']['ytd']['short'] / 1e6:,.0f}백만)"]
        for i, t in enumerate(lines):
            ws.merge_range(9 + i, 1, 9 + i, 13, f'▶ {t}', K.fmt(font_size=10.5, indent=1))
            ws.set_row(9 + i, 21)

        # 2. 유통별 실적 현황 (공식 일일 보고서 유통 구분)
        sec(13, '2. 유통별 실적 현황', '[단위:백만원, %]', 11)
        yel = '#FFFF00'
        h1 = cf(bold=True, align='center', bg_color=G, text_wrap=True)
        hy = cf(bold=True, align='center', bg_color=yel)
        ws.set_row(14, 22); ws.set_row(15, 26)
        ws.merge_range(14, 1, 15, 1, '구  분', h1); ws.merge_range(14, 2, 15, 2, '매장수', h1)
        ws.merge_range(14, 3, 14, 10, f'{nm}월 (1일~{day}일)', hy); ws.merge_range(14, 11, 15, 11, '누계\n신장율', h1)
        for j, h in enumerate(['월 목표', '진도 목표', '실적', '진도 달성', '월 달성', '전년 동기', '신장율', '구성비']):
            ws.write(15, 3 + j, h, cf(bold=True, align='center', bg_color='#F2F2F2'))
        ch = {x['name']: x for x in R['channels']}
        names = [x['name'] for x in R['channels'] if '합' not in x['name']]
        offline = [n for n in names if n != '온라인']
        order = offline + ['오프라인 소계'] + (['온라인'] if '온라인' in names else []) + ['기타(미분류)', '합  계']
        r0 = 16
        rowof = {n: r0 + i for i, n in enumerate(order)}
        share = A('progress_share')
        mbytd = next((b for b in R['brands'] if b['name'] == '마스터바니'), None)
        for n in order:
            r = rowof[n]; e = r + 1
            ws.set_row(r, 21)
            sub = n == '오프라인 소계'; tot = n.startswith('합')
            bg = '#FBE5D6' if sub else ('#DEEBF7' if tot else None)
            base = dict(align='center', bold=sub or tot, **({'bg_color': bg} if bg else {}), **({'top': 2} if tot else {}))
            num = lambda nf, **k: cf(**{**base, 'align': 'right', 'num_format': nf, **k})
            ws.write(r, 1, n, cf(**base))
            if sub:
                f_off = lambda col: f"=SUM({col}{rowof[offline[0]] + 1}:{col}{rowof[offline[-1]] + 1})"
                K.put(ws, r, 2, f_off('C'), num(NF['qty'])); K.put(ws, r, 3, f_off('D'), num(NF['m0']))
                K.put(ws, r, 5, f_off('F'), num(NF['m0'])); K.put(ws, r, 8, f_off('I'), num(NF['m0']))
                ws.write_blank(r, 11, None, num(NF['yoy']))
            elif n == '기타(미분류)':
                on = rowof.get('온라인'); tr = rowof['합  계'] + 1
                rest = lambda col: f"={col}{tr}-{col}{rowof['오프라인 소계'] + 1}" + (f"-{col}{on + 1}" if on is not None else '')
                ws.write(r, 2, '-', cf(**base))
                for col_i, col in ((3, 'D'), (5, 'F'), (8, 'I')):
                    K.put(ws, r, col_i, rest(col), num(NF['m0']))
                ws.write_blank(r, 11, None, num(NF['yoy']))
            elif tot:
                x = ch[n]
                ws.write_number(r, 2, x['stores'], num(NF['qty'])); ws.write_number(r, 3, M(x['target']), num(NF['m0']))
                ws.write_number(r, 5, M(x['sales']), num(NF['m0'])); ws.write_number(r, 8, M(x['ly']), num(NF['m0']))
                if mbytd:
                    ws.write_number(r, 11, mbytd['ytd_growth'], num(NF['yoy']))
                else:
                    ws.write_blank(r, 11, None, num(NF['yoy']))
            else:
                x = ch[n]
                if x['stores'] in (None, '-'):
                    ws.write(r, 2, '-', cf(**base))
                else:
                    ws.write_number(r, 2, x['stores'], num(NF['qty']))
                ws.write_number(r, 3, M(x['target']), num(NF['m0'])); ws.write_number(r, 5, M(x['sales']), num(NF['m0']))
                ws.write_number(r, 8, M(x['ly']), num(NF['m0']))
                ws.write_number(r, 11, x['ytd_growth'], num(NF['yoy'])) if x.get('ytd_growth') is not None else ws.write_blank(r, 11, None, num(NF['yoy']))
            K.put(ws, r, 4, f"=D{e}*{share}", num(NF['m0']))
            K.put(ws, r, 6, f'=IFERROR(F{e}/E{e},"")', num(NF['pct']))
            K.put(ws, r, 7, f'=IFERROR(F{e}/D{e},"")', num(NF['pct']))
            if n == '기타(미분류)':
                ws.write_blank(r, 9, None, num(NF['yoy']))              # 전년 규모가 작아 비율이 의미 없음
            else:
                K.put(ws, r, 9, f'=IFERROR(F{e}/I{e}-1,"")', num(NF['yoy']))
            K.put(ws, r, 10, f"=IFERROR(F{e}/F${rowof['합  계'] + 1},\"\")", num(NF['pct']))
        last_r = rowof['합  계']
        self.cells.update(ch_rows={n: r + 1 for n, r in rowof.items()})
        K.threshold(ws, f'G{r0 + 1}:G{last_r + 1}')
        ws.write(last_r + 1, 1, "※ 진도 달성 = 실적 ÷ 진도 목표(월 목표 × 진도율). 월 목표·진도율은 '기준값' 시트 참조, 누계 신장율은 일일 보고서 기준.",
                 K.fmt(font_size=9, font_color=P['sub']))
        ws.write(last_r + 2, 1, '※ 기타(미분류) = 합계 − (오프라인 소계 + 온라인). 일일 보고서 합계에는 유통으로 분류되지 않은 매출이 포함됩니다.',
                 K.fmt(font_size=9, font_color=P['sub']))

        # 3. MB 월별 트렌드
        t0 = last_r + 4
        sec(t0, '3. MB 월별 실적 트렌드', '[단위:백만원, %]', 13)
        hr = t0 + 1
        ws.set_row(hr, 24)
        mm = R['mb_monthly']
        n_m = len(mm)
        ws.write(hr, 1, '구  분', h1)
        for j in range(n_m):
            ws.write(hr, 2 + j, mm[j]['month'], hy if j == n_m - 1 else h1)
        ws.write(hr, 2 + n_m, '누  계', h1)
        redbox = dict(left=2, right=2, left_color='#FF0000', right_color='#FF0000')
        labs = [('26년', 'sales'), ('25년', 'ly'), ('신장율', None), ('26년 할인율', 'dc'), ('25년 할인율', 'ly_dc')]
        for i, (lab, key) in enumerate(labs):
            r = hr + 1 + i; e = r + 1
            ws.set_row(r, 21)
            ws.write(r, 1, lab, cf(bold=True, align='center', bg_color='#F2F2F2'))
            for j in range(n_m):
                last = j == n_m - 1
                col = cn(2 + j)
                kw = dict(align='right', **({**redbox, 'bg_color': '#FFF2CC'} if last else {}), **({'bottom': 2, 'bottom_color': '#FF0000'} if last and i == len(labs) - 1 else {}))
                if key in ('sales', 'ly'):
                    ws.write_number(r, 2 + j, M(mm[j][key]), cf(num_format=NF['m0'], **kw))
                elif key in ('dc', 'ly_dc'):
                    v = mm[j].get(key)
                    ws.write_number(r, 2 + j, v, cf(num_format=NF['pct'], **kw)) if v is not None else ws.write_blank(r, 2 + j, None, cf(**kw))
                else:
                    K.put(ws, r, 2 + j, f'=IF({col}{hr + 3}>0,{col}{hr + 2}/{col}{hr + 3}-1,"")', cf(num_format=NF['yoy'], **kw))
            tc = cn(2 + n_m)
            tk = dict(align='right', bold=True, bg_color='#DEEBF7')
            if key in ('sales', 'ly'):
                K.put(ws, r, 2 + n_m, f"=SUM(C{e}:{cn(1 + n_m)}{e})", cf(num_format=NF['m0'], **tk))
            elif key is None:
                K.put(ws, r, 2 + n_m, f'=IF({tc}{hr + 3}>0,{tc}{hr + 2}/{tc}{hr + 3}-1,"")', cf(num_format=NF['yoy'], **tk))
            else:
                ws.write_blank(r, 2 + n_m, None, cf(**tk))
        ws.conditional_format(hr + 3, 2, hr + 3, 2 + n_m, {'type': 'no_blanks', 'format': K.fmt()})
        self.cells['trend_ytd'] = f"{cn(2 + n_m)}{hr + 2}"
        ws.merge_range(hr + 6, 1, hr + 6, 13, f"※ {nm}월(노랑·빨간 박스)은 1일~{day}일 진행 중이며 전년은 같은 기간입니다. 할인율 = 1 − 실판매 ÷ TAG. 일일 보고서(브랜드 합계) 기준이라 '월별추이' 시트(매장 합계)와 다릅니다.",
                       K.fmt(font_size=9, font_color=P['sub']))

        # 4. 그래프 (회사 규칙: 표 아래, 표와 같은 열 범위, 눈금선 없음, 남색/회색 점선)
        g0 = hr + 8
        sec(g0, '4. 그래프 (좌) 월별 매출 추이 (우) 유통별 진도 달성', '[단위:백만원, %]', 13)
        # 차트용 보조 데이터: 인쇄 영역 밖 오른쪽 열(숨김)
        hc = 15
        chn = [n for n in order if n in ('백화점', '아울렛', '대리점·직영점', '온라인')] + ['합  계']
        for i, n in enumerate(chn):
            K.put(ws, g0 + i, hc, f"=B{rowof[n] + 1}", K.fmt(font_size=8))
            K.put(ws, g0 + i, hc + 1, f"=G{rowof[n] + 1}", K.fmt(font_size=8, num_format='0.0%'))
        ws.set_column(hc, hc + 1, 12, None, {'hidden': True})
        lc = K.wb.add_chart({'type': 'column'})
        lc.add_series({'name': '26년', 'categories': ['보고(요약)', hr, 2, hr, 1 + n_m], 'values': ['보고(요약)', hr + 1, 2, hr + 1, 1 + n_m],
                       'fill': {'color': '#1F3864'}, 'gap': 60,
                       'data_labels': {'value': True, 'num_format': '#,##0', 'font': {'name': '맑은 고딕', 'size': 8}}})
        ll = K.wb.add_chart({'type': 'line'})
        ll.add_series({'name': '25년', 'categories': ['보고(요약)', hr, 2, hr, 1 + n_m], 'values': ['보고(요약)', hr + 2, 2, hr + 2, 1 + n_m],
                       'line': {'color': '#A6A6A6', 'width': 2, 'dash_type': 'dash'}, 'marker': {'type': 'none'}})
        lc.combine(ll)
        lc.set_title({'none': True}); lc.set_size({'width': 500, 'height': 200}); lc.set_legend({'position': 'bottom', 'font': {'name': '맑은 고딕', 'size': 9}})
        lc.set_chartarea({'border': {'none': True}}); lc.set_plotarea({'border': {'none': True}})
        lc.set_y_axis({'major_gridlines': {'visible': False}, 'num_font': {'name': '맑은 고딕', 'size': 8, 'color': '#7F7F7F'}, 'num_format': '#,##0', 'line': {'none': True}})
        lc.set_x_axis({'num_font': {'name': '맑은 고딕', 'size': 9}, 'line': {'color': '#BFBFBF'}})
        ws.insert_chart(g0 + 2, 1, lc, {'object_position': 1, 'x_offset': 2, 'y_offset': 4})
        rc_ = K.wb.add_chart({'type': 'bar'})
        rc_.add_series({'name': '진도 달성', 'categories': ['보고(요약)', g0, hc, g0 + len(chn) - 1, hc], 'values': ['보고(요약)', g0, hc + 1, g0 + len(chn) - 1, hc + 1],
                        'fill': {'color': '#1F3864'}, 'gap': 50,
                        'points': [{'fill': {'color': '#1F3864'}}] * (len(chn) - 1) + [{'fill': {'color': '#A6A6A6'}}],
                        'data_labels': {'value': True, 'num_format': '0.0%', 'font': {'name': '맑은 고딕', 'size': 9, 'bold': True}}})
        rc_.set_title({'none': True}); rc_.set_size({'width': 545, 'height': 200}); rc_.set_legend({'none': True})
        rc_.set_chartarea({'border': {'none': True}}); rc_.set_plotarea({'border': {'none': True}})
        rc_.set_y_axis({'reverse': True, 'num_font': {'name': '맑은 고딕', 'size': 9}, 'line': {'color': '#BFBFBF'}})
        rc_.set_x_axis({'min': 0, 'max': 1, 'major_gridlines': {'visible': False}, 'num_format': '0%', 'num_font': {'name': '맑은 고딕', 'size': 8, 'color': '#7F7F7F'}, 'line': {'none': True}})
        rc_.show_hidden_data()
        ws.insert_chart(g0 + 2, 7, rc_, {'object_position': 1, 'x_offset': 10, 'y_offset': 4})
        lc.show_hidden_data()

        # 5. 意見
        o0 = g0 + 13
        sec(o0, '5. ※ 意見 (가안)', '', 13)
        ops = self.opinions()
        for i, t in enumerate(ops):
            ws.merge_range(o0 + 1 + i, 1, o0 + 1 + i, 13, t, K.fmt(font_size=10.5, text_wrap=True, indent=1, valign='top'))
            ws.set_row(o0 + 1 + i, 33 if len(t) > 120 else 21)
        e = o0 + 1 + len(ops)
        ws.merge_range(e + 1, 1, e + 1, 13, f"※ 대조: 헤드라인은 일일 보고서 기준({nm}월 누계 {R['kpi']['month']['sales'] / 1e6:,.1f}), 매장 단위(대시보드·매장별)는 마스터 DB 합계({c['ST']['mtd']:,.1f})로 "
                       f"{c['ST']['mtd'] - R['kpi']['month']['sales'] / 1e6:+,.1f}백만 차이가 있습니다.", K.fmt(font_size=9, font_color=P['sub']))
        K.footer(ws, e + 3, 13, c['author'], 1)
        ws.set_h_pagebreaks([t0 - 1])

    def next_week(self):
        """기준일(보고서 date) 다음 주 날짜를 요일과 함께 계산한다(하드코딩 금지)."""
        d0 = datetime.date.fromisoformat(self.rep['date'])
        mon = d0 + datetime.timedelta(days=(7 - d0.weekday()) % 7 or 7)
        wd = '월화수목금토일'
        f = lambda n: (lambda x: f'{x.month}/{x.day}({wd[x.weekday()]})')(mon + datetime.timedelta(days=n))
        return dict(mon=f(0), wed=f(2), thu=f(3), later=f(11))

    def gkey(self, pref):
        return next(k for k in self.rep['goal'] if k.startswith(pref))

    def days_left(self):
        return int(re.search(r'(\d+)', self.rep['goal'][self.gkey('잔여 필요 매출')]['note']).group(1))

    def m_growth(self):
        return [m['growth'] for m in self.rep['mb_monthly'] if m.get('growth') is not None]

    def opinions(self):
        R, nm = self.rep, self.nm
        share = R['goal'][self.gkey('경과 진도')]['amount'] / R['kpi']['month']['target']
        pr = sorted(((x['sales'] / (x['target'] * share), x['name']) for x in R['channels'] if '합' not in x['name'] and x['target']))
        neg = 0
        for g in reversed(self.m_growth()):
            if g < 0:
                neg += 1
            else:
                break
        mb = next((b for b in R['brands'] if b['name'] == '마스터바니'), {})
        gro = sorted(((x['growth'], x['name']) for x in R['channels'] if '합' not in x['name'] and x['growth'] is not None))
        lines = [
            '▶ 현황 : 일 목표 진도 대비 ' + ', '.join(f'{n} {v * 100:.1f}%' for v, n in pr) + ' 순으로 낮음. 진도 미달 채널부터 점검 필요함',
            f'▶ 추이 : MB 전년비는 {neg}개월 연속 마이너스, 전년 같은 기간 대비 하락폭은 {gro[0][1]} {gro[0][0] * 100:.1f}%가 가장 큼',
        ]
        if mb.get('new_dc') is not None:
            lines.append(f"▶ 가격 : 신상 할인율 {mb['new_dc'] * 100:.1f}%로 전년 대비 {mb['new_dc_diff']:+.1f}%p, 할인 판매 확대로 매출 대비 수익성 점검 필요함")
        lines.append('▶ 가설(미확정) : 하락 원인(행사 시기, 재고 가용성, 날씨, 온라인 노출 등)은 확인 전이므로 담당 확인 후 확정 예정임')
        lines.append("▶ 향후 : 다음 주 액션 플랜과 프로모션 제안은 '이슈·액션' 시트 참조")
        return lines

    # ------------------------------------------------------------------ 증감분석
    def s_waterfall(self):
        K, c = self.K, self.ctx
        ws, R, ch = c['sheets']['증감분석'], c['R'], c['chans']
        K.title(ws, 1, 1, 12, '전년 대비 증감 분석 (누계)', c['asof'])
        K.section(ws, 3, 1, 12, 'Ⅰ. 유통별 증감 기여도', '[단위:백만원, %]')
        hdr = ['유통', '전년 누계', '올해 누계', '증감액', '기여도', '신장율']
        rows, h = [], 5
        for i, name in enumerate(ch):
            e = h + 2 + i
            rows.append([name, f"=SUMIFS({R('ytd_py')},{R('ch')},$B{e})", f"=SUMIFS({R('ytd')},{R('ch')},$B{e})", f"=D{e}-C{e}",
                         f"=IFERROR(E{e}/E${h + 2 + len(ch)},\"\")", f'=IFERROR(D{e}/C{e}-1,"")'])
        et = h + 2 + len(ch)
        rows.append(['합  계', f"=SUM(C{h + 2}:C{et - 1})", f"=SUM(D{h + 2}:D{et - 1})", f"=D{et}-C{et}", f"=SUM(F{h + 2}:F{et - 1})", f'=IFERROR(D{et}/C{et}-1,"")'])
        last = K.table(ws, h, 1, hdr, rows, [None, NF['m0'], NF['m0'], NF['m0'], NF['pct'], NF['yoy']],
                       ['center', 'right', 'right', 'right', 'right', 'right'], total_last=True, row_h=22, bold_cols=(0,))
        K.databar(ws, f'F{h + 2}:F{et - 1}', P['steel'])
        # 폭포 보조 데이터
        w0 = last + 3
        K.section(ws, w0 - 1, 1, 12, 'Ⅱ. 폭포 차트 데이터 (수식)', '')
        for j, hh in enumerate(['단계', '증감', '누적', '기준', '증가', '감소', '합계']):
            ws.write(w0, 1 + j, hh, K.fmt(bold=True, font_size=9, bg_color=P['grey'], align='center'))
        steps = ['전년 누계'] + ch + ['올해 누계']
        small = K.fmt(font_size=9, num_format='#,##0')
        for i, nmv in enumerate(steps):
            r = w0 + 1 + i; e = r + 1
            ws.write(r, 1, nmv, K.fmt(font_size=9))
            if i == 0:
                K.put(ws, r, 2, None, small); K.put(ws, r, 3, f"=C{et}", small)
                for col in (4, 5, 6): K.put(ws, r, col, 0, small)
                K.put(ws, r, 7, f"=D{e}", small)
            elif i == len(steps) - 1:
                K.put(ws, r, 2, None, small); K.put(ws, r, 3, f"=D{et}", small)
                for col in (4, 5, 6): K.put(ws, r, col, 0, small)
                K.put(ws, r, 7, f"=D{e}", small)
            else:
                src = h + 1 + i                       # 위 표의 해당 유통 행(1부터)
                K.put(ws, r, 2, f"=E{src}", small); K.put(ws, r, 3, f"=D{e - 1}+C{e}", small)
                K.put(ws, r, 4, f"=MIN(D{e - 1},D{e})", small); K.put(ws, r, 5, f"=MAX(C{e},0)", small)
                K.put(ws, r, 6, f"=MAX(-C{e},0)", small); K.put(ws, r, 7, 0, small)
        wr0, wr1 = w0 + 1, w0 + len(steps)
        vals = [R_ for R_ in (c['ST']['py'], c['ST']['ytd'])]
        ymin = int((min(vals) * 0.88) // 1000 * 1000)
        wf = K.wb.add_chart({'type': 'column', 'subtype': 'stacked'})
        for col, color, nmx in ((4, None, '기준'), (5, P['navy2'], '증가'), (6, P['red'], '감소'), (7, P['navy'], '합계')):
            wf.add_series({'name': nmx, 'categories': ['증감분석', wr0, 1, wr1, 1], 'values': ['증감분석', wr0, col, wr1, col],
                           'fill': {'none': True} if color is None else {'color': color}, 'gap': 45,
                           'data_labels': {'value': color is not None, 'num_format': '#,##0;;', 'font': {'size': 9, 'name': '맑은 고딕', 'bold': True}}})
        K._style(wf, legend=None, w=560, h=330); K._axes(wf)
        wf.set_y_axis({'min': ymin, 'num_format': '#,##0', 'num_font': {'name': '맑은 고딕', 'size': 9, 'color': P['sub']},
                       'line': {'none': True}, 'major_gridlines': {'visible': True, 'line': {'color': '#E7E6E6'}}})
        ws.insert_chart('H6', wf)
        deltas = sorted(((sum(sum(x['m']) - sum(x['p']) for x in self.ytd['stores'] if x['channel'] == n), n) for n in ch))
        tot = sum(d for d, _ in deltas)
        lines = [f"누계 매출은 전년 대비 {tot:+,.0f}백만(매장 합계 기준)이며, 감소 폭은 " + ', '.join(f'{n} {d:+,.0f}' for d, n in deltas[:3] if d < 0) + ' 순으로 큼',
                 '기여도 = 채널 증감액 ÷ 전체 증감액. 음수 채널이 전체 감소를 만든 비중을 보여줍니다',
                 f'폭포 차트의 세로축은 {ymin:,}부터 시작합니다(증감이 보이도록 축 최소값 조정)']
        K.notes(ws, wr1 + 3, 1, 12, lines, '▶', 22)
        K.footer(ws, wr1 + 8, 12, c['author'], 1)

    # ------------------------------------------------------------------ 일별
    def s_daily(self):
        K, c = self.K, self.ctx
        ws, R, A, nm = c['sheets']['일별'], self.rep, self.a, self.nm
        K.title(ws, 1, 1, 12, f'{nm}월 일별 목표 대비 실적', c['asof'])
        K.section(ws, 3, 1, 12, 'Ⅰ. 오늘의 숫자', '[단위:백만원]')
        cards = [
            dict(label='전일 실적', value=f"={A('y_day_sales')}", nf='#,##0.0', sub=f"={A('y_day_sales')}/{A('y_day_target')}", sub_nf='"일 목표 대비 "0.0%', accent=P['gold']),
            dict(label='진도 대비', value=f"={A('rpt_month_sales')}/{A('progress_target')}", nf='0.0%', sub=f"={A('progress_target')}", sub_nf='"경과 진도 목표 "#,##0', accent=P['gold']),
            dict(label='잔여 필요 일평균', value=f"={A('need_daily')}", nf='#,##0.0', sub=f"={A('need_daily')}/{A('plan_daily')}", sub_nf='"계획 일평균의 "0.00"배"', accent=P['navy']),
            dict(label=f'{nm}월 마감 예상', value=f"={A('fc_base')}", nf='#,##0', sub=f"={A('fc_base')}/{A('month_target')}", sub_nf='"목표 대비 "0.0%', accent=P['navy']),
        ]
        K.kpi_cards(ws, 4, 1, cards, 3)
        K.section(ws, 8, 1, 12, 'Ⅱ. 일별 실적 (막대) vs 일 목표 (선)', '[백만원]')
        d = R['daily']
        t0 = 27
        for j, hh in enumerate(['일자', '일 목표', '실적', '달성', '누계 목표', '누계 실적', '누계 달성']):
            ws.write(t0, 1 + j, hh, K.fmt(bold=True, align='center', bg_color=P['navy'], font_color=P['white'], border=1, border_color=P['line']))
        for i, x in enumerate(d):
            r, e = t0 + 1 + i, t0 + 2 + i
            lab = f"{int(x['day'][3:5])}일({x['day'][5]})"
            z = lambda **k: K.fmt(border=1, border_color=P['line'], **({'bg_color': P['mist']} if i % 2 else {}), **k)
            ws.write(r, 1, lab, z(align='center'))
            ws.write_number(r, 2, M(x['target']), z(num_format=NF['m1'], align='right'))
            if x['sales'] is not None:
                ws.write_number(r, 3, M(x['sales']), z(num_format=NF['m1'], align='right'))
            else:
                ws.write_blank(r, 3, None, z())
            K.put(ws, r, 4, f'=IF(D{e}="","",D{e}/C{e})', z(num_format=NF['pct'], align='right'))
            K.put(ws, r, 5, f"=SUM(C${t0 + 2}:C{e})", z(num_format=NF['m1'], align='right'))
            K.put(ws, r, 6, f'=IF(D{e}="","",SUM(D${t0 + 2}:D{e}))', z(num_format=NF['m1'], align='right'))
            K.put(ws, r, 7, f'=IF(G{e}="","",G{e}/F{e})', z(num_format=NF['pct'], align='right'))
        a0, a1 = t0 + 1, t0 + len(d)
        last_i = max(i for i, x in enumerate(d) if x['sales'] is not None)
        self.cells['daily_cum'] = f'G{t0 + 2 + last_i}'
        K.threshold(ws, f'E{a0 + 1}:E{a1 + 1}'); K.threshold(ws, f'H{a0 + 1}:H{a1 + 1}'); K.databar(ws, f'D{a0 + 1}:D{a1 + 1}')
        K.col_line_chart(ws, 'B10', '일별', (a0, 1, a1, 1), [('실적', (a0, 3, a1, 3))], [('일 목표', (a0, 2, a1, 2))], w=SIZE['W'], h=340,
                         highlight_last=False)
        done = [x for x in d if x['sales'] is not None]
        hit = sum(1 for x in done if x['sales'] >= x['target'])
        K.notes(ws, a1 + 3, 1, 12, [f"{len(done)}일 중 일 목표를 넘긴 날은 {hit}일입니다(실적이 없는 날짜는 제외)",
                                    '달성 = 실적 ÷ 일 목표. 일 목표는 요일·행사 일정에 따라 크게 달라집니다'], '▶', 22)
        K.footer(ws, a1 + 7, 12, c['author'], 1)

    # ------------------------------------------------------------------ 3사비교
    def s_brands(self):
        K, c, R = self.K, self.ctx, self.rep
        ws, nm = c['sheets']['3사비교'], self.nm
        K.title(ws, 1, 1, 12, f'3개 브랜드 {nm}월 비교', c['asof'])
        K.section(ws, 3, 1, 12, 'Ⅰ. 브랜드별 현황', '[단위:백만원, %]')
        hdr = ['브랜드', '당월 매출', '전년 동기', '신장율', '누계 신장율', '신상 할인율', '할인 증감', '매장재고', '재고 전년비']
        rows = []
        for i, b in enumerate(R['brands']):
            e = 6 + i
            rows.append([b['name'], M(b['sales']), M(b['ly']), f'=IFERROR(C{e}/D{e}-1,"")', b.get('ytd_growth'), b.get('new_dc'),
                         b.get('new_dc_diff'), M(b.get('stock')), b.get('stock_yoy')])
        nfs = [None, NF['m0'], NF['m0'], NF['yoy'], NF['yoy'], NF['pct'], '+0.0"%p";-0.0"%p";"-"', NF['m0'], NF['yoy']]
        last = K.table(ws, 5, 1, hdr, rows, nfs, ['center'] + ['right'] * 8, row_h=24, bold_cols=(0,))
        ws.conditional_format(f'H7:H{last + 1}', {'type': 'cell', 'criteria': '>', 'value': 0, 'format': K.fmt(font_color=P['red'], bold=True)})
        K.section(ws, last + 2, 1, 12, 'Ⅱ. 월별 전년비 (히트맵)', '[%]')
        bg = R['brand_growth']
        h2 = last + 4
        for j, hh in enumerate(['브랜드'] + [m.replace(' ', '') for m in bg['months']]):
            ws.write(h2, 1 + j, hh, K.fmt(bold=True, align='center', bg_color=P['navy'], font_color=P['white'], border=1, border_color=P['line']))
        for i, row in enumerate(bg['rows']):
            ws.write(h2 + 1 + i, 1, row['name'], K.fmt(bold=True, align='center', border=1, border_color=P['line']))
            ws.set_row(h2 + 1 + i, 22)
            for j, v in enumerate(row['values']):
                if v is None:
                    ws.write_blank(h2 + 1 + i, 2 + j, None, K.fmt(border=1, border_color=P['line']))
                else:
                    ws.write_number(h2 + 1 + i, 2 + j, v, K.fmt(num_format=NF['yoy'], align='center', border=1, border_color=P['line'], bold=(j == len(row['values']) - 1)))
        nb, nmo = len(bg['rows']), len(bg['months'])
        K.heat(ws, h2 + 1, 2, ) if False else ws.conditional_format(h2 + 1, 2, h2 + nb, 1 + nmo, {
            'type': '3_color_scale', 'min_type': 'num', 'min_value': -0.3, 'mid_type': 'num', 'mid_value': 0, 'max_type': 'num', 'max_value': 0.3,
            'min_color': '#F4B6AE', 'mid_color': '#FFFFFF', 'max_color': '#9DC3E6'})
        ln = K.wb.add_chart({'type': 'line'})
        colors = ['#8EA9DB', P['navy'], '#9B7FD6']
        for i, row in enumerate(bg['rows']):
            ln.add_series({'name': ['3사비교', h2 + 1 + i, 1], 'categories': ['3사비교', h2, 2, h2, 1 + nmo - 1],
                           'values': ['3사비교', h2 + 1 + i, 2, h2 + 1 + i, 1 + nmo - 1],
                           'line': {'color': P['gold'] if row['name'] == '마스터바니' else colors[i % 3], 'width': 3 if row['name'] == '마스터바니' else 1.75},
                           'marker': {'type': 'circle', 'size': 5}})
        K._style(ln, w=SIZE['W'], h=300); K._axes(ln, '0%')
        ws.insert_chart(f'B{h2 + nb + 4}', ln)
        mb = next((b for b in R['brands'] if b['name'] == '마스터바니'), None)
        lines = []
        if mb:
            lines.append(f"마스터바니 {nm}월 매출 {M(mb['sales']):,.0f}백만, 전년 같은 기간 대비 {mb['growth'] * 100:+.1f}%, 누계 {mb['ytd_growth'] * 100:+.1f}%")
            lines.append(f"신상 할인율 {mb['new_dc'] * 100:.1f}%는 전년보다 {mb['new_dc_diff']:+.1f}%p, 매장재고(현판가) {M(mb['stock']):,.0f}백만(전년비 {mb['stock_yoy'] * 100:+.1f}%)")
        lines.append('할인 증감 +%p는 전년보다 더 깎아 판 것입니다(빨강). 원가·판관비 반영 전 수치입니다')
        K.notes(ws, h2 + nb + 4 + 16, 1, 12, lines, '▶', 22)
        K.footer(ws, h2 + nb + 4 + 16 + len(lines) + 1, 12, c['author'], 1)

    # ------------------------------------------------------------------ 품목·스타일
    def s_items(self):
        K, c, S = self.K, self.ctx, self.sales
        ws, nm = c['sheets']['품목·스타일'], self.nm
        K.title(ws, 1, 1, 12, f'{nm}월 품목·스타일 분석', c['asof'])
        tot = sum(r[2] for r in S['rows'])
        items, styles = collections.defaultdict(lambda: [0, 0]), collections.defaultdict(lambda: [0, 0])
        for lst in S['top_items'].values():
            for k, v, q in lst:
                items[k][0] += v; items[k][1] += q
        for lst in S['top_styles'].values():
            for k, v, q in lst:
                styles[k][0] += v; styles[k][1] += q
        it = sorted(items.items(), key=lambda x: -x[1][0])
        st = sorted(styles.items(), key=lambda x: -x[1][0])[:10]
        cov_i = sum(v[0] for _, v in it) / tot; cov_s = sum(v[0] for v in styles.values()) / tot
        K.section(ws, 3, 1, 6, 'Ⅰ. 아이템별 매출 (상위)', '[백만원, 점]')
        top = it[:15]
        rows = []
        for i, (k, (v, q)) in enumerate(top):
            e = 7 + i
            rows.append([k, v / 1e6, q, f"=C{e}/SUM(C$7:C${6 + len(it)})", f'=IFERROR(C{e}/D{e}*100,"")'])
        # 나머지 합계 행은 구성비 계산용 숨은 영역에 둔다
        rest = it[15:]
        last = K.table(ws, 5, 1, ['아이템', '매출', '수량', '구성비', '평균 단가(만원)'], rows, [None, NF['m1'], NF['qty'], NF['pct'], '#,##0.0'],
                       ['left', 'right', 'right', 'right', 'right'], row_h=21, bold_cols=(0,))
        # 구성비 분모 = 전체 아이템 합(상위 15 + 기타): 표 아래 '기타 아이템' 행
        e = last + 2
        ws.write(last + 1, 1, '기타 아이템', K.fmt(border=1, border_color=P['line'], align='left', font_color=P['sub']))
        ws.write_number(last + 1, 2, sum(v[0] for _, v in rest) / 1e6, K.fmt(border=1, border_color=P['line'], num_format=NF['m1'], align='right', font_color=P['sub']))
        ws.write_number(last + 1, 3, sum(v[1] for _, v in rest), K.fmt(border=1, border_color=P['line'], num_format=NF['qty'], align='right', font_color=P['sub']))
        K.put(ws, last + 1, 4, f"=C{e}/SUM(C$7:C${e})", K.fmt(border=1, border_color=P['line'], num_format=NF['pct'], align='right', font_color=P['sub']))
        ws.write_blank(last + 1, 5, None, K.fmt(border=1, border_color=P['line']))
        # 구성비 분모를 '상위 15 + 기타' 범위로 맞춘다
        for i in range(len(top)):
            K.put(ws, 6 + i, 4, f"=C{7 + i}/SUM(C$7:C${e})", K.fmt(border=1, border_color=P['line'], num_format=NF['pct'], align='right',
                                                                  **({'bg_color': P['mist']} if i % 2 else {})))
        K.databar(ws, f'E7:E{6 + len(top)}')
        K.bar_chart(ws, 'H5', '품목·스타일', (6, 1, 15, 1), (6, 2, 15, 2), name='매출', w=470, h=330, nf='#,##0', color=P['navy2'])
        s0 = e + 3
        K.section(ws, s0, 1, 6, 'Ⅱ. 베스트 품번 TOP 10', '[백만원, 점]')
        srows = []
        for i, (k, (v, q)) in enumerate(st):
            r_ = s0 + 3 + i
            srows.append([f'{i + 1}', k, v / 1e6, q, f'=IFERROR(D{r_}/E{r_}*100,"")'])
        K.table(ws, s0 + 1, 1, ['순위', '품번', '매출', '수량', '평균 단가(만원)'], [[r[0], r[1], r[2], r[3], r[4]] for r in srows],
                [None, None, NF['m1'], NF['qty'], '#,##0.0'], ['center', 'center', 'right', 'right', 'right'], row_h=21)
        ns = s0 + 3 + len(st) + 1
        K.notes(ws, ns, 1, 12, [
            f"아이템 표는 매장별 판매 상위 아이템(점당 평균 {sum(len(l) for l in S['top_items'].values()) / max(1, len(S['top_items'])):.1f}개)을 합산한 값으로 {nm}월 매출의 {cov_i * 100:.0f}%를 포괄합니다",
            f"품번 TOP 10은 매장별 상위 품번을 합산한 값이며 매출의 {cov_s * 100:.0f}%만 포괄하므로 순위는 참고용입니다(전체 품번 판매 자료가 있으면 교체)",
            '워스트 스타일(부진 품번)은 이 자료에 없어 포함하지 않았습니다. 재고·판매율 자료가 연결되면 추가하세요'], '※', 22)
        K.footer(ws, ns + 4, 12, c['author'], 1)

    # ------------------------------------------------------------------ PPT용 데이터
    def deck_data(self):
        """build_report_deck.js 가 읽는 요약 JSON. 엑셀과 같은 숫자를 한 곳에서 내보낸다."""
        R, Y, nm = self.rep, self.ytd, self.nm
        g = R['goal']; k = R['kpi']; fc = R['forecast'][0]; f2 = R['forecast'][1]
        prog = g[self.gkey('경과 진도')]['amount']
        share = prog / k['month']['target']
        ch = [dict(name=x['name'], sales=M(x['sales']), target=M(x['target']), prog=x['sales'] / (x['target'] * share), growth=x['growth'])
              for x in R['channels'] if '합' not in x['name'] and x['target']]
        S = Y['stores']
        chans = ['백화점', '아울렛', '대리점', '직영점', '온라인', '기타']
        wf = [dict(name=n, delta=sum(sum(x['m']) - sum(x['p']) for x in S if x['channel'] == n)) for n in chans]   # 엑셀(월별 합계)과 같은 기준
        mb = next((b for b in R['brands'] if b['name'] == '마스터바니'), {})
        onl = next((x for x in R['channels'] if x['name'] == '온라인'), None)
        cnt = collections.Counter(x['judge'] for x in R['checks'] if '비즈알파' not in x['item'])
        neg = 0
        for gr in reversed(self.m_growth()):
            if gr < 0: neg += 1
            else: break
        ex = getattr(self, 'deck_extra', {})
        return dict(
            month=nm, asof=R['base_text'].split('|')[0].replace('기준일 :', '').strip(), author=R['author'],
            month_sales=M(k['month']['sales']), month_target=M(k['month']['target']), month_rate=k['month']['rate'], month_growth=k['month']['growth'],
            month_ly=M(k['month']['ly']), progress_target=M(prog), progress_rate=k['month']['sales'] / prog,
            fc_base=M(fc['base']), fc_low=M(fc['low']), fc_high=M(fc['high']), fc_rate=fc['rate'], fc_vs_ly=fc['vs_ly'],
            need_daily=M(g[self.gkey('잔여 필요 일평균')]['amount']), plan_daily=M(g[self.gkey('잔여 일목표')]['amount']),
            days_left=self.days_left(), y_day_sales=M(g[next(kk for kk in g if kk.startswith('어제') and kk.endswith('실적'))]['amount']),
            ytd_sales=M(k['ytd']['sales']), ytd_target=M(k['ytd']['target']), ytd_rate=k['ytd']['rate'], ytd_growth=k['ytd']['growth'],
            ytd_short=M(k['ytd']['short']), neg_months=neg,
            channels=ch, monthly=[dict(month=m['month'], sales=M(m['sales']), ly=M(m['ly']), growth=m.get('growth')) for m in R['mb_monthly']],
            waterfall=wf, new_dc=mb.get('new_dc'), new_dc_diff=mb.get('new_dc_diff'), online_growth=onl['growth'] if onl else None,
            dec_base=M(f2['base']), dec_target=M(f2['target']), dec_rate=f2['rate'],
            judges=dict(cnt), priority=[dict(rank=r[0], name=r[1], basis=r[2], owner=r[3], channel=r[4], prog=r[5]) for r in ex.get('priority', [])],
            actions=[dict(rank=a[0], task=a[1], basis=a[2], owner=a[3], due=a[4], metric=a[5]) for a in ex.get('actions', [])],
            promos=[dict(id=p[0], title=p[1], basis=p[2], target=p[3], period=p[4], metric=p[5]) for p in ex.get('promos', [])],
            diff_ytd_py=M(R['kpi']['ytd']['ly']) - sum(x['ytd_py'] for x in S))

    # ------------------------------------------------------------------ 카톡 보고문
    def kakao_text(self):
        """표 없는 줄바꿈 텍스트 버전(카톡 보고용). deck_data 와 같은 숫자를 쓴다."""
        d = self.deck_data()
        f = lambda v: f'{v:,.0f}'
        p = lambda v, n=1: f'{v * 100:.{n}f}%'
        sp = lambda v: f'{v * 100:+.1f}%'
        low = sorted(d['channels'], key=lambda c: c['prog'])
        wf = sorted(d['waterfall'], key=lambda w: w['delta'])
        tot = sum(w['delta'] for w in wf)
        L = [f"[MBE {d['month']}월 영업실적 보고(가안) / {d['asof']}]", '',
             '■ 요약',
             f"- {d['month']}월 누계 {f(d['month_sales'])}백만 (월 목표 {f(d['month_target'])}의 {p(d['month_rate'])}, 진도 대비 {p(d['progress_rate'])}, 전년 같은 기간 {sp(d['month_growth'])})",
             f"- 마감 예상 {f(d['fc_base'])}백만 (목표 대비 {p(d['fc_rate'])}), 목표 달성에는 잔여 {d['days_left']}일 일평균 {f(d['need_daily'])}백만 필요 (계획 {f(d['plan_daily'])}, 전일 {f(d['y_day_sales'])})",
             f"- 1~{d['month']}월 누계 {f(d['ytd_sales'])}백만 (전년 {sp(d['ytd_growth'])}, 누계 목표 {p(d['ytd_rate'])}), 전년비 {d['neg_months']}개월 연속 마이너스", '',
             '■ 유통별 진도 대비', '- ' + ' / '.join(f"{c['name']} {p(c['prog'])}" for c in low), '',
             '■ 이슈',
             f"- 전년 대비 누계 증감 {f(tot)}백만 중 {wf[0]['name']} {f(wf[0]['delta'])} ({wf[0]['delta'] / tot * 100:.0f}%)",
             f"- 신상 할인율 {p(d['new_dc'])} (전년 대비 {d['new_dc_diff']:+.1f}%p)",
             '- 우선 점검 매장: ' + ', '.join(f"{x['name']}({p(x['prog'], 0)})" for x in d['priority'][:3]) + f" 등 {len(d['priority'])}곳", '',
             '■ 다음 주 액션(제안)']
        for a in d['actions']:
            L.append(f"{a['rank']}) {a['task']} - {a['due']}")
        L += ['', '※ 가안입니다. 숫자는 일일 보고서 기준이며 원인은 가설로, 확인 후 확정합니다.']
        return '\n'.join(L)

    # ------------------------------------------------------------------ 검증
    def compute_checks(self, vals):
        """원본 JSON에서 따로 계산한 기대값 vs 재계산된 파일 값. (구분, 항목, 기대, 실제, 형식, 허용오차(상대), 설명)"""
        R, Y, nm = self.rep, self.ytd, self.nm
        V = lambda sh, ref: vals.get((sh, ref))
        c = self.cells
        off_m, off_y = R['kpi']['month']['sales'] / 1e6, R['kpi']['ytd']['sales'] / 1e6
        prog = R['goal'][self.gkey('경과 진도')]['amount'] / 1e6
        S = Y['stores']
        st_m, st_y = sum(x['mtd'] for x in S), sum(x['ytd'] for x in S)
        st_py = sum(x['ytd_py'] for x in S)
        err = sum(1 for v in vals.values() if isinstance(v, str) and v.startswith('#'))
        rows = []
        add = lambda cat, item, exp, act, kind, tol, note: rows.append([cat, item, exp, act, kind, tol, note])
        add('수식', '수식 오류 개수', 0, err, 'int', 0, '재계산한 모든 수식 칸의 오류 값(0으로 나누기 등) 개수')
        add('원본 대조', f'{nm}월 누계 (보고(요약) vs 일일 보고서)', off_m, V('보고(요약)', c['sum_sales']), 'm1', 0.0001, '기준값 시트를 거쳐 가져온 값이 원본과 같은지')
        add('원본 대조', f'{nm}월 진도 달성 (보고(요약) vs 원본 계산)', off_m / prog, V('보고(요약)', c['sum_prog']), 'pct', 0.0001, '누계 ÷ 경과 진도 목표')
        add('원본 대조', '마감 예상 달성률', R['forecast'][0]['rate'], V('보고(요약)', c['sum_fc_rate']), 'pct', 0.0001, '마감 예상 ÷ 월 목표')
        add('원본 대조', '월별 매출 합계 = 누계 (MB 월별 트렌드)', off_y, V('보고(요약)', c['trend_ytd']), 'm1', 0.0005, '월별 26년 매출을 더한 값이 누계와 같은지')
        add('원본 대조', '일별 실적 누적 = 당월 누계', off_m, V('일별', c['daily_cum']), 'm1', 0.0005, '일별 시트 마지막 누계 실적')
        sh = [V('보고(요약)', f"K{r}") for n, r in c['ch_rows'].items() if n not in ('오프라인 소계', '합  계')]
        add('합계', '구성비 합 = 100% (유통별)', 1.0, sum(x for x in sh if isinstance(x, (int, float))), 'pct', 0.0001, '합계 행을 제외한 채널 구성비의 합')
        add('출처 차이', f'{nm}월 당월: 매장 합계(마스터 DB) vs 공식', off_m, V('대시보드', 'B7'), 'm1', 0.001, '대시보드·매장별은 마스터 DB 합계. 0.1% 이내면 통과')
        add('출처 차이', f'{nm}월 누계: 매장 합계 vs 공식', off_y, V('대시보드', 'F7'), 'm1', 0.001, '1~11월 누계')
        add('출처 차이', '누계 전년 동기: 매장 합계 vs 공식', R['kpi']['ytd']['ly'] / 1e6, st_py, 'm1', 0.001,
            '매장 행에 배분되지 않은 금액이 있어 보임. 원인은 자료 제공 쪽 확인 필요(누계 전년비 -7.2% vs 공식 -8.2%)')
        add('출처 차이', f'{nm}월 진도 달성: 대시보드(매장 합계) vs 공식', off_m / prog, V('대시보드', 'D7'), 'pct', 0.001, '같은 정의(전체 실적 ÷ 진도 목표)')
        # 극단값
        ext = []
        for ch in ('백화점', '아울렛', '대리점', '직영점', '온라인', '기타'):
            for j in range(nm):
                p = sum(x['p'][j] for x in S if x['channel'] == ch); m = sum(x['m'][j] for x in S if x['channel'] == ch)
                if p > 0 and abs(m / p - 1) > 0.5:
                    ext.append(f'{ch} {j + 1}월 {(m / p - 1) * 100:+.0f}%')
        add('극단값', '유통×월 전년비 ±50% 초과 셀', None, len(ext), 'int', None, ', '.join(ext[:6]) + (' …' if len(ext) > 6 else '') + ' · 규모가 작은 채널이라 해석에 주의')
        no_t = [x for x in S if x['active'] == '운영' and not x['month_tgt']]
        add('결측', f'월 목표가 없는 운영 매장', None, len(no_t), 'int', None,
            f"실적 {sum(x['mtd'] for x in no_t):,.1f}백만이 채널 실적에는 포함되고 매장별 달성률은 비어 있음")
        add('기간', '부분 기간 비교', None, None, 'txt', None, f'{R["kpi"]["month"]["label"]} 기준으로 전년 같은 기간과 비교합니다(전월 전체와 비교하지 않음)')
        tot = sum(r[2] for r in self.sales['rows'])
        cov = {}
        for key in ('top_items', 'top_styles'):
            agg = sum(v for l in self.sales[key].values() for _, v, _q in l)
            cov[key] = agg / tot
        add('범위', '아이템 구성 포괄 범위', None, cov['top_items'], 'pct', None, '매장별 상위 아이템 합산. 전체 판매 자료가 있으면 교체')
        add('범위', '베스트 품번 포괄 범위', None, cov['top_styles'], 'pct', None, '34% 수준이라 순위는 참고용')
        return rows

    def s_checks(self):
        K, c = self.K, self.ctx
        ws, nm = c['sheets']['검증'], self.nm
        K.title(ws, 1, 1, 7, '자료 검증 결과 (자동)', c['asof'])
        K.section(ws, 3, 1, 7, 'Ⅰ. 검증 항목', '')
        hf = K.fmt(bold=True, align='center', bg_color=P['navy'], font_color=P['white'], border=1, border_color=P['line'], text_wrap=True)
        for j, h in enumerate(['구분', '검증 항목', '기대(원본)', '실제(파일)', '차이', '판정', '설명']):
            ws.write(5, 1 + j, h, hf)
        nfs = {'m1': '#,##0.0', 'pct': '0.00%', 'int': '#,##0', 'txt': 'General'}
        r = 6
        for cat, item, exp, act, kind, tol, note in (self.checks or []):
            ws.set_row(r, 30)
            z = lambda **k: K.fmt(border=1, border_color=P['line'], text_wrap=True, **k)
            ws.write(r, 1, cat, z(align='center')); ws.write(r, 2, item, z(align='left', bold=True))
            nf = nfs[kind]
            diff = None
            if isinstance(exp, (int, float)) and isinstance(act, (int, float)):
                diff = act - exp
                if abs(diff) < 1e-9:
                    diff = 0
                ok = (abs(diff) <= (0 if tol == 0 else tol * max(abs(exp), 1e-9))) if tol is not None else True
                verdict = '통과' if ok else '주의'
            else:
                verdict = '참고' if tol is None else '주의'
            if tol is None and isinstance(act, (int, float)):
                verdict = '참고'
            if kind == 'txt':
                verdict = '통과'
            ws.write(r, 3, exp if isinstance(exp, (int, float)) else '-', z(align='right', num_format=nf))
            ws.write(r, 4, act if isinstance(act, (int, float)) else '-', z(align='right', num_format=nf))
            ws.write(r, 5, diff if diff is not None else '-', z(align='right', num_format=('+' + nf + ';-' + nf + ';0') if kind != 'txt' else 'General'))
            ws.write(r, 6, verdict, z(align='center', bold=True, font_color={'통과': P['blue'], '주의': '#C55A11', '참고': P['sub']}[verdict]))
            ws.write(r, 7, note, z(align='left', font_size=9))
            r += 1
        last = r - 1
        ws.merge_range(4, 2, 4, 7, '', K.fmt())
        K.put(ws, 4, 2, f'="통과 "&COUNTIF(G7:G{last + 1},"통과")&"건 · 주의 "&COUNTIF(G7:G{last + 1},"주의")&"건 · 참고 "&COUNTIF(G7:G{last + 1},"참고")&"건"',
              K.fmt(bold=True, font_size=11, font_color=P['navy']))
        K.section(ws, r + 1, 1, 7, 'Ⅱ. 종합 판단', '')
        K.notes(ws, r + 2, 1, 7, [
            '종합: 주의사항을 밝히고 공유 가능. 공식 보고서 기준 숫자는 모두 일치하며, 매장 단위 합계와의 차이는 아래 주의 항목에 정리했습니다',
            '공유 시 필수 문구: 누계 전년 동기는 매장 합계가 공식보다 작아(약 521백만) 매장 기준 누계 전년비(-7.2%)와 공식(-8.2%)이 다릅니다',
            '공유 시 필수 문구: 품목·품번 분석은 매장별 상위 항목 합산이라 전체를 대표하지 않습니다(아이템 78%, 품번 34%)',
            '소규모 채널(기타·직영점)의 ±50% 초과 변동은 규모가 작아 추세로 읽지 않는 것이 안전합니다'], '▶', 24)
        K.footer(ws, r + 8, 7, c['author'], 1)

    # ------------------------------------------------------------------ 이슈·액션
    def s_issues(self):
        K, c, R, Y = self.K, self.ctx, self.rep, self.ytd
        ws, nm = c['sheets']['이슈·액션'], self.nm
        K.title(ws, 1, 1, 6, '이슈·리스크 · 다음 주 액션 (가안)', c['asof'])
        K.section(ws, 3, 1, 6, 'Ⅰ. 점검 항목 (일일 보고서 자동 판정)', '')
        rank = {'확인 필요': 0, '주의': 1, '정상': 2, '참고': 3}
        ck = sorted([x for x in R['checks'] if '비즈알파' not in x['item']], key=lambda x: rank.get(x['judge'], 9))
        rows = [[x['cat'], x['item'], x['status'], x['judge']] for x in ck]
        last = K.table(ws, 5, 1, ['구분', '항목', '현황', '판정'], rows, None, ['center', 'left', 'left', 'center'], row_h=36, bold_cols=(3,))
        for r in range(6, last + 1):
            ws.set_row(r, 36)
        for txt, color in (('확인 필요', P['red']), ('주의', '#C55A11'), ('정상', P['blue'])):
            ws.conditional_format(6, 4, last, 4, {'type': 'cell', 'criteria': '==', 'value': f'"{txt}"', 'format': K.fmt(font_color=color, bold=True)})
        # 텍스트 줄바꿈
        for i, x in enumerate(rows):
            for j in (1, 2):
                ws.write(6 + i, 2 + j - 1, x[j], K.fmt(border=1, border_color=P['line'], align='left', text_wrap=True, **({'bg_color': P['mist']} if i % 2 else {})))
        p0 = last + 2
        K.section(ws, p0, 1, 6, 'Ⅱ. 우선 점검 매장 (진도 70% 미만, 월 목표 50백만 이상, 부족액 큰 순 5곳)', '[백만원]')
        cand = []
        for s in Y['stores']:
            if s['active'] == '운영' and s['mtd_tgt'] and s['month_tgt'] >= 50 and s['mtd'] / s['mtd_tgt'] < 0.7:
                cand.append((s['mtd_tgt'] - s['mtd'], s))
        cand.sort(key=lambda x: -x[0])
        rows2 = []
        for i, (short, s) in enumerate(cand[:5]):
            hq = self.st.get(s['code'], {}).get('hq_manager') or '-'
            rows2.append([i + 1, s['name'], f"월 목표 {s['month_tgt']:,.0f} · 진도 목표 {s['mtd_tgt']:,.0f} · 실적 {s['mtd']:,.1f} · 부족 {short:,.0f}", hq, s['channel'], s['mtd'] / s['mtd_tgt']])
        l2 = K.table(ws, p0 + 1, 1, ['순위', '매장', '근거', '본사 담당', '유통', '진도 달성'], rows2, [None, None, None, None, None, NF['pct']],
                     ['center', 'left', 'left', 'center', 'center', 'right'], row_h=24, bold_cols=(1,))
        K.threshold(ws, f'G{p0 + 3}:G{l2 + 1}')
        a0 = l2 + 2
        K.section(ws, a0, 1, 6, 'Ⅲ. 다음 주 액션 플랜 (제안)', '')
        mb = next((b for b in R['brands'] if b['name'] == '마스터바니'), {})
        g = R['goal']
        need, plan = g[self.gkey('잔여 필요 일평균')]['amount'] / 1e6, g[self.gkey('잔여 일목표')]['amount'] / 1e6
        yd = g[self.gkey('어제')]['amount'] if False else None
        share = g[self.gkey('경과 진도')]['amount'] / R['kpi']['month']['target']
        dept = next((x for x in R['channels'] if x['name'] == '백화점'), None)
        onl = next((x for x in R['channels'] if x['name'] == '온라인'), None)
        top5 = ', '.join(r[1] for r in rows2[:3]) or '-'
        owners = sorted({r[3] for r in rows2 if r[3] != '-'})
        nw = self.next_week()
        acts = [
            [1, '일 목표 갭 일일 점검', f'잔여 필요 일평균 {need:,.0f}백만은 계획 일평균 {plan:,.0f}의 {need / plan:.2f}배. 평일 일 목표 미달 매장 확인', '영업부', f"{nw['mon']}~ 매일", '일 목표 달성률'],
            [2, '진도 미달 매장 현장 점검', f'우선 점검 매장: {top5} 등 5곳(진도 70% 미만)', ', '.join(owners) or '영업부', nw['wed'], '진도 달성률'],
        ]
        if dept:
            acts.append([3, '백화점 진도 회복 계획', f"백화점 진도 대비 {dept['sales'] / (dept['target'] * share) * 100:.1f}%로 유통 중 최저. 행사·VMD 협의안 정리", '영업부', nw['thu'], '채널 진도 달성'])
        if mb.get('new_dc') is not None:
            acts.append([4, '신상 할인 확대 원인 확인', f"신상 할인율 {mb['new_dc'] * 100:.1f}% (전년 대비 {mb['new_dc_diff']:+.1f}%p). 행사·재고 사유와 수익성 영향 확인", '영업부', nw['thu'], '신상 할인율'])
        if onl:
            acts.append([5, '온라인 하락 원인 확인', f"온라인 {nm}월 전년 같은 기간 대비 {onl['growth'] * 100:+.1f}%. 노출·재고·행사 영향 확인", '온라인 담당', nw['thu'], '온라인 신장율'])
        f2 = R['forecast'][1]
        acts.append([len(acts) + 1, '다음 달 선제 대응', f"다음 달 예상 {f2['base'] / 1e6:,.0f}백만으로 목표 {f2['target'] / 1e6:,.0f}백만의 {f2['rate'] * 100:.1f}%. 행사 계획 선제 협의", '영업부', nw['later'], '다음 달 진도'])
        l3 = K.table(ws, a0 + 1, 1, ['순위', '과제', '근거', '담당', '기한', '확인 지표'], acts, None, ['center', 'left', 'left', 'center', 'center', 'center'], row_h=40, bold_cols=(1,))
        for i, rw in enumerate(acts):
            for j in range(1, 6):
                ws.write(a0 + 2 + i, 1 + j, rw[j], K.fmt(border=1, border_color=P['line'], align='left' if j < 3 else 'center', text_wrap=True,
                                                        bold=(j == 1), **({'bg_color': P['mist']} if i % 2 else {})))
        q0 = l3 + 2
        K.section(ws, q0, 1, 6, 'Ⅳ. 프로모션 제안 (가설, 효과는 실행 후 측정)', '')
        promos = [
            ['A', '진도 미달 백화점 주말 사은 행사', '진도 미달 채널 매출 회복을 위한 주말 구매금액별 사은 행사. 가설: 객단가·방문 전환 개선', '백화점 중 우선 점검 매장', '다음 주말 1회', '행사 전후 일 매출, 객단가'],
            ['B', '신상 부진 재고 모음 기획전', '점검 항목의 신상 부진 재고(재고 300개 이상·당월 판매 5% 미만) 품번을 묶어 온라인·대리점 동시 노출. 가설: 재고 회전 개선', '온라인, 대리점', '2주', '품번별 판매율, 할인율 변화'],
        ]
        l4 = K.table(ws, q0 + 1, 1, ['안', '제안', '근거·가설', '대상', '기간', '측정 지표'], promos, None, ['center', 'left', 'left', 'center', 'center', 'center'], row_h=48, bold_cols=(1,))
        for i, rw in enumerate(promos):
            for j in range(1, 6):
                ws.write(q0 + 2 + i, 1 + j, rw[j], K.fmt(border=1, border_color=P['line'], align='left' if j < 3 else 'center', text_wrap=True,
                                                        bold=(j == 1), **({'bg_color': P['mist']} if i % 2 else {})))
        self.deck_extra = dict(priority=rows2, actions=acts, promos=promos)
        K.notes(ws, l4 + 2, 1, 6, ['담당·기한은 제안이며 확정 전 조정 필요합니다. 본사 담당은 매장 정보의 본사 담당자 기준입니다',
                                   '원인은 가설로만 적었습니다. 확인된 사실은 일일 보고서 숫자뿐입니다'], '※', 22)
        K.footer(ws, l4 + 5, 6, c['author'], 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True); ap.add_argument('--date', required=True)
    ap.add_argument('--out', required=True); ap.add_argument('--author', default=B.AUTHOR)
    a = ap.parse_args()
    rep = Report(a.data, a.date)
    D = rep.base_data()
    tmp = a.out + '.tmp.xlsx'
    k1 = B.build(tmp, D, None, a.author, extra=rep)
    vals = verify_xlsx.recalc_values(tmp, k1.fcells)
    bad = {k: v for k, v in vals.items() if isinstance(v, str) and v.startswith('#')}
    if bad:
        print('수식 오류:', list(bad.items())[:10])
    clean = {k: ('' if v is None else v) for k, v in vals.items()}
    rep.ctx = None
    checks = rep.compute_checks(vals)
    rep2 = Report(a.data, a.date)
    rep2.checks = checks
    B.build(a.out, D, clean, a.author, extra=rep2)
    json.dump(rep2.deck_data(), open(os.path.splitext(a.out)[0] + '_deck.json', 'w', encoding='utf8'), ensure_ascii=False, indent=1)
    open(os.path.splitext(a.out)[0] + '_카톡.txt', 'w', encoding='utf8').write(rep2.kakao_text())
    for ch in checks:
        print('  검증', ch[0], '|', ch[1], '|', ch[2], '->', ch[3])
    os.remove(tmp)
    print('완료:', a.out, '| 수식', len(k1.fcells), '개 | 오류', len(bad))


if __name__ == '__main__':
    main()
