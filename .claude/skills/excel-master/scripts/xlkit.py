"""xlkit: 영업 엑셀을 임팩트 있게 만드는 xlsxwriter 헬퍼.

원칙: 수식은 살아 있게(값 복붙 금지), 숫자서식은 회사 관례(▲ 빨강), 시각 요소는 절제된 한 가지 팔레트.
2단계 빌드: 1차로 만들고 LibreOffice로 재계산 -> 값을 모아 cached 로 넘겨 다시 만들면
휴대폰·미리보기에서도 수식 칸에 숫자가 보입니다(verify_xlsx.py 참고).
"""
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell as rc, xl_col_to_name as cn

FONT = '맑은 고딕'
P = dict(navy='#1F3864', navy2='#2F5597', steel='#8EA9DB', sky='#DEEBF7', mist='#F2F6FC',
         grey='#D9D9D9', grey2='#F2F2F2', line='#BFBFBF', ink='#262626', sub='#7F7F7F',
         gold='#FFC000', yellow='#FFFF00', blue='#0070C0', red='#FF0000',
         redsoft='#FDE9E7', greensoft='#E2F0D9', white='#FFFFFF')
NF = dict(
    m0='#,##0;[Red]"▲"#,##0;"-"',            # 백만원(정수), 마이너스 빨간 ▲
    m1='#,##0.0;[Red]"▲"#,##0.0;"-"',        # 백만원(소수1)
    qty='#,##0;[Red]"▲"#,##0;"-"',
    yoy='0.0%;[Red]"▲"0.0%;"-"',             # 신장율
    pct='0.0%',                              # 달성률·구성비·판매율
    won_m='#,##0,,;[Red]"▲"#,##0,,;"-"',     # 값은 원, 표시는 백만원
    won_k='#,##0,;[Red]"▲"#,##0,;"-"',       # 값은 원, 표시는 천원
    wk='0.0"주"',
)


class Kit:
    def __init__(self, path, cached=None):
        self.wb = xlsxwriter.Workbook(path, {'nan_inf_to_errors': True})
        self.cached = cached or {}
        self.fcells = []          # (시트명, 셀) 수식 칸 기록 -> 2차 빌드 때 값 주입
        self._f = {}

    # ---------- 서식 ----------
    def fmt(self, **kw):
        base = {'font_name': FONT, 'font_size': 10, 'valign': 'vcenter'}
        base.update(kw)
        key = tuple(sorted(base.items()))
        if key not in self._f:
            self._f[key] = self.wb.add_format(base)
        return self._f[key]

    def fx(self, ws, r, c, formula, fmt=None):
        """수식 쓰기(2차 빌드면 계산된 값을 같이 저장)."""
        key = (ws.get_name(), rc(r, c))
        self.fcells.append(key)
        ws.write_formula(r, c, formula, fmt, self.cached.get(key, 0))

    def put(self, ws, r, c, v, fmt=None):
        if isinstance(v, str) and v.startswith('='):
            self.fx(ws, r, c, v, fmt)
        elif v is None:
            ws.write_blank(r, c, None, fmt)
        else:
            ws.write(r, c, v, fmt)

    # ---------- 시트 ----------
    def sheet(self, name, tab=None, zoom=85, widths=None, landscape=True, hide_grid=True):
        ws = self.wb.add_worksheet(name)
        if hide_grid:
            ws.hide_gridlines(2)
        ws.set_zoom(zoom)
        if tab:
            ws.set_tab_color(tab)
        if landscape:
            ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.center_horizontally()
        ws.set_margins(left=0.3, right=0.3, top=0.35, bottom=0.35)
        for spec in (widths or []):          # [(첫열, 끝열, 너비), ...] 0부터
            ws.set_column(spec[0], spec[1], spec[2])
        return ws

    def title(self, ws, r, c1, c2, text, asof=None, style='impact'):
        """impact: 남색 배너 / company: ■ 제목 ■ (회사 양식)."""
        if style == 'company':
            ws.merge_range(r, c1, r, c2, f'■ {text} ■', self.fmt(font_size=22, bold=True, align='center'))
            ws.set_row(r, 32)
        else:
            ws.merge_range(r, c1, r, c2, f'  {text}', self.fmt(font_size=20, bold=True, font_color=P['white'],
                                                              bg_color=P['navy'], bottom=5, bottom_color=P['gold']))
            ws.set_row(r, 38)
        if asof:
            ws.write(r + 1, c1, f'- 기준일 : {asof}', self.fmt(font_size=9, font_color=P['sub']))

    def section(self, ws, r, c1, c2, text, unit=None):
        """섹션 제목 + 표 오른쪽 위 끝에 단위."""
        f = self.fmt(font_size=12, bold=True, font_color=P['navy'], bottom=2, bottom_color=P['navy'])
        ws.merge_range(r, c1, r, c2, text, f)
        if unit:
            ws.write(r, c2, unit, self.fmt(font_size=9, font_color=P['sub'], align='right',
                                           bottom=2, bottom_color=P['navy']))
        ws.set_row(r, 22)

    def footer(self, ws, r, c, author, first_col=0):
        ws.write(r, c, f'작성자 : {author}', self.fmt(font_size=9, align='right', font_color=P['sub']))
        ws.print_area(0, first_col, r, c)

    def notes(self, ws, r, c1, c2, lines, bullet='▶', height=20, bold_first=False):
        for i, t in enumerate(lines):
            ws.merge_range(r + i, c1, r + i, c2, f'{bullet} {t}',
                           self.fmt(font_size=10.5, bold=bold_first and i == 0, text_wrap=True, indent=1))
            ws.set_row(r + i, height)
        return r + len(lines)

    # ---------- KPI 카드 ----------
    def kpi_cards(self, ws, r, c, cards, span=2):
        """cards=[{label,value(수식|값),nf,sub(수식|값),sub_nf,accent}] . 카드당 span칸, 3행 차지."""
        ws.set_row(r, 18); ws.set_row(r + 1, 40); ws.set_row(r + 2, 20)
        for i, k in enumerate(cards):
            c1, c2 = c + i * span, c + i * span + span - 1
            acc = k.get('accent', P['navy'])
            box = dict(bg_color=P['mist'], left=1, right=1, left_color=P['line'], right_color=P['line'])
            lab = self.fmt(font_size=9.5, font_color=P['sub'], bold=True, top=5, top_color=acc, indent=1, **box)
            val = self.fmt(font_size=26, bold=True, font_color=P['navy'], num_format=k.get('nf', '#,##0'), indent=1, **box)
            sub = self.fmt(font_size=10, font_color=P['ink'], num_format=k.get('sub_nf', 'General'), bottom=1,
                           bottom_color=P['line'], indent=1, **box)
            ws.merge_range(r, c1, r, c2, k['label'], lab)
            ws.merge_range(r + 1, c1, r + 1, c2, '', val)
            self.put(ws, r + 1, c1, k['value'], val)
            ws.merge_range(r + 2, c1, r + 2, c2, '', sub)
            self.put(ws, r + 2, c1, k.get('sub'), sub)
        return r + 3

    # ---------- 표 ----------
    def table(self, ws, r, c, headers, rows, nfs=None, aligns=None, header='navy', total_last=False,
              zebra=True, row_h=19, head_h=26, bold_cols=(), bcolor=None):
        """bcolor='#000000' 이면 회사 양식(검정 실선), 기본은 연한 회색 선."""
        bc = bcolor or P['line']
        nfs = nfs or [None] * len(headers)
        aligns = aligns or ['center'] * len(headers)
        hbg, hfc = (P['navy'], P['white']) if header == 'navy' else (P['grey'], P['ink'])
        hf = self.fmt(bold=True, align='center', bg_color=hbg, font_color=hfc, border=1, border_color=bc, text_wrap=True)
        ws.set_row(r, head_h)
        for j, h in enumerate(headers):
            ws.write(r, c + j, h, hf)
        for i, row in enumerate(rows):
            rr = r + 1 + i
            is_tot = total_last and i == len(rows) - 1
            ws.set_row(rr, row_h)
            for j, v in enumerate(row):
                kw = dict(border=1, border_color=bc, align=aligns[j])
                if nfs[j]:
                    kw['num_format'] = nfs[j]
                if is_tot:
                    kw.update(bold=True, bg_color=P['sky'], top=2, top_color=P['navy'])
                elif zebra and i % 2 == 1:
                    kw['bg_color'] = P['mist']
                if j in bold_cols and not is_tot:
                    kw['bold'] = True
                self.put(ws, rr, c + j, v, self.fmt(**kw))
        return r + len(rows)          # 마지막 행 번호

    # ---------- 조건부 서식 ----------
    def databar(self, ws, rng, color=None):
        ws.conditional_format(rng, {'type': 'data_bar', 'bar_solid': True, 'data_bar_2010': True,
                                    'bar_color': color or P['steel'], 'bar_no_border': True})

    def heat(self, ws, rng, kind='seq'):
        if kind == 'div':      # 증감: 빨강-흰색-파랑
            ws.conditional_format(rng, {'type': '3_color_scale', 'min_type': 'num', 'min_value': -0.3,
                                        'mid_type': 'num', 'mid_value': 0, 'max_type': 'num', 'max_value': 0.3,
                                        'min_color': '#F4B6AE', 'mid_color': '#FFFFFF', 'max_color': '#9DC3E6'})
        else:                  # 순차: 흰색-남색
            ws.conditional_format(rng, {'type': '2_color_scale', 'min_color': '#FFFFFF', 'max_color': P['steel']})

    def threshold(self, ws, rng, hi=1.0, lo=0.9):
        """달성률: hi 이상 파랑 굵게, lo 미만 빨강 (회사 규칙: 글자색만)."""
        ws.conditional_format(rng, {'type': 'cell', 'criteria': '>=', 'value': hi,
                                    'format': self.fmt(font_color=P['blue'], bold=True)})
        ws.conditional_format(rng, {'type': 'cell', 'criteria': '<', 'value': lo,
                                    'format': self.fmt(font_color=P['red'])})

    def spark(self, ws, cell, data_rng, kind='column', color=None):
        ws.add_sparkline(cell, {'range': data_rng, 'type': kind, 'series_color': color or P['navy'],
                                'last_point': True, 'high_point': True, 'markers': kind == 'line', 'show_hidden': True,
                                'negative_points': True, 'negative_color': P['red'],
                                'high_color': P['gold'], 'last_color': P['navy2']})

    # ---------- 차트 ----------
    def _style(self, ch, legend='bottom', w=560, h=300):
        ch.set_title({'none': True})
        ch.set_size({'width': w, 'height': h})
        ch.set_chartarea({'border': {'none': True}, 'fill': {'color': P['white']}})
        if legend:
            ch.set_legend({'position': legend, 'font': {'name': FONT, 'size': 9}})
        else:
            ch.set_legend({'none': True})

    def _axes(self, ch, ynf='#,##0', xfont=9):
        ch.set_x_axis({'num_font': {'name': FONT, 'size': xfont, 'color': P['ink']}, 'line': {'color': P['line']}})
        ch.set_y_axis({'num_font': {'name': FONT, 'size': 9, 'color': P['sub']}, 'num_format': ynf,
                       'line': {'none': True}, 'major_gridlines': {'visible': True, 'line': {'color': '#E7E6E6', 'width': 0.75}}})

    def col_line_chart(self, ws, anchor, sheet, cats, cols, lines=(), markers=(), w=560, h=300, ynf='#,##0',
                       highlight_last=True, last_color=None):
        """막대(올해)+선(전년)+마커(목표). cols/lines/markers=[(이름, 값범위)], 범위는 (r1,c1,r2,c2)."""
        ch = self.wb.add_chart({'type': 'column'})
        for n, (name, rg) in enumerate(cols):
            npts = (rg[3] - rg[1] + 1) if rg[0] == rg[2] else (rg[2] - rg[0] + 1)
            pts = [{'fill': {'color': P['navy']}}] * (npts - 1) + [{'fill': {'color': last_color or P['gold']}}] \
                if highlight_last and n == 0 else None
            ch.add_series({'name': name, 'categories': [sheet, *cats], 'values': [sheet, *rg],
                           'fill': {'color': P['navy']}, 'gap': 60, 'points': pts,
                           'data_labels': ({'value': True, 'num_format': '#,##0', 'position': 'inside_end',
                                            'font': {'name': FONT, 'size': 8, 'bold': True, 'color': P['white']},
                                            'custom': [None] * (npts - 1) + [{'font': {'name': FONT, 'size': 8, 'bold': True, 'color': P['ink']}}]}
                                           if n == 0 else {'value': False})})
        ln = self.wb.add_chart({'type': 'line'})
        for name, rg in lines:
            ln.add_series({'name': name, 'categories': [sheet, *cats], 'values': [sheet, *rg],
                           'line': {'color': '#7F7F7F', 'width': 2, 'dash_type': 'dash'},
                           'marker': {'type': 'circle', 'size': 5, 'fill': {'color': P['white']}, 'border': {'color': '#7F7F7F'}}})
        for name, rg in markers:
            ln.add_series({'name': name, 'categories': [sheet, *cats], 'values': [sheet, *rg],
                           'line': {'none': True},
                           'marker': {'type': 'diamond', 'size': 8, 'fill': {'color': P['red']}, 'border': {'color': P['red']}}})
        ch.combine(ln)
        self._style(ch, w=w, h=h); self._axes(ch, ynf)
        ws.insert_chart(anchor, ch, {'object_position': 1})
        return ch

    def bar_chart(self, ws, anchor, sheet, cats, vals, name='', w=560, h=300, horizontal=True, nf='0%',
                  color=None, ref_line=None, max_=None):
        ch = self.wb.add_chart({'type': 'bar' if horizontal else 'column'})
        ch.add_series({'name': name, 'categories': [sheet, *cats], 'values': [sheet, *vals],
                       'fill': {'color': color or P['navy2']}, 'gap': 50,
                       'data_labels': {'value': True, 'num_format': nf, 'font': {'name': FONT, 'size': 9, 'bold': True}}})
        self._style(ch, legend=None, w=w, h=h)
        ch.set_y_axis({'reverse': horizontal, 'num_font': {'name': FONT, 'size': 9}, 'line': {'color': P['line']}}) if horizontal else None
        ch.set_x_axis({'num_font': {'name': FONT, 'size': 9, 'color': P['sub']}, 'num_format': nf, 'min': 0, 'max': max_,
                       'major_gridlines': {'visible': True, 'line': {'color': '#E7E6E6'}}, 'line': {'none': True}}) if horizontal else self._axes(ch, nf)
        ws.insert_chart(anchor, ch, {'object_position': 1})
        return ch

    def doughnut_chart(self, ws, anchor, sheet, cats, vals, w=360, h=300, colors=None):
        ch = self.wb.add_chart({'type': 'doughnut'})
        cols = colors or [P['navy'], P['navy2'], P['steel'], '#BDD7EE', P['gold'], P['grey']]
        ch.add_series({'categories': [sheet, *cats], 'values': [sheet, *vals],
                       'points': [{'fill': {'color': c}} for c in cols],
                       'data_labels': {'percentage': True, 'font': {'name': FONT, 'size': 9, 'bold': True, 'color': P['white']}}})
        ch.set_hole_size(55)
        self._style(ch, legend='right', w=w, h=h)
        ws.insert_chart(anchor, ch, {'object_position': 1})
        return ch

    def scatter_chart(self, ws, anchor, sheet, xs, ys, xname='', yname='', w=560, h=320, xnf='0%', ynf='0%'):
        ch = self.wb.add_chart({'type': 'scatter'})
        ch.add_series({'categories': [sheet, *xs], 'values': [sheet, *ys],
                       'marker': {'type': 'circle', 'size': 7, 'fill': {'color': P['navy2'], 'transparency': 35},
                                  'border': {'color': P['white']}}})
        self._style(ch, legend=None, w=w, h=h)
        ch.set_x_axis({'name': xname, 'name_font': {'name': FONT, 'size': 9}, 'num_format': xnf,
                       'num_font': {'name': FONT, 'size': 9}, 'major_gridlines': {'visible': True, 'line': {'color': '#E7E6E6'}}})
        ch.set_y_axis({'name': yname, 'name_font': {'name': FONT, 'size': 9}, 'num_format': ynf,
                       'num_font': {'name': FONT, 'size': 9}, 'major_gridlines': {'visible': True, 'line': {'color': '#E7E6E6'}}})
        ws.insert_chart(anchor, ch, {'object_position': 1})
        return ch

    def close(self):
        self.wb.close()
