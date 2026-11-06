# 엑셀 레시피 (xlsxwriter 기준, xlkit와 함께 사용)

실행해서 직접 볼 수 있는 파일: `python scripts/recipes_demo.py 기법카탈로그.xlsx`

## 1. 표 만들기
| 목적 | 방법 |
|---|---|
| 머리글 | 남색 `1F3864` + 흰 글씨(임팩트) 또는 회색 `D9D9D9`(회사 양식), 굵게, 가운데, 행 높이 26 |
| 줄무늬 | 짝수 행 `F2F6FC`(아주 연하게). 진한 줄무늬는 금지 |
| 합계 행 | `DEEBF7` + 위쪽 두꺼운 남색 선 + 굵게 |
| 소계 행 | `F2F2F2` + 굵게 |
| 틀 고정 | `ws.freeze_panes(머리글행+1, 구분열수)` |
| 필터 | `ws.autofilter(r1, c1, r2, c2)` |
| 인쇄 제목행 | `ws.repeat_rows(머리글행)` |
| 열 그룹(상세 접기) | `ws.set_column(c1, c2, 너비, None, {'level': 1})` |
| 행 그룹 | `ws.set_row(r, None, None, {'level': 1, 'hidden': False})` |
| 시트 간 이동 | `ws.write_url(r, c, "internal:'시트'!A1", fmt, '표시글')` |
| 메모(주석) | `ws.write_comment('B2', '출처: ERP 11/5 마감')` |
| 입력 제한 | `ws.data_validation(r, c, r, c, {'validate': 'list', 'source': ['백화점','아울렛']})` |
| 이름 정의 | `wb.define_name('목표', "='원데이터'!$B$2")` |

## 2. 조건부 서식 8종
```python
# 1) 데이터 막대
ws.conditional_format('E7:E40', {'type': 'data_bar', 'bar_solid': True, 'data_bar_2010': True, 'bar_color': '#8EA9DB', 'bar_no_border': True})
# 2) 2색 색조(순차)        3) 3색 색조(증감, 0 기준)
ws.conditional_format('L7:L40', {'type': '2_color_scale', 'min_color': '#FFFFFF', 'max_color': '#8EA9DB'})
ws.conditional_format('J7:J40', {'type': '3_color_scale', 'min_type': 'num', 'min_value': -.3, 'mid_type': 'num', 'mid_value': 0,
                                 'max_type': 'num', 'max_value': .3, 'min_color': '#F4B6AE', 'mid_color': '#FFFFFF', 'max_color': '#9DC3E6'})
# 4) 임계값 글자색          5) 아이콘(신호등)
ws.conditional_format('F7:F40', {'type': 'cell', 'criteria': '>=', 'value': 1, 'format': wb.add_format({'font_color': '#0070C0', 'bold': True})})
ws.conditional_format('H7:H40', {'type': 'icon_set', 'icon_style': '3_traffic_lights',
                                 'icons': [{'criteria': '>=', 'type': 'number', 'value': 1}, {'criteria': '>=', 'type': 'number', 'value': .9}]})
# 6) 수식 조건(행 전체 강조)  7) 상위 n개  8) 중복·텍스트 포함
ws.conditional_format('B7:M40', {'type': 'formula', 'criteria': '=$G7<0.7', 'format': wb.add_format({'bg_color': '#FDE9E7'})})
ws.conditional_format('H7:H40', {'type': 'top', 'value': 3, 'format': wb.add_format({'bold': True, 'bg_color': '#FFF2CC'})})
ws.conditional_format('N7:N40', {'type': 'text', 'criteria': 'containing', 'value': '미만', 'format': wb.add_format({'font_color': '#FF0000'})})
```
- 한 표에 시각 요소는 최대 2종류. 행 전체 강조는 정말 꼭 봐야 하는 예외에만.

## 3. 차트 선택과 모양
- 모양 공통: 3D 금지, 제목 없음(섹션 제목에 씀), 눈금선 연한 회색, 범례 아래, 데이터 레이블은 핵심 계열에만.
- 색: 올해 남색 `1F3864`, 당월 금색 `FFC000`, 전년 회색 점선, 목표 빨강 마름모.
- 위치: 표 위·옆의 격자에 맞춰 `insert_chart('B11', chart, {'object_position': 1})`. 데이터 표와 겹치지 않게 아래쪽에 모은다.
- 콤보: `column.combine(line)`; 보조축은 `'y2_axis': True`.
- 폭포: 누적 막대 4계열(기준 투명, 증가, 감소, 합계). 값 차이가 작으면 `y_axis.min`을 올린다.
- 파레토: 내림차순 막대 + 누적% 선(보조축 0~100%).
- 산점도: 분포·상관. 사분면 해석 문구를 같이 쓴다.

## 4. 스파크라인(셀 안 미니 차트)
```python
ws.add_sparkline('M7', {'range': "'원데이터'!F2:P2", 'type': 'column', 'series_color': '#1F3864', 'high_point': True, 'last_point': True, 'show_hidden': True})
```
- 숨긴 열의 데이터를 쓰려면 `'show_hidden': True`(없으면 엑셀이 그리지 않는다).

## 5. 수식 패턴(피벗 없이 집계)
| 목적 | 수식 |
|---|---|
| 조건 합계 | `=SUMIFS(합계열, 유통열, $B7, 목표열, ">0")` |
| 조건 개수 | `=COUNTIFS(유통열, $B7, 운영열, "운영")` |
| 가중 달성률 | `=SUMIFS(실적, 목표, ">0") / SUM(목표)` |
| 순위 | `=RANK(H7, $H$7:$H$106)` |
| N번째 큰 값과 이름 | `=INDEX(이름열, MATCH(LARGE(값열, k), 값열, 0))` |
| 선택값 조회 | `=INDEX(결과열, MATCH($C$6, 키열, 0))` |
| 전년비 | `=IF(전년>0, 올해/전년-1, "")` |
| 문구 합치기 | `="목표 "&TEXT(B5,"#,##0")&"백만"` |

## 6. 인쇄·공유
- 한 장 대시보드: `ws.set_portrait(); ws.fit_to_pages(1, 1)`. 상세표: `fit_to_pages(1, 0)` + `repeat_rows`.
- 인쇄 영역: `ws.print_area(r1, c1, r2, c2)`(작성자 행이 마지막).
- 확인 필요 항목은 노랑 셀 + `※ 사유`.

## 7. 한계와 대안
- xlsxwriter는 기존 파일 편집, 피벗 테이블, 슬라이서를 지원하지 않는다. 편집은 openpyxl, 피벗 대용은 SUMIFS+드롭다운.
- LibreOffice 렌더링은 스파크라인을 일부 못 그린다. 엑셀에서는 정상이므로 렌더 이미지만 믿지 말고 구조를 확인한다.
