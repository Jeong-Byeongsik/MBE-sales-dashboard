"""verify_xlsx: 만든 엑셀을 LibreOffice로 재계산·검증하고 PNG로 렌더링합니다.

  python verify_xlsx.py 파일.xlsx            # 수식 오류 검사 + 요약
  python verify_xlsx.py 파일.xlsx --render   # 시트별 PNG 미리보기도 생성(./preview_*.png)
라이브러리로도 사용: recalc_values(path) -> {(시트, 셀): 값} (2단계 빌드용), render(path, outdir)
"""
import os, shutil, subprocess, sys, tempfile, glob


_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def _profile():
    """xlsxwriter 파일은 LibreOffice가 '엑셀이 만든 파일'로 보고 재계산을 건너뜁니다 -> 항상 재계산하도록 설정."""
    prof = os.path.join(tempfile.gettempdir(), 'lo_profile_xl_recalc')
    user = os.path.join(prof, 'user')
    os.makedirs(user, exist_ok=True)
    xcu = os.path.join(user, 'registrymodifications.xcu')
    if not os.path.exists(xcu):
        open(xcu, 'w', encoding='utf8').write(_XCU)
    return prof


def _soffice(args, timeout=180):
    prof = _profile()
    cmd = ['soffice', f'-env:UserInstallation=file://{prof}', '--headless', '--norestore'] + args
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)


def recalc_copy(path):
    """사본을 LibreOffice로 재저장(=재계산)한 경로를 돌려줍니다. 원본은 그대로."""
    d = tempfile.mkdtemp(prefix='xl_recalc_')
    src = os.path.join(d, 'in.xlsx')
    shutil.copy(path, src)
    out = os.path.join(d, 'out')
    os.makedirs(out)
    _soffice(['--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', out, src])
    res = os.path.join(out, 'in.xlsx')
    if not os.path.exists(res):
        raise RuntimeError('LibreOffice 재계산 실패')
    return res


def recalc_values(path, cells=None):
    """수식 칸의 재계산 값을 {(시트,셀): 값} 으로. cells 를 주면 그 칸만."""
    from openpyxl import load_workbook
    wb = load_workbook(recalc_copy(path), data_only=True)
    out = {}
    if cells:
        for sh, ref in cells:
            out[(sh, ref)] = wb[sh][ref].value
    else:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        out[(ws.title, c.coordinate)] = c.value
    return out


def check(path):
    from openpyxl import load_workbook
    vals = load_workbook(recalc_copy(path), data_only=True)
    forms = load_workbook(path)
    errors, nform = [], 0
    for ws in forms.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    nform += 1
                    v = vals[ws.title][c.coordinate].value
                    if isinstance(v, str) and v.startswith('#'):
                        errors.append((ws.title, c.coordinate, v, c.value[:60]))
    return {'sheets': forms.sheetnames, 'formulas': nform, 'errors': errors}


def render(path, outdir='.', dpi=70):
    """PDF로 변환 후 페이지별 PNG 생성. 파일 경로 목록을 돌려줌."""
    os.makedirs(outdir, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix='xl_pdf_')
    _soffice(['--convert-to', 'pdf', '--outdir', tmp, path])
    pdf = os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + '.pdf')
    if not os.path.exists(pdf):
        raise RuntimeError('PDF 변환 실패')
    pre = os.path.join(outdir, 'preview')
    subprocess.run(['pdftoppm', '-r', str(dpi), '-png', pdf, pre], check=True)
    return sorted(glob.glob(pre + '*.png'))


if __name__ == '__main__':
    p = sys.argv[1]
    r = check(p)
    print(f"시트 {r['sheets']}\n수식 {r['formulas']}개, 오류 {len(r['errors'])}개")
    for e in r['errors'][:20]:
        print('  ', e)
    if '--render' in sys.argv:
        print(render(p, os.path.dirname(os.path.abspath(p)) or '.'))
    sys.exit(1 if r['errors'] else 0)
