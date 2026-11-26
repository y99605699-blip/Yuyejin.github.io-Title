import json, io
import numpy as np
import openpyxl
book = None

def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)

def initialize(raw):
    global book, codes, names, L, M, V, sheets, valid
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)
    base = book['국산투입계수(Ad)']
    codes = [str(c) for c in base[4][2:] if c is not None and str(c).isdigit() and int(str(c)) < 9000]
    if len(codes) < 2 or len(set(codes)) != len(codes):
        raise ValueError('부문 코드를 읽지 못했습니다. 엑셀 5행(헤더) 구조를 확인하세요.')
    nm = {str(r[0]): str(r[1]) for r in base[6:]}
    names = [nm[c] for c in codes]
    def matrix(s):
        rows = book[s]
        ri = {str(r[0]): i for i,r in enumerate(rows) if r}
        ci = {str(c): i for i,c in enumerate(rows[4])}
        a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        if not np.isfinite(a).all(): raise ValueError(s + ': 비수치 또는 결측 데이터')
        return a
    L = matrix('생산유발계수'); M = matrix('수입유발계수'); V = matrix('부가가치유발계수')
    ad = matrix('국산투입계수(Ad)'); am = matrix('수입투입계수(Am)')
    rows = book['총투입계수(A)']
    ci = {str(c): i for i, c in enumerate(rows[4])}
    vr = next((r for r in rows if r and str(r[0]) == '9690'), None)
    if vr is not None:
        v = np.array([vr[ci[c]] for c in codes], dtype=float)
    else:
        ri2 = {str(r[0]): i for i, r in enumerate(rows) if r}
        A = np.array([[rows[ri2[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        v = 1 - A.sum(axis=0)
    residuals = {'생산':float(np.max(np.abs((np.eye(len(codes))-ad)@L-np.eye(len(codes))))),
                 '수입':float(np.max(np.abs(am@L-M))),
                 '부가가치':float(np.max(np.abs(v[:,None]*L-V)))}
    valid = max(residuals.values()) < 1e-8
    if not valid: raise ValueError('계수 검증 실패: '+str(residuals))
    return emit({'sectors':[{'code':c,'name':n} for c,n in zip(codes,names)],
                 'sheets':sheets,'residuals':residuals})

def explore(sheet, rowcode):
    rows=book[sheet]
    choices=[{'code':str(r[0]),'name':str(r[1])} for r in rows[6:] if len(r)>1 and r[0] is
