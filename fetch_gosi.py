# 토지이음 데이터개방의 월간 고시정보 zip을 받아 정비사업 고시만 골라 gosi.json 으로 저장한다.
import datetime
import io
import json
import re
import ssl
import urllib.parse
import urllib.request
import zipfile

import openpyxl

UA = {"User-Agent": "Mozilla/5.0"}
CTX = ssl._create_unverified_context()  # 정부 누리집 인증서 체인이 표준 저장소에 없다
LIST_URL = "https://eum.go.kr/web/op/sv/svItemDet.jsp"
AJAX_URL = "https://eum.go.kr/web/op/sv/svItemAjaxXml.jsp"
DOWN_URL = "https://map.eum.go.kr:8002/OpenData/opDownloader.jsp"
REDEV = re.compile(r"정비구역|정비계획|재개발|재건축|사업시행계획|관리처분|정비사업")
KEEP_DAYS = 730


def call(url, data=None, timeout=60):
    body = urllib.parse.urlencode(data).encode() if data else None
    return urllib.request.urlopen(urllib.request.Request(url, data=body, headers=UA), timeout=timeout, context=CTX).read()


def main():
    page = call(LIST_URL, {"dataCd": "008", "dataTypeCd": "CSV", "currentPageNo": "1", "selectType": "subject", "keyword": ""}).decode("cp949", "ignore")
    file_id = re.search(r"dataDownload\('(\d+)'\)", page).group(1)
    info = call(AJAX_URL, {"function": "selectFileInfo", "dataCd": "008", "dataTypeCd": "CSV", "refDt": "", "fileId": file_id}).decode("cp949", "ignore")
    key, name, ref = (re.search(rf"<{t}>([^<]+)", info).group(1) for t in ("key", "fileNm", "refDt"))
    # 사이트가 내려받기 전에 묻는 이용 목적을 같은 방식으로 남긴다(통계용이라 실패해도 진행)
    try:
        call(AJAX_URL, {"function": "insertStatInfo", "dataCd": "008", "dataTypeCd": "CSV", "useType": "01", "useDesc": "재개발·재건축 사업 현황 관리"})
    except Exception:
        pass
    data = call(f"{DOWN_URL}?key={key}&filename={name}", timeout=900)
    if not data.startswith(b"PK"):
        raise SystemExit("zip 아님: " + data[:200].decode("cp949", "ignore"))
    z = zipfile.ZipFile(io.BytesIO(data))
    ws = openpyxl.load_workbook(io.BytesIO(z.read(z.namelist()[0])), read_only=True)["고시"]
    ws.reset_dimensions()
    rows_iter = ws.iter_rows(values_only=True)
    next(rows_iter)
    cutoff = (datetime.date.today() - datetime.timedelta(days=KEEP_DAYS)).isoformat()
    rows = []
    for no, date, url, title in rows_iter:
        d, t = str(date or "")[:10], str(title or "")
        if d >= cutoff and REDEV.search(t):
            rows.append({"no": str(no or ""), "date": d, "org": re.sub(r"\s*(고시|공고).*$", "", str(no or "")).strip(), "url": str(url or ""), "title": t})
    rows.sort(key=lambda r: r["date"], reverse=True)
    out = {"file": name, "refDt": ref, "fetchedAt": datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M"), "rows": rows}
    with open("gosi.json", "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    print(f"{name} 기준일 {ref} · zip {len(data) / 1048576:.1f}MB · 정비사업 고시 {len(rows)}건({cutoff} 이후)")


if __name__ == "__main__":
    main()
