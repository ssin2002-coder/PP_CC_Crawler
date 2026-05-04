"""Word에서 열린 문서의 실제 셀 내용을 진단하는 스크립트"""
import pythoncom
import win32com.client
from word_crawler import clean_cell_text, split_items, find_main_table_index

pythoncom.CoInitialize()

try:
    word = win32com.client.GetActiveObject('Word.Application')
except Exception:
    print("Word가 실행되지 않았거나 문서가 열려있지 않습니다.")
    exit()

for i in range(1, word.Documents.Count + 1):
    doc = word.Documents(i)
    print(f"\n{'='*60}")
    print(f"문서: {doc.Name}")
    print(f"표 개수: {doc.Tables.Count}")

    # 각 표 행 수
    row_counts = []
    for t in range(1, doc.Tables.Count + 1):
        try:
            rc = doc.Tables(t).Rows.Count
            row_counts.append(rc)
            print(f"  표 {t}: {rc}행")
        except:
            row_counts.append(0)

    main_idx = find_main_table_index(row_counts)
    print(f"\n선택된 메인 표: {main_idx + 1 if main_idx is not None else 'None'}")

    if main_idx is None:
        continue

    table = doc.Tables(main_idx + 1)

    # 헤더
    row1 = table.Rows(1)
    headers = []
    for c in range(1, row1.Cells.Count + 1):
        h_raw = row1.Cells(c).Range.Text
        h_clean = clean_cell_text(h_raw)
        headers.append(h_clean)
    print(f"헤더: {headers}")

    # 첫 2개 데이터 행 상세 출력
    for r in range(2, min(4, table.Rows.Count + 1)):
        print(f"\n--- 행 {r} ---")
        row = table.Rows(r)
        for c in range(1, row.Cells.Count + 1):
            raw = row.Cells(c).Range.Text
            cleaned = clean_cell_text(raw)
            items = split_items(cleaned)

            print(f"  [열 {c}] 헤더: {headers[c-1] if c-1 < len(headers) else '?'}")
            print(f"    raw repr: {repr(raw[:200])}")
            print(f"    cleaned repr: {repr(cleaned[:200])}")
            print(f"    split_items 결과 ({len(items)}건):")
            for idx, item in enumerate(items):
                print(f"      {idx+1}) {item[:100]}")

pythoncom.CoUninitialize()
