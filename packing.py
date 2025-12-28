from docx import Document
import json
import re
from datetime import date

def norm(s):
    return str(s).strip()

def extract_ship_id(doc):
    """
    문서 전체에서 '출고주문번호:' 다음 값을 찾음
    """
    pattern = re.compile(r"출고주문번호[:\s]*([A-Z0-9\-]+)")
    for p in doc.paragraphs:
        m = pattern.search(p.text)
        if m:
            return m.group(1)
    return None

def main(docx_path, out_path="orders.json"):
    doc = Document(docx_path)

    ship_id = extract_ship_id(doc)
    if not ship_id:
        raise Exception("❌ 출고주문번호를 찾을 수 없습니다.")

    orders = {
        "orderId": ship_id,
        "items": []
    }

    # 모든 표 탐색
    for table in doc.tables:
        # 헤더 위치 찾기
        headers = [cell.text.strip() for cell in table.rows[0].cells]

        try:
            idx_name = headers.index("주문상품")
            idx_qty = headers.index("주문수량")
            idx_barcode = headers.index("상품 바코드")
        except ValueError:
            continue  # 이 표는 우리가 원하는 표가 아님

        # 데이터 행
        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]

            name = norm(cells[idx_name])
            qty = int(cells[idx_qty])
            barcode = norm(cells[idx_barcode])

            if not barcode or qty <= 0:
                continue

            orders["items"].append({
                "barcode": barcode,
                "name": name,
                "qty": qty
            })

    out = {
        "date": str(date.today()),
        "orders": [orders]
    }

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"✅ 변환 완료: {out_path}")
    print(f"출고번호: {ship_id}, 상품 {len(orders['items'])}건")

if __name__ == "__main__":
    main("input.docx", "orders.json")
