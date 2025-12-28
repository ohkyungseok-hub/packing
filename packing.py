import io
import json
import re
from datetime import date
from typing import Dict, List, Optional, Tuple

import streamlit as st
from docx import Document


def norm(s: str) -> str:
    return str(s).strip()


SHIP_RE = re.compile(r"출고주문번호[:\s]*([A-Za-z0-9\-]+)")


def find_ship_ids_in_order(doc: Document) -> List[str]:
    """문서 전체에서 출고주문번호를 등장 순서대로 수집(중복 제거)."""
    ids: List[str] = []
    for p in doc.paragraphs:
        m = SHIP_RE.search(p.text)
        if m:
            ship = norm(m.group(1))
            if ship and (not ids or ids[-1] != ship):
                ids.append(ship)
    return ids


def table_headers(table) -> List[str]:
    if not table.rows:
        return []
    return [norm(c.text) for c in table.rows[0].cells]


def parse_items_from_table(table) -> Tuple[List[Dict], List[str]]:
    """
    한 개 표에서 items 추출.
    - name: '주문상품' 우선
    - qty: '주문수량'
    - barcode: '상품 바코드'가 있으면 우선, 없으면 '상품연동코드' 사용
    """
    warnings: List[str] = []
    headers = table_headers(table)
    if not headers:
        return [], warnings

    # 필수
    if "주문상품" not in headers or "주문수량" not in headers:
        return [], warnings

    idx_name = headers.index("주문상품")
    idx_qty = headers.index("주문수량")

    # 바코드 컬럼 후보들
    barcode_idx = None
    if "상품 바코드" in headers:
        barcode_idx = headers.index("상품 바코드")
    elif "상품연동코드" in headers:
        barcode_idx = headers.index("상품연동코드")
        warnings.append("표에 '상품 바코드' 컬럼이 없어 '상품연동코드'를 barcode로 사용합니다.")

    items: List[Dict] = []
    for r_i, row in enumerate(table.rows[1:], start=2):
        cells = [norm(c.text) for c in row.cells]
        if len(cells) <= max(idx_name, idx_qty, (barcode_idx or 0)):
            continue

        name = norm(cells[idx_name])
        qty_raw = norm(cells[idx_qty])
        m = re.search(r"\d+", qty_raw)
        qty = int(m.group(0)) if m else 0

        # 합계 행 스킵
        if "합계" in name:
            continue

        if qty <= 0:
            continue

        barcode = ""
        if barcode_idx is not None:
            barcode = norm(cells[barcode_idx])

        # barcode가 비면 검증 불가라 제외(필요하면 포함하도록 바꿀 수 있음)
        if not barcode:
            warnings.append(f"표 {r_i}행: barcode가 비어 있어 제외됨 (상품='{name}')")
            continue

        items.append({"barcode": barcode, "name": name or "(상품명 없음)", "qty": qty})

    return items, warnings


def build_orders_json_multi(doc: Document) -> Tuple[Dict, List[str]]:
    warnings: List[str] = []
    ship_ids = find_ship_ids_in_order(doc)

    if not ship_ids:
        warnings.append("문서에서 '출고주문번호'를 하나도 찾지 못했습니다.")
        ship_ids = ["UNKNOWN"]

    # 전략: 문서에 있는 표들을 순서대로 훑어서, "출고번호 수"만큼 표를 매칭해본다.
    # (대부분 출고번호 1개당 표 1개 구조라서 잘 맞습니다.)
    tables = list(doc.tables)

    orders: List[Dict] = []
    table_cursor = 0

    for ship in ship_ids:
        items: List[Dict] = []
        local_warn: List[str] = []

        # 다음에 나오는 "유효한(주문상품/주문수량 포함)" 표를 찾는다
        while table_cursor < len(tables):
            t = tables[table_cursor]
            table_cursor += 1

            it, w = parse_items_from_table(t)
            if it:
                items = it
                local_warn.extend(w)
                break
            else:
                # items가 0이어도 headers 경고는 굳이 쌓지 않음
                continue

        if not items:
            local_warn.append(f"{ship}: 매칭되는 상품 표를 찾지 못했거나, 표에서 상품을 추출하지 못했습니다.")

        orders.append({"orderId": ship, "items": items})
        warnings.extend(local_warn)

    out = {"date": str(date.today()), "orders": orders}
    return out, warnings


st.set_page_config(page_title="DOCX → orders.json 변환", layout="centered")
st.title("워드(.docx) → orders.json 변환기 (멀티 출고 지원)")

uploaded = st.file_uploader("워드(.docx) 업로드", type=["docx"])

if uploaded is not None:
    try:
