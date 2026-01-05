# app.py
# -*- coding: utf-8 -*-

import re
from io import BytesIO
from collections import defaultdict, Counter
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# =============================
# 格式识别（看文件头，不看后缀）
# =============================
def detect_excel_format_from_bytes(data: bytes) -> str:
    head = data[:8]
    # xls (OLE2/CFB)
    if head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
        return "xls"
    # xlsx (ZIP)
    if head.startswith(b"PK\x03\x04") or head.startswith(b"PK\x05\x06") or head.startswith(b"PK\x07\x08"):
        return "xlsx"
    raise RuntimeError("无法识别文件格式：既不是 OLE2 xls，也不是 ZIP xlsx")


# =============================
# 工具函数
# =============================
def norm_str(x: Any) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def money_to_cents(x: Any) -> int:
    if x is None:
        return 0
    if isinstance(x, int):
        return x * 100
    if isinstance(x, float):
        return int(round(x * 100))
    s = str(x).strip().replace(",", "")
    if s == "":
        return 0
    try:
        return int(round(float(s) * 100))
    except ValueError:
        return 0


def cents_to_money_str(cents: int) -> str:
    sign = "-" if cents < 0 else ""
    cents = abs(cents)
    return f"{sign}{cents // 100}.{cents % 100:02d}"


def find_header_row_in_rows(rows: List[List[Any]], required_cols: List[str]) -> Optional[Tuple[int, Dict[str, int]]]:
    required = set(required_cols)
    for r, row in enumerate(rows):
        row_vals = [norm_str(v) for v in row]
        row_set = set([v for v in row_vals if v])
        if required.issubset(row_set):
            col_map = {name: row_vals.index(name) for name in required_cols}
            return r, col_map
    return None


# =============================
# 统一读取：从 bytes 得到“数据行迭代器”
# =============================
def open_excel_iter_from_bytes(data: bytes, required_cols: List[str], scan_rows: int = 80):
    fmt = detect_excel_format_from_bytes(data)

    if fmt == "xls":
        import xlrd
        book = xlrd.open_workbook(file_contents=data, formatting_info=False)
        sh = book.sheet_by_index(0)

        max_r = min(scan_rows, sh.nrows)
        sample = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(max_r)]
        found = find_header_row_in_rows(sample, required_cols)
        if not found:
            raise RuntimeError(f"未在前 {scan_rows} 行找到表头（xls）：{required_cols}")
        header_r, cmap = found

        def row_iter():
            for r in range(header_r + 1, sh.nrows):
                yield [sh.cell_value(r, c) for c in range(sh.ncols)]

        return cmap, row_iter(), fmt

    # xlsx
    import openpyxl
    wb = openpyxl.load_workbook(BytesIO(data), read_only=True, data_only=True)
    ws = wb.worksheets[0]

    sample = []
    for row in ws.iter_rows(min_row=1, max_row=scan_rows, values_only=True):
        sample.append(list(row))
    found = find_header_row_in_rows(sample, required_cols)
    if not found:
        wb.close()
        raise RuntimeError(f"未在前 {scan_rows} 行找到表头（xlsx）：{required_cols}")
    header_r, cmap = found  # 0-based

    def row_iter():
        for row in ws.iter_rows(min_row=header_r + 2, values_only=True):
            yield list(row)

    return cmap, row_iter(), fmt


# =============================
# 聚合结构
# =============================
@dataclass
class Agg:
    cents_sum: int = 0
    row_count: int = 0
    meta_counter: Counter = None

    def __post_init__(self):
        if self.meta_counter is None:
            self.meta_counter = Counter()


def agg_add(a: Agg, cents: int, meta: Optional[str] = None):
    a.cents_sum += cents
    a.row_count += 1
    if meta:
        a.meta_counter[meta] += 1


# =============================
# 读取并聚合：支付表 / 菜品表（缓存）
# =============================
@st.cache_data(show_spinner=False)
def read_payment_agg(file_bytes: bytes) -> Tuple[Dict[str, Agg], Dict[str, Agg], str]:
    required = ["POS销售单号", "POS退款单号", "总金额", "支付类型", "门店"]
    cmap, rows, fmt = open_excel_iter_from_bytes(file_bytes, required_cols=required)

    sales: Dict[str, Agg] = defaultdict(Agg)
    refund: Dict[str, Agg] = defaultdict(Agg)

    for row in rows:
        def v(col):
            idx = cmap[col]
            return row[idx] if idx < len(row) else None

        pos_sale = norm_str(v("POS销售单号"))
        pos_refund = norm_str(v("POS退款单号"))
        cents = money_to_cents(v("总金额"))
        pay_type = norm_str(v("支付类型"))
        store = norm_str(v("门店"))

        if not pos_sale and not pos_refund:
            continue

        meta = f"{store}|{pay_type}" if (store or pay_type) else None
        if pos_refund:
            agg_add(refund[pos_refund], cents, meta)
        else:
            agg_add(sales[pos_sale], cents, meta)

    return sales, refund, fmt


@st.cache_data(show_spinner=False)
def read_items_agg(file_bytes: bytes) -> Tuple[Dict[str, Agg], Dict[str, Agg], str]:
    required = ["POS销售单号", "POS退款单号", "优惠后小计价格", "单据类型", "菜品状态", "门店", "菜品名称"]
    cmap, rows, fmt = open_excel_iter_from_bytes(file_bytes, required_cols=required)

    sales: Dict[str, Agg] = defaultdict(Agg)
    refund: Dict[str, Agg] = defaultdict(Agg)

    for row in rows:
        def v(col):
            idx = cmap[col]
            return row[idx] if idx < len(row) else None

        pos_sale = norm_str(v("POS销售单号"))
        pos_refund = norm_str(v("POS退款单号"))
        cents = money_to_cents(v("优惠后小计价格"))

        doc_type = norm_str(v("单据类型"))
        dish_status = norm_str(v("菜品状态"))
        store = norm_str(v("门店"))
        dish_name = norm_str(v("菜品名称"))

        if not pos_sale and not pos_refund:
            continue

        meta = f"{store}|{doc_type}|{dish_status}|{dish_name}" if (store or doc_type or dish_status or dish_name) else None
        if pos_refund:
            agg_add(refund[pos_refund], cents, meta)
        else:
            agg_add(sales[pos_sale], cents, meta)

    return sales, refund, fmt


# =============================
# 对账
# =============================
def compare_aggs(pay: Dict[str, Agg], item: Dict[str, Agg], kind: str, tolerance_cents: int) -> pd.DataFrame:
    keys = set(pay.keys()) | set(item.keys())
    out = []

    for k in keys:
        pa = pay.get(k)
        ia = item.get(k)

        ps = pa.cents_sum if pa else 0
        is_ = ia.cents_sum if ia else 0
        diff = ps - is_

        if abs(diff) <= tolerance_cents:
            continue

        if pa is None:
            reason = "仅菜品存在：支付缺失/未导出/单号不一致"
        elif ia is None:
            reason = "仅支付存在：菜品缺失/未导出/单号不一致"
        else:
            if kind == "REFUND":
                if (ps > 0 and is_ < 0) or (ps < 0 and is_ > 0):
                    reason = "退款正负号方向不一致（一个表正一个表负）"
                else:
                    reason = "退款金额不一致：可能部分退款/多次退款/抹零/口径差异"
            else:
                if pa.row_count > 1:
                    reason = "支付多笔汇总仍不一致：可能混合支付/重复/拆单"
                elif ia.row_count > 20:
                    reason = "菜品行很多仍不一致：可能退菜/作废/状态口径差异"
                else:
                    reason = "金额不一致：可能抹零/服务费/包装费/非菜品费用/口径差异"

        pay_meta = "; ".join([f"{m}*{c}" for m, c in (pa.meta_counter.most_common(3) if pa else [])])
        item_meta = "; ".join([f"{m}*{c}" for m, c in (ia.meta_counter.most_common(3) if ia else [])])

        out.append({
            "单据类型（销售/退款）": "销售" if kind == "SALE" else "退款",
            "POS单号": k,
            "支付表总金额": cents_to_money_str(ps),
            "菜品表优惠后金额": cents_to_money_str(is_),
            "金额差异（支付-菜品）": cents_to_money_str(diff),
            "支付记录行数": pa.row_count if pa else 0,
            "菜品记录行数": ia.row_count if ia else 0,
            "差异原因判断": reason,
            "支付侧关键信息（门店|支付类型 Top3）": pay_meta,
            "菜品侧关键信息（门店|单据类型|状态|菜品 Top3）": item_meta,
        })

    df = pd.DataFrame(out)
    if not df.empty:
        df["_abs"] = df["金额差异（支付-菜品）"].apply(money_to_cents).abs()
        df = df.sort_values("_abs", ascending=False).drop(columns=["_abs"])
    return df


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# =============================
# Streamlit 页面
# =============================
st.set_page_config(page_title="POS 对账：支付 vs 菜品", layout="wide")
st.title("POS 对账：订单支付报告 vs 订单菜品报告")

st.markdown(
    """
- 上传 **两个文件**：订单支付报告、订单菜品报告（支持 `.xls` / `.xlsx` / “真 xls 假 xlsx”）。
- **销售单**按 `POS销售单号` 对账；**退款单**按 `POS退款单号` 对账（只要非空就走退款逻辑）。
- 支付表金额：`总金额`；菜品表金额：`优惠后小计价格`（同单号汇总）。
"""
)

col1, col2 = st.columns(2)
with col1:
    pay_file = st.file_uploader("上传【订单支付报告】", type=["xls", "xlsx"])
with col2:
    item_file = st.file_uploader("上传【订单菜品报告】", type=["xls", "xlsx"])

tolerance_yuan = st.slider("容差（元）", min_value=0.00, max_value=1.00, value=0.10, step=0.01)
tolerance_cents = int(round(tolerance_yuan * 100))
preview_rows = st.number_input("差异明细预览行数", min_value=50, max_value=5000, value=500, step=50)

run_btn = st.button("开始对账", type="primary", disabled=(pay_file is None or item_file is None))

if run_btn:
    try:
        pay_bytes = pay_file.getvalue()
        item_bytes = item_file.getvalue()

        with st.spinner("读取并聚合中（大文件可能需要一会儿）…"):
            pay_sales, pay_refund, pay_fmt = read_payment_agg(pay_bytes)
            item_sales, item_refund, item_fmt = read_items_agg(item_bytes)

        st.success(f"文件识别：支付表={pay_fmt}，菜品表={item_fmt}；容差={tolerance_yuan:.2f} 元")

        with st.spinner("对账计算中…"):
            df_sale = compare_aggs(pay_sales, item_sales, kind="SALE", tolerance_cents=tolerance_cents)
            df_refund = compare_aggs(pay_refund, item_refund, kind="REFUND", tolerance_cents=tolerance_cents)
            df_all = pd.concat([df_sale, df_refund], ignore_index=True)

        if df_all.empty:
            st.info("✅ 未发现差异（在容差范围内）")
            stat = pd.DataFrame(columns=["差异原因判断", "数量"])
        else:
            stat = (df_all["差异原因判断"]
                    .value_counts()
                    .reset_index()
                    .rename(columns={"index": "差异原因判断", "差异原因判断": "数量"}))

        left, right = st.columns([1, 2])

        with left:
            st.subheader("差异原因统计")
            st.dataframe(stat, use_container_width=True, height=360)
            st.download_button(
                "下载：差异原因统计.csv",
                data=df_to_csv_bytes(stat),
                file_name="差异原因统计.csv",
                mime="text/csv",
                use_container_width=True
            )

        with right:
            st.subheader(f"差异明细（共 {len(df_all)} 条）")
            st.dataframe(df_all.head(int(preview_rows)), use_container_width=True, height=520)
            st.download_button(
                "下载：差异明细.csv",
                data=df_to_csv_bytes(df_all),
                file_name="差异明细.csv",
                mime="text/csv",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"运行失败：{e}")
        st.exception(e)
