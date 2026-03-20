import io
import re
import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# =========================
# Config
# =========================
st.set_page_config(page_title="Bank Statement Expense Classifier", layout="wide")

APP_DIR = Path(".")
RULES_DIR = APP_DIR / "rules_data"
RULES_DIR.mkdir(exist_ok=True)

MERCHANT_RULES_FILE = RULES_DIR / "merchant_rules.csv"
KEYWORD_RULES_FILE = RULES_DIR / "keyword_rules.csv"
OVERRIDE_RULES_FILE = RULES_DIR / "manual_overrides.csv"

CATEGORY_OPTIONS = [
    "广告费（TK Ads）",
    "物流成本",
    "穿戴甲进货成本",
    "配套耗材进货成本",
    "停车补助费",
    "午餐补助费",
    "房租成本",
    "行政费用",
    "办公用品采购",
    "办公软件会员费用",
    "销售税预提（CA）",
    "银行手续费",
    "USPS 水单面单费用",
    "员工工资",
    "拍摄费用",
    "主播工资",
    "国内转账",
    "Shopify 扣款",
    "ADP TAX",
    "律师费用",
    "收入/回款",
    "转账/内部往来",
    "其他待确认",
]

DEFAULT_MERCHANT_RULES = [
    ("TIKTOK ADS", "广告费（TK Ads）", 0.99, "merchant"),
    ("ADS.TIKTOK", "广告费（TK Ads）", 0.99, "merchant"),
    ("TIKTOK SHOP", "Shopify 扣款", 0.70, "merchant"),
    ("SHOPIFY", "Shopify 扣款", 0.95, "merchant"),
    ("ADP TAX", "ADP TAX", 0.99, "merchant"),
    ("ADP WAGE PAY", "员工工资", 0.99, "merchant"),
    ("ADP FEES", "办公软件会员费用", 0.90, "merchant"),
    ("ADP PAY-BY-PAY", "办公软件会员费用", 0.90, "merchant"),
    ("REGUS", "房租成本", 0.98, "merchant"),
    ("HANKIN PATENT LAW", "律师费用", 0.99, "merchant"),
    ("OPENAI", "办公软件会员费用", 0.98, "merchant"),
    ("CHATGPT SUBSCR", "办公软件会员费用", 0.98, "merchant"),
    ("INTUIT", "办公软件会员费用", 0.98, "merchant"),
    ("QBOOKS", "办公软件会员费用", 0.98, "merchant"),
    ("GOOGLE *WORKSP", "办公软件会员费用", 0.95, "merchant"),
    ("GOOGLE WORKSPACE", "办公软件会员费用", 0.95, "merchant"),
    ("USPS", "USPS 水单面单费用", 0.95, "merchant"),
    ("PITNEY BOWES", "USPS 水单面单费用", 0.95, "merchant"),
    ("MARS SHIPPING", "物流成本", 0.95, "merchant"),
    ("SHANTOU PANJIA", "配套耗材进货成本", 0.80, "merchant"),
    ("ALADDIN GLOBAL", "穿戴甲进货成本", 0.90, "merchant"),
    ("H.K ALADDIN GLOBAL", "穿戴甲进货成本", 0.90, "merchant"),
    ("DBS BANK", "穿戴甲进货成本", 0.75, "merchant"),
    ("BANK OF AMERICA", "银行手续费", 0.80, "merchant"),
]

DEFAULT_KEYWORD_RULES = [
    ("PAYROLL", "员工工资", 0.95, "keyword"),
    ("LUNCH", "午餐补助费", 0.98, "keyword"),
    ("PARKING", "停车补助费", 0.98, "keyword"),
    ("INTERNALTRANSFER", "转账/内部往来", 0.98, "keyword"),
    ("INTERNAL TRANSFER", "转账/内部往来", 0.98, "keyword"),
    ("REFUND", "收入/回款", 0.70, "keyword"),
    ("WIRE TYPE:FX OUT", "穿戴甲进货成本", 0.85, "keyword"),
    ("EXTERNAL TRANSFER FEE", "银行手续费", 0.98, "keyword"),
    ("SERVICE FEE", "银行手续费", 0.98, "keyword"),
]

LIKELY_AMBIGUOUS_KEYWORDS = [
    "AMAZON", "SQ *", "UBER", "FANTUAN", "HUNGRYPANDA", "PHO", "MEAL", "RESTAURANT"
]

PDF_ENGINE_HELP = """
推荐安装：
pip install streamlit pandas pdfplumber openpyxl xlsxwriter pypdf
"""

# =========================
# Utility + persistence
# =========================
def ensure_rule_files() -> None:
    if not MERCHANT_RULES_FILE.exists():
        pd.DataFrame(DEFAULT_MERCHANT_RULES, columns=["pattern", "category", "confidence", "rule_type"]).to_csv(
            MERCHANT_RULES_FILE, index=False
        )
    if not KEYWORD_RULES_FILE.exists():
        pd.DataFrame(DEFAULT_KEYWORD_RULES, columns=["pattern", "category", "confidence", "rule_type"]).to_csv(
            KEYWORD_RULES_FILE, index=False
        )
    if not OVERRIDE_RULES_FILE.exists():
        pd.DataFrame(columns=["merchant_key", "category", "created_at"]).to_csv(OVERRIDE_RULES_FILE, index=False)


def load_rules() -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    ensure_rule_files()
    merchant_df = pd.read_csv(MERCHANT_RULES_FILE)
    keyword_df = pd.read_csv(KEYWORD_RULES_FILE)
    override_df = pd.read_csv(OVERRIDE_RULES_FILE)
    return merchant_df, keyword_df, override_df


def save_override_rule(merchant_key: str, category: str) -> None:
    _, _, override_df = load_rules()
    merchant_key = normalize_text(merchant_key)
    override_df = override_df[override_df["merchant_key"].astype(str) != merchant_key]
    new_row = pd.DataFrame(
        [{"merchant_key": merchant_key, "category": category, "created_at": datetime.now().isoformat(timespec="seconds")}]
    )
    override_df = pd.concat([override_df, new_row], ignore_index=True)
    override_df.to_csv(OVERRIDE_RULES_FILE, index=False)


def save_new_pattern_rule(pattern: str, category: str, rule_bucket: str = "merchant", confidence: float = 0.95) -> None:
    pattern = normalize_text(pattern)
    if rule_bucket == "merchant":
        df = pd.read_csv(MERCHANT_RULES_FILE)
        if not ((df["pattern"].astype(str) == pattern) & (df["category"].astype(str) == category)).any():
            df = pd.concat(
                [df, pd.DataFrame([{"pattern": pattern, "category": category, "confidence": confidence, "rule_type": "merchant"}])],
                ignore_index=True,
            )
            df.to_csv(MERCHANT_RULES_FILE, index=False)
    else:
        df = pd.read_csv(KEYWORD_RULES_FILE)
        if not ((df["pattern"].astype(str) == pattern) & (df["category"].astype(str) == category)).any():
            df = pd.concat(
                [df, pd.DataFrame([{"pattern": pattern, "category": category, "confidence": confidence, "rule_type": "keyword"}])],
                ignore_index=True,
            )
            df.to_csv(KEYWORD_RULES_FILE, index=False)


def normalize_text(text: str) -> str:
    text = str(text or "")
    text = text.replace("’", "'").replace("“", '"').replace("”", '"')
    text = text.upper().strip()
    text = re.sub(r"\s+", " ", text)
    return text


def safe_float(x) -> float:
    try:
        return float(str(x).replace(",", "").replace("$", "").strip())
    except Exception:
        return 0.0


def amount_to_display(x: float) -> str:
    return f"${x:,.2f}"


def to_excel_bytes(detail_df: pd.DataFrame, summary_df: pd.DataFrame, unknown_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        detail_df.to_excel(writer, index=False, sheet_name="交易明细")
        summary_df.to_excel(writer, index=False, sheet_name="分类汇总")
        unknown_df.to_excel(writer, index=False, sheet_name="待确认交易")
    output.seek(0)
    return output.getvalue()


# =========================
# PDF / CSV parsing
# =========================
def parse_uploaded_file(uploaded_file) -> pd.DataFrame:
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".csv":
        return parse_csv(uploaded_file)
    if suffix in [".xlsx", ".xls"]:
        return parse_excel(uploaded_file)
    if suffix == ".pdf":
        return parse_pdf_statement(uploaded_file)

    raise ValueError("暂不支持该文件类型，请上传 PDF / CSV / Excel。")


def parse_csv(uploaded_file) -> pd.DataFrame:
    df = pd.read_csv(uploaded_file)
    return standardize_transaction_df(df)


def parse_excel(uploaded_file) -> pd.DataFrame:
    xls = pd.ExcelFile(uploaded_file)
    df = pd.read_excel(xls, xls.sheet_names[0])
    return standardize_transaction_df(df)


def standardize_transaction_df(df: pd.DataFrame) -> pd.DataFrame:
    raw_cols = {c.lower().strip(): c for c in df.columns}
    date_col = next((raw_cols[c] for c in raw_cols if "date" in c), None)
    desc_col = next((raw_cols[c] for c in raw_cols if any(k in c for k in ["description", "memo", "details", "merchant", "name"])), None)
    amount_col = next((raw_cols[c] for c in raw_cols if "amount" in c), None)

    if not all([date_col, desc_col, amount_col]):
        raise ValueError("CSV / Excel 未识别到 date / description / amount 列，请检查表头。")

    result = pd.DataFrame({
        "date": pd.to_datetime(df[date_col], errors="coerce"),
        "description": df[desc_col].astype(str),
        "amount": pd.to_numeric(df[amount_col], errors="coerce"),
    })

    result = result.dropna(subset=["date", "description", "amount"]).copy()
    result["raw_description"] = result["description"]
    result["direction"] = result["amount"].apply(lambda x: "debit" if x < 0 else "credit")
    result["amount_abs"] = result["amount"].abs()
    return result.reset_index(drop=True)


def parse_pdf_statement(uploaded_file) -> pd.DataFrame:
    text = extract_pdf_text(uploaded_file)
    txns = extract_boa_transactions_from_text(text)
    if not txns:
        raise ValueError("未能从 PDF 中识别出交易明细。请先试 CSV，或检查 statement 格式。")

    df = pd.DataFrame(txns)
    df["date"] = pd.to_datetime(df["date"], format="%m/%d/%y", errors="coerce")
    df["amount"] = df["amount"].apply(safe_float)
    df["direction"] = df["amount"].apply(lambda x: "debit" if x < 0 else "credit")
    df["amount_abs"] = df["amount"].abs()
    df["raw_description"] = df["description"]
    return df.dropna(subset=["date", "description", "amount"]).reset_index(drop=True)


def extract_pdf_text(uploaded_file) -> str:
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    errors = []
    text = ""

    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            parts = []
            for page in pdf.pages:
                t = page.extract_text() or ""
                parts.append(t)
            text = "\n".join(parts)
    except Exception as e:
        errors.append(f"pdfplumber: {e}")

    if not text.strip():
        try:
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(file_bytes))
            parts = []
            for page in reader.pages:
                parts.append(page.extract_text() or "")
            text = "\n".join(parts)
        except Exception as e:
            errors.append(f"pypdf: {e}")

    if not text.strip():
        raise ValueError("PDF 解析失败。请安装 pdfplumber / pypdf，或改传 CSV。\n" + "\n".join(errors))

    return text


def extract_boa_transactions_from_text(text: str) -> List[Dict]:
    """
    面向 Bank of America statement 的简单解析器：
    1. 识别以 mm/dd/yy 开头的行
    2. 允许 description 跨行，直到遇到金额
    3. 金额形如 -1,234.56 或 1,234.56
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    transactions = []

    txn_start = re.compile(r"^(?P<date>\d{2}/\d{2}/\d{2})\s+(?P<rest>.+)$")
    amount_end = re.compile(r"(?P<amount>-?\d[\d,]*\.\d{2})$")

    in_transaction_sections = False
    current = None

    section_markers = [
        "Deposits and other credits",
        "Withdrawals and other debits",
        "Service fees",
        "Checks",
        "Card account #",
        "Date Description Amount",
        "Date Transaction description Amount",
    ]

    section_stop_markers = [
        "Daily ledger balances",
        "Check images",
        "This page intentionally left blank",
        "Total service fees",
        "Subtotal for card account",
        "Total checks",
        "Page ",
    ]

    for line in lines:
        if any(marker in line for marker in section_markers):
            in_transaction_sections = True

        if any(marker in line for marker in section_stop_markers):
            if current is not None and current.get("description"):
                maybe = finalize_txn(current)
                if maybe:
                    transactions.append(maybe)
                current = None

        if not in_transaction_sections:
            continue

        start_match = txn_start.match(line)
        if start_match:
            if current is not None:
                maybe = finalize_txn(current)
                if maybe:
                    transactions.append(maybe)

            current = {
                "date": start_match.group("date"),
                "description_lines": [start_match.group("rest")],
            }
            continue

        if current is not None:
            current["description_lines"].append(line)

    if current is not None:
        maybe = finalize_txn(current)
        if maybe:
            transactions.append(maybe)

    return transactions


def finalize_txn(current: Dict) -> Optional[Dict]:
    merged = " ".join(current["description_lines"]).strip()
    merged = re.sub(r"\s+", " ", merged)

    amount_match = re.search(r"(-?\d[\d,]*\.\d{2})$", merged)
    if not amount_match:
        return None

    amount_str = amount_match.group(1)
    desc = merged[:amount_match.start()].strip()

    if not desc:
        return None

    # 过滤 page / subtotal / total 行
    bad_prefixes = [
        "TOTAL ",
        "SUBTOTAL ",
        "PAGE ",
        "DATE DESCRIPTION AMOUNT",
        "DATE TRANSACTION DESCRIPTION AMOUNT",
        "DATE CHECK # AMOUNT",
    ]
    if any(desc.upper().startswith(x) for x in bad_prefixes):
        return None

    return {
        "date": current["date"],
        "description": desc,
        "amount": safe_float(amount_str),
    }


# =========================
# Classification
# =========================
def extract_merchant_key(description: str) -> str:
    desc = normalize_text(description)

    replacements = [
        r"CONF#\s*[A-Z0-9]+",
        r"CONFIRMATION#\s*\d+",
        r"ID:\s*[A-Z0-9]+",
        r"CO ID:\s*[A-Z0-9]+",
        r"INDN:\s*[A-Z0-9X ]+",
        r"TRN:\d+",
        r"FX:[A-Z]{3}\s*[\d\.]+",
        r"DATE:\d+",
        r"TIME:\d+ ET",
        r"\b\d{2}/\d{2}/\d{2}\b",
        r"\b\d{4,}\b",
        r"\bX{4,}\b",
        r"PMT INFO:.*",
        r'"[^"]+"',
    ]

    for pattern in replacements:
        desc = re.sub(pattern, "", desc)

    desc = re.sub(r"\bPURCHASE\b", "", desc)
    desc = re.sub(r"\bMOBILE PURCHASE\b", "", desc)
    desc = re.sub(r"\bTRANSFER\b", "", desc)
    desc = re.sub(r"\bZELLE PAYMENT TO\b", "", desc)
    desc = re.sub(r"\bZELLE PAYMENT\b", "", desc)
    desc = re.sub(r"\bDES:\b", "", desc)
    desc = re.sub(r"\bCCD\b", "", desc)
    desc = re.sub(r"\bCKCD\b", "", desc)
    desc = re.sub(r"\s+", " ", desc).strip()

    tokens = desc.split()
    if len(tokens) > 8:
        desc = " ".join(tokens[:8])

    return desc


def classify_transactions(df: pd.DataFrame) -> pd.DataFrame:
    merchant_rules, keyword_rules, override_rules = load_rules()

    override_map = {
        normalize_text(str(row["merchant_key"])): row["category"]
        for _, row in override_rules.iterrows()
    }

    results = []
    for _, row in df.iterrows():
        description = str(row["description"])
        desc_norm = normalize_text(description)
        merchant_key = extract_merchant_key(description)
        amount = float(row["amount"])

        category = "其他待确认"
        confidence = 0.0
        rule_source = "unmatched"
        needs_review = True
        note = ""

        # 0) override rule
        if merchant_key in override_map:
            category = override_map[merchant_key]
            confidence = 1.0
            rule_source = "manual_override"
            needs_review = False

        # 1) credit heuristics
        elif amount > 0:
            if "TIKTOK INC" in desc_norm or "PAYMENT" in desc_norm or "REFUND" in desc_norm:
                category = "收入/回款"
                confidence = 0.95
                rule_source = "credit_heuristic"
                needs_review = False
            else:
                category = "收入/回款"
                confidence = 0.60
                rule_source = "generic_credit"
                needs_review = True

        # 2) merchant rules
        if rule_source == "unmatched":
            for _, rule in merchant_rules.iterrows():
                pattern = normalize_text(rule["pattern"])
                if pattern and pattern in desc_norm:
                    category = rule["category"]
                    confidence = float(rule["confidence"])
                    rule_source = "merchant_rule"
                    needs_review = confidence < 0.90
                    break

        # 3) keyword rules
        if rule_source == "unmatched":
            for _, rule in keyword_rules.iterrows():
                pattern = normalize_text(rule["pattern"])
                if pattern and pattern in desc_norm:
                    category = rule["category"]
                    confidence = float(rule["confidence"])
                    rule_source = "keyword_rule"
                    needs_review = confidence < 0.90
                    break

        # 4) special heuristics
        if rule_source == "unmatched":
            if "WIRE TYPE:FX OUT" in desc_norm and ("ALADDIN" in desc_norm or "DBS BANK" in desc_norm):
                category = "穿戴甲进货成本"
                confidence = 0.92
                rule_source = "wire_supplier_heuristic"
                needs_review = False
            elif "EXTERNAL TRANSFER FEE" in desc_norm:
                category = "银行手续费"
                confidence = 0.99
                rule_source = "fee_heuristic"
                needs_review = False
            elif "ZELLE" in desc_norm and "PAYROLL" in desc_norm:
                category = "员工工资"
                confidence = 0.96
                rule_source = "zelle_payroll_heuristic"
                needs_review = False
            elif any(k in desc_norm for k in LIKELY_AMBIGUOUS_KEYWORDS):
                category = "其他待确认"
                confidence = 0.30
                rule_source = "ambiguous_spend"
                needs_review = True
                note = "疑似餐饮 / Amazon / 杂费，建议人工确认。"
            elif amount < 0:
                category = "其他待确认"
                confidence = 0.20
                rule_source = "generic_debit"
                needs_review = True

        results.append(
            {
                **row.to_dict(),
                "merchant_key": merchant_key,
                "auto_category": category,
                "final_category": category,
                "confidence": round(confidence, 2),
                "rule_source": rule_source,
                "needs_review": needs_review,
                "note": note,
            }
        )

    out = pd.DataFrame(results)
    out["month"] = out["date"].dt.to_period("M").astype(str)
    return out


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    debit_df = df[df["amount"] < 0].copy()
    if debit_df.empty:
        return pd.DataFrame(columns=["序号", "支出项目名称", "金额", "备注说明"])

    summary = (
        debit_df.groupby("final_category", dropna=False)["amount_abs"]
        .sum()
        .reset_index()
        .rename(columns={"final_category": "支出项目名称", "amount_abs": "金额"})
        .sort_values("金额", ascending=False)
        .reset_index(drop=True)
    )
    summary["金额"] = summary["金额"].round(2)
    summary.insert(0, "序号", range(1, len(summary) + 1))
    summary["备注说明"] = ""
    return summary[["序号", "支出项目名称", "金额", "备注说明"]]


def mark_categories_from_editor(base_df: pd.DataFrame, edited_df: pd.DataFrame) -> pd.DataFrame:
    df = base_df.copy()
    if edited_df is None or edited_df.empty:
        return df

    if "row_id" not in df.columns or "row_id" not in edited_df.columns:
        return df

    edit_map = edited_df.set_index("row_id")["final_category"].to_dict()
    df["final_category"] = df["row_id"].map(edit_map).fillna(df["final_category"])
    return df


# =========================
# UI
# =========================
def init_session():
    for key in ["transactions_df", "classified_df", "edited_df"]:
        if key not in st.session_state:
            st.session_state[key] = None


def sidebar_rule_manager():
    st.sidebar.header("规则库管理")
    merchant_rules, keyword_rules, override_rules = load_rules()

    with st.sidebar.expander("查看商户规则", expanded=False):
        st.dataframe(merchant_rules, use_container_width=True)

    with st.sidebar.expander("查看关键词规则", expanded=False):
        st.dataframe(keyword_rules, use_container_width=True)

    with st.sidebar.expander("查看人工覆盖规则", expanded=False):
        st.dataframe(override_rules, use_container_width=True)

    with st.sidebar.expander("新增规则", expanded=False):
        rule_type = st.selectbox("规则类型", ["merchant", "keyword"], key="new_rule_type")
        pattern = st.text_input("匹配文本 / 关键词", key="new_rule_pattern")
        category = st.selectbox("分类", CATEGORY_OPTIONS, key="new_rule_category")
        confidence = st.slider("置信度", 0.5, 1.0, 0.95, 0.01, key="new_rule_conf")
        if st.button("保存新规则"):
            if pattern.strip():
                save_new_pattern_rule(pattern, category, rule_bucket=rule_type, confidence=float(confidence))
                st.success("规则已保存，刷新后生效。")
            else:
                st.warning("请先填写匹配文本。")


def render_top_intro():
    st.title("银行 Statement 分类汇总助手")
    st.caption("上传 PDF / CSV / Excel → 自动分类 → 只确认少量不确定项 → 导出汇总表")

    with st.expander("适合你的工作流", expanded=False):
        st.markdown(
            """
1. 上传银行 statement  
2. 程序自动解析交易  
3. 根据规则库自动分类  
4. 你只需要改“待确认”项目  
5. 导出 Excel 汇总表，直接做月度费用明细
            """
        )


def main():
    init_session()
    sidebar_rule_manager()
    render_top_intro()

    upload = st.file_uploader("上传银行 statement", type=["pdf", "csv", "xlsx", "xls"])

    col_a, col_b = st.columns([2, 1])
    with col_a:
        review_threshold = st.slider("低于该置信度自动标记为“待确认”", 0.50, 1.00, 0.90, 0.01)
    with col_b:
        only_debits = st.checkbox("仅展示支出", value=True)

    if upload is None:
        st.info("先上传文件。PDF 推荐 BoA / Chase / AMEX 的标准 statement；CSV/Excel 更稳。")
        st.code(PDF_ENGINE_HELP)
        st.stop()

    try:
        raw_df = parse_uploaded_file(upload)
        raw_df = raw_df.sort_values(["date", "amount"], ascending=[True, True]).reset_index(drop=True)
        raw_df["row_id"] = range(1, len(raw_df) + 1)
        st.session_state["transactions_df"] = raw_df
    except Exception as e:
        st.error(f"文件解析失败：{e}")
        st.stop()

    classified_df = classify_transactions(st.session_state["transactions_df"])
    classified_df["needs_review"] = (classified_df["confidence"] < review_threshold) | (classified_df["final_category"] == "其他待确认")
    st.session_state["classified_df"] = classified_df.copy()

    working_df = classified_df.copy()
    if only_debits:
        working_df = working_df[working_df["amount"] < 0].copy()

    st.subheader("1）交易明细")
    show_cols = [
        "row_id", "date", "description", "amount", "merchant_key",
        "auto_category", "final_category", "confidence", "rule_source", "needs_review", "note"
    ]
    st.dataframe(
        working_df[show_cols].sort_values(["date", "row_id"]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "date": st.column_config.DateColumn("日期"),
            "amount": st.column_config.NumberColumn("金额", format="$%.2f"),
            "confidence": st.column_config.NumberColumn("置信度", format="%.2f"),
            "needs_review": st.column_config.CheckboxColumn("待确认"),
        },
    )

    st.subheader("2）待确认交易")
    unknown_df = working_df[working_df["needs_review"]].copy()

    if unknown_df.empty:
        st.success("本次没有待确认项目。")
        edited_unknown_df = unknown_df.copy()
    else:
        st.caption("这里只需要改不确定的项目，不需要逐笔全部重选。")
        edited_unknown_df = st.data_editor(
            unknown_df[["row_id", "date", "description", "amount", "merchant_key", "final_category", "note"]].copy(),
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "date": st.column_config.DateColumn("日期"),
                "amount": st.column_config.NumberColumn("金额", format="$%.2f"),
                "final_category": st.column_config.SelectboxColumn("最终分类", options=CATEGORY_OPTIONS),
                "note": st.column_config.TextColumn("备注"),
            },
            key="unknown_editor",
        )

    updated_df = mark_categories_from_editor(classified_df.copy(), edited_unknown_df)
    st.session_state["edited_df"] = updated_df.copy()

    st.subheader("3）一键把你的修改写入规则库")
    with st.expander("保存本次确认结果", expanded=False):
        candidate_rows = updated_df.merge(
            classified_df[["row_id", "final_category"]].rename(columns={"final_category": "old_category"}),
            on="row_id",
            how="left",
        )
        changed_rows = candidate_rows[candidate_rows["final_category"] != candidate_rows["old_category"]].copy()

        if changed_rows.empty:
            st.info("当前没有新的人工修改。")
        else:
            st.dataframe(
                changed_rows[["row_id", "description", "merchant_key", "old_category", "final_category"]],
                use_container_width=True,
                hide_index=True,
            )

            save_mode = st.radio(
                "保存方式",
                [
                    "按 merchant_key 保存（推荐）",
                    "按 description 前缀保存为 merchant 规则",
                    "不保存，仅本次使用",
                ],
                horizontal=False,
            )

            if st.button("保存修改到规则库"):
                if save_mode == "按 merchant_key 保存（推荐）":
                    for _, r in changed_rows.iterrows():
                        save_override_rule(str(r["merchant_key"]), str(r["final_category"]))
                elif save_mode == "按 description 前缀保存为 merchant 规则":
                    for _, r in changed_rows.iterrows():
                        save_new_pattern_rule(str(r["merchant_key"]), str(r["final_category"]), rule_bucket="merchant", confidence=0.95)
                st.success("已保存。下次遇到类似商户会自动分类。")

    st.subheader("4）分类汇总")
    summary_df = build_summary(updated_df)

    st.dataframe(
        summary_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "金额": st.column_config.NumberColumn("金额", format="$%.2f")
        },
    )

    total_expense = summary_df["金额"].sum() if not summary_df.empty else 0.0
    c1, c2, c3 = st.columns(3)
    c1.metric("支出总笔数", int((updated_df["amount"] < 0).sum()))
    c2.metric("待确认笔数", int(updated_df["needs_review"].sum()))
    c3.metric("支出合计", amount_to_display(total_expense))

    st.subheader("5）导出")
    excel_bytes = to_excel_bytes(
        detail_df=updated_df.sort_values(["date", "row_id"]),
        summary_df=summary_df,
        unknown_df=updated_df[updated_df["needs_review"]].sort_values(["date", "row_id"]),
    )

    st.download_button(
        label="下载 Excel 结果",
        data=excel_bytes,
        file_name=f"bank_statement_classification_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    csv_bytes = updated_df.sort_values(["date", "row_id"]).to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="下载交易明细 CSV",
        data=csv_bytes,
        file_name=f"bank_statement_detail_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
    )

    st.subheader("6）建议你下一步怎么用")
    st.markdown(
        """
- 第一次跑：把未知项改正确，顺手保存到规则库  
- 第二次开始：大部分交易会自动命中  
- 后面每个月：基本只需要确认少量新商户  
- 如果你以后想更强，我建议再加：
  - 供应商主数据表
  - 员工名单表
  - Zelle / Wire / Card 独立规则
  - 自动生成“与你这张月度费用明细表一致”的格式
        """
    )


if __name__ == "__main__":
    main()
