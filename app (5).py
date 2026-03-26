def parse_bofa(data: bytes) -> pd.DataFrame:
    try:
        import pdfplumber
    except ImportError:
        st.error("缺少依赖: pip install pdfplumber")
        return pd.DataFrame()

    with pdfplumber.open(io.BytesIO(data)) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages)

    SKIP_PREFIXES = (
        "Date Description", "continued on the next page", "BUSINESS ADVANTAGE",
        "Your checking account", "Page ", "COLORFOUR LLC !", "See the big picture",
        "To learn more", "When you use the QRC", "Mobile Banking requires",
        "Message and data rates", "Help prevent", "Consider writing",
        "You can also set up", "Scan the code", "Moving from checks",
        "Please see", "SSM", "PULL:", "NEW:", "Find more", "Explore your",
        "Bank of America", "Equal Housing", "bofa.com", "bankofamerica.com",
    )

    DATE_LINE = re.compile(r'^(\d{2}/\d{2})/\d{2}\s+(.+?)\s+(-?[\d,]+\.\d{2})\s*$')
    CHECK_LINE = re.compile(r'^(\d{2}/\d{2})/\d{2}\s+(\d+)\s+([\d,]+\.\d{2})\s*$')
    FEE_LINE   = re.compile(r'^(\d{2}/\d{2})/\d{2}\s+(.+?)\s+([\d,]+\.\d{2})\s*$')

    def skip(line):
        if not line: return True
        for p in SKIP_PREFIXES:
            if line.startswith(p): return True
        return False

    lines = text.split("\n")
    rows = []
    section = None
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        if line in ("Deposits and other credits", "Deposits and other credits - continued"):
            section = "dep"; i += 1; continue
        if line in ("Withdrawals and other debits", "Withdrawals and other debits - continued"):
            section = "wdl"; i += 1; continue
        if line.startswith("Checks") and "Total" not in line and "image" not in line:
            section = "chk"; i += 1; continue
        if "Service fees" in line and "Total" not in line and len(line) < 20:
            section = "fee"; i += 1; continue
        if any(x in line for x in ["Daily ledger", "Check images",
               "Total deposits", "Total withdrawals", "Total checks", "Total service",
               "Total # of", "Note your Ending"]):
            i += 1; continue

        if section in ("dep", "wdl", "chk", "fee") and re.match(r'^\d{2}/\d{2}/\d{2}', line):
            parts = [line]
            j = i + 1
            while j < len(lines):
                nxt = lines[j].strip()
                if not nxt or re.match(r'^\d{2}/\d{2}/\d{2}', nxt) or skip(nxt): break
                if nxt in ("Deposits and other credits", "Deposits and other credits - continued",
                           "Withdrawals and other debits", "Withdrawals and other debits - continued",
                           "Checks", "Service fees"): break
                parts.append(nxt)
                j += 1

            full = " ".join(parts)

            if section == "chk":
                m = CHECK_LINE.match(full)
                if m:
                    rows.append({"date": m.group(1), "desc": f"Check #{m.group(2)}",
                                 "amount": -float(m.group(3).replace(",", "")), "src": "BofA"})
                    i = j; continue
            else:
                m = DATE_LINE.match(full)
                if m:
                    amt = float(m.group(3).replace(",", ""))
                    if section in ("wdl", "fee") and amt > 0:
                        amt = -amt
                    rows.append({"date": m.group(1), "desc": m.group(2).strip(),
                                 "amount": amt, "src": "BofA"})
                    i = j; continue
        i += 1

    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).drop_duplicates(subset=["date", "desc", "amount"])
    df["cat"]    = df["desc"].apply(cls_bank)
    df["status"] = df["cat"].apply(lambda c: "auto" if c != "待确认" else "pending")
    return df
