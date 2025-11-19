# generate_voucher.py
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
from pathlib import Path


# voucher and upload sheet is being generated from DSR report


# ------------- CONFIG -------------
DSR_FILE = "dsr_report.xlsx"
RUN_NUMBER = 1
OUTPUT_ROOT = Path("E-tollAcquiringSettlement/Processing")

# Column names (must match Excel)
COL_SETTLEMENT_DATE = "Settlement Date"
COL_TRANSACTION_CYCLE = "Transaction Cycle"
COL_TRANSACTION_TYPE = "Transaction Type"
COL_CHANNEL = "Channel"
COL_SETAMTDR = "SETAMTDR"
COL_SETAMTCR = "SETAMTCR"
COL_SERVICE_FEE_DR = "Service Fee Amt Dr"
COL_SERVICE_FEE_CR = "Service Fee Amt Cr"
COL_FINAL_NET_AMT = "Final Net Amt"
COL_INWARD_OUTWARD = "Inward/Outward"

# TEMPLATE (unchanged)
TEMPLATE = [
    ("0103SLRGTSRC", "NPCIR5{yyyymmdd} {ddmmyy}_{cycle} ETCAC", "Final Net Amt"),
    ("", "", ""),

    ("0103SLETCACQ", "Etoll acq {dd_mm_yy}_{cycle}", "NETC Settled Transaction"),
    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} Dr.Adj_{cycle}", "Debit Adjustment"),
    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} GF Accp_{cycle}", "Good Faith Acceptance Credit"),

    ("", "", ""),

    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} Cr.Adj_{cycle}", "Credit Adjustment"),
    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} Chbk_{cycle}", "Chargeback Acceptance"),
    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} GF Accp_{cycle}", "Good Faith Acceptance Debit"),
    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} PrArbtAc_{cycle}", "Pre-Arbitration Acceptance"),
    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} DrPrAbAc_{cycle}", "Pre-Arbitration Deemed Acceptance"),
    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} DrChbAc_{cycle}", "Debit chargeback deemed Acceptance"),
    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} ArbtAc_{cycle}", "Arbitration Acceptance"),
    ("0103SLETCACQ", "Etoll acq {dd_mm_yy} ArbtVer_{cycle}", "Arbitration Vedict"),

    ("", "", ""),

    ("0103CNETCACQ", "Etoll acq {dd_mm_yy}_{cycle}", "Income Debit"),
    ("0103SLPPCIGT", "Etoll acq {dd_mm_yy}_{cycle}", "GST Debit"),
    ("0103CNETCACQ", "Etoll acq {dd_mm_yy}_{cycle}", "Income Credit"),
    ("0103SLPPCIGT", "Etoll acq {dd_mm_yy}_{cycle}", "GST Credit"),
]

# RULES (unchanged)
RULES = {
    "NETC Settled Transaction":         {"cycles": ["netc settled transaction"], "sum_col": COL_SETAMTCR, "side": "credit"},
    "Debit Adjustment":                 {"cycles": ["debitadjustment", "debit adjustment"], "sum_col": COL_SETAMTCR, "side": "credit"},
    "Good Faith Acceptance Credit":     {"cycles": ["good faith acceptance"], "sum_col": COL_SETAMTCR, "side": "credit"},
    "Credit Adjustment":                {"cycles": ["credit adjustment"], "sum_col": COL_SETAMTDR, "side": "debit"},
    "Chargeback Acceptance":            {"cycles": ["chargeback acceptance"], "sum_col": COL_SETAMTDR, "side": "debit"},
    "Good Faith Acceptance Debit":      {"cycles": ["good faith acceptance"], "sum_col": None, "side": "goodfaith"},
    "Pre-Arbitration Acceptance":       {"cycles": ["pre-arbitration acceptance"], "sum_col": COL_SETAMTDR, "side": "debit"},
    "Pre-Arbitration Deemed Acceptance":{"cycles": ["pre-arbitration deemed acceptance"], "sum_col": COL_SETAMTDR, "side": "debit"},
    "Debit chargeback deemed Acceptance":{"cycles": ["debit chargeback deemed acceptance"], "sum_col": COL_SETAMTDR, "side": "debit"},
    "Arbitration Acceptance":           {"cycles": ["arbitration acceptance"], "sum_col": COL_SETAMTDR, "side": "debit"},
    "Arbitration Vedict":               {"cycles": ["arbitration vedict"], "sum_col": COL_SETAMTDR, "side": "debit"},
    "Income Debit":                     {"special": "inward_dr"},
    "GST Debit":                        {"special": "inward_dr"},
    "Income Credit":                    {"special": "inward_cr"},
    "GST Credit":                       {"special": "inward_cr"},
    "Final Net Amt":                    {"special": "final"},
}

# ---------------- helpers ----------------
def to_decimal(x):
    try:
        if pd.isna(x):
            return Decimal("0")
    except:
        pass
    s = str(x).replace(",", "").strip()
    if s == "" or s.lower() == "nan":
        return Decimal("0")
    try:
        return Decimal(s)
    except:
        try:
            return Decimal(str(float(s)))
        except:
            return Decimal("0")

def round2(d):
    # Accept Decimal, int, float — convert to Decimal before quantize
    if not isinstance(d, Decimal):
        try:
            d = Decimal(str(d))
        except:
            d = Decimal("0")
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


# ---------------- MAIN ----------------
def generate_voucher():
    p = Path(DSR_FILE)
    if not p.exists():
        print("DSR not found:", p.resolve())
        return

    df = pd.read_excel(p)
    df.columns = df.columns.str.strip()

    # --------- IMPORTANT: only ffill TC and TT (NOT Channel) ---------
    if COL_TRANSACTION_CYCLE in df.columns:
        df[COL_TRANSACTION_CYCLE] = df[COL_TRANSACTION_CYCLE].ffill()
    if COL_TRANSACTION_TYPE in df.columns:
        df[COL_TRANSACTION_TYPE] = df[COL_TRANSACTION_TYPE].ffill()
    # DO NOT ffill Channel

    # settlement date (first non-empty)
    settlement = datetime.today().date()
    if COL_SETTLEMENT_DATE in df.columns:
        for v in df[COL_SETTLEMENT_DATE].astype(str):
            if v.strip() and v.strip().lower() != "nan":
                try:
                    settlement = datetime.strptime(v.strip(), "%d-%m-%Y").date()
                except:
                    pass
                break

    yyyymmdd = settlement.strftime("%Y%m%d")
    ddmmyy = settlement.strftime("%d%m%y")
    dd_mm_yy = settlement.strftime("%d.%m.%y")
    cycle = f"{RUN_NUMBER}C"

    # normalized helper columns
    df["TC"] = df[COL_TRANSACTION_CYCLE].astype(str).str.strip().str.lower()
    df["TT"] = df[COL_TRANSACTION_TYPE].astype(str).str.strip().str.lower()
    df["CH"] = df[COL_CHANNEL].astype(str).str.strip().str.lower()

    # ---------- FINAL NET AMT: RIGHTMOST + LOWEST non-empty cell in the Final Net Amt column ----------
    total_final = Decimal("0.00")
    if COL_FINAL_NET_AMT in df.columns:
        # take last non-empty value
        series = df[COL_FINAL_NET_AMT].dropna().astype(str).map(lambda x: x.strip())
        series = series[series != ""]
        if not series.empty:
            raw = series.iloc[-1]
            total_final = round2(to_decimal(raw))
    else:
        total_final = Decimal("0.00")

    print("Final Net Amt (Rightmost+Lowest) =", total_final)

    if total_final < 0:
        print("Final Net Amt negative:", total_final, "Terminating.")
        return

    # ---------- INWARD / INWARD GST (Option A: row above INWARD GST = INCOME; INWARD GST row = GST) ----------
    income_debit = income_credit = gst_debit = gst_credit = Decimal("0")
    if COL_INWARD_OUTWARD in df.columns:
        cond_gst = df[COL_INWARD_OUTWARD].astype(str).str.upper().str.strip() == "INWARD GST"
        inward_gst_rows = df[cond_gst]
        if not inward_gst_rows.empty:
            idx_gst = inward_gst_rows.index[0]
            # row above => income values
            if idx_gst > 0:
                row_above = df.iloc[idx_gst - 1]
                income_debit = round2(to_decimal(row_above.get(COL_SERVICE_FEE_DR, 0)))
                income_credit = round2(to_decimal(row_above.get(COL_SERVICE_FEE_CR, 0)))
            # GST row => gst values
            row_gst = inward_gst_rows.iloc[0]
            gst_debit = round2(to_decimal(row_gst.get(COL_SERVICE_FEE_DR, 0)))
            gst_credit = round2(to_decimal(row_gst.get(COL_SERVICE_FEE_CR, 0)))

    print("Derived INWARD values -> Income Debit: {}, GST Debit: {}, Income Credit: {}, GST Credit: {}".format(
        income_debit, gst_debit, income_credit, gst_credit
    ))

    # ---------------- BUILD VOUCHER ----------------
    voucher = []

    for acct, tmpl, desc in TEMPLATE:
        narration = tmpl.replace("{yyyymmdd}", yyyymmdd)\
                        .replace("{ddmmyy}", ddmmyy)\
                        .replace("{dd_mm_yy}", dd_mm_yy)\
                        .replace("{cycle}", cycle)

        # spacer
        if acct == "" and desc == "":
            voucher.append(["", "", "", "", ""])
            continue

        rule = RULES.get(desc, {})

        # final net
        if rule.get("special") == "final":
            voucher.append([acct, float(total_final) if total_final != Decimal("0.00") else "", "", narration, desc])
            continue

        # inward dr
        if rule.get("special") == "inward_dr":
            if desc == "Income Debit":
                voucher.append([acct, float(income_debit) if income_debit != Decimal("0.00") else "", "", narration, desc])
            elif desc == "GST Debit":
                voucher.append([acct, float(gst_debit) if gst_debit != Decimal("0.00") else "", "", narration, desc])
            else:
                voucher.append([acct, "", "", narration, desc])
            continue

        # inward cr
        if rule.get("special") == "inward_cr":
            if desc == "Income Credit":
                voucher.append([acct, "", float(income_credit) if income_credit != Decimal("0.00") else "", narration, desc])
            elif desc == "GST Credit":
                voucher.append([acct, "", float(gst_credit) if gst_credit != Decimal("0.00") else "", narration, desc])
            else:
                voucher.append([acct, "", "", narration, desc])
            continue

        # Arbitration Vedict special: sum SETAMTDR where TC=arbitration vedict, TT in {debit, non_fin} and Channel present (no ffill on channel)
        if desc == "Arbitration Vedict":
            arb_sel = df[
                (df["TC"] == "arbitration vedict") &
                (df["TT"].isin(["debit", "non_fin"])) &
                (df[COL_CHANNEL].notna()) &
                (df[COL_CHANNEL].astype(str).str.strip() != "")
            ]
            arb_amt = round2(sum(arb_sel[COL_SETAMTDR].apply(to_decimal), Decimal("0")))
            voucher.append([acct, float(arb_amt) if arb_amt != Decimal("0.00") else "", "", narration, desc])
            print("Arbitration Vedict: summed SETAMTDR = {} (rows counted: {})".format(arb_amt, len(arb_sel)))
            continue

        # Normal aggregation (no strict validation — use TC only)
        cycles = [c.lower() for c in rule.get("cycles", [])] if rule.get("cycles") else []
        sum_col = rule.get("sum_col")
        side = rule.get("side")

        sel = df[df["TC"].isin(cycles)]

        if side == "goodfaith":
            amt_dr = round2(sum(sel[COL_SETAMTDR].apply(to_decimal), Decimal("0")))
            amt_cr = round2(sum(sel[COL_SETAMTCR].apply(to_decimal), Decimal("0")))
            if amt_dr != Decimal("0.00"):
                voucher.append([acct, float(amt_dr), "", narration, desc])
            elif amt_cr != Decimal("0.00"):
                voucher.append([acct, "", float(amt_cr), narration, desc])
            else:
                voucher.append([acct, "", "", narration, desc])
            continue

        amt = Decimal("0.00")
        if sum_col:
            amt = round2(sum(sel[sum_col].apply(to_decimal), Decimal("0")))

        if side == "credit":
            voucher.append([acct, "", float(amt) if amt != Decimal("0.00") else "", narration, desc])
        else:
            voucher.append([acct, float(amt) if amt != Decimal("0.00") else "", "", narration, desc])

        print(f"{desc}: summed {sum_col} = {amt} (rows counted: {len(sel)})")

    # Prepare voucher DataFrame
    voucher_df = pd.DataFrame(voucher, columns=["Account No", "Debit", "Credit", "Narration", "Description"])

    # ---------------- CREATE UPLOAD SHEET ----------------
    # Copy Account No, Debit, Credit, Narration into Upload sheet (values only)
    upload_df = voucher_df[["Account No", "Debit", "Credit", "Narration"]].copy()

    # Insert C/D column before Debit (i.e. new column named 'C/D' goes at index 1)
    # C/D logic: if Debit not null and >0 => 'D', elif Credit not null and >0 => 'C', else blank
    def cd_flag(row):
        try:
            d = to_decimal(row["Debit"])
        except:
            d = Decimal("0")
        try:
            c = to_decimal(row["Credit"])
        except:
            c = Decimal("0")
        if d != Decimal("0"):
            return "D"
        if c != Decimal("0"):
            return "C"
        return ""

    upload_df.insert(1, "C/D", upload_df.apply(cd_flag, axis=1))

    # Amount column: if Debit present use Debit else use Credit (Option 1 -> keep positive)
    def amount_val(row):
        d = to_decimal(row["Debit"])
        if d != Decimal("0"):
            return float(d)
        c = to_decimal(row["Credit"])
        if c != Decimal("0"):
            return float(c)
        return ""

    upload_df["Amount"] = upload_df.apply(amount_val, axis=1)

    # Drop original Credit column as instruction says remove Column D named Credit later
    upload_df = upload_df[["Account No", "C/D", "Amount", "Narration"]]

    # Remove rows where Amount is zero/blank
    upload_df = upload_df[upload_df["Amount"].apply(lambda x: x != "" and x is not None and float(x) != 0)]

    # Reset index for neatness
    upload_df = upload_df.reset_index(drop=True)

    # ---------------- VALIDATE TALLY: Voucher Debit total == Voucher Credit total ----------------
    voucher_debit_total = sum(voucher_df["Debit"].apply(lambda x: to_decimal(x)))
    voucher_credit_total = sum(voucher_df["Credit"].apply(lambda x: to_decimal(x)))

    voucher_debit_total = round2(voucher_debit_total)
    voucher_credit_total = round2(voucher_credit_total)

    print("Voucher totals -> Debit:", voucher_debit_total, "Credit:", voucher_credit_total)

    out_folder = OUTPUT_ROOT / settlement.strftime("%Y") / settlement.strftime("%m") / settlement.strftime("%d")
    out_folder.mkdir(parents=True, exist_ok=True)

    outfile = out_folder / f"ETOLL_ACQUIRING_VOUCHER_{ddmmyy}_N{RUN_NUMBER}.xlsx"

    # If mismatch -> write error file and stop (user requested this behavior)
    if voucher_debit_total != voucher_credit_total:
        errfile = out_folder / f"ERROR_ETOLL_ACQUIRING_VOUCHER_{ddmmyy}_N{RUN_NUMBER}.xlsx"
        with pd.ExcelWriter(errfile, engine="openpyxl") as ew:
            voucher_df.to_excel(ew, sheet_name="Voucher", index=False)
            upload_df.to_excel(ew, sheet_name="Upload", index=False)
        print("ERROR: Debit and credit not tallied. Written error file:", errfile.resolve())
        # STOP processing as per requirement
        return

    # Save voucher and upload into same workbook (Voucher sheet + Upload sheet)
    with pd.ExcelWriter(outfile, engine="openpyxl") as ew:
        voucher_df.to_excel(ew, sheet_name="Voucher", index=False)
        upload_df.to_excel(ew, sheet_name="Upload", index=False)

    print("Voucher + Upload saved at:", outfile.resolve())


if __name__ == "__main__":
    generate_voucher()
