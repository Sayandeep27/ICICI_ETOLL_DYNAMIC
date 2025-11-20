# generate_voucher.py
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
from pathlib import Path


# voucher and upload sheet is being generated from DSR repo
# date folder creation problem solved

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

# helpers
def to_decimal(x):
    try:
        if pd.isna(x):
            return Decimal("0")
    except:
        pass
    s = str(x).replace(",", "").strip()
    if not s or s.lower() == "nan":
        return Decimal("0")
    try:
        return Decimal(s)
    except:
        try:
            return Decimal(str(float(s)))
        except:
            return Decimal("0")

def round2(d):
    if not isinstance(d, Decimal):
        d = Decimal(str(d))
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


# ---------------- MAIN ----------------
def generate_voucher():
    df = pd.read_excel(DSR_FILE)
    df.columns = df.columns.str.strip()

    # ffill TC & TT only
    if COL_TRANSACTION_CYCLE in df.columns:
        df[COL_TRANSACTION_CYCLE] = df[COL_TRANSACTION_CYCLE].ffill()
    if COL_TRANSACTION_TYPE in df.columns:
        df[COL_TRANSACTION_TYPE] = df[COL_TRANSACTION_TYPE].ffill()

    # ------------------- FIXED SETTLEMENT DATE LOGIC -------------------
    settlement = None

    if COL_SETTLEMENT_DATE in df.columns:
        for v in df[COL_SETTLEMENT_DATE]:
            if pd.notna(v):

                # CASE 1: Excel datetime / pandas Timestamp
                if isinstance(v, (datetime, pd.Timestamp)):
                    settlement = v.date()
                    break

                # CASE 2: String formats
                s = str(v).strip()
                if s and s.lower() != "nan":
                    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
                        try:
                            settlement = datetime.strptime(s, fmt).date()
                            break
                        except:
                            pass
                    if settlement:
                        break

    if settlement is None:
        settlement = datetime.today().date()
    # -------------------------------------------------------------------

    yyyymmdd = settlement.strftime("%Y%m%d")
    ddmmyy = settlement.strftime("%d%m%y")
    dd_mm_yy = settlement.strftime("%d.%m.%y")
    cycle = f"{RUN_NUMBER}C"

    # normalize helpers
    df["TC"] = df[COL_TRANSACTION_CYCLE].astype(str).str.strip().str.lower()
    df["TT"] = df[COL_TRANSACTION_TYPE].astype(str).str.strip().str.lower()
    df["CH"] = df[COL_CHANNEL].astype(str).str.strip().str.lower()

    # ---------- FINAL NET AMT: LAST NON-EMPTY ----------
    total_final = Decimal("0")
    if COL_FINAL_NET_AMT in df.columns:
        col = df[COL_FINAL_NET_AMT].dropna().astype(str).str.strip()
        col = col[col != ""]
        if not col.empty:
            total_final = round2(to_decimal(col.iloc[-1]))

    print("Final Net Amt (Rightmost+Lowest) =", total_final)

    # ---------- INWARD GST ----------
    income_debit = income_credit = gst_debit = gst_credit = Decimal("0")
    if COL_INWARD_OUTWARD in df.columns:
        cond = df[COL_INWARD_OUTWARD].astype(str).str.upper().str.strip() == "INWARD GST"
        rows = df[cond]
        if not rows.empty:
            idx = rows.index[0]

            if idx > 0:
                ra = df.iloc[idx - 1]
                income_debit = round2(to_decimal(ra.get(COL_SERVICE_FEE_DR, 0)))
                income_credit = round2(to_decimal(ra.get(COL_SERVICE_FEE_CR, 0)))

            rg = rows.iloc[0]
            gst_debit = round2(to_decimal(rg.get(COL_SERVICE_FEE_DR, 0)))
            gst_credit = round2(to_decimal(rg.get(COL_SERVICE_FEE_CR, 0)))

    print(f"Derived INWARD values -> Income Debit: {income_debit}, GST Debit: {gst_debit}, Income Credit: {income_credit}, GST Credit: {gst_credit}")

    # ---------------- BUILD VOUCHER ----------------
    voucher = []

    for acct, tmpl, desc in TEMPLATE:
        narration = tmpl.replace("{yyyymmdd}", yyyymmdd)\
                        .replace("{ddmmyy}", ddmmyy)\
                        .replace("{dd_mm_yy}", dd_mm_yy)\
                        .replace("{cycle}", cycle)

        # spacer
        if acct == "" and desc == "":
            voucher.append(["","","","",""])
            continue

        rule = RULES.get(desc, {})

        # final net
        if rule.get("special") == "final":
            voucher.append([acct, float(total_final), "", narration, desc])
            continue

        # inward dr
        if rule.get("special") == "inward_dr":
            amt = income_debit if desc == "Income Debit" else gst_debit
            voucher.append([acct, float(amt) if amt != 0 else "", "", narration, desc])
            continue

        # inward cr
        if rule.get("special") == "inward_cr":
            amt = income_credit if desc == "Income Credit" else gst_credit
            voucher.append([acct, "", float(amt) if amt != 0 else "", narration, desc])
            continue

        # arbitration vedict
        if desc == "Arbitration Vedict":
            rows = df[
                (df["TC"] == "arbitration vedict") &
                (df["TT"].isin(["debit", "non_fin"])) &
                (df[COL_CHANNEL].notna()) &
                (df[COL_CHANNEL].astype(str).str.strip() != "")
            ]
            amt = round2(sum(rows[COL_SETAMTDR].apply(to_decimal), Decimal("0")))
            voucher.append([acct, float(amt) if amt != 0 else "", "", narration, desc])
            print("Arbitration Vedict: summed SETAMTDR =", amt, "rows:", len(rows))
            continue

        # normal
        cycles = [c.lower() for c in rule.get("cycles", [])]
        sel = df[df["TC"].isin(cycles)]
        sum_col = rule.get("sum_col")
        side = rule.get("side")

        if side == "goodfaith":
            dr = round2(sum(sel[COL_SETAMTDR].apply(to_decimal), Decimal("0")))
            cr = round2(sum(sel[COL_SETAMTCR].apply(to_decimal), Decimal("0")))
            if dr != 0:
                voucher.append([acct, float(dr), "", narration, desc])
            elif cr != 0:
                voucher.append([acct, "", float(cr), narration, desc])
            else:
                voucher.append([acct,"","",narration,desc])
            continue

        amt = round2(sum(sel[sum_col].apply(to_decimal), Decimal("0"))) if sum_col else Decimal("0")

        if side == "credit":
            voucher.append([acct, "", float(amt) if amt != 0 else "", narration, desc])
        else:
            voucher.append([acct, float(amt) if amt != 0 else "", "", narration, desc])

    voucher_df = pd.DataFrame(voucher, columns=["Account No","Debit","Credit","Narration","Description"])

    # ---------------- UPLOAD SHEET ----------------
    upload_df = voucher_df[["Account No","Debit","Credit","Narration"]].copy()

    def cd(row):
        d = to_decimal(row["Debit"])
        c = to_decimal(row["Credit"])
        if d != 0:
            return "D"
        if c != 0:
            return "C"
        return ""

    upload_df.insert(1,"C/D", upload_df.apply(cd, axis=1))

    def amount(row):
        d = to_decimal(row["Debit"])
        if d != 0:
            return float(d)
        c = to_decimal(row["Credit"])
        if c != 0:
            return float(c)
        return ""

    upload_df["Amount"] = upload_df.apply(amount, axis=1)

    upload_df = upload_df[["Account No","C/D","Amount","Narration"]]
    upload_df = upload_df[upload_df["Amount"] != ""].reset_index(drop=True)

    # TALLY
    d_total = round2(sum(voucher_df["Debit"].apply(to_decimal)))
    c_total = round2(sum(voucher_df["Credit"].apply(to_decimal)))

    print("Voucher totals -> Debit:", d_total, "Credit:", c_total)

    folder = OUTPUT_ROOT / settlement.strftime("%Y") / settlement.strftime("%m") / settlement.strftime("%d")
    folder.mkdir(parents=True, exist_ok=True)

    file = folder / f"ETOLL_ACQUIRING_VOUCHER_{ddmmyy}_N{RUN_NUMBER}.xlsx"

    if d_total != c_total:
        err = folder / f"ERROR_ETOLL_ACQUIRING_VOUCHER_{ddmmyy}_N{RUN_NUMBER}.xlsx"
        with pd.ExcelWriter(err, engine="openpyxl") as w:
            voucher_df.to_excel(w, "Voucher", index=False)
            upload_df.to_excel(w, "Upload", index=False)
        print("ERROR: Debit and credit not tallied:", err)
        return

    with pd.ExcelWriter(file, engine="openpyxl") as w:
        voucher_df.to_excel(w, "Voucher", index=False)
        upload_df.to_excel(w, "Upload", index=False)

    print("Voucher + Upload saved at:", file)


if __name__ == "__main__":
    generate_voucher()
