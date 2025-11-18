# generate_voucher.py
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
from pathlib import Path

# dynamic + every value is correct


# ------------- CONFIG -------------
DSR_FILE = "dsr_report.xlsx"
RUN_NUMBER = 1
OUTPUT_ROOT = Path("E-tollAcquiringSettlement/Processing")

# Column names
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

# Template (unchanged)
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

# Mapping rules
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

# Helpers
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

def round2(d: Decimal):
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


# ======================= MAIN =======================
def generate_voucher():

    df = pd.read_excel(DSR_FILE)
    df.columns = df.columns.str.strip()

    # --------- IMPORTANT CHANGES ---------
    # ffill ONLY for Transaction Cycle and Transaction Type
    df[COL_TRANSACTION_CYCLE] = df[COL_TRANSACTION_CYCLE].ffill()
    df[COL_TRANSACTION_TYPE] = df[COL_TRANSACTION_TYPE].ffill()

    # DO NOT FFILL CHANNEL
    # df[COL_CHANNEL] stays as is
    # -------------------------------------

    # Settlement date
    settlement = datetime.today().date()
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
    cycle_suffix = f"{RUN_NUMBER}C"

    # Normalize TC, TT, CH
    df["TC"] = df[COL_TRANSACTION_CYCLE].astype(str).str.strip().str.lower()
    df["TT"] = df[COL_TRANSACTION_TYPE].astype(str).str.strip().str.lower()
    df["CH"] = df[COL_CHANNEL].astype(str).str.strip().str.lower()

    # --- FINAL NET AMT (row above INWARD GST)
    final_amt = Decimal("0")
    inward_gst_rows = df[df[COL_INWARD_OUTWARD].astype(str).str.upper().str.strip() == "INWARD GST"]

    if not inward_gst_rows.empty:
        idx = inward_gst_rows.index[0]
        if idx > 0:
            final_amt = round2(to_decimal(df.iloc[idx - 1][COL_FINAL_NET_AMT]))

    if final_amt < 0:
        print("Final Net Amt negative. Terminating.")
        return

    # --- INCOME/GST (Option A)
    income_debit_val = income_credit_val = Decimal("0")
    gst_debit_val = gst_credit_val = Decimal("0")

    if not inward_gst_rows.empty:
        idx = inward_gst_rows.index[0]

        # INCOME values from row above
        if idx > 0:
            row = df.iloc[idx - 1]
            income_debit_val = round2(to_decimal(row.get(COL_SERVICE_FEE_DR, 0)))
            income_credit_val = round2(to_decimal(row.get(COL_SERVICE_FEE_CR, 0)))

        # GST from INWARD GST row
        rowg = inward_gst_rows.iloc[0]
        gst_debit_val = round2(to_decimal(rowg.get(COL_SERVICE_FEE_DR, 0)))
        gst_credit_val = round2(to_decimal(rowg.get(COL_SERVICE_FEE_CR, 0)))

    # ----- BUILD VOUCHER -----
    voucher_rows = []

    for acct, narr_tmpl, desc in TEMPLATE:

        narration = (narr_tmpl.replace("{yyyymmdd}", yyyymmdd)
                               .replace("{ddmmyy}", ddmmyy)
                               .replace("{dd_mm_yy}", dd_mm_yy)
                               .replace("{cycle}", cycle_suffix))

        if acct == "" and desc == "":
            voucher_rows.append(["","","","",""])
            continue

        rule = RULES.get(desc, {})

        # Final Net
        if rule.get("special") == "final":
            voucher_rows.append([acct, float(final_amt), "", narration, desc])
            continue

        # Inward DR
        if rule.get("special") == "inward_dr":
            amt = income_debit_val if desc == "Income Debit" else gst_debit_val
            voucher_rows.append([acct, float(amt) if amt != 0 else "", "", narration, desc])
            continue

        # Inward CR
        if rule.get("special") == "inward_cr":
            amt = income_credit_val if desc == "Income Credit" else gst_credit_val
            voucher_rows.append([acct, "", float(amt) if amt != 0 else "", narration, desc])
            continue

        # -------- Arbitration Vedict special logic --------
        if desc == "Arbitration Vedict":

            arb_rows = df[(df["TC"] == "arbitration vedict") &
                          (df["TT"].isin(["debit", "non_fin"])) &
                          (df[COL_CHANNEL].notna()) &
                          (df[COL_CHANNEL] != "")]

            amt = round2(sum(arb_rows[COL_SETAMTDR].apply(to_decimal), Decimal("0")))

            voucher_rows.append([acct,
                                 float(amt) if amt != 0 else "",
                                 "",
                                 narration,
                                 desc])
            continue

        # Normal rules
        cycles = rule.get("cycles", [])
        cycles = [c.lower() for c in cycles]
        sum_col = rule.get("sum_col")
        side = rule.get("side")

        sel = df[df["TC"].isin(cycles)]

        if side == "goodfaith":
            amt_dr = round2(sum(sel[COL_SETAMTDR].apply(to_decimal), Decimal("0")))
            amt_cr = round2(sum(sel[COL_SETAMTCR].apply(to_decimal), Decimal("0")))

            if amt_dr != 0:
                voucher_rows.append([acct, float(amt_dr), "", narration, desc])
            elif amt_cr != 0:
                voucher_rows.append([acct, "", float(amt_cr), narration, desc])
            else:
                voucher_rows.append([acct, "", "", narration, desc])
            continue

        amt = Decimal("0")
        if sum_col:
            amt = round2(sum(sel[sum_col].apply(to_decimal), Decimal("0")))

        if side == "credit":
            voucher_rows.append([acct, "", float(amt) if amt != 0 else "", narration, desc])
        else:
            voucher_rows.append([acct, float(amt) if amt != 0 else "", "", narration, desc])

    # Save
    out_folder = OUTPUT_ROOT / settlement.strftime("%Y") / settlement.strftime("%m") / settlement.strftime("%d")
    out_folder.mkdir(parents=True, exist_ok=True)
    outfile = out_folder / f"ETOLL_ACQUIRING_VOUCHER_{ddmmyy}_N{RUN_NUMBER}.xlsx"
    pd.DataFrame(voucher_rows,
                 columns=["Account No","Debit","Credit","Narration","Description"]).to_excel(outfile,index=False)

    print("Voucher saved at:", outfile)


if __name__ == "__main__":
    generate_voucher()
