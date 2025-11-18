# generate_voucher.py
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
from pathlib import Path

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

# Strict valid combos (as you provided)
VALID = {
    ("NETC Settled Transaction", "DEBIT", "PARKING"),
    ("NETC Settled Transaction", "DEBIT", "TOLL"),
    ("NETC Settled Transaction", "NON_FIN", "PARKING"),
    ("NETC Settled Transaction", "NON_FIN", "TOLL"),
    ("NETC Settled Transaction", "NON_FIN", "APTRIP"),
    ("DebitAdjustment", "DEBIT", "TOLL"),
    ("DebitAdjustment", "NON_FIN", "TOLL"),
    ("Pre-Arbitration Deemed Acceptance", "DEBIT", "TOLL"),
    ("Pre-Arbitration Deemed Acceptance", "DEBIT", "PARKING"),
    ("Debit chargeback deemed Acceptance", "DEBIT", "TOLL"),
    ("Debit chargeback deemed Acceptance", "DEBIT", "PARKING"),
    ("Good Faith Acceptance", "DEBIT", "TOLL"),
    ("Credit Adjustment", "DEBIT", "TOLL"),
    ("Credit Adjustment", "DEBIT", "PARKING"),
    ("Arbitration Vedict", "DEBIT", "TOLL"),
}

# Template order (output order)
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

# Per-description rules (explicit, follow your spec)
RULES = {
    "NETC Settled Transaction":         {"cycles": ["netc settled transaction"], "sum_col": COL_SETAMTCR, "side": "credit"},
    "Debit Adjustment":                 {"cycles": ["debitadjustment", "debit adjustment"], "sum_col": COL_SETAMTCR, "side": "credit"},
    "Good Faith Acceptance Credit":     {"cycles": ["good faith acceptance"], "sum_col": COL_SETAMTCR, "side": "credit"},
    "Credit Adjustment":                {"cycles": ["credit adjustment", "creditadjustment"], "sum_col": COL_SETAMTDR, "side": "debit"},
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
    if x is None:
        return Decimal("0")
    try:
        if pd.isna(x):
            return Decimal("0")
    except Exception:
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

def round2(d: Decimal):
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

# ---------------- main ----------------
def generate_voucher():
    p = Path(DSR_FILE)
    if not p.exists():
        print("DSR not found:", p.resolve())
        return

    df = pd.read_excel(p)
    df.columns = df.columns.str.strip()

    # forward-fill merged cells for TC, TT, CH
    for col in (COL_TRANSACTION_CYCLE, COL_TRANSACTION_TYPE, COL_CHANNEL):
        if col in df.columns:
            df[col] = df[col].ffill()

    # settlement date
    sd_raw = ""
    if COL_SETTLEMENT_DATE in df.columns:
        for v in df[COL_SETTLEMENT_DATE].astype(str).tolist():
            if str(v).strip() and str(v).strip().lower() != "nan":
                sd_raw = str(v).strip()
                break
    try:
        settlement = datetime.strptime(sd_raw, "%d-%m-%Y").date() if sd_raw else datetime.today().date()
    except:
        settlement = datetime.today().date()

    yyyymmdd = settlement.strftime("%Y%m%d")
    ddmmyy = settlement.strftime("%d%m%y")
    dd_mm_yy = settlement.strftime("%d.%m.%y")
    cycle_suffix = f"{RUN_NUMBER}C"

    # normalize small helper columns (lower case trimmed versions)
    df2 = df.copy()
    df2["TC"] = df2[COL_TRANSACTION_CYCLE].fillna("").astype(str).str.strip().str.lower()
    df2["TT"] = df2[COL_TRANSACTION_TYPE].fillna("").astype(str).str.strip().str.lower()
    df2["CH"] = df2[COL_CHANNEL].fillna("").astype(str).str.strip().str.lower()

    # Validate transaction cycles: build set of valid cycles from RULES
    valid_cycles = set()
    for v in RULES.values():
        for c in v.get("cycles", []) if isinstance(v.get("cycles", []), list) else []:
            valid_cycles.add(c)

    # collect DSR cycles where channel non-empty and TC non-empty
    dsr_cycles = set(df2[(df2["CH"] != "") & (df2["TC"] != "")]["TC"].unique())
    unknown = sorted([c for c in dsr_cycles if c not in valid_cycles])
    if unknown:
        print("New Type of Transaction Received:", unknown)
        print("Terminating.")
        return

    # Strict valid tuple filtering using VALID set
    valid_lower = {(a.lower(), b.lower(), c.lower()) for (a,b,c) in VALID}
    def matches_valid(r):
        return (r["TC"], r["TT"], r["CH"]) in valid_lower

    df_strict = df2[df2.apply(matches_valid, axis=1)].copy()

    # compute final net (sum of Final Net Amt for rows where Channel non-empty and TC in valid_cycles)
    mask_final = (df2["CH"] != "") & (df2["TC"] != "") & (df2["TC"].isin(valid_cycles))
    if COL_FINAL_NET_AMT in df2.columns:
        total_final = round2(sum(df2[mask_final][COL_FINAL_NET_AMT].apply(to_decimal), Decimal("0")))
    else:
        total_final = Decimal("0.00")

    if total_final < 0:
        print("Final Net Amt negative:", total_final, "Terminating.")
        return

    # --- Correct INWARD selection: try "row above INWARD GST" method first ---
    income_debit_val = Decimal("0")
    gst_debit_val = Decimal("0")
    income_credit_val = Decimal("0")
    gst_credit_val = Decimal("0")

    if COL_INWARD_OUTWARD in df.columns:
        # find indices (first occurrence)
        cond_inward_gst = df[COL_INWARD_OUTWARD].astype(str).str.upper().str.strip() == "INWARD GST"
        idxs_gst = df[cond_inward_gst].index.tolist()
        if idxs_gst:
            # take first INWARD GST index and pick the row just above as the INWARD total
            idx_gst = idxs_gst[0]
            idx_inward_candidate = idx_gst - 1
            if idx_inward_candidate >= 0:
                row_inward = df.iloc[[idx_inward_candidate]]
                # sanity: check that candidate row is INWARD or blank; if not, fallback
                maybe_tag = str(row_inward[COL_INWARD_OUTWARD].iat[0]).upper().strip()
                if maybe_tag == "INWARD" or maybe_tag == "" or maybe_tag == "TOTAL" or maybe_tag.startswith("INWARD"):
                    # use this candidate row
                    if COL_SERVICE_FEE_DR in row_inward.columns:
                        income_debit_val = round2(to_decimal(row_inward[COL_SERVICE_FEE_DR].iat[0]))
                    if COL_SERVICE_FEE_CR in row_inward.columns:
                        income_credit_val = round2(to_decimal(row_inward[COL_SERVICE_FEE_CR].iat[0]))
                else:
                    # fallback to direct row selection by INWARD
                    inward_row = df[df[COL_INWARD_OUTWARD].astype(str).str.upper().str.strip() == "INWARD"]
                    inward_row = inward_row if not inward_row.empty else df.iloc[0:0]
                    if not inward_row.empty:
                        income_debit_val = round2(to_decimal(inward_row[COL_SERVICE_FEE_DR].iloc[0])) if COL_SERVICE_FEE_DR in inward_row.columns else Decimal("0")
                        income_credit_val = round2(to_decimal(inward_row[COL_SERVICE_FEE_CR].iloc[0])) if COL_SERVICE_FEE_CR in inward_row.columns else Decimal("0")
            else:
                # fallback
                inward_row = df[df[COL_INWARD_OUTWARD].astype(str).str.upper().str.strip() == "INWARD"]
                inward_row = inward_row if not inward_row.empty else df.iloc[0:0]
                if not inward_row.empty:
                    income_debit_val = round2(to_decimal(inward_row[COL_SERVICE_FEE_DR].iloc[0])) if COL_SERVICE_FEE_DR in inward_row.columns else Decimal("0")
                    income_credit_val = round2(to_decimal(inward_row[COL_SERVICE_FEE_CR].iloc[0])) if COL_SERVICE_FEE_CR in inward_row.columns else Decimal("0")
            # gst row values (direct)
            inward_gst_row = df[cond_inward_gst]
            if not inward_gst_row.empty:
                gst_debit_val = round2(to_decimal(inward_gst_row[COL_SERVICE_FEE_DR].iloc[0])) if COL_SERVICE_FEE_DR in inward_gst_row.columns else Decimal("0")
                gst_credit_val = round2(to_decimal(inward_gst_row[COL_SERVICE_FEE_CR].iloc[0])) if COL_SERVICE_FEE_CR in inward_gst_row.columns else Decimal("0")
        else:
            # fallback: pick direct rows
            inward_row = df[df[COL_INWARD_OUTWARD].astype(str).str.upper().str.strip() == "INWARD"]
            inward_gst_row = df[df[COL_INWARD_OUTWARD].astype(str).str.upper().str.strip() == "INWARD GST"]
            if not inward_row.empty:
                income_debit_val = round2(to_decimal(inward_row[COL_SERVICE_FEE_DR].iloc[0])) if COL_SERVICE_FEE_DR in inward_row.columns else Decimal("0")
                income_credit_val = round2(to_decimal(inward_row[COL_SERVICE_FEE_CR].iloc[0])) if COL_SERVICE_FEE_CR in inward_row.columns else Decimal("0")
            if not inward_gst_row.empty:
                gst_debit_val = round2(to_decimal(inward_gst_row[COL_SERVICE_FEE_DR].iloc[0])) if COL_SERVICE_FEE_DR in inward_gst_row.columns else Decimal("0")
                gst_credit_val = round2(to_decimal(inward_gst_row[COL_SERVICE_FEE_CR].iloc[0])) if COL_SERVICE_FEE_CR in inward_gst_row.columns else Decimal("0")

    print(f"Derived INWARD values -> Income Debit: {income_debit_val}, GST Debit: {gst_debit_val}, Income Credit: {income_credit_val}, GST Credit: {gst_credit_val}")

    # build voucher rows per TEMPLATE
    voucher_rows = []
    for acct, narr_tmpl, desc in TEMPLATE:
        narration = narr_tmpl.replace("{yyyymmdd}", yyyymmdd).replace("{ddmmyy}", ddmmyy).replace("{dd_mm_yy}", dd_mm_yy).replace("{cycle}", cycle_suffix)

        # spacer
        if acct == "" and desc == "":
            voucher_rows.append(["", "", "", "", ""])
            continue

        # Final Net
        rule = RULES.get(desc, {})
        if rule.get("special") == "final":
            voucher_rows.append([acct, float(total_final) if total_final != Decimal("0.00") else "", "", narration, desc])
            continue

        # Income/GST (exact-row logic)
        if rule.get("special") == "inward_dr":
            if desc == "Income Debit":
                voucher_rows.append([acct, float(income_debit_val) if income_debit_val != Decimal("0.00") else "", "", narration, desc])
            elif desc == "GST Debit":
                voucher_rows.append([acct, float(gst_debit_val) if gst_debit_val != Decimal("0.00") else "", "", narration, desc])
            else:
                voucher_rows.append([acct, "", "", narration, desc])
            print(f"{desc}: placed inward DR value.")
            continue

        if rule.get("special") == "inward_cr":
            if desc == "Income Credit":
                voucher_rows.append([acct, "", float(income_credit_val) if income_credit_val != Decimal("0.00") else "", narration, desc])
            elif desc == "GST Credit":
                voucher_rows.append([acct, "", float(gst_credit_val) if gst_credit_val != Decimal("0.00") else "", narration, desc])
            else:
                voucher_rows.append([acct, "", "", narration, desc])
            print(f"{desc}: placed inward CR value.")
            continue

        # Special explicit handler for Arbitration Vedict: use df_strict only (strict tuple match)
        if desc == "Arbitration Vedict":
            sel_arb = df_strict[df_strict["TC"] == "arbitration vedict"]
            amt_arb = round2(sum(sel_arb[COL_SETAMTDR].apply(to_decimal), Decimal("0"))) if not sel_arb.empty and COL_SETAMTDR in sel_arb.columns else Decimal("0")
            voucher_rows.append([acct, float(amt_arb) if amt_arb != Decimal("0.00") else "", "", narration, desc])
            print(f"{desc}: summed SETAMTDR (strict) = {amt_arb} (rows: {len(sel_arb)})")
            continue

        # normal cycle-based aggregation
        cycles = [c.strip().lower() for c in rule.get("cycles", [])]
        sum_col = rule.get("sum_col")
        side = rule.get("side")

        # select strict rows where TC in cycles and Channel non-empty
        sel = df_strict[(df_strict["CH"] != "") & (df_strict["TC"].isin(cycles))].copy()

        if side == "goodfaith":
            # sum both dr and cr
            amt_dr = round2(sum(sel[COL_SETAMTDR].apply(to_decimal), Decimal("0"))) if not sel.empty and COL_SETAMTDR in sel.columns else Decimal("0.00")
            amt_cr = round2(sum(sel[COL_SETAMTCR].apply(to_decimal), Decimal("0"))) if not sel.empty and COL_SETAMTCR in sel.columns else Decimal("0.00")
            if amt_dr != Decimal("0.00"):
                voucher_rows.append([acct, float(amt_dr), "", narration, desc])
            elif amt_cr != Decimal("0.00"):
                voucher_rows.append([acct, "", float(amt_cr), narration, desc])
            else:
                voucher_rows.append([acct, "", "", narration, desc])
            print(f"{desc}: SETAMTDR={amt_dr} SETAMTCR={amt_cr} rows={len(sel)}")
            continue

        # normal: sum specified column
        amt = Decimal("0.00")
        if sum_col and not sel.empty and sum_col in sel.columns:
            amt = round2(sum(sel[sum_col].apply(to_decimal), Decimal("0")))

        if side == "credit":
            voucher_rows.append([acct, "", float(amt) if amt != Decimal("0.00") else "", narration, desc])
        else:
            voucher_rows.append([acct, float(amt) if amt != Decimal("0.00") else "", "", narration, desc])

        print(f"{desc}: summed {sum_col} = {amt} (rows counted: {len(sel)})")

    # write output
    out_folder = OUTPUT_ROOT / settlement.strftime("%Y") / settlement.strftime("%m") / settlement.strftime("%d")
    out_folder.mkdir(parents=True, exist_ok=True)
    outfile = out_folder / f"ETOLL_ACQUIRING_VOUCHER_{ddmmyy}_N{RUN_NUMBER}.xlsx"
    pd.DataFrame(voucher_rows, columns=["Account No", "Debit", "Credit", "Narration", "Description"]).to_excel(outfile, index=False)
    print("Voucher saved to:", outfile.resolve())

if __name__ == "__main__":
    generate_voucher()
