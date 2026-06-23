import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO
from rapidfuzz import process, fuzz
import re
from datetime import datetime

st.set_page_config(page_title="Universal Bank → Tally Mapper", layout="wide")

# Custom CSS for SUSPENSE Helper
st.markdown("""
<style>
    .suspense-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 20px 0;
    }
    .group-card {
        background: white;
        border-left: 5px solid #3498db;
        padding: 15px;
        margin: 10px 0;
        border-radius: 5px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .group-card-green {
        border-left-color: #2ecc71;
    }
    .group-card-orange {
        border-left-color: #e67e22;
    }
    .group-card-red {
        border-left-color: #e74c3c;
    }
    .amount-high {
        color: #e74c3c;
        font-weight: bold;
        font-size: 1.1em;
    }
    .amount-medium {
        color: #e67e22;
        font-weight: bold;
    }
    .amount-low {
        color: #2ecc71;
    }
    .success-box {
        background: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------- CLEAN NARRATION ---------------------- #
def clean_narration(text):
    if not isinstance(text, str):
        return ""
    text = text.upper()
    text = re.sub(r'\b(RTGS|NEFT|IMPS|UPI|DR|CR|CNRBR|ICIC|FAST|BANK|REF|FROM|TO|BY|PAYMENT|TRANSFER|CREDIT|DEBIT)\b', '', text)
    text = re.sub(r'[^A-Z ]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# ---------------------- DATE CONVERT ---------------------- #
def to_date_str(val):
    try:
        if isinstance(val, str):
            val = val.strip()
            for fmt in ("%d-%b-%Y", "%d/%b/%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
                try:
                    dt = datetime.strptime(val, fmt)
                    return dt.strftime("%d-%m-%Y")
                except ValueError:
                    continue
            dt = pd.to_datetime(val, dayfirst=True, errors='coerce')
            if pd.notna(dt):
                return dt.strftime("%d-%m-%Y")
        elif pd.notna(val):
            dt = pd.to_datetime(val, errors='coerce')
            if pd.notna(dt):
                return dt.strftime("%d-%m-%Y")
    except Exception:
        pass
    return ""

# ---------------------- PDF PARSER ---------------------- #
def parse_custom_pdf(file, header_text):
    rows = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                if not any(cell and str(cell).strip() for cell in row):
                    continue
                text_join = " ".join(str(c).upper() for c in row)
                if any(x in text_join for x in [
                    "STATEMENT OF ACCOUNT", "PAGE NO", "ACCOUNT BRANCH",
                    "JOINT HOLDERS", "NOMINATION", "EMAIL", "LIMITED", "SUMMARY"
                ]):
                    continue
                rows.append(row)

    if not rows:
        raise ValueError("No valid table data found in PDF. Try adjusting headers or check if it's scanned.")

    headers = [h.strip().upper() for h in header_text.split(",")]
    sample_len = len(rows[0])
    if len(headers) != sample_len:
        st.warning(f"⚠️ Adjusting headers: PDF has {sample_len} columns but you typed {len(headers)}.")
        headers = headers[:sample_len]

    df = pd.DataFrame(rows[1:], columns=headers)

    def normalize_col(col): return col.replace(" ", "").replace(".", "").upper()
    norm_cols = [normalize_col(c) for c in df.columns]
    debit_col = credit_col = date_col = narr_col = None

    for i, c in enumerate(norm_cols):
        if any(k in c for k in ["DEBIT", "WITHDRAW"]):
            debit_col = df.columns[i]
        if any(k in c for k in ["CREDIT", "DEPOSIT"]):
            credit_col = df.columns[i]
        if "DATE" in c and not date_col:
            date_col = df.columns[i]
        if any(k in c for k in ["DESC", "PARTICULAR", "NARR"]):
            narr_col = df.columns[i]

    if not date_col:
        date_col = df.columns[0]
    if not narr_col:
        narr_col = df.columns[1]

    for col in [debit_col, credit_col]:
        if col:
            df[col] = df[col].astype(str).str.replace(",", "").str.strip()
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df_out = pd.DataFrame({
        "DATE": df[date_col],
        "NARRITION": df[narr_col],
        "DEBIT": df[debit_col] if debit_col else 0,
        "CREDIT": df[credit_col] if credit_col else 0
    })
    return df_out

# ---------------------- SUSPENSE HELPER FUNCTIONS ---------------------- #
def group_similar_narrations(df_suspense):
    """Group similar narrations together"""
    groups = {}
    for idx, row in df_suspense.iterrows():
        narr = clean_narration(str(row['NARRITION']))
        found = False
        for key in groups.keys():
            if fuzz.token_set_ratio(narr, key) >= 80:
                groups[key].append(idx)
                found = True
                break
        if not found:
            groups[narr] = [idx]
    return groups

def get_amount_category(amount):
    """Categorize amount for color coding"""
    if amount >= 10000:
        return "high", "#e74c3c"
    elif amount >= 1000:
        return "medium", "#e67e22"
    else:
        return "low", "#2ecc71"

# ---------------------- MAIN APP ---------------------- #
def main():
    st.title("🏦 Universal Bank Statement → Tally Mapper (Darshan Pro v3.1)")
    st.markdown("💡 Fixed: SUSPENSE mapping now updates in final Excel! ✨")

    # Initialize session state for dataframe persistence
    if 'df_main' not in st.session_state:
        st.session_state.df_main = None
    if 'mapping_done' not in st.session_state:
        st.session_state.mapping_done = False

    file = st.file_uploader("📂 Upload Bank Statement (Excel or PDF)", type=["xlsx","pdf"])
    ledger_file = st.file_uploader("📘 Upload Ledger Master (Excel or CSV)", type=["xlsx","csv"])
    header_text = st.text_input("📝 PDF Header Names (comma-separated):", 
                                value="DATE, NARRATION, CHQ./REF.NO., VALUE DT, WITHDRAWAL AMT., DEPOSIT AMT., CLOSING BALANCE")
    fuzzy_thresh = st.slider("🎯 Matching Accuracy Threshold", 50, 100, 75)
    give_bank_name = st.text_input("🏦 Bank Name for Output:", value="HDFC BANK")

    if file is None:
        st.info("⬆️ Upload your statement file to begin.")
        return

    # ---- Read File (Only if new file or not in session) ----
    if st.session_state.df_main is None or st.button("🔄 Reload File"):
        if str(file.name).lower().endswith(".pdf"):
            df = parse_custom_pdf(file, header_text)
            st.success("✅ PDF parsed successfully!")
        else:
            df = pd.read_excel(file, engine="openpyxl")

        df.columns = [str(c).strip().upper() for c in df.columns]
        for col in ["DEBIT","CREDIT"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        # Store original index for maintaining order
        df['ORIGINAL_INDEX'] = df.index
        df["MAPPED_LEDGER"] = ""
        
        # Load Ledger Master
        ledger_list = []
        if ledger_file is not None:
            if str(ledger_file.name).lower().endswith(".csv"):
                df_led = pd.read_csv(ledger_file)
            else:
                df_led = pd.read_excel(ledger_file, engine="openpyxl")
            ledger_list = df_led.iloc[:,0].astype(str).str.strip().tolist()
            ledger_list = [clean_narration(x) for x in ledger_list if x]
            st.success(f"✅ Loaded {len(ledger_list)} ledger names.")

        # Auto BANK CHARGES
        df.loc[(df["DEBIT"] > 0) & (df["DEBIT"] <= 55), "MAPPED_LEDGER"] = "BANK CHARGES"

        # Fuzzy Mapping
        if ledger_list:
            for idx in df.index:
                if df.at[idx, "MAPPED_LEDGER"] == "":
                    narr = clean_narration(str(df.at[idx, "NARRITION"]))
                    if narr:
                        best = process.extractOne(narr, ledger_list, scorer=fuzz.token_set_ratio)
                        if best and best[1] >= fuzzy_thresh:
                            df.at[idx, "MAPPED_LEDGER"] = best[0]
        
        st.session_state.df_main = df
        st.session_state.mapping_done = False

    df = st.session_state.df_main

    st.subheader("📄 Statement Preview")
    st.dataframe(df.head(10))

    # ---- SUSPENSE Helper Section ----
    suspense_df = df[df["MAPPED_LEDGER"] == ""].copy()
    
    if len(suspense_df) > 0:
        st.markdown(f"""
        <div class="suspense-header">
            <h2>🚨 SUSPENSE Helper</h2>
            <p>{len(suspense_df)} transactions need mapping</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Initialize session state for view mode
        if 'suspense_view' not in st.session_state:
            st.session_state.suspense_view = None
        
        # View Selection
        st.subheader("📋 Choose Your View:")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📋 Group Similar Narrations", use_container_width=True):
                st.session_state.suspense_view = "group"
                
        with col2:
            if st.button("💰 Sort by Amount (High→Low)", use_container_width=True):
                st.session_state.suspense_view = "amount"
                
        with col3:
            if st.button("🎨 Group Together (Color Blocks)", use_container_width=True):
                st.session_state.suspense_view = "color"
        
        st.divider()
        
        # Display based on selected view
        if st.session_state.suspense_view == "group":
            st.subheader("📋 Grouped by Similar Narrations")
            
            groups = group_similar_narrations(suspense_df)
            
            for group_idx, (narr_key, indices) in enumerate(groups.items()):
                group_data = suspense_df.loc[indices]
                total_amount = group_data['DEBIT'].sum() + group_data['CREDIT'].sum()
                
                with st.expander(f"🔵 Group {group_idx + 1}: {narr_key[:50]}... ({len(indices)} items - ₹{total_amount:,.2f})", expanded=True):
                    st.dataframe(group_data[['DATE', 'NARRITION', 'DEBIT', 'CREDIT']], use_container_width=True)
                    
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        ledger_name = st.text_input(
                            f"Map all {len(indices)} items to:",
                            key=f"group_{group_idx}",
                            placeholder="Enter ledger name (e.g., UPI PAYMENTS)..."
                        )
                    with col2:
                        st.write("")
                        st.write("")
                        if st.button("✅ Apply", key=f"apply_group_{group_idx}"):
                            if ledger_name.strip():
                                for idx in indices:
                                    st.session_state.df_main.at[idx, "MAPPED_LEDGER"] = ledger_name.strip().upper()
                                st.success(f"✅ Mapped {len(indices)} items to {ledger_name.strip().upper()}!")
                                st.session_state.mapping_done = True
                                st.rerun()
        
        elif st.session_state.suspense_view == "amount":
            st.subheader("💰 Sorted by Amount (High to Low)")
            
            suspense_sorted = suspense_df.copy()
            suspense_sorted['TOTAL_AMT'] = suspense_sorted['DEBIT'] + suspense_sorted['CREDIT']
            suspense_sorted = suspense_sorted.sort_values('TOTAL_AMT', ascending=False)
            
            for idx, row in suspense_sorted.iterrows():
                amount = row['DEBIT'] if row['DEBIT'] > 0 else row['CREDIT']
                category, color = get_amount_category(amount)
                
                st.markdown(f"""
                <div class="group-card">
                    <strong style="color: {color};">₹{amount:,.2f}</strong> | 
                    {row['DATE']} | {row['NARRITION'][:80]}
                </div>
                """, unsafe_allow_html=True)
                
                col1, col2 = st.columns([3, 1])
                with col1:
                    ledger_name = st.text_input(
                        "Map to:",
                        key=f"amount_{idx}",
                        placeholder="Enter ledger name..."
                    )
                with col2:
                    st.write("")
                    st.write("")
                    if st.button("✅", key=f"apply_amount_{idx}"):
                        if ledger_name.strip():
                            st.session_state.df_main.at[idx, "MAPPED_LEDGER"] = ledger_name.strip().upper()
                            st.success(f"✅ Mapped to {ledger_name.strip().upper()}!")
                            st.session_state.mapping_done = True
                            st.rerun()
        
        elif st.session_state.suspense_view == "color":
            st.subheader("🎨 Color-Coded Groups")
            
            groups = group_similar_narrations(suspense_df)
            colors = ['#3498db', '#2ecc71', '#e67e22', '#e74c3c', '#9b59b6', '#1abc9c']
            
            for group_idx, (narr_key, indices) in enumerate(groups.items()):
                color = colors[group_idx % len(colors)]
                group_data = suspense_df.loc[indices]
                total_amount = group_data['DEBIT'].sum() + group_data['CREDIT'].sum()
                
                st.markdown(f"""
                <div style="background: {color}20; padding: 15px; border-radius: 10px; margin: 15px 0; border-left: 5px solid {color};">
                    <h4 style="color: {color}; margin: 0;">🎨 Group {group_idx + 1}</h4>
                    <p style="margin: 5px 0;"><strong>{len(indices)} transactions</strong> | Total: ₹{total_amount:,.2f}</p>
                </div>
                """, unsafe_allow_html=True)
                
                for idx in indices:
                    row = suspense_df.loc[idx]
                    amount = row['DEBIT'] if row['DEBIT'] > 0 else row['CREDIT']
                    
                    st.markdown(f"""
                    <div class="group-card" style="border-left-color: {color}; background: {color}10;">
                        <strong>₹{amount:,.2f}</strong> | {row['DATE']} | {row['NARRITION'][:80]}
                    </div>
                    """, unsafe_allow_html=True)
                
                col1, col2 = st.columns([3, 1])
                with col1:
                    ledger_name = st.text_input(
                        f"Map entire group ({len(indices)} items) to:",
                        key=f"color_group_{group_idx}",
                        placeholder="Enter ledger name..."
                    )
                with col2:
                    st.write("")
                    st.write("")
                    if st.button("✅ Apply", key=f"apply_color_group_{group_idx}"):
                        if ledger_name.strip():
                            for idx in indices:
                                st.session_state.df_main.at[idx, "MAPPED_LEDGER"] = ledger_name.strip().upper()
                            st.success(f"✅ Mapped {len(indices)} items to {ledger_name.strip().upper()}!")
                            st.session_state.mapping_done = True
                            st.rerun()
                
                st.divider()
        
        else:
            st.info("👆 Select a view option above to start mapping SUSPENSE items")

    else:
        st.markdown("""
        <div class="success-box">
            <h3>🎉 All Transactions Mapped Successfully!</h3>
            <p>No SUSPENSE items remaining. You can now download the final Excel.</p>
        </div>
        """, unsafe_allow_html=True)

    # ---- Build Final Tally Format ----
    st.divider()
    st.subheader("📊 Final Tally Import Format")
    
    # Sort by original index to maintain bank statement order
    df_final = st.session_state.df_main.sort_values('ORIGINAL_INDEX').copy()
    
    out_rows = []
    for _, row in df_final.iterrows():
        date_str = to_date_str(row["DATE"])
        debit, credit = row["DEBIT"], row["CREDIT"]
        amount = debit if debit > 0 else credit
        mapped = row["MAPPED_LEDGER"] if row["MAPPED_LEDGER"] else "SUSPENSE"
        narration = str(row["NARRITION"])

        if debit > 0 and credit == 0:
            vtype = "Payment"
            by_dr = mapped
            to_cr = give_bank_name
        else:
            vtype = "Receipt"
            by_dr = give_bank_name
            to_cr = mapped

        out_rows.append({
            "DATE": date_str,
            "VOUCHER NO.": "",
            "BY / DR": by_dr,
            "TO / CR": to_cr,
            "AMOUNT": amount,
            "NARRATION": narration,
            "VOUCHER TYPE": vtype,
            "DAY": date_str[:2] if date_str else ""
        })

    out_df = pd.DataFrame(out_rows)
    
    # Show preview
    st.dataframe(out_df.head(20), use_container_width=True)

    # ---- Summary Totals ----
    total_payment = df_final["DEBIT"].sum()
    total_receipt = df_final["CREDIT"].sum()
    diff = total_receipt - total_payment
    suspense_count = len(out_df[out_df["TO / CR"] == "SUSPENSE"]) + len(out_df[out_df["BY / DR"] == "SUSPENSE"])

    st.subheader("📊 Summary")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💸 Total Payment", f"₹{total_payment:,.2f}")
    col2.metric("💰 Total Receipt", f"₹{total_receipt:,.2f}")
    col3.metric("⚖️ Difference", f"₹{diff:,.2f}")
    col4.metric("🚨 SUSPENSE Items", suspense_count)

    # ---- Download Excel ----
    st.divider()
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Tally_Import")

    st.download_button(
        "⬇️ Download Final Excel for Tally",
        buffer.getvalue(),
        "tally_mapped_output.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    if suspense_count > 0:
        st.warning(f"⚠️ Note: {suspense_count} transactions are still marked as SUSPENSE. Map them above before importing to Tally.")
    else:
        st.success("✅ Perfect! All transactions mapped. Ready for Tally import!")

if __name__ == "__main__":
    main()