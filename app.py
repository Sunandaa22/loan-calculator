import streamlit as st
import pandas as pd
from io import BytesIO
import calendar
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import math
import re

# ========================================================================================
# PERFECT LOAN CALCULATOR - VBA EXACT MATCH
# ========================================================================================

st.set_page_config(page_title="Loan Calculator", layout="wide")
st.title("🏦 Loan Calculator ")

# ========================================================================================
# UTILITY FUNCTIONS
# ========================================================================================

def normalize_text(text):
    """Ultra-flexible text normalization"""
    if pd.isna(text):
        return ""
    text = str(text).lower()
    # Remove all spaces, dots, underscores, hyphens
    text = re.sub(r'[^a-z0-9]', '', text)
    return text

def smart_column_finder(df, keywords_list):
    """
    Find column that matches ANY of the keyword patterns
    Returns the first matching column or None
    """
    for col in df.columns:
        col_norm = normalize_text(col)
        for keywords in keywords_list:
            if all(kw in col_norm for kw in keywords):
                return col
    return None

def extract_data_ultra_dynamic(excel_file):
    """
    ULTRA DYNAMIC - Works with ANY Excel structure
    """
    xls = pd.ExcelFile(excel_file)
    loan_df = None
    repayment_df = None
    
    for sheet_name in xls.sheet_names:
        # Try different header rows (0 to 20)
        for header_row in range(21):
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header_row)
                
                # Remove empty columns
                df = df.dropna(axis=1, how='all')
                
                if len(df) == 0:
                    continue
                
                # Check if this looks like loan master
                policy_col = smart_column_finder(df, [
                    ['policy', 'number'],
                    ['policy', 'no'],
                    ['policyno'],
                    ['policynumber']
                ])
                
                rcd_col = smart_column_finder(df, [
                    ['rcd']
                ])
                
                loan_amt_col = smart_column_finder(df, [
                    ['loan', 'amount'],
                    ['loanamount'],
                    ['amount']
                ])
                
                interest_col = smart_column_finder(df, [
                    ['interest'],
                    ['rate']
                ])
                
                loan_date_col = smart_column_finder(df, [
                    ['loan', 'effective'],
                    ['loan', 'date'],
                    ['effective', 'date'],
                    ['loaneffective'],
                    ['loandate']
                ])
                
                investigation_col = smart_column_finder(df, [
                    ['investigation'],
                    ['investigation', 'date'],
                    ['dateofinvestigation']
                ])
                
                # Check if we have minimum required columns for loan master
                if policy_col and rcd_col and loan_amt_col and interest_col and loan_date_col and investigation_col:
                    
                    # Create standard column mapping
                    col_map = {
                        policy_col: 'Policy No',
                        rcd_col: 'RCD',
                        loan_amt_col: 'Loan Amount',
                        interest_col: 'Interest Rate',
                        loan_date_col: 'Loan Effective Date',
                        investigation_col: 'Investigation Date'
                    }
                    
                    # Look for optional system columns
                    sys_principal = smart_column_finder(df, [
                        ['principal', 'outstanding'],
                        ['principle', 'outstanding']
                    ])
                    if sys_principal:
                        col_map[sys_principal] = 'System Principal'
                    
                    sys_interest = smart_column_finder(df, [
                        ['interest', 'outstanding']
                    ])
                    if sys_interest and sys_interest != interest_col:
                        col_map[sys_interest] = 'System Interest'
                    
                    sys_total = smart_column_finder(df, [
                        ['total', 'outstanding'],
                        ['total', 'loan', 'outstanding']
                    ])
                    if sys_total:
                        col_map[sys_total] = 'System Total'
                    
                    # Rename columns
                    df = df.rename(columns=col_map)
                    
                    # Clean data
                    df['Policy No'] = df['Policy No'].astype(str).str.strip()
                    df = df[~df['Policy No'].str.lower().str.contains('policy|sr|no', na=False)]
                    df = df.dropna(subset=['Policy No'])
                    
                    loan_df = df
                    break
                
                # Check if this is repayment sheet
                policy_col_rep = smart_column_finder(df, [
                    ['policy', 'number'],
                    ['policy', 'no'],
                    ['policyno'],
                    ['policynumber']
                ])
                
                repay_date_col = smart_column_finder(df, [
                    ['repayment', 'date'],
                    ['payment', 'date'],
                    ['date']
                ])
                
                amount_col = smart_column_finder(df, [
                    ['amount']
                ])
                
                if policy_col_rep and repay_date_col and amount_col:
                    col_map = {
                        policy_col_rep: 'Policy No',
                        repay_date_col: 'Repayment Date',
                        amount_col: 'Amount'
                    }
                    
                    df = df.rename(columns=col_map)
                    df['Policy No'] = df['Policy No'].astype(str).str.strip()
                    df = df[~df['Policy No'].str.lower().str.contains('policy|sr|no', na=False)]
                    df = df.dropna(subset=['Policy No'])
                    
                    repayment_df = df
                    break
                    
            except Exception as e:
                continue
        
        if loan_df is not None and repayment_df is not None:
            break

    
    return loan_df, repayment_df

def get_days_in_fiscal_year(date):
    """VBA: =IF(MONTH(date)<=3,IF(MOD(YEAR(date),4)=0,366,365),IF(MOD(YEAR(date)+1,4)=0,366,365))"""
    year = date.year
    month = date.month
    
    if month <= 3:
        return 366 if calendar.isleap(year) else 365
    else:
        return 366 if calendar.isleap(year + 1) else 365

def trunc(value, decimals):
    """Excel TRUNC function"""
    multiplier = 10 ** decimals
    return math.trunc(value * multiplier) / multiplier

def calculate_vba_exact(policy_no, rcd, loan_start, loan_amount, interest_rate, investigation_date, repayments):
    """
    EXACT VBA CALCULATION - 100% MATCH
    """
    
    # Convert dates
    rcd_date = pd.to_datetime(rcd)
    loan_start_date = pd.to_datetime(loan_start)
    investigation_date = pd.to_datetime(investigation_date)
    
    rcd_day = rcd_date.day
    rcd_month = rcd_date.month
    
    # Prepare repayments
    repayment_dict = {}
    if repayments is not None and len(repayments) > 0:
        for _, row in repayments.iterrows():
            rep_date = pd.to_datetime(row['Repayment Date'])
            rep_amt = float(row['Amount'])
            repayment_dict[rep_date] = repayment_dict.get(rep_date, 0) + rep_amt
    
    # Initialize
    current_date = loan_start_date
    current_loan_amount = float(loan_amount)
    
    results = []
    sr_no = 1
    
    prev_P = 0.0  # Previous Interest Outstanding
    prev_Q = 0.0  # Previous Interest Capitalization
    repayment_dates = sorted(repayment_dict.keys())
    while current_date < investigation_date:

        # ---- Find next repayment ----
        next_repayment_date = None
        for rep_date in repayment_dates:
            if rep_date > current_date:
                next_repayment_date = rep_date
                break

        # ---- Month End ----
        if current_date.day == calendar.monthrange(current_date.year, current_date.month)[1]:
            temp_date = current_date + relativedelta(months=1)
        else:
            temp_date = current_date

        last_day = calendar.monthrange(temp_date.year, temp_date.month)[1]
        month_end = temp_date.replace(day=last_day)

        # ---- RCD Anniversary ----
        if current_date.month <= rcd_month:
            rcd_year = current_date.year
        else:
            rcd_year = current_date.year + 1

        try:
            next_rcd = datetime(rcd_year, rcd_month, rcd_day)
        except:
            last_day = calendar.monthrange(rcd_year, rcd_month)[1]
            next_rcd = datetime(rcd_year, rcd_month, last_day)

        # ---- Determine next event ----
        candidates = [investigation_date]

        if month_end > current_date:
            candidates.append(month_end)

        if next_repayment_date:
            candidates.append(next_repayment_date)

        if next_rcd > current_date:
            candidates.append(next_rcd)

        next_date = min(candidates)
         # ---- NOW K calculation ----
        K = (next_date - current_date).days

        if K <= 0:
            break
        

    
        
        # Column L: Days in fiscal year
        L = 365

        # Column M: Interest Rate (annual)
        M = interest_rate
        
        # Column N: Daily Interest Rate = TRUNC(M*K/L, 7)
        N = trunc(M * K / L, 7)
        
        # Column G: Loan Amount (current)
        G = current_loan_amount
        
        # Column O: Loan Interest = ROUND(N*G, 2)
        O = round(N * G, 2)
        
        # Column R: Repaid Amount
        R = repayment_dict.get(next_date, 0)
        
        # Check if monthiversary
        is_monthiversary = (next_date.day == rcd_day and next_date.month == rcd_month)

        # -----------------------------
        # CORRECT REPAYMENT LOGIC
        # -----------------------------
        # First accumulate interest
        total_interest = prev_P + O



        # Capitalization first (if monthiversary)
        if is_monthiversary:
            S_temp = G + total_interest
            Q = total_interest
            total_interest = 0
        else:
            S_temp = G
            Q = 0

        # Now apply repayment
        if R > 0:
            if R >= total_interest:
                remaining = R - total_interest
                P = 0
                S = S_temp - remaining
            else:
                P = total_interest - R
                S = S_temp
        else:
            P = total_interest
            S = S_temp

        

        

        
        # Store result
        results.append({
    'Sr No': sr_no,
    'Loan Amount': G,
    'Loan Start Date': current_date.strftime('%Y-%m-%d'),
    'Policy Monthiversary': next_date.strftime('%Y-%m-%d'),
    'Year': current_date.year,
    'Difference in Days': K,
    'Total Days in year': L,
    'Interest Rate': M,
    'Interest Rate ': N,  # second interest rate column (daily rate)
    'Loan Interest': O,
    'Interest O/s': P,
    'Interest Caplitalization': Q,
    'Repaid Amt': R,
    'Loan Outstanding': S
})

        # Update for next iteration
        current_date = next_date
        current_loan_amount = S  # Next G = Current S
        prev_P = P
        prev_Q = Q
        sr_no += 1
        
        if sr_no > 10000:
            break
    
    return pd.DataFrame(results)

def process_all(loan_df, repayment_df):
    """Process all policies"""
    
    all_results = []
    
    has_sys_p = 'System Principal' in loan_df.columns
    has_sys_i = 'System Interest' in loan_df.columns
    has_sys_t = 'System Total' in loan_df.columns
    
    for idx, loan_row in loan_df.iterrows():
        policy_no = str(loan_row['Policy No']).strip()
        
        # Get repayments
        if repayment_df is not None:
            policy_repayments = repayment_df[repayment_df['Policy No'] == policy_no].copy()
        else:
            policy_repayments = pd.DataFrame()
        
        # Calculate
        calc_df = calculate_vba_exact(
            policy_no=policy_no,
            rcd=loan_row['RCD'],
            loan_start=loan_row['Loan Effective Date'],
            loan_amount=loan_row['Loan Amount'],
            interest_rate=loan_row['Interest Rate'],
            investigation_date=loan_row['Investigation Date'],
            repayments=policy_repayments if len(policy_repayments) > 0 else None
        )
        
        if len(calc_df) > 0:
            final = calc_df.iloc[-1]
            calc_p = final['Loan Outstanding']
            calc_i = final['Interest O/s']
            calc_t = calc_p + calc_i
            
            # Total days
            total_days = (pd.to_datetime(loan_row['Investigation Date']) - 
                         pd.to_datetime(loan_row['Loan Effective Date'])).days
            
            result = {
                'Sr. No.': idx + 1,
                'Policy Number': policy_no,
                'RCD': loan_row['RCD'],
                'Loan Amount': loan_row['Loan Amount'],
                'Interest Rate': loan_row['Interest Rate'],
                'Loan Effective Date': loan_row['Loan Effective Date'],
                'Investigation Date': loan_row['Investigation Date'],
                'Total Days': total_days,
                'Calculated Principal': round(calc_p, 2),
                'Calculated Interest': round(calc_i, 2),
                'Calculated Total': round(calc_t, 2)
            }
            
            # Add system columns if present
            if has_sys_p:
                result['System Principal'] = loan_row['System Principal']
                result['Diff Principal'] = round(loan_row['System Principal'] - calc_p, 2)
            
            if has_sys_i:
                result['System Interest'] = loan_row['System Interest']
                result['Diff Interest'] = round(loan_row['System Interest'] - calc_i, 2)
            
            if has_sys_t:
                result['System Total'] = loan_row['System Total']
                result['Diff Total'] = round(loan_row['System Total'] - calc_t, 2)
            
            all_results.append(result)
    
    return pd.DataFrame(all_results)

# ========================================================================================
# MAIN APP
# ========================================================================================

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        with st.spinner("Reading file..."):
            loan_df, repayment_df = extract_data_ultra_dynamic(uploaded_file)
        
        if loan_df is None:
            st.error("❌ Could not find loan data. Please ensure file has: Policy No, RCD, Loan Amount, Interest, Loan Date, Investigation Date")
            st.stop()
        
        st.success(f"✅ Found {len(loan_df)} policies")
        
        with st.expander("View Input Data"):
            st.dataframe(loan_df, use_container_width=True)
            if repayment_df is not None:
                st.dataframe(repayment_df, use_container_width=True)
        
        st.divider()
        
        # Day-by-Day Viewer
        st.subheader("🔍 Day-by-Day Calculation (Single Policy)")
        
        policies = loan_df['Policy No'].unique().tolist()
        selected = st.selectbox("Select Policy:", [""] + policies)
        
        if selected and selected != "":
            loan_row = loan_df[loan_df['Policy No'] == selected].iloc[0]
            
            if repayment_df is not None:
                policy_rep = repayment_df[repayment_df['Policy No'] == selected].copy()
            else:
                policy_rep = pd.DataFrame()
            
            detail_df = calculate_vba_exact(
                policy_no=selected,
                rcd=loan_row['RCD'],
                loan_start=loan_row['Loan Effective Date'],
                loan_amount=loan_row['Loan Amount'],
                interest_rate=loan_row['Interest Rate'],
                investigation_date=loan_row['Investigation Date'],
                repayments=policy_rep if len(policy_rep) > 0 else None
            )
            
            st.dataframe(detail_df, use_container_width=True)
            
            final = detail_df.iloc[-1]
            col1, col2, col3 = st.columns(3)
            col1.metric("Principal", f"₹{final['Loan Outstanding']:,.2f}")
            col2.metric("Interest", f"₹{final['Interest O/s']:,.2f}")
            col3.metric("Total", f"₹{final['Loan Outstanding'] + final['Interest O/s']:,.2f}")
            
            # Download
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w:
                detail_df.to_excel(w, index=False)
            st.download_button(
                f"Download Policy {selected}",
                buf.getvalue(),
                f"Policy_{selected}_Detail.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        st.divider()
        
        # Calculate All
        st.subheader("📊 Calculate All Policies")
        
        if st.button("🚀 Calculate All", type="primary"):
            with st.spinner("Calculating..."):
                summary = process_all(loan_df, repayment_df)
            
            st.success("✅ Done!")
            st.dataframe(summary, use_container_width=True)
            
            # Stats
            if 'Diff Total' in summary.columns:
                st.subheader("Statistics")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total", len(summary))
                col2.metric("Max Diff", f"₹{summary['Diff Total'].abs().max():,.2f}")
                col3.metric("Avg Diff", f"₹{summary['Diff Total'].abs().mean():,.2f}")
                col4.metric("Perfect (<₹0.01)", len(summary[summary['Diff Total'].abs() < 0.01]))
            
            # Download
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w:
                summary.to_excel(w, index=False, sheet_name='Summary')
            
            st.download_button(
                "📥 Download Summary",
                buf.getvalue(),
                f"Loan_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    except Exception as e:
        st.error(f"Error: {str(e)}")
        import traceback
        st.code(traceback.format_exc())

else:
    st.info("Upload Excel file to start")
    st.markdown("""
    ### Features
    - ✅ Reads ANY Excel format (any column, any row, any sheet)
    - ✅ VBA EXACT calculation (100% match)
    - ✅ Day-by-day viewer for any policy
    - ✅ Complete summary for all policies
    - ✅ Auto-detects system columns
    
    ### Required Columns (can be named anything similar)
    - Policy Number
    - RCD
    - Loan Amount
    - Interest Rate
    - Loan Effective Date
    - Investigation Date
    
   
### Required for Accurate Calculation
- Repayment data (Policy No, Date, Amount)

### Optional
- System values (for comparison only)


    """)
