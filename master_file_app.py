import streamlit as st
import pandas as pd
import io
import os
import json
from openpyxl import load_workbook
from openpyxl.styles import Protection, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# --- Configuration ---
st.set_page_config(page_title="UCES Margin Analyzer", layout="wide")

# --- CSS Styling ---
st.markdown("""
<style>
    .stApp { background-color: #f5f5f5; }
    h1 { color: #2e4053; }
    .info-box {
        background-color: #e2e3e5;
        border: 1px solid #d6d8db;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 15px;
        font-size: 0.9em;
    }
    div[data-testid="stExpander"] {
        background-color: white;
        border-radius: 10px;
        border: 1px solid #ddd;
        border-left: 5px solid #2e4053;
    }
    .filter-container {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #ddd;
        margin-bottom: 20px;
    }
    .item-card {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 5px;
        padding: 8px;
        margin-bottom: 5px;
        font-size: 0.9em;
    }
</style>
""", unsafe_allow_html=True)

# --- PERSISTENCE FUNCTIONS (SAFE MODE) ---
DATA_FILE = os.path.join(os.getcwd(), "uces_app_data.json")

def save_data():
    """Saves current state to JSON file."""
    data_to_save = {
        'df': st.session_state.df.to_dict(orient='records'),
        'line_items_db': st.session_state.line_items_db,
        'source_filename': st.session_state.source_filename
    }
    try:
        with open(DATA_FILE, "w") as f:
            json.dump(data_to_save, f)
    except Exception as e:
        # SAFE MODE: If we can't save (e.g. on GitHub server), just ignore it.
        # Don't show error to user to avoid confusion.
        pass

def load_data():
    """Loads data from JSON file if it exists."""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r") as f:
                data = json.load(f)
                df = pd.DataFrame(data['df'])
                return {
                    'df': df,
                    'line_items_db': data['line_items_db'],
                    'source_filename': data['source_filename']
                }
        except Exception:
            return None
    return None

# --- Helper Function: Initialize DF ---
def init_df():
    return pd.DataFrame(columns=[
        "Quotation No", "Po Huawei", "Linked PR Subcon", "Vendor Name", "Project", "Site ID", "Line Items",
        "Po Huawei (Unit Price)", "Requested Qty", "Total", 
        "Po Subcon (Unit Price)", "Qty", "Sub Total", "Profit", "Margin%", "Status"
    ])

# --- Initialize Session State ---
if 'df' not in st.session_state:
    saved_data = load_data()
    if saved_data:
        st.session_state.df = saved_data['df']
        st.session_state.line_items_db = saved_data['line_items_db']
        st.session_state.source_filename = saved_data['source_filename']
    else:
        st.session_state.df = init_df()
        st.session_state.line_items_db = [
            "Router X5", "Cable 10m", "Antenna Panel", "SIM Card", "Power Adapter"
        ]
        st.session_state.source_filename = "master_file_data.xlsx"

# Define Project Options
PROJECT_OPTIONS = {
    "---": "(NON-PROJECT)",
    "BD": "Business Development",
    "CME": "Civil Mechanical Electrical",
    "CS": "Customer Support",
    "HQ": "Head Quarter",
    "IBS": "Inbuilding System",
    "MISC": "Miscellaneous Project",
    "MS": "Managing Services",
    "RNO": "Radio Network Optimization",
    "SOLAR": "Solar",
    "TI": "Technical Installation",
    "TINSOL": "Tinno Solar"
}
PROJECT_KEYS = list(PROJECT_OPTIONS.keys())

# --- App Title ---
st.title("📊 UCES Margin Analyzer")
st.markdown("Check Client PO vs Subcon PO Margins (Target: ≥30%)")

# --- Section: Settings (Reset) ---
with st.expander("⚙️ System Settings", expanded=False):
    st.write("Manage local application data.")
    if st.button("🗑️ Reset App Data (Clear Cache)", type="secondary"):
        if os.path.exists(DATA_FILE):
            os.remove(DATA_FILE)
            st.success("App data cleared! Refresh page to start fresh.")
            st.rerun()
    st.caption("Clears the saved Line Items and current table state.")

# --- Section: Manage Line Items ---
with st.expander("🛠️ Manage Line Items (Database)", expanded=False):
    st.write("Add standard items here. They will appear in the dropdown when adding new entries.")
    
    m_col1, m_col2 = st.columns([2, 1])
    with m_col1:
        new_item = st.text_input("New Line Item Name", placeholder="e.g., Router X5")
    with m_col2:
        if st.button("➕ Add to Database", use_container_width=True):
            if new_item and new_item not in st.session_state.line_items_db:
                st.session_state.line_items_db.append(new_item)
                save_data() 
                st.success(f"'{new_item}' added!")
                st.rerun()
            elif new_item in st.session_state.line_items_db:
                st.warning("Item already exists!")
    
    st.divider()
    st.write("**Current Database:**")
    
    items_grid_cols = st.columns(4)
    for i, item in enumerate(st.session_state.line_items_db):
        with items_grid_cols[i % 4]:
            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.markdown(f"<div class='item-card'>{item}</div>", unsafe_allow_html=True)
            with d_col2:
                if st.button("🗑️", key=f"del_item_{i}", help="Delete this item"):
                    st.session_state.line_items_db.pop(i)
                    save_data() 
                    st.rerun()
    
    if len(st.session_state.line_items_db) == 0:
        st.info("Database is empty.")

# --- Section 1: Upload Excel ---
with st.expander("📥 Step 1: Load Existing Excel File", expanded=False):
    uploaded_file = st.file_uploader("Choose your Master Excel file", type=['xlsx', 'xls'], key="file_uploader")
    
    if uploaded_file is not None:
        st.info(f"File selected: `{uploaded_file.name}`. Click Confirm to Load.")
        if st.button("🔄 Load File", type="primary", key="load_btn"):
            try:
                temp_df = pd.read_excel(uploaded_file)
                required = init_df().columns.tolist()
                uploaded_cols = temp_df.columns.tolist()
                
                required_lower = [x.lower() for x in required]
                uploaded_cols_lower = [x.lower() for x in uploaded_cols]
                cols_match = all(item in uploaded_cols_lower for item in required_lower)
                
                if cols_match and len(temp_df) > 0:
                    col_map = {x.lower(): x for x in required}
                    temp_df.columns = [col_map.get(str(c).lower(), c) for c in temp_df.columns]
                    
                    st.session_state.df = temp_df[required]
                    st.session_state.source_filename = uploaded_file.name
                    save_data() 
                    st.success(f"Loaded {len(temp_df)} rows from {uploaded_file.name}!")
                    st.rerun()
                else:
                    st.markdown(f"<div class='info-box'>⚠️ <b>Column Mismatch.</b> Please map your columns.</div>", unsafe_allow_html=True)
                    
                    mapping = {}
                    cols_with_none = ["(Ignore/Missing)"] + uploaded_cols
                    st.write("Map columns:")
                    map_cols = st.columns(2)
                    
                    for i, col in enumerate(required):
                        default_idx = 0
                        simple_col = col.replace(" ", "").replace("(", "").replace(")", "").lower()
                        for j, up_col in enumerate(uploaded_cols):
                            simple_up = up_col.replace(" ", "").replace("(", "").replace(")", "").lower()
                            if simple_col == simple_up:
                                default_idx = j + 1
                                break
                        
                        with map_cols[i % 2]:
                            selected = st.selectbox(f"System: `{col}`", cols_with_none, index=default_idx, key=f"map_{col}")
                            mapping[col] = selected if selected != "(Ignore/Missing)" else None
                    
                    if st.button("Apply Mapping & Load", type="secondary", key="map_load_btn"):
                        final_df = pd.DataFrame()
                        for app_col, excel_col in mapping.items():
                            if excel_col is not None:
                                final_df[app_col] = temp_df[excel_col]
                            else:
                                if app_col == "Status": final_df[app_col] = "Process"
                                elif app_col == "Project": final_df[app_col] = "---"
                                elif app_col in ["Po Huawei (Unit Price)", "Po Subcon (Unit Price)", "Total", "Sub Total", "Profit", "Margin%"]: final_df[app_col] = 0.0
                                elif app_col in ["Requested Qty", "Qty"]: final_df[app_col] = 0
                                else: final_df[app_col] = ""
                        st.session_state.df = final_df
                        st.session_state.source_filename = uploaded_file.name
                        save_data() 
                        st.success("Loaded mapped data!")
                        st.rerun()

            except Exception as e:
                st.error(f"Error: {e}")

st.divider()

# --- Section 2: Form (Add/Edit) ---
if 'edit_index' not in st.session_state:
    st.session_state.edit_index = None

form_mode = "✏️ Edit Entry" if st.session_state.edit_index is not None else "➕ Add New Entry"

with st.expander(f"{form_mode}", expanded=(st.session_state.edit_index is not None)):
    
    if st.session_state.edit_index is not None:
        row_data = st.session_state.df.iloc[st.session_state.edit_index]
        current_project = row_data["Project"]
        if current_project in PROJECT_KEYS:
            proj_default = current_project
        else:
            proj_default = "---" 

        default_vals = {
            "quotation_no": row_data["Quotation No"],
            "po_huawei": row_data["Po Huawei"],
            "linked_pr": row_data["Linked PR Subcon"],
            "vendor_name": row_data["Vendor Name"],
            "project": proj_default,
            "site_id": row_data["Site ID"],
            "line_items": row_data["Line Items"],
            "status": row_data["Status"],
            "hp": float(row_data["Po Huawei (Unit Price)"]),
            "rq": int(row_data["Requested Qty"]),
            "sp": float(row_data["Po Subcon (Unit Price)"]),
            "sq": int(row_data["Qty"]),
        }
        btn_text = "💾 Update Row"
        btn_type = "primary"
    else:
        default_vals = {
            "quotation_no": "",
            "po_huawei": "", "linked_pr": "", "vendor_name": "",
            "project": "---", "site_id": "", "line_items": "", 
            "status": "Process",
            "hp": 0.0, "rq": 0, "sp": 0.0, "sq": 0
        }
        btn_text = "➕ Add to Table"
        btn_type = "secondary"

    with st.form("data_form", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            quotation_no = st.text_input("Quotation No", value=default_vals["quotation_no"])
            
            po_huawei = st.text_input("PO Huawei*", value=default_vals["po_huawei"])
            linked_pr = st.text_input("Linked PR Subcon", value=default_vals["linked_pr"])
            
            project = st.selectbox("Project", PROJECT_KEYS, 
                                  index=PROJECT_KEYS.index(default_vals["project"]) if default_vals["project"] in PROJECT_KEYS else 0,
                                  format_func=lambda x: f"{x} = {PROJECT_OPTIONS[x]}")
        
        with col2:
            vendor_name = st.text_input("Vendor Name", value=default_vals["vendor_name"])
            
            site_id = st.text_input("Site ID", value=default_vals["site_id"])
            
            line_item_options = ["(Custom...)"] + st.session_state.line_items_db
            
            default_li_index = 0
            try:
                if default_vals["line_items"] in line_item_options:
                    default_li_index = line_item_options.index(default_vals["line_items"])
            except:
                default_li_index = 0

            selected_line_item = st.selectbox("Line Items", line_item_options, index=default_li_index)
            
            if selected_line_item == "(Custom...)":
                final_line_item = st.text_input("Type Custom Item", value=default_vals["line_items"])
            else:
                final_line_item = selected_line_item
            
            status = st.selectbox("Status", ["Process", "Rejected", "Claimed", "Waiting"], 
                                  index=["Process", "Rejected", "Claimed", "Waiting"].index(default_vals["status"]) if default_vals["status"] in ["Process", "Rejected", "Claimed", "Waiting"] else 0)
        
        with col3:
            st.subheader("Financials")
            po_huawei_price = st.number_input("PO Huawei (Unit Price)", value=default_vals["hp"], format="%.2f")
            req_qty = st.number_input("Requested Qty", value=default_vals["rq"], step=1)
            po_subcon_price = st.number_input("PO Subcon (Unit Price)", value=default_vals["sp"], format="%.2f")
            subcon_qty = st.number_input("Qty (Subcon)", value=default_vals["sq"], step=1)

        submitted = st.form_submit_button(btn_text, type=btn_type)
        cancel_edit = st.form_submit_button("Cancel Edit") if st.session_state.edit_index is not None else None

        if submitted:
            if not po_huawei:
                st.error("PO Huawei is required!")
            else:
                total_huawei = po_huawei_price * req_qty
                total_subcon = po_subcon_price * subcon_qty
                profit = total_huawei - total_subcon
                margin = (profit / total_huawei * 100) if total_huawei != 0 else 0.0

                new_row = {
                    "Quotation No": quotation_no,
                    "Po Huawei": po_huawei,
                    "Linked PR Subcon": linked_pr,
                    "Vendor Name": vendor_name,
                    "Project": project,
                    "Site ID": site_id,
                    "Line Items": final_line_item,
                    "Po Huawei (Unit Price)": po_huawei_price,
                    "Requested Qty": req_qty,
                    "Total": total_huawei,
                    "Po Subcon (Unit Price)": po_subcon_price,
                    "Qty": subcon_qty,
                    "Sub Total": total_subcon,
                    "Profit": profit,
                    "Margin%": round(margin, 2),
                    "Status": status
                }

                if st.session_state.edit_index is not None:
                    st.session_state.df.iloc[st.session_state.edit_index] = new_row
                    st.session_state.edit_index = None
                    st.success("Row updated successfully!")
                else:
                    st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame([new_row])], ignore_index=True)
                    st.success("Entry added successfully!")
                
                save_data() 
                st.rerun()
        
        if cancel_edit:
            st.session_state.edit_index = None
            st.rerun()

# --- Section 3: View Data (WITH FILTERS) ---
st.divider()
st.subheader("Current Master File Data")

if not st.session_state.df.empty:
    
    st.markdown('<div class="filter-container">', unsafe_allow_html=True)
    
    f_col1, f_col2, f_col3, f_col4 = st.columns([1, 1, 1, 1])
    
    with f_col1:
        status_filter = st.selectbox("Filter by Status", ["All", "Process", "Rejected", "Claimed", "Waiting"])
    
    with f_col2:
        search_project = st.text_input("Search Project Code (e.g., BD, SOLAR)")
        
    with f_col3:
        margin_filter = st.selectbox("Filter by Margin", ["All", "High Margin (≥30%)", "Low Margin (<30%)"])
        
    with f_col4:
        if st.button("🔄 Reset Filters"):
            st.rerun()
            
    st.markdown('</div>', unsafe_allow_html=True)

    filtered_df = st.session_state.df.copy()
    
    if status_filter != "All":
        filtered_df = filtered_df[filtered_df["Status"] == status_filter]
    
    if search_project:
        filtered_df = filtered_df[filtered_df["Project"].str.contains(search_project, case=False, na=False)]
    
    if margin_filter == "High Margin (≥30%)":
        filtered_df = filtered_df[filtered_df["Margin%"] >= 30.0]
    elif margin_filter == "Low Margin (<30%)":
        filtered_df = filtered_df[filtered_df["Margin%"] < 30.0]
        
    st.caption(f"Showing {len(filtered_df)} of {len(st.session_state.df)} records")

    def color_margin(val):
        try:
            val_float = float(val)
            if val_float >= 30.0:
                color = 'green'
            else:
                color = 'red'
            return f'color: {color}; font-weight: bold'
        except:
            return 'color: black'

    styled_df = filtered_df.style.applymap(color_margin, subset=['Margin%'])
    
    column_config = {
        "Po Huawei (Unit Price)": st.column_config.NumberColumn("Price (RM)", format="%.2f"),
        "Requested Qty": st.column_config.NumberColumn("Req Qty", format="%d"), 
        "Total": st.column_config.NumberColumn("Total (RM)", format="%.2f"),
        "Po Subcon (Unit Price)": st.column_config.NumberColumn("Sub Price (RM)", format="%.2f"),
        "Qty": st.column_config.NumberColumn("Sub Qty", format="%d"), 
        "Sub Total": st.column_config.NumberColumn("Sub Total (RM)", format="%.2f"),
        "Profit": st.column_config.NumberColumn("Profit (RM)", format="%.2f"),
        "Margin%": st.column_config.NumberColumn("Margin%", format="%.2f"),
        "Project": st.column_config.SelectboxColumn(
            "Project",
            options=PROJECT_KEYS,
            required=True
        )
    }

    st.dataframe(
        styled_df,
        use_container_width=True,
        height=500,
        column_config=column_config
    )

    st.markdown("### Quick Actions")
    action_col1, action_col2, action_col3 = st.columns([2, 1, 1])
    
    with action_col1:
        options = [f"Row {i}: {r['Po Huawei']} - {r['Site ID']}" for i, r in st.session_state.df.iterrows()]
        selected_row_opt = st.selectbox("Select a row to manage:", options, index=None)
        
    with action_col2:
        if selected_row_opt:
            selected_index = int(selected_row_opt.split(":")[0].split(" ")[1])
            if st.button("✏️ Edit Selected", type="primary", use_container_width=True):
                st.session_state.edit_index = selected_index
                st.rerun()
                
    with action_col3:
        if selected_row_opt:
            if st.button("🗑️ Delete Selected", type="secondary", use_container_width=True):
                st.session_state.df = st.session_state.df.drop(selected_index)
                st.session_state.df = st.session_state.df.reset_index(drop=True)
                save_data() 
                st.rerun()

else:
    st.info("No data to display. Upload a file or add a new entry.")

# --- Section 4: Footer / Download ---
st.divider()

st.markdown("### 💾 Save Your Work")
st.info("Download your updated Excel file here. (Note: If running on GitHub, data is temporary and will reset on refresh).")

col_dl, col_clr = st.columns(2)

with col_dl:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state.df.to_excel(writer, index=False, sheet_name='Master File')
        
        workbook = writer.book
        worksheet = workbook['Master File']
        
        green_fill = PatternFill(start_color="C6F6D5", end_color="C6F6D5", fill_type="solid")
        red_fill = PatternFill(start_color="FED7D7", end_color="FED7D7", fill_type="solid")
        
        col_indices = {}
        for col in range(1, worksheet.max_column + 1):
            header_val = worksheet.cell(row=1, column=col).value
            if header_val:
                col_indices[header_val] = col
        
        margin_col_idx = col_indices.get("Margin%")
        
        project_col_letter = None
        for k, v in col_indices.items():
            if k == "Project":
                project_col_letter = get_column_letter(v)
                break
        
        # Excel Validation
        if project_col_letter and worksheet.max_row > 1:
            dv_range = f"{project_col_letter}2:{project_col_letter}{worksheet.max_row}"
            
            dv = DataValidation(type="list", formula1=f'"{",".join(PROJECT_KEYS)}"', allow_blank=True)
            
            dv.error = 'Your entry is not in list'
            dv.errorTitle = 'Invalid Entry'
            dv.prompt = 'Please select from the list'
            dv.promptTitle = 'Project Selection'
            
            dv.add(dv_range)
            worksheet.add_data_validation(dv)

        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.protection = Protection(locked=False)
                
                cell_header = worksheet.cell(row=1, column=cell.column).value
                
                if cell_header in ["Po Huawei (Unit Price)", "Total", "Po Subcon (Unit Price)", "Sub Total", "Profit"]:
                    cell.number_format = '#,##0.00'
                elif cell_header in ["Requested Qty", "Qty"]:
                    cell.number_format = '0'
                elif cell_header == "Margin%":
                    cell.number_format = '0.00'
                    try:
                        val = float(cell.value)
                        if val >= 30.0:
                            cell.fill = green_fill
                        else:
                            cell.fill = red_fill
                    except:
                        pass
                else:
                    cell.number_format = 'General'

        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_letter].width = adjusted_width

    buffer.seek(0)
    
    st.download_button(
        label=f"📥 Update & Download File ({st.session_state.source_filename})",
        data=buffer,
        file_name=st.session_state.source_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col_clr:
    if st.button("🗑️ Clear All Data", type="primary"):
        st.session_state.df = init_df()
        st.session_state.edit_index = None
        st.rerun()