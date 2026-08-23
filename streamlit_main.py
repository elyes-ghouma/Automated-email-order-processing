import os
import pandas as pd
import streamlit as st
from functions_real import *

path="C:/Users/ELITE/Desktop/Work/Vectorys Planification"

# --- Page Configuration ---
st.set_page_config(
    page_title="Email & Order Management",
    page_icon="📦",
    layout="wide"
)

# Initialize persistent session states
if "orders_db" not in st.session_state:
    st.session_state["orders_db"] = []

# List of company dictionary objects conforming strictly to your schema
if "clients_db" not in st.session_state:
    st.session_state["clients_db"] = []

# ==========================================
# SIDEBAR: Work Folder, Credentials & Execution
# ==========================================
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # 1. Workspace Folder Selection
    st.subheader("1. Working Directory")
    default_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Vectorys Planification")
    work_folder = st.text_input("📁 Work Folder Path", value=default_dir)
    
    if work_folder and not os.path.exists(work_folder):
        os.makedirs(work_folder, exist_ok=True)
        st.success("Work folder created!")

    st.markdown("---")
    
    # 2. IMAP Credentials
    st.subheader("2. IMAP Account")
    email_user = st.text_input("Email Address", value="")
    email_pass = st.text_input("Password / App Password", type="password")
    
    st.markdown("---")
    
    # 3. Execution Button
    btn_analyse = st.button("🚀 Analyse Mails", type="primary", use_container_width=True)

st.title("📦 Email & Order Processing Hub")

# Trigger IMAP & Ollama pipeline when button is clicked
if btn_analyse:
    if not email_user or not email_pass:
        st.error("Please fill in your Email and Password in the sidebar before analyzing.")
    else:
        with st.spinner("Connecting to inbox, downloading attachments, and analyzing mails..."):
            try:
                # -------------------------------------------------------------
                # CALL YOUR EXISTING PIPELINE HERE:
                # e.g., analyze_all_mails(email_user, email_pass, work_folder)
                # -------------------------------------------------------------
                
                # Mock result payload for UI testing
                st.session_state["orders_db"] = [
                    {
                        "local_mail_id": "101",
                        "sender": "client1@domain.com",
                        "subject": "Commande Urgente",
                        "is_order": True,
                        "client": "Client A",
                        "stockeur": "CO3 CMS",
                        "status": "Ready",
                        "attachment": "order_101.xlsm"
                    },
                    {
                        "local_mail_id": "102",
                        "sender": "newsletter@domain.com",
                        "subject": "Monthly Updates",
                        "is_order": False,
                        "client": "N/A",
                        "stockeur": "N/A",
                        "status": "Non-Order",
                        "attachment": "None"
                    }
                ]
                st.success("Analysis complete!")
            except Exception as e:
                st.error(f"Error during execution: {e}")

# ==========================================
# MAIN SECTIONS (TABS)
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs([
    "👥 Add Client / Add Stocker",
    "📊 Show Orders",
    "📧 Show Non Orders",
    "📩 Show Mails"
])

# ------------------------------------------
# 1. ADD CLIENT / ADD STOCKER SECTION
# ------------------------------------------
with tab1:
    st.header("Manage Clients & Stocker Locations")
    
    col_client, col_stocker = st.columns(2)
    
    # --- Form 1: Add New Client (Strict dict structure) ---
    with col_client:
        st.subheader("Add New Client")
        with st.form("add_client_form", clear_on_submit=True):
            company_name = st.text_input("Company Name *").upper()
            st.markdown("**Initial Stocker Details:**")
            stockeur_name = st.text_input("Stockeur * (e.g., CO3 CMS)").upper()
            num_stockeur = st.text_input("Num Stockeur *").upper()
            num_dep = st.text_input("Num Département *").upper()
            ville = st.text_input("Ville *").upper()
            pays = st.text_input("Pays *").upper()
            
            submit_client = st.form_submit_button("Add Client")
            
            if submit_client:
                if not (company_name and stockeur_name and num_stockeur and num_dep and ville and pays):
                    st.error("All fields are mandatory!")
                elif any(c["company"] == company_name for c in read_jsonl_file(path+"/Clients.jsonl")):
                    st.warning(f"Company '{company_name}' already exists.")
                else:
                    new_client = {
                        "company": company_name,
                        "stockeurs": [
                            {
                                "stockeur": stockeur_name,
                                "num_stockeur": num_stockeur,
                                "num_dep": num_dep,
                                "ville": ville,
                                "pays": pays
                            }
                        ]
                    }
                    add_client(path,new_client)
                    st.success(f"Company '{company_name}' created with stockeur '{stockeur_name}'!")

    # --- Form 2: Add Additional Stocker to Existing Company ---
    with col_stocker:
        st.subheader("Add Additional Stocker")
        company_names = [c["company"] for c in read_jsonl_file(path+"/Clients.jsonl")]
        if company_names:
            with st.form("add_stocker_form", clear_on_submit=True):
                selected_company = st.selectbox("Select Company", company_names).upper()
                st.markdown("**New Stocker Details:**")
                extra_stockeur = st.text_input("Stockeur * (e.g., Dépôt Nord)").upper()
                extra_num_stockeur = st.text_input("Num Stockeur *").upper()
                extra_num_dep = st.text_input("Num Département *").upper()
                extra_ville = st.text_input("Ville *").upper()
                extra_pays = st.text_input("Pays *").upper()
                
                submit_stocker = st.form_submit_button("Add Stocker")
                
                if submit_stocker:
                    if not (extra_stockeur and extra_num_stockeur and extra_num_dep and extra_ville and extra_pays):
                        st.error("All Stocker fields are mandatory!")
                    else:
                        # Find targeted company entry
                        target_entry = next(c for c in read_jsonl_file(path+"/Clients.jsonl") if c["company"] == selected_company)
                        
                        if any(s["stockeur"] == extra_stockeur for s in target_entry["stockeurs"]):
                            st.warning("A stockeur with this name already exists for this company.")
                        else:
                            new_stockeur_obj = {
                                "stockeur": extra_stockeur,
                                "num_stockeur": extra_num_stockeur,
                                "num_dep": extra_num_dep,
                                "ville": extra_ville,
                                "pays": extra_pays
                            }
                            target_entry["stockeurs"].append(new_stockeur_obj)
                            add_stocker(target_entry["company"],path,new_stockeur_obj)
                            st.success(f"Stockeur '{extra_stockeur}' added to {selected_company}!")
        else:
            st.info("No companies registered yet. Add a client first.")

    st.markdown("---")

# ------------------------------------------
# 2. SHOW ORDERS SECTION
# ------------------------------------------
with tab2:
    st.header("Validated Order Requests")
    
    orders = [m for m in st.session_state["orders_db"] if m.get("is_order") is True]
    
    if orders:
        df_orders = pd.DataFrame(orders)
        st.dataframe(df_orders, use_container_width=True)
    else:
        st.info("No valid orders found yet. Click 'Analyse Mails' to process your inbox.")

# ------------------------------------------
# 3. SHOW NON ORDERS SECTION
# ------------------------------------------
with tab3:
    st.header("Non-Order Emails")
    
    non_orders = [m for m in st.session_state["orders_db"] if m.get("is_order") is False]
    
    if non_orders:
        df_non_orders = pd.DataFrame(non_orders)
        st.dataframe(df_non_orders, use_container_width=True)
    else:
        st.info("No non-order emails found.")

# ------------------------------------------
# 4. SHOW MAILS SECTION
# ------------------------------------------
with tab4:
    st.header("All Processed Emails")
    
    if st.session_state["orders_db"]:
        df_all = pd.DataFrame(st.session_state["orders_db"])
        st.dataframe(df_all, use_container_width=True)
    else:
        st.info("No email records loaded.")