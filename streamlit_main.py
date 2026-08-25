import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import streamlit as st
from functions_real import *

# Helper function to launch native OS Folder Browser
def select_folder():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder_path = filedialog.askdirectory(master=root)
    root.destroy()
    return folder_path

# --- Page Configuration ---
st.set_page_config(
    page_title="Email & Order Management",
    page_icon="📦",
    layout="wide"
)

# Initialize persistent session states
default_dir = os.path.join(os.path.expanduser("~"), "Desktop")
if "work_folder" not in st.session_state:
    st.session_state["work_folder"] = default_dir

# ==========================================
# SIDEBAR: Work Folder, Credentials & Execution
# ==========================================
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # 1. Workspace Folder Selection with Native Picker
    st.subheader("1. Working Directory")
    
    col_path, col_btn = st.columns([3, 1])
    
    with col_path:
        work_folder = st.text_input(
            "📁 Folder Path", 
            value=st.session_state["work_folder"],
            label_visibility="collapsed"
        )
        st.session_state["work_folder"] = work_folder

    with col_btn:
        if st.button("📂 Browse", use_container_width=True):
            selected = select_folder()
            if selected:
                st.session_state["work_folder"] = selected
                st.rerun()

    # Dynamic path update based on active workspace
    path = st.session_state["work_folder"]

    if path and not os.path.exists(path):
        set_main_folder_path(path)
        parametres = read_json_file("parametres.json")
        path = parametres["main_directory_path"]
        st.success("Work folder created!")
    else:
        set_main_folder_path(path)
        parametres = read_json_file("parametres.json")
        path = parametres["main_directory_path"]
        st.caption(f"Active Path: `{path}`")

    st.markdown("---")
    
    # 2. Execution & Stop Controls
    btn_analyse = st.button("🚀 Analyse Mails", type="primary", use_container_width=True)
    
    col_fast, col_slow = st.columns(2)
    with col_fast:
        btn_fast_stop = st.button("⚡ Fast Stop", use_container_width=True)
    with col_slow:
        btn_slow_stop = st.button("🐢 Slow Stop", use_container_width=True)

    # --- Direct JSON File Parameter Control ---
    if btn_fast_stop:
        parametres=read_json_file("parametres.json")
        parametres["fast_stop"]=True
        parametres["slow_stop"]=True
        create_json_file(parametres,"parametres.json")
        st.error("⚡ Immediate Fast Stop requested in config.")

    if btn_slow_stop:
        parametres=read_json_file("parametres.json")
        parametres["slow_stop"]=True
        create_json_file(parametres,"parametres.json")
        st.warning("🐢 Graceful Slow Stop requested in config.")

st.title("📦 Email & Order Processing Hub")

# Trigger IMAP & Ollama pipeline when button is clicked
if btn_analyse:
    with st.spinner("Connecting to inbox, downloading attachments, and analyzing mails..."):
        try:
            parametres=read_json_file("parametres.json")
            parametres["fast_stop"]=False
            parametres["slow_stop"]=False
            create_json_file(parametres,"parametres.json")
            main()
            st.success("Analysis complete!")
        except Exception as e:
            with st.sidebar:
                st.error(f"Error during execution: {e}")

# ==========================================
# MAIN SECTIONS (TABS)
# ==========================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "👥 Add Client / Add Stocker",
    "🏢 Show Clients",
    "📊 Show Orders",
    "📧 Show Non Orders",
    "📩 Show Mails",
    "✏️ Edit Order"
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
                    add_client(path, new_client)
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
                            add_stocker(target_entry["company"], path, new_stockeur_obj)
                            st.success(f"Stockeur '{extra_stockeur}' added to {selected_company}!")
        else:
            st.info("No companies registered yet in the current working directory.")

    st.markdown("---")

# ------------------------------------------
# 2. SHOW CLIENTS SECTION
# ------------------------------------------
with tab2:
    st.header("Registered Clients & Stockers")
    
    clients_file_path = os.path.join(path, "Clients.jsonl")
    if os.path.exists(clients_file_path):
        clients_data = read_jsonl_file(clients_file_path)
        
        if clients_data:
            # Flatten the nested dictionary into a tabular format
            flattened_data = []
            for client in clients_data:
                company = client.get("company", "N/A")
                for stocker in client.get("stockeurs", []):
                    flattened_data.append({
                        "Company": company,
                        "Stockeur": stocker.get("stockeur", ""),
                        "Num Stockeur": stocker.get("num_stockeur", ""),
                        "Num Département": stocker.get("num_dep", ""),
                        "Ville": stocker.get("ville", ""),
                        "Pays": stocker.get("pays", "")
                    })
            
            if flattened_data:
                df_clients = pd.DataFrame(flattened_data)
                st.dataframe(df_clients, use_container_width=True)
            else:
                st.info("Clients found, but no stockers are registered.")
                
            with st.expander("View Raw JSON Data"):
                st.json(clients_data)
        else:
            st.info("The Clients database is empty.")
    else:
        st.info("No Clients.jsonl file found yet.")

# ------------------------------------------
# 3. SHOW ORDERS SECTION
# ------------------------------------------
with tab3:
    st.header("Validated Order Requests")
    orders = read_jsonl_file(os.path.join(path,"orders.jsonl"))
    if orders:
        st.dataframe(pd.DataFrame(orders), use_container_width=True)
    else:
        st.info("No valid orders found yet. Click 'Analyse Mails' to process your inbox.")

# ------------------------------------------
# 4. SHOW NON ORDERS SECTION
# ------------------------------------------
with tab4:
    st.header("Non-Order Emails")
    non_orders = read_jsonl_file(os.path.join(path,"non_orders.jsonl"))
    if non_orders:
        st.dataframe(pd.DataFrame(non_orders), use_container_width=True)
    else:
        st.info("No non-order emails found.")

# ------------------------------------------
# 5. SHOW MAILS SECTION
# ------------------------------------------
with tab5:
    st.header("All Processed Emails")
    emails = read_jsonl_file(os.path.join(path,"all_mails.jsonl"))
    if emails:
        st.dataframe(pd.DataFrame(emails), use_container_width=True)
    else:
        st.info("No email records loaded.")

# ------------------------------------------
# 6. EDIT ORDER SECTION
# ------------------------------------------
with tab6:
    st.header("Lookup & Edit Order")
    
    orders = read_jsonl_file(os.path.join(path, "orders.jsonl"))
    all_mails = read_jsonl_file(os.path.join(path, "all_mails.jsonl"))
    
    if orders:
        order_ids = [str(o.get("local_mail_id")) for o in orders if "local_mail_id" in o]
        selected_id = st.selectbox("Select Order (local_mail_id)", order_ids)
        
        target_order = next((o for o in orders if str(o.get("local_mail_id")) == selected_id), None)
        target_mail = next((m for m in all_mails if str(m.get("local_mail_id")) == selected_id), None)
        
        col_mail_info, col_order_edit = st.columns(2)
        
        with col_mail_info:
            st.subheader("📩 Corresponding Email Details")
            if target_mail:
                st.json(target_mail)
            else:
                st.warning("No matching raw email record found in `all_mails.jsonl`.")
                
        with col_order_edit:
            st.subheader("✏️ Edit Order Content")
            if target_order:
                with st.form("edit_order_form"):
                    edited_fields = {}
                    for k, v in target_order.items():
                        if k == "local_mail_id" or k == "local_order_id":
                            st.text_input(f"{k} (Read-Only)", value=str(v), disabled=True)
                            edited_fields[k] = v
                        else:
                            edited_fields[k] = st.text_input(str(k), value=str(v))
                    submit_update = st.form_submit_button("Save Order Updates")
                    if submit_update:
                        if "update_order" in globals():
                            update_order(path, edited_fields)
                        st.success("Order updated successfully!")
    else:
        st.info("No order records available to edit.")