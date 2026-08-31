import io
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import streamlit as st
from functions import *

# Helper function to launch native OS Folder Browser
def select_folder():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder_path = filedialog.askdirectory(master=root)
    root.destroy()
    return folder_path

# Helper function to safely fetch unread mail count
def get_unread_mail_count():
    try:
        # Calls helper from functions.py or IMAP checker
        mail=log_in()
        res=len(extract_mail_ids(mail))
        log_out(mail)
        return res
    except Exception:
        pass
    return 0

# --- Page Configuration ---
st.set_page_config(
    page_title="Email & Order Management",
    page_icon="📦",
    layout="wide"
)

# Initialize persistent session states from parametres.json
default_dir = os.path.join(os.path.expanduser("~"), "Desktop")
parametres = read_json_file("parametres.json")

# 1. Load existing path from JSON or initialize default
saved_path = parametres.get("paths", {}).get("main_directory_path")

if "work_folder" not in st.session_state:
    if saved_path and os.path.exists(saved_path):
        st.session_state["work_folder"] = saved_path
    else:
        init_main_folder_path(default_dir)
        parametres = read_json_file("parametres.json")
        st.session_state["work_folder"] = parametres["paths"]["main_directory_path"]

if "is_running" not in st.session_state:
    st.session_state["is_running"] = False

if "trigger_main" not in st.session_state:
    st.session_state["trigger_main"] = False

if "execution_logs" not in st.session_state:
    st.session_state["execution_logs"] = ""

# ==========================================
# SIDEBAR: Work Folder, Mode & Execution
# ==========================================
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # 1. Workspace Folder Selection
    st.subheader("1. Working Directory")
    
    col_path, col_btn = st.columns([3, 1])
    
    with col_path:
        input_folder = st.text_input(
            "📁 Folder Path", 
            value=st.session_state["work_folder"],
            label_visibility="collapsed",
            disabled=st.session_state["is_running"]
        )
        
        # Trigger update ONLY if the user typed a new path
        if input_folder != st.session_state["work_folder"]:
            init_main_folder_path(input_folder)
            updated_params = read_json_file("parametres.json")
            st.session_state["work_folder"] = updated_params["paths"]["main_directory_path"]
            st.rerun()

    with col_btn:
        if st.button("📂 Browse", use_container_width=True, disabled=st.session_state["is_running"]):
            selected = select_folder()
            if selected:
                init_main_folder_path(selected)
                updated_params = read_json_file("parametres.json")
                st.session_state["work_folder"] = updated_params["paths"]["main_directory_path"]
                st.rerun()

    # Active path display
    path = st.session_state["work_folder"]
    st.caption(f"Active Path: `{path}`")
  
    check_temp_files()

    st.markdown("---")
    
    # 2. Pre-Execution Speed Toggle
    st.subheader("2. Processing Mode")
    exec_mode = st.radio(
        "Select Execution Strategy",
        options=["🐢 Slow (Safe) Execution", "⚡ Fast Execution"],
        help="Fast mode runs without extra delays; Slow mode pauses between requests to prevent rate limits.",
        disabled=st.session_state["is_running"]
    )
    
    st.markdown("---")

    # 3. Execution Trigger
    btn_analyse = st.button(
        "🚀 Start Analysis", 
        type="primary", 
        use_container_width=True, 
        disabled=st.session_state["is_running"]
    )

    # --- Unread Mails Count Message ---
    unread_count = get_unread_mail_count()
    st.caption(f"📩 **Unread Mails in Inbox:** `{unread_count}` mail(s) pending")

    st.markdown("---")
    
    # 4. Console Logs Display
    st.subheader("📋 Execution Logs")
    output_container = st.empty()
    
    if st.session_state["execution_logs"]:
        output_container.code(st.session_state["execution_logs"], language="text")

st.title("📦 Email & Order Processing Hub")

# 1. Trigger phase: Lock controls, set config file based on toggle, and trigger execution
if btn_analyse:
    st.session_state["is_running"] = True
    st.session_state["trigger_main"] = True
    
    # Set mode flags in config based on user toggle
    parametres = read_json_file("parametres.json")
    if "⚡ Fast" in exec_mode:
        parametres["fast_stop"] = True
        parametres["slow_stop"] = False
    else:
        parametres["fast_stop"] = False
        parametres["slow_stop"] = True
    create_json_file(parametres, "parametres.json")

    st.session_state["execution_logs"] = f"[{exec_mode}] Mode configured. Launching...\n"
    st.rerun()

# 2. Execution phase: Run main() in thread pool and record logs
if st.session_state["trigger_main"]:
    st.session_state["trigger_main"] = False
    log_file_path = os.path.join(path, "execution.log")

    output_buffer = io.StringIO()
    sys.stdout = output_buffer

    with st.spinner("Connecting to inbox, downloading attachments, and analyzing mails..."):
        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(main)
            
            while not future.done():
                new_logs = output_buffer.getvalue()
                if new_logs:
                    st.session_state["execution_logs"] += new_logs
                    output_container.code(st.session_state["execution_logs"], language="text")
                    output_buffer.truncate(0)
                    output_buffer.seek(0)
                time.sleep(0.5)
            
            final_logs = output_buffer.getvalue()
            if final_logs:
                st.session_state["execution_logs"] += final_logs
                output_container.code(st.session_state["execution_logs"], language="text")
            
            with open(log_file_path, "a", encoding="utf-8") as f:
                f.write(st.session_state["execution_logs"] + "\n--- END OF RUN ---\n")

            exception = future.exception()
            if exception:
                st.session_state["execution_logs"] += f"\n[ERROR]: {exception}\n"
                output_container.code(st.session_state["execution_logs"], language="text")
                with st.sidebar:
                    st.error(f"Error: {exception}")
            else:
                st.success("Analysis complete!")

    sys.stdout = sys.__stdout__
    st.session_state["is_running"] = False
    st.rerun()

# ==========================================
# MAIN SECTIONS (TABS)
# ==========================================
if st.session_state["is_running"]:
    st.warning("⏳ Execution in progress... Tab controls and database entries are locked.")
else:
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Show Orders",
        "🏢 Show Clients",
        "📧 Show Non Orders",
        "📩 Show Mails",
        "👥 Add Client / Add Stocker",
        "✏️ Edit Order"
    ])

    # ------------------------------------------
    # 1. SHOW ORDERS SECTION
    # ------------------------------------------
    with tab1:
        st.header("Validated Order Requests")
        orders = read_jsonl_file(os.path.join(path, "orders.jsonl"))
        if orders:
            st.dataframe(pd.DataFrame(orders), use_container_width=True)
        else:
            st.info("No valid orders found yet. Click 'Start Analysis' to process your inbox.")

    # ------------------------------------------
    # 2. SHOW CLIENTS SECTION
    # ------------------------------------------
    with tab2:
        st.header("Registered Clients & Stockers")
        
        clients_file_path = os.path.join(path, "Clients.jsonl")
        if os.path.exists(clients_file_path):
            clients_data = read_jsonl_file(clients_file_path)
            
            if clients_data:
                flattened_data = []
                for client in clients_data:
                    company = client.get("company", "N/A")
                    for stocker in client.get("stockeurs", []):
                        flattened_data.append({
                            "Company": company,
                            "Stockeur": stocker.get("stockeur", ""),
                            "Num Stockeur": stocker.get("num_stockeur", ""),
                            "Num Destinataire": stocker.get("num_dest", ""),
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
    # 3. SHOW NON ORDERS SECTION
    # ------------------------------------------
    with tab3:
        st.header("Non-Order Emails")
        non_orders = read_jsonl_file(os.path.join(path, "non_orders.jsonl"))
        if non_orders:
            st.dataframe(pd.DataFrame(non_orders), use_container_width=True)
        else:
            st.info("No non-order emails found.")

    # ------------------------------------------
    # 4. SHOW MAILS SECTION
    # ------------------------------------------
    with tab4:
        st.header("All Processed Emails")
        emails = read_jsonl_file(os.path.join(path, "all_mails.jsonl"))
        if emails:
            st.dataframe(pd.DataFrame(emails), use_container_width=True)
        else:
            st.info("No email records loaded.")

    # ------------------------------------------
    # 5. ADD CLIENT / ADD STOCKER SECTION
    # ------------------------------------------
    with tab5:
        st.header("Manage Clients & Stocker Locations")
        
        col_client, col_stocker = st.columns(2)
        
        with col_client:
            st.subheader("Add New Client")
            with st.form("add_client_form", clear_on_submit=True):
                company_name = st.text_input("Company Name *").upper()
                st.markdown("**Initial Stocker Details:**")
                stockeur_name = st.text_input("Stockeur * (e.g., CO3 CMS)").upper()
                num_stockeur = st.text_input("Num Stockeur *").upper()
                num_dest = st.text_input("Num Destinataire *").upper()
                num_dep = st.text_input("Num Département *").upper()
                ville = st.text_input("Ville *").upper()
                pays = st.text_input("Pays *").upper()
                
                submit_client = st.form_submit_button("Add Client")
                
                if submit_client:
                    if not (company_name and stockeur_name and num_stockeur and num_dep and ville and pays):
                        st.error("All fields are mandatory!")
                    elif any(c["company"] == company_name for c in read_jsonl_file(os.path.join(path, "Clients.jsonl"))):
                        st.warning(f"Company '{company_name}' already exists.")
                    else:
                        new_client = {
                            "company": company_name,
                            "stockeurs": [
                                {
                                    "stockeur": stockeur_name,
                                    "num_stockeur": num_stockeur,
                                    "num_dest": num_dest,
                                    "num_dep": num_dep,
                                    "ville": ville,
                                    "pays": pays
                                }
                            ]
                        }
                        add_client(path, new_client)
                        st.success(f"Company '{company_name}' created with stockeur '{stockeur_name}'!")

        with col_stocker:
            st.subheader("Add Additional Stocker")
            clients_list = read_jsonl_file(os.path.join(path, "Clients.jsonl"))
            company_names = [c["company"] for c in clients_list] if clients_list else []
            
            if company_names:
                with st.form("add_stocker_form", clear_on_submit=True):
                    selected_company = st.selectbox("Select Company", company_names).upper()
                    st.markdown("**New Stocker Details:**")
                    extra_stockeur = st.text_input("Stockeur * (e.g., Dépôt Nord)").upper()
                    extra_num_stockeur = st.text_input("Num Stockeur *").upper()
                    extra_num_dest = st.text_input("Num Destinataire *").upper()
                    extra_num_dep = st.text_input("Num Département *").upper()
                    extra_ville = st.text_input("Ville *").upper()
                    extra_pays = st.text_input("Pays *").upper()
                    
                    submit_stocker = st.form_submit_button("Add Stocker")
                    
                    if submit_stocker:
                        if not (extra_stockeur and extra_num_stockeur and extra_num_dep and extra_ville and extra_pays):
                            st.error("All Stocker fields are mandatory!")
                        else:
                            target_entry = next(c for c in clients_list if c["company"] == selected_company)
                            
                            if any(s["stockeur"] == extra_stockeur for s in target_entry.get("stockeurs", [])):
                                st.warning("A stockeur with this name already exists for this company.")
                            else:
                                new_stockeur_obj = {
                                    "stockeur": extra_stockeur,
                                    "num_stockeur": extra_num_stockeur,
                                    "num_dest": extra_num_dest,
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