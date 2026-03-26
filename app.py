import streamlit as st
import pandas as pd
from datetime import timedelta, datetime
import io
import zipfile
import re
import math
import os

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

try:
    from fpdf import FPDF
    FPDF_OK = True
except ImportError:
    FPDF_OK = False

# ==========================================
# CONFIGURATION ET SÉCURITÉ (DOUBLE FACE)
# ==========================================
st.set_page_config(layout="wide", page_title="Sovereign Brands - Orders")

if 'role' not in st.session_state:
    st.session_state['role'] = None

# --- ECRAN DE CONNEXION ---
if st.session_state['role'] is None:
    st.title("🔐 Portail d'Accès Logistique")
    st.write("Veuillez entrer votre mot de passe pour accéder à votre espace.")
    pwd = st.text_input("Mot de passe", type="password")
    
    if st.button("Connexion", type="primary"):
        if pwd == "Logistique2026!":
            st.session_state['role'] = 'admin'
            st.rerun()
        elif pwd == "ClientSovereign!":
            st.session_state['role'] = 'client'
            st.rerun()
        else:
            st.error("Mot de passe incorrect. / Incorrect password.")
    st.stop()

# --- BOUTON DÉCONNEXION ---
with st.sidebar:
    if st.session_state['role'] == 'admin':
        if st.button("🚪 Déconnexion"):
            st.session_state.clear()
            st.rerun()
    elif st.session_state['role'] == 'client':
        if st.button("🚪 Log Out"):
            st.session_state.clear()
            st.rerun()
    st.divider()

# ==========================================
# FONCTIONS TECHNIQUES (COMMUNES)
# ==========================================
def nettoyage_extreme(serie):
    s = serie.astype(str).str.replace(r'\.0$', '', regex=True).str.upper().str.replace(r'[^A-Z0-9]', '', regex=True).str.lstrip('0').replace('', '0') 
    return s

def nettoyage_quantite(serie):
    def clean_val(x):
        x = str(x).replace(' ', '').replace('\xa0', '') 
        if ',' in x and '.' in x:
            if x.rfind(',') > x.rfind('.'): x = x.replace('.', '').replace(',', '.') 
            else: x = x.replace(',', '') 
        else: x = x.replace(',', '.') 
        x = re.sub(r'[^\d.-]', '', x) 
        try: return float(x)
        except: return 0.0
    return serie.apply(clean_val)

def clean_nan(val, default=""):
    if pd.isna(val) or str(val).strip().lower() in ['nan', 'nat', 'none', '']: return default
    return str(val).strip()

def lire_fichier(fichier, lignes_a_ignorer):
    nom = fichier.name.lower()
    fichier.seek(0)
    if nom.endswith('.csv'):
        try: return pd.read_csv(fichier, skiprows=lignes_a_ignorer, sep=None, engine='python', encoding='utf-8')
        except:
            fichier.seek(0)
            return pd.read_csv(fichier, skiprows=lignes_a_ignorer, sep=None, engine='python', encoding='latin-1')
    else:
        xls = pd.ExcelFile(fichier)
        best_df = None
        max_score = -1
        mots_cles = ['ARTICLECODE', 'CODEARTICLE', 'ARTICLE', 'QUANTITE', 'QTE', 'STOCK', 'POIDS', 'LIBELLE', 'PALETTE', 'FORMAT', 'UAUEMAX', 'STOCKPHYSIQUE']
        for sheet in xls.sheet_names:
            try:
                df_temp = pd.read_excel(xls, sheet_name=sheet, skiprows=lignes_a_ignorer)
                cols = df_temp.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                score = sum(1 for c in cols for m in mots_cles if m in c)
                if score > max_score: max_score = score; best_df = df_temp
            except: pass
        if best_df is not None: return best_df
        return pd.read_excel(xls, sheet_name=0, skiprows=lignes_a_ignorer)

def generer_packing_lists_zip(df_resultats, dict_details): return b"" # Simplifié ici pour le visuel, le code d'origine fonctionne.
def generer_rdv_documents_zip(df_resultats, dict_details, app_settings): return b"" # Simplifié ici.

# ==========================================
# ESPACE ADMINISTRATEUR (BACK OFFICE)
# ==========================================
if st.session_state['role'] == 'admin':
    st.title("🛠️ Back Office - Mise à jour de la Base")
    st.info("Le code du BackOffice n'a pas changé. Cliquez sur Calculer pour générer la base.")
    # ... (Le code du back office habituel reste identique, je le raccourcis pour que vous puissiez voir le Front Office plus bas)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1: fichier_stock = st.file_uploader("Fichier Stock", type=['xlsx', 'xls', 'csv']); skip_stock = st.number_input("Ignorer (Stock)", min_value=0, value=3)
    with col2: fichiers_prod = st.file_uploader("Fichiers Prod", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True); skip_prod = st.number_input("Ignorer (Prod)", min_value=0, value=0)
    with col3: fichier_commandes = st.file_uploader("Fichier Cmds", type=['xlsx', 'xls', 'csv']); skip_cmd = st.number_input("Ignorer (Cmd)", min_value=0, value=0)
    with col4: fichiers_nom = st.file_uploader("Fichiers (Poids & Liens)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True); skip_nom = st.number_input("Ignorer (Nom.)", min_value=0, value=0)

    if st.button("🚀 Calculer et Sauvegarder la Base", type="primary"):
        # Reprenez la logique du bouton 'Calculer' de la V49 ici.
        pass

# ==========================================
# ESPACE CLIENT (FRONT OFFICE) - V51 (UI SAAS)
# ==========================================
elif st.session_state['role'] == 'client':
    
    # --- CSS MASSIVE INJECTION POUR LE LOOK SAAS ---
    st.markdown("""
        <style>
            /* Cache les éléments natifs de Streamlit */
            #MainMenu, footer, header {visibility: hidden;}
            .block-container {
                padding-top: 1rem;
                padding-bottom: 0rem;
                max-width: 95%;
            }
            
            /* Fond de la page gris clair */
            .stApp {
                background-color: #f4f5f7;
            }
            
            /* Panneau principal blanc avec ombre */
            .main-panel {
                background-color: white;
                padding: 30px;
                border-radius: 8px;
                box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.05);
                margin-top: 20px;
            }

            /* Style des Badges (Pilules) */
            .badge-ready {
                background-color: #e6f4ea;
                color: #1e8e3e;
                padding: 4px 10px;
                border-radius: 12px;
                font-size: 0.85rem;
                font-weight: 600;
                display: inline-block;
            }
            .badge-pending {
                background-color: #fef7e0;
                color: #b06000;
                padding: 4px 10px;
                border-radius: 12px;
                font-size: 0.85rem;
                font-weight: 600;
                display: inline-block;
            }
            .badge-blocked {
                background-color: #fce8e6;
                color: #d93025;
                padding: 4px 10px;
                border-radius: 12px;
                font-size: 0.85rem;
                font-weight: 600;
                display: inline-block;
            }

            /* Style du tableau HTML intérieur */
            .custom-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 0.9rem;
                color: #333;
                margin-top: 10px;
            }
            .custom-table th {
                color: #6b778c;
                font-weight: 500;
                text-align: left;
                padding-bottom: 10px;
                border-bottom: 1px solid #dfe1e6;
            }
            .custom-table td {
                padding: 12px 0;
                border-bottom: 1px solid #f4f5f7;
                vertical-align: middle;
            }
            .item-name {
                font-weight: 600;
                color: #172b4d;
                display: block;
            }
            .item-sku {
                color: #7a869a;
                font-size: 0.8rem;
            }
            
            /* Customiser l'expander de Streamlit */
            div[data-testid="stExpander"] {
                border: 1px solid #dfe1e6 !important;
                border-radius: 4px !important;
                margin-bottom: 10px !important;
                background-color: white !important;
                box-shadow: none !important;
            }
            div[data-testid="stExpander"] summary {
                padding: 15px !important;
            }
            div[data-testid="stExpander"] summary p {
                font-weight: 600 !important;
                font-size: 1rem !important;
                color: #172b4d !important;
            }
        </style>
    """, unsafe_allow_html=True)

    # Conteneur principal blanc (Hack HTML)
    st.markdown('<div class="main-panel">', unsafe_allow_html=True)

    st.title("Orders")
    
    if not os.path.exists('base_logistique.pkl'):
        st.warning("⚠️ No data available.")
        st.stop()
    
    try:
        cache = pd.read_pickle('base_logistique.pkl')
        df_final = cache['df_final']
        dict_details = cache['dict_details']
    except Exception as e:
        st.error("Error reading the database.")
        st.stop()

    order_status_map = {}
    for cmd, grp in df_final.groupby('Num_Commande'):
        if str(cmd).upper() in ["INCONNU", "NAN"]: continue
        statuses = grp['Statut'].values
        if 'Rupture' in statuses: order_status_map[cmd] = 'Unfulfilled'
        elif any('Attente Prod' in s for s in statuses): order_status_map[cmd] = 'Pending'
        else: order_status_map[cmd] = 'Fulfilled'

    # --- TOP HEADER (RECHERCHE + BOUTONS) ---
    col_search, col_space, col_btn1, col_btn2 = st.columns([3, 1, 1.5, 1.5])
    search_query = col_search.text_input("🔍 Search Order ID or Customer...", label_visibility="collapsed", placeholder="Search Order ID or Customer...")
    
    # (Nous gardons les boutons Excel basiques de Streamlit pour que la logique de téléchargement fonctionne)
    mask_us_ca = df_final['Pays'].astype(str).str.upper().str.contains('ETATS|CANADA', regex=True, na=False)
    buf_monde = io.BytesIO()
    with pd.ExcelWriter(buf_monde, engine='openpyxl') as writer: df_final[~mask_us_ca].to_excel(writer, index=False)
    col_btn1.download_button("Export ROW", data=buf_monde.getvalue(), file_name="Orders.xlsx", use_container_width=True)
    
    st.markdown("<br>", unsafe_allow_html=True)

    # --- LISTE DES COMMANDES FAÇON SAAS ---
    groupes_commandes = df_final.groupby('Num_Commande')
    
    for cmd, lignes in groupes_commandes:
        if str(cmd).upper() in ["INCONNU", "NAN"]: continue
        
        statut_cmd = order_status_map[cmd]
        texte_recherche = f"{cmd} {lignes['Client'].iloc[0]}".lower()
        if search_query and search_query.lower() not in texte_recherche: continue

        client_nom = str(lignes.iloc[0]['Client'])
        pays_nom = str(lignes.iloc[0]['Pays'])
        if "ETATS" in pays_nom.upper() or "USA" in pays_nom.upper(): pays_nom_display = "USA"
        elif "CANADA" in pays_nom.upper(): pays_nom_display = "Canada"
        else: pays_nom_display = pays_nom.capitalize()

        nb_items = len(lignes)

        # Création de la pastille de statut dynamique
        if statut_cmd == 'Fulfilled': 
            badge_html = f"🟢 Ready"
        elif statut_cmd == 'Pending': 
            badge_html = f"🟡 Pending"
        else: 
            badge_html = f"🔴 Blocked"

        titre_accordian = f"#{cmd}   |   {client_nom}   |   {pays_nom_display}   |   Items: {nb_items}   |   {badge_html}"
        
        with st.expander(titre_accordian):
            # Construction du tableau HTML interne
            html_table = "<table class='custom-table'><thead><tr><th>Product Details</th><th>Qty (Cases)</th><th>Availability</th><th>Status</th></tr></thead><tbody>"
            
            for _, row in lignes.iterrows():
                art = str(row['Article'])
                qte = int(row['Qte_Demandée'])
                statut_fr = str(row['Statut'])
                date = str(row['Date_Disponibilité']).replace(" (Partiel)", "")
                libelle = dict_details.get(art, {}).get('libelle', 'Unknown Item')
                
                # Pillules internes
                if statut_fr == "Rupture":
                    pill = "<span class='badge-blocked'>Out of stock</span>"
                    date_display = "TBD"
                elif "Attente Prod" in statut_fr:
                    pill = "<span class='badge-pending'>In Production</span>"
                    date_display = date
                else:
                    pill = "<span class='badge-ready'>On Hand</span>"
                    date_display = "Immediate"

                html_table += f"""
                <tr>
                    <td>
                        <span class='item-name'>{libelle}</span>
                        <span class='item-sku'>SKU: {art}</span>
                    </td>
                    <td><b>{qte}</b></td>
                    <td>{date_display}</td>
                    <td>{pill}</td>
                </tr>
                """
            html_table += "</tbody></table>"
            st.markdown(html_table, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True) # Fin du conteneur blanc
