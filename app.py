import streamlit as st
import pandas as pd
from datetime import timedelta
import io
import zipfile
import re

try:
    from fpdf import FPDF
    FPDF_OK = True
except ImportError:
    FPDF_OK = False

# ==========================================
# 1. SYSTÈME DE SÉCURITÉ
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == "Logistique2026!":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Mot de passe", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Mot de passe", type="password", on_change=password_entered, key="password")
        st.error("Mot de passe incorrect 😕")
        return False
    return True

if not check_password():
    st.stop()

# ==========================================
# OUTILS
# ==========================================
with st.sidebar:
    st.write("🛠️ **Outils techniques**")
    if st.button("🗑️ Vider le cache et Redémarrer"):
        st.session_state.clear()
        st.rerun()

if 'calcul_ok' not in st.session_state:
    st.session_state['calcul_ok'] = False

def nettoyage_extreme(serie):
    s = serie.astype(str)
    s = s.str.replace(r'\.0$', '', regex=True) 
    s = s.str.upper() 
    s = s.str.replace(r'[^A-Z0-9]', '', regex=True) 
    # Le destructeur de zéros qui a sauvé l'article 85633 !
    s = s.str.lstrip('0') 
    s = s.replace('', '0') 
    return s

def nettoyage_quantite(serie):
    def clean_val(x):
        x = str(x).replace(' ', '').replace('\xa0', '') 
        if ',' in x and '.' in x:
            if x.rfind(',') > x.rfind('.'):
                x = x.replace('.', '').replace(',', '.') 
            else:
                x = x.replace(',', '') 
        else:
            x = x.replace(',', '.') 
            
        x = re.sub(r'[^\d.-]', '', x) 
        try:
            return float(x)
        except:
            return 0.0
    return serie.apply(clean_val)

def clean_text_pdf(t):
    return str(t).encode('latin-1', 'replace').decode('latin-1')

def lire_fichier(fichier, lignes_a_ignorer):
    nom = fichier.name.lower()
    fichier.seek(0)
    if nom.endswith('.csv'):
        try:
            return pd.read_csv(fichier, skiprows=lignes_a_ignorer, sep=None, engine='python', encoding='utf-8')
        except:
            fichier.seek(0)
            return pd.read_csv(fichier, skiprows=lignes_a_ignorer, sep=None, engine='python', encoding='latin-1')
    else:
        return pd.read_excel(fichier, skiprows=lignes_a_ignorer)

def generer_packing_lists_zip(df_resultats, dict_details):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        commandes = df_resultats['Num_Commande'].unique()
        for cmd in commandes:
            if str(cmd).upper() in ["INCONNU", "NAN"]: continue
            lignes = df_resultats[df_resultats['Num_Commande'] == cmd]
            client = str(lignes.iloc[0]['Client'])
            
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, txt=f"PACKING LIST - COMMANDE: {clean_text_pdf(cmd)}", ln=True, align='C')
            pdf.set_font("Arial", "", 12)
            pdf.cell(0, 10, txt=f"Client: {clean_text_pdf(client)}", ln=True, align='C')
            pdf.ln(10)
            
            pdf.set_font("Arial", "B", 9)
            pdf.cell(25, 8, "Code", 1)
            pdf.cell(85, 8, "Description", 1)
            pdf.cell(30, 8, "Format", 1)
            pdf.cell(25, 8, "Bouteilles", 1, align='C')
            pdf.cell(25, 8, "Cartons", 1, align='C')
            pdf.ln()
            
            pdf.set_font("Arial", "", 8)
            for _, row in lignes.iterrows():
                art = str(row['Article'])
                qte = row['Qte_Demandée']
                details = dict_details.get(art, {'libelle': 'Inconnu', 'format': '', 'uc': 6})
                
                libelle = details['libelle'][:45]
                fmt = details['format'][:15]
                uc = details['uc']
                if pd.isna(uc) or uc <= 0: uc = 6
                cartons = int(qte / uc) if qte > 0 else 0
                
                pdf.cell(25, 8, clean_text_pdf(art), 1)
                pdf.cell(85, 8, clean_text_pdf(libelle), 1)
                pdf.cell(30, 8, clean_text_pdf(fmt), 1)
                pdf.cell(25, 8, str(int(qte)), 1, align='C')
                pdf.cell(25, 8, str(cartons), 1, align='C')
                pdf.ln()
                
            safe_name = str(cmd).replace('/', '_').replace('\\', '_')
            out = pdf.output(dest='S')
            pdf_bytes = out.encode('latin1') if isinstance(out, str) else out
            zip_file.writestr(f"Packing_List_{safe_name}.pdf", pdf_bytes)
            
    return zip_buffer.getvalue()

# ==========================================
# 2. INTERFACE VISUELLE
# ==========================================
st.set_page_config(layout="wide", page_title="Portail Logistique V20")
st.title("📦 Portail de Disponibilité - VERSION 20 🔴")
st.write("Équilibre parfait : Zéros retirés pour le Stock ET Colonnes strictes pour la Prod.")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.subheader("1. Stock")
    fichier_stock = st.file_uploader("Fichier Stock", type=['xlsx', 'xls', 'csv'])
    skip_stock = st.number_input("Ignorer (Stock)", min_value=0, value=3)

with col2:
    st.subheader("2. Production")
    fichiers_prod = st.file_uploader("Fichiers Prod", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True)
    skip_prod = st.number_input("Ignorer (Prod)", min_value=0, value=0)

with col3:
    st.subheader("3. Commandes")
    fichier_commandes = st.file_uploader("Fichier Cmds", type=['xlsx', 'xls', 'csv'])
    skip_cmd = st.number_input("Ignorer (Cmd)", min_value=0, value=0)

with col4:
    st.subheader("4. Nomenclature")
    fichier_nom = st.file_uploader("Fichier Nomencl.", type=['xlsx', 'xls', 'csv'])
    skip_nom = st.number_input("Ignorer (Nom.)", min_value=0, value=0)

# ==========================================
# 3. LE MOTEUR DE CALCUL
# ==========================================
st.divider()

if st.button("🚀 Calculer les disponibilités (V20)", type="primary", use_container_width=True):
    if fichier_stock and fichiers_prod and fichier_commandes:
        with st.spinner('Calcul et chaînage final en cours...'):
            try:
                # --- A. LECTURE NOMENCLATURE ---
                dict_prepa = {}
                dict_details = {}
                if fichier_nom:
                    df_nom_brut = lire_fichier(fichier_nom, skip_nom)
                    st.session_state['df_nom_brut'] = df_nom_brut.copy() 
                    
                    df_nom_brut.columns = df_nom_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                    c_art = next((c for c in ['ARTICLECODE', 'CODEARTICLE'] if c in df_nom_brut.columns), None)
                    c_prepa = next((c for c in ['ARTPREPA', 'PRODUITDEBASECODE'] if c in df_nom_brut.columns), None)
                    c_lib = next((c for c in ['ARTICLELIBELLE', 'LIBELLE'] if c in df_nom_brut.columns), None)
                    c_fmt = next((c for c in ['FORMAT'] if c in df_nom_brut.columns), None)
                    c_uc = next((c for c in ['UCUA', 'UC', 'PCB'] if c in df_nom_brut.columns), None)
                    
                    if c_art:
                        df_nom_brut['CLEAN_ART'] = nettoyage_extreme(df_nom_brut[c_art])
                        if c_prepa: df_nom_brut['CLEAN_PREPA'] = nettoyage_extreme(df_nom_brut[c_prepa])
                            
                        for _, r in df_nom_brut.iterrows():
                            art_id = str(r['CLEAN_ART'])
                            prepa_id = str(r['CLEAN_PREPA']) if c_prepa else ""
                            
                            if prepa_id and prepa_id != "0" and prepa_id != "NAN" and prepa_id != art_id:
                                dict_prepa[art_id] = prepa_id
                                
                            dict_details[art_id] = {
                                'libelle': str(r[c_lib]) if c_lib else "Inconnu",
                                'format': str(r[c_fmt]) if c_fmt else "",
                                'uc': float(nettoyage_quantite(pd.Series([r[c_uc]]))[0]) if c_uc else 6.0
                            }
                st.session_state['dict_details'] = dict_details

                # --- B. LECTURE STOCK ---
                df_stock_brut = lire_fichier(fichier_stock, skip_stock)
                st.session_state['df_stock_brut'] = df_stock_brut.copy() 
                
                df_stock_brut.columns = df_stock_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                # Retour à la liste stricte et sécurisée
                col_art_stock = next((c for c in ['CODEARTICLE', 'ARTICLECODE', 'ARTICLE'] if c in df_stock_brut.columns), None)
                col_qte_stock = next((c for c in ['STOCKDISPONIBLE', 'QTESTOCK', 'QUANTITE', 'STOCK', 'TOTAL', 'TOTALGNRAL', 'TOTALGENERAL'] if c in df_stock_brut.columns), None)
                
                df_stock = pd.DataFrame()
                df_stock['CODE_ARTICLE'] = nettoyage_extreme(df_stock_brut[col_art_stock])
                df_stock['STOCK_DISPO'] = nettoyage_quantite(df_stock_brut[col_qte_stock]) if col_qte_stock else 0
                stock_actuel = df_stock.groupby('CODE_ARTICLE')['STOCK_DISPO'].sum().to_dict()

                # --- C. LECTURE PRODUCTION ---
                liste_prod = []
                df_prod_brut_total = pd.DataFrame() 

                for f in fichiers_prod:
                    df_temp = lire_fichier(f, skip_prod)
                    df_temp_copy = df_temp.copy()
                    df_temp_copy['SOURCE'] = f.name
                    df_prod_brut_total = pd.concat([df_prod_brut_total, df_temp_copy], ignore_index=True)

                    df_temp.columns = df_temp.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                    # Retour à la liste stricte (Adieu le chasseur élastique qui cassait les dates)
                    col_art_prod = next((c for c in ['ARTICLECODEAE', 'CODEARTENTREE', 'ARTENTREE', 'ARTICLECODE', 'CODEARTICLE', 'ARTICLE'] if c in df_temp.columns), None)
                    col_qte_prod = next((c for c in ['QTEAE', 'QTEARTENTREE', 'QTEENTREE', 'QUANTITE', 'QTE', 'TOTAL', 'TOTALGNRAL', 'TOTALGENERAL'] if c in df_temp.columns), None)
                    
                    if col_art_prod and col_qte_prod:
                        df_ext = pd.DataFrame()
                        df_ext['ARTICLE'] = nettoyage_extreme(df_temp[col_art_prod])
                        df_ext['QTE_PRODUITE'] = nettoyage_quantite(df_temp[col_qte_prod])
                        
                        date_series = None
                        for col in ['DATEREALISATION', 'DATEPLANIF', 'DATEFIN', 'DATEPREVUE', 'ECHEANCE', 'DATE']:
                            if col in df_temp.columns:
                                s_test = pd.to_datetime(df_temp[col], dayfirst=True, errors='coerce')
                                date_series = s_test if date_series is None else date_series.fillna(s_test)
                                    
                        if date_series is None or date_series.isna().all():
                            for col in df_temp.columns:
                                if 'DATE' in col:
                                    s_test = pd.to_datetime(df_temp[col], dayfirst=True, errors='coerce')
                                    date_series = s_test if date_series is None else date_series.fillna(s_test)
                                        
                        df_ext['DATE_PROD'] = date_series if date_series is not None else pd.Series(pd.NaT, index=df_temp.index)
                        liste_prod.append(df_ext)
                
                st.session_state['df_prod_brut'] = df_prod_brut_total 
                df_production = pd.concat(liste_prod, ignore_index=True)
                df_production_valide = df_production.dropna(subset=['DATE_PROD']).copy()
                df_production_valide = df_production_valide[df_production_valide['QTE_PRODUITE'] > 0]
                
                df_production_valide['Date_Dispo_Reelle'] = df_production_valide['DATE_PROD'] + timedelta(days=2)
                df_production_valide = df_production_valide.sort_values(by=['ARTICLE', 'Date_Dispo_Reelle'])
                productions_futures = df_production_valide.to_dict('records')

                # --- D. LECTURE COMMANDES ---
                df_commandes_brut = lire_fichier(fichier_commandes, skip_cmd)
                df_commandes_brut.columns = df_commandes_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                
                col_art_cmd = next((c for c in ['ARTICLECODE', 'CODEARTICLE', 'ARTICLE'] if c in df_commandes_brut.columns), None)
                col_date_cmd = next((c for c in ['DATECDE', 'DATECOMMANDE', 'DATECREATION', 'DATE'] if c in df_commandes_brut.columns), None)
                col_qte_cmd = next((c for c in ['QTEUBCDETOTAL', 'QTEUBCDE', 'QUANTITE', 'QTE', 'TOTAL', 'TOTALGNRAL', 'TOTALGENERAL'] if c in df_commandes_brut.columns), None)
                col_num_cmd = next((c for c in ['NUMCDE', 'NUMCOMMANDE', 'COMMANDE'] if c in df_commandes_brut.columns), None)
                col_client = next((c for c in ['EXPENOMCLIENT', 'CLIENT', 'NOMCLIENT'] if c in df_commandes_brut.columns), None)
                
                df_commandes = pd.DataFrame()
                df_commandes['ARTICLE_CODE'] = nettoyage_extreme(df_commandes_brut[col_art_cmd])
                df_commandes['DATE_CDE'] = pd.to_datetime(df_commandes_brut[col_date_cmd], dayfirst=True, errors='coerce')
                df_commandes['QUANTITE'] = nettoyage_quantite(df_commandes_brut[col_qte_cmd])
                df_commandes['NUM_CDE'] = df_commandes_brut[col_num_cmd] if col_num_cmd else 'Inconnu'
                df_commandes['CLIENT'] = df_commandes_brut[col_client] if col_client else 'Inconnu'
                df_commandes['URGENCE'] = 0
                
                df_commandes = df_commandes.dropna(subset=['DATE_CDE'])
                df_commandes = df_commandes.sort_values(by=['URGENCE', 'DATE_CDE'], ascending=[False, True])

                # --- E. ALGORITHME AVEC CHAÎNAGE ---
                resultats = []
                for index, commande in df_commandes.iterrows():
                    article = commande['ARTICLE_CODE']
                    qte_restante = commande['QUANTITE']
                    
                    qte_prise_stock = 0
                    qte_prise_prod = 0
                    dates_trouvees = []
                    
                    def consommer(code_a_chercher, qte_a_trouver):
                        q_stk, q_prd = 0, 0
                        s = stock_actuel.get(code_a_chercher, 0)
                        if s > 0:
                            prise = min(s, qte_a_trouver)
                            stock_actuel[code_a_chercher] -= prise
                            q_stk += prise
                            qte_a_trouver -= prise
                            
                        if qte_a_trouver > 0:
                            for prod in productions_futures:
                                if prod['ARTICLE'] == code_a_chercher and prod['QTE_PRODUITE'] > 0:
                                    prise = min(prod['QTE_PRODUITE'], qte_a_trouver)
                                    prod['QTE_PRODUITE'] -= prise
                                    q_prd += prise
                                    qte_a_trouver -= prise
                                    dates_trouvees.append(prod['Date_Dispo_Reelle'])
                                    if qte_a_trouver == 0: break
                        return q_stk, q_prd, qte_a_trouver

                    qs1, qp1, qte_restante = consommer(article, qte_restante)
                    qte_prise_stock += qs1
                    qte_prise_prod += qp1
                    
                    utilise_prepa = "Non"
                    if qte_restante > 0 and article in dict_prepa:
                        prepa = dict_prepa[article]
                        qs2, qp2, qte_restante = consommer(prepa, qte_restante)
                        qte_prise_stock += qs2
                        qte_prise_prod += qp2
                        if (qs2 + qp2) > 0: utilise_prepa = f"Oui ({prepa})"

                    if qte_restante > 0:
                        statut = "Rupture"
                        date_dispo = "Pas de date"
                    else:
                        if len(dates_trouvees) == 0:
                            date_dispo = "Immédiate"
                            statut = "En Stock"
                        else:
                            date_dispo = max(dates_trouvees).strftime('%d/%m/%Y')
                            statut = "Attente Prod"
                        
                    resultats.append({
                        'Num_Commande': commande['NUM_CDE'],
                        'Client': commande['CLIENT'],
                        'Article': article,
                        'Qte_Demandée': int(commande['QUANTITE']),
                        'Tiré_Stock': int(qte_prise_stock),
                        'Tiré_Prod': int(qte_prise_prod),
                        'Remplacement_Prepa': utilise_prepa,
                        'Manquant': int(qte_restante),
                        'Statut': statut,
                        'Date_Disponibilité': date_dispo
                    })

                st.session_state['df_final'] = pd.DataFrame(resultats)
                st.session_state['calcul_ok'] = True

            except Exception as e:
                st.error(f"Une erreur s'est produite. Détails : {e}")
                st.session_state['calcul_ok'] = False
    else:
        st.warning("Veuillez déposer Stock, Prod, Commandes et Nomenclature.")

# ==========================================
# 4. AFFICHAGE ET EXPORT PDF
# ==========================================
if st.session_state['calcul_ok']:
    st.success("✅ Calcul terminé avec succès !")
    st.dataframe(st.session_state['df_final'], use_container_width=True)

    c_btn1, c_btn2 = st.columns(2)
    with c_btn1:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state['df_final'].to_excel(writer, index=False, sheet_name='Analyse')
        st.download_button("📥 Télécharger l'Excel Détaillé", data=buffer, file_name="Analyse_V20.xlsx", type="primary")

    with c_btn2:
        if FPDF_OK:
            zip_data = generer_packing_lists_zip(st.session_state['df_final'], st.session_state['dict_details'])
            st.download_button("📦 Télécharger les Packing Lists PDF (.zip)", data=zip_data, file_name="Packing_Lists.zip", type="secondary")

    # ==========================================
    # SCANNER GLOBAL V20
    # ==========================================
    st.divider()
    st.subheader("🕵️‍♂️ Scanner Global Absolu V20")
    recherche = st.text_input("Tapez votre numéro (ex: 85633) et appuyez sur Entrée :")
    
    if recherche:
        rech_clean = re.sub(r'[^A-Z0-9]', '', recherche.strip().upper()).lstrip('0')
        col_s1, col_s2, col_s3 = st.columns(3)
        
        def display_scan(df_name, title, col):
            if df_name in st.session_state:
                df = st.session_state[df_name]
                def match_cell(val):
                    val_c = re.sub(r'[^A-Z0-9]', '', str(val).upper().replace('.0', '')).lstrip('0')
                    return rech_clean in val_c if rech_clean else False
                
                mask = df.applymap(match_cell)
                res = df[mask.any(axis=1)].copy().dropna(axis=1, how='all')
                col.write(f"**{title} : {len(res)} ligne(s)**")
                if not res.empty:
                    for c in res.columns:
                        if pd.api.types.is_datetime64_any_dtype(res[c]):
                            res[c] = res[c].dt.strftime('%d/%m/%Y')
                    col.dataframe(res, use_container_width=True)
                else:
                    col.info("Introuvable.")

        display_scan('df_stock_brut', '📦 STOCK', col_s1)
        display_scan('df_prod_brut', '🏭 PRODUCTION', col_s2)
        display_scan('df_nom_brut', '🧠 NOMENCLATURE', col_s3)
