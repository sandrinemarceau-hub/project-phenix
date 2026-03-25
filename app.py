import streamlit as st
import pandas as pd
from datetime import timedelta
import io

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
# BOUTON D'URGENCE
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
    return s

# NOUVEAU : Lecteur Universel V13 avec Auto-détection (sep=None)
def lire_fichier(fichier, lignes_a_ignorer):
    nom = fichier.name.lower()
    fichier.seek(0)
    if nom.endswith('.csv'):
        try:
            # sep=None force Pandas à deviner le séparateur (, ou ;) tout seul !
            return pd.read_csv(fichier, skiprows=lignes_a_ignorer, sep=None, engine='python', encoding='utf-8')
        except:
            fichier.seek(0)
            return pd.read_csv(fichier, skiprows=lignes_a_ignorer, sep=None, engine='python', encoding='latin-1')
    else:
        return pd.read_excel(fichier, skiprows=lignes_a_ignorer)

# ==========================================
# 2. INTERFACE VISUELLE
# ==========================================
st.set_page_config(layout="wide", page_title="Portail Logistique V13")
st.title("📦 Portail de Disponibilité - VERSION 13 🔴")
st.write("Le lecteur CSV intelligent à auto-détection est activé.")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. Fichier Stocks")
    fichier_stock = st.file_uploader("Fichier Stock (.csv, .xls, .xlsx)", type=['xlsx', 'xls', 'csv'])
    skip_stock = st.number_input("Lignes à ignorer (Stock)", min_value=0, value=3)

with col2:
    st.subheader("2. Fichiers Production")
    fichiers_prod = st.file_uploader("Fichiers Prod (.csv, .xls, .xlsx)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True)
    skip_prod = st.number_input("Lignes à ignorer (Prod)", min_value=0, value=0)

with col3:
    st.subheader("3. Fichier Commandes")
    fichier_commandes = st.file_uploader("Fichier Commandes (.csv, .xls, .xlsx)", type=['xlsx', 'xls', 'csv'])
    skip_cmd = st.number_input("Lignes à ignorer (Cmd)", min_value=0, value=0)

# ==========================================
# 3. LE MOTEUR DE CALCUL
# ==========================================
st.divider()

if st.button("🚀 Calculer les disponibilités (V13)", type="primary", use_container_width=True):
    if fichier_stock and fichiers_prod and fichier_commandes:
        with st.spinner('Lecture auto-détectée et calcul en cours...'):
            try:
                rapport = {}

                # --- A. LECTURE STOCK ---
                df_stock_brut = lire_fichier(fichier_stock, skip_stock)
                df_stock_brut.columns = df_stock_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                
                col_art_stock = next((c for c in ['CODEARTICLE', 'ARTICLECODE', 'ARTICLE'] if c in df_stock_brut.columns), None)
                col_qte_stock = next((c for c in ['STOCKDISPONIBLE', 'QTESTOCK', 'QUANTITE', 'STOCK'] if c in df_stock_brut.columns), None)
                
                if not col_art_stock:
                    st.error("❌ Erreur STOCK : La colonne Code Article est introuvable.")
                    st.info(f"🔍 Colonnes lues par l'outil : {df_stock_brut.columns.tolist()}")
                    st.stop()

                df_stock = pd.DataFrame()
                df_stock['CODE_ARTICLE'] = nettoyage_extreme(df_stock_brut[col_art_stock])
                df_stock['STOCK_DISPO'] = pd.to_numeric(df_stock_brut[col_qte_stock].astype(str).str.replace(',', '.'), errors='coerce').fillna(0) if col_qte_stock else 0
                
                stock_actuel = df_stock.groupby('CODE_ARTICLE')['STOCK_DISPO'].sum().to_dict()
                rapport['stock_lignes'] = len(df_stock)

                # --- B. LECTURE PRODUCTION ---
                liste_prod = []
                df_prod_brut_total = pd.DataFrame() 

                for f in fichiers_prod:
                    df_temp = lire_fichier(f, skip_prod)
                    
                    df_temp_copy = df_temp.copy()
                    df_temp_copy['SOURCE'] = f.name
                    df_prod_brut_total = pd.concat([df_prod_brut_total, df_temp_copy], ignore_index=True)

                    df_temp.columns = df_temp.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                    
                    col_art_prod = next((c for c in ['ARTICLECODEAE', 'CODEARTENTREE', 'ARTENTREE', 'ARTICLECODE', 'CODEARTICLE', 'ARTICLE'] if c in df_temp.columns), None)
                    col_qte_prod = next((c for c in ['QTEAE', 'QTEARTENTREE', 'QTEENTREE', 'QUANTITE', 'QTE'] if c in df_temp.columns), None)
                    
                    if col_art_prod and col_qte_prod:
                        df_extracted = pd.DataFrame()
                        df_extracted['ARTICLE'] = nettoyage_extreme(df_temp[col_art_prod])
                        df_extracted['QTE_PRODUITE'] = pd.to_numeric(df_temp[col_qte_prod].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                        
                        date_series = None
                        colonnes_dates_possibles = ['DATEREALISATION', 'DATEPLANIF', 'DATEFIN', 'DATEPREVUE', 'ECHEANCE', 'DATE']
                        
                        for col in colonnes_dates_possibles:
                            if col in df_temp.columns:
                                s_test = pd.to_datetime(df_temp[col], dayfirst=True, errors='coerce')
                                if date_series is None:
                                    date_series = s_test
                                else:
                                    date_series = date_series.fillna(s_test)
                                    
                        if date_series is None or date_series.isna().all():
                            for col in df_temp.columns:
                                if 'DATE' in col:
                                    s_test = pd.to_datetime(df_temp[col], dayfirst=True, errors='coerce')
                                    if date_series is None:
                                        date_series = s_test
                                    else:
                                        date_series = date_series.fillna(s_test)
                                        
                        if date_series is None:
                            date_series = pd.Series(pd.NaT, index=df_temp.index)
                                    
                        df_extracted['DATE_PROD'] = date_series
                        liste_prod.append(df_extracted)
                
                if not liste_prod:
                    st.error("❌ Erreur PRODUCTION : Impossible de lire les colonnes d'Article et de Quantité.")
                    st.stop()
                    
                st.session_state['df_prod_brut'] = df_prod_brut_total 
                
                df_production = pd.concat(liste_prod, ignore_index=True)
                lignes_prod_initiales = len(df_production)
                
                df_production_valide = df_production.dropna(subset=['DATE_PROD']).copy()
                df_production_valide = df_production_valide[df_production_valide['QTE_PRODUITE'] > 0]
                
                rapport['prod_initiales'] = lignes_prod_initiales
                rapport['prod_valides'] = len(df_production_valide)
                rapport['prod_ignorees'] = lignes_prod_initiales - len(df_production_valide)

                df_production_valide['Date_Dispo_Reelle'] = df_production_valide['DATE_PROD'] + timedelta(days=2)
                df_production_valide = df_production_valide.sort_values(by=['ARTICLE', 'Date_Dispo_Reelle'])
                productions_futures = df_production_valide.to_dict('records')

                # --- C. LECTURE COMMANDES ---
                df_commandes_brut = lire_fichier(fichier_commandes, skip_cmd)
                df_commandes_brut.columns = df_commandes_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                
                col_art_cmd = next((c for c in ['ARTICLECODE', 'CODEARTICLE', 'ARTICLE'] if c in df_commandes_brut.columns), None)
                col_date_cmd = next((c for c in ['DATECDE', 'DATECOMMANDE', 'DATECREATION', 'DATE'] if c in df_commandes_brut.columns), None)
                col_qte_cmd = next((c for c in ['QTEUBCDETOTAL', 'QTEUBCDE', 'QUANTITE', 'QTE'] if c in df_commandes_brut.columns), None)
                col_num_cmd = next((c for c in ['NUMCDE', 'NUMCOMMANDE', 'COMMANDE'] if c in df_commandes_brut.columns), None)
                col_client = next((c for c in ['EXPENOMCLIENT', 'CLIENT', 'NOMCLIENT'] if c in df_commandes_brut.columns), None)
                col_urgence = next((c for c in ['URGENCE', 'PRIORITE'] if c in df_commandes_brut.columns), None)

                if not col_art_cmd or not col_date_cmd:
                    st.error("❌ Erreur COMMANDES : Colonnes Article ou Date introuvables.")
                    st.stop()

                df_commandes = pd.DataFrame()
                df_commandes['ARTICLE_CODE'] = nettoyage_extreme(df_commandes_brut[col_art_cmd])
                df_commandes['DATE_CDE'] = pd.to_datetime(df_commandes_brut[col_date_cmd], dayfirst=True, errors='coerce')
                df_commandes['QUANTITE'] = pd.to_numeric(df_commandes_brut[col_qte_cmd].astype(str).str.replace(',', '.'), errors='coerce').fillna(0) if col_qte_cmd else 0
                df_commandes['NUM_CDE'] = df_commandes_brut[col_num_cmd] if col_num_cmd else 'Inconnu'
                df_commandes['CLIENT'] = df_commandes_brut[col_client] if col_client else 'Inconnu'
                df_commandes['URGENCE'] = pd.to_numeric(df_commandes_brut[col_urgence], errors='coerce').fillna(0) if col_urgence else 0
                
                lignes_cmd_initiales = len(df_commandes)
                df_commandes = df_commandes.dropna(subset=['DATE_CDE'])
                rapport['cmd_valides'] = len(df_commandes)
                
                df_commandes = df_commandes.sort_values(by=['URGENCE', 'DATE_CDE'], ascending=[False, True])

                # --- D. ALGORITHME D'ATTRIBUTION ---
                resultats = []
                for index, commande in df_commandes.iterrows():
                    article = commande['ARTICLE_CODE']
                    qte_demandee = commande['QUANTITE']
                    qte_restante = qte_demandee
                    date_dispo = "Immédiate (En Stock)"
                    
                    stock_dispo = stock_actuel.get(article, 0)
                    if stock_dispo > 0:
                        qte_prise = min(stock_dispo, qte_restante)
                        stock_actuel[article] -= qte_prise
                        qte_restante -= qte_prise
                        
                    if qte_restante > 0:
                        date_dispo = None
                        for prod in productions_futures:
                            if prod['ARTICLE'] == article and prod['QTE_PRODUITE'] > 0:
                                qte_prise = min(prod['QTE_PRODUITE'], qte_restante)
                                prod['QTE_PRODUITE'] -= qte_prise
                                qte_restante -= qte_prise
                                if qte_restante == 0:
                                    date_dispo = prod['Date_Dispo_Reelle'].strftime('%d/%m/%Y')
                                    break
                                    
                    if qte_restante > 0:
                        date_dispo = f"⚠️ Manque {int(qte_restante)} unités"
                        
                    resultats.append({
                        'Num_Commande': commande['NUM_CDE'],
                        'Article': article,
                        'Client': commande['CLIENT'],
                        'Quantité': qte_demandee,
                        'Date_Disponibilité': date_dispo
                    })

                st.session_state['df_final'] = pd.DataFrame(resultats)
                st.session_state['rapport'] = rapport
                st.session_state['calcul_ok'] = True

            except Exception as e:
                st.error(f"Une erreur s'est produite. Détails : {e}")
                st.session_state['calcul_ok'] = False
    else:
        st.warning("Veuillez déposer tous les fichiers demandés avant de lancer le calcul.")

# ==========================================
# 4. AFFICHAGE DES RÉSULTATS
# ==========================================
if st.session_state['calcul_ok']:
    st.success("✅ Calcul terminé avec succès !")
    
    rapport = st.session_state['rapport']
    with st.expander("📊 Voir le rapport de lecture des données", expanded=False):
        st.write(f"- **Stock :** {rapport.get('stock_lignes', 0)} articles lus.")
        st.write(f"- **Commandes :** {rapport.get('cmd_valides', 0)} commandes à traiter.")
        st.write(f"- **Production :** {rapport.get('prod_valides', 0)} lignes de production valides.")

    st.dataframe(st.session_state['df_final'], use_container_width=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        st.session_state['df_final'].to_excel(writer, index=False, sheet_name='Dispos')
    
    st.download_button(
        label="📥 Télécharger le fichier des disponibilités",
        data=buffer,
        file_name="Commandes_Disponibles.xlsx",
        mime="application/vnd.ms-excel",
        type="primary"
    )

    # ==========================================
    # SCANNER GLOBAL V13
    # ==========================================
    st.divider()
    st.subheader("🕵️‍♂️ Scanner Global V13")
    recherche = st.text_input("Tapez votre numéro (ex: 39586) et appuyez sur Entrée :")
    
    if recherche:
        recherche = str(recherche).strip().upper()
        df_prod_brut = st.session_state['df_prod_brut']
        
        mask = df_prod_brut.astype(str).apply(lambda x: x.str.contains(recherche, case=False, na=False))
        prods_trouvees = df_prod_brut[mask.any(axis=1)].copy()
        
        if not prods_trouvees.empty:
            prods_trouvees = prods_trouvees.dropna(axis=1, how='all')
            st.success(f"🏭 **{len(prods_trouvees)} ligne(s) trouvée(s) :**")
            for col in prods_trouvees.columns:
                if pd.api.types.is_datetime64_any_dtype(prods_trouvees[col]):
                    prods_trouvees[col] = prods_trouvees[col].dt.strftime('%d/%m/%Y')
            st.dataframe(prods_trouvees, use_container_width=True)
        else:
            st.error(f"⚠️ Le numéro {recherche} n'existe pas.")
