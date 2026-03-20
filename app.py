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

if 'calcul_ok' not in st.session_state:
    st.session_state['calcul_ok'] = False

# NOUVEAU : Fonction de nettoyage extrême des valeurs
def nettoyage_extreme(serie):
    """Force en texte, supprime le .0, met en majuscules, et supprime TOUT ce qui n'est pas lettre ou chiffre."""
    s = serie.astype(str)
    s = s.str.replace(r'\.0$', '', regex=True) 
    s = s.str.upper() 
    s = s.str.replace(r'[^A-Z0-9]', '', regex=True) 
    return s

# ==========================================
# 2. INTERFACE VISUELLE
# ==========================================
st.set_page_config(layout="wide", page_title="Portail Logistique")
st.title("📦 Portail de Disponibilité des Commandes")
st.write("Déposez vos exports ci-dessous. Le système est blindé contre les espaces invisibles de l'ERP.")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. Fichier Stocks")
    fichier_stock = st.file_uploader("Export_Stock.xlsx", type=['xlsx'])
    skip_stock = st.number_input("Lignes à ignorer (Stock)", min_value=0, value=3)

with col2:
    st.subheader("2. Fichiers Production")
    fichiers_prod = st.file_uploader("Glissez les fichiers (Max 3)", type=['xlsx'], accept_multiple_files=True)
    skip_prod = st.number_input("Lignes à ignorer (Prod)", min_value=0, value=0)

with col3:
    st.subheader("3. Fichier Commandes")
    fichier_commandes = st.file_uploader("Export_Commandes.xlsx", type=['xlsx'])
    skip_cmd = st.number_input("Lignes à ignorer (Cmd)", min_value=0, value=0)

# ==========================================
# 3. LE MOTEUR DE CALCUL
# ==========================================
st.divider()

if st.button("🚀 Calculer les disponibilités", type="primary", use_container_width=True):
    if fichier_stock and fichiers_prod and fichier_commandes:
        with st.spinner('Nettoyage brutal des en-têtes et calcul en cours...'):
            try:
                rapport = {}

                # --- A. LECTURE STOCK ---
                df_stock = pd.read_excel(fichier_stock, skiprows=skip_stock)
                # Nettoyage brutal des en-têtes (Supprime tous les espaces/caractères spéciaux)
                df_stock.columns = df_stock.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                
                if 'CODEARTICLE' not in df_stock.columns:
                    st.error("❌ Erreur STOCK : La colonne Code Article est introuvable.")
                    st.stop()

                df_stock['CODEARTICLE'] = nettoyage_extreme(df_stock['CODEARTICLE'])
                st.session_state['df_stock_brut'] = df_stock.copy()
                
                col_stock = 'STOCKDISPONIBLE' if 'STOCKDISPONIBLE' in df_stock.columns else df_stock.columns[-1]
                try:
                    stock_actuel = df_stock.set_index('CODEARTICLE')[col_stock].to_dict()
                except:
                    stock_actuel = {}
                    
                rapport['stock_lignes'] = len(df_stock)

                # --- B. LECTURE PRODUCTION ---
                liste_prod = []
                for f in fichiers_prod:
                    df_temp = pd.read_excel(f, skiprows=skip_prod)
                    # Nettoyage brutal des en-têtes pour supprimer les espaces invisibles
                    df_temp.columns = df_temp.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                    
                    colonnes_a_renommer = {
                        'ARTICLECODEAE': 'ARTICLE', 'CODEARTENTREE': 'ARTICLE', 'ARTENTREE': 'ARTICLE', 
                        'ARTICLE': 'ARTICLE', 'CODEARTICLE': 'ARTICLE',
                        'QTEAE': 'QTE_PRODUITE', 'QTEARTENTREE': 'QTE_PRODUITE', 'QTEENTREE': 'QTE_PRODUITE',
                        'QTE': 'QTE_PRODUITE', 'QUANTITE': 'QTE_PRODUITE',
                        'DATEPLANIF': 'DATE_PROD', 'DATEREALISATION': 'DATE_PROD', 'DATE': 'DATE_PROD'
                    }
                    df_temp = df_temp.rename(columns=colonnes_a_renommer)
                        
                    liste_prod.append(df_temp) 
                
                df_production = pd.concat(liste_prod, ignore_index=True)
                lignes_prod_initiales = len(df_production)
                
                df_production['ARTICLE'] = nettoyage_extreme(df_production['ARTICLE'])
                df_production['QTE_PRODUITE'] = pd.to_numeric(df_production['QTE_PRODUITE'], errors='coerce').fillna(0)

                st.session_state['df_prod_brut'] = df_production.copy()

                df_production['DATE_PROD'] = pd.to_datetime(df_production['DATE_PROD'], dayfirst=True, errors='coerce')
                df_production_valide = df_production.dropna(subset=['DATE_PROD']).copy()
                
                rapport['prod_initiales'] = lignes_prod_initiales
                rapport['prod_valides'] = len(df_production_valide)
                rapport['prod_ignorees'] = lignes_prod_initiales - len(df_production_valide)

                df_production_valide['Date_Dispo_Reelle'] = df_production_valide['DATE_PROD'] + timedelta(days=2)
                df_production_valide = df_production_valide.sort_values(by=['ARTICLE', 'Date_Dispo_Reelle'])
                productions_futures = df_production_valide.to_dict('records')

                # --- C. LECTURE COMMANDES ---
                df_commandes = pd.read_excel(fichier_commandes, skiprows=skip_cmd)
                df_commandes.columns = df_commandes.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                
                # Fallbacks pour les commandes
                col_art_cmd = 'ARTICLECODE' if 'ARTICLECODE' in df_commandes.columns else 'ARTICLE'
                col_date_cmd = 'DATECDE' if 'DATECDE' in df_commandes.columns else 'DATE'
                
                df_commandes['ARTICLECODE_NET'] = nettoyage_extreme(df_commandes[col_art_cmd])
                df_commandes['DATE_CDE_NET'] = pd.to_datetime(df_commandes[col_date_cmd], dayfirst=True, errors='coerce')
                
                lignes_cmd_initiales = len(df_commandes)
                df_commandes = df_commandes.dropna(subset=['DATE_CDE_NET'])
                
                rapport['cmd_valides'] = len(df_commandes)
                
                if 'URGENCE' not in df_commandes.columns:
                    df_commandes['URGENCE'] = 0
                df_commandes = df_commandes.sort_values(by=['URGENCE', 'DATE_CDE_NET'], ascending=[False, True])

                # --- D. ALGORITHME D'ATTRIBUTION ---
                resultats = []
                for index, commande in df_commandes.iterrows():
                    article = commande['ARTICLECODE_NET']
                    col_qte = 'QTEUBCDETOTAL' if 'QTEUBCDETOTAL' in df_commandes.columns else ('QTEUBCDE' if 'QTEUBCDE' in df_commandes.columns else 'QUANTITE')
                    
                    try:
                        qte_demandee = float(commande.get(col_qte, 0))
                    except:
                        qte_demandee = 0
                        
                    num_cmd = commande.get('NUMCDE', 'Inconnu')
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
                        'Num_Commande': num_cmd,
                        'Article': article,
                        'Client': commande.get('EXPENOMCLIENT', 'Inconnu'),
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
    # SCANNER GLOBAL (Mode Détective Ultime)
    # ==========================================
    st.divider()
    st.subheader("🕵️‍♂️ Scanner Global : Chercher un Numéro")
    st.write("Tapez votre numéro (ex: 39586). L'outil va fouiller dans TOUTES les colonnes de votre fichier brut.")
    
    recherche = st.text_input("Tapez votre numéro et appuyez sur Entrée :")
    
    if recherche:
        recherche = str(recherche).strip().upper()
        df_prod_brut = st.session_state['df_prod_brut']
        
        mask = df_prod_brut.astype(str).apply(lambda x: x.str.contains(recherche, case=False, na=False))
        prods_trouvees = df_prod_brut[mask.any(axis=1)]
        
        if not prods_trouvees.empty:
            st.success(f"🏭 **Bingo ! {len(prods_trouvees)} ligne(s) trouvée(s) contenant '{recherche}' :**")
            st.dataframe(prods_trouvees, use_container_width=True)
            st.info("👆 Vérifiez la colonne DATE_PROD. Si elle est vide ou marquée 'NaT', l'outil l'ignore par sécurité.")
        else:
            st.error(f"⚠️ Le numéro {recherche} n'existe ABSOLUMENT PAS dans les fichiers importés.")
