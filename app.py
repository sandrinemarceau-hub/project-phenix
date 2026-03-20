import streamlit as st
import pandas as pd
from datetime import timedelta
import io

# ==========================================
# 1. SYSTÈME DE SÉCURITÉ (Mot de passe)
# ==========================================
def check_password():
    """Vérifie si l'utilisateur a le bon mot de passe."""
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
# 2. INTERFACE VISUELLE PRINCIPALE
# ==========================================
st.title("📦 Portail de Disponibilité des Commandes")
st.write("Bienvenue ! Déposez vos exports ci-dessous pour calculer les dates de disponibilité.")

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. Fichier Stocks")
    fichier_stock = st.file_uploader("Export_Stock.xlsx", type=['xlsx'])

with col2:
    st.subheader("2. Fichiers Production")
    fichiers_prod = st.file_uploader("Glissez les fichiers (Max 3)", type=['xlsx'], accept_multiple_files=True)

with col3:
    st.subheader("3. Fichier Commandes")
    fichier_commandes = st.file_uploader("Export_Commandes.xlsx", type=['xlsx'])

# ==========================================
# 3. LE MOTEUR DE CALCUL
# ==========================================
if st.button("🚀 Calculer les disponibilités", type="primary"):
    if fichier_stock and fichiers_prod and fichier_commandes:
        with st.spinner('Calcul en cours...'):
            try:
                # --- A. LECTURE STOCK ---
                df_stock = pd.read_excel(fichier_stock)
                
                # NETTOYAGE EXTRÊME : tout en majuscules, suppression des espaces invisibles
                df_stock.columns = df_stock.columns.astype(str).str.strip().str.upper()
                
                # DIAGNOSTIC AUTOMATIQUE SI ERREUR
                if 'CODE ARTICLE' not in df_stock.columns:
                    st.error("❌ Erreur dans le fichier STOCK : La colonne 'CODE ARTICLE' est introuvable.")
                    st.info(f"🔍 Voici les colonnes que l'outil a trouvées dans votre Excel (Ligne 1) : {df_stock.columns.tolist()}")
                    st.warning("👉 Si la liste ci-dessus affiche des choses comme 'Unnamed' ou 'Nom de la société', c'est que votre tableau ne commence pas à la Ligne 1 d'Excel.")
                    st.stop()

                stock_actuel = df_stock.set_index('CODE ARTICLE')['STOCK DISPONIBLE'].to_dict()

                # --- B. LECTURE PRODUCTION ---
                liste_prod = []
                for f in fichiers_prod:
                    df_temp = pd.read_excel(f)
                    df_temp.columns = df_temp.columns.astype(str).str.strip().str.upper() # Nettoyage Extreme
                    
                    # Dictionnaire de traduction
                    colonnes_a_renommer = {
                        'CODE ART ENTREE': 'ARTICLE', 'ARTICLE CODE AE': 'ARTICLE', 'ART ENTREE': 'ARTICLE',
                        'QTE ART ENTREE': 'QTE_PRODUITE', 'QTE AE': 'QTE_PRODUITE', 'QTE ENTREE': 'QTE_PRODUITE',
                        'DATE PLANIF': 'DATE_PROD', 'DATE REALISATION': 'DATE_PROD'
                    }
                    df_temp = df_temp.rename(columns=colonnes_a_renommer)
                    
                    if 'ARTICLE' in df_temp.columns and 'QTE_PRODUITE' in df_temp.columns and 'DATE_PROD' in df_temp.columns:
                        df_temp = df_temp[['ARTICLE', 'QTE_PRODUITE', 'DATE_PROD']]
                        liste_prod.append(df_temp)
                
                if not liste_prod:
                    st.error("❌ Erreur dans les fichiers de PRODUCTION : Colonnes introuvables.")
                    st.stop()

                df_production = pd.concat(liste_prod, ignore_index=True)
                df_production['DATE_PROD'] = pd.to_datetime(df_production['DATE_PROD'], errors='coerce')
                df_production['Date_Dispo_Reelle'] = df_production['DATE_PROD'] + timedelta(days=2)
                df_production = df_production.sort_values(by=['ARTICLE', 'Date_Dispo_Reelle'])
                productions_futures = df_production.to_dict('records')

                # --- C. LECTURE COMMANDES ---
                df_commandes = pd.read_excel(fichier_commandes)
                df_commandes.columns = df_commandes.columns.astype(str).str.strip().str.upper() # Nettoyage extreme
                
                if 'ARTICLE_CODE' not in df_commandes.columns:
                    st.error("❌ Erreur dans le fichier COMMANDES : La colonne 'ARTICLE_CODE' est introuvable.")
                    st.info(f"🔍 Voici les colonnes trouvées : {df_commandes.columns.tolist()}")
                    st.stop()

                df_commandes['DATE_CDE'] = pd.to_datetime(df_commandes['DATE_CDE'], errors='coerce')
                if 'URGENCE' not in df_commandes.columns:
                    df_commandes['URGENCE'] = 0
                df_commandes = df_commandes.sort_values(by=['URGENCE', 'DATE_CDE'], ascending=[False, True])

                # --- D. ALGORITHME D'ATTRIBUTION ---
                resultats = []
                for index, commande in df_commandes.iterrows():
                    article = commande['ARTICLE_CODE']
                    col_qte = 'QTE_UB_CDE TOTAL' if 'QTE_UB_CDE TOTAL' in df_commandes.columns else 'QTE_UB_CDE'
                    qte_demandee = commande.get(col_qte, 0)
                    num_cmd = commande['NUM_CDE']
                    qte_restante = qte_demandee
                    date_dispo = "Immédiate (En Stock)"
                    
                    # 1. Taper dans le stock
                    stock_dispo = stock_actuel.get(article, 0)
                    if stock_dispo > 0:
                        qte_prise = min(stock_dispo, qte_restante)
                        stock_actuel[article] -= qte_prise
                        qte_restante -= qte_prise
                        
                    # 2. Taper dans la prod
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
                        date_dispo = f"⚠️ Manque {qte_restante} unités"
                        
                    resultats.append({
                        'Num_Commande': num_cmd,
                        'Article': article,
                        'Client': commande.get('EXPE_NOM_CLIENT', 'Inconnu'),
                        'Quantité': qte_demandee,
                        'Date_Disponibilité': date_dispo
                    })

                df_final = pd.DataFrame(resultats)
                
                # --- E. AFFICHAGE ET TÉLÉCHARGEMENT ---
                st.success("Calcul terminé avec succès ! 🎉")
                st.dataframe(df_final)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Dispos')
                
                st.download_button(
                    label="📥 Télécharger le fichier mis à jour",
                    data=buffer,
                    file_name="Commandes_Disponibles.xlsx",
                    mime="application/vnd.ms-excel"
                )

            except Exception as e:
                st.error(f"Une erreur inattendue s'est produite : {e}")
    else:
        st.warning("Veuillez déposer tous les fichiers demandés avant de lancer le calcul.")
