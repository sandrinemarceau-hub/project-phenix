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
        if st.session_state["password"] == "Logistique2026!": # <-- Changez le mot de passe ici
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

# Si le mot de passe n'est pas bon, on arrête tout ici.
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
    fichiers_prod = st.file_uploader("Glissez les 3 fichiers de Prod", type=['xlsx'], accept_multiple_files=True)

with col3:
    st.subheader("3. Fichier Commandes")
    fichier_commandes = st.file_uploader("Export_Commandes.xlsx", type=['xlsx'])

# ==========================================
# 3. LE MOTEUR DE CALCUL (S'active au clic)
# ==========================================
if st.button("🚀 Calculer les disponibilités", type="primary"):
    if fichier_stock and fichiers_prod and fichier_commandes:
        with st.spinner('Calcul en cours...'):
            try:
                # A. Lecture Stock
                df_stock = pd.read_excel(fichier_stock)
                # La ligne magique qui nettoie les en-têtes (enlève les espaces invisibles et met en majuscules)
                df_stock.columns = df_stock.columns.str.strip().str.upper()
                
                stock_actuel = df_stock.set_index('CODE ARTICLE')['STOCK DISPONIBLE'].to_dict()

                # B. Lecture Production
                liste_prod = []
                for f in fichiers_prod:
                    df_temp = pd.read_excel(f)
                    df_temp = df_temp.rename(columns={
                        'CODE ART ENTREE': 'Article', 'ARTICLE CODE AE': 'Article', 'art entree': 'Article',
                        'QTE ART ENTREE': 'Qte_Produite', 'QTE AE': 'Qte_Produite', 'qte entree': 'Qte_Produite',
                        'DATE PLANIF': 'Date_Prod', 'DATE REALISATION': 'Date_Prod', 'Date Realisation': 'Date_Prod'
                    })
                    if 'Article' in df_temp.columns and 'Qte_Produite' in df_temp.columns and 'Date_Prod' in df_temp.columns:
                        df_temp = df_temp[['Article', 'Qte_Produite', 'Date_Prod']]
                        liste_prod.append(df_temp)
                
                df_production = pd.concat(liste_prod, ignore_index=True)
                df_production['Date_Prod'] = pd.to_datetime(df_production['Date_Prod'])
                df_production['Date_Dispo_Reelle'] = df_production['Date_Prod'] + timedelta(days=2)
                df_production = df_production.sort_values(by=['Article', 'Date_Dispo_Reelle'])
                productions_futures = df_production.to_dict('records')

                # C. Lecture Commandes
                df_commandes = pd.read_excel(fichier_commandes)
                df_commandes['DATE_CDE'] = pd.to_datetime(df_commandes['DATE_CDE'])
                if 'URGENCE' not in df_commandes.columns:
                    df_commandes['URGENCE'] = 0
                df_commandes = df_commandes.sort_values(by=['URGENCE', 'DATE_CDE'], ascending=[False, True])

                # D. Calcul (Le FIFO)
                resultats = []
                for index, commande in df_commandes.iterrows():
                    article = commande['ARTICLE_CODE']
                    qte_demandee = commande['QTE_UB_CDE Total']
                    num_cmd = commande['NUM_CDE']
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
                            if prod['Article'] == article and prod['Qte_Produite'] > 0:
                                qte_prise = min(prod['Qte_Produite'], qte_restante)
                                prod['Qte_Produite'] -= qte_prise
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
                
                # E. Affichage du résultat
                st.success("Calcul terminé avec succès ! 🎉")
                st.dataframe(df_final)

                # F. Bouton de téléchargement
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
                st.error(f"Une erreur s'est produite avec vos fichiers. Vérifiez les colonnes. Détail : {e}")
    else:
        st.warning("Veuillez déposer tous les fichiers demandés avant de lancer le calcul.")
