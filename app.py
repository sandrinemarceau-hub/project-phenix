import streamlit as st
import pandas as pd
from datetime import timedelta
import io
import zipfile
import re
import math

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

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

def safe_xml(texte):
    return str(texte).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def clean_nan(val, default=""):
    if pd.isna(val) or str(val).strip().lower() in ['nan', 'nat', 'none', '']:
        return default
    return str(val).strip()

def format_num(val):
    if not isinstance(val, (int, float)): return val
    s = f"{val:.4f}".rstrip('0')
    if s.endswith('.'): s = s[:-1]
    return s if s else "0"

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
        xls = pd.ExcelFile(fichier)
        best_df = None
        max_score = -1
        mots_cles = ['ARTICLECODE', 'CODEARTICLE', 'ARTICLE', 'QUANTITE', 'QTE', 'STOCK', 'POIDS', 'LIBELLE', 'PALETTE', 'FORMAT', 'UAUEMAX', 'STOCKPHYSIQUE']
        
        for sheet in xls.sheet_names:
            try:
                df_temp = pd.read_excel(xls, sheet_name=sheet, skiprows=lignes_a_ignorer)
                cols = df_temp.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                score = sum(1 for c in cols for m in mots_cles if m in c)
                if score > max_score:
                    max_score = score
                    best_df = df_temp
            except:
                pass
        
        if best_df is not None: return best_df
        return pd.read_excel(xls, sheet_name=0, skiprows=lignes_a_ignorer)

# --- GÉNÉRATEUR PACKING LISTS REPORTLAB ---
def generer_packing_lists_zip(df_resultats, dict_details):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        commandes = df_resultats['Num_Commande'].unique()
        for cmd in commandes:
            if str(cmd).upper() in ["INCONNU", "NAN"]: continue
            lignes = df_resultats[df_resultats['Num_Commande'] == cmd]
            
            client = str(lignes.iloc[0]['Client'])
            adresse = clean_nan(lignes.iloc[0]['Adresse'])
            ville = clean_nan(lignes.iloc[0]['Ville'])
            pays = clean_nan(lignes.iloc[0]['Pays'])
            exportateur = clean_nan(lignes.iloc[0]['Exportateur']).upper()
            
            txt_exp = "<b>EXPORTER:</b><br/>LUC BELAIRE INTERNATIONAL, LTD<br/>DUBLIN, IRELAND"
            if "FRANCE" in exportateur or "SOVEREIGN" in exportateur:
                txt_exp = "<b>EXPORTER:</b><br/>SOVEREIGN BRANDS FRANCE<br/>10 RUE DE LA LOGISTIQUE<br/>75000 PARIS, FRANCE"
            elif "USA" in exportateur or "AMERICA" in exportateur:
                txt_exp = "<b>EXPORTER:</b><br/>SOVEREIGN BRANDS USA<br/>123 BROADWAY AVE<br/>NEW YORK, NY 10001, USA"

            consignee_lines = [f"<b>CONSIGNEE:</b><br/>{safe_xml(client)}"]
            if adresse: consignee_lines.append(safe_xml(adresse))
            if ville: consignee_lines.append(safe_xml(ville))
            if pays: consignee_lines.append(safe_xml(pays))
            txt_con = "<br/>".join(consignee_lines)
            
            pdf_buffer = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, margin=1.2*cm)
            elements = []
            styles = getSampleStyleSheet()

            style_desc = ParagraphStyle('Desc', parent=styles['Normal'], fontSize=8.5, leading=11)
            style_header = ParagraphStyle('H', parent=styles['Normal'], fontSize=9, leading=11)
            style_title = ParagraphStyle('T', parent=styles['Title'], fontSize=22, alignment=2)

            elements.append(Paragraph("PACKING LIST", style_title))
            elements.append(HRFlowable(width=18.5*cm, thickness=1.5, color=colors.black, spaceAfter=15, hAlign='CENTER'))

            t_adr = Table([[Paragraph(txt_exp, style_header), "", Paragraph(txt_con, style_header)]], colWidths=[8.5*cm, 1.5*cm, 8.5*cm])
            t_adr.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('BOTTOMPADDING',(0,0),(-1,-1), 15)]))
            elements.append(t_adr)

            headers = ['SKU / REF', 'CASES (COLIS)', 'UNIT QTY', 'DESCRIPTION']
            data = [headers]
            
            t_q, t_c, t_p, t_pal = 0, 0, 0.0, 0.0
            type_pal_label = "N/A"

            for _, row in lignes.iterrows():
                art = str(row['Article'])
                qte = int(row['Qte_Demandée'])
                
                d = dict_details.get(art, {
                    'libelle': 'Inconnu', 'format': '', 'degres': '', 'couleur': '',
                    'uc': 6.0, 'poids': 0.0, 'type_pal': 'N/A', 'cas_pal': 100.0
                })
                
                uc = d['uc'] if d['uc'] > 0 else 6.0
                cartons = int(qte / uc) if qte > 0 else 0
                
                poids_ligne = qte * d['poids']
                cas_pal = d['cas_pal'] if d['cas_pal'] > 0 else 100.0
                palettes_ligne = cartons / cas_pal if cartons > 0 else 0
                
                if d['type_pal'] not in ["N/A", "", "NAN"]:
                    type_pal_label = d['type_pal']
                
                t_q += qte
                t_c += cartons
                t_p += poids_ligne
                t_pal += palettes_ligne
                
                desc_html = f"<b>{safe_xml(d['libelle'])}</b><br/>"
                
                sub1 = []
                if d['format']: sub1.append(f"Fmt: {safe_xml(d['format'])}")
                if d['degres']: sub1.append(f"Vol: {safe_xml(d['degres'])}%")
                if d['couleur']: sub1.append(f"Coul: {safe_xml(d['couleur'])}")
                if sub1: desc_html += f"<font color='#333333'>{' | '.join(sub1)}</font><br/>"
                
                sub2 = []
                sub2.append(f"Carton: {int(uc)} btls")
                sub2.append(f"Palette: {int(cas_pal)} ctns")
                if d['poids'] > 0: sub2.append(f"Poids: {format_num(d['poids'])} kg/btl")
                desc_html += f"<font color='#666666'>{' | '.join(sub2)}</font>"
                
                data.append([
                    safe_xml(art),
                    str(cartons),
                    str(int(qte)),
                    Paragraph(desc_html, style_desc)
                ])

            t_art = Table(data, colWidths=[3*cm, 3*cm, 3*cm, 9.5*cm], repeatRows=1)
            t_art.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.black),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('ALIGN', (0,0), (2,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.2, colors.grey),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ]))
            elements.append(t_art)

            elements.append(Spacer(1, 20))
            t_pal_int = int(math.ceil(t_pal)) 
            
            tot_data = [
                [f"TOTAL UNITS: {int(t_q)}", f"TOTAL WEIGHT: {format_num(t_p)} kg"],
                [f"TOTAL CASES: {int(t_c)}", f"TOTAL PALLETS: {t_pal_int} ({type_pal_label})"],
            ]
            t_tot = Table(tot_data, colWidths=[9*cm, 9.5*cm])
            t_tot.setStyle(TableStyle([
                ('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'),
                ('FONTSIZE',(0,0),(-1,-1),10),
                ('BOX',(0,0),(-1,-1),1.5,colors.black),
                ('LEFTPADDING',(0,0),(-1,-1), 10),
                ('TOPPADDING',(0,0),(-1,-1), 10),
                ('BOTTOMPADDING',(0,0),(-1,-1), 10),
            ]))
            elements.append(t_tot)

            elements.append(Spacer(1, 40))
            elements.append(Paragraph("________________________________<br/>Authorized Signature & Stamp", styles['Normal']))

            doc.build(elements)
            safe_name = str(cmd).replace('/', '_').replace('\\', '_')
            zip_file.writestr(f"Packing_List_{safe_name}.pdf", pdf_buffer.getvalue())
            
    return zip_buffer.getvalue()

# ==========================================
# 2. INTERFACE VISUELLE
# ==========================================
st.set_page_config(layout="wide", page_title="Portail Logistique V27")
st.title("📦 Portail de Disponibilité - VERSION 27 🔴")
st.write("Mode 'Cerveau Fusionné' : Déposez plusieurs fichiers Nomenclatures en même temps !")

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
    # NOUVEAU : Accept Multiple Files pour la nomenclature !
    st.subheader("4. Nomenclatures")
    fichiers_nom = st.file_uploader("Fichiers (Poids & Liens)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True)
    skip_nom = st.number_input("Ignorer (Nom.)", min_value=0, value=0)

# ==========================================
# 3. LE MOTEUR DE CALCUL
# ==========================================
st.divider()

if st.button("🚀 Calculer les disponibilités (V27)", type="primary", use_container_width=True):
    if fichier_stock and fichiers_prod and fichier_commandes:
        with st.spinner('Analyse et fusion des nomenclatures en cours...'):
            try:
                # --- A. LECTURE NOMENCLATURES MULTIPLES ---
                dict_prepa = {}
                dict_details = {}
                
                # Le Scanner a besoin d'une vue d'ensemble : on va concaténer les nomenclatures lues
                df_nom_scanner = pd.DataFrame()

                if fichiers_nom:
                    for f_nom in fichiers_nom:
                        df_nom_brut = lire_fichier(f_nom, skip_nom)
                        
                        # Ajout au Scanner
                        df_nom_scanner = pd.concat([df_nom_scanner, df_nom_brut.copy()], ignore_index=True)
                        
                        df_nom_brut.columns = df_nom_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                        
                        c_art = next((c for c in ['ARTICLECODE', 'CODEARTICLE'] if c in df_nom_brut.columns), None)
                        c_prepa = next((c for c in ['ARTPREPA', 'PRODUITDEBASECODE'] if c in df_nom_brut.columns), None)
                        c_lib = next((c for c in ['ARTICLELIBELLE', 'LIBELLE', 'DESCRIPTION', 'DESCRIPTIONARTICLE'] if c in df_nom_brut.columns), None)
                        c_fmt = next((c for c in ['FORMAT'] if c in df_nom_brut.columns), None)
                        c_uc = next((c for c in ['UCUA', 'UC', 'PCB'] if c in df_nom_brut.columns), None)
                        c_poids = next((c for c in ['POIDSBTLLES', 'POIDS', 'WEIGHT'] if c in df_nom_brut.columns), None)
                        c_pal_type = next((c for c in ['PALETTE', 'TYPEPALETTE'] if c in df_nom_brut.columns), None)
                        c_cas_pal = next((c for c in ['UAUEMAX', 'PAL', 'CASESPERPALLET'] if c in df_nom_brut.columns), None)
                        c_degres = next((c for c in ['DEGRES', 'DEGRE'] if c in df_nom_brut.columns), None)
                        c_couleur = next((c for c in ['COULEUR', 'COLOR'] if c in df_nom_brut.columns), None)
                        
                        if c_art:
                            df_nom_brut['CLEAN_ART'] = nettoyage_extreme(df_nom_brut[c_art])
                            if c_prepa: df_nom_brut['CLEAN_PREPA'] = nettoyage_extreme(df_nom_brut[c_prepa])
                                
                            for _, r in df_nom_brut.iterrows():
                                art_id = str(r['CLEAN_ART'])
                                prepa_id = str(r['CLEAN_PREPA']) if c_prepa else ""
                                
                                # Ajout du lien de parenté
                                if prepa_id and prepa_id not in ["0", "NAN", art_id]:
                                    dict_prepa[art_id] = prepa_id
                                
                                # Initialisation sécurisée
                                if art_id not in dict_details:
                                    dict_details[art_id] = {
                                        'libelle': 'Inconnu', 'format': '', 'degres': '', 'couleur': '',
                                        'uc': 6.0, 'poids': 0.0, 'type_pal': 'N/A', 'cas_pal': 100.0
                                    }
                                
                                # Mise à jour conditionnelle (on écrase si la donnée est valide)
                                if c_lib:
                                    val = clean_nan(r[c_lib])
                                    if val and val != "NAN": dict_details[art_id]['libelle'] = val
                                if c_fmt:
                                    val = clean_nan(r[c_fmt])
                                    if val and val != "NAN": dict_details[art_id]['format'] = val
                                if c_degres:
                                    val = clean_nan(r[c_degres])
                                    if val and val != "NAN": dict_details[art_id]['degres'] = val
                                if c_couleur:
                                    val = clean_nan(r[c_couleur])
                                    if val and val != "NAN": dict_details[art_id]['couleur'] = val
                                if c_uc:
                                    val = float(nettoyage_quantite(pd.Series([r[c_uc]]))[0])
                                    if val > 0: dict_details[art_id]['uc'] = val
                                if c_poids:
                                    val = float(nettoyage_quantite(pd.Series([r[c_poids]]))[0])
                                    if val > 0: dict_details[art_id]['poids'] = val
                                if c_pal_type:
                                    val = clean_nan(r[c_pal_type])
                                    if val and val not in ["NAN", "N/A"]: dict_details[art_id]['type_pal'] = val
                                if c_cas_pal:
                                    val = float(nettoyage_quantite(pd.Series([r[c_cas_pal]]))[0])
                                    if val > 0: dict_details[art_id]['cas_pal'] = val
                
                st.session_state['dict_details'] = dict_details
                st.session_state['df_nom_brut'] = df_nom_scanner

                # --- B. LECTURE STOCK ---
                df_stock_brut = lire_fichier(fichier_stock, skip_stock)
                st.session_state['df_stock_brut'] = df_stock_brut.copy() 
                
                df_stock_brut.columns = df_stock_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                col_art_stock = next((c for c in ['CODEARTICLE', 'ARTICLECODE', 'ARTICLE', 'REFERENCE', 'CODE'] if c in df_stock_brut.columns), None)
                col_qte_stock = next((c for c in ['STOCKPHYSIQUE', 'STOCKDISPONIBLE', 'QTESTOCK', 'QUANTITE', 'STOCK', 'TOTAL', 'TOTALGNRAL', 'TOTALGENERAL'] if c in df_stock_brut.columns), None)
                
                if not col_art_stock or not col_qte_stock:
                    st.error("❌ Erreur STOCK : Colonnes introuvables.")
                    st.stop()
                
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

                    colonnes_temp = df_temp.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                    df_temp.columns = colonnes_temp
                    
                    liste_articles_prod = ['ARTICLECODEAE', 'CODEARTENTREE', 'ARTENTREE', 'ARTICLECODE', 'CODEARTICLE', 'ARTICLE', 'REFERENCE', 'CODE', 'ARTPREPA', 'CODEPREPA', 'PRODUIT']
                    liste_qtes_prod = ['QTEAE', 'QTEARTENTREE', 'QTEENTREE', 'QUANTITE', 'QTE', 'TOTAL', 'TOTALGNRAL', 'TOTALGENERAL', 'QTEPREVUE', 'QUANTITEPREVUE', 'RESTEAFAIRE', 'QTEFABRIQUEE']
                    
                    col_art_prod = next((c for c in liste_articles_prod if c in colonnes_temp), None)
                    col_qte_prod = next((c for c in liste_qtes_prod if c in colonnes_temp), None)
                    
                    if col_art_prod and col_qte_prod:
                        df_ext = pd.DataFrame()
                        df_ext['ARTICLE'] = nettoyage_extreme(df_temp[col_art_prod])
                        df_ext['QTE_PRODUITE'] = nettoyage_quantite(df_temp[col_qte_prod])
                        
                        date_series = None
                        for col in ['DATEREALISATION', 'DATEPLANIF', 'DATEFIN', 'DATEPREVUE', 'ECHEANCE', 'DATE']:
                            if col in colonnes_temp:
                                s_test = pd.to_datetime(df_temp[col], dayfirst=True, errors='coerce')
                                date_series = s_test if date_series is None else date_series.fillna(s_test)
                                    
                        if date_series is None or date_series.isna().all():
                            for col in colonnes_temp:
                                if 'DATE' in col:
                                    s_test = pd.to_datetime(df_temp[col], dayfirst=True, errors='coerce')
                                    date_series = s_test if date_series is None else date_series.fillna(s_test)
                                        
                        df_ext['DATE_PROD'] = date_series if date_series is not None else pd.Series(pd.NaT, index=df_temp.index)
                        liste_prod.append(df_ext)
                    else:
                        st.warning(f"⚠️ Alerte Fichier Ignoré : Le fichier '{f.name}' n'a pas pu être lu. Je ne trouve pas la colonne d'Article ou de Quantité. \n\n 🔍 Voici les colonnes que j'ai trouvées dedans : {colonnes_temp.tolist()}")
                
                st.session_state['df_prod_brut'] = df_prod_brut_total 
                if liste_prod:
                    df_production = pd.concat(liste_prod, ignore_index=True)
                    df_production_valide = df_production.dropna(subset=['DATE_PROD']).copy()
                    df_production_valide = df_production_valide[df_production_valide['QTE_PRODUITE'] > 0]
                    
                    df_production_valide['Date_Dispo_Reelle'] = df_production_valide['DATE_PROD'] + timedelta(days=2)
                    df_production_valide = df_production_valide.sort_values(by=['ARTICLE', 'Date_Dispo_Reelle'])
                    productions_futures = df_production_valide.to_dict('records')
                else:
                    productions_futures = []

                # --- D. LECTURE COMMANDES ---
                df_commandes_brut = lire_fichier(fichier_commandes, skip_cmd)
                colonnes_cmd = df_commandes_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                df_commandes_brut.columns = colonnes_cmd
                
                col_art_cmd = next((c for c in ['ARTICLECODE', 'CODEARTICLE', 'ARTICLE'] if c in colonnes_cmd), None)
                col_date_cmd = next((c for c in ['DATECDE', 'DATECOMMANDE', 'DATECREATION', 'DATE'] if c in colonnes_cmd), None)
                col_qte_cmd = next((c for c in ['QTEUBCDETOTAL', 'QTEU
