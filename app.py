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
st.set_page_config(layout="wide", page_title="Portail Sovereign Brands")

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
            st.error("Mot de passe incorrect. Veuillez réessayer.")
    st.stop()

# --- BOUTON DÉCONNEXION ---
with st.sidebar:
    if st.button("🚪 Déconnexion"):
        st.session_state.clear()
        st.rerun()
    st.divider()

# ==========================================
# FONCTIONS TECHNIQUES (COMMUNES)
# ==========================================
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
            if x.rfind(',') > x.rfind('.'): x = x.replace('.', '').replace(',', '.') 
            else: x = x.replace(',', '') 
        else: x = x.replace(',', '.') 
        x = re.sub(r'[^\d.-]', '', x) 
        try: return float(x)
        except: return 0.0
    return serie.apply(clean_val)

def safe_xml(texte):
    return str(texte).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def clean_nan(val, default=""):
    if pd.isna(val) or str(val).strip().lower() in ['nan', 'nat', 'none', '']: return default
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
                if score > max_score:
                    max_score = score
                    best_df = df_temp
            except: pass
        if best_df is not None: return best_df
        return pd.read_excel(xls, sheet_name=0, skiprows=lignes_a_ignorer)

# --- GÉNÉRATEURS DE PDF ---
def generer_packing_lists_zip(df_resultats, dict_details):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        commandes = df_resultats['Num_Commande'].unique()
        for cmd in commandes:
            if str(cmd).upper() in ["INCONNU", "NAN"]: continue
            lignes = df_resultats[df_resultats['Num_Commande'] == cmd]
            client = str(lignes.iloc[0]['Client']); adresse = clean_nan(lignes.iloc[0]['Adresse']); ville = clean_nan(lignes.iloc[0]['Ville']); pays = clean_nan(lignes.iloc[0]['Pays']); exportateur = clean_nan(lignes.iloc[0]['Exportateur']).upper()
            txt_exp = "<b>EXPORTER:</b><br/>LUC BELAIRE INTERNATIONAL, LTD<br/>DUBLIN, IRELAND"
            if "FRANCE" in exportateur or "SOVEREIGN" in exportateur: txt_exp = "<b>EXPORTER:</b><br/>SOVEREIGN BRANDS FRANCE<br/>10 RUE DE LA LOGISTIQUE<br/>75000 PARIS, FRANCE"
            elif "USA" in exportateur or "AMERICA" in exportateur: txt_exp = "<b>EXPORTER:</b><br/>SOVEREIGN BRANDS USA<br/>123 BROADWAY AVE<br/>NEW YORK, NY 10001, USA"
            consignee_lines = [f"<b>CONSIGNEE:</b><br/>{safe_xml(client)}"]
            if adresse: consignee_lines.append(safe_xml(adresse))
            if ville: consignee_lines.append(safe_xml(ville))
            if pays: consignee_lines.append(safe_xml(pays))
            txt_con = "<br/>".join(consignee_lines)
            
            pdf_buffer = io.BytesIO(); doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, margin=1.2*cm); elements = []; styles = getSampleStyleSheet()
            style_desc = ParagraphStyle('Desc', parent=styles['Normal'], fontSize=8.5, leading=11); style_header = ParagraphStyle('H', parent=styles['Normal'], fontSize=9, leading=11); style_title = ParagraphStyle('T', parent=styles['Title'], fontSize=22, alignment=2)
            elements.append(Paragraph("PACKING LIST", style_title)); elements.append(HRFlowable(width=18.5*cm, thickness=1.5, color=colors.black, spaceAfter=15, hAlign='CENTER'))
            t_adr = Table([[Paragraph(txt_exp, style_header), "", Paragraph(txt_con, style_header)]], colWidths=[8.5*cm, 1.5*cm, 8.5*cm]); t_adr.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('BOTTOMPADDING',(0,0),(-1,-1), 15)])); elements.append(t_adr)

            data = [['SKU / REF', 'CASES (COLIS)', 'UNIT QTY', 'DESCRIPTION']]; t_q, t_c, t_p, t_pal = 0, 0, 0.0, 0.0; type_pal_label = "N/A"
            for _, row in lignes.iterrows():
                art = str(row['Article']); qte = int(row['Qte_Demandée'])
                d = dict_details.get(art, {'libelle': 'Inconnu', 'format': '', 'degres': '', 'couleur': '', 'uc': 6.0, 'poids': 0.0, 'type_pal': 'N/A', 'cas_pal': 100.0})
                uc = d['uc'] if d['uc'] > 0 else 6.0; cartons = int(qte / uc) if qte > 0 else 0; poids_ligne = qte * d['poids']; cas_pal = d['cas_pal'] if d['cas_pal'] > 0 else 100.0; palettes_ligne = cartons / cas_pal if cartons > 0 else 0
                if d['type_pal'] not in ["N/A", "", "NAN"]: type_pal_label = d['type_pal']
                t_q += qte; t_c += cartons; t_p += poids_ligne; t_pal += palettes_ligne
                
                desc_html = f"<b>{safe_xml(d['libelle'])}</b><br/>"
                sub1 = []
                if d['format']: sub1.append(f"Fmt: {safe_xml(d['format'])}")
                if d['degres']: sub1.append(f"Vol: {safe_xml(d['degres'])}%")
                if d['couleur']: sub1.append(f"Coul: {safe_xml(d['couleur'])}")
                if sub1: desc_html += f"<font color='#333333'>{' | '.join(sub1)}</font><br/>"
                sub2 = [f"Carton: {int(uc)} btls", f"Palette: {int(cas_pal)} ctns"]
                if d['poids'] > 0: sub2.append(f"Poids: {format_num(d['poids'])} kg/btl")
                desc_html += f"<font color='#666666'>{' | '.join(sub2)}</font>"
                data.append([safe_xml(art), str(cartons), str(int(qte)), Paragraph(desc_html, style_desc)])

            t_art = Table(data, colWidths=[3*cm, 3*cm, 3*cm, 9.5*cm], repeatRows=1)
            t_art.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.black), ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), ('ALIGN', (0,0), (2,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.2, colors.grey), ('TOPPADDING', (0,0), (-1,-1), 8), ('BOTTOMPADDING', (0,0), (-1,-1), 8)]))
            elements.append(t_art); elements.append(Spacer(1, 20))
            
            tot_data = [[f"TOTAL UNITS: {int(t_q)}", f"TOTAL WEIGHT: {format_num(t_p)} kg"], [f"TOTAL CASES: {int(t_c)}", f"TOTAL PALLETS: {int(math.ceil(t_pal))} ({type_pal_label})"]]
            t_tot = Table(tot_data, colWidths=[9*cm, 9.5*cm])
            t_tot.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'), ('FONTSIZE',(0,0),(-1,-1),10), ('BOX',(0,0),(-1,-1),1.5,colors.black), ('LEFTPADDING',(0,0),(-1,-1), 10), ('TOPPADDING',(0,0),(-1,-1), 10), ('BOTTOMPADDING',(0,0),(-1,-1), 10)]))
            elements.append(t_tot); elements.append(Spacer(1, 40)); elements.append(Paragraph("________________________________<br/>Authorized Signature & Stamp", styles['Normal']))
            doc.build(elements); safe_name = str(cmd).replace('/', '_').replace('\\', '_'); zip_file.writestr(f"Packing_List_{safe_name}.pdf", pdf_buffer.getvalue())
    return zip_buffer.getvalue()

if FPDF_OK:
    class RDVPDF(FPDF):
        def header(self):
            self.ln(35); self.set_font("Helvetica", "B", 24); self.cell(0, 15, 'RDV DOCUMENT', 0, 1, 'C'); self.ln(2)
        def get_lines_count(self, w, line_height, text):
            try: return len(self.multi_cell(w, line_height, text, split_only=True))
            except:
                try: return len(self.multi_cell(w, line_height, text, dry_run=True, output="LINES"))
                except: return max(1, math.ceil(self.get_string_width(text) / (w - 2)))
        def draw_harmonized_row(self, label, value):
            label = str(label).replace("’", "'").replace("–", "-"); value = str(value).replace("’", "'").replace("–", "-")
            w_label, w_value, marge_x, line_height = 75, 105, 15, 6
            self.set_font("Helvetica", "", 10); lines_label = self.get_lines_count(w_label, line_height, label)
            self.set_font("Helvetica", "B", 10); lines_value = self.get_lines_count(w_value, line_height, value)
            total_h = max(max(lines_label, lines_value) * line_height + 4, 12); x_curr, y_curr = marge_x, self.get_y()
            self.set_xy(x_curr, y_curr); self.cell(w_label, total_h, "", border=1); self.cell(w_value, total_h, "", border=1)
            self.set_font("Helvetica", "", 10); self.set_xy(x_curr, y_curr + (total_h - lines_label * line_height) / 2); self.multi_cell(w_label, line_height, label, align='C')
            self.set_font("Helvetica", "B", 10); self.set_xy(x_curr + w_label, y_curr + (total_h - lines_value * line_height) / 2); self.multi_cell(w_value, line_height, value, align='C')
            self.set_xy(marge_x, y_curr + total_h)

def generer_rdv_documents_zip(df_resultats, dict_details, app_settings):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        commandes = df_resultats['Num_Commande'].unique()
        for cmd in commandes:
            if str(cmd).upper() in ["INCONNU", "NAN"]: continue
            lignes = df_resultats[df_resultats['Num_Commande'] == cmd]
            pire_date_obj = None; en_rupture = False
            for _, r in lignes.iterrows():
                statut = r['Statut']; d_str = r['Date_Disponibilité'].replace(" (Partiel)", "")
                if statut == "Rupture": en_rupture = True
                elif statut in ["Attente Prod", "Attente Prod (Partiel)"] and d_str != "Pas de date":
                    try:
                        d_obj = datetime.strptime(d_str, "%d/%m/%Y")
                        if pire_date_obj is None or d_obj > pire_date_obj: pire_date_obj = d_obj
                    except: pass
            if en_rupture: date_finale = "A DEFINIR (Rupture Partielle)"
            elif pire_date_obj: date_finale = pire_date_obj.strftime("%d/%m/%Y")
            else: date_finale = "ASAP (En Stock)"

            t_poids = 0.0; t_palettes = 0.0
            for _, r in lignes.iterrows():
                art = str(r['Article']); qte = int(r['Qte_Demandée'])
                d = dict_details.get(art, {'uc': 6.0, 'poids': 0.0, 'cas_pal': 100.0})
                uc = d['uc'] if d['uc'] > 0 else 6.0; cas_pal = d['cas_pal'] if d['cas_pal'] > 0 else 100.0
                cartons = qte / uc if qte > 0 else 0; t_poids += qte * d['poids']; t_palettes += cartons / cas_pal if cartons > 0 else 0

            client = str(lignes.iloc[0]['Client']); pays = clean_nan(lignes.iloc[0]['Pays']); exportateur = clean_nan(lignes.iloc[0]['Exportateur']).upper()
            adresse_enlevement = app_settings['adresse_defaut']
            if "VEUVE" in exportateur or "AMBAL" in exportateur: adresse_enlevement = app_settings['adresse_veuve']

            if FPDF_OK:
                pdf = RDVPDF(); pdf.add_page(); pdf.set_font("Helvetica", "B", 14)
                txt_noir = "Available for collection on : "; txt_rouge = date_finale
                largeur_totale = pdf.get_string_width(txt_noir) + pdf.get_string_width(txt_rouge)
                pdf.set_x((pdf.w - largeur_totale) / 2); pdf.set_text_color(0, 0, 0); pdf.cell(pdf.get_string_width(txt_noir), 10, txt_noir); pdf.set_text_color(200, 0, 0); pdf.cell(pdf.get_string_width(txt_rouge), 10, txt_rouge, 0, 1); pdf.set_text_color(0, 0, 0); pdf.ln(10)
                pdf.draw_harmonized_row("Pick Up address / Adresse d'enlèvement", adresse_enlevement)
                pdf.draw_harmonized_row("Loading Hours / Horaires d'ouverture", app_settings['horaires'])
                pdf.draw_harmonized_row("Contact", app_settings['contact'])
                pdf.draw_harmonized_row("Order number / Numéro de commande", str(cmd))
                pdf.draw_harmonized_row("Country of delivery", pays)
                pdf.draw_harmonized_row("Customer / Client", client)
                pdf.draw_harmonized_row("Number and size of pallets /\nNombre et dimensions des palettes", f"{int(math.ceil(t_palettes))} Palettes")
                pdf.draw_harmonized_row("Total Weight / Poids", f"{format_num(t_poids)} KG")
                pdf.draw_harmonized_row("Shipping costs / Frais de port", "-")
                pdf.ln(15); pdf.set_font("Helvetica", "B", 8.5); pdf.set_text_color(200, 0, 0)
                w_en = "Reminder : we need a 48 hours delay to prepare the order before collection. ALL SHIPPER COMING WITHOUT AN APPOINTMENT AND NOT RESPECTING OUR 48 HOURS DELAY WILL BE REFUSED AND NOT LOADED."
                w_fr = "Pour rappel : un délai de 48h est nécessaire afin que notre entrepôt prépare la commande avant le chargement. TOUT TRANSPORTEUR SE PRÉSENTANT SANS RDV ET SANS RESPECTER CE DÉLAI SERA REFUSÉ ET NON CHARGÉ."
                pdf.multi_cell(0, 5, w_en, align='C'); pdf.ln(5); pdf.multi_cell(0, 5, w_fr, align='C')
                safe_name = str(cmd).replace('/', '_').replace('\\', '_'); pdf_bytes = pdf.output(dest="S").encode("latin-1"); zip_file.writestr(f"RDV_{safe_name}.pdf", pdf_bytes)
    return zip_buffer.getvalue()

# ==========================================
# ESPACE ADMINISTRATEUR (BACK OFFICE)
# ==========================================
if st.session_state['role'] == 'admin':
    with st.sidebar:
        st.write("📝 **Paramètres des PDF RDV**")
        pdf_contact = st.text_input("Contact Email", "logistique@sovereignbrands.com")
        pdf_horaires = st.text_input("Horaires", "08:00 - 16:00 (Du Lundi au Vendredi)")
        pdf_adresse_defaut = st.text_area("Adresse Enlèvement (MGC)", "MGC\nZone Industrielle\n21200 Beaune", height=100)
        pdf_adresse_veuve = st.text_area("Adresse Enlèvement (Veuve Ambal)", "VEUVE AMBAL\n32 rue de la Croix Clément\n71530 Champforgeuil", height=100)
    settings_pdf = {'contact': pdf_contact, 'horaires': pdf_horaires, 'adresse_defaut': pdf_adresse_defaut, 'adresse_veuve': pdf_adresse_veuve}

    st.title("🛠️ Back Office - V48 (Dashboard & Filtres)")
    st.write("Importez vos fichiers usine ici. Les résultats seront sauvegardés pour les clients.")

    col1, col2, col3, col4 = st.columns(4)
    with col1: fichier_stock = st.file_uploader("Fichier Stock", type=['xlsx', 'xls', 'csv']); skip_stock = st.number_input("Ignorer (Stock)", min_value=0, value=3)
    with col2: fichiers_prod = st.file_uploader("Fichiers Prod", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True); skip_prod = st.number_input("Ignorer (Prod)", min_value=0, value=0)
    with col3: fichier_commandes = st.file_uploader("Fichier Cmds", type=['xlsx', 'xls', 'csv']); skip_cmd = st.number_input("Ignorer (Cmd)", min_value=0, value=0)
    with col4: fichiers_nom = st.file_uploader("Fichiers (Poids & Liens)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True); skip_nom = st.number_input("Ignorer (Nom.)", min_value=0, value=0)

    st.divider()

    if st.button("🚀 Calculer et Sauvegarder la Base", type="primary", use_container_width=True):
        if fichier_stock and fichiers_prod and fichier_commandes:
            with st.spinner('Analyse, Auto-Apprentissage et Sauvegarde en cours...'):
                try:
                    dict_prepa = {}; dict_details = {}; df_nom_scanner = pd.DataFrame()
                    if fichiers_nom:
                        for f_nom in fichiers_nom:
                            df_nom_brut = lire_fichier(f_nom, skip_nom); df_nom_scanner = pd.concat([df_nom_scanner, df_nom_brut.copy()], ignore_index=True); df_nom_brut.columns = df_nom_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                            c_art = next((c for c in ['ARTICLECODE', 'CODEARTICLE'] if c in df_nom_brut.columns), None); c_prepa = next((c for c in ['ARTPREPA', 'CODEPREPA', 'COMPOSANT'] if c in df_nom_brut.columns), None)
                            c_lib = next((c for c in ['ARTICLELIBELLE', 'LIBELLE', 'DESCRIPTION', 'DESCRIPTIONARTICLE'] if c in df_nom_brut.columns), None); c_fmt = next((c for c in ['FORMAT'] if c in df_nom_brut.columns), None)
                            c_uc = next((c for c in ['UCUA', 'UC', 'PCB'] if c in df_nom_brut.columns), None); c_poids = next((c for c in ['POIDSBTLLES', 'POIDS', 'WEIGHT'] if c in df_nom_brut.columns), None)
                            c_pal_type = next((c for c in ['PALETTE', 'TYPEPALETTE'] if c in df_nom_brut.columns), None); c_cas_pal = next((c for c in ['UAUEMAX', 'PAL', 'CASESPERPALLET'] if c in df_nom_brut.columns), None)
                            if c_art:
                                df_nom_brut['CLEAN_ART'] = nettoyage_extreme(df_nom_brut[c_art])
                                if c_prepa: df_nom_brut['CLEAN_PREPA'] = nettoyage_extreme(df_nom_brut[c_prepa])
                                for _, r in df_nom_brut.iterrows():
                                    art_id = str(r['CLEAN_ART']); prepa_id = str(r['CLEAN_PREPA']) if c_prepa else ""
                                    if prepa_id and prepa_id not in ["0", "NAN", "NONE", ".", art_id]: dict_prepa[art_id] = prepa_id
                                    if art_id not in dict_details: dict_details[art_id] = {'libelle': 'Inconnu', 'format': '', 'degres': '', 'couleur': '', 'uc': 6.0, 'poids': 0.0, 'type_pal': 'N/A', 'cas_pal': 100.0}
                                    if c_lib: val_lib = clean_nan(r[c_lib]); 
                                    if val_lib and val_lib != "NAN": dict_details[art_id]['libelle'] = val_lib
                                    if c_poids: val_pds = float(nettoyage_quantite(pd.Series([r[c_poids]]))[0]); 
                                    if val_pds > 0: dict_details[art_id]['poids'] = val_pds
                                    if c_cas_pal: val_pal = float(nettoyage_quantite(pd.Series([r[c_cas_pal]]))[0]); 
                                    if val_pal > 0: dict_details[art_id]['cas_pal'] = val_pal

                    df_stock_brut = lire_fichier(fichier_stock, skip_stock); mask_total = df_stock_brut.astype(str).apply(lambda x: x.str.contains('TOTAL', case=False, na=False)).any(axis=1); df_stock_brut = df_stock_brut[~mask_total]
                    df_stock_brut.columns = df_stock_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                    col_art_stock = next((c for c in ['CODEARTICLE', 'ARTICLECODE', 'ARTICLE', 'REFERENCE', 'CODE'] if c in df_stock_brut.columns), None)
                    col_qte_stock = next((c for c in ['STOCKPHYSIQUE', 'PHYSIQUE', 'STOCKDISPONIBLE', 'DISPONIBLE', 'QTEDISPO', 'QTESTOCK', 'QUANTITE', 'STOCK'] if c in df_stock_brut.columns), None)
                    if not col_art_stock or not col_qte_stock: st.error("❌ Erreur STOCK : Colonnes introuvables."); st.stop()
                    df_stock = pd.DataFrame(); df_stock['CODE_ARTICLE'] = nettoyage_extreme(df_stock_brut[col_art_stock]); df_stock['STOCK_DISPO'] = nettoyage_quantite(df_stock_brut[col_qte_stock]) if col_qte_stock else 0; stock_actuel = df_stock.groupby('CODE_ARTICLE')['STOCK_DISPO'].sum().to_dict()

                    liste_prod = []; df_prod_brut_total = pd.DataFrame() 
                    for f in fichiers_prod:
                        df_temp = lire_fichier(f, skip_prod); df_prod_brut_total = pd.concat([df_prod_brut_total, df_temp], ignore_index=True); colonnes_temp = df_temp.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True); df_temp.columns = colonnes_temp
                        col_sortie_auto = next((c for c in colonnes_temp if 'SORTIE' in c), None); col_entree_auto = next((c for c in colonnes_temp if 'ENTREE' in c or 'PREPA' in c), None)
                        if col_sortie_auto and col_entree_auto:
                            for _, r in df_temp.iterrows():
                                parent = nettoyage_extreme(pd.Series([r[col_sortie_auto]]))[0]; enfant = nettoyage_extreme(pd.Series([r[col_entree_auto]]))[0]
                                if parent and enfant and parent != enfant and parent not in ["0", "NAN", "NONE"] and enfant not in ["0", "NAN", "NONE"]:
                                    if parent not in dict_prepa: dict_prepa[parent] = enfant
                        arts_cols = [c for c in colonnes_temp if any(k in c for k in ['ART', 'CODE', 'REF', 'PRODUIT', 'COMPOSANT']) and not any(k in c for k in ['QTE', 'QUANT', 'DATE', 'ECH'])]
                        qtes_cols = [c for c in colonnes_temp if any(k in c for k in ['QTE', 'QUANT', 'RESTE', 'AFAIRE', 'BESOIN', 'LANCE', 'PREVU', 'PROD', 'ORDRE']) and not any(k in c for k in ['ART', 'CODE', 'DATE', 'ECH', 'REF'])]
                        dates_cols = [c for c in colonnes_temp if any(k in c for k in ['DATE', 'ECH', 'FIN', 'LIV', 'DISPO', 'BESOIN', 'PLANIF', 'REALISATION', 'PREVU', 'CREA', 'DELAI']) and not any(k in c for k in ['QTE', 'QUANT', 'ART', 'CODE'])]
                        if not arts_cols: continue
                        for c in arts_cols: df_temp[c] = nettoyage_extreme(df_temp[c])
                        for c in qtes_cols: df_temp[c] = nettoyage_quantite(df_temp[c])
                        df_temp['OMNI_DATE'] = pd.NaT
                        for c in dates_cols: df_temp['OMNI_DATE'] = df_temp['OMNI_DATE'].fillna(pd.to_datetime(df_temp[c], dayfirst=True, errors='coerce'))
                        if qtes_cols: df_temp['OMNI_QTE'] = df_temp[qtes_cols].max(axis=1)
                        else: df_temp['OMNI_QTE'] = 0
                        for idx, row in df_temp.iterrows():
                            qte = row.get('OMNI_QTE', 0); d = row.get('OMNI_DATE')
                            if pd.notna(d):
                                if qte <= 0: qte = 99999
                                for c in arts_cols:
                                    code = str(row[c])
                                    if code and code not in ["0", "NAN", "NONE"]: liste_prod.append({'ARTICLE': code, 'QTE_PRODUITE': qte, 'DATE_PROD': d})
                    if liste_prod:
                        df_production = pd.DataFrame(liste_prod); df_production['Date_Dispo_Reelle'] = df_production['DATE_PROD'] + timedelta(days=2); df_production = df_production.sort_values(by=['ARTICLE', 'Date_Dispo_Reelle'])
                        productions_futures = df_production.to_dict('records')
                    else: productions_futures = []

                    df_commandes_brut = lire_fichier(fichier_commandes, skip_cmd); colonnes_cmd = df_commandes_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True); df_commandes_brut.columns = colonnes_cmd
                    col_art_cmd = next((c for c in ['ARTICLECODE', 'CODEARTICLE', 'ARTICLE'] if c in colonnes_cmd), None); col_date_cmd = next((c for c in ['DATECDE', 'DATECOMMANDE', 'DATECREATION', 'DATE'] if c in colonnes_cmd), None)
                    col_qte_cmd = next((c for c in ['QTEUBCDETOTAL', 'QTEUBCDE', 'QUANTITE', 'QTE', 'TOTAL', 'TOTALGNRAL', 'TOTALGENERAL'] if c in colonnes_cmd), None); col_num_cmd = next((c for c in ['NUMCDE', 'NUMCOMMANDE', 'COMMANDE'] if c in colonnes_cmd), None)
                    col_client = next((c for c in ['EXPENOMCLIENT', 'CLIENT', 'NOMCLIENT'] if c in colonnes_cmd), None); col_adresse = next((c for c in colonnes_cmd if 'ADRESSE' in c or 'ADR' in c), None)
                    col_ville = next((c for c in colonnes_cmd if 'VILLE' in c or 'CITY' in c), None); col_pays = next((c for c in colonnes_cmd if 'PAYS' in c or 'COUNTRY' in c), None); col_exportateur = next((c for c in colonnes_cmd if 'EXPORT' in c or 'SOCIETE' in c or 'FILIALE' in c or 'STEAPP' in c), None)
                    df_commandes = pd.DataFrame(); df_commandes['ARTICLE_CODE'] = nettoyage_extreme(df_commandes_brut[col_art_cmd]); df_commandes['DATE_CDE'] = pd.to_datetime(df_commandes_brut[col_date_cmd], dayfirst=True, errors='coerce'); df_commandes['QUANTITE'] = nettoyage_quantite(df_commandes_brut[col_qte_cmd]); df_commandes['NUM_CDE'] = df_commandes_brut[col_num_cmd] if col_num_cmd else 'Inconnu'; df_commandes['CLIENT'] = df_commandes_brut[col_client] if col_client else 'Inconnu'; df_commandes['ADRESSE'] = df_commandes_brut[col_adresse] if col_adresse else ""; df_commandes['VILLE'] = df_commandes_brut[col_ville] if col_ville else ""; df_commandes['PAYS'] = df_commandes_brut[col_pays] if col_pays else ""; df_commandes['EXPORTATEUR'] = df_commandes_brut[col_exportateur] if col_exportateur else "DEFAUT"
                    df_commandes = df_commandes.dropna(subset=['DATE_CDE']).sort_values(by=['DATE_CDE'])

                    def get_cascade_prepas(art_code):
                        cascade = []; courant = dict_prepa.get(art_code)
                        for _ in range(5):
                            if courant and courant not in cascade: cascade.append(courant); courant = dict_prepa.get(courant)
                            else: break
                        return cascade

                    resultats = []
                    for index, commande in df_commandes.iterrows():
                        article = commande['ARTICLE_CODE']; qte_restante = commande['QUANTITE']; qte_prise_stock = 0; qte_prise_prod = 0; dates_trouvees = []
                        def consommer(code_a_chercher, qte_a_trouver):
                            q_stk, q_prd = 0, 0; s = stock_actuel.get(code_a_chercher, 0)
                            if s > 0: prise = min(s, qte_a_trouver); stock_actuel[code_a_chercher] -= prise; q_stk += prise; qte_a_trouver -= prise
                            if qte_a_trouver > 0:
                                for prod in productions_futures:
                                    match = False
                                    if prod['ARTICLE'] == code_a_chercher or str(prod['ARTICLE']).startswith(code_a_chercher) or code_a_chercher in str(prod['ARTICLE']): match = True
                                    if match and prod['QTE_PRODUITE'] > 0:
                                        prise = min(prod['QTE_PRODUITE'], qte_a_trouver); prod['QTE_PRODUITE'] -= prise; q_prd += prise; qte_a_trouver -= prise; dates_trouvees.append(prod['Date_Dispo_Reelle'])
                                        if qte_a_trouver == 0: break
                            return q_stk, q_prd, qte_a_trouver
                        qs1, qp1, qte_restante = consommer(article, qte_restante); qte_prise_stock += qs1; qte_prise_prod += qp1; utilise_prepa = "Non"
                        if qte_restante > 0:
                            cascade = get_cascade_prepas(article)
                            for prepa in cascade:
                                if qte_restante <= 0: break
                                qs2, qp2, qte_restante = consommer(prepa, qte_restante); qte_prise_stock += qs2; qte_prise_prod += qp2
                                if (qs2 + qp2) > 0: utilise_prepa = f"Oui ({prepa})"
                        if qte_restante > 0:
                            if dates_trouvees: date_dispo = max(dates_trouvees).strftime('%d/%m/%Y') + " (Partiel)"; statut = "Attente Prod (Partiel)"
                            else: date_dispo = "Pas de date"; statut = "Rupture"
                        else:
                            if not dates_trouvees: date_dispo = "Immédiate"; statut = "En Stock"
                            else: date_dispo = max(dates_trouvees).strftime('%d/%m/%Y'); statut = "Attente Prod"
                        resultats.append({'Num_Commande': commande['NUM_CDE'], 'Client': commande['CLIENT'], 'Article': article, 'Qte_Demandée': int(commande['QUANTITE']), 'Tiré_Stock': int(qte_prise_stock), 'Tiré_Prod': int(qte_prise_prod), 'Remplacement_Prepa': utilise_prepa, 'Manquant': int(qte_restante), 'Statut': statut, 'Date_Disponibilité': date_dispo, 'Adresse': commande['ADRESSE'], 'Ville': commande['VILLE'], 'Pays': commande['PAYS'], 'Exportateur': commande['EXPORTATEUR']})
                    
                    df_final = pd.DataFrame(resultats)
                    
                    # SAUVEGARDE DANS LE FICHIER CACHÉ
                    cache_data = {'df_final': df_final, 'dict_details': dict_details, 'settings_pdf': settings_pdf}
                    pd.to_pickle(cache_data, 'base_logistique.pkl')
                    
                    st.success("✅ Calcul terminé ! La base a été sauvegardée et est maintenant visible par les clients sur le Front Office.")
                    colonnes_a_afficher = [c for c in df_final.columns if c not in ['Adresse', 'Ville', 'Exportateur']]
                    st.dataframe(df_final[colonnes_a_afficher], use_container_width=True)
                except Exception as e:
                    st.error(f"Erreur : {e}")
        else: st.warning("Veuillez déposer tous les fichiers.")

# ==========================================
# ESPACE CLIENT (FRONT OFFICE) - V48
# ==========================================
elif st.session_state['role'] == 'client':
    st.title("🌐 Portail Suivi Commandes & Logistique")
    st.write("Bienvenue sur votre espace. Vous trouverez ci-dessous la disponibilité de vos commandes en temps réel.")
    
    if not os.path.exists('base_logistique.pkl'):
        st.warning("⚠️ Aucune donnée n'est actuellement disponible. L'entrepôt n'a pas encore mis à jour la base.")
        st.stop()
    
    try:
        cache = pd.read_pickle('base_logistique.pkl')
        df_final = cache['df_final']
        dict_details = cache['dict_details']
        settings_pdf = cache['settings_pdf']
    except Exception as e:
        st.error("Erreur lors de la lecture de la base de données.")
        st.stop()

    # --- 1. CALCUL DES KPIs (TABLEAU DE BORD) ---
    order_status_map = {}
    for cmd, grp in df_final.groupby('Num_Commande'):
        if str(cmd).upper() in ["INCONNU", "NAN"]: continue
        statuses = grp['Statut'].values
        if 'Rupture' in statuses: 
            order_status_map[cmd] = 'Rupture'
        elif any('Attente Prod' in s for s in statuses): 
            order_status_map[cmd] = 'Attente'
        else: 
            order_status_map[cmd] = 'Prêt'
            
    nb_total = len(order_status_map)
    nb_pret = sum(1 for v in order_status_map.values() if v == 'Prêt')
    nb_attente = sum(1 for v in order_status_map.values() if v == 'Attente')
    nb_rupture = sum(1 for v in order_status_map.values() if v == 'Rupture')

    st.subheader("📊 Tableau de Bord du Portefeuille")
    col_k1, col_k2, col_k3, col_k4 = st.columns(4)
    col_k1.metric("📦 Total Commandes", nb_total)
    col_k2.metric("🟢 100% Prêtes (En Stock)", nb_pret)
    col_k3.metric("⏳ En Attente Production", nb_attente)
    col_k4.metric("🔴 Bloquées (Rupture Partielle)", nb_rupture)
    st.divider()

    # --- 2. BOUTONS DE TÉLÉCHARGEMENT ---
    mask_us_ca = df_final['Pays'].astype(str).str.upper().str.contains('ETATS|CANADA', regex=True, na=False)
    df_us_ca = df_final[mask_us_ca].copy()
    df_monde = df_final[~mask_us_ca].copy()

    st.subheader("📥 Téléchargements Groupés")
    col_d1, col_d2, col_d3, col_d4 = st.columns(4)
    with col_d1:
        if not df_us_ca.empty:
            buf1 = io.BytesIO()
            with pd.ExcelWriter(buf1, engine='openpyxl') as writer: df_us_ca.to_excel(writer, index=False)
            st.download_button("🇺🇸🇨🇦 Excel (US & CANADA)", data=buf1.getvalue(), file_name="Commandes_US_CANADA.xlsx", type="primary", use_container_width=True)
    with col_d2:
        if not df_monde.empty:
            buf2 = io.BytesIO()
            with pd.ExcelWriter(buf2, engine='openpyxl') as writer: df_monde.to_excel(writer, index=False)
            st.download_button("🌍 Excel (EUROPE & INTER)", data=buf2.getvalue(), file_name="Commandes_MONDE.xlsx", type="primary", use_container_width=True)
    with col_d3:
        if REPORTLAB_OK and not df_final.empty:
            zip_pack = generer_packing_lists_zip(df_final, dict_details)
            st.download_button("📦 Packing Lists (PDF)", data=zip_pack, file_name="Packing_Lists.zip", type="secondary", use_container_width=True)
    with col_d4:
        if FPDF_OK and not df_final.empty:
            zip_rdv = generer_rdv_documents_zip(df_final, dict_details, settings_pdf)
            st.download_button("📅 RDV Documents (PDF)", data=zip_rdv, file_name="RDV_Documents.zip", type="secondary", use_container_width=True)

    st.divider()

    # --- 3. RECHERCHE ET FILTRES DYNAMIQUES ---
    st.subheader("🔍 Recherche et Vision Détaillée")
    col_f1, col_f2 = st.columns([1, 2])
    search_query = col_f1.text_input("Rechercher (N° Commande, Pays, Article...)")
    
    # Filtres par cases à cocher (multiselect)
    options_statut = ["🟢 100% Prêtes", "⏳ En Attente Production", "🔴 Bloquées (Rupture)"]
    filtre_statut = col_f2.multiselect("Filtrer par état de la commande :", options_statut, default=options_statut)

    # --- 4. AFFICHAGE ACCORDÉON (FILTRÉ) ---
    groupes_commandes = df_final.groupby('Num_Commande')
    commandes_affichees = 0

    for cmd, lignes in groupes_commandes:
        if str(cmd).upper() in ["INCONNU", "NAN"]: continue
        
        # Déterminer le statut de l'accordéon pour le filtre
        statut_cmd = order_status_map[cmd]
        if statut_cmd == 'Prêt' and "🟢 100% Prêtes" not in filtre_statut: continue
        if statut_cmd == 'Attente' and "⏳ En Attente Production" not in filtre_statut: continue
        if statut_cmd == 'Rupture' and "🔴 Bloquées (Rupture)" not in filtre_statut: continue

        # Filtrer par la barre de recherche (recherche globale sur la commande)
        texte_recherche = f"{cmd} {lignes['Client'].iloc[0]} {lignes['Pays'].iloc[0]} {' '.join(lignes['Article'].astype(str))}".lower()
        if search_query and search_query.lower() not in texte_recherche: continue

        commandes_affichees += 1
        client_nom = str(lignes.iloc[0]['Client'])
        pays_nom = str(lignes.iloc[0]['Pays'])
        
        if statut_cmd == 'Rupture': icone = "🔴"
        elif statut_cmd == 'Attente': icone = "⏳"
        else: icone = "🟢"
            
        titre_accordian = f"{icone} Commande N° {cmd} — {client_nom} ({pays_nom})"
        
        with st.expander(titre_accordian):
            for _, row in lignes.iterrows():
                art = str(row['Article'])
                qte = int(row['Qte_Demandée'])
                statut = str(row['Statut'])
                date = str(row['Date_Disponibilité'])
                libelle = dict_details.get(art, {}).get('libelle', 'Article inconnu')
                
                if statut == "Rupture":
                    st.error(f"**{art} - {libelle}** : {qte} cartons ➔ **{statut}** (Date inconnue)")
                elif "Attente Prod" in statut:
                    st.warning(f"**{art} - {libelle}** : {qte} cartons ➔ **{statut}** (Prêt le : {date})")
                else:
                    st.success(f"**{art} - {libelle}** : {qte} cartons ➔ **{statut}** (Immédiate)")

    if commandes_affichees == 0:
        st.info("Aucune commande ne correspond à votre recherche ou vos filtres.")
