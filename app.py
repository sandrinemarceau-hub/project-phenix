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
# CONFIGURATION ET SÉCURITÉ (SECRETS)
# ==========================================
st.set_page_config(layout="wide", page_title="Sovereign Brands - Orders")

if 'role' not in st.session_state:
    st.session_state['role'] = None

PASS_ADMIN = st.secrets.get("PASS_ADMIN")
PASS_CLIENT = st.secrets.get("PASS_CLIENT")

# --- ECRAN DE CONNEXION ---
if st.session_state['role'] is None:
    st.title("🔐 Logistics Access Portal")
    st.write("Please enter your password to access your space.")
    pwd = st.text_input("Password", type="password")
    
    if st.button("Login", type="primary"):
        if pwd == PASS_ADMIN:
            st.session_state['role'] = 'admin'
            st.rerun()
        elif pwd == PASS_CLIENT:
            st.session_state['role'] = 'client'
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

# --- BOUTON DÉCONNEXION ---
with st.sidebar:
    if st.session_state['role'] == 'admin':
        if st.button("🚪 Déconnexion (Admin)"):
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

def safe_xml(texte): return str(texte).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
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
                if score > max_score: max_score = score; best_df = df_temp
            except: pass
        if best_df is not None: return best_df
        return pd.read_excel(xls, sheet_name=0, skiprows=lignes_a_ignorer)

# ==========================================
# MOTEUR PDF UNITAIRE & STRUCTURE RDV
# ==========================================
if FPDF_OK:
    class RDVPDF(FPDF):
        def header(self):
            self.ln(35); self.set_font("Helvetica", "B", 24); self.cell(0, 15, 'COLLECTION APPOINTMENT', 0, 1, 'C'); self.ln(2)
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

def generer_pl_unique(cmd, lignes, dict_details, app_settings):
    # NOUVEAUTÉ : Distinction client facturé (pour entête) et livré (pour consignee)
    client_fact = str(lignes.iloc[0].get('Client_Facturation', lignes.iloc[0]['Client']))
    client_liv = str(lignes.iloc[0]['Client'])
    
    adresse = clean_nan(lignes.iloc[0]['Adresse'])
    ville = clean_nan(lignes.iloc[0]['Ville'])
    pays = clean_nan(lignes.iloc[0]['Pays'])
    ref_client = clean_nan(lignes.iloc[0].get('Ref_Client', ''))

    client_fact_upper = client_fact.upper()
    
    # NOUVEAUTÉ : ADRESSES DYNAMIQUES SELON L'ENTITÉ JURIDIQUE
    if "SOVEREIGN BRANDS" in client_fact_upper or "LUC BELAIRE LLC" in client_fact_upper:
        exp_text_html = "SOVEREIGN BRANDS, LLC / Luc Belaire LLC<br/>1300 Old Skokie Valley Rd, Suite A<br/>Highland Park, IL 60035, USA"
    elif "LUC BELAIRE INTERNATIONAL" in client_fact_upper:
        exp_text_html = "LUC BELAIRE INTERNATIONAL, LTD<br/>5th Floor, 76 Sir John Rogerson's Quay<br/>Dublin Docklands, DUBLIN 2<br/>D02 C9D0, IRELAND"
    else:
        exp_text_html = safe_xml(app_settings.get('exp_row', '')).replace('\n', '<br/>')
        
    txt_exp = f"<b>EXPORTER:</b><br/>{exp_text_html}"
    
    consignee_lines = [f"<b>CONSIGNEE:</b><br/>{safe_xml(client_liv)}"]
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
    
    header_str = f"<b>Order Ref:</b> {cmd}"
    if ref_client: header_str += f" | <b>PO / Client Ref:</b> {safe_xml(ref_client)}"
    elements.append(Paragraph(header_str, style_header))
    
    elements.append(Spacer(1, 10))
    elements.append(HRFlowable(width=18.5*cm, thickness=1.5, color=colors.black, spaceAfter=15, hAlign='CENTER'))
    
    t_adr = Table([[Paragraph(txt_exp, style_header), "", Paragraph(txt_con, style_header)]], colWidths=[8.5*cm, 1.5*cm, 8.5*cm])
    t_adr.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('BOTTOMPADDING',(0,0),(-1,-1), 15)]))
    elements.append(t_adr)

    data = [['CASES', 'SKU', 'DESCRIPTION', 'UNITS', 'NET WGT']]
    t_q, t_c, t_p, t_pal = 0, 0, 0.0, 0.0
    type_pal_label = "N/A"
    
    for _, row in lignes.iterrows():
        art = str(row['Article'])
        qte = int(row['Qte_Demandée']) 
        d = dict_details.get(art, {'libelle': 'Inconnu', 'format': '', 'degres': '', 'couleur': '', 'uc': 6.0, 'poids': 0.0, 'type_pal': 'N/A', 'cas_pal': 100.0})
        uc = d['uc'] if d['uc'] > 0 else 6.0
        cas_pal = d['cas_pal'] if d['cas_pal'] > 0 else 100.0
        if d['type_pal'] not in ["N/A", "", "NAN"]: type_pal_label = d['type_pal']
        
        units = int(qte * uc)
        poids_ligne = units * (d['poids'] if d['poids'] > 0 else 1.5) 
        palettes_ligne = qte / cas_pal if cas_pal > 0 else 0
        
        t_q += units; t_c += qte; t_p += poids_ligne; t_pal += palettes_ligne
        
        desc_html = f"<b>{safe_xml(d['libelle'])}</b>"
        sub1 = []
        if d['format']: sub1.append(f"Size: {safe_xml(d['format'])}")
        if sub1: desc_html += f"<br/><font color='#555555'>{' | '.join(sub1)}</font>"
        
        data.append([str(qte), safe_xml(art), Paragraph(desc_html, style_desc), str(units), format_num(poids_ligne)])

    t_art = Table(data, colWidths=[2*cm, 2.5*cm, 9*cm, 2*cm, 3*cm], repeatRows=1)
    t_art.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.black), 
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke), 
        ('ALIGN', (0,0), (0,-1), 'CENTER'), 
        ('ALIGN', (3,0), (-1,-1), 'RIGHT'), 
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), 
        ('GRID', (0,0), (-1,-1), 0.2, colors.grey), 
        ('TOPPADDING', (0,0), (-1,-1), 8), 
        ('BOTTOMPADDING', (0,0), (-1,-1), 8)
    ]))
    elements.append(t_art)
    elements.append(Spacer(1, 20))
    
    tot_data = [
        [f"TOTAL UNITS: {int(t_q)}", f"TOTAL NET WEIGHT: {format_num(t_p)} kg"], 
        [f"TOTAL CASES: {int(t_c)}", f"TOTAL PALLETS: {int(math.ceil(t_pal))} ({type_pal_label})"]
    ]
    t_tot = Table(tot_data, colWidths=[9*cm, 9.5*cm])
    t_tot.setStyle(TableStyle([('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'), ('FONTSIZE',(0,0),(-1,-1),10), ('BOX',(0,0),(-1,-1),1.5,colors.black), ('LEFTPADDING',(0,0),(-1,-1), 10), ('TOPPADDING',(0,0),(-1,-1), 10), ('BOTTOMPADDING',(0,0),(-1,-1), 10)]))
    elements.append(t_tot); elements.append(Spacer(1, 40))
    elements.append(Paragraph("________________________________<br/>Authorized Signature & Stamp", styles['Normal']))
    
    doc.build(elements)
    return pdf_buffer.getvalue()

def generer_rdv_unique(cmd, lignes, dict_details, app_settings):
    pire_date_obj = None; en_rupture = False
    for _, r in lignes.iterrows():
        statut = r['Statut']
        d_str = str(r['Date_Disponibilité']).replace(" (Partiel)", "")
        if statut == "Rupture": en_rupture = True
        elif "Attente Prod" in statut and d_str != "Pas de date":
            try:
                d_obj = datetime.strptime(d_str, "%d/%m/%Y")
                if pire_date_obj is None or d_obj > pire_date_obj: pire_date_obj = d_obj
            except: pass
            
    if en_rupture: date_finale = "TBD (Partial Out of Stock)"
    elif pire_date_obj: date_finale = pire_date_obj.strftime("%d/%m/%Y")
    else: date_finale = "ASAP (In Stock)"

    t_poids = 0.0; t_palettes = 0.0
    for _, r in lignes.iterrows():
        art = str(r['Article']); qte = int(r['Qte_Demandée'])
        d = dict_details.get(art, {'uc': 6.0, 'poids': 0.0, 'cas_pal': 100.0})
        uc = d['uc'] if d['uc'] > 0 else 6.0; cas_pal = d['cas_pal'] if d['cas_pal'] > 0 else 100.0
        units = int(qte * uc)
        t_poids += units * (d['poids'] if d['poids'] > 0 else 1.5)
        t_palettes += qte / cas_pal if cas_pal > 0 else 0

    client_fact = str(lignes.iloc[0].get('Client_Facturation', lignes.iloc[0]['Client']))
    pays = clean_nan(lignes.iloc[0]['Pays'])
    ref_client = clean_nan(lignes.iloc[0].get('Ref_Client', ''))
    adresse_enlevement = app_settings['adresse_veuve']

    if not FPDF_OK: return b""
    
    pdf = RDVPDF(); pdf.add_page(); pdf.set_font("Helvetica", "B", 14)
    txt_noir = "Available for collection on: "; txt_rouge = date_finale
    largeur_totale = pdf.get_string_width(txt_noir) + pdf.get_string_width(txt_rouge)
    pdf.set_x((pdf.w - largeur_totale) / 2); pdf.set_text_color(0, 0, 0); pdf.cell(pdf.get_string_width(txt_noir), 10, txt_noir)
    pdf.set_text_color(200, 0, 0); pdf.cell(pdf.get_string_width(txt_rouge), 10, txt_rouge, 0, 1); pdf.set_text_color(0, 0, 0); pdf.ln(10)
    
    pdf.draw_harmonized_row("Pick Up Address", adresse_enlevement)
    pdf.draw_harmonized_row("Loading Hours", app_settings['horaires'])
    pdf.draw_harmonized_row("Contact", app_settings['contact'])
    pdf.draw_harmonized_row("Order Number", str(cmd))
    if ref_client:
        pdf.draw_harmonized_row("Customer PO / Ref", ref_client)
    pdf.draw_harmonized_row("Country of Delivery", pays)
    pdf.draw_harmonized_row("Customer", client_fact)
    pdf.draw_harmonized_row("Number of Pallets", f"{int(math.ceil(t_palettes))} Pallet(s)")
    pdf.draw_harmonized_row("Total Weight", f"{format_num(t_poids)} KG")
    pdf.draw_harmonized_row("Shipping Costs", "-")
    
    pdf.ln(15); pdf.set_font("Helvetica", "B", 8.5); pdf.set_text_color(200, 0, 0)
    w_en = "Reminder: We require a 48-hour notice to prepare the order before collection. ANY CARRIER ARRIVING WITHOUT AN APPOINTMENT OR FAILING TO RESPECT THIS NOTICE PERIOD WILL BE REFUSED AND NOT LOADED."
    pdf.multi_cell(0, 5, w_en, align='C')
    return pdf.output(dest="S").encode("latin-1")

def generer_packing_lists_zip(df_resultats, dict_details, app_settings):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for cmd in df_resultats['Num_Commande'].unique():
            if str(cmd).upper() in ["INCONNU", "NAN"]: continue
            lignes = df_resultats[df_resultats['Num_Commande'] == cmd]
            pdf_bytes = generer_pl_unique(cmd, lignes, dict_details, app_settings)
            safe_name = str(cmd).replace('/', '_').replace('\\', '_')
            zip_file.writestr(f"Packing_List_{safe_name}.pdf", pdf_bytes)
    return zip_buffer.getvalue()

def generer_rdv_documents_zip(df_resultats, dict_details, app_settings):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for cmd in df_resultats['Num_Commande'].unique():
            if str(cmd).upper() in ["INCONNU", "NAN"]: continue
            lignes = df_resultats[df_resultats['Num_Commande'] == cmd]
            pdf_bytes = generer_rdv_unique(cmd, lignes, dict_details, app_settings)
            safe_name = str(cmd).replace('/', '_').replace('\\', '_')
            zip_file.writestr(f"RDV_{safe_name}.pdf", pdf_bytes)
    return zip_buffer.getvalue()

# ==========================================
# ESPACE ADMINISTRATEUR (BACK OFFICE)
# ==========================================
if st.session_state['role'] == 'admin':
    with st.sidebar:
        st.write("⭐ **Système de Priorité**")
        st.caption("Les clients listés ici (séparés par des virgules) seront servis en premier sur les stocks disponibles.")
        clients_vip = st.text_area("Clients VIP", "SOVEREIGN BRANDS USA, LUC BELAIRE LLC")
        liste_vip = [c.strip().upper() for c in clients_vip.split(',')]
        
        st.divider()
        st.write("📝 **Paramètres d'Enlèvement**")
        pdf_contact = st.text_input("Contact Email", "logistique@sovereignbrands.com")
        pdf_horaires = st.text_input("Horaires", "08:00 - 16:00 (Monday - Friday)") 
        pdf_adresse_veuve = st.text_area("Adresse (Veuve Ambal)", "VEUVE AMBAL\n32 rue de la Croix Clément\n71530 Champforgeuil", height=80)
        
        st.divider()
        st.write("🌍 **Adresses par défaut (Divers)**")
        exp_row = st.text_area("Exportateur (Fallback Divers)", "SOVEREIGN BRANDS FRANCE\n10 Rue de la Logistique\n75000 Paris", height=80)
        
    settings_pdf = {
        'contact': pdf_contact, 'horaires': pdf_horaires, 'adresse_veuve': pdf_adresse_veuve, 'exp_row': exp_row
    }

    st.title("🛠️ Back Office - Mise à jour de la Base")
    st.write("Importez vos fichiers usine ici. Les résultats seront sauvegardés pour les clients.")

    col1, col2, col3, col4 = st.columns(4)
    with col1: fichier_stock = st.file_uploader("Fichier Stock", type=['xlsx', 'xls', 'csv']); skip_stock = st.number_input("Ignorer (Stock)", min_value=0, value=1)
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
                                    if c_lib: 
                                        val_lib = clean_nan(r[c_lib])
                                        if val_lib and val_lib != "NAN": dict_details[art_id]['libelle'] = val_lib
                                    if c_poids: 
                                        val_pds = float(nettoyage_quantite(pd.Series([r[c_poids]]))[0])
                                        if val_pds > 0: dict_details[art_id]['poids'] = val_pds
                                    if c_cas_pal: 
                                        val_pal = float(nettoyage_quantite(pd.Series([r[c_cas_pal]]))[0])
                                        if val_pal > 0: dict_details[art_id]['cas_pal'] = val_pal

                    df_stock_brut = lire_fichier(fichier_stock, skip_stock); mask_total = df_stock_brut.astype(str).apply(lambda x: x.str.contains('TOTAL', case=False, na=False)).any(axis=1); df_stock_brut = df_stock_brut[~mask_total]
                    df_stock_brut.columns = df_stock_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True)
                    col_art_stock = next((c for c in ['CODEARTICLE', 'ARTICLECODE', 'ARTICLE', 'REFERENCE', 'CODE'] if c in df_stock_brut.columns), None)
                    col_qte_stock = next((c for c in ['STOCKPHYSIQUE', 'PHYSIQUE', 'STOCKDISPONIBLE', 'DISPONIBLE', 'QTEDISPO', 'QTESTOCK', 'QUANTITE', 'STOCK'] if c in df_stock_brut.columns), None)
                    if not col_art_stock or not col_qte_stock: st.error("❌ Erreur STOCK : Colonnes introuvables."); st.stop()
                    
                    df_stock = pd.DataFrame()
                    df_stock['CODE_ARTICLE'] = nettoyage_extreme(df_stock_brut[col_art_stock])
                    df_stock['STOCK_DISPO'] = nettoyage_quantite(df_stock_brut[col_qte_stock]) if col_qte_stock else 0
                    
                    # --- CONVERSION BOUTEILLES -> CARTONS (STOCK) ---
                    df_stock['STOCK_DISPO'] = df_stock.apply(lambda r: r['STOCK_DISPO'] / (dict_details.get(r['CODE_ARTICLE'], {}).get('uc', 6.0) or 6.0), axis=1)
                    stock_actuel = df_stock.groupby('CODE_ARTICLE')['STOCK_DISPO'].sum().to_dict()

                    liste_prod = []; df_prod_brut_total = pd.DataFrame() 
                    for f in fichiers_prod:
                        df_temp = lire_fichier(f, skip_prod); df_prod_brut_total = pd.concat([df_prod_brut_total, df_temp], ignore_index=True); colonnes_temp = df_temp.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True); df_temp.columns = colonnes_temp
                        col_sortie_auto = next((c for c in colonnes_temp if 'SORTIE' in c), None); col_entree_auto = next((c for c in colonnes_temp if 'ENTREE' in c or 'PREPA' in c), None)
                        if col_sortie_auto and col_entree_auto:
                            for _, r in df_temp.iterrows():
                                parent = nettoyage_extreme(pd.Series([r[col_sortie_auto]]))[0]; enfant = nettoyage_extreme(pd.Series([r[col_entree_auto]]))[0]
                                if parent and enfant and parent != enfant and parent not in ["0", "NAN", "NONE"] and enfant not in ["0", "NAN", "NONE"]: dict_prepa[parent] = enfant
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
                                    if code and code not in ["0", "NAN", "NONE"]: 
                                        # --- CONVERSION BOUTEILLES -> CARTONS (PROD) ---
                                        uc_prod = dict_details.get(code, {}).get('uc', 6.0) or 6.0
                                        qte_cartons = qte / uc_prod
                                        liste_prod.append({'ARTICLE': code, 'QTE_PRODUITE': qte_cartons, 'DATE_PROD': d})
                    if liste_prod:
                        df_production = pd.DataFrame(liste_prod); df_production['Date_Dispo_Reelle'] = df_production['DATE_PROD'] + timedelta(days=2); df_production = df_production.sort_values(by=['ARTICLE', 'Date_Dispo_Reelle'])
                        productions_futures = df_production.to_dict('records')
                    else: productions_futures = []

                    df_commandes_brut = lire_fichier(fichier_commandes, skip_cmd); colonnes_cmd = df_commandes_brut.columns.astype(str).str.upper().str.replace(r'[^A-Z]', '', regex=True); df_commandes_brut.columns = colonnes_cmd
                    
                    # NOUVEAUTÉ : DÉTECTION DES NOUVELLES COLONNES (FACTURATION, LIVRAISON, CARTONS)
                    col_art_cmd = next((c for c in ['ARTICLECODE', 'CODEARTICLE', 'ARTICLE'] if c in colonnes_cmd), None)
                    col_date_cmd = next((c for c in ['DATECDE', 'DATECOMMANDE', 'DATECREATION', 'DATE'] if c in colonnes_cmd), None)
                    
                    col_qte_cmd_cartons = next((c for c in ['NBCARTONS', 'CARTONS'] if c in colonnes_cmd), None)
                    col_qte_cmd_bouteilles = next((c for c in ['QTEUBCDETOTAL', 'QTEUBCDE', 'QUANTITE', 'QTE', 'TOTAL', 'TOTALGNRAL', 'TOTALGENERAL'] if c in colonnes_cmd), None)
                    
                    col_num_cmd = next((c for c in ['NUMCDE', 'NUMCOMMANDE', 'COMMANDE'] if c in colonnes_cmd), None)
                    
                    col_client_fact = next((c for c in ['CLIENTNOM', 'NOMCLIENTFACTURATION', 'CLIENT'] if c in colonnes_cmd), None)
                    col_client_liv = next((c for c in ['EXPENOMCLIENT', 'CONSIGNEE', 'DESTINATAIRE'] if c in colonnes_cmd), None)
                    if not col_client_liv: col_client_liv = col_client_fact
                    
                    col_adresse = next((c for c in colonnes_cmd if 'ADRESSE' in c or 'ADR' in c), None)
                    col_ville = next((c for c in colonnes_cmd if 'EXPEVILLE', 'VILLE', 'CITY'] if c in colonnes_cmd), None)
                    col_pays = next((c for c in colonnes_cmd if 'EXPEPAYS', 'PAYS', 'COUNTRY'] if c in colonnes_cmd), None)
                    
                    col_ref_client = next((c for c in ['REFCMDCLT', 'REFCLIENT', 'REFERENCECLIENT', 'PO', 'CDECLIENT', 'CUSTOMERREF', 'VOTREREF'] if c in colonnes_cmd), None)
                    
                    df_commandes = pd.DataFrame()
                    df_commandes['ARTICLE_CODE'] = nettoyage_extreme(df_commandes_brut[col_art_cmd])
                    df_commandes['DATE_CDE'] = pd.to_datetime(df_commandes_brut[col_date_cmd], dayfirst=True, errors='coerce')
                    
                    # NOUVEAUTÉ : GESTION INTELLIGENTE DES QUANTITÉS
                    if col_qte_cmd_cartons:
                        # Si on trouve NB CARTONS, on prend la valeur telle quelle !
                        df_commandes['QUANTITE'] = nettoyage_quantite(df_commandes_brut[col_qte_cmd_cartons])
                    else:
                        # Sinon on prend les bouteilles et on les divise
                        df_commandes['QUANTITE'] = nettoyage_quantite(df_commandes_brut[col_qte_cmd_bouteilles])
                        df_commandes['QUANTITE'] = df_commandes.apply(lambda r: r['QUANTITE'] / (dict_details.get(r['ARTICLE_CODE'], {}).get('uc', 6.0) or 6.0), axis=1)

                    df_commandes['NUM_CDE'] = df_commandes_brut[col_num_cmd] if col_num_cmd else 'Inconnu'
                    df_commandes['CLIENT_FACT'] = df_commandes_brut[col_client_fact] if col_client_fact else 'Inconnu'
                    df_commandes['CLIENT_LIV'] = df_commandes_brut[col_client_liv] if col_client_liv else df_commandes['CLIENT_FACT']
                    df_commandes['ADRESSE'] = df_commandes_brut[col_adresse] if col_adresse else ""
                    df_commandes['VILLE'] = df_commandes_brut[col_ville] if col_ville else ""
                    df_commandes['PAYS'] = df_commandes_brut[col_pays] if col_pays else ""
                    df_commandes['REF_CLIENT'] = df_commandes_brut[col_ref_client].astype(str).replace('nan', '', regex=True).replace('None', '', regex=True) if col_ref_client else ""
                    
                    # NOUVEAUTÉ : ATTRIBUTION DU NIVEAU DE PRIORITÉ
                    df_commandes['NIVEAU_PRIO'] = df_commandes['CLIENT_FACT'].apply(lambda x: 0 if any(vip in str(x).upper() for vip in liste_vip) else 1)
                    
                    # Tri par Priorité (0 d'abord) PUIS par Date de commande
                    df_commandes = df_commandes.dropna(subset=['DATE_CDE']).sort_values(by=['NIVEAU_PRIO', 'DATE_CDE'])

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
                        
                        resultats.append({
                            'Num_Commande': commande['NUM_CDE'], 
                            'Ref_Client': clean_nan(commande['REF_CLIENT']),
                            'Date_Commande': commande['DATE_CDE'], 
                            'Client_Facturation': commande['CLIENT_FACT'],
                            'Client': commande['CLIENT_LIV'], # Gardé pour compatibilité Front-end
                            'Article': article, 
                            'Qte_Demandée': int(commande['QUANTITE']), 
                            'Tiré_Stock': int(qte_prise_stock), 
                            'Tiré_Prod': int(qte_prise_prod), 
                            'Remplacement_Prepa': utilise_prepa, 
                            'Manquant': int(qte_restante), 
                            'Statut': statut, 
                            'Date_Disponibilité': date_dispo, 
                            'Adresse': commande['ADRESSE'], 
                            'Ville': commande['VILLE'], 
                            'Pays': commande['PAYS']
                        })
                    
                    df_final = pd.DataFrame(resultats)
                    
                    cache_data = {'df_final': df_final, 'dict_details': dict_details, 'settings_pdf': settings_pdf}
                    pd.to_pickle(cache_data, 'base_logistique.pkl')
                    
                    st.success("✅ Calcul terminé ! La base a été sauvegardée et est maintenant visible par les clients sur le Front Office.")
                    
                    df_admin_view = df_final.copy()
                    df_admin_view.rename(columns={'Manquant': 'Cartons_Manquants', 'Qte_Demandée': 'Cartons_Commandés', 'Ref_Client': 'Référence_Client', 'Client': 'Client_Livraison'}, inplace=True)
                    colonnes_a_afficher = [c for c in df_admin_view.columns if c not in ['Adresse', 'Ville']]
                    st.dataframe(df_admin_view[colonnes_a_afficher], use_container_width=True)
                    
                    buf_admin = io.BytesIO()
                    with pd.ExcelWriter(buf_admin, engine='openpyxl') as writer: df_admin_view.to_excel(writer, index=False)
                    st.download_button("📥 Télécharger le Résultat Global (Excel)", data=buf_admin.getvalue(), file_name="Resultat_Global_Logistique.xlsx", type="primary", use_container_width=True)
                    
                except Exception as e:
                    import traceback
                    st.error(f"Erreur : {e}\n{traceback.format_exc()}")
        else: st.warning("Veuillez déposer tous les fichiers.")

# ==========================================
# ESPACE CLIENT (FRONT OFFICE)
# ==========================================
elif st.session_state['role'] == 'client':
    
    st.markdown("""
        <style>
            #MainMenu, footer, header {visibility: hidden;}
            .block-container { padding-top: 1rem; max-width: 95%; }
            .stApp { background-color: #f4f5f7; }
            .main-panel { background-color: white; padding: 30px; border-radius: 8px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.05); margin-top: 20px; }
            .badge-ready { background-color: #e6f4ea; color: #1e8e3e; padding: 4px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: 600; display: inline-block; white-space: nowrap;}
            .badge-pending { background-color: #fef7e0; color: #b06000; padding: 4px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: 600; display: inline-block; white-space: nowrap;}
            .badge-blocked { background-color: #fce8e6; color: #d93025; padding: 4px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: 600; display: inline-block; white-space: nowrap;}
            
            .custom-table { width: 100%; border-collapse: collapse; font-size: 0.9rem; color: #333333; margin-top: 10px; }
            .custom-table th { color: #6b778c; font-weight: 500; text-align: left; padding-bottom: 10px; border-bottom: 1px solid #dfe1e6; }
            .custom-table td { padding: 12px 0; border-bottom: 1px solid #f4f5f7; vertical-align: middle; color: #333333; }
            .item-name { font-weight: 600; color: #172b4d; display: block; }
            .item-sku { color: #7a869a; font-size: 0.8rem; font-family: monospace;}
            
            div[data-testid="stExpander"] { border: 1px solid #dfe1e6 !important; border-radius: 4px !important; margin-bottom: 10px !important; background-color: white !important; box-shadow: none !important; }
            div[data-testid="stExpander"] summary { padding: 15px !important; }
            @media (max-width: 768px) { .main-panel { padding: 15px; margin-top: 10px; } .custom-table { display: block; overflow-x: auto; white-space: nowrap; } div[data-testid="stMetricValue"] {font-size: 1.5rem !important;} }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-panel">', unsafe_allow_html=True)
    st.title("Orders")
    
    if not os.path.exists('base_logistique.pkl'):
        st.warning("⚠️ No data available. The warehouse has not updated the database yet.")
        st.stop()
    
    try:
        cache = pd.read_pickle('base_logistique.pkl')
        df_final = cache['df_final']
        dict_details = cache['dict_details']
        settings_pdf = cache['settings_pdf']
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

    col_search, col_space, col_btn_excel, col_btn_import, col_btn_new = st.columns([3, 0.5, 1.2, 1.2, 1.2])
    search_query = col_search.text_input("🔍 Search Order ID or Customer...", label_visibility="collapsed", placeholder="Search Order ID or Customer...")
    
    def preparer_excel_client(df_source):
        colonnes_dispos = df_source.columns.tolist()
        cols_a_garder = ['Num_Commande']
        if 'Date_Commande' in colonnes_dispos: cols_a_garder.append('Date_Commande')
        if 'Ref_Client' in colonnes_dispos: cols_a_garder.append('Ref_Client')
        cols_a_garder.extend(['Client', 'Pays', 'Article', 'Qte_Demandée', 'Manquant', 'Statut', 'Date_Disponibilité'])
        
        df_client = df_source[cols_a_garder].copy()
        
        rename_map = {
            'Num_Commande': 'Order_No',
            'Date_Commande': 'Order_Date',
            'Ref_Client': 'PO_Number',
            'Client': 'Consignee',
            'Pays': 'Country',
            'Article': 'SKU',
            'Qte_Demandée': 'Ordered_Cases',
            'Manquant': 'Missing_Cases',
            'Statut': 'Status',
            'Date_Disponibilité': 'Availability_Date'
        }
        df_client.rename(columns=rename_map, inplace=True)
            
        def translate_status(s):
            if "Rupture" in s: return "Out of stock"
            if "Attente Prod" in s: return "In Production"
            return "Ready"
            
        if 'Status' in df_client.columns:
            df_client['Status'] = df_client['Status'].apply(translate_status)
        if 'Availability_Date' in df_client.columns:
            df_client['Availability_Date'] = df_client['Availability_Date'].astype(str).str.replace(" (Partiel)", "").replace({'Immédiate': 'Immediate', 'Pas de date': 'TBD'})
        return df_client

    df_client_export = preparer_excel_client(df_final)
    buf_global = io.BytesIO()
    with pd.ExcelWriter(buf_global, engine='openpyxl') as writer: df_client_export.to_excel(writer, index=False)
    col_btn_excel.download_button("Export to Excel", data=buf_global.getvalue(), file_name="Sovereign_Orders.xlsx", use_container_width=True)
    
    col_btn_import.button("Import Orders", use_container_width=True, disabled=True)
    col_btn_new.button("+ New Order", type="primary", use_container_width=True, disabled=True)

    st.markdown("<br>", unsafe_allow_html=True)
    
    col_p1, col_p2, col_p3 = st.columns([2, 1.5, 1.5])
    if REPORTLAB_OK and not df_final.empty:
        zip_pack = generer_packing_lists_zip(df_final, dict_details, settings_pdf)
        col_p2.download_button("📦 Bulk Packing Lists (ZIP)", data=zip_pack, file_name="Packing_Lists.zip", type="secondary", use_container_width=True)
    if FPDF_OK and not df_final.empty:
        zip_rdv = generer_rdv_documents_zip(df_final, dict_details, settings_pdf)
        col_p3.download_button("📅 Bulk RDV Documents (ZIP)", data=zip_rdv, file_name="RDV_Documents.zip", type="secondary", use_container_width=True)

    st.markdown("<hr style='margin-top: 10px; margin-bottom: 20px; border-top: 1px solid #dfe1e6;'>", unsafe_allow_html=True)

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

        if 'Date_Commande' in lignes.columns and pd.notna(lignes['Date_Commande'].iloc[0]):
            order_date_simulated = pd.to_datetime(lignes['Date_Commande'].iloc[0]).strftime("%m/%d/%Y")
        else:
            order_date_simulated = "N/A"
            
        items_count = len(lignes)

        if statut_cmd == 'Fulfilled': badge_text = "🟢 Ready"
        elif statut_cmd == 'Pending': badge_text = "🟡 Pending"
        else: badge_text = "🔴 Unfulfilled"

        titre_accordian = f"#{cmd}   |   {order_date_simulated}   |   {client_nom}   |   {pays_nom_display}   |   Items: {items_count}   |   {badge_text}"
        
        with st.expander(titre_accordian):
            
            ref_client = clean_nan(lignes.iloc[0].get('Ref_Client', ''))
            if ref_client:
                st.markdown(f"<span style='color:#6b778c; font-size:0.9rem;'><strong>PO / Ref:</strong> {ref_client}</span><br><br>", unsafe_allow_html=True)

            html_table = "<div style='overflow-x:auto;'><table class='custom-table'><thead><tr><th>Product Details</th><th>Qty (Cases)</th><th>Availability</th><th>Status</th></tr></thead><tbody>"
            
            for _, row in lignes.iterrows():
                art = str(row['Article'])
                qte = int(row['Qte_Demandée'])
                manquant = int(row['Manquant'])
                statut_fr = str(row['Statut'])
                date = str(row['Date_Disponibilité']).replace(" (Partiel)", "")
                libelle = dict_details.get(art, {}).get('libelle', 'Unknown Item')
                
                if manquant > 0:
                    qty_html = f"<b>{qte}</b><br><span style='color:#d93025; font-size:0.75rem; font-weight:600;'>⚠ {manquant} missing</span>"
                else:
                    qty_html = f"<b>{qte}</b>"

                if statut_fr == "Rupture": 
                    pill = "<span class='badge-blocked'>Out of stock</span>"
                    date_display = "TBD"
                elif "Attente Prod" in statut_fr: 
                    pill = "<span class='badge-pending'>In Production</span>"
                    date_display = date
                else: 
                    pill = "<span class='badge-ready'>On Hand</span>"
                    date_display = "Immediate"

                html_table += f"<tr><td><span class='item-name'>{libelle}</span><span class='item-sku'>SKU: {art}</span></td><td>{qty_html}</td><td>{date_display}</td><td>{pill}</td></tr>"
            
            html_table += "</tbody></table></div>"
            st.markdown(html_table, unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            col_a1, col_a2, col_a3 = st.columns([1.5, 1.5, 2])
            
            if FPDF_OK:
                rdv_bytes = generer_rdv_unique(cmd, lignes, dict_details, settings_pdf)
                if rdv_bytes:
                    col_a1.download_button(f"📅 Download RDV Doc", data=rdv_bytes, file_name=f"RDV_{cmd}.pdf", key=f"rdv_{cmd}", use_container_width=True)
                    
            if REPORTLAB_OK:
                pl_bytes = generer_pl_unique(cmd, lignes, dict_details, settings_pdf)
                col_a2.download_button(f"📦 Download Packing List", data=pl_bytes, file_name=f"Packing_List_{cmd}.pdf", key=f"pl_{cmd}", use_container_width=True)
            
            subject = f"Update on your Sovereign Brands Order #{cmd}"
            body = f"Hello,%0A%0AHere is an update regarding your order #{cmd}.%0A"
            mail_btn_html = f'<a href="mailto:?subject={subject}&body={body}" style="display: block; text-align: center; background-color: #f4f5f7; color: #172b4d; padding: 0.5rem 1rem; border-radius: 0.25rem; text-decoration: none; border: 1px solid #dfe1e6; font-weight: 600; font-size: 0.9rem;">✉ Forward Tracking via Email</a>'
            col_a3.markdown(mail_btn_html, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)
