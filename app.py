import io
import os
from copy import copy
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
from PIL import Image

st.set_page_config(page_title="Production & Changements de Moules - OK PLAST", layout="wide")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "suivie des changement.xlsx")
SHEET_NAME = "Suivie"
LOGO_FILE = os.path.join(BASE_DIR, "logo.png")

col_logo, col_title = st.columns([1, 4])

with col_logo:
    if os.path.exists(LOGO_FILE):
        img = Image.open(LOGO_FILE)
        st.image(img, width=180)
    else:
        st.info("Logo non trouvé (nommez l'image logo.png)")

with col_title:
    st.title("📋 Production & Changements de Moules")
    st.markdown("👨‍💻 *Application développée par **Fersi Ahmed***")

st.markdown("---")


def convertir_date_ou_serial(val):
    """Convertit une date ou un numéro de série Excel (ex: 45762) en JJ/MM/AAAA."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, date):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if 20000 < val < 80000:
            dt = datetime(1899, 12, 30) + timedelta(days=int(val))
            return dt.strftime("%d/%m/%Y")
        return str(int(val))

    s = str(val).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return ""

    # Gère les nombres de série Excel stockés sous forme de texte (ex: "45762")
    try:
        num = float(s.replace(",", "."))
        if num.is_integer() and 20000 < int(num) < 80000:
            dt = datetime(1899, 12, 30) + timedelta(days=int(num))
            return dt.strftime("%d/%m/%Y")
    except ValueError:
        pass

    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            pass

    try:
        dt = pd.to_datetime(s, dayfirst=True)
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass

    return s


def formater_colonne_date(serie):
    return serie.apply(convertir_date_ou_serial)


def charger_donnees():
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
            df.columns = df.columns.astype(str).str.strip()

            if "MACH" in df.columns and "MACHINE" in df.columns:
                df = df.drop(columns=["MACH"])

            if "DATE" in df.columns:
                df["DATE"] = formater_colonne_date(df["DATE"])

            if "PIECES PRODUITES" in df.columns:
                df["PIECES PRODUITES"] = (
                    pd.to_numeric(df["PIECES PRODUITES"], errors="coerce")
                    .fillna(0)
                    .astype(int)
                )

            for col in df.columns:
                if col not in ["PIECES PRODUITES", "DATE"]:
                    df[col] = df[col].fillna("").astype(str).str.strip()
                    df[col] = df[col].replace({"nan": "", "None": ""})

            return df
        except Exception as e:
            st.error(f"⚠️ Erreur de lecture du fichier Excel : {e}")
            return pd.DataFrame()
    else:
        st.warning(f"⚠️ Le fichier '{os.path.basename(EXCEL_FILE)}' est introuvable.")
        return pd.DataFrame()


df_display = charger_donnees()

st.sidebar.header("🔐 Authentification")
mode_admin = st.sidebar.checkbox("Mode Administrateur (Saisie / Modification)")

acces_accorde = False
if mode_admin:
    password = st.sidebar.text_input("Entrez le mot de passe administrateur", type="password")
    if password == "102030Tn":
        st.sidebar.success("✅ Accès Administrateur autorisé")
        acces_accorde = True
    elif password != "":
        st.sidebar.error("❌ Mot de passe incorrect")
else:
    st.sidebar.info("👁️ Mode consultation / recherche actif.")

if acces_accorde:
    st.subheader("✍️ Nouvelle Saisie Intelligente (Admin)")

    machines_existantes = (
        sorted([m for m in df_display["MACHINE"].unique().tolist() if m])
        if "MACHINE" in df_display.columns and not df_display.empty
        else ["M01", "M02A"]
    )
    sections_existantes = (
        sorted([s for s in df_display["SECTION"].unique().tolist() if s])
        if "SECTION" in df_display.columns and not df_display.empty
        else ["HENRI(01)"]
    )
    codes_internes_connus = (
        sorted([c for c in df_display["Code Interne"].unique().tolist() if c])
        if "Code Interne" in df_display.columns and not df_display.empty
        else []
    )

    def chercher_infos(code_interne):
        if not code_interne or df_display.empty:
            return "", "", ""

        match = df_display[df_display["Code Interne"] == str(code_interne).strip()]
        if not match.empty:
            row = match.iloc[-1]
            n_mou = row["N MOULE"] if "N MOULE" in df_display.columns else ""
            nom = ""
            if "MOULES PROGRAMMES" in df_display.columns and row["MOULES PROGRAMMES"] != "":
                nom = row["MOULES PROGRAMMES"]
            elif "MOULES DEPROGRAMMES" in df_display.columns and row["MOULES DEPROGRAMMES"] != "":
                nom = row["MOULES DEPROGRAMMES"]
            mat = row["MATIERES"] if "MATIERES" in df_display.columns else ""
            return nom, n_mou, mat

        return "", "", ""

    for champ in [
        "txt_m_deprog",
        "txt_n_deprog",
        "txt_mat_deprog",
        "txt_m_prog",
        "txt_n_prog",
        "txt_mat_prog",
    ]:
        if champ not in st.session_state:
            st.session_state[champ] = ""

    def update_deprog():
        code = st.session_state.get("sel_c_deprog", "")
        nom, n_mou, mat = chercher_infos(code)
        st.session_state["txt_m_deprog"] = nom
        st.session_state["txt_n_deprog"] = n_mou
        st.session_state["txt_mat_deprog"] = mat

    def update_prog():
        code = st.session_state.get("sel_c_prog", "")
        nom, n_mou, mat = chercher_infos(code)
        st.session_state["txt_m_prog"] = nom
        st.session_state["txt_n_prog"] = n_mou
        st.session_state["txt_mat_prog"] = mat

    col1, col2, col3 = st.columns(3)

    with col1:
        s_date = st.date_input("Date", value=date.today())
        choix_mach = st.selectbox("Machine", machines_existantes)
        choix_section = st.selectbox("Section", sections_existantes)

    with col2:
        st.markdown("##### 🔴 Moule Déprogrammé")
        st.selectbox(
            "Code Interne",
            [""] + codes_internes_connus,
            key="sel_c_deprog",
            on_change=update_deprog,
        )
        st.text_input("Nom du Moule", key="txt_m_deprog")
        st.text_input("N° Moule", key="txt_n_deprog")
        st.text_input("Matière", key="txt_mat_deprog")

    with col3:
        st.markdown("##### 🟢 Moule Programmé")
        st.selectbox(
            "Code Interne",
            [""] + codes_internes_connus,
            key="sel_c_prog",
            on_change=update_prog,
        )
        st.text_input("Nom du Moule", key="txt_m_prog")
        st.text_input("N° Moule", key="txt_n_prog")
        st.text_input("Matière", key="txt_mat_prog")

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        pieces = st.number_input("Nombre de pièces produites", min_value=0, step=1)
    with col_b2:
        remarque = st.text_area("Remarque", value="R.A.S")

    btn_submit = st.button("💾 Enregistrer la saisie")

    if btn_submit:
        date_formatee = s_date.strftime("%d/%m/%Y")
        lignes_a_ajouter = []

        code_final_deprog = st.session_state.get("sel_c_deprog", "")
        code_final_prog = st.session_state.get("sel_c_prog", "")

        nom_deprog_val = st.session_state.get("txt_m_deprog", "").strip()
        n_deprog_val = st.session_state.get("txt_n_deprog", "").strip()
        mat_deprog_val = st.session_state.get("txt_mat_deprog", "").strip()

        nom_prog_val = st.session_state.get("txt_m_prog", "").strip()
        n_prog_val = st.session_state.get("txt_n_prog", "").strip()
        mat_prog_val = st.session_state.get("txt_mat_prog", "").strip()

        if nom_deprog_val or n_deprog_val or code_final_deprog:
            lignes_a_ajouter.append({
                "DATE": date_formatee,
                "N MOULE": n_deprog_val,
                "Code Interne": code_final_deprog,
                "MOULES PROGRAMMES": "",
                "MOULES DEPROGRAMMES": nom_deprog_val,
                "MACHINE": choix_mach,
                "SECTION": choix_section,
                "MATIERES": mat_deprog_val,
                "REMARQUE": remarque.strip(),
                "PIECES PRODUITES": 0,
            })

        if nom_prog_val or n_prog_val or code_final_prog:
            lignes_a_ajouter.append({
                "DATE": date_formatee,
                "N MOULE": n_prog_val,
                "Code Interne": code_final_prog,
                "MOULES PROGRAMMES": nom_prog_val,
                "MOULES DEPROGRAMMES": "",
                "MACHINE": choix_mach,
                "SECTION": choix_section,
                "MATIERES": mat_prog_val,
                "REMARQUE": remarque.strip(),
                "PIECES PRODUITES": int(pieces),
            })

        if lignes_a_ajouter:
            try:
                wb = openpyxl.load_workbook(EXCEL_FILE)
                ws = wb[SHEET_NAME]

                headers = [
                    str(ws.cell(row=1, column=c).value).strip()
                    for c in range(1, ws.max_column + 1)
                ]

                fill_transparent = PatternFill(fill_type=None)
                fill_bleu = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
                thin_side = Side(style="thin", color="D9D9D9")
                border_ligne = Border(
                    left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
                )

                last_row_idx = ws.max_row
                for ligne_dict in lignes_a_ajouter:
                    new_row_idx = ws.max_row + 1
                    est_bleu = new_row_idx % 2 == 1

                    for col_idx, col_name in enumerate(headers, start=1):
                        val = ligne_dict.get(col_name, "")
                        if col_name == "PIECES PRODUITES":
                            try:
                                val = int(val) if val != "" else 0
                            except Exception:
                                val = 0
                        cell = ws.cell(row=new_row_idx, column=col_idx, value=val)

                        if last_row_idx >= 2:
                            ref_cell = ws.cell(row=last_row_idx, column=col_idx)
                            cell.font = copy(ref_cell.font)
                            cell.alignment = copy(ref_cell.alignment)
                            cell.number_format = ref_cell.number_format
                        else:
                            cell.font = Font(name="Calibri", size=10)
                            cell.alignment = Alignment(vertical="center")

                        cell.fill = fill_bleu if est_bleu else fill_transparent
                        cell.border = border_ligne

                col_date_idx = headers.index("DATE") + 1 if "DATE" in headers else None

                for r in range(2, ws.max_row + 1):
                    fill_r = fill_bleu if (r % 2 == 1) else fill_transparent
                    for c in range(1, ws.max_column + 1):
                        cell_r = ws.cell(row=r, column=c)
                        if col_date_idx and c == col_date_idx:
                            cell_r.value = convertir_date_ou_serial(cell_r.value)
                            cell_r.number_format = "@"
                        cell_r.fill = fill_r
                        cell_r.border = border_ligne

                ws.freeze_panes = "A2"
                last_col_letter = get_column_letter(ws.max_column)
                ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

                largeurs_min = {
                    "DATE": 14,
                    "N MOULE": 16,
                    "Code Interne": 15,
                    "MOULES PROGRAMMES": 42,
                    "MOULES DEPROGRAMMES": 42,
                    "MACHINE": 12,
                    "SECTION": 16,
                    "MATIERES": 34,
                    "REMARQUE": 26,
                    "PIECES PRODUITES": 18,
                }
                for idx, h in enumerate(headers, start=1):
                    col_letter = get_column_letter(idx)
                    actuelle = ws.column_dimensions[col_letter].width or 0
                    cible = largeurs_min.get(h, 15)
                    if actuelle < cible:
                        ws.column_dimensions[col_letter].width = cible

                wb.save(EXCEL_FILE)
                st.success("✅ Saisies enregistrées avec succès !")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur lors de l'enregistrement : {e}")
        else:
            st.warning("⚠️ Veuillez remplir au moins les informations d'un moule.")
else:
    st.info("🔒 Le formulaire de saisie est verrouillé. Activez le Mode Administrateur dans le menu latéral.")

st.markdown("---")

st.subheader("📊 Historique et Recherche Intelligente")

if not df_display.empty:
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)

    with col_f1:
        machines_dispo = (
            ["Toutes"] + sorted([m for m in df_display["MACHINE"].unique().tolist() if m])
            if "MACHINE" in df_display.columns
            else ["Toutes"]
        )
        filtre_mach = st.selectbox("⚙️ Machine", machines_dispo)

    with col_f2:
        sections_dispo = (
            ["Toutes"] + sorted([s for s in df_display["SECTION"].unique().tolist() if s])
            if "SECTION" in df_display.columns
            else ["Toutes"]
        )
        filtre_sect = st.selectbox("🏭 Section", sections_dispo)

    with col_f3:
        codes_dispo = (
            ["Tous"] + sorted([c for c in df_display["Code Interne"].unique().tolist() if c])
            if "Code Interne" in df_display.columns
            else ["Tous"]
        )
        filtre_code = st.selectbox("🏷️ Code Interne", codes_dispo)

    with col_f4:
        filtre_type = st.selectbox(
            "🔄 Type de ligne", ["Tous", "Moules Programmés", "Moules Déprogrammés"]
        )

    recherche_libre = st.text_input(
        "🔎 Recherche rapide (nom de moule, N° moule, matière, remarque...)"
    )

    df_filtre = df_display.copy()

    if filtre_mach != "Toutes" and "MACHINE" in df_filtre.columns:
        df_filtre = df_filtre[df_filtre["MACHINE"] == filtre_mach]

    if filtre_sect != "Toutes" and "SECTION" in df_filtre.columns:
        df_filtre = df_filtre[df_filtre["SECTION"] == filtre_sect]

    if filtre_code != "Tous" and "Code Interne" in df_filtre.columns:
        df_filtre = df_filtre[df_filtre["Code Interne"] == filtre_code]

    if filtre_type == "Moules Programmés" and "MOULES PROGRAMMES" in df_filtre.columns:
        df_filtre = df_filtre[df_filtre["MOULES PROGRAMMES"] != ""]
    elif filtre_type == "Moules Déprogrammés" and "MOULES DEPROGRAMMES" in df_filtre.columns:
        df_filtre = df_filtre[df_filtre["MOULES DEPROGRAMMES"] != ""]

    if recherche_libre.strip():
        terme = recherche_libre.strip().lower()
        masque = df_filtre.apply(
            lambda row: row.astype(str).str.lower().str.contains(terme, na=False).any(), axis=1
        )
        df_filtre = df_filtre[masque]

    col_info, col_dl = st.columns([3, 1])
    with col_info:
        st.write(f"Affichage de **{len(df_filtre)}** ligne(s) sur **{len(df_display)}** au total.")

    if acces_accorde and not df_filtre.empty:
        with col_dl:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_filtre.to_excel(writer, sheet_name=SHEET_NAME, index=False)
            st.download_button(
                label="📊 Exporter en Excel (.xlsx)",
                data=output.getvalue(),
                file_name="suivie_changements_filtre.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    column_config_map = {
        "DATE": st.column_config.TextColumn("DATE", width="small"),
        "N MOULE": st.column_config.TextColumn("N MOULE", width="small"),
        "Code Interne": st.column_config.TextColumn("Code Interne", width="small"),
        "MOULES PROGRAMMES": st.column_config.TextColumn("MOULES PROGRAMMES", width="large"),
        "MOULES DEPROGRAMMES": st.column_config.TextColumn("MOULES DEPROGRAMMES", width="large"),
        "MACHINE": st.column_config.TextColumn("MACHINE", width="small"),
        "SECTION": st.column_config.TextColumn("SECTION", width="medium"),
        "MATIERES": st.column_config.TextColumn("MATIERES", width="medium"),
        "REMARQUE": st.column_config.TextColumn("REMARQUE", width="medium"),
        "PIECES PRODUITES": st.column_config.NumberColumn(
            "PIECES PRODUITES", width="small", format="%d"
        ),
    }
    active_col_config = {k: v for k, v in column_config_map.items() if k in df_filtre.columns}

    st.dataframe(
        df_filtre,
        use_container_width=True,
        hide_index=True,
        height=500,
        column_config=active_col_config,
    )
else:
    st.info("Aucune donnée à afficher pour le moment.")