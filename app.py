import streamlit as st
import pandas as pd
import os
from datetime import date
from PIL import Image
import io

# Configuration de la page
st.set_page_config(page_title="Production & Changements de Moules", layout="wide")

EXCEL_FILE = "suivie des changement.xlsx"
SHEET_NAME = "Suivie"
LOGO_FILE = "logo.png"

# En-tête avec Logo, Titre et Signature
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


def formater_colonne_date(serie):
    def _parse_une_date(val):
        if pd.isna(val):
            return ""
        s = str(val).strip()
        if not s or s.lower() in ("nan", "nat", "none"):
            return ""
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                dt = pd.to_datetime(s, format=fmt)
                return dt.strftime("%d/%m/%Y")
            except Exception:
                pass
        try:
            dt = pd.to_datetime(s, dayfirst=True)
            if pd.notna(dt):
                return dt.strftime("%d/%m/%Y")
        except Exception:
            pass
        return s

    return serie.apply(_parse_une_date)


# Fonction pour charger et nettoyer les données en toute sécurité
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
                df["PIECES PRODUITES"] = pd.to_numeric(df["PIECES PRODUITES"], errors="coerce").fillna(0).astype(int)

            for col in df.columns:
                if col not in ["PIECES PRODUITES", "DATE"]:
                    df[col] = df[col].fillna("").astype(str).str.strip()
                    df[col] = df[col].replace({"nan": "", "None": ""})

            return df
        except Exception as e:
            st.error(f"⚠️ Erreur de lecture du fichier Excel : {e}")
            return pd.DataFrame()
    else:
        st.warning(f"⚠️ Le fichier '{EXCEL_FILE}' est introuvable.")
        return pd.DataFrame()


df_display = charger_donnees()

# Gestion des accès Administrateur
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

# Formulaire de saisie intelligent (Admin)
if acces_accorde:
    st.subheader("✍️ Nouvelle Saisie Intelligente (Admin)")

    machines_existantes = sorted([m for m in df_display["MACHINE"].unique().tolist() if m]) if "MACHINE" in df_display.columns and not df_display.empty else ["M01", "M02A"]
    sections_existantes = sorted([s for s in df_display["SECTION"].unique().tolist() if s]) if "SECTION" in df_display.columns and not df_display.empty else ["HENRI(01)"]
    codes_internes_connus = sorted([c for c in df_display["Code Interne"].unique().tolist() if c]) if "Code Interne" in df_display.columns and not df_display.empty else []

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

    for champ in ["txt_m_deprog", "txt_n_deprog", "txt_mat_deprog", "txt_m_prog", "txt_n_prog", "txt_mat_prog"]:
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
        st.selectbox("Code Interne", [""] + codes_internes_connus, key="sel_c_deprog", on_change=update_deprog)
        m_deprog = st.text_input("Nom du Moule", key="txt_m_deprog")
        n_moule_deprog = st.text_input("N° Moule", key="txt_n_deprog")
        matiere_deprog = st.text_input("Matière", key="txt_mat_deprog")

    with col3:
        st.markdown("##### 🟢 Moule Programmé")
        st.selectbox("Code Interne", [""] + codes_internes_connus, key="sel_c_prog", on_change=update_prog)
        m_prog = st.text_input("Nom du Moule", key="txt_m_prog")
        n_moule_prog = st.text_input("N° Moule", key="txt_n_prog")
        matiere_prog = st.text_input("Matière", key="txt_mat_prog")

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
                "PIECES PRODUITES": 0
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
                "PIECES PRODUITES": int(pieces)
            })

        if lignes_a_ajouter:
            try:
                df_actuel = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
                df_actuel.columns = df_actuel.columns.astype(str).str.strip()
                if "MACH" in df_actuel.columns:
                    df_actuel = df_actuel.drop(columns=["MACH"])

                df_nouveau = pd.concat([df_actuel, pd.DataFrame(lignes_a_ajouter)], ignore_index=True)
                with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df_nouveau.to_excel(writer, sheet_name=SHEET_NAME, index=False)

                st.success("✅ Saisies enregistrées avec succès !")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur lors de l'enregistrement : {e}")
        else:
            st.warning("⚠️ Veuillez remplir au moins les informations d'un moule.")
else:
    st.info("🔒 Le formulaire de saisie est verrouillé. Activez le Mode Administrateur dans le menu latéral.")

st.markdown("---")

# Historique et Recherche Intelligente
st.subheader("📊 Historique et Recherche Intelligente")

if not df_display.empty:
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)

    with col_f1:
        machines_dispo = ["Toutes"] + sorted([m for m in df_display["MACHINE"].unique().tolist() if m]) if "MACHINE" in df_display.columns else ["Toutes"]
        filtre_mach = st.selectbox("⚙️ Machine", machines_dispo)

    with col_f2:
        sections_dispo = ["Toutes"] + sorted([s for s in df_display["SECTION"].unique().tolist() if s]) if "SECTION" in df_display.columns else ["Toutes"]
        filtre_sect = st.selectbox("🏭 Section", sections_dispo)

    with col_f3:
        codes_dispo = ["Tous"] + sorted([c for c in df_display["Code Interne"].unique().tolist() if c]) if "Code Interne" in df_display.columns else ["Tous"]
        filtre_code = st.selectbox("🏷️ Code Interne", codes_dispo)

    with col_f4:
        filtre_type = st.selectbox("🔄 Type de ligne", ["Tous", "Moules Programmés", "Moules Déprogrammés"])

    recherche_libre = st.text_input("🔎 Recherche rapide (nom de moule, N° moule, matière, remarque...)")

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
        masque = df_filtre.apply(lambda row: row.astype(str).str.lower().str.contains(terme, na=False).any(), axis=1)
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
                use_container_width=True
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
        "PIECES PRODUITES": st.column_config.NumberColumn("PIECES PRODUITES", width="small", format="%d")
    }
    active_col_config = {k: v for k, v in column_config_map.items() if k in df_filtre.columns}

    st.dataframe(
        df_filtre,
        use_container_width=True,
        hide_index=True,
        height=500,
        column_config=active_col_config
    )
else:
    st.info("Aucune donnée à afficher pour le moment.")