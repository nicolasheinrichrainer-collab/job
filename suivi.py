import io
import shutil
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Suivi clients", page_icon="🚗", layout="wide")

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
DEFAULT_FILE = "clients.xlsx"

INFO_COLS = ["NOM", "PRENOM", "CHASSY", "IMMAT"]

# L'ordre = l'ordre des étapes du dossier. La "prochaine étape" est la
# première colonne encore à NON.
STEPS = [
    "VEHICULE_PROCHE",
    "BDC_ENVOYE",
    "BDC_SIGNE",
    "COMMANDE_VISEE",
    "PAIEMENT",
    "ADMIN_EFFICOM",
    "GARANTIE_SPOTICAR",
    "FACTURE_DCSNT",
    "CARTE_GRISE",
]

LABELS = {
    "VEHICULE_PROCHE": "Véhicule proche",
    "BDC_ENVOYE": "BDC envoyé",
    "BDC_SIGNE": "BDC signé",
    "COMMANDE_VISEE": "Commande visée",
    "PAIEMENT": "Paiement",
    "ADMIN_EFFICOM": "Admin Efficom",
    "GARANTIE_SPOTICAR": "Garantie Spoticar",
    "FACTURE_DCSNT": "Facture DCSNT",
    "CARTE_GRISE": "Carte grise",
}

DONE = "Terminé ✅"
TRUE_VALUES = {"OUI", "O", "YES", "Y", "1", "TRUE", "VRAI", "X"}


# ----------------------------------------------------------------------------
# Fonctions utilitaires
# ----------------------------------------------------------------------------
def to_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return False
    return str(value).strip().upper() in TRUE_VALUES


def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie le tableau : en-têtes, texte, OUI/NON -> booléens."""
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    for c in INFO_COLS:
        if c not in df.columns:
            df[c] = ""
    for c in STEPS:
        if c not in df.columns:
            df[c] = False
    extras = [c for c in df.columns if c not in INFO_COLS + STEPS]
    df = df[INFO_COLS + STEPS + extras]
    for c in INFO_COLS:
        df[c] = df[c].fillna("").astype(str).str.strip()
    for c in STEPS:
        df[c] = df[c].map(to_bool).astype(bool)
    return df.reset_index(drop=True)


def to_excel_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Reconvertit les booléens en OUI/NON pour l'export Excel."""
    out = df.copy()
    for c in STEPS:
        out[c] = out[c].map(lambda b: "OUI" if b else "NON")
    return out


def save_to_excel(df: pd.DataFrame, path: str) -> None:
    p = Path(path)
    if p.exists():
        shutil.copy(p, p.with_suffix(".bak.xlsx"))  # petite sauvegarde de sécurité
    to_excel_frame(df).to_excel(p, index=False)


def excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    to_excel_frame(df).to_excel(buf, index=False)
    return buf.getvalue()


def next_step(row) -> str:
    for s in STEPS:
        if not row[s]:
            return LABELS[s]
    return DONE


def add_computed(df: pd.DataFrame) -> pd.DataFrame:
    view = df.copy()
    view["PROGRESSION"] = (view[STEPS].sum(axis=1) / len(STEPS) * 100).round()
    view["PROCHAINE ÉTAPE"] = view.apply(next_step, axis=1)
    return view


# ----------------------------------------------------------------------------
# Chargement des données
# ----------------------------------------------------------------------------
st.title("🚗 Suivi des clients")

with st.sidebar:
    st.header("Fichier")
    path = st.text_input("Chemin du fichier Excel", DEFAULT_FILE)
    autosave = st.checkbox("Sauvegarde automatique", value=True,
                           help="Enregistre dans le fichier Excel à chaque modification.")
    if st.button("🔄 Recharger depuis le fichier", use_container_width=True):
        st.session_state.pop("df", None)

if st.session_state.get("path") != path:
    st.session_state.pop("df", None)
    st.session_state["path"] = path

if "df" not in st.session_state:
    if Path(path).exists():
        try:
            st.session_state.df = normalize(pd.read_excel(path, dtype=object))
        except Exception as e:  # fichier ouvert/corrompu, etc.
            st.error(f"Impossible de lire « {path} » : {e}")
            st.stop()
    else:
        st.info(f"Le fichier « {path} » n'existe pas encore.")
        up = st.file_uploader("Importer ton fichier Excel", type=["xlsx", "xls"])
        if up is not None:
            st.session_state.df = normalize(pd.read_excel(up, dtype=object))
            st.rerun()
        if st.button("Démarrer avec un tableau vide"):
            st.session_state.df = normalize(pd.DataFrame(columns=INFO_COLS + STEPS))
            st.rerun()
        st.stop()

df: pd.DataFrame = st.session_state.df


def persist():
    if autosave:
        try:
            save_to_excel(st.session_state.df, path)
        except Exception as e:
            st.error(f"Sauvegarde impossible (fichier ouvert dans Excel ?) : {e}")


# ----------------------------------------------------------------------------
# Onglets
# ----------------------------------------------------------------------------
tab_suivi, tab_dash, tab_gestion = st.tabs(["📋 Suivi", "📊 Tableau de bord", "➕ Ajouter / supprimer"])

# ---------------------------- Suivi -----------------------------------------
with tab_suivi:
    full = add_computed(df)

    c1, c2, c3 = st.columns([2, 2, 1])
    search = c1.text_input("🔎 Recherche (nom, prénom, châssis, immat)", "")
    steps_filter = c2.multiselect(
        "Prochaine étape",
        options=[LABELS[s] for s in STEPS] + [DONE],
        placeholder="Toutes",
    )
    hide_done = c3.toggle("Masquer les terminés", value=False)

    mask = pd.Series(True, index=full.index)
    if search.strip():
        q = search.strip().lower()
        mask &= full[INFO_COLS].apply(
            lambda col: col.str.lower().str.contains(q, regex=False)
        ).any(axis=1)
    if steps_filter:
        mask &= full["PROCHAINE ÉTAPE"].isin(steps_filter)
    if hide_done:
        mask &= full["PROCHAINE ÉTAPE"] != DONE

    view = full[mask]
    st.caption(f"{len(view)} dossier(s) affiché(s) sur {len(full)}")

    edit_cols = INFO_COLS + STEPS
    column_config = {
        "NOM": st.column_config.TextColumn("Nom"),
        "PRENOM": st.column_config.TextColumn("Prénom"),
        "CHASSY": st.column_config.TextColumn("Châssis"),
        "IMMAT": st.column_config.TextColumn("Immat"),
        "PROGRESSION": st.column_config.ProgressColumn(
            "Progression", min_value=0, max_value=100, format="%d%%"
        ),
        "PROCHAINE ÉTAPE": st.column_config.TextColumn("Prochaine étape"),
    }
    for s in STEPS:
        column_config[s] = st.column_config.CheckboxColumn(LABELS[s])

    ordered = INFO_COLS + ["PROGRESSION", "PROCHAINE ÉTAPE"] + STEPS
    edited = st.data_editor(
        view[ordered],
        column_config=column_config,
        disabled=["PROGRESSION", "PROCHAINE ÉTAPE"],
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        height=min(600, 80 + 35 * max(len(view), 1)),
        key="editor",
    )

    if not edited[edit_cols].equals(view[edit_cols]):
        st.session_state.df.loc[edited.index, edit_cols] = edited[edit_cols]
        st.session_state.df[STEPS] = st.session_state.df[STEPS].astype(bool)
        persist()
        st.rerun()

    b1, b2, _ = st.columns([1, 1, 4])
    if b1.button("💾 Enregistrer", type="primary"):
        try:
            save_to_excel(st.session_state.df, path)
            st.success("Enregistré dans le fichier Excel.")
        except Exception as e:
            st.error(f"Sauvegarde impossible : {e}")
    b2.download_button(
        "⬇️ Télécharger",
        data=excel_bytes(st.session_state.df),
        file_name="clients.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------- Tableau de bord --------------------------------
with tab_dash:
    full = add_computed(df)
    total = len(full)
    termines = int((full["PROCHAINE ÉTAPE"] == DONE).sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Dossiers", total)
    m2.metric("Terminés", termines)
    m3.metric("En cours", total - termines)
    m4.metric("Progression moyenne", f"{full['PROGRESSION'].mean():.0f}%" if total else "—")

    st.subheader("Dossiers en attente par étape")
    pending = pd.DataFrame(
        {"Étape": [LABELS[s] for s in STEPS], "À faire": [int((~df[s]).sum()) for s in STEPS]}
    ).set_index("Étape")
    st.bar_chart(pending)

    st.subheader("Cas à surveiller")
    alerts = {
        "BDC signé mais pas de paiement": df["BDC_SIGNE"] & ~df["PAIEMENT"],
        "Payé mais pas de facture": df["PAIEMENT"] & ~df["FACTURE_DCSNT"],
        "Facturé mais pas de carte grise": df["FACTURE_DCSNT"] & ~df["CARTE_GRISE"],
        "BDC envoyé mais pas signé": df["BDC_ENVOYE"] & ~df["BDC_SIGNE"],
    }
    for title, cond in alerts.items():
        with st.expander(f"{title} ({int(cond.sum())})"):
            if cond.any():
                st.dataframe(df.loc[cond, INFO_COLS], hide_index=True, use_container_width=True)
            else:
                st.write("Aucun dossier concerné.")

# ---------------------------- Ajouter / supprimer ----------------------------
with tab_gestion:
    st.subheader("Nouveau client")
    with st.form("new_client", clear_on_submit=True):
        a, b = st.columns(2)
        nom = a.text_input("Nom")
        prenom = b.text_input("Prénom")
        chassy = a.text_input("Châssis")
        immat = b.text_input("Immat")
        submitted = st.form_submit_button("Ajouter", type="primary")

    if submitted:
        if not (nom.strip() or prenom.strip()):
            st.warning("Renseigne au moins un nom ou un prénom.")
        else:
            dup = df[
                ((df["IMMAT"] != "") & (df["IMMAT"].str.upper() == immat.strip().upper()))
                | ((df["CHASSY"] != "") & (df["CHASSY"].str.upper() == chassy.strip().upper()))
            ]
            if not dup.empty and (immat.strip() or chassy.strip()):
                st.warning("Un client avec cette immat ou ce châssis existe déjà.")
            else:
                row = {c: "" for c in df.columns}
                row.update(NOM=nom.strip(), PRENOM=prenom.strip(),
                           CHASSY=chassy.strip(), IMMAT=immat.strip())
                row.update({s: False for s in STEPS})
                st.session_state.df = pd.concat(
                    [df, pd.DataFrame([row])], ignore_index=True
                )
                persist()
                st.success(f"{nom} {prenom} ajouté.")
                st.rerun()

    st.divider()
    st.subheader("Supprimer des clients")
    labels = {
        i: f"{r.NOM} {r.PRENOM} — {r.IMMAT or r.CHASSY or 'sans immat'}"
        for i, r in df.iterrows()
    }
    to_delete = st.multiselect("Clients à supprimer", options=list(labels), format_func=labels.get)
    confirm = st.checkbox("Je confirme la suppression")
    if st.button("🗑️ Supprimer", disabled=not (to_delete and confirm)):
        st.session_state.df = df.drop(index=to_delete).reset_index(drop=True)
        persist()
        st.rerun()
