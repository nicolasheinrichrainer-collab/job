import io
import mimetypes
import re
import shutil
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Suivi clients", page_icon="🚗", layout="wide")

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
DEFAULT_FILE = "clients.xlsx"

INFO_COLS = ["NOM", "PRENOM", "CHASSY", "IMMAT", "EMAIL"]

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
# Mail : modèle "documents à retourner" + brouillon Outlook (.eml)
# ----------------------------------------------------------------------------
# --- MAIL-START (bloc testé séparément) ---
# Dossier (à côté de ce fichier) où déposer les pièces jointes communes :
# le mandat vierge, le RIB... JAMAIS de documents personnels de clients.
PJ_DIR = Path(__file__).resolve().parent / "pieces_jointes"

IBAN = "FR76 1474 9000 0400 8117 8001 053"
SIGNATURE = "Nicolas,\nL'équipe E-Store Vo Corporate\n09 73 96 00 00"

CENTRES = {
    "POISSY": "CENTRE DE LIVRAISON POISSY CLRP2 - 45 rue Jean Pierre Timbaud - 78300 Poissy",
    "RENNES": "CENTRE DE LIVRAISON RENNES - PSA La Janais - 35131 Chartres de Bretagne",
    "HORDAIN": "CENTRE DE LIVRAISON HORDAIN - Avenue Jean Monnet - 59111 Lieu-St-Amand",
    "SOCHAUX": "CENTRE DE LIVRAISON SOCHAUX - 16 Av. d'Helvétie - 25200 Montbéliard",
    "MULHOUSE": "USINE STELLANTIS - Rue de Chalampé - 68300 Sausheim",
    "VITROLLES": "ZAC DES CADESTEAUX - Avenue Joseph Cugnot - 13127 Vitrolles",
}
TOUS_CENTRES = "Tous les centres (liste complète)"

DISCLAIMER = (
    "« Ce message électronique et tous les fichiers attachés qu'il contient sont "
    "confidentiels et destinés exclusivement à l'usage de la personne à laquelle ils "
    "sont adressés. Si vous avez reçu ce message par erreur, merci de le retourner à "
    "son émetteur. La publication, l'usage, la distribution, l'impression ou la copie "
    "non autorisée de ce message et des attachements qu'il contient sont strictement "
    "interdits. »"
)

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_valid_email(value: str) -> bool:
    return bool(EMAIL_RE.match((value or "").strip()))


def mail_subject(immat: str) -> str:
    suffix = f" {immat}" if immat else ""
    return f"Votre commande{suffix} – documents à nous retourner"


def mail_body(nom: str, prenom: str, immat: str, centre: str, avec_proche: bool) -> str:
    """Construit le texte du mail à partir des infos du client."""
    ref = f"{immat} – {nom} {prenom}".strip() if immat else "IMMAT DU VEHICULE – NOM ET PRENOM du collaborateur"

    if centre == TOUS_CENTRES:
        livraison = "\n".join(f"{ville} : {adr}" for ville, adr in CENTRES.items())
    else:
        livraison = f"{centre} : {CENTRES[centre]}"

    parts = [
        "Bonjour,",
        "Félicitations pour votre choix ! Nous sommes ravis de vous accompagner dans cette "
        "acquisition et de préparer ensemble la livraison de votre véhicule.",
        "Afin de pouvoir lancer rapidement les démarches logistiques et garantir une mise à "
        "disposition dans les meilleures conditions, nous vous invitons à nous retourner "
        "l'ensemble des documents demandés en réponse à ce mail dans un délai maximum de 7 jours.",
        "Ce délai nous permet de réserver votre véhicule et d'organiser sa livraison le plus "
        "rapidement possible.\n"
        "(Passé ce délai, la commande sera annulée et le véhicule remis en vente.)",
        "Documents à nous transmettre par retour de ce mail :\n"
        "- Le bon de commande et les conditions générales de vente (les 23 premières pages uniquement)\n"
        "    - Le bon de commande, paraphé à chaque page\n"
        "    - Les conditions générales de vente, datées et signées uniquement page 21.\n"
        "- Le règlement sera à effectuer dans un délai de 15 jours à compter de la réception du "
        "dossier complet par notre service, en effectuant un virement (RIB joint).\n"
        f"    IBAN {IBAN}\n"
        f"    La référence du virement doit être : {ref}\n"
        "- Copie recto/verso d'une pièce d'identité en cours de validité\n"
        "- Un justificatif de domicile de moins de 3 mois (facture d'électricité, gaz, téléphone, "
        "box Internet, ou avis d'imposition)\n"
        "- Photocopie recto/verso du permis de conduire\n"
        "- Le mandat (joint en pièce jointe) autorisant STELLANTIS AND YOU à effectuer la demande "
        "d'immatriculation, dûment complété, daté et signé.",
    ]

    if avec_proche:
        parts.append(
            "Si votre commande est pour un proche :\n"
            "- Copie recto/verso de la pièce d'identité en cours de validité\n"
            "- Justificatif de domicile de moins de 3 mois (facture d'électricité, gaz, téléphone, "
            "box Internet, ou avis d'imposition)\n"
            "- Copie recto/verso du permis de conduire\n"
            "- Le mandat à son nom (joint en pièce jointe) autorisant STELLANTIS AND YOU à effectuer "
            "la demande d'immatriculation, dûment complété, daté et signé.\n"
            "- Adresse mail personnelle du proche"
        )

    parts += [
        "Votre Livraison :\n"
        "Votre véhicule sera livré à l'adresse suivante :\n\n" + livraison,
        "Dès que votre véhicule sera prêt, vous serez recontacté afin de convenir d'un rendez-vous. "
        "Vous disposerez alors de 15 jours maximum pour en prendre possession.",
        "Les modifications de rendez-vous sont tolérées, mais doivent rester exceptionnelles.",
    ]

    if avec_proche:
        parts.append(
            "Rappel des règles d'une commande pour un proche :\n"
            "- Le règlement doit être effectué au comptant : aucune possibilité de financement "
            "pour une commande proche\n"
            "- Aucun cotitulaire autre que le collaborateur n'est autorisé\n"
            "- Le bon de commande et la facture seront au nom du collaborateur\n"
            "- Vous ne pouvez pas effectuer plus de 4 commandes par an pour un proche"
        )

    parts += [
        "Nous restons à votre disposition pour toute question ou précision.",
        "Cordialement,",
        SIGNATURE,
        DISCLAIMER,
    ]
    return "\n\n".join(parts)


def build_eml(to: str, subject: str, body: str, attachments: list) -> bytes:
    """Brouillon .eml (en-tête X-Unsent) : Outlook l'ouvre comme un mail à envoyer.

    attachments : liste de tuples (nom_de_fichier, contenu_en_bytes)
    """
    msg = EmailMessage()
    if to:
        msg["To"] = to
    msg["Subject"] = subject
    msg["X-Unsent"] = "1"
    msg.set_content(body)
    for name, data in attachments:
        ctype, _ = mimetypes.guess_type(name)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
    return msg.as_bytes()


def list_pj() -> list:
    if not PJ_DIR.is_dir():
        return []
    return sorted(p for p in PJ_DIR.iterdir() if p.is_file() and not p.name.startswith("."))
# --- MAIL-END ---


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
tab_suivi, tab_new, tab_mail, tab_dash, tab_gestion = st.tabs(
    ["📋 Suivi", "➕ Nouveau client", "✉️ Mail", "📊 Tableau de bord", "🗑️ Supprimer"]
)

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
        "EMAIL": st.column_config.TextColumn("Email"),
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

# ---------------------------- Mail -------------------------------------------
with tab_mail:
    st.subheader("Préparer le mail « documents à retourner »")

    flash_mail = st.session_state.pop("flash_mail", None)
    if flash_mail:
        st.success(flash_mail)

    if df.empty:
        st.info("Aucun client dans le tableau pour le moment.")
    else:
        mail_labels = {
            i: f"{r.NOM} {r.PRENOM} — {r.IMMAT or r.CHASSY or 'sans immat'}"
            for i, r in df.iterrows()
        }
        idx = st.selectbox(
            "Client",
            options=list(mail_labels),
            format_func=mail_labels.get,
            index=None,
            placeholder="Choisis un client (tu peux taper pour chercher)…",
        )

        if idx is not None:
            r = df.loc[idx]

            c1, c2 = st.columns(2)
            centre = c1.selectbox("Centre de livraison", [TOUS_CENTRES] + list(CENTRES))
            avec_proche = c2.checkbox(
                "Inclure la partie « commande pour un proche »",
                value=True,
                help="Décoche si le client commande pour lui-même.",
            )

            to = st.text_input("Adresse mail du client", value=r["EMAIL"], key=f"mail_to_{idx}")
            if to and not is_valid_email(to):
                st.warning("Cette adresse mail ne semble pas valide.")
            if to.strip() != r["EMAIL"] and is_valid_email(to):
                if st.button("💾 Enregistrer cette adresse dans le tableau"):
                    st.session_state.df.at[idx, "EMAIL"] = to.strip()
                    persist()
                    st.session_state["flash_mail"] = "✅ Adresse mail enregistrée."
                    st.rerun()
            if not r["IMMAT"]:
                st.warning(
                    "Pas d'immatriculation pour ce client : la référence du virement restera "
                    "à compléter dans le texte du mail."
                )

            subject = st.text_input("Objet", value=mail_subject(r["IMMAT"]),
                                    key=f"mail_subj_{idx}")
            # La clé change quand le centre ou l'option "proche" change : le texte est
            # alors régénéré. Tant qu'elles ne changent pas, tes retouches sont conservées.
            body = st.text_area(
                "Texte du mail (modifiable avant l'envoi)",
                value=mail_body(r["NOM"], r["PRENOM"], r["IMMAT"], centre, avec_proche),
                height=480,
                key=f"mail_body_{idx}_{centre}_{avec_proche}",
            )

            st.markdown("**Pièces jointes**")
            pj_files = list_pj()
            chosen = st.multiselect(
                "Fichiers du dossier « pieces_jointes »",
                options=[p.name for p in pj_files],
                default=[p.name for p in pj_files],
            )
            extra = st.file_uploader("Ajouter d'autres fichiers (facultatif)",
                                     accept_multiple_files=True)
            if not pj_files and not extra:
                st.info(
                    "Aucune pièce jointe. Dépose le mandat vierge et le RIB dans un dossier "
                    "« pieces_jointes » à côté de ce script, ou ajoute-les ci-dessus. "
                    "N'y mets jamais de documents personnels de clients."
                )

            attachments = [(p.name, p.read_bytes()) for p in pj_files if p.name in chosen]
            attachments += [(f.name, f.getvalue()) for f in extra]

            eml = build_eml(to.strip(), subject, body, attachments)
            safe = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{r['NOM']}_{r['IMMAT']}").strip("_") or "client"

            d1, d2 = st.columns(2)
            d1.download_button(
                "📨 Télécharger le brouillon Outlook",
                data=eml,
                file_name=f"mail_{safe}.eml",
                mime="message/rfc822",
                type="primary",
                use_container_width=True,
                help="Ouvre le fichier téléchargé : le mail s'affiche dans Outlook, "
                     "prêt à être relu et envoyé.",
            )
            if d2.button("✅ Marquer « BDC envoyé » pour ce client", use_container_width=True,
                         disabled=bool(r["BDC_ENVOYE"])):
                st.session_state.df.at[idx, "BDC_ENVOYE"] = True
                persist()
                st.session_state["flash_mail"] = f"✅ BDC envoyé coché pour {r['NOM']} {r['PRENOM']}."
                st.rerun()
            if r["BDC_ENVOYE"]:
                st.caption("« BDC envoyé » est déjà coché pour ce client.")

            with st.expander("Le brouillon ne s'ouvre pas ? Copier le texte à la main"):
                st.code(body, language=None)

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

# ---------------------------- Nouveau client ---------------------------------
with tab_new:
    st.subheader("Saisir un nouveau client")

    # Message de confirmation conservé après le rechargement de la page
    flash = st.session_state.pop("flash", None)
    if flash:
        st.success(flash)

    with st.form("new_client", clear_on_submit=True):
        st.markdown("**Informations client**")
        a, b = st.columns(2)
        nom = a.text_input("Nom *")
        prenom = b.text_input("Prénom *")
        chassy = a.text_input("Châssis")
        immat = b.text_input("Immat")
        email = st.text_input("Adresse mail du client")

        st.markdown("**Étapes déjà réalisées** (laisse décoché ce qui reste à faire)")
        step_cols = st.columns(3)
        new_steps = {}
        for i, s in enumerate(STEPS):
            new_steps[s] = step_cols[i % 3].checkbox(LABELS[s], key=f"new_{s}")

        submitted = st.form_submit_button("➕ Ajouter au tableau", type="primary")

    if submitted:
        nom, prenom = nom.strip(), prenom.strip()
        chassy, immat, email = chassy.strip(), immat.strip(), email.strip()
        if not (nom and prenom):
            st.warning("Le nom et le prénom sont obligatoires.")
        elif email and not is_valid_email(email):
            st.warning("L'adresse mail ne semble pas valide. Rien n'a été ajouté.")
        else:
            dup = df[
                ((df["IMMAT"] != "") & (df["IMMAT"].str.upper() == immat.upper()) & bool(immat))
                | ((df["CHASSY"] != "") & (df["CHASSY"].str.upper() == chassy.upper()) & bool(chassy))
            ]
            if not dup.empty:
                d = dup.iloc[0]
                st.warning(
                    f"Un client avec cette immat ou ce châssis existe déjà : "
                    f"{d['NOM']} {d['PRENOM']}. Rien n'a été ajouté."
                )
            else:
                row = {c: "" for c in df.columns}
                row.update(NOM=nom, PRENOM=prenom, CHASSY=chassy, IMMAT=immat, EMAIL=email)
                row.update(new_steps)
                st.session_state.df = pd.concat(
                    [df, pd.DataFrame([row])], ignore_index=True
                )
                persist()
                st.session_state["flash"] = f"✅ {nom} {prenom} a été ajouté au tableau."
                st.rerun()

# ---------------------------- Supprimer --------------------------------------
with tab_gestion:
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
