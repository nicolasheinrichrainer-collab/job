import io
import mimetypes
import re
import shutil
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Suivi clients", page_icon="🚗", layout="wide")

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
DEFAULT_FILE = "clients.xlsx"
STALE_DAYS_DEFAULT = 7  # un dossier est "bloqué" après ce nombre de jours sans avancée

CORE_COLS = ["NOM", "PRENOM", "CHASSY", "IMMAT"]
CONTACT_COLS = ["EMAIL", "TELEPHONE"]
INFO_COLS = CORE_COLS + CONTACT_COLS

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

# Colonnes de suivi ajoutées au fichier Excel
TEXT_EXTRA = ["NOTES", "STATUT", "MOTIF_ANNULATION"]
STEP_DATE_COLS = [f"DATE_{s}" for s in STEPS]
DATE_COLS = ["DERNIER_CONTACT", "RELANCE_LE"] + STEP_DATE_COLS

STATUT_ACTIF = "EN COURS"
STATUT_ANNULE = "ANNULÉ / PERDU"

DONE = "Terminé ✅"
TRUE_VALUES = {"OUI", "O", "YES", "Y", "1", "TRUE", "VRAI", "X"}
DATE_FMT = "DD/MM/YYYY"


# ----------------------------------------------------------------------------
# Fonctions utilitaires
# ----------------------------------------------------------------------------
def today_fr() -> date:
    """Date du jour en heure de Paris (le serveur Streamlit Cloud est en UTC)."""
    try:
        return datetime.now(ZoneInfo("Europe/Paris")).date()
    except Exception:
        return date.today()


def to_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return False
    return str(value).strip().upper() in TRUE_VALUES


def clean_text(v) -> str:
    """Texte propre : pas de NaN, pas de '612345678.0' venant d'Excel."""
    if v is None or v is pd.NaT:
        return ""
    if isinstance(v, float):
        if pd.isna(v):
            return ""
        if v.is_integer():
            return str(int(v))
    return str(v).strip()


def norm_phone(v) -> str:
    raw = clean_text(v)
    if not raw or raw.startswith("+"):
        return raw
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 9:  # le 0 de tête a sauté dans Excel
        digits = "0" + digits
    if len(digits) == 10 and digits[0] == "0":
        return " ".join(digits[i:i + 2] for i in range(0, 10, 2))
    return raw


def norm_immat(v):
    """Retourne (immat_propre, avertissement ou None). Format attendu : AB-123-CD."""
    raw = re.sub(r"[\s\-]", "", clean_text(v)).upper()
    if not raw:
        return "", None
    m = re.fullmatch(r"([A-Z]{2})(\d{3})([A-Z]{2})", raw)
    if m:
        return "-".join(m.groups()), None
    return clean_text(v).upper(), "Immat au format inhabituel (attendu : AB-123-CD)."


VIN_RE = re.compile(r"^[A-HJ-NPR-Z0-9]{17}$")


def norm_chassis(v):
    """Retourne (châssis_propre, avertissement ou None). Un numéro de série fait 17 caractères."""
    raw = re.sub(r"\s", "", clean_text(v)).upper()
    if not raw:
        return "", None
    if VIN_RE.match(raw):
        return raw, None
    if len(raw) != 17:
        return raw, f"Le châssis fait {len(raw)} caractères au lieu de 17."
    return raw, "Le châssis contient des caractères non autorisés (I, O, Q ou symboles)."


def dedupe_key(v) -> str:
    return re.sub(r"[\s\-]", "", clean_text(v)).upper()


def find_duplicates(df: pd.DataFrame, immat: str, chassy: str, exclude=None) -> pd.DataFrame:
    """Clients ayant déjà cette immat ou ce châssis (sans tenir compte des tirets/espaces/casse)."""
    ki, kc = dedupe_key(immat), dedupe_key(chassy)
    mask = pd.Series(False, index=df.index)
    if ki:
        mask |= df["IMMAT"].map(dedupe_key) == ki
    if kc:
        mask |= df["CHASSY"].map(dedupe_key) == kc
    if exclude is not None:
        mask &= df.index != exclude
    return df[mask]


def is_valid_email_basic(value: str) -> bool:
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", (value or "").strip()))


def norm_statut(v) -> str:
    s = clean_text(v).upper()
    return STATUT_ANNULE if ("ANNUL" in s or "PERDU" in s) else STATUT_ACTIF


def date_col(step: str) -> str:
    return f"DATE_{step}"


def fdate(ts) -> str:
    return "" if pd.isna(ts) else pd.Timestamp(ts).strftime("%d/%m/%Y")


def normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie le tableau : en-têtes, texte, OUI/NON -> booléens, dates."""
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    text_cols = INFO_COLS + TEXT_EXTRA
    for c in text_cols:
        if c not in df.columns:
            df[c] = ""
    for c in STEPS:
        if c not in df.columns:
            df[c] = False
    for c in DATE_COLS:
        if c not in df.columns:
            df[c] = pd.NaT
    known = INFO_COLS + STEPS + TEXT_EXTRA + DATE_COLS
    extras = [c for c in df.columns if c not in known]
    df = df[known + extras]
    for c in text_cols:
        df[c] = df[c].map(clean_text).astype(object)
    df["TELEPHONE"] = df["TELEPHONE"].map(norm_phone).astype(object)
    df["STATUT"] = df["STATUT"].map(norm_statut).astype(object)
    for c in STEPS:
        df[c] = df[c].map(to_bool).astype(bool)
    for c in DATE_COLS:
        df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True).dt.normalize()
    return df.reset_index(drop=True)


def to_excel_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Reconvertit les booléens en OUI/NON et les dates en vraies dates pour Excel."""
    out = df.copy()
    for c in STEPS:
        out[c] = out[c].map(lambda b: "OUI" if b else "NON")
    for c in DATE_COLS:
        out[c] = out[c].dt.date
    return out


def write_excel(df: pd.DataFrame, target) -> None:
    out = to_excel_frame(df)
    with pd.ExcelWriter(target, engine="openpyxl") as w:
        out.to_excel(w, index=False)
        ws = w.book.worksheets[0]
        # dates affichées JJ/MM/AAAA dans Excel (pandas ignore ce réglage, on le pose ici)
        for j, name in enumerate(out.columns, start=1):
            if name in DATE_COLS:
                ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = 14
                for col in ws.iter_cols(min_col=j, max_col=j, min_row=2):
                    for cell in col:
                        cell.number_format = "DD/MM/YYYY"


def save_to_excel(df: pd.DataFrame, path: str) -> None:
    p = Path(path)
    if p.exists():
        shutil.copy(p, p.with_suffix(".bak.xlsx"))  # petite sauvegarde de sécurité
    write_excel(df, p)


def excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    write_excel(df, buf)
    return buf.getvalue()


def next_step(row) -> str:
    for s in STEPS:
        if not row[s]:
            return LABELS[s]
    return DONE


def add_computed(df: pd.DataFrame) -> pd.DataFrame:
    """Ajoute progression, prochaine étape, jours sans avancée et relance à faire."""
    view = df.copy()
    view["PROGRESSION"] = (view[STEPS].sum(axis=1) / len(STEPS) * 100).round()
    view["PROCHAINE ÉTAPE"] = [next_step(r) for _, r in view.iterrows()]
    actif = view["STATUT"] != STATUT_ANNULE
    ref = pd.Timestamp(today_fr())
    if view.empty:
        view["JOURS_BLOQUE"] = pd.Series(dtype=float)
        view["RELANCE_DUE"] = pd.Series(dtype=bool)
        return view
    last_progress = view[STEP_DATE_COLS].max(axis=1)  # date de la dernière étape cochée
    ouvert = actif & (view["PROCHAINE ÉTAPE"] != DONE)
    view["JOURS_BLOQUE"] = (ref - last_progress).dt.days.astype(float).where(ouvert)
    view["RELANCE_DUE"] = actif & view["RELANCE_LE"].notna() & (view["RELANCE_LE"] <= ref)
    return view


def set_step(dframe: pd.DataFrame, idx, step: str, done: bool, when=None) -> None:
    """Coche/décoche une étape et tient sa date à jour."""
    dframe.at[idx, step] = bool(done)
    dframe.at[idx, date_col(step)] = pd.Timestamp(when or today_fr()) if done else pd.NaT


def frames_equal(a: pd.DataFrame, b: pd.DataFrame) -> bool:
    """Compare deux tableaux sur les valeurs (indifférent au type str/object)."""
    if not a.index.equals(b.index) or a.shape != b.shape:
        return False
    return bool((a.astype(object).to_numpy() == b.astype(object).to_numpy()).all())


def rowkey(r, cols) -> int:
    """Signature d'une ligne : sert à recharger les champs si les données changent ailleurs."""
    return abs(hash(tuple(str(r[c]) for c in cols))) % 10**8


def client_label(r) -> str:
    label = f"{r['NOM']} {r['PRENOM']} — {r['IMMAT'] or 'sans immat'}"
    if r["CHASSY"]:
        label += f" — {r['CHASSY']}"
    if r["STATUT"] == STATUT_ANNULE:
        label += " [annulé]"
    return label


def optional_date(label: str, current, default: date, key: str):
    """Case à cocher + date : retourne un Timestamp, ou NaT si la case est décochée."""
    on = st.checkbox(label, value=bool(pd.notna(current)), key=f"{key}_on")
    shown = pd.Timestamp(current).date() if pd.notna(current) else default
    d = st.date_input(label, value=shown, key=key, format=DATE_FMT,
                      disabled=not on, label_visibility="collapsed")
    return pd.Timestamp(d) if (on and d is not None) else pd.NaT


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
    st.header("Alertes")
    stale_days = int(st.number_input(
        "Dossier « bloqué » après (jours)", min_value=1, max_value=60,
        value=STALE_DAYS_DEFAULT,
        help="Nombre de jours sans nouvelle étape cochée avant d'alerter.",
    ))

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
            st.session_state.df = normalize(pd.DataFrame(columns=CORE_COLS + STEPS))
            st.rerun()
        st.stop()

df: pd.DataFrame = st.session_state.df
computed = add_computed(df)
actif_mask = computed["STATUT"] != STATUT_ANNULE


def persist():
    if autosave:
        try:
            save_to_excel(st.session_state.df, path)
        except Exception as e:
            st.error(f"Sauvegarde impossible (fichier ouvert dans Excel ?) : {e}")


n_relances = int(computed["RELANCE_DUE"].sum())
if n_relances:
    st.info(f"📞 {n_relances} relance(s) à faire aujourd'hui ou en retard — "
            "onglet Suivi, filtre « À relancer ».")


# ----------------------------------------------------------------------------
# Onglets
# ----------------------------------------------------------------------------
tab_suivi, tab_fiche, tab_new, tab_mail, tab_dash, tab_gestion = st.tabs(
    ["📋 Suivi", "👤 Fiche client", "➕ Nouveau client", "✉️ Mail",
     "📊 Tableau de bord", "🗑️ Supprimer"]
)

# ---------------------------- Suivi -----------------------------------------
with tab_suivi:
    for msg in st.session_state.pop("toasts", []):
        st.warning(msg)

    c1, c2, c3 = st.columns([2, 2, 2])
    search = c1.text_input("🔎 Recherche (nom, prénom, châssis, immat, mail, tél.)", "")
    steps_filter = c2.multiselect(
        "Prochaine étape",
        options=[LABELS[s] for s in STEPS] + [DONE],
        placeholder="Toutes",
    )
    statut_view = c3.radio("Afficher", ["En cours", "Annulés / perdus", "Tous"], horizontal=True)

    t1, t2, t3 = st.columns(3)
    hide_done = t1.toggle("Masquer les terminés", value=False)
    only_blocked = t2.toggle(f"Bloqués depuis plus de {stale_days} jours", value=False)
    only_relance = t3.toggle("À relancer (aujourd'hui ou en retard)", value=False)

    mask = pd.Series(True, index=computed.index)
    if statut_view == "En cours":
        mask &= actif_mask
    elif statut_view == "Annulés / perdus":
        mask &= ~actif_mask
    if search.strip():
        q = search.strip().lower()
        mask &= computed[INFO_COLS].apply(
            lambda col: col.str.lower().str.contains(q, regex=False)
        ).any(axis=1)
    if steps_filter:
        mask &= computed["PROCHAINE ÉTAPE"].isin(steps_filter)
    if hide_done:
        mask &= computed["PROCHAINE ÉTAPE"] != DONE
    if only_blocked:
        mask &= computed["JOURS_BLOQUE"] > stale_days
    if only_relance:
        mask &= computed["RELANCE_DUE"]

    view = computed[mask]
    st.caption(f"{len(view)} dossier(s) affiché(s) sur {len(computed)}")

    edit_cols = CORE_COLS + STEPS
    computed_cols = ["PROGRESSION", "PROCHAINE ÉTAPE", "JOURS_BLOQUE", "RELANCE_LE"]
    column_config = {
        "NOM": st.column_config.TextColumn("Nom"),
        "PRENOM": st.column_config.TextColumn("Prénom"),
        "CHASSY": st.column_config.TextColumn("Châssis"),
        "IMMAT": st.column_config.TextColumn("Immat"),
        "PROGRESSION": st.column_config.ProgressColumn(
            "Progression", min_value=0, max_value=100, format="%d%%"
        ),
        "PROCHAINE ÉTAPE": st.column_config.TextColumn("Prochaine étape"),
        "JOURS_BLOQUE": st.column_config.NumberColumn("Sans avancée (j)", format="%d"),
        "RELANCE_LE": st.column_config.DateColumn("À rappeler le", format=DATE_FMT),
    }
    for s in STEPS:
        column_config[s] = st.column_config.CheckboxColumn(LABELS[s])

    ordered = CORE_COLS + computed_cols + STEPS
    edited = st.data_editor(
        view[ordered],
        column_config=column_config,
        disabled=computed_cols,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        height=min(600, 80 + 35 * max(len(view), 1)),
        # la clé dépend des lignes affichées : changer de filtre ne réapplique
        # pas des modifications à la mauvaise ligne
        key=f"editor_{abs(hash(tuple(view.index))) % 10**8}",
    )

    cur = edited[edit_cols].copy()
    cur[CORE_COLS] = cur[CORE_COLS].fillna("")
    old = view[edit_cols]
    if not frames_equal(cur, old):
        new = cur.copy()
        msgs = []
        for i in new.index:
            new.at[i, "NOM"] = clean_text(new.at[i, "NOM"])
            new.at[i, "PRENOM"] = clean_text(new.at[i, "PRENOM"])
            label = f"{new.at[i, 'NOM']} {new.at[i, 'PRENOM']}".strip()
            for col, fn, nom_col in (("IMMAT", norm_immat, "L'immat"),
                                     ("CHASSY", norm_chassis, "Le châssis")):
                if new.at[i, col] != old.at[i, col]:
                    val, warn = fn(new.at[i, col])
                    new.at[i, col] = val
                    if warn:
                        msgs.append(f"{label} : {warn}")
                    dup = find_duplicates(
                        df, val if col == "IMMAT" else "", val if col == "CHASSY" else "", exclude=i
                    )
                    if not dup.empty:
                        d = dup.iloc[0]
                        msgs.append(f"{label} : {nom_col} est déjà utilisé par {d['NOM']} {d['PRENOM']}.")
        if not frames_equal(new, old):  # (évite une boucle si la correction ne change rien)
            for i in new.index:
                for s in STEPS:
                    if bool(new.at[i, s]) != bool(old.at[i, s]):
                        set_step(st.session_state.df, i, s, bool(new.at[i, s]))
            st.session_state.df.loc[new.index, CORE_COLS] = new[CORE_COLS]
            persist()
            if msgs:
                st.session_state["toasts"] = msgs
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

# ---------------------------- Fiche client ----------------------------------
with tab_fiche:
    st.subheader("Fiche client")

    flash_fiche = st.session_state.pop("flash_fiche", None)
    if flash_fiche:
        st.success(flash_fiche)

    if df.empty:
        st.info("Aucun client dans le tableau pour le moment.")
    else:
        fiche_labels = {i: client_label(r) for i, r in df.iterrows()}
        fidx = st.selectbox(
            "Client",
            options=list(fiche_labels),
            format_func=fiche_labels.get,
            index=None,
            placeholder="Tape un nom, une immat ou un châssis…",
            key="fiche_client",
        )

        if fidx is not None:
            r = df.loc[fidx]
            c = computed.loc[fidx]
            annule = r["STATUT"] == STATUT_ANNULE
            today = pd.Timestamp(today_fr())

            # ---- En-tête ----
            st.markdown(f"### {r['NOM']} {r['PRENOM']}")
            coords = " · ".join(x for x in [r["TELEPHONE"], r["EMAIL"]] if x)
            if coords:
                st.caption(coords)
            pct = int(c["PROGRESSION"])
            st.progress(pct, text=f"{pct} % — prochaine étape : {c['PROCHAINE ÉTAPE']}")
            if annule:
                motif_txt = f" — {r['MOTIF_ANNULATION']}" if r["MOTIF_ANNULATION"] else ""
                st.error(f"🚫 Dossier annulé / perdu{motif_txt}")
            else:
                if pd.notna(c["JOURS_BLOQUE"]) and c["JOURS_BLOQUE"] > stale_days:
                    st.warning(f"⏳ Aucune avancée depuis {int(c['JOURS_BLOQUE'])} jours.")
                if bool(c["RELANCE_DUE"]):
                    st.warning(f"📞 Relance prévue le {fdate(r['RELANCE_LE'])}.")
            infos = []
            if pd.notna(r["DERNIER_CONTACT"]):
                infos.append(f"Dernier contact : {fdate(r['DERNIER_CONTACT'])}")
            if pd.notna(r["RELANCE_LE"]) and not bool(c["RELANCE_DUE"]):
                infos.append(f"Relance prévue le {fdate(r['RELANCE_LE'])}")
            if infos:
                st.caption(" · ".join(infos))

            # ---- Actions rapides pendant l'appel ----
            q1, q2, q3, q4 = st.columns(4)
            quick = None
            if q1.button("📞 Contacté aujourd'hui", key=f"q1_{fidx}", use_container_width=True):
                quick = ("contact", None)
            if q2.button("⏰ Rappeler demain", key=f"q2_{fidx}", use_container_width=True):
                quick = ("later", 1)
            if q3.button("⏰ Rappeler dans 3 jours", key=f"q3_{fidx}", use_container_width=True):
                quick = ("later", 3)
            if q4.button("⏰ Rappeler dans 1 semaine", key=f"q4_{fidx}", use_container_width=True):
                quick = ("later", 7)
            if quick:
                d2 = st.session_state.df
                d2.at[fidx, "DERNIER_CONTACT"] = today
                if quick[0] == "later":
                    when = today + pd.Timedelta(days=quick[1])
                    d2.at[fidx, "RELANCE_LE"] = when
                    msg_q = f"✅ Contact noté, relance fixée au {fdate(when)}."
                else:
                    if pd.notna(r["RELANCE_LE"]) and r["RELANCE_LE"] <= today:
                        d2.at[fidx, "RELANCE_LE"] = pd.NaT  # relance faite
                    msg_q = "✅ Contact du jour noté."
                persist()
                st.session_state["flash_fiche"] = msg_q
                st.rerun()

            left, right = st.columns(2)

            # ---- Étapes (enregistrées dès que tu coches) ----
            with left:
                st.markdown("**Étapes** (enregistrées dès que tu coches)")
                for s in STEPS:
                    done = bool(r[s])
                    d_cur = r[date_col(s)]
                    a, b = st.columns([3, 2])
                    sk = f"{fidx}_{s}_{done}_{fdate(d_cur)}"
                    new_done = a.checkbox(LABELS[s], value=done, key=f"chk_{sk}")
                    if done and pd.isna(d_cur):
                        b.caption("date inconnue")
                        if b.button("📅 Dater aujourd'hui", key=f"dt_{sk}"):
                            st.session_state.df.at[fidx, date_col(s)] = today
                            persist()
                            st.rerun()
                        new_date = None
                        shown = None
                    else:
                        shown = d_cur.date() if pd.notna(d_cur) else today.date()
                        new_date = b.date_input(
                            LABELS[s], value=shown, key=f"dt_{sk}", format=DATE_FMT,
                            disabled=not done, label_visibility="collapsed",
                        )
                    if new_done != done:
                        set_step(st.session_state.df, fidx, s, new_done)
                        persist()
                        st.rerun()
                    elif done and new_date is not None and new_date != shown:
                        st.session_state.df.at[fidx, date_col(s)] = pd.Timestamp(new_date)
                        persist()
                        st.rerun()

            # ---- Informations, notes, relance ----
            with right:
                st.markdown("**Informations**")
                sig = rowkey(r, INFO_COLS + ["NOTES"])

                def k(name):
                    return f"fiche_{name}_{fidx}_{sig}"

                a, b = st.columns(2)
                f_nom = a.text_input("Nom", value=r["NOM"], key=k("nom"))
                f_prenom = b.text_input("Prénom", value=r["PRENOM"], key=k("prenom"))
                f_immat = a.text_input("Immat", value=r["IMMAT"], key=k("immat"))
                f_chassy = b.text_input("Châssis", value=r["CHASSY"], key=k("chassy"))
                f_tel = a.text_input("Téléphone", value=r["TELEPHONE"], key=k("tel"))
                f_email = b.text_input("Email", value=r["EMAIL"], key=k("email"))
                f_notes = st.text_area("Notes", value=r["NOTES"], height=150, key=k("notes"))

                st.markdown("**Suivi**")
                f_rel = optional_date("Relance prévue le", r["RELANCE_LE"],
                                      (today + pd.Timedelta(days=2)).date(),
                                      f"fiche_rel_{fidx}_{rowkey(r, ['RELANCE_LE'])}")
                f_lc = optional_date("Dernier contact le", r["DERNIER_CONTACT"], today.date(),
                                     f"fiche_lc_{fidx}_{rowkey(r, ['DERNIER_CONTACT'])}")

                ignore = st.checkbox(
                    "Enregistrer même si le format de l'immat ou du châssis est inhabituel",
                    key=k("ign"),
                )
                if st.button("💾 Enregistrer la fiche", type="primary", key=k("save")):
                    errors, warns = [], []
                    nom_v, prenom_v = f_nom.strip(), f_prenom.strip()
                    if not (nom_v and prenom_v):
                        errors.append("Le nom et le prénom sont obligatoires.")
                    immat_v, w = norm_immat(f_immat)
                    if w:
                        warns.append(w)
                    chassy_v, w = norm_chassis(f_chassy)
                    if w:
                        warns.append(w)
                    email_v = f_email.strip()
                    if email_v and not is_valid_email_basic(email_v):
                        errors.append("L'adresse mail ne semble pas valide.")
                    dup = find_duplicates(df, immat_v, chassy_v, exclude=fidx)
                    if not dup.empty:
                        d = dup.iloc[0]
                        errors.append(f"Cette immat ou ce châssis est déjà utilisé par {d['NOM']} {d['PRENOM']}.")
                    if warns and not ignore:
                        errors += warns + ["Corrige, ou coche « Enregistrer même si le format… » puis réessaie."]
                    if errors:
                        for e in errors:
                            st.error(e)
                    else:
                        d2 = st.session_state.df
                        d2.at[fidx, "NOM"] = nom_v
                        d2.at[fidx, "PRENOM"] = prenom_v
                        d2.at[fidx, "IMMAT"] = immat_v
                        d2.at[fidx, "CHASSY"] = chassy_v
                        d2.at[fidx, "EMAIL"] = email_v
                        d2.at[fidx, "TELEPHONE"] = norm_phone(f_tel)
                        d2.at[fidx, "NOTES"] = f_notes.strip()
                        d2.at[fidx, "RELANCE_LE"] = f_rel
                        d2.at[fidx, "DERNIER_CONTACT"] = f_lc
                        persist()
                        st.session_state["flash_fiche"] = "✅ Fiche enregistrée."
                        st.rerun()

            # ---- Annulation / réactivation ----
            st.divider()
            if annule:
                if st.button("♻️ Réactiver ce dossier", key=f"react_{fidx}"):
                    st.session_state.df.at[fidx, "STATUT"] = STATUT_ACTIF
                    st.session_state.df.at[fidx, "MOTIF_ANNULATION"] = ""
                    persist()
                    st.session_state["flash_fiche"] = "♻️ Dossier réactivé."
                    st.rerun()
            else:
                with st.expander("🚫 Annuler / marquer ce dossier comme perdu"):
                    motif = st.text_input("Motif (facultatif)", key=f"motif_{fidx}")
                    if st.button("Confirmer l'annulation", key=f"cancel_{fidx}"):
                        st.session_state.df.at[fidx, "STATUT"] = STATUT_ANNULE
                        st.session_state.df.at[fidx, "MOTIF_ANNULATION"] = motif.strip()
                        persist()
                        st.session_state["flash_fiche"] = "🚫 Dossier marqué comme annulé / perdu."
                        st.rerun()

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
        chassy = a.text_input("Châssis (17 caractères)")
        immat = b.text_input("Immat (ex. AB-123-CD)")
        email = a.text_input("Adresse mail du client")
        tel = b.text_input("Téléphone")

        st.markdown("**Étapes déjà réalisées** (laisse décoché ce qui reste à faire)")
        step_cols = st.columns(3)
        new_steps = {}
        for i, s in enumerate(STEPS):
            new_steps[s] = step_cols[i % 3].checkbox(LABELS[s], key=f"new_{s}")

        ignore_new = st.checkbox("Ajouter même si le format de l'immat ou du châssis est inhabituel")
        submitted = st.form_submit_button("➕ Ajouter au tableau", type="primary")

    if submitted:
        nom, prenom, email = nom.strip(), prenom.strip(), email.strip()
        immat_v, w1 = norm_immat(immat)
        chassy_v, w2 = norm_chassis(chassy)
        warns = [w for w in (w1, w2) if w]
        dup = find_duplicates(df, immat_v, chassy_v)
        if not (nom and prenom):
            st.warning("Le nom et le prénom sont obligatoires.")
        elif email and not is_valid_email_basic(email):
            st.warning("L'adresse mail ne semble pas valide. Rien n'a été ajouté.")
        elif not dup.empty:
            d = dup.iloc[0]
            st.warning(
                f"Un client avec cette immat ou ce châssis existe déjà : "
                f"{d['NOM']} {d['PRENOM']}. Rien n'a été ajouté."
            )
        elif warns and not ignore_new:
            for w in warns:
                st.warning(w)
            st.warning("Rien n'a été ajouté. Corrige la saisie, ou coche « Ajouter même si le "
                       "format… » et réessaie.")
        else:
            row = {c: "" for c in df.columns}
            for c in DATE_COLS:
                row[c] = pd.NaT
            row.update(NOM=nom, PRENOM=prenom, CHASSY=chassy_v, IMMAT=immat_v,
                       EMAIL=email, TELEPHONE=norm_phone(tel), STATUT=STATUT_ACTIF)
            for s in STEPS:
                row[s] = bool(new_steps[s])
                row[date_col(s)] = pd.Timestamp(today_fr()) if new_steps[s] else pd.NaT
            st.session_state.df = normalize(
                pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            )
            persist()
            st.session_state["flash"] = f"✅ {nom} {prenom} a été ajouté au tableau."
            st.rerun()

# ---------------------------- Mail -------------------------------------------
with tab_mail:
    st.subheader("Préparer le mail « documents à retourner »")

    flash_mail = st.session_state.pop("flash_mail", None)
    if flash_mail:
        st.success(flash_mail)

    if df.empty:
        st.info("Aucun client dans le tableau pour le moment.")
    else:
        mail_labels = {i: client_label(r) for i, r in df.iterrows()}
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

            d1, d2c = st.columns(2)
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
            if d2c.button("✅ Marquer « BDC envoyé » pour ce client", use_container_width=True,
                          disabled=bool(r["BDC_ENVOYE"])):
                set_step(st.session_state.df, idx, "BDC_ENVOYE", True)
                persist()
                st.session_state["flash_mail"] = f"✅ BDC envoyé coché pour {r['NOM']} {r['PRENOM']}."
                st.rerun()
            if r["BDC_ENVOYE"]:
                st.caption("« BDC envoyé » est déjà coché pour ce client.")

            with st.expander("Le brouillon ne s'ouvre pas ? Copier le texte à la main"):
                st.code(body, language=None)

# ---------------------------- Tableau de bord --------------------------------
with tab_dash:
    actifs = computed[actif_mask]
    total = len(actifs)
    termines = int((actifs["PROCHAINE ÉTAPE"] == DONE).sum())
    bloques = int((actifs["JOURS_BLOQUE"] > stale_days).sum())

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Dossiers en cours", total)
    m2.metric("Terminés", termines)
    m3.metric(f"Bloqués > {stale_days} j", bloques)
    m4.metric("Relances à faire", n_relances)
    m5.metric("Progression moyenne", f"{actifs['PROGRESSION'].mean():.0f}%" if total else "—")
    st.caption(f"{int((~actif_mask).sum())} dossier(s) annulé(s) / perdu(s), non comptés ci-dessus.")

    st.subheader("Dossiers en attente par étape")
    pending = pd.DataFrame(
        {"Étape": [LABELS[s] for s in STEPS],
         "À faire": [int((~df.loc[actif_mask, s]).sum()) for s in STEPS]}
    ).set_index("Étape")
    st.bar_chart(pending)

    st.subheader("Cas à surveiller")
    show_cols = CORE_COLS
    alerts = {
        f"Bloqués depuis plus de {stale_days} jours": (
            computed[computed["JOURS_BLOQUE"] > stale_days],
            ["NOM", "PRENOM", "IMMAT", "PROCHAINE ÉTAPE", "JOURS_BLOQUE"],
        ),
        "Relances à faire": (
            computed[computed["RELANCE_DUE"]],
            ["NOM", "PRENOM", "IMMAT", "TELEPHONE", "RELANCE_LE", "NOTES"],
        ),
        "BDC signé mais pas de paiement": (
            computed[df["BDC_SIGNE"] & ~df["PAIEMENT"] & actif_mask], show_cols),
        "Payé mais pas de facture": (
            computed[df["PAIEMENT"] & ~df["FACTURE_DCSNT"] & actif_mask], show_cols),
        "Facturé mais pas de carte grise": (
            computed[df["FACTURE_DCSNT"] & ~df["CARTE_GRISE"] & actif_mask], show_cols),
        "BDC envoyé mais pas signé": (
            computed[df["BDC_ENVOYE"] & ~df["BDC_SIGNE"] & actif_mask], show_cols),
    }
    for title, (sub, cols) in alerts.items():
        with st.expander(f"{title} ({len(sub)})"):
            if len(sub):
                st.dataframe(
                    sub[cols],
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        "JOURS_BLOQUE": st.column_config.NumberColumn("Sans avancée (j)", format="%d"),
                        "RELANCE_LE": st.column_config.DateColumn("À rappeler le", format=DATE_FMT),
                    },
                )
            else:
                st.write("Aucun dossier concerné.")

# ---------------------------- Supprimer --------------------------------------
with tab_gestion:
    st.subheader("Supprimer des clients")
    st.caption("Pour un dossier perdu, préfère « Annuler » dans la fiche client : "
               "il sort de ta liste mais reste dans l'historique.")
    labels = {i: client_label(r) for i, r in df.iterrows()}
    to_delete = st.multiselect("Clients à supprimer", options=list(labels), format_func=labels.get)
    confirm = st.checkbox("Je confirme la suppression")
    if st.button("🗑️ Supprimer", disabled=not (to_delete and confirm)):
        st.session_state.df = df.drop(index=to_delete).reset_index(drop=True)
        persist()
        st.rerun()
