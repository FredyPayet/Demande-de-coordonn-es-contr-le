# -*- coding: utf-8 -*-
"""
Application Streamlit d'automatisation :
  1) Lecture du "tableau 1" (export de contrôle, une ligne = une opération)
  2) Report des colonnes "RAISON SOCIALE du demandeur" -> "SIRET de l'entreprise
     ayant réalisé l'opération" dans les matrices "tableau 2" (une matrice = un
     modèle de fiche CEE, ex. BAR-TH-104, BAR-EN-105, ...)
  3) Génération d'un fichier tableau 2 rempli PAR CLIENT et PAR FICHE, avec
     uniquement les opérations de ce client pour cette fiche
  4) Génération d'un mail type par client dans Outlook (application locale),
     avec les fichiers générés en pièces jointes

Lancement :
    pip install streamlit openpyxl pandas pywin32
    streamlit run app.py

NB pywin32 / Outlook : la génération automatique des mails dans Outlook ne
fonctionne que si l'application tourne SUR le même poste Windows où Outlook
(application de bureau) est installé et ouvert. Sur un serveur / Mac / Linux,
cette étape est remplacée par un export .eml + un récapitulatif à copier-coller.
"""

import io
import os
import re
import zipfile
import unicodedata
from collections import defaultdict
from datetime import datetime

import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Essai d'import pywin32 (uniquement disponible / utile sous Windows)
# ---------------------------------------------------------------------------
try:
    import win32com.client  # type: ignore
    OUTLOOK_AVAILABLE = True
except Exception:
    OUTLOOK_AVAILABLE = False


# ---------------------------------------------------------------------------
# Constantes métier
# ---------------------------------------------------------------------------

# Colonnes à NE JAMAIS remplir dans les matrices, même si un nom de colonne
# source correspondait par erreur.
COLONNES_EXCLUES = [
    "numero de telephone du beneficiaire",
    "adresse de courriel du beneficiaire",
    "montant du role actif et incitatif",
]

# Nom (normalisé) de la colonne source qui contient le code de la fiche
# (ex : BAR-TH-104) et qui sert à choisir la bonne matrice.
COLONNE_CODE_FICHE = "reference de la fiche d'operation standardisee"

# Séparateur utilisé lors de la concaténation de plusieurs colonnes sources
# (ex: "Marque isolant" + "Référence isolant")
SEPARATEUR_CONCATENATION = " "


# Nom (normalisé) de la colonne qui identifie le client (raison sociale du
# bénéficiaire de l'opération). Utilisé pour le regroupement et le mail.
COLONNE_CLIENT = "raison sociale du beneficiaire de l'operation"

FEUILLE_MATRICE = "Personnes morales"  # nom de la feuille dans les matrices


# ---------------------------------------------------------------------------
# Fonctions utilitaires de normalisation / correspondance de colonnes
# ---------------------------------------------------------------------------

def normaliser(texte):
    """Normalise un texte pour comparaison : minuscule, sans accent, sans
    parenthèses, espaces multiples réduits."""
    if texte is None:
        return ""
    s = str(texte)
    s = s.replace("\u2019", "'")
    # retire tout contenu entre parenthèses (ex: "(sous la forme : XXX-XX-XX)")
    s = re.sub(r"\([^)]*\)", " ", s)
    # retire les accents
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = re.sub(r"[\s\n\r]+", " ", s).strip()
    return s


def colonne_exclue(nom_normalise):
    return any(nom_normalise.startswith(exc) for exc in COLONNES_EXCLUES)


def trouver_meilleure_correspondance(nom_cible, index_source):
    """Cherche le nom de colonne source correspondant à nom_cible (déjà
    normalisés). Tolère les suffixes/préfixes différents (ex: 'reference
    interne de l'operation' vs '... du demandeur')."""
    if nom_cible in index_source:
        return nom_cible
    # correspondance par préfixe dans les deux sens
    candidats = []
    for cle in index_source:
        if not cle or not nom_cible:
            continue
        if cle.startswith(nom_cible) or nom_cible.startswith(cle):
            candidats.append(cle)
    if len(candidats) == 1:
        return candidats[0]
    if len(candidats) > 1:
        # on prend le plus proche en longueur
        candidats.sort(key=lambda c: abs(len(c) - len(nom_cible)))
        return candidats[0]
    return None


# ---------------------------------------------------------------------------
# Lecture du tableau 1 (source)
# ---------------------------------------------------------------------------

def lire_tableau_source(fichier):
    """Retourne (liste_de_dicts_operations, index_colonnes_normalisees,
    nom_feuille) depuis le fichier Excel source. Détecte automatiquement la
    ligne d'en-têtes (celle où figure 'RAISON SOCIALE du demandeur' en
    colonne A ou plus loin) et la première ligne de données."""
    wb = openpyxl.load_workbook(fichier, data_only=True)
    ws = wb[wb.sheetnames[0]]

    header_row = None
    for r in range(1, min(10, ws.max_row) + 1):
        for c in range(1, min(80, ws.max_column) + 1):
            v = normaliser(ws.cell(row=r, column=c).value)
            if v == "raison sociale du demandeur":
                header_row = r
                break
        if header_row:
            break
    if header_row is None:
        raise ValueError(
            "Impossible de trouver la ligne d'en-têtes (colonne "
            "'RAISON SOCIALE du demandeur') dans le tableau 1."
        )

    # index colonne_normalisee -> numero de colonne (on garde la 1ere occurrence)
    index_cols = {}
    for c in range(1, ws.max_column + 1):
        v = normaliser(ws.cell(row=header_row, column=c).value)
        if v and v not in index_cols:
            index_cols[v] = c

    operations = []
    for r in range(header_row + 1, ws.max_row + 1):
        # ligne vide -> on l'ignore (mais on continue de scanner jusqu'au bout)
        if all(ws.cell(row=r, column=c).value in (None, "") for c in range(1, ws.max_column + 1)):
            continue
        ligne = {}
        for nom, c in index_cols.items():
            ligne[nom] = ws.cell(row=r, column=c).value
        ligne["_ligne_source"] = r
        operations.append(ligne)

    return operations, index_cols, header_row


# ---------------------------------------------------------------------------
# Détection des codes fiches à partir des noms de fichiers matrices
# ---------------------------------------------------------------------------

_FULL_CODE = re.compile(r"(BAR|BAT)[\s_-](EN|TH)[\s_-](\d{3})(-SE)?", re.IGNORECASE)
_BARE_NUM = re.compile(r"(?<!\d)(\d{3})(?!\d)")


def extraire_codes_fiches(nom_fichier):
    """Devine, à partir du nom du fichier matrice, la ou les codes de fiches
    CEE couverts (ex: 'BAR-TH-145 et 164.xlsx' -> ['BAR-TH-145','BAR-TH-164'])."""
    base = os.path.splitext(os.path.basename(nom_fichier))[0]
    norm = base.replace("_", " ")
    norm = re.sub(r"\bet\b", " ", norm, flags=re.IGNORECASE)
    norm = re.sub(r"\s+", " ", norm)

    tokens = []
    full_spans = []
    for m in _FULL_CODE.finditer(norm):
        tokens.append((m.start(), "full", m))
        full_spans.append((m.start(), m.end()))
    for m in _BARE_NUM.finditer(norm):
        if any(m.start() >= s and m.end() <= e for s, e in full_spans):
            continue
        tokens.append((m.start(), "bare", m))
    tokens.sort(key=lambda t: t[0])

    codes, last_fam = [], None
    for _, typ, m in tokens:
        if typ == "full":
            fam = f"{m.group(1).upper()}-{m.group(2).upper()}"
            se = "-SE" if m.group(4) else ""
            codes.append(f"{fam}-{m.group(3)}{se}")
            last_fam = fam
        elif last_fam:
            codes.append(f"{last_fam}-{m.group(1)}")
    # dédoublonne en conservant l'ordre
    vu, resultat = set(), []
    for c in codes:
        if c not in vu:
            vu.add(c)
            resultat.append(c)
    return resultat


def construire_mapping_fiches(fichiers_matrices):
    """fichiers_matrices : dict {nom_fichier: bytes}. Retourne une liste de
    lignes {code_fiche, fichier_matrice} pour affichage/édition dans l'appli."""
    lignes = []
    for nom in sorted(fichiers_matrices.keys()):
        codes = extraire_codes_fiches(nom)
        if not codes:
            lignes.append({"code_fiche": "(non détecté)", "fichier_matrice": nom})
        for code in codes:
            lignes.append({"code_fiche": code, "fichier_matrice": nom})
    return lignes


# ---------------------------------------------------------------------------
# Lecture des en-têtes d'une matrice (tableau 2) + détection ligne d'en-tête
# ---------------------------------------------------------------------------

def lire_entetes_matrice(contenu_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(contenu_bytes), data_only=False)
    nom_feuille = wb.sheetnames[0]
    for candidat in wb.sheetnames:
        if normaliser(candidat) == normaliser(FEUILLE_MATRICE):
            nom_feuille = candidat
            break
    ws = wb[nom_feuille]

    header_row = None
    for r in range(1, min(10, ws.max_row) + 1):
        v = normaliser(ws.cell(row=r, column=1).value)
        if v == "raison sociale du demandeur":
            header_row = r
            break
    if header_row is None:
        return None, None, None

    entetes = {}
    for c in range(1, ws.max_column + 1):
        v = normaliser(ws.cell(row=header_row, column=c).value)
        if v:
            entetes[c] = v
    return nom_feuille, header_row, entetes


# ---------------------------------------------------------------------------
# Lecture du tableau de correspondance "données techniques" (dépend de la
# fiche CEE, ex: BAR-EN-103 -> colonnes isolation à remplir depuis le tableau 1)
# ---------------------------------------------------------------------------

def lire_tableau_technique(fichier):
    """Lit le fichier de correspondance données techniques.
    Colonnes attendues : Fiche CEE | Valeur à remplir sur le tableau 2 |
    équivalent à récupérer sur le tableau 1.
    La colonne 'Fiche CEE' n'est renseignée que sur la 1ère ligne de chaque
    groupe (fusion visuelle) : on la propage sur les lignes suivantes.
    Retourne une liste de dicts : {code_fiche, colonne_cible, expression}."""
    wb = openpyxl.load_workbook(fichier, data_only=True)
    ws = wb[wb.sheetnames[0]]

    header_row = 1
    for r in range(1, min(5, ws.max_row) + 1):
        v = normaliser(ws.cell(row=r, column=1).value)
        if "fiche" in v:
            header_row = r
            break

    lignes = []
    dernier_code = None
    for r in range(header_row + 1, ws.max_row + 1):
        code = ws.cell(row=r, column=1).value
        cible = ws.cell(row=r, column=2).value
        expression = ws.cell(row=r, column=3).value
        if code:
            dernier_code = str(code).strip().upper()
        if not cible or not expression:
            continue
        lignes.append(
            {
                "code_fiche": dernier_code or "",
                "colonne_cible": str(cible).strip(),
                "expression_source": str(expression).strip(),
            }
        )
    return lignes


def resoudre_expression_technique(expression, index_cols_source):
    """Détermine si 'expression' (ex: 'Marque isolant + Référence isolant' ou
    'Surface isolant (m²)' ou '3.85') correspond à une ou plusieurs colonnes
    du tableau 1, ou s'il s'agit d'une valeur constante à écrire telle quelle.
    Retourne ('colonnes', [noms_colonnes_source]) ou ('constante', valeur)."""
    parties = [p.strip() for p in expression.split("+") if p.strip()]
    if not parties:
        return ("constante", expression)

    colonnes_resolues = []
    for partie in parties:
        cible = trouver_meilleure_correspondance(normaliser(partie), index_cols_source)
        if not cible:
            # au moins une partie ne correspond à aucune colonne source
            # -> on considère l'expression entière comme une valeur fixe
            return ("constante", expression)
        colonnes_resolues.append(cible)
    return ("colonnes", colonnes_resolues)


def valeur_constante_typee(texte):
    """Convertit '3.85' ou '3,85' en float si possible, sinon renvoie le texte."""
    try:
        return float(texte.replace(",", "."))
    except (ValueError, AttributeError):
        return texte


def construire_regles_techniques(lignes_techniques, index_cols_source):
    """Regroupe les règles techniques par code fiche, avec leur résolution
    (colonnes sources à concaténer, ou valeur constante), prêtes à l'emploi."""
    regles_par_fiche = defaultdict(list)
    for ligne in lignes_techniques:
        type_resolution, donnee = resoudre_expression_technique(
            ligne["expression_source"], index_cols_source
        )
        regles_par_fiche[ligne["code_fiche"]].append(
            {
                "colonne_cible_normalisee": normaliser(ligne["colonne_cible"]),
                "colonne_cible_affichage": ligne["colonne_cible"],
                "type": type_resolution,
                "donnee": donnee,
            }
        )
    return regles_par_fiche


def valeur_regle_technique(operation, regle):
    if regle["type"] == "constante":
        return valeur_constante_typee(regle["donnee"])
    # concaténation d'une ou plusieurs colonnes sources
    parties = []
    for nom_col in regle["donnee"]:
        v = operation.get(nom_col)
        if v not in (None, ""):
            parties.append(str(v))
    if not parties:
        return None
    if len(parties) == 1:
        return parties[0]
    return SEPARATEUR_CONCATENATION.join(parties)


# ---------------------------------------------------------------------------
# Génération des matrices remplies (une par client x fiche)
# ---------------------------------------------------------------------------

def remplir_matrice(contenu_bytes, operations_client, index_cols_source, regles_techniques=None):
    """Copie la matrice modèle (contenu_bytes) et y ajoute une ligne par
    opération. 'regles_techniques' (optionnel) est la liste des règles
    supplémentaires à appliquer pour LA fiche de ces opérations (colonnes
    techniques propres à la fiche, ex: isolation).
    Retourne (bytes_du_fichier, colonnes_non_mappees, cibles_techniques_non_trouvees)."""
    wb = openpyxl.load_workbook(io.BytesIO(contenu_bytes))
    nom_feuille, header_row, entetes = lire_entetes_matrice(contenu_bytes)
    if nom_feuille is None:
        raise ValueError("Ligne d'en-têtes introuvable dans cette matrice.")
    ws = wb[nom_feuille]

    # Au-delà de la colonne SIRET (+ quelques colonnes voisines : téléphone,
    # email, montant... explicitement exclues), les matrices contiennent la
    # partie "audit terrain" du bureau de contrôle, qui n'a pas d'équivalent
    # dans le tableau 1. On ne signale donc les colonnes non reconnues que
    # jusqu'à cette limite, pour ne pas polluer les avertissements.
    col_siret = None
    for c, nom in entetes.items():
        if nom.startswith("siret de l'entreprise ayant realise l'operation"):
            col_siret = c
            break
    limite_signalement = (col_siret + 8) if col_siret else max(entetes.keys(), default=0)

    # correspondance colonne matrice -> colonne source (nom normalisé)
    correspondance = {}  # col_matrice -> nom_colonne_source
    colonnes_non_mappees = []
    for col_matrice, nom_entete in entetes.items():
        if colonne_exclue(nom_entete):
            continue
        cible = trouver_meilleure_correspondance(nom_entete, index_cols_source)
        if cible:
            correspondance[col_matrice] = cible
        elif col_matrice <= limite_signalement:
            colonnes_non_mappees.append(nom_entete)

    # règles techniques spécifiques à la fiche (colonnes situées plus loin
    # dans la matrice, ex: "Surface déclarée dans l'AH/facture (m2)")
    regles_par_colonne = {}  # col_matrice -> regle
    cibles_techniques_non_trouvees = []
    for regle in (regles_techniques or []):
        col_matrice = None
        cible_norm = regle["colonne_cible_normalisee"]
        # correspondance exacte d'abord, puis par préfixe (dans les deux sens)
        for c, nom_entete in entetes.items():
            if nom_entete == cible_norm:
                col_matrice = c
                break
        if col_matrice is None:
            for c, nom_entete in entetes.items():
                if nom_entete and (nom_entete.startswith(cible_norm) or cible_norm.startswith(nom_entete)):
                    col_matrice = c
                    break
        if col_matrice:
            regles_par_colonne[col_matrice] = regle
        else:
            cibles_techniques_non_trouvees.append(regle["colonne_cible_affichage"])

    ligne_ecriture = header_row + 1
    for operation in operations_client:
        for col_matrice, nom_source in correspondance.items():
            valeur = operation.get(nom_source)
            cellule = ws.cell(row=ligne_ecriture, column=col_matrice, value=valeur)
            # Certains modèles ont des cellules pré-formatées en date/monnaie
            # sur ces colonnes (héritage du modèle Excel). On repasse en
            # format standard pour éviter qu'un nombre (ex. SIREN) ne
            # s'affiche comme une date.
            if not isinstance(valeur, datetime):
                cellule.number_format = "General"
        for col_matrice, regle in regles_par_colonne.items():
            valeur = valeur_regle_technique(operation, regle)
            cellule = ws.cell(row=ligne_ecriture, column=col_matrice, value=valeur)
            if not isinstance(valeur, datetime):
                cellule.number_format = "General"
        ligne_ecriture += 1

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.read(), colonnes_non_mappees, cibles_techniques_non_trouvees


# ---------------------------------------------------------------------------
# Interface Streamlit
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Automatisation fiches CEE par client", layout="wide")
st.title("Automatisation — Tableau 1 → Fiches par client (Tableau 2)")

st.markdown(
    """
Cette application :
1. lit le **tableau 1** (export de contrôle) ;
2. reporte les colonnes *"RAISON SOCIALE du demandeur"* → *"SIRET de l'entreprise
   ayant réalisé l'opération"* dans la bonne **matrice** (tableau 2) selon le code
   de fiche (ex. BAR-TH-104) ;
3. reporte aussi les **colonnes techniques propres à chaque fiche** (ex.
   surface, épaisseur, marque d'isolant...) selon un tableau de correspondance
   que vous fournissez ;
4. **regroupe les opérations par client** et génère un fichier par client et par
   type de fiche ;
5. prépare un **mail type par client** (Outlook local si disponible).
"""
)

if "resultats" not in st.session_state:
    st.session_state["resultats"] = None

# --- Étape 1 : fichiers -----------------------------------------------------
st.header("1. Fichiers")

col1, col2 = st.columns(2)
with col1:
    fichier_source = st.file_uploader(
        "Tableau 1 — export de contrôle (.xlsx)", type=["xlsx"], key="source"
    )
with col2:
    fichiers_matrices_up = st.file_uploader(
        "Matrices (tableau 2) — fichiers .xlsx ou une archive .zip les contenant",
        type=["xlsx", "zip"],
        accept_multiple_files=True,
        key="matrices",
    )

fichier_technique = st.file_uploader(
    "Tableau de correspondance des données techniques par fiche (.xlsx) — optionnel",
    type=["xlsx"],
    key="technique",
    help=(
        "Colonnes attendues : 'Fiche CEE tableau 1' | 'Valeur à remplir sur le "
        "tableau 2' | 'équivalent à récupérer sur le tableau 1'. Une valeur qui "
        "ne correspond à aucune colonne du tableau 1 (ex: '3.85') est écrite "
        "telle quelle, comme valeur fixe."
    ),
)

fichiers_matrices = {}  # nom -> bytes
if fichiers_matrices_up:
    for f in fichiers_matrices_up:
        if f.name.lower().endswith(".zip"):
            with zipfile.ZipFile(f) as z:
                for info in z.infolist():
                    if info.filename.lower().endswith(".xlsx") and not info.is_dir():
                        nom_court = os.path.basename(info.filename)
                        if not nom_court or nom_court.startswith("~$"):
                            continue
                        fichiers_matrices[nom_court] = z.read(info.filename)
        else:
            fichiers_matrices[f.name] = f.read()

if fichiers_matrices:
    st.caption(f"{len(fichiers_matrices)} fichier(s) matrice détecté(s).")

# --- Étape 2 : mapping code fiche -> matrice --------------------------------
mapping_valide = None
if fichiers_matrices:
    st.header("2. Correspondance code fiche → fichier matrice")
    st.caption(
        "Détection automatique à partir des noms de fichiers. Vérifiez et "
        "corrigez si besoin avant de générer les fichiers."
    )
    lignes_mapping = construire_mapping_fiches(fichiers_matrices)
    mapping_valide = st.data_editor(
        lignes_mapping,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "code_fiche": st.column_config.TextColumn("Code fiche (ex: BAR-TH-104)"),
            "fichier_matrice": st.column_config.SelectboxColumn(
                "Fichier matrice", options=sorted(fichiers_matrices.keys())
            ),
        },
        key="editeur_mapping",
    )

# --- Étape 2bis : tableau des données techniques par fiche -----------------
lignes_techniques = []
if fichier_technique:
    st.header("2bis. Colonnes techniques par fiche")
    st.caption(
        "Vérifiez la correspondance détectée : colonne(s) source du tableau 1 "
        "(concaténées avec '+') ou valeur fixe si aucune colonne ne correspond."
    )
    try:
        lignes_techniques = lire_tableau_technique(fichier_technique)
    except Exception as e:
        st.error(f"Erreur de lecture du tableau technique : {e}")
        lignes_techniques = []

    if lignes_techniques:
        apercu_technique = st.data_editor(
            lignes_techniques,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "code_fiche": st.column_config.TextColumn("Code fiche"),
                "colonne_cible": st.column_config.TextColumn("Colonne à remplir (tableau 2)"),
                "expression_source": st.column_config.TextColumn(
                    "Colonne(s) source (tableau 1) ou valeur fixe"
                ),
            },
            key="editeur_technique",
        )
        lignes_techniques = apercu_technique

# --- Étape 3 : analyse -------------------------------------------------------
if fichier_source and fichiers_matrices and mapping_valide:
    st.header("3. Analyse du tableau 1")

    try:
        operations, index_cols_source, header_row_source = lire_tableau_source(fichier_source)
    except Exception as e:
        st.error(f"Erreur de lecture du tableau 1 : {e}")
        st.stop()

    st.success(f"{len(operations)} opération(s) trouvée(s) dans le tableau 1.")

    if COLONNE_CLIENT not in index_cols_source:
        st.error(
            "Colonne client introuvable dans le tableau 1 "
            "('RAISON SOCIALE du bénéficiaire de l'opération')."
        )
        st.stop()
    if COLONNE_CODE_FICHE not in index_cols_source:
        st.error(
            "Colonne 'REFERENCE DE LA FICHE d'opération standardisée' "
            "introuvable dans le tableau 1."
        )
        st.stop()

    # code_fiche -> fichier_matrice (à partir du tableau édité)
    code_vers_fichier = {}
    for ligne in mapping_valide:
        code = (ligne.get("code_fiche") or "").strip().upper()
        fichier = ligne.get("fichier_matrice")
        if code and fichier:
            code_vers_fichier[code] = fichier

    # règles techniques par fiche (colonnes propres à chaque fiche CEE)
    regles_techniques_par_fiche = (
        construire_regles_techniques(lignes_techniques, index_cols_source)
        if lignes_techniques
        else {}
    )

    # Regroupement client -> fiche -> [operations]
    groupes = defaultdict(lambda: defaultdict(list))
    codes_sans_matrice = set()
    for op in operations:
        client = (op.get(COLONNE_CLIENT) or "(client non renseigné)").strip() \
            if isinstance(op.get(COLONNE_CLIENT), str) else (op.get(COLONNE_CLIENT) or "(client non renseigné)")
        code_fiche_brut = op.get(COLONNE_CODE_FICHE)
        code_fiche = str(code_fiche_brut).strip().upper() if code_fiche_brut else ""
        if code_fiche not in code_vers_fichier:
            codes_sans_matrice.add(code_fiche or "(vide)")
            continue
        groupes[client][code_fiche].append(op)

    if codes_sans_matrice:
        st.warning(
            "Codes fiches présents dans le tableau 1 mais sans matrice associée "
            "(opérations ignorées) : " + ", ".join(sorted(codes_sans_matrice))
        )

    # tableau récapitulatif
    recap = []
    for client, par_fiche in groupes.items():
        for code_fiche, ops in par_fiche.items():
            recap.append(
                {
                    "Client": client,
                    "Code fiche": code_fiche,
                    "Fichier matrice": code_vers_fichier[code_fiche],
                    "Nb opérations": len(ops),
                }
            )
    st.dataframe(recap, use_container_width=True)

    st.header("4. Génération des fichiers")
    if st.button("Générer les fichiers tableau 2 par client", type="primary"):
        resultats = {}  # client -> [(nom_fichier, bytes)]
        avertissements_colonnes = {}
        avertissements_technique = {}
        with st.spinner("Génération en cours..."):
            for client, par_fiche in groupes.items():
                fichiers_client = []
                for code_fiche, ops in par_fiche.items():
                    nom_matrice = code_vers_fichier[code_fiche]
                    contenu = fichiers_matrices[nom_matrice]
                    regles_fiche = regles_techniques_par_fiche.get(code_fiche, [])
                    try:
                        octets, colonnes_non_mappees, cibles_non_trouvees = remplir_matrice(
                            contenu, ops, index_cols_source, regles_fiche
                        )
                    except Exception as e:
                        st.error(f"Erreur sur {client} / {code_fiche} : {e}")
                        continue
                    base_nom = os.path.splitext(nom_matrice)[0]
                    client_safe = re.sub(r"[\\/:*?\"<>|]+", "_", str(client))[:80]
                    nom_sortie = f"{client_safe} - {base_nom}.xlsx"
                    fichiers_client.append((nom_sortie, octets))
                    if colonnes_non_mappees:
                        avertissements_colonnes[nom_matrice] = colonnes_non_mappees
                    if cibles_non_trouvees:
                        avertissements_technique[f"{code_fiche} ({nom_matrice})"] = cibles_non_trouvees
                resultats[client] = fichiers_client

        st.session_state["resultats"] = resultats
        st.session_state["avertissements_colonnes"] = avertissements_colonnes
        st.success("Génération terminée.")

        if avertissements_colonnes:
            with st.expander("⚠️ Colonnes de matrices non reconnues (laissées vides)"):
                for nom_matrice, cols in avertissements_colonnes.items():
                    st.write(f"**{nom_matrice}**")
                    st.write(", ".join(cols))

        if avertissements_technique:
            with st.expander("⚠️ Colonnes techniques non trouvées dans la matrice (à vérifier)"):
                for cle, cibles in avertissements_technique.items():
                    st.write(f"**{cle}**")
                    st.write(", ".join(cibles))

# --- Étape 5 : téléchargement + mails ---------------------------------------
resultats = st.session_state.get("resultats")
if resultats:
    st.header("5. Téléchargement")

    # zip global
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for client, fichiers in resultats.items():
            client_safe = re.sub(r"[\\/:*?\"<>|]+", "_", str(client))[:80]
            for nom_fichier, octets in fichiers:
                z.writestr(f"{client_safe}/{nom_fichier}", octets)
    zip_buffer.seek(0)

    st.download_button(
        "📦 Télécharger tous les fichiers (.zip)",
        data=zip_buffer,
        file_name=f"fiches_par_client_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
        mime="application/zip",
    )

    with st.expander("Voir / télécharger fichier par fichier"):
        for client, fichiers in resultats.items():
            st.subheader(client)
            for nom_fichier, octets in fichiers:
                st.download_button(
                    f"Télécharger : {nom_fichier}",
                    data=octets,
                    file_name=nom_fichier,
                    key=f"dl_{client}_{nom_fichier}",
                )

    # --- Mails -----------------------------------------------------------
    st.header("6. Mails clients")

    sujet_defaut = "Contrôle de vos opérations CEE — {client}"
    corps_defaut = (
        "Bonjour,\n\n"
        "Veuillez trouver ci-joint le(s) tableau(x) de contrôle relatif(s) à "
        "vos opérations ({nb_fiches} fiche(s), {nb_operations} opération(s) "
        "au total).\n\n"
        "N'hésitez pas à revenir vers nous pour toute question.\n\n"
        "Cordialement,"
    )

    sujet_modele = st.text_input("Sujet du mail (variables : {client}, {nb_fiches}, {nb_operations})", value=sujet_defaut)
    corps_modele = st.text_area(
        "Corps du mail (variables : {client}, {nb_fiches}, {nb_operations})",
        value=corps_defaut,
        height=200,
    )

    if not OUTLOOK_AVAILABLE:
        st.info(
            "Outlook (application de bureau Windows) n'est pas détecté sur cette "
            "machine — l'ouverture automatique des mails n'est disponible que si "
            "l'application est lancée sur un poste Windows avec Outlook installé "
            "et pywin32 (`pip install pywin32`). Vous pouvez malgré tout "
            "prévisualiser les mails ci-dessous et les composer manuellement."
        )

    with st.expander("Aperçu des mails par client"):
        for client, fichiers in resultats.items():
            nb_fiches = len(fichiers)
            nb_operations = sum(
                len(ops)
                for par_fiche in [groupes.get(client, {})]
                for ops in par_fiche.values()
            ) if "groupes" in dir() else nb_fiches
            sujet = sujet_modele.format(client=client, nb_fiches=nb_fiches, nb_operations=nb_operations)
            corps = corps_modele.format(client=client, nb_fiches=nb_fiches, nb_operations=nb_operations)
            st.markdown(f"**{client}** — pièces jointes : {', '.join(n for n, _ in fichiers)}")
            st.text(f"Objet : {sujet}\n\n{corps}")
            st.divider()

    generer_label = (
        "✉️ Générer les mails dans Outlook (brouillons)"
        if OUTLOOK_AVAILABLE
        else "✉️ Générer les mails (indisponible sans Outlook local)"
    )
    if st.button(generer_label, disabled=not OUTLOOK_AVAILABLE):
        outlook = win32com.client.Dispatch("Outlook.Application")
        nb_crees = 0
        # les pièces jointes doivent exister sur disque pour Outlook COM
        dossier_tmp = os.path.join(os.environ.get("TEMP", "."), "fiches_ceee_tmp")
        os.makedirs(dossier_tmp, exist_ok=True)
        for client, fichiers in resultats.items():
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.Subject = sujet_modele.format(
                client=client, nb_fiches=len(fichiers), nb_operations=len(fichiers)
            )
            mail.Body = corps_modele.format(
                client=client, nb_fiches=len(fichiers), nb_operations=len(fichiers)
            )
            for nom_fichier, octets in fichiers:
                chemin = os.path.join(dossier_tmp, nom_fichier)
                with open(chemin, "wb") as f:
                    f.write(octets)
                mail.Attachments.Add(chemin)
            mail.Display()  # ouvre le brouillon pour relecture (n'envoie pas)
            nb_crees += 1
        st.success(f"{nb_crees} brouillon(s) ouvert(s) dans Outlook.")

st.divider()
st.caption(
    "Astuce : le mapping code fiche → fichier matrice et les colonnes non "
    "reconnues sont affichés pour vérification — pensez à les contrôler avant "
    "un envoi réel aux clients."
)
