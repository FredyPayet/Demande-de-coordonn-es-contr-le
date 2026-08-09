# Automatisation fiches CEE — Tableau 1 → Tableau 2 par client

## Installation

```bash
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run app.py
```

Une page s'ouvre dans votre navigateur (généralement http://localhost:8501).

## Utilisation

1. **Fichiers** : déposez le tableau 1 (export de contrôle .xlsx) et les fichiers
   matrices (tableau 2) — soit un par un, soit directement vos deux .zip
   (`MATRICE_FICHE_BAR_EN.zip`, `Matrice_fiches_BAR_TH.zip`).
2. **Correspondance code fiche → fichier matrice** : l'application détecte
   automatiquement, à partir des noms de fichiers, quel code de fiche (ex.
   `BAR-TH-104`) correspond à quelle matrice. **Vérifiez ce tableau** avant de
   continuer — vous pouvez corriger une ligne ou en ajouter/supprimer
   directement dans le tableau éditable.
3. **Colonnes techniques par fiche** : la correspondance (surface, épaisseur,
   marque d'isolant...) est **intégrée directement dans le code** — plus besoin
   de fournir de fichier à chaque lancement. Elle reste affichée et modifiable
   ponctuellement à l'écran pour vérification, mais toute modification faite
   ici ne persiste pas après fermeture de la page (voir plus bas pour la
   modifier de façon permanente).
4. **Analyse** : l'application lit le tableau 1, détecte le numéro de lot,
   regroupe les opérations par client, et affiche un récapitulatif (nombre
   d'opérations par client et par fiche). Si un code de fiche du tableau 1
   n'a pas de matrice associée, un avertissement s'affiche et ces opérations
   sont ignorées.
5. **Génération** : deux boutons indépendants sont disponibles :
   - **"Générer les fichiers tableau 2 par client"** : un fichier par client et
     par type de fiche, plus l'objet et le corps du mail type correspondants.
   - **"Générer le tableau 2 complet (tous clients)"** : un fichier par fiche,
     mais contenant **toutes** les opérations de **tous** les clients — pratique
     pour une vue d'ensemble ou un contrôle global.
   Vous pouvez utiliser l'un, l'autre, ou les deux.
6. **Actions par client** : pour chaque client (et chaque fiche s'il y en a
   plusieurs), trois boutons sont disponibles :
   - **📋 Copier l'objet** — copie l'objet du mail dans le presse-papiers.
   - **📋 Copier le mail** — copie le corps du mail type (coordonnées du
     client, du lot et du bureau de contrôle déjà intégrées) dans le
     presse-papiers.
   - **⬇️ Télécharger le tableau 2** — télécharge directement le fichier Excel
     rempli pour ce client (téléchargement géré par le navigateur, sans passer
     par le serveur Streamlit).
   Chaque bouton passe en **vert** une fois l'action effectuée, pour un suivi
   visuel rapide de ce qui a déjà été traité.

## Le mail type

Le corps du mail reprend le modèle fourni (`Mail_typee_DDC.docx`), avec les
champs suivants complétés automatiquement à partir du tableau 2 généré :

| Champ du modèle | Source |
|---|---|
| Fiche CEE | Code de la fiche (ex. `BAR-EN-103`) |
| Organisme de contrôle | Colonne "RAISON sociale de l'organisme de contrôle" du tableau 1 (ou son SIREN si la raison sociale n'est pas disponible) |
| Plage de colonnes à remplir | Calculée automatiquement : de "Nom interlocuteur - contrôle sur site" jusqu'à juste avant "INFORMATIONS COMPLEMENTAIRES", dans la matrice utilisée |
| Date limite de retour | Date du jour + 1 semaine |
| Liste des dossiers concernés | 6 premiers caractères de "REFERENCE interne de l'opération", dédupliqués, un par ligne |

Si l'organisme de contrôle ou la plage de colonnes n'est pas trouvé pour un
client (donnée absente du tableau 1, ou structure de matrice non reconnue),
le mail est quand même généré avec une mention explicite à compléter
manuellement (`[organisme de contrôle non trouvé]` ou équivalent), et un
avertissement récapitulatif s'affiche après la génération.

**Hypothèse retenue pour l'objet du mail** : en l'absence d'un numéro client
identifié dans le tableau 1, l'objet utilise le **nom du client** (raison
sociale) : `"Demande de coordonnées" - {numéro de lot} - {nom du client}`. Si
vous disposez d'un véritable numéro client distinct du nom, signalez-le pour
adapter la formule.

## Comment fonctionne le report des colonnes

Le report ne se fait **pas** par position de colonne (l'ordre diffère d'une
matrice à l'autre selon la fiche) mais par **correspondance du nom d'en-tête**
entre le tableau 1 et chaque matrice, restreinte à la zone "identité" de la
matrice (jusqu'à la colonne SIRET + quelques colonnes voisines). Les colonnes
suivantes ne sont jamais remplies, comme demandé :
- numéro de téléphone du bénéficiaire
- Adresse de courriel du bénéficiaire
- Montant du rôle actif et incitatif (€)
- Fonction (colonnes de la partie "audit terrain" du bureau de contrôle)

Toutes les cellules remplies par l'application reçoivent une mise en forme
uniforme : police noire, non grasse — quelle que soit la mise en forme
héritée du modèle Excel d'origine sur cette cellule.

Un point de vigilance identifié pendant les tests : certaines cellules des
modèles matrices avaient un format "date" préréglé qui aurait déformé
l'affichage de valeurs numériques (ex. un SIREN). Le script réinitialise
systématiquement le format des cellules qu'il remplit pour éviter ce piège.

## Nommage des fichiers et de l'objet du mail

- **Numéro de lot** : détecté automatiquement dans le nom du fichier tableau 1
  (motif attendu : `..._lot_de_controle_NUMERO_...`, ex.
  `ODICEE_Export_lot_de_controle_XX818P_...xlsx` → `XX818P`). Il est affiché
  et modifiable à l'écran (étape 4) avant génération — si le motif n'est pas
  détecté, un avertissement s'affiche et vous devez le renseigner vous-même.
- **Fichier groupé** ("tableau 2 complet") : nommé `{numéro de lot}.xlsx`. Si
  plusieurs types de fiches sont présents dans le lot, un fichier est généré
  par fiche et le code fiche est ajouté au nom pour éviter les collisions
  (ex. `XX818P - BAR-EN-103.xlsx`).
- **Fichiers par client** : nommés `Demande de coordonnées - {numéro de lot}
  - {nom du client}.xlsx`. Si un même client a des opérations relevant de
  plusieurs types de fiches différents, le code fiche est ajouté en fin de
  nom pour éviter que les fichiers ne s'écrasent entre eux (ex. `Demande de
  coordonnées - XX818P - Client 3 - BAR-EN-103.xlsx`).
- **Objet du mail** : `"Demande de coordonnées" - {numéro de lot} - {nom du
  client}` (voir hypothèse ci-dessus).

## Colonnes techniques propres à chaque fiche CEE

En plus des colonnes communes, chaque fiche CEE peut nécessiter le report de
colonnes techniques spécifiques (ex. surface, épaisseur, marque et référence
d'isolant, résistance...). Cette correspondance est **intégrée directement
dans le code** (variable `DONNEES_TECHNIQUES_PAR_FICHE` en haut du fichier
`app.py`), avec pour chaque ligne : le code fiche, le nom de la colonne à
remplir dans le tableau 2, et soit un ou plusieurs noms de colonnes du
tableau 1 à concaténer avec un `+` (ex. `Marque isolant + Référence isolant`),
soit une **valeur fixe** si aucune colonne ne correspond (ex. `3.85` pour
BAR-EN-102) — l'application écrit alors cette valeur telle quelle sur toutes
les opérations concernées par cette fiche.

**Pour ajouter ou modifier une règle de façon permanente** (nouvelle fiche,
nouvelle colonne technique), éditez la liste `DONNEES_TECHNIQUES_PAR_FICHE`
dans `app.py` — chaque entrée suit le même format :
```python
{"code_fiche": "BAR-TH-XXX", "colonne_cible": "Nom de la colonne dans le tableau 2", "expression_source": "Colonne du tableau 1"},
```
L'écran "2bis. Colonnes techniques par fiche" de l'application permet de
vérifier et corriger ponctuellement ces règles pour une session, sans toucher
au code — mais ces changements ne sont pas conservés après fermeture de la
page.

## Limites connues / à vérifier avec vous

- Le mapping automatique code fiche → fichier est basé sur le nom des
  fichiers matrices. Pour les fichiers couvrant plusieurs fiches avec une
  numérotation ambiguë (ex. `BAR-EN-102-107`), l'algorithme peut se tromper
  sur les codes couverts — **vérifiez toujours le tableau de correspondance**
  avant de générer.
- Certaines colonnes propres à une fiche (ex. `date de la facture` au lieu de
  `date d'achèvement de l'opération`) n'ont pas d'équivalent dans le tableau 1
  et restent donc vides : c'est un choix volontaire (pas de donnée fiable à y
  mettre), à confirmer avec vous si un report différent est souhaité.
- Si une colonne "cible" du tableau technique n'est trouvée dans aucune
  colonne de la matrice correspondante, un avertissement s'affiche après la
  génération (elle n'est alors pas remplie) — vérifiez l'orthographe exacte
  dans votre tableau de correspondance dans ce cas.
- Le bouton "Copier" utilise l'API presse-papiers du navigateur (avec un
  repli automatique si elle est indisponible) — fonctionne dans tous les
  navigateurs modernes, y compris via `streamlit run` local.
- L'objet du mail utilise le nom du client, pas un numéro client dédié (voir
  section "Le mail type" ci-dessus).
