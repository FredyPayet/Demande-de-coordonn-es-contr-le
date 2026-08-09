# Automatisation fiches CEE — Tableau 1 → Tableau 2 par client

## Installation

```bash
pip install -r requirements.txt
```

Sous Windows, si vous voulez la génération automatique des mails Outlook, ajoutez :
```bash
pip install pywin32
```

## Lancement

```bash
streamlit run app.py
```

Une page s'ouvre dans votre navigateur (généralement http://localhost:8501).

## Utilisation

1. **Fichiers** : déposez le tableau 1 (export de contrôle .xlsx), les fichiers
   matrices (tableau 2) — soit un par un, soit directement vos deux .zip
   (`MATRICE_FICHE_BAR_EN.zip`, `Matrice_fiches_BAR_TH.zip`) — et, si besoin,
   le tableau de correspondance des données techniques par fiche.
2. **Correspondance code fiche → fichier matrice** : l'application détecte
   automatiquement, à partir des noms de fichiers, quel code de fiche (ex.
   `BAR-TH-104`) correspond à quelle matrice. **Vérifiez ce tableau** avant de
   continuer — vous pouvez corriger une ligne ou en ajouter/supprimer
   directement dans le tableau éditable.
3. **Colonnes techniques par fiche** (si un tableau technique est fourni) :
   l'application affiche, pour vérification, la résolution de chaque ligne de
   votre tableau (quelle(s) colonne(s) du tableau 1 sont utilisées, ou si la
   valeur est fixe). Vous pouvez corriger ici aussi.
4. **Analyse** : l'application lit le tableau 1, regroupe les opérations par
   client, et affiche un récapitulatif (nombre d'opérations par client et par
   fiche). Si un code de fiche du tableau 1 n'a pas de matrice associée,
   un avertissement s'affiche et ces opérations sont ignorées.
5. **Génération** : cliquez sur "Générer les fichiers tableau 2 par client".
   Un fichier est créé par client et par type de fiche (un client avec deux
   types d'opérations différents recevra deux fichiers).
6. **Téléchargement** : récupérez tout en un .zip, ou fichier par fichier.
7. **Mails** : personnalisez le sujet et le corps du mail type (variables
   disponibles : `{client}`, `{nb_fiches}`, `{nb_operations}`), prévisualisez,
   puis :
   - **Sur un poste Windows avec Outlook installé** : le bouton crée un
     brouillon par client dans Outlook, pièces jointes incluses, prêt à
     relire avant envoi (rien n'est envoyé automatiquement).
   - **Ailleurs** (Mac/Linux, ou test) : l'aperçu texte reste disponible pour
     copier-coller manuellement ; le bouton Outlook est désactivé.

## Comment fonctionne le report des colonnes

Le report ne se fait **pas** par position de colonne (l'ordre diffère d'une
matrice à l'autre selon la fiche) mais par **correspondance du nom d'en-tête**
entre le tableau 1 et chaque matrice. Les 3 colonnes suivantes, présentes
uniquement dans le tableau 2, ne sont jamais remplies, comme demandé :
- numéro de téléphone du bénéficiaire
- Adresse de courriel du bénéficiaire
- Montant du rôle actif et incitatif (€)

### Colonnes techniques propres à chaque fiche CEE

En plus des colonnes communes, chaque fiche CEE peut nécessiter le report de
colonnes techniques spécifiques (ex. surface, épaisseur, marque et référence
d'isolant, résistance...). Ces correspondances sont lues depuis le fichier
"tableau de correspondance des données techniques" que vous fournissez, avec 3
colonnes : le code fiche, le nom de la colonne à remplir dans le tableau 2, et
soit un ou plusieurs noms de colonnes du tableau 1 à concaténer avec un `+`
(ex. `Marque isolant + Référence isolant`), soit une **valeur fixe** si aucune
colonne ne correspond (ex. `3.85` pour BAR-EN-102) — l'application écrit alors
cette valeur telle quelle sur toutes les opérations concernées par cette fiche.

Ce fichier peut être complété au fil du temps (nouvelles fiches, nouvelles
colonnes) sans modification du code : il suffit d'ajouter des lignes.

Un point de vigilance identifié pendant les tests : certaines cellules des
modèles matrices avaient un format "date" préréglé qui aurait déformé
l'affichage de valeurs numériques (ex. un SIREN). Le script réinitialise
systématiquement le format des cellules qu'il remplit pour éviter ce piège.

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
- La génération Outlook nécessite que l'application tourne sur le poste
  Windows où Outlook (application de bureau) est ouvert.
