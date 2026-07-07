# MS Project → Excel — Module de transfert autonome

Deux outils **100 % hors ligne** (aucune connexion internet, aucune librairie
externe, aucun add-in) pour reprendre un planning MS Project et le réafficher
dans Excel avec le même look (Gantt coloré, tâches récapitulatives en gras,
jalons en losange, chemin critique en rouge, week-ends grisés).

Les deux outils lisent le même format d'entrée : l'**export XML natif de MS
Project** (`Fichier > Enregistrer sous > Type : XML`).

## 1. `index.html` — Outil web autonome (recommandé)

- Ouvrez simplement `index.html` en le double-cliquant (aucun serveur requis).
- Glissez-déposez votre export XML MS Project, ou cliquez sur la zone pour le
  choisir, ou testez avec le bouton **« Charger un exemple »**.
- Un aperçu du Gantt s'affiche directement dans la page.
- Cliquez sur **« Générer le fichier Excel (.xlsx) »** : un classeur `.xlsx`
  est généré et téléchargé, avec :
  - Colonnes N°, Nom de la tâche (indentée selon la hiérarchie), Durée, Début,
    Fin, % achevé, Prédécesseurs ;
  - Grille temporelle (jour, semaine ou mois selon la durée du projet) ;
  - Barres bleu foncé (réalisé) / bleu clair (restant), barres grises pour les
    tâches récapitulatives, losanges noirs pour les jalons, rouge pour le
    chemin critique (optionnel), grisé pour les week-ends (optionnel) ;
  - Volets figés (colonnes de gauche + lignes d'en-tête) pour naviguer
    facilement dans un grand planning.
- Ce fichier `.xlsx` s'ouvre nativement dans Excel, LibreOffice Calc, Google
  Sheets (import), etc. — aucune macro n'est nécessaire pour le consulter.

Le fichier `.xlsx` est construit "à la main" en JavaScript pur (générateur
ZIP + XML Office Open XML minimal embarqué dans la page) : aucune librairie
externe (type SheetJS) n'est chargée, donc **aucune connexion internet
n'est requise**, y compris pour l'export.

## 2. `ImportMSProjectXML.bas` — Macro Excel (VBA)

Pour ceux qui préfèrent importer directement dans un classeur Excel existant
et retravailler le résultat avec des formules/macros Excel.

Comme un fichier `.xlsm` compilé ne peut pas être généré sans Excel lui-même,
la macro est fournie sous forme de **module VBA texte (`.bas`)** à importer :

1. Ouvrez Excel, créez un classeur puis enregistrez-le en **`.xlsm`**
   (Classeur Excel prenant en charge les macros).
2. Activez l'onglet Développeur si besoin (`Fichier > Options > Ruban`).
3. `Alt+F11` pour ouvrir l'éditeur VBA, puis
   `Fichier > Importer un fichier...` et sélectionnez
   `ImportMSProjectXML.bas`.
4. `Alt+F8`, choisissez `ImporterProjetMSProject`, cliquez sur **Exécuter**.
5. Sélectionnez votre export XML MS Project (ou `ExempleProjet.xml` fourni
   pour tester) puis répondez aux deux questions (chemin critique en rouge ?
   week-ends grisés ?).
6. Une feuille **« Gantt »** est créée/recréée dans le classeur, avec la même
   charte visuelle que l'outil HTML.

La macro utilise `MSXML2.DOMDocument.6.0`, fourni en standard avec Windows —
**aucune connexion ni add-in n'est nécessaire**. Elle fonctionne sous Excel
Windows ; Excel pour Mac ne dispose pas du composant MSXML et n'est donc pas
supporté par cette macro (utilisez l'outil HTML dans ce cas, qui fonctionne
sur toutes les plateformes via n'importe quel navigateur).

## 3. `ExempleProjet.xml` — Fichier d'exemple

Petit planning fictif ("Construction Pyramide") utilisable pour tester les
deux outils sans avoir besoin d'un vrai fichier MS Project sous la main.

## Limites connues / simplifications

- Les durées MS Project sont ré-exprimées en jours pleins (8h = 1 jour),
  comme c'est l'usage courant en reporting simplifié.
- Les prédécesseurs sont affichés par numéro de tâche uniquement (sans le
  type de liaison FS/SS/FF ni le décalage), pour rester lisible dans le
  format Gantt compact.
- Le calendrier des jours ouvrés est simplifié (week-end = samedi/dimanche) ;
  les calendriers personnalisés / jours fériés de MS Project ne sont pas
  repris.
- Pour les très longs plannings, la grille temporelle bascule automatiquement
  en vue hebdomadaire puis mensuelle afin de rester lisible et compacte.
