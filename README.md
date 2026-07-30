# Add-in "Planning Visites" — version finale prête à déployer

## Ce que j'ai fait
- Domaine d'hébergement fixé sur `https://lanpark.github.io/planning_visites/`
- GUID unique généré pour le manifeste : `c18d7d4f-f69a-4536-bcb7-03e761f3cfc5`
- URL de l'app Power Apps intégrée dans `dialog.html`

⚠️ **Note sur l'URL Power Apps** : j'ai retiré les paramètres `hint=...` et `sourcetime=...` de l'URL que tu m'as donnée. Ce sont des paramètres de session temporaires (liés à ta session d'édition dans Studio), pas nécessaires pour lancer l'app — l'URL avec juste `tenantId` suffit et reste stable dans le temps.

## Étapes restantes (rapide)

### 1. Publier le repo GitHub Pages
Dans le repo `lanpark/planning_visites` (ou celui qui sert `lanpark.github.io/planning_visites/`), place ces 4 fichiers **à la racine** :
- `manifest.json` *(celui-ci n'a pas besoin d'être sur GitHub Pages — voir étape 3, mais tu peux le stocker dans le même repo par simplicité)*
- `commands.html`
- `commands.js`
- `dialog.html`

Ajoute aussi un dossier `assets/` avec 3 icônes PNG (16×16, 32×32, 80×80 px) — n'importe quelle icône simple fait l'affaire pour commencer.

### 2. Vérifier l'accès une fois publié
Une fois GitHub Pages actif (peut prendre 1-2 minutes après un push), teste dans un navigateur :
- `https://lanpark.github.io/planning_visites/commands.html` → doit s'afficher (page vide, normal)
- `https://lanpark.github.io/planning_visites/dialog.html` → doit rediriger automatiquement vers ton app Power Apps

### 3. Déployer le manifeste dans Microsoft 365
- **Test perso rapide** : Outlook (nouveau ou web) → Paramètres → Add-ins → "Ajouter mon propre add-in" → charger `manifest.json` (tu peux le télécharger depuis GitHub ou le fournir directement depuis ce fichier).
- **Déploiement pour toute l'organisation** : [admin.microsoft.com](https://admin.microsoft.com) → **Paramètres** → **Apps intégrées** → **Charger une app personnalisée** → sélectionner `manifest.json` → assigner aux utilisateurs/groupes concernés.

Une fois assigné, le bouton "Planning Visites" apparaît automatiquement dans le ruban, identique sur nouveau Outlook et Outlook desktop classique — sans action supplémentaire de ta part.

### 4. Test final
Clique sur le bouton "Planning Visites" dans le ruban Outlook → une fenêtre de dialogue doit s'ouvrir et charger directement ton app Power Apps.

## En cas de mise à jour future
- **Changement de logique métier / interface** → modifie uniquement l'app Power Apps elle-même, rien à republier côté Outlook.
- **Changement d'URL d'app (nouvel environnement, republication ailleurs)** → mets à jour l'URL dans `dialog.html` sur GitHub, republie sur Pages — pas besoin de retoucher le manifeste ni de le redéployer dans Microsoft 365.
- **Changement de domaine d'hébergement** → là seulement, il faut modifier et redéployer `manifest.json`.
