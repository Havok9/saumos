# Saumos · Le carnet de la caserne

Site statique : aucune installation, aucun build.

## Contenu
- `index.html` : le carnet (Accueil, Santé, Fugue)
- `urgence.html` : la page urgences
- `assets/` : photos de Saumos

## Mettre en ligne avec Vercel
1. Crée un dépôt GitHub et envoie-y le contenu de ce dossier (index.html à la racine).
2. Sur vercel.com : Add New → Project → importe le dépôt.
3. Framework Preset : **Other**. Laisse Build Command et Output Directory vides.
4. Deploy. L'adresse ressemble à `saumos.vercel.app`.


## À savoir
- Le site se met à jour toutes les 20 secondes et quand on revient sur l'onglet.
- Une tâche cochée ne peut plus être décochée (vérifié aussi côté Google Sheet).
- Après 20 h, si une tâche du jour manque, l'accueil passe en rouge.
