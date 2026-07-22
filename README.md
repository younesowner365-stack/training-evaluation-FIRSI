# Training Evaluation FIRSI — WOW V2

## Fonctionnalités
- Page d'accès unique : Admin/RH et Collaborateur
- Comptes Admin et RH avec rôles
- Comptes collaborateurs avec mot de passe temporaire
- Changement obligatoire du mot de passe à la première connexion
- Gestion collaborateurs, formations, affectations et évaluations
- Dashboard interactif Chart.js
- Export Excel
- SQLite local / Neon PostgreSQL sur Render

## Installation locale
```powershell
python -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m uvicorn main:app --reload
```

## Premier accès
- Identifiant : `admin`
- Mot de passe temporaire : `Admin@2026`
- Le changement de mot de passe est obligatoire.

En production, configurez dans Render :
- `DATABASE_URL` : connexion Neon
- `SECRET_KEY` : chaîne aléatoire longue
- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`
- `ENVIRONMENT=production`

## Mise à jour GitHub / Render
```powershell
git add .
git commit -m "Mise à jour V2 WOW"
git push
```
