# FIRSI Formation Experience — Version Pro V4

Plateforme FastAPI avec trois espaces sécurisés : Administrateur, RH et Collaborateur.

## Installation locale

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python reset_admin.py
python -m uvicorn main:app --reload
```

Ouvrir : `http://127.0.0.1:8000`

Compte initial local :

- Identifiant : `admin`
- Mot de passe temporaire : `Admin@2026`

## Render / Neon

Variables recommandées :

- `DATABASE_URL` : chaîne PostgreSQL Neon
- `SECRET_KEY` : chaîne aléatoire longue
- `ENVIRONMENT` : `production`
- `ADMIN_USERNAME` : identifiant administrateur
- `ADMIN_PASSWORD` : mot de passe temporaire initial
- `ADMIN_EMAIL` : facultatif

Pour forcer une réinitialisation unique de l’administrateur sur Render, ajouter temporairement :

- `RESET_ADMIN_ON_START=true`

Après connexion et changement du mot de passe, supprimer cette variable ou la passer à `false`.

Commande de démarrage Render :

```bash
uvicorn main:app --host 0.0.0.0 --port $PORT
```

## Footer

`© FIRSI - UM6P`
