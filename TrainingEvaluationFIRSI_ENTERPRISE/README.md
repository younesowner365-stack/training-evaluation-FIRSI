# FIRSI - UM6P Enterprise

Fonctionnalités incluses :
- Thématiques dynamiques : Formation, Team Building, Ftour Ramadan et autres
- Questionnaires par thématique, modifiables sans coder
- Sessions et événements
- OTP par e-mail
- Invitations et relances e-mail
- Dashboard Chart.js
- Export Excel et PDF
- Analyse locale des commentaires
- Multi-départements
- Comptes Admin / RH
- PostgreSQL Neon en production, SQLite en local

Installation :
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m uvicorn main:app --reload

Compte initial :
admin / Admin@2026

Render :
Build: pip install -r requirements.txt
Start: uvicorn main:app --host 0.0.0.0 --port $PORT

Variables :
DATABASE_URL
SECRET_KEY
ENVIRONMENT=production
ADMIN_USERNAME
ADMIN_PASSWORD
SMTP_HOST
SMTP_PORT
SMTP_USERNAME
SMTP_PASSWORD
SMTP_FROM

Logo :
L’interface utilise l’identité texte FIRSI — UM6P. Le vrai fichier du logo officiel n’a pas été fourni.


## Accès collaborateur
La RH crée le collaborateur avec un mot de passe temporaire. À la première connexion, le collaborateur doit définir son propre mot de passe.
