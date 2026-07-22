# Training Evaluation Platform — Full V1

Connexion RH : admin / admin123

Lancement :
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m uvicorn app:app --reload
```

Ouvrir http://127.0.0.1:8000
