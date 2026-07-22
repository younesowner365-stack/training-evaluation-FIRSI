import os
from main import SessionLocal, User, hash_password

USERNAME = os.getenv("ADMIN_USERNAME", "admin")
NEW_PASSWORD = os.getenv("ADMIN_PASSWORD", "Admin@2026")


def main() -> None:
    db = SessionLocal()
    try:
        user = db.query(User).filter(User.username == USERNAME).first()
        if user is None:
            user = User(
                fullname="Administrateur principal",
                username=USERNAME,
                email=os.getenv("ADMIN_EMAIL"),
                password_hash=hash_password(NEW_PASSWORD),
                role="ADMIN",
                active=True,
                must_change_password=True,
                failed_attempts=0,
            )
            db.add(user)
            action = "créé"
        else:
            user.password_hash = hash_password(NEW_PASSWORD)
            user.role = "ADMIN"
            user.active = True
            user.must_change_password = True
            user.failed_attempts = 0
            user.locked_until = None
            action = "réinitialisé"
        db.commit()
        print(f"Compte administrateur {action}.")
        print(f"Identifiant : {USERNAME}")
        print(f"Mot de passe temporaire : {NEW_PASSWORD}")
    finally:
        db.close()


if __name__ == "__main__":
    main()
