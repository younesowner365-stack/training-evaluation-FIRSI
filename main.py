from __future__ import annotations

from datetime import datetime
from pathlib import Path
import io
import secrets

from fastapi import FastAPI, Request, Form, Depends, HTTPException
from fastapi.responses import RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from starlette.middleware.sessions import SessionMiddleware

from sqlalchemy import (
    create_engine,
    Column,
    Integer,
    String,
    Boolean,
    DateTime,
    ForeignKey,
    Text,
    UniqueConstraint,
)
from sqlalchemy.orm import (
    declarative_base,
    sessionmaker,
    Session,
    relationship,
)

from openpyxl import Workbook


# =========================================================
# CONFIGURATION GÉNÉRALE
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

DATABASE_URL = f"sqlite:///{BASE_DIR / 'evaluation.db'}"


# =========================================================
# CONFIGURATION BASE DE DONNÉES
# =========================================================

engine = create_engine(
    DATABASE_URL,
    connect_args={
        "check_same_thread": False
    },
)

SessionLocal = sessionmaker(
    bind=engine,
    autoflush=False,
    autocommit=False,
)

Base = declarative_base()


# =========================================================
# MODÈLES DE DONNÉES
# =========================================================

class Collaborateur(Base):

    __tablename__ = "collaborateurs"

    id = Column(
        Integer,
        primary_key=True,
        index=True,
    )

    code = Column(
        String(20),
        unique=True,
        nullable=False,
        index=True,
    )

    nom = Column(
        String(120),
        nullable=False,
    )

    prenom = Column(
        String(120),
        nullable=False,
    )

    email = Column(
        String(180),
        nullable=True,
    )

    direction = Column(
        String(180),
        nullable=True,
    )

    fonction = Column(
        String(180),
        nullable=True,
    )

    token = Column(
        String(100),
        unique=True,
        nullable=False,
        index=True,
    )

    actif = Column(
        Boolean,
        default=True,
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow,
    )

    affectations = relationship(
        "Affectation",
        back_populates="collaborateur",
        cascade="all, delete-orphan",
    )


class Formation(Base):

    __tablename__ = "formations"

    id = Column(
        Integer,
        primary_key=True,
        index=True,
    )

    titre = Column(
        String(255),
        nullable=False,
    )

    type_formation = Column(
        String(50),
        nullable=False,
    )

    description = Column(
        Text,
        nullable=True,
    )

    active = Column(
        Boolean,
        default=True,
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow,
    )

    affectations = relationship(
        "Affectation",
        back_populates="formation",
        cascade="all, delete-orphan",
    )


class Affectation(Base):

    __tablename__ = "affectations"

    __table_args__ = (
        UniqueConstraint(
            "collaborateur_id",
            "formation_id",
            name="uq_affectation_collaborateur_formation",
        ),
    )

    id = Column(
        Integer,
        primary_key=True,
        index=True,
    )

    collaborateur_id = Column(
        Integer,
        ForeignKey("collaborateurs.id"),
        nullable=False,
    )

    formation_id = Column(
        Integer,
        ForeignKey("formations.id"),
        nullable=False,
    )

    statut = Column(
        String(30),
        default="À évaluer",
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow,
    )

    collaborateur = relationship(
        "Collaborateur",
        back_populates="affectations",
    )

    formation = relationship(
        "Formation",
        back_populates="affectations",
    )

    evaluation = relationship(
        "Evaluation",
        back_populates="affectation",
        uselist=False,
        cascade="all, delete-orphan",
    )


class Evaluation(Base):

    __tablename__ = "evaluations"

    id = Column(
        Integer,
        primary_key=True,
        index=True,
    )

    affectation_id = Column(
        Integer,
        ForeignKey("affectations.id"),
        unique=True,
        nullable=False,
    )

    contenu = Column(
        Integer,
        nullable=False,
    )

    formateur = Column(
        Integer,
        nullable=False,
    )

    organisation = Column(
        Integer,
        nullable=False,
    )

    application_poste = Column(
        Integer,
        nullable=False,
    )

    satisfaction = Column(
        Integer,
        nullable=False,
    )

    recommande = Column(
        String(3),
        nullable=False,
    )

    observations = Column(
        Text,
        nullable=True,
    )

    points_forts = Column(
        Text,
        nullable=True,
    )

    points_ameliorer = Column(
        Text,
        nullable=True,
    )

    date_reponse = Column(
        DateTime,
        default=datetime.utcnow,
    )

    affectation = relationship(
        "Affectation",
        back_populates="evaluation",
    )


Base.metadata.create_all(bind=engine)


# =========================================================
# INITIALISATION FASTAPI
# =========================================================

app = FastAPI(
    title="Training Evaluation Platform",
    description="Plateforme d’évaluation des formations",
    version="1.0.0",
)


app.add_middleware(
    SessionMiddleware,
    secret_key="training-evaluation-secret-key-2026",
    same_site="lax",
    https_only=False,
)


app.mount(
    "/static",
    StaticFiles(
        directory=BASE_DIR / "app" / "static"
    ),
    name="static",
)


templates = Jinja2Templates(
    directory=BASE_DIR / "app" / "templates"
)


# =========================================================
# OUTILS INTERNES
# =========================================================

def get_db():

    db = SessionLocal()

    try:
        yield db

    finally:
        db.close()


def verifier_connexion_rh(request: Request):

    if not request.session.get("rh_authenticated"):

        return RedirectResponse(
            url="/",
            status_code=303,
        )

    return None


def generer_code_collaborateur(db: Session) -> str:

    dernier_collaborateur = (
        db.query(Collaborateur)
        .order_by(Collaborateur.id.desc())
        .first()
    )

    if dernier_collaborateur is None:

        prochain_numero = 1

    else:

        prochain_numero = dernier_collaborateur.id + 1

    return f"COL-{prochain_numero:03d}"


# =========================================================
# AUTHENTIFICATION RH
# =========================================================

@app.get("/")
def page_connexion(request: Request):

    if request.session.get("rh_authenticated"):

        return RedirectResponse(
            url="/dashboard",
            status_code=303,
        )

    return templates.TemplateResponse(
        request=request,
        name="login.html",
        context={
            "error": request.query_params.get("error"),
        },
    )


@app.post("/login")
def connexion(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
):

    if username == "admin" and password == "admin123":

        request.session["rh_authenticated"] = True

        request.session["username"] = username

        return RedirectResponse(
            url="/dashboard",
            status_code=303,
        )

    return RedirectResponse(
        url="/?error=1",
        status_code=303,
    )


@app.get("/logout")
def deconnexion(request: Request):

    request.session.clear()

    return RedirectResponse(
        url="/",
        status_code=303,
    )


# =========================================================
# TABLEAU DE BORD RH
# =========================================================

@app.get("/dashboard")
def dashboard(
    request: Request,
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    total_collaborateurs = (
        db.query(Collaborateur)
        .filter(Collaborateur.actif.is_(True))
        .count()
    )

    total_formations = (
        db.query(Formation)
        .filter(Formation.active.is_(True))
        .count()
    )

    total_affectations = (
        db.query(Affectation)
        .count()
    )

    total_evaluations = (
        db.query(Evaluation)
        .count()
    )

    if total_affectations > 0:

        participation = round(
            total_evaluations
            / total_affectations
            * 100,
            1,
        )

    else:

        participation = 0

    evaluations = (
        db.query(Evaluation)
        .all()
    )

    if evaluations:

        satisfaction = round(
            sum(
                evaluation.satisfaction
                for evaluation in evaluations
            )
            / len(evaluations),
            2,
        )

        recommandation = round(
            sum(
                1
                for evaluation in evaluations
                if evaluation.recommande == "Oui"
            )
            / len(evaluations)
            * 100,
            1,
        )

    else:

        satisfaction = 0

        recommandation = 0

    return templates.TemplateResponse(
        request=request,
        name="dashboard.html",
        context={
            "active_page": "dashboard",
            "total_collaborateurs": total_collaborateurs,
            "total_formations": total_formations,
            "total_evaluations": total_evaluations,
            "participation": participation,
            "satisfaction": satisfaction,
            "recommandation": recommandation,
        },
    )


# =========================================================
# GESTION DES COLLABORATEURS
# =========================================================

@app.get("/collaborateurs")
def page_collaborateurs(
    request: Request,
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    collaborateurs = (
        db.query(Collaborateur)
        .order_by(Collaborateur.id.desc())
        .all()
    )

    return templates.TemplateResponse(
        request=request,
        name="collaborateurs.html",
        context={
            "active_page": "collaborateurs",
            "collaborateurs": collaborateurs,
        },
    )


@app.post("/collaborateurs")
def ajouter_collaborateur(
    request: Request,
    nom: str = Form(...),
    prenom: str = Form(...),
    email: str = Form(""),
    direction: str = Form(""),
    fonction: str = Form(""),
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    collaborateur = Collaborateur(
        code=generer_code_collaborateur(db),
        nom=nom.strip(),
        prenom=prenom.strip(),
        email=email.strip() or None,
        direction=direction.strip() or None,
        fonction=fonction.strip() or None,
        token=secrets.token_urlsafe(24),
        actif=True,
    )

    db.add(collaborateur)

    db.commit()

    return RedirectResponse(
        url="/collaborateurs",
        status_code=303,
    )


@app.post("/collaborateurs/{collaborateur_id}/toggle")
def activer_desactiver_collaborateur(
    collaborateur_id: int,
    request: Request,
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    collaborateur = db.get(
        Collaborateur,
        collaborateur_id,
    )

    if collaborateur is None:

        raise HTTPException(
            status_code=404,
            detail="Collaborateur introuvable",
        )

    collaborateur.actif = not collaborateur.actif

    db.commit()

    return RedirectResponse(
        url="/collaborateurs",
        status_code=303,
    )


# =========================================================
# GESTION DES FORMATIONS
# =========================================================

@app.get("/formations")
def page_formations(
    request: Request,
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    formations = (
        db.query(Formation)
        .order_by(Formation.id.desc())
        .all()
    )

    return templates.TemplateResponse(
        request=request,
        name="formations.html",
        context={
            "active_page": "formations",
            "formations": formations,
        },
    )


@app.post("/formations")
def ajouter_formation(
    request: Request,
    titre: str = Form(...),
    type_formation: str = Form(...),
    description: str = Form(""),
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    formation = Formation(
        titre=titre.strip(),
        type_formation=type_formation,
        description=description.strip() or None,
        active=True,
    )

    db.add(formation)

    db.commit()

    return RedirectResponse(
        url="/formations",
        status_code=303,
    )


@app.post("/formations/{formation_id}/toggle")
def activer_desactiver_formation(
    formation_id: int,
    request: Request,
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    formation = db.get(
        Formation,
        formation_id,
    )

    if formation is None:

        raise HTTPException(
            status_code=404,
            detail="Formation introuvable",
        )

    formation.active = not formation.active

    db.commit()

    return RedirectResponse(
        url="/formations",
        status_code=303,
    )


# =========================================================
# GESTION DES AFFECTATIONS
# =========================================================

@app.get("/affectations")
def page_affectations(
    request: Request,
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    collaborateurs = (
        db.query(Collaborateur)
        .filter(Collaborateur.actif.is_(True))
        .order_by(Collaborateur.code.asc())
        .all()
    )

    formations = (
        db.query(Formation)
        .filter(Formation.active.is_(True))
        .order_by(Formation.titre.asc())
        .all()
    )

    affectations = (
        db.query(Affectation)
        .order_by(Affectation.id.desc())
        .all()
    )

    return templates.TemplateResponse(
        request=request,
        name="affectations.html",
        context={
            "active_page": "affectations",
            "collaborateurs": collaborateurs,
            "formations": formations,
            "affectations": affectations,
        },
    )


@app.post("/affectations")
def ajouter_affectation(
    request: Request,
    collaborateur_id: int = Form(...),
    formation_id: int = Form(...),
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    affectation_existante = (
        db.query(Affectation)
        .filter(
            Affectation.collaborateur_id
            == collaborateur_id,
            Affectation.formation_id
            == formation_id,
        )
        .first()
    )

    if affectation_existante is None:

        affectation = Affectation(
            collaborateur_id=collaborateur_id,
            formation_id=formation_id,
            statut="À évaluer",
        )

        db.add(affectation)

        db.commit()

    return RedirectResponse(
        url="/affectations",
        status_code=303,
    )


@app.post("/affectations/{affectation_id}/delete")
def supprimer_affectation(
    affectation_id: int,
    request: Request,
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    affectation = db.get(
        Affectation,
        affectation_id,
    )

    if affectation:

        db.delete(affectation)

        db.commit()

    return RedirectResponse(
        url="/affectations",
        status_code=303,
    )


# =========================================================
# CONSULTATION DES ÉVALUATIONS RH
# =========================================================

@app.get("/evaluations")
def page_evaluations(
    request: Request,
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    evaluations = (
        db.query(Evaluation)
        .order_by(Evaluation.date_reponse.desc())
        .all()
    )

    return templates.TemplateResponse(
        request=request,
        name="evaluations.html",
        context={
            "active_page": "evaluations",
            "evaluations": evaluations,
        },
    )


# =========================================================
# INTERFACE COLLABORATEUR
# =========================================================

@app.get("/evaluation/{token}")
def formulaire_evaluation(
    token: str,
    request: Request,
    db: Session = Depends(get_db),
):

    collaborateur = (
        db.query(Collaborateur)
        .filter(
            Collaborateur.token == token,
            Collaborateur.actif.is_(True),
        )
        .first()
    )

    if collaborateur is None:

        raise HTTPException(
            status_code=404,
            detail="Lien invalide ou désactivé",
        )

    affectations = (
        db.query(Affectation)
        .filter(
            Affectation.collaborateur_id
            == collaborateur.id
        )
        .order_by(Affectation.id.asc())
        .all()
    )

    return templates.TemplateResponse(
        request=request,
        name="evaluation_form.html",
        context={
            "collaborateur": collaborateur,
            "affectations": affectations,
        },
    )


@app.post("/evaluation/{token}/{affectation_id}")
def envoyer_evaluation(
    token: str,
    affectation_id: int,
    request: Request,
    contenu: int = Form(...),
    formateur: int = Form(...),
    organisation: int = Form(...),
    application_poste: int = Form(...),
    satisfaction: int = Form(...),
    recommande: str = Form(...),
    observations: str = Form(""),
    points_forts: str = Form(""),
    points_ameliorer: str = Form(""),
    db: Session = Depends(get_db),
):

    collaborateur = (
        db.query(Collaborateur)
        .filter(Collaborateur.token == token)
        .first()
    )

    affectation = db.get(
        Affectation,
        affectation_id,
    )

    if (
        collaborateur is None
        or affectation is None
        or affectation.collaborateur_id
        != collaborateur.id
    ):

        raise HTTPException(
            status_code=404,
            detail="Évaluation introuvable",
        )

    evaluation_existante = (
        db.query(Evaluation)
        .filter(
            Evaluation.affectation_id
            == affectation_id
        )
        .first()
    )

    if evaluation_existante:

        return RedirectResponse(
            url=f"/evaluation/{token}?already=1",
            status_code=303,
        )

    notes = [
        contenu,
        formateur,
        organisation,
        application_poste,
        satisfaction,
    ]

    for note in notes:

        if note < 1 or note > 5:

            raise HTTPException(
                status_code=400,
                detail=(
                    "Les notes doivent être "
                    "comprises entre 1 et 5"
                ),
            )

    evaluation = Evaluation(
        affectation_id=affectation_id,
        contenu=contenu,
        formateur=formateur,
        organisation=organisation,
        application_poste=application_poste,
        satisfaction=satisfaction,
        recommande=recommande,
        observations=observations.strip() or None,
        points_forts=points_forts.strip() or None,
        points_ameliorer=(
            points_ameliorer.strip()
            or None
        ),
    )

    affectation.statut = "Évaluée"

    db.add(evaluation)

    db.commit()

    return RedirectResponse(
        url=f"/evaluation/{token}?success=1",
        status_code=303,
    )


# =========================================================
# EXPORT EXCEL
# =========================================================

@app.get("/export/evaluations.xlsx")
def exporter_evaluations(
    request: Request,
    db: Session = Depends(get_db),
):

    protection = verifier_connexion_rh(request)

    if protection:
        return protection

    classeur = Workbook()

    feuille = classeur.active

    feuille.title = "Évaluations"

    feuille.append(
        [
            "Code collaborateur",
            "Formation",
            "Type de formation",
            "Contenu",
            "Formateur",
            "Organisation",
            "Application au poste",
            "Satisfaction globale",
            "Recommande",
            "Observations",
            "Points forts",
            "Points à améliorer",
            "Date de réponse",
        ]
    )

    evaluations = (
        db.query(Evaluation)
        .order_by(Evaluation.date_reponse.desc())
        .all()
    )

    for evaluation in evaluations:

        affectation = evaluation.affectation

        feuille.append(
            [
                affectation.collaborateur.code,
                affectation.formation.titre,
                affectation.formation.type_formation,
                evaluation.contenu,
                evaluation.formateur,
                evaluation.organisation,
                evaluation.application_poste,
                evaluation.satisfaction,
                evaluation.recommande,
                evaluation.observations or "",
                evaluation.points_forts or "",
                evaluation.points_ameliorer or "",
                evaluation.date_reponse.strftime(
                    "%d/%m/%Y %H:%M"
                ),
            ]
        )

    fichier_excel = io.BytesIO()

    classeur.save(fichier_excel)

    fichier_excel.seek(0)

    headers = {
        "Content-Disposition":
            'attachment; filename="evaluations_formations.xlsx"'
    }

    return StreamingResponse(
        fichier_excel,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers=headers,
    )