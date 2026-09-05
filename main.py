
from __future__ import annotations
from datetime import datetime, timedelta
from pathlib import Path
from collections import Counter
from email.message import EmailMessage
import base64, hashlib, hmac, io, os, secrets, smtplib, string

from fastapi import FastAPI, Request, Form, Depends, HTTPException
from fastapi.responses import RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from sqlalchemy import create_engine, Column, Integer, String, Boolean, DateTime, ForeignKey, Text, UniqueConstraint, func, Float, text, inspect
from sqlalchemy.orm import declarative_base, sessionmaker, Session, relationship
from sqlalchemy.exc import IntegrityError
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
except ImportError:
    colors = None

BASE_DIR = Path(__file__).resolve().parent
SQLITE_URL = f"sqlite:///{BASE_DIR / 'evaluation.db'}"
DATABASE_URL = os.getenv("DATABASE_URL", SQLITE_URL)
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)
kwargs = {"pool_pre_ping": True}
if DATABASE_URL.startswith("sqlite"):
    kwargs["connect_args"] = {"check_same_thread": False}
engine = create_engine(DATABASE_URL, **kwargs)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()

class User(Base):
    __tablename__="app_users"
    id=Column(Integer,primary_key=True)
    fullname=Column(String(180),nullable=False)
    username=Column(String(80),unique=True,nullable=False,index=True)
    email=Column(String(180))
    password_hash=Column(String(255),nullable=False)
    role=Column(String(20),default="RH",nullable=False)
    active=Column(Boolean,default=True)
    must_change_password=Column(Boolean,default=True)
    created_at=Column(DateTime,default=datetime.utcnow)

class Departement(Base):
    __tablename__="departements"
    id=Column(Integer,primary_key=True)
    nom=Column(String(180),unique=True,nullable=False)
    actif=Column(Boolean,default=True)

class Collaborateur(Base):
    __tablename__="collaborateurs"
    id=Column(Integer,primary_key=True)
    code=Column(String(30),unique=True,nullable=False,index=True)
    nom=Column(String(120),nullable=False)
    prenom=Column(String(120),nullable=False)
    email=Column(String(180),unique=True,nullable=False)
    token=Column(String(96),nullable=False,default=lambda: secrets.token_hex(24))
    password_hash=Column(String(255),nullable=True)
    must_change_password=Column(Boolean,default=True,nullable=False)
    fonction=Column(String(180))
    departement_id=Column(Integer,ForeignKey("departements.id"))
    actif=Column(Boolean,default=True)
    a_quitte=Column(Boolean,default=False,nullable=False)
    date_depart=Column(DateTime,nullable=True)
    motif_depart=Column(String(80),nullable=True)
    created_at=Column(DateTime,default=datetime.utcnow)
    departement=relationship("Departement")
    affectations=relationship("Affectation",back_populates="collaborateur",cascade="all, delete-orphan")

class Categorie(Base):
    __tablename__="categories"
    id=Column(Integer,primary_key=True)
    nom=Column(String(120),unique=True,nullable=False)
    description=Column(Text)
    actif=Column(Boolean,default=True)

class Questionnaire(Base):
    __tablename__="questionnaires"
    id=Column(Integer,primary_key=True)
    titre=Column(String(255),nullable=False)
    categorie_id=Column(Integer,ForeignKey("categories.id"),nullable=False)
    description=Column(Text)
    actif=Column(Boolean,default=True)
    categorie=relationship("Categorie")
    questions=relationship("Question",back_populates="questionnaire",cascade="all, delete-orphan",order_by="Question.ordre")

class Question(Base):
    __tablename__="questions"
    id=Column(Integer,primary_key=True)
    questionnaire_id=Column(Integer,ForeignKey("questionnaires.id"),nullable=False)
    libelle=Column(Text,nullable=False)
    type_reponse=Column(String(30),default="note5",nullable=False)
    options=Column(Text)
    obligatoire=Column(Boolean,default=True)
    ordre=Column(Integer,default=1)
    questionnaire=relationship("Questionnaire",back_populates="questions")

class SessionFormation(Base):
    __tablename__="sessions"
    id=Column(Integer,primary_key=True)
    titre=Column(String(255),nullable=False)
    categorie_id=Column(Integer,ForeignKey("categories.id"),nullable=False)
    questionnaire_id=Column(Integer,ForeignKey("questionnaires.id"),nullable=False)
    date_debut=Column(String(20))
    date_fin=Column(String(20))
    lieu=Column(String(255))
    animateur=Column(String(255))
    description=Column(Text)
    type_formation=Column(String(50),nullable=True)
    active=Column(Boolean,default=True)
    categorie=relationship("Categorie")
    questionnaire=relationship("Questionnaire")
    affectations=relationship("Affectation",back_populates="session",cascade="all, delete-orphan")

class Affectation(Base):
    __tablename__="affectations"
    __table_args__=(UniqueConstraint("collaborateur_id","session_id",name="uq_collab_session"),)
    id=Column(Integer,primary_key=True)
    collaborateur_id=Column(Integer,ForeignKey("collaborateurs.id"),nullable=False)
    session_id=Column(Integer,ForeignKey("sessions.id"),nullable=False)
    statut=Column(String(30),default="À évaluer")
    invitation_envoyee=Column(Boolean,default=False)
    beneficie=Column(Boolean,nullable=True)
    created_at=Column(DateTime,default=datetime.utcnow)
    collaborateur=relationship("Collaborateur",back_populates="affectations")
    session=relationship("SessionFormation",back_populates="affectations")
    reponse=relationship("Reponse",back_populates="affectation",uselist=False,cascade="all, delete-orphan")

class Reponse(Base):
    __tablename__="reponses"
    id=Column(Integer,primary_key=True)
    affectation_id=Column(Integer,ForeignKey("affectations.id"),unique=True,nullable=False)
    satisfaction_globale=Column(Float,default=0)
    recommande=Column(Boolean,default=False)
    commentaire=Column(Text)
    date_reponse=Column(DateTime,default=datetime.utcnow)
    affectation=relationship("Affectation",back_populates="reponse")
    details=relationship("ReponseDetail",back_populates="reponse",cascade="all, delete-orphan")

class ReponseDetail(Base):
    __tablename__="reponse_details"
    id=Column(Integer,primary_key=True)
    reponse_id=Column(Integer,ForeignKey("reponses.id"),nullable=False)
    question_id=Column(Integer,ForeignKey("questions.id"),nullable=False)
    valeur=Column(Text)
    reponse=relationship("Reponse",back_populates="details")
    question=relationship("Question")

class OTPCode(Base):
    __tablename__="otp_codes"
    id=Column(Integer,primary_key=True)
    email=Column(String(180),nullable=False,index=True)
    code_hash=Column(String(255),nullable=False)
    expires_at=Column(DateTime,nullable=False)
    used=Column(Boolean,default=False)
    created_at=Column(DateTime,default=datetime.utcnow)

class AuditLog(Base):
    __tablename__="audit_logs"
    id=Column(Integer,primary_key=True)
    acteur=Column(String(180))
    action=Column(String(255),nullable=False)
    details=Column(Text)
    created_at=Column(DateTime,default=datetime.utcnow)

Base.metadata.create_all(bind=engine)

def ensure_runtime_schema():
    """Ajoute les colonnes nécessaires sans supprimer les données existantes."""
    with engine.begin() as conn:
        if engine.url.get_backend_name() == "postgresql":
            conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN IF NOT EXISTS password_hash VARCHAR(255)"))
            conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN IF NOT EXISTS must_change_password BOOLEAN NOT NULL DEFAULT TRUE"))
            conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN IF NOT EXISTS a_quitte BOOLEAN NOT NULL DEFAULT FALSE"))
            conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN IF NOT EXISTS date_depart TIMESTAMP"))
            conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN IF NOT EXISTS motif_depart VARCHAR(80)"))
            conn.execute(text("ALTER TABLE affectations ADD COLUMN IF NOT EXISTS beneficie BOOLEAN"))
            conn.execute(text("ALTER TABLE sessions ADD COLUMN IF NOT EXISTS type_formation VARCHAR(50)"))
        else:
            cols={row[1] for row in conn.execute(text("PRAGMA table_info(collaborateurs)"))}
            if "password_hash" not in cols:
                conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN password_hash VARCHAR(255)"))
            if "must_change_password" not in cols:
                conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN must_change_password BOOLEAN NOT NULL DEFAULT 1"))
            if "a_quitte" not in cols:
                conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN a_quitte BOOLEAN NOT NULL DEFAULT 0"))
            if "date_depart" not in cols:
                conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN date_depart DATETIME"))
            if "motif_depart" not in cols:
                conn.execute(text("ALTER TABLE collaborateurs ADD COLUMN motif_depart VARCHAR(80)"))
            affect_cols={row[1] for row in conn.execute(text("PRAGMA table_info(affectations)"))}
            if "beneficie" not in affect_cols:
                conn.execute(text("ALTER TABLE affectations ADD COLUMN beneficie BOOLEAN"))
            session_cols={row[1] for row in conn.execute(text("PRAGMA table_info(sessions)"))}
            if "type_formation" not in session_cols:
                conn.execute(text("ALTER TABLE sessions ADD COLUMN type_formation VARCHAR(50)"))

ensure_runtime_schema()

app=FastAPI(title="FIRSI - UM6P",version="5.0")
app.add_middleware(SessionMiddleware,secret_key=os.getenv("SECRET_KEY","firsi-local-change-me"),same_site="lax",https_only=os.getenv("ENVIRONMENT")=="production",max_age=28800)
app.mount("/static",StaticFiles(directory=BASE_DIR/"app"/"static"),name="static")
templates=Jinja2Templates(directory=BASE_DIR/"app"/"templates")

def get_db():
    db=SessionLocal()
    try: yield db
    finally: db.close()

def hash_password(password:str)->str:
    salt=secrets.token_bytes(16); it=390000
    digest=hashlib.pbkdf2_hmac("sha256",password.encode(),salt,it)
    return f"pbkdf2_sha256${it}${base64.urlsafe_b64encode(salt).decode()}${base64.urlsafe_b64encode(digest).decode()}"

def verify_password(password:str,stored:str)->bool:
    try:
        _,it,salt_b64,digest_b64=stored.split("$",3)
        actual=hashlib.pbkdf2_hmac("sha256",password.encode(),base64.urlsafe_b64decode(salt_b64),int(it))
        return hmac.compare_digest(actual,base64.urlsafe_b64decode(digest_b64))
    except Exception:return False

def require_staff(request):
    return request.session.get("role") in {"ADMIN","RH"}

def require_admin(request):
    return request.session.get("role")=="ADMIN"

def audit(db,request,action,details=""):
    db.add(AuditLog(acteur=request.session.get("name","système"),action=action,details=details));db.commit()

def delete_legacy_evaluations(db:Session,affectation_ids:list[int]|None=None):
    """Supprime aussi les évaluations de l’ancienne V4 si la table existe encore."""
    if not inspect(engine).has_table("evaluations"):
        return
    if affectation_ids:
        for aid in affectation_ids:
            db.execute(text("DELETE FROM evaluations WHERE affectation_id = :aid"),{"aid":aid})
    else:
        db.execute(text("DELETE FROM evaluations"))

def send_email(to_email,subject,body):
    """Envoie un e-mail SMTP. Retourne (succès, message utilisateur)."""
    host=os.getenv("SMTP_HOST");user=os.getenv("SMTP_USERNAME");pwd=os.getenv("SMTP_PASSWORD")
    port=int(os.getenv("SMTP_PORT","587"));sender=os.getenv("SMTP_FROM",user or "")
    if not all([host,user,pwd,sender]):
        print(f"[EMAIL NON CONFIGURÉ] {to_email} | {subject}\n{body}",flush=True)
        return False,"La messagerie SMTP n’est pas encore configurée sur Render."
    try:
        msg=EmailMessage();msg["From"]=sender;msg["To"]=to_email;msg["Subject"]=subject;msg.set_content(body)
        with smtplib.SMTP(host,port,timeout=25) as server:
            server.ehlo();server.starttls();server.ehlo();server.login(user,pwd);server.send_message(msg)
        return True,"E-mail envoyé avec succès."
    except Exception as exc:
        print(f"[ERREUR SMTP] {type(exc).__name__}: {exc}",flush=True)
        return False,"Échec de l’envoi. Vérifiez les paramètres SMTP dans Render."


def public_base_url(request: Request) -> str:
    configured = os.getenv("APP_BASE_URL", "").strip().rstrip("/")
    if configured:
        return configured
    return str(request.base_url).rstrip("/")

def invitation_link(request: Request, affectation_id: int) -> str:
    db=SessionLocal()
    try:
        affectation=db.get(Affectation,affectation_id)
        if affectation and affectation.collaborateur and affectation.collaborateur.token:
            return f"{public_base_url(request)}/acces/{affectation.collaborateur.token}"
        return public_base_url(request)
    finally:
        db.close()


def seed():
    db=SessionLocal()
    try:
        if not db.query(User).filter(User.username==os.getenv("ADMIN_USERNAME","admin")).first():
            db.add(User(fullname="Administrateur principal",username=os.getenv("ADMIN_USERNAME","admin"),email=os.getenv("ADMIN_EMAIL"),password_hash=hash_password(os.getenv("ADMIN_PASSWORD","Admin@2026")),role="ADMIN"))
        for d in ["Programme & Bourse","Ressources Humaines","Finance","Communication"]:
            if not db.query(Departement).filter_by(nom=d).first():db.add(Departement(nom=d))
        for nom,desc in {
            "Formation":"Développement des compétences.",
            "Team Building":"Cohésion et collaboration.",
            "Ftour Ramadan":"Convivialité et événements Ramadan."
        }.items():
            if not db.query(Categorie).filter_by(nom=nom).first():db.add(Categorie(nom=nom,description=desc))
        db.commit()
        defaults={
            "Formation":[("Le contenu répond-il à vos besoins professionnels ?","note5",None),("La qualité de l'animation était-elle satisfaisante ?","note5",None),("Cette formation est-elle applicable dans votre poste ?","note5",None),("Quelle formation souhaitez-vous suivre prochainement ?","texte",None),("Recommanderiez-vous cette formation ?","oui_non",None)],
            "Team Building":[("L'activité a-t-elle renforcé la cohésion ?","note5",None),("L'organisation était-elle satisfaisante ?","note5",None),("Quel thème de Team Building proposez-vous ?","texte",None),("Quel format préférez-vous ?","choix","Sport;Cuisine;Culture;Créativité;Nature"),("Recommanderiez-vous cette activité ?","oui_non",None)],
            "Ftour Ramadan":[("Comment évaluez-vous l'organisation du Ftour ?","note5",None),("Le lieu et l'accueil étaient-ils satisfaisants ?","note5",None),("Quel thème proposez-vous pour le prochain Ftour ?","texte",None),("Quelles animations souhaiteriez-vous ?","texte",None),("Recommanderiez-vous cet événement ?","oui_non",None)]
        }
        for cat_name,questions in defaults.items():
            cat=db.query(Categorie).filter_by(nom=cat_name).first();title=f"Questionnaire standard - {cat_name}"
            if not db.query(Questionnaire).filter_by(titre=title).first():
                qn=Questionnaire(titre=title,categorie_id=cat.id,description=f"Questionnaire par défaut {cat_name}");db.add(qn);db.flush()
                for i,(lib,t,opt) in enumerate(questions,1):db.add(Question(questionnaire_id=qn.id,libelle=lib,type_reponse=t,options=opt,ordre=i))
        db.commit()
        formation_cat=db.query(Categorie).filter(func.lower(Categorie.nom)=="formation").first()
        if formation_cat:
            for qn in db.query(Questionnaire).filter(Questionnaire.categorie_id==formation_cat.id).all():
                labels=[(q.libelle or "").lower() for q in qn.questions]
                if not any("type de cette formation" in x for x in labels):
                    maxo=max([q.ordre for q in qn.questions] or [0]); db.add(Question(questionnaire_id=qn.id,libelle="Quel est le type de cette formation ?",type_reponse="choix",options="Certifiante;Diplômante;Qualifiante",obligatoire=True,ordre=maxo+1))
                if not any("pourquoi ne recommanderiez-vous pas" in x for x in labels):
                    maxo=max([q.ordre for q in qn.questions] or [0])+1; db.add(Question(questionnaire_id=qn.id,libelle="Si non, pourquoi ne recommanderiez-vous pas cette formation ?",type_reponse="paragraphe",obligatoire=False,ordre=maxo))
        db.commit()
    finally:db.close()
seed()

@app.get("/")
def landing(request:Request):
    return templates.TemplateResponse("landing.html",{"request":request,"error":request.query_params.get("error"),"email":request.query_params.get("email","")})

@app.post("/auth/staff")
def staff_login(request:Request,username:str=Form(...),password:str=Form(...),db:Session=Depends(get_db)):
    u=db.query(User).filter(User.username==username.strip(),User.active.is_(True)).first()
    if not u or not verify_password(password,u.password_hash):return RedirectResponse("/?error=staff",303)
    request.session.clear();request.session.update({"user_id":u.id,"role":u.role,"name":u.fullname})
    return RedirectResponse("/change-password" if u.must_change_password else "/dashboard",303)


@app.post("/auth/collaborateur")
def collaborator_login(
    request:Request,
    email:str=Form(...),
    password:str=Form(...),
    db:Session=Depends(get_db),
):
    clean_email=email.strip().lower()
    collaborateur=db.query(Collaborateur).filter(
        func.lower(Collaborateur.email)==clean_email,
        Collaborateur.actif.is_(True)
    ).first()
    if not collaborateur or not collaborateur.password_hash or not verify_password(password,collaborateur.password_hash):
        return RedirectResponse("/?error=collaborateur",303)

    request.session.clear()
    request.session.update({
        "role":"COLLABORATEUR",
        "collaborateur_id":collaborateur.id,
        "name":f"{collaborateur.prenom} {collaborateur.nom}",
    })
    if collaborateur.must_change_password:
        return RedirectResponse("/collaborateur/change-password",303)
    return RedirectResponse("/mon-espace",303)

@app.get("/collaborateur/change-password")
def collaborator_change_password_page(request:Request):
    if request.session.get("role")!="COLLABORATEUR":
        return RedirectResponse("/",303)
    return templates.TemplateResponse("collaborator_change_password.html",{
        "request":request,
        "error":request.query_params.get("error"),
    })

@app.post("/collaborateur/change-password")
def collaborator_change_password(
    request:Request,
    new_password:str=Form(...),
    confirm_password:str=Form(...),
    db:Session=Depends(get_db),
):
    if request.session.get("role")!="COLLABORATEUR":
        return RedirectResponse("/",303)
    if len(new_password)<8 or new_password!=confirm_password:
        return RedirectResponse("/collaborateur/change-password?error=1",303)
    collaborateur=db.get(Collaborateur,request.session.get("collaborateur_id"))
    if not collaborateur:
        request.session.clear()
        return RedirectResponse("/",303)
    collaborateur.password_hash=hash_password(new_password)
    collaborateur.must_change_password=False
    db.commit()
    return RedirectResponse("/mon-espace?password=changed",303)

@app.get("/change-password")
def change_password_page(request:Request):
    if not require_staff(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("change_password.html",{"request":request})

@app.post("/change-password")
def change_password(request:Request,new_password:str=Form(...),confirm_password:str=Form(...),db:Session=Depends(get_db)):
    if len(new_password)<8 or new_password!=confirm_password:return RedirectResponse("/change-password?error=1",303)
    u=db.get(User,request.session.get("user_id"));u.password_hash=hash_password(new_password);u.must_change_password=False;db.commit();return RedirectResponse("/dashboard",303)

@app.get("/logout")
def logout(request:Request):request.session.clear();return RedirectResponse("/",303)

FORMATION_TYPES=("Certifiante","Diplômante","Qualifiante")

def _is_training_type_question(question):
    label=(question.libelle or "").lower()
    return "type" in label and "formation" in label

def _resolved_training_type_for_response(response):
    """Type retenu pour une évaluation : paramètre session prioritaire, sinon réponse collaborateur."""
    fixed=(response.affectation.session.type_formation or "").strip()
    if fixed in FORMATION_TYPES:
        return fixed, "Session"
    for detail in response.details:
        if _is_training_type_question(detail.question):
            value=(detail.valeur or "").strip()
            if value in FORMATION_TYPES:
                return value, "Collaborateur"
    return None, "Non renseigné"

def _formation_type_stats(db):
    counts={t:0 for t in FORMATION_TYPES}
    sources={"Session":0,"Collaborateur":0,"Non renseigné":0}
    responses=(db.query(Reponse).join(Affectation,Reponse.affectation_id==Affectation.id)
               .join(SessionFormation,Affectation.session_id==SessionFormation.id)
               .join(Categorie,SessionFormation.categorie_id==Categorie.id)
               .filter(func.lower(Categorie.nom)=="formation").all())
    for r in responses:
        typ,source=_resolved_training_type_for_response(r)
        sources[source]=sources.get(source,0)+1
        if typ: counts[typ]+=1
    total=sum(counts.values())
    items=[{"type":t,"count":counts[t],"pct":round(counts[t]/total*100,1) if total else 0} for t in FORMATION_TYPES]
    return items,sources,total

@app.get("/dashboard")
def dashboard(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)

    total_collaborateurs=db.query(Collaborateur).filter(Collaborateur.actif.is_(True)).count()
    total_departures=db.query(Collaborateur).filter(Collaborateur.a_quitte.is_(True)).count()
    total_aff=db.query(Affectation).count()
    total_rep=db.query(Reponse).count()
    pending=max(
        db.query(Affectation).filter(Affectation.beneficie.is_(True)).count()-total_rep,
        0
    )
    satisfaction=db.query(func.avg(Reponse.satisfaction_globale)).scalar() or 0

    # Taux de couverture des formations :
    # collaborateurs actifs ayant au moins une affectation à une session
    # de catégorie "Formation" / collaborateurs actifs.
    collaborateurs_formes=(
        db.query(func.count(func.distinct(Affectation.collaborateur_id)))
        .join(SessionFormation,Affectation.session_id==SessionFormation.id)
        .join(Categorie,SessionFormation.categorie_id==Categorie.id)
        .join(Collaborateur,Affectation.collaborateur_id==Collaborateur.id)
        .filter(
            func.lower(Categorie.nom)=="formation",
            Collaborateur.actif.is_(True),
            Affectation.beneficie.is_(True),
        )
        .scalar()
    ) or 0
    taux_couverture=round(
        collaborateurs_formes/total_collaborateurs*100,1
    ) if total_collaborateurs else 0

    # Taux de participation aux évaluations :
    # réponses reçues / bénéficiaires confirmés.
    total_beneficiaires=db.query(Affectation).filter(Affectation.beneficie.is_(True)).count()
    total_non_beneficiaires=db.query(Affectation).filter(Affectation.beneficie.is_(False)).count()
    total_confirmation_attente=db.query(Affectation).filter(Affectation.beneficie.is_(None)).count()
    taux_participation=round(total_rep/total_beneficiaires*100,1) if total_beneficiaires else 0

    # Questions, réponses reçues et réponses en attente par thématique
    themes=[]
    categories=db.query(Categorie).order_by(Categorie.nom).all()
    for cat in categories:
        question_count=(
            db.query(Question)
            .join(Questionnaire,Question.questionnaire_id==Questionnaire.id)
            .filter(Questionnaire.categorie_id==cat.id)
            .count()
        )
        received=(
            db.query(Reponse)
            .join(Affectation,Reponse.affectation_id==Affectation.id)
            .join(SessionFormation,Affectation.session_id==SessionFormation.id)
            .filter(SessionFormation.categorie_id==cat.id)
            .count()
        )
        assigned=(
            db.query(Affectation)
            .join(SessionFormation,Affectation.session_id==SessionFormation.id)
            .filter(SessionFormation.categorie_id==cat.id,Affectation.beneficie.is_(True))
            .count()
        )
        avg_score=(
            db.query(func.avg(Reponse.satisfaction_globale))
            .join(Affectation,Reponse.affectation_id==Affectation.id)
            .join(SessionFormation,Affectation.session_id==SessionFormation.id)
            .filter(SessionFormation.categorie_id==cat.id)
            .scalar()
        ) or 0
        themes.append({
            "nom":cat.nom,
            "questions":question_count,
            "recues":received,
            "attente":max(assigned-received,0),
            "satisfaction":round(float(avg_score),2),
        })

    # Répartition par type basée sur les évaluations : type fixé sur la session
    # si disponible, sinon choix du collaborateur dans le questionnaire.
    formation_types, formation_type_sources, total_typed_evaluations = _formation_type_stats(db)

    deps=db.query(Departement.nom,func.count(Collaborateur.id)).outerjoin(Collaborateur).group_by(Departement.nom).all()

    top_sessions=(
        db.query(SessionFormation.titre,func.avg(Reponse.satisfaction_globale).label("score"),func.count(Reponse.id).label("nb"))
        .join(Affectation,Affectation.session_id==SessionFormation.id)
        .join(Reponse,Reponse.affectation_id==Affectation.id)
        .group_by(SessionFormation.id,SessionFormation.titre)
        .order_by(func.avg(Reponse.satisfaction_globale).desc())
        .limit(5).all()
    )

    recent_pending=(
        db.query(Affectation)
        .filter(Affectation.beneficie.is_(True),Affectation.statut!="Évaluée")
        .order_by(Affectation.created_at.desc())
        .limit(5).all()
    )

    return templates.TemplateResponse("dashboard.html",{
        "request":request,
        "kpis":{
            "collaborateurs":total_collaborateurs,
            "collaborateurs_formes":collaborateurs_formes,
            "sessions":db.query(SessionFormation).filter(SessionFormation.active.is_(True)).count(),
            "couverture":taux_couverture,
            "participation":taux_participation,
            "satisfaction":round(float(satisfaction),2),
            "attente":pending,
            "recues":total_rep,
            "affectations":total_aff,
            "beneficiaires":total_beneficiaires,
            "non_beneficiaires":total_non_beneficiaires,
            "confirmation_attente":total_confirmation_attente,
            "non_formes":max(total_collaborateurs-collaborateurs_formes,0),
            "formations":db.query(SessionFormation).join(Categorie,SessionFormation.categorie_id==Categorie.id).filter(SessionFormation.active.is_(True),func.lower(Categorie.nom)=="formation").count(),
            "departures":total_departures,
        },
        "themes":themes,
        "formation_types":formation_types,
        "formation_type_sources":formation_type_sources,
        "total_typed_evaluations":total_typed_evaluations,
        "department_labels":[x[0] for x in deps],
        "department_values":[x[1] for x in deps],
        "top_sessions":top_sessions,
        "recent_pending":recent_pending,
    })

@app.get("/categories")
def categories_page(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("categories.html",{
        "request":request,
        "categories":db.query(Categorie).order_by(Categorie.nom).all(),
        "message":request.query_params.get("message"),
        "error":request.query_params.get("error")
    })

@app.post("/categories")
def create_category(request:Request,nom:str=Form(...),description:str=Form(""),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    clean=nom.strip()
    if not db.query(Categorie).filter(func.lower(Categorie.nom)==clean.lower()).first():
        db.add(Categorie(nom=clean,description=description.strip() or None));db.commit();audit(db,request,"Création catégorie",clean)
        return RedirectResponse("/categories?message=ajoutee",303)
    return RedirectResponse("/categories?error=existe",303)

@app.post("/categories/{cat_id}/delete")
def delete_category(cat_id:int,request:Request,db:Session=Depends(get_db)):
    """Supprime une thématique et ses questionnaires/questions seulement si aucune session ne l'utilise.
    Les évaluations existantes sont ainsi protégées.
    """
    if not require_staff(request):return RedirectResponse("/",303)
    cat=db.get(Categorie,cat_id)
    if not cat:return RedirectResponse("/categories?error=introuvable",303)
    in_use = (db.query(SessionFormation).filter(SessionFormation.categorie_id==cat_id).count()
              + db.query(SessionFormation).join(Questionnaire, SessionFormation.questionnaire_id==Questionnaire.id).filter(Questionnaire.categorie_id==cat_id).count())
    if in_use:
        return RedirectResponse("/categories?error=utilisee",303)
    questionnaires=list(db.query(Questionnaire).filter(Questionnaire.categorie_id==cat_id).all())
    try:
        for qn in questionnaires:
            db.query(Question).filter(Question.questionnaire_id==qn.id).delete(synchronize_session=False)
            db.delete(qn)
        name=cat.nom
        db.delete(cat);db.commit();audit(db,request,"Suppression thématique",f"{name} + questionnaires/questions associés")
        return RedirectResponse("/categories?message=supprimee",303)
    except Exception as exc:
        db.rollback();print(f"[SUPPRESSION CATEGORIE] {exc}",flush=True)
        return RedirectResponse("/categories?error=suppression_impossible",303)

@app.get("/questionnaires")
def questionnaires_page(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("questionnaires.html",{"request":request,"questionnaires":db.query(Questionnaire).order_by(Questionnaire.id.desc()).all(),"categories":db.query(Categorie).filter_by(actif=True).all()})

@app.post("/questionnaires")
def create_questionnaire(request:Request,titre:str=Form(...),categorie_id:int=Form(...),description:str=Form(""),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    q=Questionnaire(titre=titre.strip(),categorie_id=categorie_id,description=description.strip() or None);db.add(q);db.commit();return RedirectResponse(f"/questionnaires/{q.id}",303)

@app.get("/questionnaires/{qid}")
def builder(qid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    q=db.get(Questionnaire,qid)
    if not q:raise HTTPException(404)
    return templates.TemplateResponse("questionnaire_builder.html",{"request":request,"questionnaire":q})

@app.post("/questionnaires/{qid}/questions")
def add_question(qid:int,request:Request,libelle:str=Form(...),type_reponse:str=Form(...),options:str=Form(""),obligatoire:str=Form("oui"),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    if not db.get(Questionnaire,qid): raise HTTPException(404)
    allowed={"note5","note10","oui_non","choix","choix_multiple","texte","paragraphe"}
    if type_reponse not in allowed:return RedirectResponse(f"/questionnaires/{qid}",303)
    if type_reponse in {"choix","choix_multiple"} and not options.strip():return RedirectResponse(f"/questionnaires/{qid}",303)
    ordre=(db.query(func.max(Question.ordre)).filter_by(questionnaire_id=qid).scalar() or 0)+1
    db.add(Question(questionnaire_id=qid,libelle=libelle.strip(),type_reponse=type_reponse,options=options.strip() or None,obligatoire=obligatoire=="oui",ordre=ordre));db.commit();return RedirectResponse(f"/questionnaires/{qid}",303)

@app.post("/questions/{qid}/delete")
def delete_question(qid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    q=db.get(Question,qid);parent=q.questionnaire_id;db.delete(q);db.commit();return RedirectResponse(f"/questionnaires/{parent}",303)

@app.get("/departements")
def departments_page(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("departements.html",{"request":request,"departements":db.query(Departement).all()})

@app.post("/departements")
def create_department(request:Request,nom:str=Form(...),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    if not db.query(Departement).filter_by(nom=nom.strip()).first():db.add(Departement(nom=nom.strip()));db.commit()
    return RedirectResponse("/departements",303)

@app.get("/collaborateurs")
def collabs_page(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("collaborateurs.html",{
        "request":request,
        "collaborateurs":db.query(Collaborateur).order_by(Collaborateur.id.desc()).all(),
        "departements":db.query(Departement).all(),
        "message":request.query_params.get("message"),
        "error":request.query_params.get("error"),
        "base_url":public_base_url(request),
    })

@app.post("/collaborateurs")
def create_collab(request:Request,nom:str=Form(...),prenom:str=Form(...),email:str=Form(...),temporary_password:str=Form(...),fonction:str=Form(""),departement_id:int|None=Form(None),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    clean_email=email.strip().lower()
    if db.query(Collaborateur).filter(func.lower(Collaborateur.email)==clean_email).first():
        return RedirectResponse("/collaborateurs?error=email_existant",303)
    code=f"COL-{(db.query(func.max(Collaborateur.id)).scalar() or 0)+1:04d}"
    try:
        db.add(Collaborateur(
            code=code,
            nom=nom.strip(),
            prenom=prenom.strip(),
            email=clean_email,
            token=secrets.token_hex(24),
            password_hash=hash_password(temporary_password),
            must_change_password=True,
            fonction=fonction.strip() or None,
            departement_id=departement_id
        ))
        db.commit()
        audit(db,request,"Création collaborateur",f"{prenom.strip()} {nom.strip()} ({clean_email})")
        return RedirectResponse("/collaborateurs?message=collaborateur_ajoute",303)
    except IntegrityError:
        db.rollback()
        return RedirectResponse("/collaborateurs?error=donnees_dupliquees",303)

@app.post("/collaborateurs/{cid}/reset-password")
def reset_collaborateur_password(
    cid:int,
    request:Request,
    temporary_password:str=Form(...),
    db:Session=Depends(get_db)
):
    if not require_staff(request):
        return RedirectResponse("/",303)
    c=db.get(Collaborateur,cid)
    if not c:
        return RedirectResponse("/collaborateurs?error=introuvable",303)
    temporary_password=temporary_password.strip()
    if len(temporary_password)<8:
        return RedirectResponse("/collaborateurs?error=mot_de_passe_court",303)
    try:
        c.password_hash=hash_password(temporary_password)
        c.must_change_password=True
        c.token=secrets.token_hex(24)
        db.commit()
        audit(db,request,"Réinitialisation mot de passe collaborateur",f"{c.prenom} {c.nom}")
        return RedirectResponse("/collaborateurs?message=mot_de_passe_reinitialise",303)
    except Exception as exc:
        db.rollback()
        print(f"[RESET PASSWORD COLLABORATEUR] {exc}",flush=True)
        return RedirectResponse("/collaborateurs?error=reinitialisation_impossible",303)

@app.post("/collaborateurs/{cid}/departure")
def mark_departure(cid:int,request:Request,motif_depart:str=Form("Démission"),db:Session=Depends(get_db)):
    if not require_staff(request): return RedirectResponse("/",303)
    c=db.get(Collaborateur,cid)
    if not c: return RedirectResponse("/collaborateurs?error=introuvable",303)
    c.a_quitte=True; c.date_depart=datetime.utcnow(); c.motif_depart=motif_depart.strip() or "Démission"; c.actif=False
    db.commit(); audit(db,request,"Départ collaborateur",f"{c.prenom} {c.nom} — {c.motif_depart}")
    return RedirectResponse("/collaborateurs?message=depart_enregistre",303)

@app.post("/collaborateurs/{cid}/delete")
def delete_collaborateur(cid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    c=db.get(Collaborateur,cid)
    if not c:return RedirectResponse("/collaborateurs?error=introuvable",303)
    label=f"{c.prenom} {c.nom}";email=c.email
    try:
        # Nettoyage explicite pour PostgreSQL et les anciennes contraintes.
        affectations=db.query(Affectation).filter(Affectation.collaborateur_id==cid).all()
        delete_legacy_evaluations(db,[a.id for a in affectations])
        for a in affectations:
            if a.reponse:
                db.query(ReponseDetail).filter(ReponseDetail.reponse_id==a.reponse.id).delete(synchronize_session=False)
                db.delete(a.reponse)
            db.delete(a)
        db.query(OTPCode).filter(OTPCode.email==email).delete(synchronize_session=False)
        if inspect(engine).has_table("collaborator_accounts"):
            db.execute(text("DELETE FROM collaborator_accounts WHERE collaborateur_id = :cid"),{"cid":cid})
        db.delete(c);db.commit()
        audit(db,request,"Suppression collaborateur",label)
        return RedirectResponse("/collaborateurs?message=collaborateur_supprime",303)
    except Exception as exc:
        db.rollback();print(f"[SUPPRESSION COLLABORATEUR] {exc}",flush=True)
        return RedirectResponse("/collaborateurs?error=suppression_impossible",303)

@app.get("/sessions")
def sessions_page(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("sessions.html",{
        "request":request,"sessions":db.query(SessionFormation).order_by(SessionFormation.id.desc()).all(),
        "categories":db.query(Categorie).all(),"questionnaires":db.query(Questionnaire).all(),
        "message":request.query_params.get("message"),"error":request.query_params.get("error")
    })

@app.post("/sessions")
def create_session(request:Request,titre:str=Form(...),categorie_id:int=Form(...),questionnaire_id:int=Form(...),type_formation:str=Form(""),date_debut:str=Form(""),date_fin:str=Form(""),lieu:str=Form(""),animateur:str=Form(""),description:str=Form(""),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    allowed={"Certifiante","Diplômante","Qualifiante"}
    clean_type=type_formation.strip()
    if clean_type not in allowed: clean_type=None
    db.add(SessionFormation(titre=titre.strip(),categorie_id=categorie_id,questionnaire_id=questionnaire_id,type_formation=clean_type,date_debut=date_debut or None,date_fin=date_fin or None,lieu=lieu.strip() or None,animateur=animateur.strip() or None,description=description.strip() or None));db.commit();return RedirectResponse("/sessions?message=session_creee",303)

@app.post("/sessions/{sid}/delete")
def delete_session(sid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    session=db.get(SessionFormation,sid)
    if not session:return RedirectResponse("/sessions?error=introuvable",303)
    try:
        session_affectations=list(session.affectations)
        delete_legacy_evaluations(db,[a.id for a in session_affectations])
        for a in session_affectations:
            if a.reponse:
                db.query(ReponseDetail).filter(ReponseDetail.reponse_id==a.reponse.id).delete(synchronize_session=False)
                db.delete(a.reponse)
            db.delete(a)
        titre=session.titre;db.delete(session);db.commit();audit(db,request,"Suppression session",titre)
        return RedirectResponse("/sessions?message=session_supprimee",303)
    except Exception as exc:
        db.rollback();print(f"[SUPPRESSION SESSION] {exc}",flush=True)
        return RedirectResponse("/sessions?error=suppression_impossible",303)

@app.get("/affectations")
def assignments_page(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("affectations.html",{
        "request":request,"affectations":db.query(Affectation).order_by(Affectation.id.desc()).all(),
        "collaborateurs":db.query(Collaborateur).filter_by(actif=True).all(),
        "sessions":db.query(SessionFormation).filter_by(active=True).all(),
        "mail_status":request.query_params.get("mail_status"),
        "invite_link":request.query_params.get("invite_link"),
        "message":request.query_params.get("message"),"error":request.query_params.get("error")
    })

@app.post("/affectations")
def create_assignment(request:Request,collaborateur_id:int=Form(...),session_id:int=Form(...),envoyer_invitation:str=Form("non"),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    if not db.query(Affectation).filter_by(collaborateur_id=collaborateur_id,session_id=session_id).first():
        a=Affectation(collaborateur_id=collaborateur_id,session_id=session_id,statut="Confirmation requise",beneficie=None);db.add(a);db.commit();db.refresh(a)
        if envoyer_invitation=="oui":
            link=invitation_link(request,a.id)
            sent,_=send_email(
                a.collaborateur.email,
                "Invitation FIRSI - UM6P",
                f"Bonjour {a.collaborateur.prenom},\n\n"
                f"Merci d’évaluer : {a.session.titre}.\n"
                f"Accédez à la plateforme : {link}\n"
                f"Connectez-vous avec votre adresse e-mail et le mot de passe temporaire communiqué par la RH."
            )
            a.invitation_envoyee=sent;db.commit()
            return RedirectResponse(
                f"/affectations?mail_status={'sent' if sent else 'link'}&invite_link={link}",
                303
            )
    return RedirectResponse("/affectations?message=affectation_ajoutee",303)

@app.post("/affectations/bulk")
def bulk_assign(request:Request,session_id:int=Form(...),collaborateur_ids:list[int]=Form(...),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    session=db.get(SessionFormation,session_id)
    if not session:return RedirectResponse("/affectations?error=session_introuvable",303)
    created=0
    for cid in collaborateur_ids:
        c=db.get(Collaborateur,cid)
        if not c or not c.actif:continue
        if db.query(Affectation).filter_by(collaborateur_id=cid,session_id=session_id).first():continue
        db.add(Affectation(collaborateur_id=cid,session_id=session_id,statut="Confirmation requise",beneficie=None));created+=1
    db.commit();audit(db,request,"Affectation en masse",f"{created} collaborateur(s) - {session.titre}")
    return RedirectResponse(f"/affectations?message=affectation_masse&created={created}",303)

@app.post("/affectations/{aid}/invite")
def invite(aid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    a=db.get(Affectation,aid)
    if not a:return RedirectResponse("/affectations?error=introuvable",303)
    link=invitation_link(request,a.id)
    sent,_=send_email(
        a.collaborateur.email,
        "Rappel d’évaluation FIRSI - UM6P",
        f"Bonjour {a.collaborateur.prenom},\n\n"
        f"Merci de compléter l’évaluation de : {a.session.titre}.\n"
        f"Accédez à la plateforme : {link}\n"
        f"Connectez-vous avec votre adresse e-mail et le mot de passe temporaire communiqué par la RH."
    )
    a.invitation_envoyee=sent;db.commit()
    return RedirectResponse(
        f"/affectations?mail_status={'sent' if sent else 'link'}&invite_link={link}",
        303
    )

@app.post("/affectations/{aid}/delete")
def delete_affectation(aid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    a=db.get(Affectation,aid)
    if not a:return RedirectResponse("/affectations?error=introuvable",303)
    try:
        delete_legacy_evaluations(db,[a.id])
        if a.reponse:
            db.query(ReponseDetail).filter(ReponseDetail.reponse_id==a.reponse.id).delete(synchronize_session=False)
            db.delete(a.reponse)
        db.delete(a);db.commit()
        return RedirectResponse("/affectations?message=affectation_supprimee",303)
    except Exception as exc:
        db.rollback();print(f"[SUPPRESSION AFFECTATION] {exc}",flush=True)
        return RedirectResponse("/affectations?error=suppression_impossible",303)


@app.get("/acces/{token}")
def access_by_token(token:str,request:Request,db:Session=Depends(get_db)):
    collaborateur=db.query(Collaborateur).filter(Collaborateur.token==token,Collaborateur.actif.is_(True)).first()
    if not collaborateur:return RedirectResponse("/?error=lien_invalide",303)
    from urllib.parse import quote
    return RedirectResponse(f"/?email={quote(collaborateur.email)}",303)

@app.get("/mon-espace")
def collaborator_space(request:Request,db:Session=Depends(get_db)):
    cid=request.session.get("collaborateur_id")
    if request.session.get("role")!="COLLABORATEUR" or not cid:return RedirectResponse("/",303)
    c=db.get(Collaborateur,cid);return templates.TemplateResponse("collaborator_space.html",{"request":request,"collaborateur":c,"affectations":c.affectations})

@app.post("/mon-espace/affectations/{aid}/participation")
def confirm_participation(aid:int,request:Request,beneficie:str=Form(...),db:Session=Depends(get_db)):
    cid=request.session.get("collaborateur_id")
    if request.session.get("role")!="COLLABORATEUR" or not cid:return RedirectResponse("/",303)
    a=db.query(Affectation).filter(Affectation.id==aid,Affectation.collaborateur_id==cid).first()
    if not a:return RedirectResponse("/mon-espace?error=affectation_introuvable",303)
    if a.reponse:return RedirectResponse("/mon-espace?already=1",303)
    choix=beneficie.strip().lower()
    if choix=="non":a.beneficie=False;a.statut="Non bénéficié";db.commit();return RedirectResponse("/mon-espace?participation=non",303)
    if choix=="oui":a.beneficie=True;a.statut="À évaluer";db.commit();return RedirectResponse(f"/mon-espace/evaluer/{a.id}",303)
    return RedirectResponse("/mon-espace?error=choix_invalide",303)

@app.get("/mon-espace/evaluer/{aid}")
def evaluation_form(aid:int,request:Request,db:Session=Depends(get_db)):
    a=db.get(Affectation,aid)
    if request.session.get("role")!="COLLABORATEUR" or not a or a.collaborateur_id!=request.session.get("collaborateur_id"):raise HTTPException(403)
    if a.reponse:return RedirectResponse("/mon-espace?already=1",303)
    if a.beneficie is not True:return RedirectResponse("/mon-espace?error=participation_requise",303)
    return templates.TemplateResponse("evaluation_form.html",{"request":request,"affectation":a,"questions":a.session.questionnaire.questions,"fixed_type_formation":a.session.type_formation,"error":request.query_params.get("error")})

@app.post("/mon-espace/evaluer/{aid}")
async def submit_evaluation(aid:int,request:Request,db:Session=Depends(get_db)):
    a=db.get(Affectation,aid)
    if request.session.get("role")!="COLLABORATEUR" or not a or a.collaborateur_id!=request.session.get("collaborateur_id"):raise HTTPException(403)
    if a.reponse:return RedirectResponse("/mon-espace?already=1",303)
    if a.beneficie is not True:return RedirectResponse("/mon-espace?error=participation_requise",303)
    form=await request.form();scores=[];comments=[];recommend=False;values=[];recommendation_negative=False;recommendation_reason=""
    for q in a.session.questionnaire.questions:
        vlist=form.getlist(f"q_{q.id}")
        v=", ".join(str(x).strip() for x in vlist if str(x).strip()) if q.type_reponse=="choix_multiple" else str(form.get(f"q_{q.id}","")).strip(); label=(q.libelle or "").lower()
        # Si le type a été défini par RH/Admin sur la session, il est la valeur de référence.
        if _is_training_type_question(q) and a.session.type_formation in FORMATION_TYPES:
            v=a.session.type_formation
        if q.obligatoire and not v:return RedirectResponse(f"/mon-espace/evaluer/{aid}?error=required",303)
        if q.type_reponse in {"note5","note10"} and v:
            try:scores.append(float(v))
            except:pass
        if q.type_reponse=="oui_non" and ("recommander" in label or "suggér" in label or "suggere" in label):
            recommend=(v.lower()=="oui"); recommendation_negative=(v.lower()=="non")
        if "pourquoi ne recommanderiez-vous pas" in label: recommendation_reason=v
        if q.type_reponse in {"texte","paragraphe"} and v:comments.append(v)
        values.append((q,v))
    if recommendation_negative and not recommendation_reason:return RedirectResponse(f"/mon-espace/evaluer/{aid}?error=recommendation_reason",303)
    r=Reponse(affectation_id=aid,satisfaction_globale=round(sum(scores)/len(scores),2) if scores else 0,recommande=recommend,commentaire="\n".join(comments) or None);db.add(r);db.flush()
    for q,v in values:db.add(ReponseDetail(reponse_id=r.id,question_id=q.id,valeur=v))
    a.statut="Évaluée";db.commit();return RedirectResponse("/mon-espace?success=1",303)

@app.get("/evaluations")
def evaluations_page(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("evaluations.html",{"request":request,"reponses":db.query(Reponse).order_by(Reponse.date_reponse.desc()).all()})

@app.get("/analyse")
def analysis_page(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    comments=[r.commentaire for r in db.query(Reponse).filter(Reponse.commentaire.isnot(None)).all() if r.commentaire]
    tokens=[]
    for c in comments:tokens.extend("".join(ch.lower() if ch.isalnum() else " " for ch in c).split())
    freq=Counter(w for w in tokens if len(w)>4)
    return templates.TemplateResponse("analyse.html",{"request":request,"count":len(comments),"themes":freq.most_common(12)})

@app.get("/export/evaluations.xlsx")
def export_xlsx(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    wb=Workbook()
    # Synthèse
    ws=wb.active;ws.title="Synthèse"
    total_rep=db.query(Reponse).count(); avg=db.query(func.avg(Reponse.satisfaction_globale)).scalar() or 0
    total_coll=db.query(Collaborateur).filter(Collaborateur.actif.is_(True)).count()
    total_ben=db.query(Affectation).filter(Affectation.beneficie.is_(True)).count()
    total_form=db.query(SessionFormation).join(Categorie,SessionFormation.categorie_id==Categorie.id).filter(SessionFormation.active.is_(True),func.lower(Categorie.nom)=="formation").count()
    ws.append(["FIRSI – UM6P | Export des données",""]);ws.append(["Indicateur","Valeur"])
    for row in [("Collaborateurs actifs",total_coll),("Sessions actives",db.query(SessionFormation).filter_by(active=True).count()),("Sessions Formation",total_form),("Bénéficiaires confirmés",total_ben),("Évaluations reçues",total_rep),("Satisfaction moyenne /5",round(float(avg),2))]:ws.append(row)
    # Collaborateurs
    ws=wb.create_sheet("Collaborateurs");ws.append(["Code","Nom","Prénom","Email","Fonction","Département","Actif","A quitté","Date départ","Motif départ"])
    for c in db.query(Collaborateur).order_by(Collaborateur.id):ws.append([c.code,c.nom,c.prenom,c.email,c.fonction or "",c.departement.nom if c.departement else "","Oui" if c.actif else "Non","Oui" if c.a_quitte else "Non",c.date_depart.strftime("%d/%m/%Y") if c.date_depart else "",c.motif_depart or ""])
    # Affectations
    ws=wb.create_sheet("Affectations");ws.append(["Collaborateur","Email","Session","Thématique","Type de formation","Bénéficie","Statut","Invitation envoyée","Date affectation"])
    for a in db.query(Affectation).order_by(Affectation.id):ws.append([f"{a.collaborateur.prenom} {a.collaborateur.nom}",a.collaborateur.email,a.session.titre,a.session.categorie.nom,a.session.type_formation,a.beneficie if a.beneficie is not None else "En attente",a.statut,"Oui" if a.invitation_envoyee else "Non",a.created_at.strftime("%d/%m/%Y")])
    # Évaluations
    ws=wb.create_sheet("Évaluations");ws.append(["Collaborateur","Email","Thématique","Session","Type de formation retenu","Source du type","Satisfaction","Recommande","Commentaire","Date"])
    for r in db.query(Reponse).order_by(Reponse.date_reponse.desc()):
        a=r.affectation; resolved_type, type_source=_resolved_training_type_for_response(r);ws.append([f"{a.collaborateur.prenom} {a.collaborateur.nom}",a.collaborateur.email,a.session.categorie.nom,a.session.titre,resolved_type or "",type_source,r.satisfaction_globale,"Oui" if r.recommande else "Non",r.commentaire or "",r.date_reponse.strftime("%d/%m/%Y")])
    # Réponses détaillées
    ws=wb.create_sheet("Réponses détaillées");ws.append(["Collaborateur","Thématique","Session","Question","Type de question","Réponse","Date"])
    for r in db.query(Reponse).order_by(Reponse.date_reponse.desc()):
        a=r.affectation
        for d in r.details:ws.append([f"{a.collaborateur.prenom} {a.collaborateur.nom}",a.session.categorie.nom,a.session.titre,d.question.libelle,d.question.type_reponse,d.valeur or "",r.date_reponse.strftime("%d/%m/%Y")])
    # Sessions
    ws=wb.create_sheet("Sessions");ws.append(["Session","Thématique","Type de formation","Date début","Date fin","Lieu","Animateur","Active"])
    for x in db.query(SessionFormation).order_by(SessionFormation.id):ws.append([x.titre,x.categorie.nom,x.type_formation,x.date_debut or "",x.date_fin or "",x.lieu or "",x.animateur or "","Oui" if x.active else "Non"])
    # Questionnaires / questions
    ws=wb.create_sheet("Questionnaires");ws.append(["Questionnaire","Thématique","Description","Actif"])
    for qn in db.query(Questionnaire).order_by(Questionnaire.id):ws.append([qn.titre,qn.categorie.nom,qn.description or "","Oui" if qn.actif else "Non"])
    ws=wb.create_sheet("Questions");ws.append(["Questionnaire","Thématique","Ordre","Question","Type de réponse","Options","Obligatoire"])
    for q in db.query(Question).join(Questionnaire).order_by(Question.questionnaire_id,Question.ordre):ws.append([q.questionnaire.titre,q.questionnaire.categorie.nom,q.ordre,q.libelle,q.type_reponse,q.options or "","Oui" if q.obligatoire else "Non"])
    # Mise en forme simple et exploitable
    for ws in wb.worksheets:
        ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
        for cell in ws[1]:cell.font=Font(bold=True);cell.fill=PatternFill("solid",fgColor="D9EAF7");cell.alignment=Alignment(vertical="center")
        for col in range(1,ws.max_column+1):
            maxlen=max((len(str(ws.cell(r,col).value or "")) for r in range(1,min(ws.max_row,100)+1)),default=10)
            ws.column_dimensions[get_column_letter(col)].width=min(max(maxlen+2,12),55)
    out=io.BytesIO();wb.save(out);out.seek(0)
    return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":'attachment; filename="FIRSI_donnees_evaluations.xlsx"'})

def _report_data(db):
    total_coll=db.query(Collaborateur).filter(Collaborateur.actif.is_(True)).count()
    total_departures=db.query(Collaborateur).filter(Collaborateur.a_quitte.is_(True)).count()
    total_rep=db.query(Reponse).count(); avg=float(db.query(func.avg(Reponse.satisfaction_globale)).scalar() or 0)
    total_ben=db.query(Affectation).filter(Affectation.beneficie.is_(True)).count()
    collaborateurs_formes=(db.query(func.count(func.distinct(Affectation.collaborateur_id))).join(SessionFormation,Affectation.session_id==SessionFormation.id).join(Categorie,SessionFormation.categorie_id==Categorie.id).join(Collaborateur,Affectation.collaborateur_id==Collaborateur.id).filter(func.lower(Categorie.nom)=="formation",Collaborateur.actif.is_(True),Affectation.beneficie.is_(True)).scalar() or 0)
    coverage=round(collaborateurs_formes/total_coll*100,1) if total_coll else 0
    participation=round(total_rep/total_ben*100,1) if total_ben else 0
    type_items,type_sources,total_typed_evaluations=_formation_type_stats(db)
    types=[(x["type"],x["count"]) for x in type_items]
    themes=[]
    for cat in db.query(Categorie).order_by(Categorie.nom):
        score=float(db.query(func.avg(Reponse.satisfaction_globale)).join(Affectation,Reponse.affectation_id==Affectation.id).join(SessionFormation,Affectation.session_id==SessionFormation.id).filter(SessionFormation.categorie_id==cat.id).scalar() or 0)
        rec=db.query(Reponse).join(Affectation,Reponse.affectation_id==Affectation.id).join(SessionFormation,Affectation.session_id==SessionFormation.id).filter(SessionFormation.categorie_id==cat.id).count()
        themes.append((cat.nom,round(score,2),rec))
    return locals()

@app.get("/export/rapport.pdf")
def export_pdf(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    if colors is None:
        return RedirectResponse("/dashboard?error=pdf_dependency",303)
    d=_report_data(db);out=io.BytesIO()
    page_w,page_h=landscape(A4)
    c=canvas.Canvas(out,pagesize=(page_w,page_h))
    navy=colors.HexColor("#0B1739"); blue=colors.HexColor("#3563FF"); indigo=colors.HexColor("#5F6FFF")
    green=colors.HexColor("#12B76A"); gold=colors.HexColor("#F79009"); red=colors.HexColor("#D92D20")
    ink=colors.HexColor("#101828"); muted=colors.HexColor("#667085"); line=colors.HexColor("#E4E7EC"); bg=colors.HexColor("#F5F7FB")
    logo_path=BASE_DIR/"app"/"static"/"img"/"logo-firsi-um6p.png"

    def txt(x,y,text,size=9,color=ink,font="Helvetica",align="left"):
        c.setFillColor(color);c.setFont(font,size)
        if align=="center": c.drawCentredString(x,y,str(text))
        elif align=="right": c.drawRightString(x,y,str(text))
        else: c.drawString(x,y,str(text))
    def card(x,y,w,h,title,value,accent=blue,subtitle=""):
        c.setFillColor(colors.white);c.setStrokeColor(line);c.setLineWidth(.6);c.roundRect(x,y,w,h,8,fill=1,stroke=1)
        c.setFillColor(accent);c.roundRect(x+8,y+h-18,26,7,3.5,fill=1,stroke=0)
        txt(x+10,y+h-28,title.upper(),7,muted,"Helvetica-Bold")
        txt(x+10,y+15,value,18,ink,"Helvetica-Bold")
        if subtitle: txt(x+10,y+5,subtitle,6.5,muted)
    def section_title(x,y,eyebrow,title):
        txt(x,y,eyebrow.upper(),6.5,indigo,"Helvetica-Bold");txt(x,y-15,title,12,ink,"Helvetica-Bold")
    def hbar(x,y,w,label,value,maxv,accent=blue,suffix=""):
        txt(x,y,label,7.5,ink,"Helvetica-Bold");txt(x+w,y,f"{value}{suffix}",7.5,ink,"Helvetica-Bold","right")
        c.setFillColor(colors.HexColor("#EEF2F6"));c.roundRect(x,y-9,w,5,2.5,fill=1,stroke=0)
        ratio=(float(value)/maxv) if maxv else 0;c.setFillColor(accent);c.roundRect(x,y-9,max(0,min(w,w*ratio)),5,2.5,fill=1,stroke=0)
    def header(page_label):
        c.setFillColor(bg);c.rect(0,0,page_w,page_h,fill=1,stroke=0)
        c.setFillColor(navy);c.rect(0,page_h-68,page_w,68,fill=1,stroke=0)
        if logo_path.exists():
            try:c.drawImage(ImageReader(str(logo_path)),24,page_h-55,width=88,height=33,preserveAspectRatio=True,mask='auto')
            except Exception:txt(28,page_h-38,"FIRSI — UM6P",15,colors.white,"Helvetica-Bold")
        else: txt(28,page_h-38,"FIRSI — UM6P",15,colors.white,"Helvetica-Bold")
        txt(130,page_h-30,"TABLEAU DE BORD DÉCISIONNEL",13,colors.white,"Helvetica-Bold")
        txt(130,page_h-46,"Formation & expérience collaborateur · Vue exécutive confidentielle",7.5,colors.HexColor("#C7D7FE"))
        txt(page_w-26,page_h-31,page_label,7.5,colors.white,"Helvetica-Bold","right")
        txt(page_w-26,page_h-46,datetime.now().strftime("%d/%m/%Y · %H:%M"),7,colors.HexColor("#C7D7FE"),align="right")

    # PAGE 1 — dashboard
    header("SYNTHÈSE EXÉCUTIVE")
    margin=24;content_top=page_h-88
    gap=10;cw=(page_w-2*margin-5*gap)/6;ch=63
    formation_count=db.query(SessionFormation).join(Categorie,SessionFormation.categorie_id==Categorie.id).filter(SessionFormation.active.is_(True),func.lower(Categorie.nom)=="formation").count()
    kpis=[
        ("Collaborateurs actifs",d["total_coll"],blue,"Effectif de référence"),
        ("Formations actives",formation_count,indigo,"Sessions Formation"),
        ("Couverture",f'{d["coverage"]}%',blue,f'{d["collaborateurs_formes"]} formé(s)'),
        ("Participation",f'{d["participation"]}%',green,f'{d["total_rep"]} réponse(s)'),
        ("Satisfaction",f'{d["avg"]:.2f}/5',gold,"Moyenne globale"),
        ("Départs",d["total_departures"],red,"Historique conservé"),
    ]
    for i,item in enumerate(kpis):card(margin+i*(cw+gap),content_top-ch,cw,ch,*item)

    y1=content_top-ch-22
    left_w=360;right_x=margin+left_w+14;right_w=page_w-margin-right_x
    # Participation panel
    c.setFillColor(colors.white);c.setStrokeColor(line);c.roundRect(margin,y1-170,left_w,170,9,fill=1,stroke=1)
    section_title(margin+16,y1-20,"Évaluations","Participation & couverture")
    # donut-like ring participation
    cx=margin+88;cy=y1-98;r=43
    c.setLineWidth(10);c.setStrokeColor(colors.HexColor("#E9EEF7"));c.circle(cx,cy,r,fill=0,stroke=1)
    c.setStrokeColor(green)
    if d["participation"]:
        c.arc(cx-r,cy-r,cx+r,cy+r,90,-360*d["participation"]/100)
    txt(cx,cy+2,f'{d["participation"]}%',14,ink,"Helvetica-Bold","center");txt(cx,cy-12,"participation",6.5,muted,align="center")
    hbar(margin+158,y1-72,175,"Couverture",d["coverage"],100,blue,"%")
    hbar(margin+158,y1-107,175,"Participation",d["participation"],100,green,"%")
    hbar(margin+158,y1-142,175,"Satisfaction",round(d["avg"],2),5,gold,"/5")

    # Types panel based on answers
    c.setFillColor(colors.white);c.setStrokeColor(line);c.roundRect(right_x,y1-170,right_w,170,9,fill=1,stroke=1)
    section_title(right_x+16,y1-20,"Formations","Répartition par type · évaluations")
    txt(right_x+16,y1-45,"Type fixé par la session si renseigné, sinon choix du collaborateur.",6.8,muted)
    max_type=max([n for _,n in d["types"]]+[1])
    accents=[blue,indigo,gold]
    yy=y1-72
    for i,(typ,n) in enumerate(d["types"]):
        hbar(right_x+18,yy-i*34,right_w-36,typ,n,max_type,accents[i])
    src=d.get("type_sources",{})
    txt(right_x+18,y1-158,f'Source : session {src.get("Session",0)} · collaborateur {src.get("Collaborateur",0)} · non renseigné {src.get("Non renseigné",0)}',6.5,muted)

    # Bottom: themes + top sessions
    y2=y1-192
    panel_h=185
    c.setFillColor(colors.white);c.setStrokeColor(line);c.roundRect(margin,y2-panel_h,left_w,panel_h,9,fill=1,stroke=1)
    section_title(margin+16,y2-20,"Thématiques","Satisfaction moyenne")
    themes=sorted(d["themes"],key=lambda x:x[1],reverse=True)[:5]
    yy=y2-54
    for i,(name,score,rec) in enumerate(themes):
        hbar(margin+18,yy-i*27,left_w-36,(name[:30]+'…') if len(name)>30 else name,round(score,2),5,indigo,"/5")
    if not themes: txt(margin+18,y2-80,"Aucune évaluation disponible",8,muted)

    c.setFillColor(colors.white);c.setStrokeColor(line);c.roundRect(right_x,y2-panel_h,right_w,panel_h,9,fill=1,stroke=1)
    section_title(right_x+16,y2-20,"Performance","Top sessions évaluées")
    top=db.query(SessionFormation.titre,func.avg(Reponse.satisfaction_globale),func.count(Reponse.id)).join(Affectation,Affectation.session_id==SessionFormation.id).join(Reponse,Reponse.affectation_id==Affectation.id).group_by(SessionFormation.id,SessionFormation.titre).order_by(func.avg(Reponse.satisfaction_globale).desc()).limit(5).all()
    yy=y2-56
    if top:
        for i,(name,score,nb) in enumerate(top):
            label=(name[:34]+'…') if len(name)>34 else name
            hbar(right_x+18,yy-i*27,right_w-36,label,round(float(score),2),5,blue,"/5")
    else: txt(right_x+18,y2-80,"Aucune session évaluée",8,muted)
    txt(page_w/2,12,"FIRSI – UM6P · Rapport généré automatiquement · Données détaillées disponibles dans l’export Excel",6.5,muted,align="center")
    c.showPage()

    # PAGE 2 — decision support detail
    header("ANALYSE DE PILOTAGE")
    topy=page_h-92
    section_title(margin,topy,"Lecture décisionnelle","Indicateurs d’action")
    # 4 action cards, numbers only + concise labels
    action_gap=12;aw=(page_w-2*margin-action_gap*3)/4;ay=topy-92
    action_data=[
        ("Bénéficiaires",d["total_ben"],green,"Base de participation"),
        ("Évaluations reçues",d["total_rep"],blue,"Questionnaires soumis"),
        ("En attente",max(d["total_ben"]-d["total_rep"],0),gold,"Relances prioritaires"),
        ("Non formés",max(d["total_coll"]-d["collaborateurs_formes"],0),red,"Couverture à renforcer"),
    ]
    for i,item in enumerate(action_data):card(margin+i*(aw+action_gap),ay,aw,66,*item)

    # thematic matrix visual
    matrix_y=ay-26;matrix_h=245
    c.setFillColor(colors.white);c.setStrokeColor(line);c.roundRect(margin,matrix_y-matrix_h,page_w-2*margin,matrix_h,9,fill=1,stroke=1)
    section_title(margin+16,matrix_y-20,"Performance par thématique","Satisfaction et volume d’évaluations")
    rows=sorted(d["themes"],key=lambda x:(x[2],x[1]),reverse=True)[:8]
    max_rec=max([r[2] for r in rows]+[1]);row_y=matrix_y-58
    for i,(name,score,rec) in enumerate(rows):
        yy=row_y-i*24
        txt(margin+18,yy,(name[:28]+'…') if len(name)>28 else name,7.3,ink,"Helvetica-Bold")
        # satisfaction bar
        bx=margin+190;bw=210;c.setFillColor(colors.HexColor("#EEF2F6"));c.roundRect(bx,yy-4,bw,5,2.5,fill=1,stroke=0);c.setFillColor(indigo);c.roundRect(bx,yy-4,bw*(score/5 if score else 0),5,2.5,fill=1,stroke=0)
        txt(bx+bw+10,yy,f"{score:.2f}/5",7,ink,"Helvetica-Bold")
        # volume bar
        vx=margin+500;vw=220;c.setFillColor(colors.HexColor("#EEF2F6"));c.roundRect(vx,yy-4,vw,5,2.5,fill=1,stroke=0);c.setFillColor(blue);c.roundRect(vx,yy-4,vw*(rec/max_rec if max_rec else 0),5,2.5,fill=1,stroke=0)
        txt(vx+vw+8,yy,str(rec),7,ink,"Helvetica-Bold")
    if not rows: txt(margin+18,matrix_y-80,"Aucune donnée disponible",8,muted)
    txt(margin+18,matrix_y-matrix_h+16,"Confidentialité : la synthèse exécutive ne présente pas de ventilation par département.",6.5,muted)
    txt(page_w/2,12,"FIRSI – UM6P · Tableau de bord décisionnel",6.5,muted,align="center")
    c.save();out.seek(0)
    return StreamingResponse(out,media_type="application/pdf",headers={"Content-Disposition":'attachment; filename="FIRSI_dashboard_decisionnel.pdf"'})

@app.post("/admin/nettoyage-tests")
def cleanup_test_data(request:Request,confirmation:str=Form(...),db:Session=Depends(get_db)):
    if not require_admin(request):return RedirectResponse("/",303)
    if confirmation.strip().upper()!="NETTOYER":return RedirectResponse("/admin/utilisateurs?cleanup=confirmation",303)
    try:
        delete_legacy_evaluations(db)
        db.query(ReponseDetail).delete(synchronize_session=False)
        db.query(Reponse).delete(synchronize_session=False)
        db.query(Affectation).delete(synchronize_session=False)
        db.query(SessionFormation).delete(synchronize_session=False)
        db.query(OTPCode).delete(synchronize_session=False)
        if inspect(engine).has_table("collaborator_accounts"):
            db.execute(text("DELETE FROM collaborator_accounts"))
        db.query(Collaborateur).delete(synchronize_session=False)
        db.commit();audit(db,request,"Nettoyage des données de test","Collaborateurs, sessions, affectations, réponses et OTP supprimés")
        return RedirectResponse("/admin/utilisateurs?cleanup=success",303)
    except Exception as exc:
        db.rollback();print(f"[NETTOYAGE TESTS] {exc}",flush=True)
        return RedirectResponse("/admin/utilisateurs?cleanup=error",303)

@app.get("/admin/utilisateurs")
def users_page(request:Request,db:Session=Depends(get_db)):
    if not require_admin(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("users.html",{"request":request,"users":db.query(User).all(),"cleanup":request.query_params.get("cleanup")})

@app.post("/admin/utilisateurs")
def create_user(request:Request,fullname:str=Form(...),username:str=Form(...),email:str=Form(""),role:str=Form("RH"),password:str=Form(...),db:Session=Depends(get_db)):
    if not require_admin(request):return RedirectResponse("/",303)
    if not db.query(User).filter_by(username=username.strip()).first():db.add(User(fullname=fullname.strip(),username=username.strip(),email=email.strip() or None,role=role,password_hash=hash_password(password)));db.commit()
    return RedirectResponse("/admin/utilisateurs",303)
