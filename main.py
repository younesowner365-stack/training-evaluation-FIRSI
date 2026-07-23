
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
    fonction=Column(String(180))
    departement_id=Column(Integer,ForeignKey("departements.id"))
    actif=Column(Boolean,default=True)
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
    finally:db.close()
seed()

@app.get("/")
def landing(request:Request):return templates.TemplateResponse("landing.html",{"request":request,"error":request.query_params.get("error")})

@app.post("/auth/staff")
def staff_login(request:Request,username:str=Form(...),password:str=Form(...),db:Session=Depends(get_db)):
    u=db.query(User).filter(User.username==username.strip(),User.active.is_(True)).first()
    if not u or not verify_password(password,u.password_hash):return RedirectResponse("/?error=staff",303)
    request.session.clear();request.session.update({"user_id":u.id,"role":u.role,"name":u.fullname})
    return RedirectResponse("/change-password" if u.must_change_password else "/dashboard",303)

@app.post("/auth/otp/request")
def otp_request(request:Request,email:str=Form(...),db:Session=Depends(get_db)):
    email=email.strip().lower();c=db.query(Collaborateur).filter_by(email=email,actif=True).first()
    if not c:return RedirectResponse("/?error=email",303)
    code="".join(secrets.choice(string.digits) for _ in range(6))
    db.query(OTPCode).filter_by(email=email,used=False).update({"used":True})
    db.add(OTPCode(email=email,code_hash=hash_password(code),expires_at=datetime.utcnow()+timedelta(minutes=10)));db.commit()
    sent,_=send_email(email,"Code de connexion FIRSI - UM6P",f"Votre code est : {code}\nValable 10 minutes.")
    request.session["otp_email"]=email
    if not sent and os.getenv("ENVIRONMENT")!="production":request.session["dev_otp"]=code
    return RedirectResponse("/otp",303)

@app.get("/otp")
def otp_page(request:Request):
    if not request.session.get("otp_email"):return RedirectResponse("/",303)
    return templates.TemplateResponse("otp.html",{"request":request,"email":request.session["otp_email"],"dev_otp":request.session.get("dev_otp")})

@app.post("/auth/otp/verify")
def otp_verify(request:Request,code:str=Form(...),db:Session=Depends(get_db)):
    email=request.session.get("otp_email")
    otp=db.query(OTPCode).filter(OTPCode.email==email,OTPCode.used.is_(False),OTPCode.expires_at>datetime.utcnow()).order_by(OTPCode.id.desc()).first()
    if not otp or not verify_password(code.strip(),otp.code_hash):return RedirectResponse("/otp?error=1",303)
    otp.used=True;c=db.query(Collaborateur).filter_by(email=email).first();db.commit()
    request.session.clear();request.session.update({"role":"COLLABORATEUR","collaborateur_id":c.id,"name":f"{c.prenom} {c.nom}"});return RedirectResponse("/mon-espace",303)

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

@app.get("/dashboard")
def dashboard(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)

    total_aff=db.query(Affectation).count()
    total_rep=db.query(Reponse).count()
    pending=max(total_aff-total_rep,0)
    satisfaction=db.query(func.avg(Reponse.satisfaction_globale)).scalar() or 0

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
            .filter(SessionFormation.categorie_id==cat.id)
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
        .filter(Affectation.statut!="Évaluée")
        .order_by(Affectation.created_at.desc())
        .limit(5).all()
    )

    return templates.TemplateResponse("dashboard.html",{
        "request":request,
        "kpis":{
            "collaborateurs":db.query(Collaborateur).filter(Collaborateur.actif.is_(True)).count(),
            "sessions":db.query(SessionFormation).filter(SessionFormation.active.is_(True)).count(),
            "participation":round(total_rep/total_aff*100,1) if total_aff else 0,
            "satisfaction":round(float(satisfaction),2),
            "attente":pending,
            "recues":total_rep,
        },
        "themes":themes,
        "department_labels":[x[0] for x in deps],
        "department_values":[x[1] for x in deps],
        "top_sessions":top_sessions,
        "recent_pending":recent_pending,
    })

@app.get("/categories")
def categories_page(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    return templates.TemplateResponse("categories.html",{"request":request,"categories":db.query(Categorie).order_by(Categorie.nom).all()})

@app.post("/categories")
def create_category(request:Request,nom:str=Form(...),description:str=Form(""),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    if not db.query(Categorie).filter_by(nom=nom.strip()).first():db.add(Categorie(nom=nom.strip(),description=description.strip() or None));db.commit();audit(db,request,"Création catégorie",nom)
    return RedirectResponse("/categories",303)

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
def create_collab(request:Request,nom:str=Form(...),prenom:str=Form(...),email:str=Form(...),fonction:str=Form(""),departement_id:int|None=Form(None),db:Session=Depends(get_db)):
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
            fonction=fonction.strip() or None,
            departement_id=departement_id
        ))
        db.commit()
        audit(db,request,"Création collaborateur",f"{prenom.strip()} {nom.strip()} ({clean_email})")
        return RedirectResponse("/collaborateurs?message=collaborateur_ajoute",303)
    except IntegrityError:
        db.rollback()
        return RedirectResponse("/collaborateurs?error=donnees_dupliquees",303)

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
def create_session(request:Request,titre:str=Form(...),categorie_id:int=Form(...),questionnaire_id:int=Form(...),date_debut:str=Form(""),date_fin:str=Form(""),lieu:str=Form(""),animateur:str=Form(""),description:str=Form(""),db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    db.add(SessionFormation(titre=titre.strip(),categorie_id=categorie_id,questionnaire_id=questionnaire_id,date_debut=date_debut or None,date_fin=date_fin or None,lieu=lieu.strip() or None,animateur=animateur.strip() or None,description=description.strip() or None));db.commit();return RedirectResponse("/sessions",303)

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
        a=Affectation(collaborateur_id=collaborateur_id,session_id=session_id);db.add(a);db.commit();db.refresh(a)
        if envoyer_invitation=="oui":
            link=invitation_link(request,a.id)
            sent,_=send_email(
                a.collaborateur.email,
                "Invitation FIRSI - UM6P",
                f"Bonjour {a.collaborateur.prenom},\n\n"
                f"Merci d’évaluer : {a.session.titre}.\n"
                f"Accédez à la plateforme : {link}\n"
                f"Puis connectez-vous avec votre adresse e-mail professionnelle."
            )
            a.invitation_envoyee=sent;db.commit()
            return RedirectResponse(
                f"/affectations?mail_status={'sent' if sent else 'link'}&invite_link={link}",
                303
            )
    return RedirectResponse("/affectations?message=affectation_ajoutee",303)

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
        f"Puis connectez-vous avec votre adresse e-mail professionnelle."
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
    collaborateur=db.query(Collaborateur).filter(
        Collaborateur.token==token,
        Collaborateur.actif.is_(True)
    ).first()
    if not collaborateur:
        return RedirectResponse("/?error=lien_invalide",303)
    request.session.clear()
    request.session.update({
        "role":"COLLABORATEUR",
        "collaborateur_id":collaborateur.id,
        "name":f"{collaborateur.prenom} {collaborateur.nom}"
    })
    return RedirectResponse("/mon-espace",303)

@app.get("/mon-espace")
def collaborator_space(request:Request,db:Session=Depends(get_db)):
    cid=request.session.get("collaborateur_id")
    if request.session.get("role")!="COLLABORATEUR" or not cid:return RedirectResponse("/",303)
    c=db.get(Collaborateur,cid);return templates.TemplateResponse("collaborator_space.html",{"request":request,"collaborateur":c,"affectations":c.affectations})

@app.get("/mon-espace/evaluer/{aid}")
def evaluation_form(aid:int,request:Request,db:Session=Depends(get_db)):
    a=db.get(Affectation,aid)
    if request.session.get("role")!="COLLABORATEUR" or not a or a.collaborateur_id!=request.session.get("collaborateur_id"):raise HTTPException(403)
    if a.reponse:return RedirectResponse("/mon-espace?already=1",303)
    return templates.TemplateResponse("evaluation_form.html",{"request":request,"affectation":a,"questions":a.session.questionnaire.questions})

@app.post("/mon-espace/evaluer/{aid}")
async def submit_evaluation(aid:int,request:Request,db:Session=Depends(get_db)):
    a=db.get(Affectation,aid)
    if request.session.get("role")!="COLLABORATEUR" or not a or a.collaborateur_id!=request.session.get("collaborateur_id"):raise HTTPException(403)
    form=await request.form();scores=[];comments=[];recommend=False;values=[]
    for q in a.session.questionnaire.questions:
        v=str(form.get(f"q_{q.id}","")).strip()
        if q.obligatoire and not v:return RedirectResponse(f"/mon-espace/evaluer/{aid}?error=required",303)
        if q.type_reponse in {"note5","note10"} and v:
            try:scores.append(float(v))
            except:pass
        if q.type_reponse=="oui_non" and v.lower()=="oui":recommend=True
        if q.type_reponse in {"texte","paragraphe"} and v:comments.append(v)
        values.append((q,v))
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
    wb=Workbook();ws=wb.active;ws.title="Évaluations";ws.append(["Collaborateur","Email","Département","Thématique","Session","Satisfaction","Recommande","Commentaire","Date"])
    for r in db.query(Reponse).all():
        a=r.affectation;ws.append([f"{a.collaborateur.prenom} {a.collaborateur.nom}",a.collaborateur.email,a.collaborateur.departement.nom if a.collaborateur.departement else "",a.session.categorie.nom,a.session.titre,r.satisfaction_globale,"Oui" if r.recommande else "Non",r.commentaire or "",r.date_reponse.strftime("%d/%m/%Y")])
    out=io.BytesIO();wb.save(out);out.seek(0);return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":'attachment; filename="evaluations_firsi.xlsx"'})

def build_simple_pdf(lines:list[str])->bytes:
    """Génère un PDF texte simple sans dépendance externe."""
    safe=[]
    for line in lines:
        cleaned=line.replace("\\","\\\\").replace("(","\\(").replace(")","\\)")
        cleaned=cleaned.encode("latin-1","replace").decode("latin-1")
        safe.append(cleaned)
    content=["BT","/F1 18 Tf","50 790 Td"]
    for i,line in enumerate(safe):
        if i:
            content.append("0 -24 Td")
        content.append(f"({line}) Tj")
    content.append("ET")
    stream="\n".join(content).encode("latin-1")
    objects=[
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        b"<< /Length %d >>\nstream\n"%len(stream)+stream+b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    pdf=bytearray(b"%PDF-1.4\n")
    offsets=[0]
    for idx,obj in enumerate(objects,1):
        offsets.append(len(pdf))
        pdf.extend(f"{idx} 0 obj\n".encode()+obj+b"\nendobj\n")
    xref=len(pdf)
    pdf.extend(f"xref\n0 {len(objects)+1}\n".encode())
    pdf.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        pdf.extend(f"{off:010d} 00000 n \n".encode())
    pdf.extend(f"trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode())
    return bytes(pdf)

@app.get("/export/rapport.pdf")
def export_pdf(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return RedirectResponse("/",303)
    total=db.query(Reponse).count()
    avg=db.query(func.avg(Reponse.satisfaction_globale)).scalar() or 0
    data=build_simple_pdf([
        "FIRSI - UM6P",
        "Rapport des evaluations",
        f"Nombre d evaluations : {total}",
        f"Satisfaction moyenne : {round(float(avg),2)}",
        "© FIRSI - UM6P",
    ])
    return StreamingResponse(io.BytesIO(data),media_type="application/pdf",headers={"Content-Disposition":'attachment; filename="rapport_firsi.pdf"'})

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
