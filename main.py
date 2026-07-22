from __future__ import annotations

from datetime import datetime
from pathlib import Path
import io
import os
import secrets
import string

from fastapi import FastAPI, Request, Form, Depends, HTTPException
from fastapi.responses import RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from sqlalchemy import (
    create_engine, Column, Integer, String, Boolean, DateTime,
    ForeignKey, Text, UniqueConstraint, func
)
from sqlalchemy.orm import declarative_base, sessionmaker, Session, relationship
from passlib.context import CryptContext
from openpyxl import Workbook

BASE_DIR = Path(__file__).resolve().parent
SQLITE_URL = f"sqlite:///{BASE_DIR / 'evaluation.db'}"
DATABASE_URL = os.getenv("DATABASE_URL", SQLITE_URL)
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine_kwargs = {"pool_pre_ping": True}
if DATABASE_URL.startswith("sqlite"):
    engine_kwargs["connect_args"] = {"check_same_thread": False}
engine = create_engine(DATABASE_URL, **engine_kwargs)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")


class User(Base):
    __tablename__ = "app_users"
    id = Column(Integer, primary_key=True)
    fullname = Column(String(180), nullable=False)
    username = Column(String(80), unique=True, nullable=False, index=True)
    email = Column(String(180), nullable=True)
    password_hash = Column(String(255), nullable=False)
    role = Column(String(20), nullable=False, default="RH")
    active = Column(Boolean, default=True)
    must_change_password = Column(Boolean, default=True)
    failed_attempts = Column(Integer, default=0)
    locked_until = Column(DateTime, nullable=True)
    last_login = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)


class Collaborateur(Base):
    __tablename__ = "collaborateurs"
    id = Column(Integer, primary_key=True)
    code = Column(String(20), unique=True, nullable=False, index=True)
    nom = Column(String(120), nullable=False)
    prenom = Column(String(120), nullable=False)
    email = Column(String(180), nullable=True)
    direction = Column(String(180), nullable=True)
    fonction = Column(String(180), nullable=True)
    token = Column(String(100), unique=True, nullable=False, index=True)
    actif = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    affectations = relationship("Affectation", back_populates="collaborateur", cascade="all, delete-orphan")
    account = relationship("CollaboratorAccount", back_populates="collaborateur", uselist=False, cascade="all, delete-orphan")


class CollaboratorAccount(Base):
    __tablename__ = "collaborator_accounts"
    id = Column(Integer, primary_key=True)
    collaborateur_id = Column(Integer, ForeignKey("collaborateurs.id"), unique=True, nullable=False)
    password_hash = Column(String(255), nullable=False)
    must_change_password = Column(Boolean, default=True)
    failed_attempts = Column(Integer, default=0)
    last_login = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    collaborateur = relationship("Collaborateur", back_populates="account")


class Formation(Base):
    __tablename__ = "formations"
    id = Column(Integer, primary_key=True)
    titre = Column(String(255), nullable=False)
    type_formation = Column(String(50), nullable=False)
    description = Column(Text, nullable=True)
    active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    affectations = relationship("Affectation", back_populates="formation", cascade="all, delete-orphan")


class Affectation(Base):
    __tablename__ = "affectations"
    __table_args__ = (UniqueConstraint("collaborateur_id", "formation_id", name="uq_affectation_collaborateur_formation"),)
    id = Column(Integer, primary_key=True)
    collaborateur_id = Column(Integer, ForeignKey("collaborateurs.id"), nullable=False)
    formation_id = Column(Integer, ForeignKey("formations.id"), nullable=False)
    statut = Column(String(30), default="À évaluer")
    created_at = Column(DateTime, default=datetime.utcnow)
    collaborateur = relationship("Collaborateur", back_populates="affectations")
    formation = relationship("Formation", back_populates="affectations")
    evaluation = relationship("Evaluation", back_populates="affectation", uselist=False, cascade="all, delete-orphan")


class Evaluation(Base):
    __tablename__ = "evaluations"
    id = Column(Integer, primary_key=True)
    affectation_id = Column(Integer, ForeignKey("affectations.id"), unique=True, nullable=False)
    contenu = Column(Integer, nullable=False)
    formateur = Column(Integer, nullable=False)
    organisation = Column(Integer, nullable=False)
    application_poste = Column(Integer, nullable=False)
    satisfaction = Column(Integer, nullable=False)
    recommande = Column(String(3), nullable=False)
    observations = Column(Text, nullable=True)
    points_forts = Column(Text, nullable=True)
    points_ameliorer = Column(Text, nullable=True)
    date_reponse = Column(DateTime, default=datetime.utcnow)
    affectation = relationship("Affectation", back_populates="evaluation")


Base.metadata.create_all(bind=engine)

app = FastAPI(title="Training Evaluation FIRSI", version="2.0")
app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SECRET_KEY", secrets.token_urlsafe(32)),
    same_site="lax",
    https_only=os.getenv("ENVIRONMENT") == "production",
    max_age=60 * 60 * 8,
)
app.mount("/static", StaticFiles(directory=BASE_DIR / "app" / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "app" / "templates")


def get_db():
    db = SessionLocal()
    try: yield db
    finally: db.close()


def random_password(length=10):
    alphabet = string.ascii_letters + string.digits + "!@#"
    return "".join(secrets.choice(alphabet) for _ in range(length))


def hash_password(password): return pwd_context.hash(password)
def verify_password(password, hashed): return pwd_context.verify(password, hashed)


def seed_admin():
    db = SessionLocal()
    try:
        if not db.query(User).filter(User.role == "ADMIN").first():
            username = os.getenv("ADMIN_USERNAME", "admin")
            password = os.getenv("ADMIN_PASSWORD", "Admin@2026")
            db.add(User(fullname="Administrateur principal", username=username, email=os.getenv("ADMIN_EMAIL"),
                        password_hash=hash_password(password), role="ADMIN", active=True,
                        must_change_password=True))
            db.commit()
    finally: db.close()
seed_admin()


def auth_user(request: Request, roles=None):
    uid = request.session.get("user_id")
    role = request.session.get("role")
    if not uid or (roles and role not in roles): return None
    return {"id": uid, "role": role, "name": request.session.get("name")}


def require_staff(request: Request, roles=("ADMIN", "RH")):
    return auth_user(request, roles)


def next_code(db):
    max_id = db.query(func.max(Collaborateur.id)).scalar() or 0
    return f"COL-{max_id + 1:04d}"


def redirect_login(): return RedirectResponse("/?login=required", status_code=303)


@app.get("/")
def landing(request: Request):
    return templates.TemplateResponse(request=request, name="landing.html", context={"error": request.query_params.get("error")})


@app.post("/auth/staff")
def staff_login(request: Request, username: str = Form(...), password: str = Form(...), db: Session = Depends(get_db)):
    user = db.query(User).filter(User.username == username.strip()).first()
    if not user or not user.active or not verify_password(password, user.password_hash):
        return RedirectResponse("/?error=staff", status_code=303)
    user.last_login = datetime.utcnow(); user.failed_attempts = 0; db.commit()
    request.session.clear(); request.session.update({"user_id": user.id, "role": user.role, "name": user.fullname})
    return RedirectResponse("/change-password" if user.must_change_password else "/dashboard", status_code=303)


@app.post("/auth/collaborateur")
def collaborator_login(request: Request, code: str = Form(...), password: str = Form(...), db: Session = Depends(get_db)):
    collab = db.query(Collaborateur).filter(Collaborateur.code == code.strip().upper(), Collaborateur.actif.is_(True)).first()
    if not collab or not collab.account or not verify_password(password, collab.account.password_hash):
        return RedirectResponse("/?error=collaborateur", status_code=303)
    collab.account.last_login = datetime.utcnow(); db.commit()
    request.session.clear(); request.session.update({"collaborateur_id": collab.id, "role": "COLLABORATEUR", "name": collab.code})
    return RedirectResponse("/change-password" if collab.account.must_change_password else "/mon-espace", status_code=303)


@app.get("/logout")
def logout(request: Request): request.session.clear(); return RedirectResponse("/", status_code=303)


@app.get("/change-password")
def change_password_page(request: Request):
    if not request.session.get("role"): return redirect_login()
    return templates.TemplateResponse(request=request, name="change_password.html", context={})


@app.post("/change-password")
def change_password(request: Request, new_password: str = Form(...), confirm_password: str = Form(...), db: Session = Depends(get_db)):
    if len(new_password) < 8 or new_password != confirm_password:
        return RedirectResponse("/change-password?error=1", status_code=303)
    if request.session.get("role") == "COLLABORATEUR":
        acc = db.query(CollaboratorAccount).filter(CollaboratorAccount.collaborateur_id == request.session["collaborateur_id"]).first()
        acc.password_hash = hash_password(new_password); acc.must_change_password = False
        target = "/mon-espace"
    else:
        user = db.get(User, request.session.get("user_id")); user.password_hash = hash_password(new_password); user.must_change_password = False
        target = "/dashboard"
    db.commit(); return RedirectResponse(target, status_code=303)


@app.get("/dashboard")
def dashboard(request: Request, db: Session = Depends(get_db)):
    user = require_staff(request)
    if not user: return redirect_login()
    evaluations = db.query(Evaluation).all(); affect_count = db.query(Affectation).count()
    kpis = {
        "collaborateurs": db.query(Collaborateur).filter(Collaborateur.actif.is_(True)).count(),
        "formations": db.query(Formation).filter(Formation.active.is_(True)).count(),
        "evaluations": len(evaluations),
        "participation": round(len(evaluations)/affect_count*100,1) if affect_count else 0,
        "satisfaction": round(sum(e.satisfaction for e in evaluations)/len(evaluations),2) if evaluations else 0,
        "recommandation": round(sum(e.recommande=="Oui" for e in evaluations)/len(evaluations)*100,1) if evaluations else 0,
    }
    by_training = db.query(Formation.titre, func.avg(Evaluation.satisfaction)).join(Affectation).join(Evaluation).group_by(Formation.id).all()
    distribution = [db.query(Evaluation).filter(Evaluation.satisfaction == i).count() for i in range(1,6)]
    return templates.TemplateResponse(request=request, name="dashboard.html", context={
        "active_page":"dashboard", "user":user, "kpis":kpis,
        "chart_labels":[x[0] for x in by_training], "chart_values":[round(float(x[1]),2) for x in by_training],
        "distribution":distribution})


@app.get("/admin/utilisateurs")
def users_page(request: Request, db: Session = Depends(get_db)):
    user = auth_user(request, ("ADMIN",))
    if not user: return redirect_login()
    return templates.TemplateResponse(request=request, name="users.html", context={"active_page":"users","user":user,"users":db.query(User).order_by(User.id.desc()).all(),"temp":request.query_params.get("temp")})


@app.post("/admin/utilisateurs")
def create_user(request: Request, fullname: str=Form(...), username: str=Form(...), email: str=Form(""), role: str=Form("RH"), db:Session=Depends(get_db)):
    if not auth_user(request,("ADMIN",)): return redirect_login()
    if db.query(User).filter(User.username==username.strip()).first(): return RedirectResponse("/admin/utilisateurs?error=exists",303)
    temp=random_password(); db.add(User(fullname=fullname.strip(),username=username.strip(),email=email.strip() or None,password_hash=hash_password(temp),role=role if role in ("ADMIN","RH") else "RH",must_change_password=True)); db.commit()
    return RedirectResponse(f"/admin/utilisateurs?temp={temp}",303)


@app.post("/admin/utilisateurs/{uid}/toggle")
def toggle_user(uid:int,request:Request,db:Session=Depends(get_db)):
    if not auth_user(request,("ADMIN",)): return redirect_login()
    u=db.get(User,uid)
    if u and u.id != request.session.get("user_id"): u.active=not u.active; db.commit()
    return RedirectResponse("/admin/utilisateurs",303)


@app.post("/admin/utilisateurs/{uid}/reset")
def reset_user(uid:int,request:Request,db:Session=Depends(get_db)):
    if not auth_user(request,("ADMIN",)): return redirect_login()
    u=db.get(User,uid); temp=random_password()
    if u: u.password_hash=hash_password(temp); u.must_change_password=True; db.commit()
    return RedirectResponse(f"/admin/utilisateurs?temp={temp}",303)


@app.get("/collaborateurs")
def collaborators_page(request:Request,db:Session=Depends(get_db)):
    user=require_staff(request)
    if not user:return redirect_login()
    return templates.TemplateResponse(request=request,name="collaborateurs.html",context={"active_page":"collaborateurs","user":user,"collaborateurs":db.query(Collaborateur).order_by(Collaborateur.id.desc()).all(),"temp":request.query_params.get("temp")})


@app.post("/collaborateurs")
def create_collaborator(request:Request,nom:str=Form(...),prenom:str=Form(...),email:str=Form(""),direction:str=Form(""),fonction:str=Form(""),db:Session=Depends(get_db)):
    if not require_staff(request):return redirect_login()
    temp=random_password(); c=Collaborateur(code=next_code(db),nom=nom.strip(),prenom=prenom.strip(),email=email.strip() or None,direction=direction.strip() or None,fonction=fonction.strip() or None,token=secrets.token_urlsafe(24),actif=True)
    db.add(c);db.flush();db.add(CollaboratorAccount(collaborateur_id=c.id,password_hash=hash_password(temp),must_change_password=True));db.commit()
    return RedirectResponse(f"/collaborateurs?temp={temp}&code={c.code}",303)


@app.post("/collaborateurs/{cid}/toggle")
def toggle_collaborator(cid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return redirect_login()
    c=db.get(Collaborateur,cid)
    if c:c.actif=not c.actif;db.commit()
    return RedirectResponse("/collaborateurs",303)


@app.post("/collaborateurs/{cid}/reset")
def reset_collaborator(cid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return redirect_login()
    c=db.get(Collaborateur,cid);temp=random_password()
    if c:
        if c.account:
            c.account.password_hash=hash_password(temp);c.account.must_change_password=True
        else:
            db.add(CollaboratorAccount(collaborateur_id=c.id,password_hash=hash_password(temp),must_change_password=True))
        db.commit()
        return RedirectResponse(f"/collaborateurs?temp={temp}&code={c.code}",303)
    return RedirectResponse("/collaborateurs",303)


@app.get("/formations")
def formations_page(request:Request,db:Session=Depends(get_db)):
    user=require_staff(request)
    if not user:return redirect_login()
    return templates.TemplateResponse(request=request,name="formations.html",context={"active_page":"formations","user":user,"formations":db.query(Formation).order_by(Formation.id.desc()).all()})


@app.post("/formations")
def create_formation(request:Request,titre:str=Form(...),type_formation:str=Form(...),description:str=Form(""),db:Session=Depends(get_db)):
    if not require_staff(request):return redirect_login()
    db.add(Formation(titre=titre.strip(),type_formation=type_formation,description=description.strip() or None));db.commit();return RedirectResponse("/formations",303)


@app.post("/formations/{fid}/toggle")
def toggle_formation(fid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return redirect_login()
    f=db.get(Formation,fid)
    if f:f.active=not f.active;db.commit()
    return RedirectResponse("/formations",303)


@app.get("/affectations")
def assignments_page(request:Request,db:Session=Depends(get_db)):
    user=require_staff(request)
    if not user:return redirect_login()
    return templates.TemplateResponse(request=request,name="affectations.html",context={"active_page":"affectations","user":user,"collaborateurs":db.query(Collaborateur).filter(Collaborateur.actif.is_(True)).all(),"formations":db.query(Formation).filter(Formation.active.is_(True)).all(),"affectations":db.query(Affectation).order_by(Affectation.id.desc()).all()})


@app.post("/affectations")
def create_assignment(request:Request,collaborateur_id:int=Form(...),formation_id:int=Form(...),db:Session=Depends(get_db)):
    if not require_staff(request):return redirect_login()
    if not db.query(Affectation).filter_by(collaborateur_id=collaborateur_id,formation_id=formation_id).first():db.add(Affectation(collaborateur_id=collaborateur_id,formation_id=formation_id));db.commit()
    return RedirectResponse("/affectations",303)


@app.post("/affectations/{aid}/delete")
def delete_assignment(aid:int,request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return redirect_login()
    a=db.get(Affectation,aid)
    if a:db.delete(a);db.commit()
    return RedirectResponse("/affectations",303)


@app.get("/evaluations")
def evaluations_page(request:Request,db:Session=Depends(get_db)):
    user=require_staff(request)
    if not user:return redirect_login()
    return templates.TemplateResponse(request=request,name="evaluations.html",context={"active_page":"evaluations","user":user,"evaluations":db.query(Evaluation).order_by(Evaluation.date_reponse.desc()).all()})


@app.get("/mon-espace")
def collaborator_space(request:Request,db:Session=Depends(get_db)):
    cid=request.session.get("collaborateur_id")
    if request.session.get("role")!="COLLABORATEUR" or not cid:return redirect_login()
    c=db.get(Collaborateur,cid)
    return templates.TemplateResponse(request=request,name="collaborator_space.html",context={"collaborateur":c,"affectations":db.query(Affectation).filter(Affectation.collaborateur_id==cid).all()})


@app.post("/mon-espace/evaluation/{aid}")
def submit_evaluation(aid:int,request:Request,contenu:int=Form(...),formateur:int=Form(...),organisation:int=Form(...),application_poste:int=Form(...),satisfaction:int=Form(...),recommande:str=Form(...),observations:str=Form(""),points_forts:str=Form(""),points_ameliorer:str=Form(""),db:Session=Depends(get_db)):
    cid=request.session.get("collaborateur_id");a=db.get(Affectation,aid)
    if request.session.get("role")!="COLLABORATEUR" or not a or a.collaborateur_id!=cid:raise HTTPException(403)
    if a.evaluation:return RedirectResponse("/mon-espace?already=1",303)
    if any(n not in range(1,6) for n in [contenu,formateur,organisation,application_poste,satisfaction]):raise HTTPException(400)
    db.add(Evaluation(affectation_id=aid,contenu=contenu,formateur=formateur,organisation=organisation,application_poste=application_poste,satisfaction=satisfaction,recommande=recommande,observations=observations.strip() or None,points_forts=points_forts.strip() or None,points_ameliorer=points_ameliorer.strip() or None));a.statut="Évaluée";db.commit()
    return RedirectResponse("/mon-espace?success=1",303)


@app.get("/export/evaluations.xlsx")
def export_excel(request:Request,db:Session=Depends(get_db)):
    if not require_staff(request):return redirect_login()
    wb=Workbook();ws=wb.active;ws.title="Évaluations";ws.append(["Code","Formation","Type","Contenu","Formateur","Organisation","Application poste","Satisfaction","Recommande","Observations","Points forts","Points à améliorer","Date"])
    for e in db.query(Evaluation).order_by(Evaluation.date_reponse.desc()).all():
        a=e.affectation;ws.append([a.collaborateur.code,a.formation.titre,a.formation.type_formation,e.contenu,e.formateur,e.organisation,e.application_poste,e.satisfaction,e.recommande,e.observations or "",e.points_forts or "",e.points_ameliorer or "",e.date_reponse.strftime("%d/%m/%Y %H:%M")])
    out=io.BytesIO();wb.save(out);out.seek(0)
    return StreamingResponse(out,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":'attachment; filename="evaluations_formations.xlsx"'})
