
import os
os.environ['DATABASE_URL']='sqlite:///./test_firsi.db'
os.environ['ENVIRONMENT']='test'
import secrets
TEST_ADMIN_PASSWORD = secrets.token_urlsafe(24)
os.environ['ADMIN_USERNAME'] = 'admin'
os.environ['ADMIN_PASSWORD'] = TEST_ADMIN_PASSWORD
from fastapi.testclient import TestClient
import main
client=TestClient(main.app)

def test_home():
    r=client.get('/')
    assert r.status_code==200
    assert 'FIRSI' in r.text

def test_defaults():
    db=main.SessionLocal()
    names={c.nom for c in db.query(main.Categorie).all()}
    assert {'Formation','Team Building','Ftour Ramadan'}.issubset(names)
    assert db.query(main.Questionnaire).count()>=3
    assert db.query(main.Question).count()>=15
    db.close()

def test_login():
    r=client.post('/auth/staff',data={'username':'admin','password':TEST_ADMIN_PASSWORD},follow_redirects=False)
    assert r.status_code==303


def test_product_owner_features():
    src=open("main.py",encoding="utf-8").read()
    assert 'a_quitte=Column(Boolean,default=False,nullable=False)' in src
    assert 'Certifiante;Diplômante;Qualifiante' in src
    assert '@app.post("/collaborateurs/{cid}/departure")' in src
    assert 'recommendationReasonBlock' in open("app/templates/evaluation_form.html",encoding="utf-8").read()
    assert 'Nombre de formations' in open("app/templates/dashboard.html",encoding="utf-8").read()


def test_exports_and_question_types():
    r=client.post('/auth/staff',data={'username':'admin','password':TEST_ADMIN_PASSWORD},follow_redirects=False)
    assert r.status_code==303
    r=client.get('/export/evaluations.xlsx',follow_redirects=False)
    assert r.status_code==200
    assert r.headers['content-type'].startswith('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    r=client.get('/export/rapport.pdf',follow_redirects=False)
    assert r.status_code==200
    assert r.headers['content-type'].startswith('application/pdf')
    html=open('app/templates/questionnaire_builder.html',encoding='utf-8').read()
    assert 'choix_multiple' in html
    eval_html=open('app/templates/evaluation_form.html',encoding='utf-8').read()
    assert "type_reponse=='choix_multiple'" in eval_html

def test_consultation_access_and_lifecycle():
    import re
    from uuid import uuid4
    guest=TestClient(main.app)
    assert guest.get('/consultations').status_code == 403
    assert guest.get('/mon-espace/consultations').status_code == 403
    staff=TestClient(main.app)
    staff.post('/auth/staff',data={'username':'admin','password':TEST_ADMIN_PASSWORD})
    page=staff.get('/consultations')
    assert page.status_code == 200
    token=re.search(r'name="csrf_token" value="([^"]+)"',page.text).group(1)
    assert staff.post('/consultations',data={'titre':'Non autorisé','csrf_token':'bad'}).status_code==403
    name='Consultation '+uuid4().hex[:8]
    r=staff.post('/consultations',data={'titre':name,'csrf_token':token},follow_redirects=False)
    assert r.status_code==303
    url=r.headers['location']; cid=int(url.split('/')[-1])
    detail=staff.get(url)
    assert detail.status_code==200
    r=staff.post(url+'/formations',data={'titre':'Power BI','dates_previsionnelles':'23-25 décembre','csrf_token':token},follow_redirects=False)
    assert r.status_code==303
    r=staff.post(url+'/statut',data={'statut':'ouverte','csrf_token':token},follow_redirects=False)
    assert r.status_code==303
    assert staff.get(url+'/export.xlsx').status_code==200
    with main.SessionLocal() as db:
        c=db.get(main.Consultation,cid)
        assert c.statut=='ouverte' and len(c.formations)==1

def test_consultation_response_lock_and_admin_unlock():
    from uuid import uuid4
    import re
    unique=uuid4().hex[:12]
    with main.SessionLocal() as db:
        admin=db.query(main.User).filter_by(username='admin').first()
        admin.must_change_password=False
        collaborator=main.Collaborateur(code='TEST-'+unique,nom='Test',prenom='Consultation',email=unique+'@example.test',password_hash=main.hash_password('TestPass!123'),must_change_password=False,actif=True)
        db.add(collaborator)
        consultation=main.Consultation(titre='Verrouillage '+unique,statut='ouverte')
        db.add(consultation);db.flush()
        formation=main.ConsultationFormation(consultation_id=consultation.id,titre='Power BI')
        db.add(formation);db.commit()
        fid=formation.id;cid=consultation.id;collab_id=collaborator.id
    collaborator_client=TestClient(main.app)
    assert collaborator_client.post('/auth/collaborateur',data={'email':unique+'@example.test','password':'TestPass!123'},follow_redirects=False).status_code==303
    page=collaborator_client.get('/mon-espace/consultations')
    token=re.search(r'name="csrf_token" value="([^"]+)"',page.text).group(1)
    endpoint=f'/mon-espace/consultations/{fid}/repondre'
    assert collaborator_client.post(endpoint,data={'choix':'oui','csrf_token':token},follow_redirects=False).status_code==303
    assert 'verrouillée' in collaborator_client.get('/mon-espace/consultations').text
    assert collaborator_client.post(endpoint,data={'choix':'non','csrf_token':token},follow_redirects=False).status_code==409
    with main.SessionLocal() as db:
        r=db.query(main.ConsultationReponse).filter_by(formation_id=fid,collaborateur_id=collab_id).one()
        assert r.choix=='oui' and not r.modification_autorisee
        rid=r.id
    staff=TestClient(main.app)
    staff.post('/auth/staff',data={'username':'admin','password':TEST_ADMIN_PASSWORD})
    detail=staff.get(f'/consultations/{cid}')
    staff_token=re.search(r'name="csrf_token" value="([^"]+)"',detail.text).group(1)
    unlock=f'/consultations/{cid}/reponses/{rid}/autoriser-modification'
    assert collaborator_client.post(unlock,data={'csrf_token':token},follow_redirects=False).status_code==403
    assert staff.post(unlock,data={'csrf_token':'invalid'},follow_redirects=False).status_code==403
    assert staff.post(unlock,data={'csrf_token':staff_token},follow_redirects=False).status_code==303
    assert collaborator_client.post(endpoint,data={'choix':'non','csrf_token':token},follow_redirects=False).status_code==303
    assert collaborator_client.post(endpoint,data={'choix':'oui','csrf_token':token},follow_redirects=False).status_code==409
    with main.SessionLocal() as db:
        r=db.get(main.ConsultationReponse,rid)
        assert r.choix=='non' and not r.modification_autorisee
