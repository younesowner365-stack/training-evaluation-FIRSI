
import os
os.environ['DATABASE_URL']='sqlite:///./test_firsi.db'
os.environ['ENVIRONMENT']='test'
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
    r=client.post('/auth/staff',data={'username':'admin','password':'Admin@2026'},follow_redirects=False)
    assert r.status_code==303


def test_product_owner_features():
    src=open("main.py",encoding="utf-8").read()
    assert 'a_quitte=Column(Boolean,default=False,nullable=False)' in src
    assert 'Certifiante;Diplômante;Qualifiante' in src
    assert '@app.post("/collaborateurs/{cid}/departure")' in src
    assert 'recommendationReasonBlock' in open("app/templates/evaluation_form.html",encoding="utf-8").read()
    assert 'Nombre de formations' in open("app/templates/dashboard.html",encoding="utf-8").read()


def test_exports_and_question_types():
    r=client.post('/auth/staff',data={'username':'admin','password':'Admin@2026'},follow_redirects=False)
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
