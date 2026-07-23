
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
