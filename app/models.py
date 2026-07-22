from sqlalchemy import Column
from sqlalchemy import Integer
from sqlalchemy import String
from sqlalchemy import Boolean
from sqlalchemy import ForeignKey

from app.database.database import Base


class Collaborateur(Base):

    __tablename__ = "collaborateurs"

    id = Column(Integer, primary_key=True)

    code = Column(String, unique=True)

    nom = Column(String)

    email = Column(String)

    actif = Column(Boolean, default=True)



class Formation(Base):

    __tablename__ = "formations"

    id = Column(Integer, primary_key=True)

    titre = Column(String)

    type = Column(String)

    active = Column(Boolean, default=True)



class Affectation(Base):

    __tablename__ = "affectations"

    id = Column(Integer, primary_key=True)

    collaborateur_id = Column(
        Integer,
        ForeignKey("collaborateurs.id")
    )

    formation_id = Column(
        Integer,
        ForeignKey("formations.id")
    )