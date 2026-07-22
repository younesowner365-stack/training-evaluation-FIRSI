BASE_DIR = Path(__file__).resolve().parent

# =========================================================
# CONFIGURATION BASE DE DONNÉES
# =========================================================

SQLITE_URL = f"sqlite:///{BASE_DIR / 'evaluation.db'}"

DATABASE_URL = os.getenv(
    "DATABASE_URL",
    SQLITE_URL,
)

# Compatibilité PostgreSQL
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace(
        "postgres://",
        "postgresql://",
        1,
    )

# SQLite en local
if DATABASE_URL.startswith("sqlite"):

    engine = create_engine(
        DATABASE_URL,
        connect_args={
            "check_same_thread": False,
        },
    )

# PostgreSQL (Render + Neon)
else:

    engine = create_engine(
        DATABASE_URL,
        pool_pre_ping=True,
        pool_recycle=300,
    )

SessionLocal = sessionmaker(
    bind=engine,
    autoflush=False,
    autocommit=False,
)

Base = declarative_base()