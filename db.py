from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

engine = create_engine("sqlite:///database/productos.db",
                       connect_args={
                           "check_same_thread": False
                       })

Session = sessionmaker(engine)
Base = declarative_base()
session = Session()
