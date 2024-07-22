from sqlalchemy import (Column, Integer, String)

import db


class Producto(db.Base):
    """"Producto:
        Es una clase para ser usada en VentanaPrincipal()
        instanciada en app.py
            args
            nombre : str nombre
            precio : float precio
            stock : float stock
            categoria : str categoria"""

    __tablename__ = 'producto'
    __table_args__ = {'sqlite_autoincrement': True}
    id = Column(Integer, primary_key=True)
    nombre = Column(String(100), nullable=False)
    precio = Column(Integer, nullable=False)
    stock = Column(Integer)
    categoria = Column(String(100))

    def __init__(self, nombre, precio, stock=0, categoria=None):
        self.nombre = nombre
        self.precio = precio
        self.stock = stock
        self.categoria = categoria
        print("Producto creado con éxito")

    def __str__(self):
        return "Categorias: {}-- {}-- {}".format(self.nombre, self.precio, self.stock, self.categoria)
