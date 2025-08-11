"""
import mysql.connector as mariadb

class Telefonica:
    # Este es el constructor de la clase, alli crearemos la base datos y la tabla
    def __init__(self):
        self.mydb = mariadb.connect(
            host="127.0.0.1",
            user="root",
            password="",
            autocommit=True
        )
        print(self.mydb)  # veo si se conectó bien

        # ---------------------VERIFICO SI LA BD existe--------------------------------------------
        self.mycursor = self.mydb.cursor()
        self.mycursor.execute("DROP database IF EXISTS GTELEFONICA")

        # --------------------------------------/// 1 CREO MI BASE ///---------------------------------
        self.mycursor = self.mydb.cursor()
        self.mycursor.execute("CREATE DATABASE GTELEFONICA")

        # ---------------------VERIFICO SI SE CREO MI BASE--------------------------------------------
        self.mycursor = self.mydb.cursor()
        self.mycursor.execute("SHOW DATABASES")

        for x in self.mycursor:
            print(x)

        # ------------------------------INTENTO CONECTAR A LA BASE CREADA------------------------------
        self.mydb = mariadb.connect(
            host="localhost",
            user="root",
            password="",    # no le puse pass a mi base por el momento
            database="GTELEFONICA"
        )

        # -----------------/// 2 CREO UNA TABLA DENTRO DE LA BASE CON UNA CLAVE PRIMARIA ///--------------
        self.mycursor = self.mydb.cursor()
        self.mycursor.execute(
            "CREATE TABLE clientes(dni VARCHAR(255) PRIMARY KEY,Nombre VARCHAR(255),apellido VARCHAR(255),direccion VARCHAR(255),telefono VARCHAR(255))")

        # --------------------/// 3 INSERTO REGISTROS A MI TABLA ///----------------------------------------
        sql = "INSERT INTO clientes(dni, nombre, apellido, direccion, telefono) VALUES (%s, %s, %s, %s, %s)"
        val = [
            (40234159, "Maria", "Ramirez", "Manzana 78", 1550236598),
            (36598124, "Tomas", "Perez", "Naranjos 54", 1541021487),
            (39584357, "Alicia", "Lopez", "Moreno 39", 1525652541),
            (37852164, "Kevin", "Sanchez", "Flores 62", 1525647136),
            (32854126, "Sabrina", "Ramirez", "Trebol 17", 1547852137),
            (41256327, "Brian", "Martinez", "Luna 85", 1585471274)
        ]
        self.mycursor.executemany(sql, val)
        self.mydb.commit()
        print(self.mycursor.rowcount, "Fueron insertados.")

    # -----------------------------/// 4 CONSULTO EL TELEFONO DE UN CLIENTE ///--------------------------
    def consultar_telefono(self, telefono):
        self.mycursor = self.mydb.cursor()
        sql = "SELECT * FROM clientes where telefono = " + str(telefono) + ";"
        self.mycursor.execute(sql)
        myresultado = self.mycursor.fetchall()
        for ind in myresultado:
            print(ind)

#---------------------------/// 5 AGREGO UN NUEVO CLIENTE CON SUS DATOS ///-------------------------
    def agregar_registro(self, val):
        self.mycursor = self.mydb.cursor()
        sql = "INSERT INTO Clientes(dni, nombre, apellido, direccion, telefono) VALUES (%s, %s, %s, %s, %s)"

        self.mycursor.execute(sql, val)

        self.mydb.commit()
        print(self.mycursor.rowcount, " Registro insertado.")

#---------------------------------/// 6 CAMBIO TELEFONO DEL CLIENTE ///----------------------------
    def cambiar_telefono(self, tele, dni):
        self.mycursor = self.mydb.cursor()
        sql = "UPDATE clientes SET telefono = " + tele + " WHERE dni = " + dni
        self.mycursor.execute(sql)
        self.mydb.commit()
        print(self.mycursor.rowcount, " registros modificados")

    # ---------------------------------VEO EL REGISTRO DESPUES DEL CAMBIO---------------------------------
    def ver_cambio(self, dni):
        self.mycursor = self.mydb.cursor()
        sql = "SELECT * FROM clientes where dni = " + dni
        self.mycursor.execute(sql)
        myresultado = self.mycursor.fetchall()
        for ind in myresultado:
          print(ind)

#------------------------------/// 7 ELIMINO TELEFONO DEL CLIENTE ///-----------------------------------
    def eliminar_por_telefono(self, condicion):
        self.mycursor = self.mydb.cursor()
        sql = "DELETE FROM clientes WHERE telefono LIKE '%" + str(condicion) + "%'"
        self.mycursor.execute(sql)
        self.mydb.commit()
        print(self.mycursor.rowcount, " registros eliminados")

#------------------------/// 8 MUESTRO EL LISTADO DE LOS CLIENTES DE MI TABLA ///------------------------
    def mostrar_todos(self):
        self.mycursor = self.mydb.cursor()
        sql = "SELECT * FROM clientes"
        self.mycursor.execute(sql)
        myresultado = self.mycursor.fetchall()

        print("Registros de la base de datos ...")
        for ind in myresultado:
          print(ind)

if __name__ == '__main__':
    telefonia = Telefonica()
    telefonia.consultar_telefono("1550236598")

    val = (40235258, "Rosa", "Lopez", "Martinez 11", 1525143698)
    telefonia.agregar_registro(val)

    telefonia.cambiar_telefono("1525143689", "40235258")
    telefonia.ver_cambio("40235258")

    telefonia.eliminar_por_telefono(1525652541)

    telefonia.mostrar_todos()

"""


import pdfplumber
import spacy
import sqlite3
from spacy.matcher import PhraseMatcher

# 1. Configuración inicial
nlp = spacy.load("es_core_news_sm")  # Modelo de español
palabra_clave = "PARQUEADERO NÚMERO"
"""
conn = sqlite3.connect('datos_extraidos.db')
cursor = conn.cursor()

# 2. Crear tabla en la base de datos
cursor.execute('''CREATE TABLE IF NOT EXISTS extracciones
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                 contexto TEXT,
                 contenido TEXT)''')
"""

# 3. Función para procesar PDF
def procesar_pdf(ruta_pdf):
    with pdfplumber.open(ruta_pdf) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text() + "\n"

    doc = nlp(texto_completo)
    matcher = PhraseMatcher(nlp.vocab)
    matcher.add("PARQUEADERO", [nlp(palabra_clave)])

    extracciones = []
    for match_id, start, end in matcher(doc):
        span = doc[end:end + 100]  # Tomar 15 tokens después de la palabra clave
        contexto = doc[max(0, start - 5):end + 100].text  # Contexto ampliado
        contenido = span.text

        extracciones.append((contexto, contenido))
    return extracciones  # Valido el de abajo
"""
        # 4. Almacenar en la base de datos
        cursor.execute("INSERT INTO extracciones (contexto, contenido) VALUES (?, ?)",
                       (contexto, contenido))

    conn.commit()
 
    return extracciones
"""

# 5. Ejecución principal
if __name__ == "__main__":
    ruta_pdf = "Doc1_Ejemplo_Extraccion.pdf"  # Cambiar por tu archivo PDF
    resultados = procesar_pdf(ruta_pdf)

    print("\nInformación extraída:")
    for i, (ctx, cont) in enumerate(resultados, 1):
        print(f"\nExtracción {i}:")
        print(f"Contexto: {ctx}")
        print(f"Contenido: {cont}")

    # 6. Cerrar conexión
"""    
    conn.close()
"""