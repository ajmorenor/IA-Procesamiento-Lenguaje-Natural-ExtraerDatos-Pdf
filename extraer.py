import pdfplumber
import spacy
import sqlite3
from spacy.matcher import PhraseMatcher

# 1. Configuración inicial
nlp = spacy.load("es_core_news_sm")  # Modelo de español
palabra_clave = "parqueadero"
conn = sqlite3.connect('datos_extraidos.db')
cursor = conn.cursor()

# 2. Crear tabla en la base de datos
cursor.execute('''CREATE TABLE IF NOT EXISTS extracciones
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                 contexto TEXT,
                 contenido TEXT)''')


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
        span = doc[end:end + 15]  # Tomar 15 tokens después de la palabra clave
        contexto = doc[max(0, start - 5):end + 15].text  # Contexto ampliado
        contenido = span.text

        extracciones.append((contexto, contenido))

        # 4. Almacenar en la base de datos
        cursor.execute("INSERT INTO extracciones (contexto, contenido) VALUES (?, ?)",
                       (contexto, contenido))

    conn.commit()
    return extracciones


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
    conn.close()
