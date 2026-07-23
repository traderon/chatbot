from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import openai
import os
from dotenv import load_dotenv
from docx import Document
from flask import session
from flask import send_file
from io import BytesIO


load_dotenv()

app = Flask(__name__, static_folder='./build')
CORS(app)
#app.secret_key = os.environ.get("SECRET_KEY")
app.secret_key = "prueba123456789"
#client = openai(api_key=os.getenv("OPENAI_API_KEY"))

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def index(path):
    print("Entró a index")
    global preguntas
    global contexto
    global pregunta
    #pregunta=""
    #contexto={}
    #preguntas=['Cual/es el/los periodo/s a evaluar','Coloque el nivel obtenido por el estudiante: S(superior), A(Alto), B(Básico), Ba(Bajo)','Cuáles son las barreras actitudinales de la familia','Cuáles son las barreras actitudinales del docente','Cuáles son las barreras curriculares']
    session["preguntas"] = [
        "Cual/es el/los periodo/s a evaluar",
        "Coloque el nivel obtenido por el estudiante: S(superior), A(Alto), B(Básico), Ba(Bajo)",
        "Cuáles son las barreras actitudinales de la familia",
        "Cuáles son las barreras actitudinales del docente",
        "Cuáles son las barreras curriculares"
    ]
    session["messages"]=[]
    session.modified = True
    
    print("SESSION:", dict(session))
    print("SET COOKIE?", app.session_interface.should_set_cookie(app, session))

    print("SESSION INDEX:", dict(session))
    session["contexto"] = {}
    session["pregunta"] = ""
    if path != "" and os.path.exists(app.static_folder + '/' + path):
        return send_from_directory(app.static_folder, path)
    else:
        return send_from_directory(app.static_folder, 'index.html')

@app.route("/chat", methods=["POST"])
def chat():
    openai.api_key = os.environ.get("OPENAI_SECRET_KEY")
    data = request.json
    message = data.get("message")
    #global contexto
    #global preguntas
    #global pregunta
    print("COOKIES:", request.cookies)
    print("SESSION:", dict(session))
    preguntas = session.get("preguntas", [])
    contexto = session.get("contexto", {})
    pregunta = session.get("pregunta", "")
    messages = session.get("messages", [])
    print(preguntas)
    messages.append({"role": "system", "content": "Eres un asistente educativo experto en ajustes razonables en matemáticas."})
    messages.append({
    "role": "user",
    "content": message
    })
    session.modified = True
    if len(preguntas)>0 and any("Iniciar generación del ajuste razonable" in msg["content"] for msg in messages):
        #messages.append({"role": "user", "content": message})
        
        if len(preguntas)<5:
            if message=='S':
                message="""
                Utiliza correctamente la notación decimal para expresar fracciones en 
                diferentes contextos y relaciono estas dos notaciones con la de los porcentajes.
                """
            elif message=='A':
                message="""
                Utiliza la notación decimal para expresar fracciones en diferentes contextos y 
                relaciono estas dos notaciones con la de los porcentajes.
                """
            elif message=='B':
                message="""
                Con alguna dificultad utiliza la notación decimal para expresar fracciones en diferentes 
                contextos y relaciono estas dos notaciones con la de los porcentajes.
                """
            elif message=='Ba':
                message="""
                Se le dificultad utilizar la notación decimal para expresar fracciones en diferentes 
                contextos y relaciono estas dos notaciones con la de los porcentajes.
                """
            contexto[pregunta]=message
        pregunta=preguntas[0]
        del preguntas[0]
        # Guardar nuevamente en la sesión
        session["preguntas"] = preguntas
        session["contexto"] = contexto
        session["pregunta"] = pregunta
        session.modified = True
        print(session["contexto"])
        return jsonify({
        "response": pregunta
        })    
    elif len(preguntas)<=0 and any("Iniciar generación del ajuste razonable" in msg["content"] for msg in messages):
        if len(preguntas)<3:
            contexto[pregunta]=message
        print(contexto)
        documento = Document()

        # Crear tabla de 2 columnas
        tabla = documento.add_table(rows=1, cols=3)
        tabla.style = 'Table Grid'
        
        # Encabezados
        encabezado = tabla.rows[0].cells
        encabezado[0].text = contexto["Cual/es el/los periodo/s a evaluar"]
        encabezado[1].text = 'Barreras actitudinales del aprendizaje'
        encabezado[2].text = 'PRIMER PERIODO Y SEGUNDO PERFIODO'

        
        # Agregar datos
        
        fila = tabla.add_row().cells
        fila[0].text = contexto["Coloque el nivel obtenido por el estudiante: S(superior), A(Alto), B(Básico), Ba(Bajo)"]
        fila[1].text = 'Familia: \n'+str(list(contexto.values())[2])+'\n Docente: \n'+str(list(contexto.values())[3])+'\n Barreras curriculares: \n'+str(list(contexto.values())[4])
        fila[2].text = 'ADAPTACIÓN CURRICULAR: \n'+"""Priorización:  .por dos periodos académicos 
        Desarrollar el reconocimiento visual y manipulativo de los números del 1 al 10 mediante el uso de material concreto, pictogramas y rutinas repetitivas, permitiendo que la estudiante identifique cada número y lo asocie con una cantidad correspondiente, sin requerir lenguaje oral o simbólico complejo.Reformulación: 
        Reconoce y compara cantidades simples (mitad, entero, vacío, lleno) usando material concreto y apoyos visuales, sin necesidad de lenguaje estructurado. Identifica cuándo una cantidad es “una parte” o “todas las partes” en diferentes situaciones manipulativas.
        """
        
        #documento.save('reporte.docx')
        #return jsonify({
        #"response": "Gracias por tus respuestas"
        # })
        buffer = BytesIO()
        documento.save(buffer)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name="reporte.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
         
    
    response = openai.ChatCompletion.create(
            model="gpt-4",  # Modelo que deseas usar
            #model="gpt-3.5-turbo",
            messages=messages,
            temperature=0.7, 
            max_tokens=2000  # Límite de tokens para la respuesta
        )
        #print(contexto)
        # Obtener la respuesta generada
    bot_reply = response['choices'][0]['message']['content']
   


    return jsonify({
        "response": bot_reply
    })

if __name__ == "__main__":
    app.run(debug=True)