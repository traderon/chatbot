from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import openai
import os
from dotenv import load_dotenv
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from flask import session
from flask import send_file
from io import BytesIO
from datos import Estandar, Desempeño


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
    global grado
    global periodo
    global componente
    global estandar
    global desempeño
    global contador
    global messages
    session.clear()
    """GRADO 1:"""
    

    componente= ["Pensamiento numérico variacional", "Pensamiento Espacial Métrico", "Pensamiento Aleatorio"]
    session["grado"]=""
    session["periodo"]=""
    session["componente"]=""
    session["estandar"]=""
    session["desempeño"]=""
    session["contador"]=0
    #pregunta=""
    #contexto={}
    #preguntas=['Cual/es el/los periodo/s a evaluar','Coloque el nivel obtenido por el estudiante: S(superior), A(Alto), B(Básico), Ba(Bajo)','Cuáles son las barreras actitudinales de la familia','Cuáles son las barreras actitudinales del docente','Cuáles son las barreras curriculares']
    session["preguntas"] = [
        "Seleccione el grado Grado 1 o Grado 2?",
        "Seleccione el período Primer período,Segundo período,Tercer período,Cuarto período?",
        "Seleccione el componente: \n\n"+componente[0]+" \n"+componente[1]+" \n"+componente[2],
        "Seleccione el estándar"+ session.get("grado", ""),
        "Seleccione el desempeño"
        #"Coloque el nivel obtenido por el estudiante: S(superior), A(Alto), B(Básico), Ba(Bajo)",
        #"Cuáles son las barreras actitudinales de la familia",
        #"Cuáles son las barreras actitudinales del docente",
        #"Cuáles son las barreras curriculares"
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

def insertar_parrafo_despues(parrafo, texto):
    nuevo = OxmlElement("w:p")
    parrafo._p.addnext(nuevo)

    nuevo_parrafo = Paragraph(nuevo, parrafo._parent)
    nuevo_parrafo.add_run(texto)

    return nuevo_parrafo

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
    grado=session.get("grado", "")
    periodo=session.get("periodo", "")
    componente=session.get("componente", "")
    estandar=session.get("estandar", "")
    desempeño=session.get("desempeño", "")
    contador=session.get("contador", 0)
    print(preguntas)
    print(messages)
    messages.append({"role": "system", "content": "Eres un asistente educativo experto en ajustes razonables en matemáticas."})
    messages.append({
    "role": "user",
    "content": message
    })
    session.modified = True
    
    if len(preguntas)>=contador and any("Iniciar generación del ajuste razonable" in msg["content"] for msg in messages):
        #messages.append({"role": "user", "content": message})
        if len(preguntas)==contador:
            session["contador"] += 1
            contador=session.get("contador", 0)
            if "Seleccione el desempeño" in pregunta:
                session["desempeño"]=message
                desempeño=session.get("desempeño", "")
        elif len(preguntas)>contador:
            if pregunta !="" and "Seleccione el grado Grado 1 o Grado 2?" == pregunta:
                session["grado"]=message
            elif pregunta !="" and "Seleccione el período Primer período,Segundo período,Tercer período,Cuarto período?" == pregunta:
                session["periodo"]=message
            elif "Seleccione el componente" in pregunta:
                session["componente"]=message
            elif "Seleccione el estándar" in pregunta:
                session["estandar"]=message    
            pregunta=preguntas[contador]
            if "Seleccione el estándar" in pregunta:
                estandares = Estandar[session["grado"]][session["periodo"]][session["componente"]]

                pregunta = "Seleccione el estándar:\n\n"
                
                for i, estandar in enumerate(estandares, start=1):
                    pregunta += f"{i}. {estandar}\n"
            elif "Seleccione el desempeño" in pregunta:
                if session["grado"]=='Grado 2':
                    desempeños = Desempeño[session["grado"]][session["periodo"]]["Pensamiento numérico variacional"]
                else:
                    desempeños = Desempeño[session["grado"]][session["periodo"]][session["componente"]]
                pregunta = "Seleccione el desempeño:\n\n"
                
                for i, (nivel,descripcion) in enumerate(desempeños.items(), start=1):
                    pregunta += f"{i}. {nivel}:{descripcion}\n"         
            # Guardar nuevamente en la sesión
            session["preguntas"] = preguntas
            session["contexto"] = contexto
            session["pregunta"] = pregunta
            session["messages"]= messages
            session["contador"] += 1
            print("SESSION:", dict(session))
            session.modified = True
            print("SESSION:", dict(session))
            #del preguntas[0]
            print(session["contexto"])
            return jsonify({
            "response": pregunta
            })    
    if contador>len(preguntas) and any("Iniciar generación del ajuste razonable" in msg["content"] for msg in messages):
        if len(preguntas)<3:
            contexto[pregunta]=message
        print(contexto)
        #documento = Document()
        doc = Document("PIAR.docx")
        
        for numeroCelda in range(1,4):
            celda = doc.tables[0].cell(1, numeroCelda)
            
            for i, p in enumerate(celda.paragraphs):
                print(i, p.text)
                if p.text.strip() == "PRIMER PERIODO" and numeroCelda==1:
                    insertar_parrafo_despues(
                        p,
                        periodo+"\n\n"+componente+"\n\n"+estandar
                    )
                    break
                if p.text.strip() == "Reformulación:" and numeroCelda==3:
                    insertar_parrafo_despues(
                        p,
                        desempeño
                    )
                    break
                
        doc.save("resultado.docx")
        # Crear tabla de 2 columnas
        #tabla = documento.add_table(rows=1, cols=3)
        #tabla.style = 'Table Grid'
        
        # Encabezados
        #encabezado = tabla.rows[0].cells
        #encabezado[0].text = contexto["Cual/es el/los periodo/s a evaluar"]
        #encabezado[1].text = 'Barreras actitudinales del aprendizaje'
        #encabezado[2].text = 'PRIMER PERIODO Y SEGUNDO PERFIODO'

        
        # Agregar datos
        
        #fila = tabla.add_row().cells
        #fila[0].text = contexto["Coloque el nivel obtenido por el estudiante: S(superior), A(Alto), B(Básico), Ba(Bajo)"]
        #fila[1].text = 'Familia: \n'+str(list(contexto.values())[2])+'\n Docente: \n'+str(list(contexto.values())[3])+'\n Barreras curriculares: \n'+str(list(contexto.values())[4])
        #fila[2].text = 'ADAPTACIÓN CURRICULAR: \n'+"""Priorización:  .por dos periodos académicos 
        #Desarrollar el reconocimiento visual y manipulativo de los números del 1 al 10 mediante el uso de material concreto, pictogramas y rutinas repetitivas, permitiendo que la estudiante identifique cada número y lo asocie con una cantidad correspondiente, sin requerir lenguaje oral o simbólico complejo.Reformulación: 
        #Reconoce y compara cantidades simples (mitad, entero, vacío, lleno) usando material concreto y apoyos visuales, sin necesidad de lenguaje estructurado. Identifica cuándo una cantidad es “una parte” o “todas las partes” en diferentes situaciones manipulativas.
        #"""
        
        #documento.save('reporte.docx')
        #return jsonify({
        #"response": "Gracias por tus respuestas"
        # })
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name="reporte.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
         
    """
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
    """
if __name__ == "__main__":
    app.run(debug=True)