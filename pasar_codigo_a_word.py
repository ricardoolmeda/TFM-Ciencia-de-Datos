import os
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Nombre del archivo Word
file_name = "codigo_python.docx"

# 🔹 Si el archivo ya existe, lo eliminamos
if os.path.exists(file_name):
    os.remove(file_name)
    print(f"🗑 Archivo '{file_name}' eliminado.")

# Código en Python que queremos escribir en Word
codigo_python = """# Dibujar las manos detectadas en la imagen y obtener coordenadas
if results.multi_hand_landmarks:
    num_hands = len(results.multi_hand_landmarks)
    for i, hand_landmarks in enumerate(results.multi_hand_landmarks):
        mp_drawing.draw_landmarks(image_bgr, hand_landmarks, mp_hands.HAND_CONNECTIONS)
        print(f"\n✋ Coordenadas de la mano {i+1}:")
        for idx, landmark in enumerate(hand_landmarks.landmark):
            x, y = int(landmark.x * width), int(landmark.y * height)
            print(f"Punto {idx}: ({x}, {y})")

"""

# Crear un nuevo documento de Word
doc = Document()
doc.add_heading('Código en Python', level=1)  # Agregar título

# Agregar el código con formato de texto monoespaciado
code_paragraph = doc.add_paragraph()
run = code_paragraph.add_run(codigo_python)
run.font.name = 'Courier New'  # Fuente monoespaciada
run.font.size = Pt(10)  # Tamaño de fuente

# Aplicar fondo gris al código
shading_elm = parse_xml(r'<w:shd {} w:fill="D9D9D9"/>'.format(nsdecls('w')))
code_paragraph._element.get_or_add_pPr().append(shading_elm)

# Guardar el documento
doc.save(file_name)
print(f"✅ Documento Word creado: '{file_name}'")
