import docx
import os
import re
import unicodedata
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

MODO_CORTE = 'REGEX_CAPITULO'

# ==========================================
# FUNCIONES DE EXTRACCIÓN AVANZADA
# ==========================================

def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Elemento no soportado")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extraer_imagenes_del_parrafo(parrafo, doc_part, dir_media):
    imagenes_md = ""
    # Ampliamos la red para atrapar formatos modernos, antiguos y formas agruadas
    nodos_imagen = parrafo._element.xpath('.//*[local-name()="blip" or local-name()="imagedata" or local-name()="shape" or local-name()="drawing"]')
    
    for nodo in nodos_imagen:
        rIds = nodo.xpath('.//@*[local-name()="embed" or local-name()="id" or local-name()="link"]')
        
        for rId in rIds:
            if rId in doc_part.related_parts:
                try:
                    image_part = doc_part.related_parts[rId]
                    if 'image' in image_part.content_type:
                        ext = image_part.content_type.split('/')[-1]
                        filename = f"img_{image_part.sha1[:8]}.{ext}"
                        filepath = os.path.join(dir_media, filename)
                        
                        if not os.path.exists(filepath):
                            with open(filepath, 'wb') as f:
                                f.write(image_part.blob)
                                
                        img_str = f"\n\n![](media/{filename})\n\n"
                        if img_str not in imagenes_md:
                            imagenes_md += img_str
                except Exception:
                    pass
                
    return imagenes_md

def extraer_texto_integrado(parrafo):
    """
    Lee el párrafo de izquierda a derecha. Si encuentra texto normal, le aplica negritas/cursivas.
    Si encuentra una ecuación, la extrae en la misma línea para que no quede 'vertical' ni desordenada.
    """
    texto_md = ""
    for child in parrafo._element:
        # 1. Si es un bloque de texto normal (w:r)
        if child.tag.endswith('r'):
            run = docx.text.run.Run(child, parrafo)
            
            # --- CORRECCIÓN AQUÍ ---
            # Si run.text es nulo, le asignamos un string vacío para evitar que .strip() colapse
            txt = run.text if run.text is not None else ""
            
            if not txt.strip(): 
                texto_md += txt
                continue
            
            # Aplicar estilos básicos
            if run.bold and run.italic: txt = f"***{txt}***"
            elif run.bold: txt = f"**{txt}**"
            elif run.italic: txt = f"*{txt}*"
            
            texto_md += txt
            
        # 2. Si es una ecuación matemática de Word (m:oMath)
        elif child.tag.endswith('oMath') or child.tag.endswith('oMathPara'):
            # Extraemos todo el texto de la fórmula, eliminando saltos de línea internos
            textos = child.xpath('.//*[local-name()="t"]')
            eq_text = "".join([t.text for t in textos if t.text]).replace('\n', ' ').strip()
            
            if eq_text:
                # Si es un párrafo matemático completo (oMathPara), usamos bloque doble $$
                if 'oMathPara' in child.tag:
                    texto_md += f"\n\n$${eq_text}$$\n\n"
                # Si es una ecuación dentro de la línea de texto (oMath), usamos un solo $
                else:
                    texto_md += f" ${eq_text}$ "
                    
    return texto_md

def tabla_a_markdown(tabla):
    if not tabla.rows: return ""
    md = "\n"
    for i, fila in enumerate(tabla.rows):
        fila_texto = [celda.text.replace('\n', ' ').strip() for celda in fila.cells]
        md += "| " + " | ".join(fila_texto) + " |\n"
        if i == 0:
            md += "| " + " | ".join(['---'] * len(fila.cells)) + " |\n"
    return md + "\n"

# ==========================================
# LÓGICA DE CORTE Y FORMATO
# ==========================================

def limpiar_nombre_archivo(texto):
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    limpio = re.sub(r'[^\w\s-]', '', texto).strip()
    return re.sub(r'[-\s]+', '-', limpio).lower()[:45]

def es_titulo_principal(texto):
    if MODO_CORTE == 'REGEX_CAPITULO':
        return bool(re.match(r'(?i)^cap[íi]tulo\s+.*', texto))
    return False

def es_seccion_final(texto):
    patron = r'(?i)^(anexos?|glosario|autores|sobre los autores|acerca de|índice|agradecimientos|identificación del autor|lista de pares revisores)\b'
    return bool(re.match(patron, texto.strip()))

def generar_yaml(dir_salida, lista_archivos):
    ruta_yaml = os.path.join(dir_salida, '_quarto.yml')
    with open(ruta_yaml, 'w', encoding='utf-8') as f:
        f.write("project:\n  type: book\n  output-dir: _book\n\n")
        f.write("book:\n  title: \"Investigaciones en Gestión del Riesgo\"\n  chapters:\n    - index.qmd\n")
        for archivo in lista_archivos: f.write(f"    - {archivo}\n")
        # Aseguramos que MathJax esté activo para renderizar bien el $ y $$
        f.write("\nformat:\n  html:\n    theme: cosmo\n    toc: true\n    number-sections: true\n    html-math-method: mathjax\n")

def procesar_documento(ruta_docx, dir_salida):
    print("Iniciando extracción nativa unificada (Texto fluido, Ecuaciones inline e Imágenes)...")
    doc = docx.Document(ruta_docx)
    
    dir_media = os.path.join(dir_salida, 'media')
    os.makedirs(dir_media, exist_ok=True)
    
    capitulo_num = 0
    lista_archivos = []
    
    archivo_actual = open(os.path.join(dir_salida, 'index.qmd'), 'w', encoding='utf-8')
    archivo_actual.write("# Preliminares {.unnumbered}\n\n")
    
    flag_post_bibliografia = False

    for block in iter_block_items(doc):
        # 1. TABLAS
        if isinstance(block, Table):
            archivo_actual.write(tabla_a_markdown(block))
            continue
            
        parrafo = block
        texto_base = parrafo.text.strip()
        
        # Extraer imágenes del bloque
        imagenes_md = extraer_imagenes_del_parrafo(parrafo, doc.part, dir_media)
        
        # Extraer texto integrado con ecuaciones
        texto_md = extraer_texto_integrado(parrafo)
        
        if not texto_md.strip() and not imagenes_md: 
            continue
            
        texto_limpio = texto_md.lstrip()
        if texto_limpio.startswith('['): texto_limpio = '\\' + texto_limpio
        
        # 2. EVALUACIÓN DE SECCIONES (Capítulos, Bibliografía, etc.)
        es_biblio = bool(re.match(r'(?i)^(bibliograf[íi]a|referencias|literatura citada)\b', texto_base))
        
        if es_titulo_principal(texto_base):
            flag_post_bibliografia = False
            archivo_actual.close()
            capitulo_num += 1
            nom = f"{capitulo_num:02d}-{limpiar_nombre_archivo(texto_base)}.qmd"
            archivo_actual = open(os.path.join(dir_salida, nom), 'w', encoding='utf-8')
            archivo_actual.write(f"# {texto_base}\n\n")
            lista_archivos.append(nom)
            print(f"-> Capítulo: {nom}")
            
        elif es_biblio:
            flag_post_bibliografia = True
            archivo_actual.write(f"## {texto_limpio}\n\n")
            
        elif flag_post_bibliografia and texto_base and es_seccion_final(texto_base):
            archivo_actual.close()
            capitulo_num += 1
            nom = f"{capitulo_num:02d}-post-{limpiar_nombre_archivo(texto_base)}.qmd"
            archivo_actual = open(os.path.join(dir_salida, nom), 'w', encoding='utf-8')
            archivo_actual.write(f"# {texto_base} {{.unnumbered}}\n\n")
            lista_archivos.append(nom)
            print(f"-> Sección Final: {nom}")
            
        else:
            if texto_limpio:
                estilo = parrafo.style.name.lower()
                if 'heading 1' in estilo or 'título 1' in estilo:
                    archivo_actual.write(f"## {texto_limpio}\n\n")
                elif 'heading 2' in estilo or 'título 2' in estilo:
                    archivo_actual.write(f"### {texto_limpio}\n\n")
                else:
                    archivo_actual.write(f"{texto_limpio}\n\n")
                
        # 3. Inyectar las imágenes debajo del párrafo al que pertenecen
        if imagenes_md:
            archivo_actual.write(imagenes_md)

    if archivo_actual: archivo_actual.close()
    generar_yaml(dir_salida, lista_archivos)
    print("\n¡Procesamiento estandarizado completado con éxito!")

if __name__ == "__main__":
    archivo_entrada = 'LGRD_CAPITULOS_V16.docx'
    carpeta_salida = 'proyecto_libro_quarto'
    if os.path.exists(archivo_entrada):
        procesar_documento(archivo_entrada, carpeta_salida)
    else:
        print("Archivo no encontrado.")