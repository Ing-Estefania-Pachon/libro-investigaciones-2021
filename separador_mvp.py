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
def es_seccion_final(texto):
    patron = r'(?i)^(anexos?|glosario|autores|sobre los autores|acerca de|índice|agradecimientos|identificación del autor|lista de pares revisores)\b'
    return bool(re.match(patron, texto.strip()))

def normalizar_comparacion(texto):
    """Normaliza texto eliminando acentos y convirtiendo a minúsculas para comparaciones robustas."""
    if not texto: return ""
    texto = unicodedata.normalize('NFD', texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != 'Mn')
    return texto.lower().strip()

def generar_yaml(dir_salida, lista_archivos):
    ruta_yaml = os.path.join(dir_salida, '_quarto.yml')
    with open(ruta_yaml, 'w', encoding='utf-8') as f:
        f.write("project:\n  type: book\n  output-dir: _book\n\n")
        f.write("book:\n  title: \"Investigaciones en Gestión del Riesgo\"\n  chapters:\n    - index.qmd\n")
        for archivo in lista_archivos: f.write(f"    - {archivo}\n")
        # Aseguramos que MathJax esté activo y desactivamos numeración de secciones
        f.write("\nformat:\n  html:\n    theme: cosmo\n    toc: true\n    number-sections: false\n    html-math-method: mathjax\n")

def procesar_documento(ruta_docx, dir_salida):
    print("Iniciando extracción nativa unificada (Texto fluido, Ecuaciones inline e Imágenes)...")
    doc = docx.Document(ruta_docx)
    
    dir_media = os.path.join(dir_salida, 'media')
    os.makedirs(dir_media, exist_ok=True)
    
    capitulo_num = 0
    lista_archivos = []
    
    # Estructura para preliminares
    PRELIM_TITLES = ["Presentación del libro", "Página Legal", "Prólogo", "Lista de autores"]
    prelim_buffers = {t: [] for t in PRELIM_TITLES}
    current_prelim_section = None
    buffer_pre_capitulo = []
    
    flag_primer_capitulo_encontrado = False
    flag_post_bibliografia = False
    
    archivo_actual = None

    for block in iter_block_items(doc):
        # 1. IDENTIFICACIÓN DE TEXTO Y TÍTULOS
        if isinstance(block, Table):
            texto_base = ""
            contenido_bloque = tabla_a_markdown(block)
        else:
            parrafo = block
            texto_base = parrafo.text.strip()
            # Extraer imágenes y texto
            imagenes_md = extraer_imagenes_del_parrafo(parrafo, doc.part, dir_media)
            texto_md = extraer_texto_integrado(parrafo)
            if not texto_md.strip() and not imagenes_md: continue
            
            texto_limpio = texto_md.lstrip()
            if texto_limpio.startswith('['): texto_limpio = '\\' + texto_limpio
            
            # Formatear según estilo
            if texto_limpio:
                estilo = parrafo.style.name.lower()
                if 'heading 1' in estilo or 'título 1' in estilo:
                    contenido_bloque = f"## {texto_limpio} {{.unnumbered}}\n\n"
                elif 'heading 2' in estilo or 'título 2' in estilo:
                    contenido_bloque = f"### {texto_limpio} {{.unnumbered}}\n\n"
                else:
                    contenido_bloque = f"{texto_limpio}\n\n"
            else:
                contenido_bloque = ""
            
            if imagenes_md:
                contenido_bloque += imagenes_md

        # 2. LÓGICA DE DISTRIBUCIÓN
        
        # Check if it's a main chapter title
        if es_titulo_principal(texto_base):
            flag_primer_capitulo_encontrado = True
            flag_post_bibliografia = False
            if archivo_actual: archivo_actual.close()
            
            capitulo_num += 1
            nom = f"{capitulo_num:02d}-{limpiar_nombre_archivo(texto_base)}.qmd"
            archivo_actual = open(os.path.join(dir_salida, nom), 'w', encoding='utf-8')
            # Los capítulos conservan su nombre sin prefijo numérico manual, Quarto no los numerará
            archivo_actual.write(f"# {texto_base} {{.unnumbered}}\n\n")
            lista_archivos.append(nom)
            print(f"-> Capítulo: {nom}")
            continue

        if not flag_primer_capitulo_encontrado:
            # Estamos en la zona de preliminares. Buscamos títulos de secciones requeridas.
            texto_norm = normalizar_comparacion(texto_base)
            found_title = None
            for pt in PRELIM_TITLES:
                if texto_norm == normalizar_comparacion(pt):
                    found_title = pt
                    break
            
            if found_title:
                current_prelim_section = found_title
            elif current_prelim_section:
                prelim_buffers[current_prelim_section].append(contenido_bloque)
        else:
            # Estamos después del primer capítulo
            es_biblio = bool(re.match(r'(?i)^(bibliograf[íi]a|referencias|literatura citada)\b', texto_base))
            
            if es_biblio:
                flag_post_bibliografia = True
                archivo_actual.write(f"## {texto_base} {{.unnumbered}}\n\n")
            elif flag_post_bibliografia and texto_base and es_seccion_final(texto_base):
                archivo_actual.close()
                capitulo_num += 1
                nom = f"{capitulo_num:02d}-post-{limpiar_nombre_archivo(texto_base)}.qmd"
                archivo_actual = open(os.path.join(dir_salida, nom), 'w', encoding='utf-8')
                archivo_actual.write(f"# {texto_base} {{.unnumbered}}\n\n")
                lista_archivos.append(nom)
                print(f"-> Sección Final: {nom}")
            else:
                archivo_actual.write(contenido_bloque)

    # 3. ESCRITURA DE PRELIMINARES (Archivos individuales)
    lista_preliminares = []
    
    # El primer preliminar con contenido (preferiblemente Presentación) será el index.qmd
    # ya que Quarto requiere que el primer archivo de capítulos sea index.qmd
    index_asignado = False
    
    for pt in PRELIM_TITLES:
        if prelim_buffers[pt]:
            nom_base = limpiar_nombre_archivo(pt)
            
            if not index_asignado:
                nom_archivo = "index.qmd"
                index_asignado = True
            else:
                nom_archivo = f"prelim-{nom_base}.qmd"
                lista_preliminares.append(nom_archivo)
            
            with open(os.path.join(dir_salida, nom_archivo), 'w', encoding='utf-8') as f_pre:
                f_pre.write(f"# {pt} {{.unnumbered}}\n\n")
                for block_content in prelim_buffers[pt]:
                    f_pre.write(block_content)
            
            print(f"-> Preliminar: {nom_archivo} ({pt})")

    # Si por alguna razón no hubo preliminares, creamos un index.qmd genérico
    if not index_asignado:
        with open(os.path.join(dir_salida, "index.qmd"), 'w', encoding='utf-8') as f_idx:
            f_idx.write("# Introducción {.unnumbered}\n\nContenido en preparación.\n")
        index_asignado = True

    if archivo_actual: archivo_actual.close()
    
    # Combinamos preliminares y capítulos (index.qmd no se añade a lista_archivos porque generar_yaml ya lo incluye)
    todos_los_archivos = lista_preliminares + lista_archivos
    generar_yaml(dir_salida, todos_los_archivos)
    print("\n¡Procesamiento estandarizado completado con éxito!")

if __name__ == "__main__":
    archivo_entrada = 'LGRD_CAPITULOS_V16.docx'
    carpeta_salida = 'proyecto_libro_quarto'
    if os.path.exists(archivo_entrada):
        procesar_documento(archivo_entrada, carpeta_salida)
    else:
        print("Archivo no encontrado.")