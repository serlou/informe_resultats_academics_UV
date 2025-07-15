import os
import glob
import shutil
from extraer_resultado_de_excel import (
    extraer_resultado_de_excel, 
    generar_diagrama_sectores, 
    generar_titulo_completo,
    obtener_info_asignatura
)
from config import (
    TIPOS_CONVOCATORIAS, ETIQUETAS_RESULTADOS, DIRECTORIO_EXCELS,
    DIRECTORIO_OUTPUT, SUBDIRECTORIO_GRAFICOS, ARCHIVO_LATEX,
    LATEX_CONFIG, TEXTOS, CURSO, AUTOR_INFORME, TITULACIONES
)

def limpiar_outputs_anteriores():
    """
    Elimina toda la carpeta output de ejecuciones anteriores.
    """
    print("🧹 Limpiando outputs de ejecuciones anteriores...")
    
    # Eliminar toda la carpeta output si existe
    if os.path.exists(DIRECTORIO_OUTPUT):
        shutil.rmtree(DIRECTORIO_OUTPUT)
        print(f"  ✅ Eliminada carpeta completa: {DIRECTORIO_OUTPUT}")
    
    print("✅ Limpieza completada\n")

def obtener_archivos_por_carpeta():
    """
    Obtiene todos los archivos .xls organizados por carpetas.
    
    Returns:
        dict: Diccionario con las carpetas como claves y listas de archivos como valores
    """
    carpetas = {}
    
    for carpeta, info in TIPOS_CONVOCATORIAS.items():
        carpetas[carpeta] = {
            "nombre": info["nombre"],
            "archivos": []
        }
    
    for carpeta in carpetas.keys():
        patron = f"{DIRECTORIO_EXCELS}/{carpeta}/*.xls"
        archivos = glob.glob(patron)
        carpetas[carpeta]["archivos"] = sorted(archivos)
    
    return carpetas

def generar_graficos_para_archivo(filename, output_dir=None):
    """
    Genera los gráficos para un archivo específico y retorna la información.
    
    Args:
        filename (str): Ruta del archivo Excel
        output_dir (str): Directorio donde guardar los gráficos
        
    Returns:
        dict: Información del archivo con resultados y ruta del gráfico
    """
    if output_dir is None:
        output_dir = os.path.join(DIRECTORIO_OUTPUT, SUBDIRECTORIO_GRAFICOS)
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Extraer resultados
    resultados = extraer_resultado_de_excel(filename)
    codigo, nombre, grupo, convocatoria = obtener_info_asignatura(filename)
    titulo = generar_titulo_completo(filename)
    
    # Generar nombre del archivo de gráfico
    base_name = os.path.basename(filename).replace('.xls', '')
    grafico_path = os.path.join(output_dir, f"{base_name}.png")
    
    # Generar gráfico
    generar_diagrama_sectores(
        resultados, 
        titulo=titulo, 
        mostrar=False,  # No mostrar en pantalla
        guardar_archivo=grafico_path
    )
    
    return {
        "filename": filename,
        "codigo": codigo,
        "nombre": nombre,
        "grupo": grupo,
        "convocatoria": convocatoria,
        "titulo": titulo,
        "resultados": resultados,
        "grafico_path": grafico_path,
        "total_matriculados": sum(resultados.values())
    }

def generar_tabla_latex(info):
    """
    Genera el código LaTeX para una tabla de resultados.
    
    Args:
        info (dict): Información del archivo con resultados
        
    Returns:
        str: Código LaTeX para la tabla
    """
    resultados = info["resultados"]
    
    # Etiquetas completas
    etiquetas = ETIQUETAS_RESULTADOS
    
    # Generar filas de la tabla
    filas = []
    for key, label in etiquetas.items():
        count = resultados[key]
        if count > 0:  # Solo mostrar categorías con valores > 0
            porcentaje = (count / info["total_matriculados"]) * 100 if info["total_matriculados"] > 0 else 0
            filas.append(f"{label} & {count} & {porcentaje:.1f}\\% \\\\")
    
    # Fila total
    filas.append(f"\\hline")
    filas.append(f"\\textbf{{{TEXTOS['tabla_total']}}} & \\textbf{{{info['total_matriculados']}}} & \\textbf{{100.0\\%}} \\\\")
    
    tabla_latex = f"""
\\begin{{table}}[H]
\\centering
\\caption{{{info['titulo']}}}
\\begin{{tabular}}{{|l|c|c|}}
\\hline
\\textbf{{{TEXTOS['tabla_resultado']}}} & \\textbf{{{TEXTOS['tabla_estudiantes']}}} & \\textbf{{{TEXTOS['tabla_porcentaje']}}} \\\\
\\hline
{chr(10).join(filas)}
\\hline
\\end{{tabular}}
\\end{{table}}
"""
    return tabla_latex

def generar_latex_completo():
    """
    Genera el archivo LaTeX completo con todas las asignaturas.
    """
    # Limpiar outputs de ejecuciones anteriores
    limpiar_outputs_anteriores()
    
    # Obtener archivos organizados por carpetas
    carpetas = obtener_archivos_por_carpeta()
    
    # Generar gráficos y recopilar información
    todas_las_asignaturas = {}
    
    for carpeta, info_carpeta in carpetas.items():
        todas_las_asignaturas[carpeta] = {
            "nombre": info_carpeta["nombre"],
            "asignaturas": []
        }
        
        for archivo in info_carpeta["archivos"]:
            try:
                info_asignatura = generar_graficos_para_archivo(archivo)
                todas_las_asignaturas[carpeta]["asignaturas"].append(info_asignatura)
                print(f"{TEXTOS['procesado']}: {archivo}")
            except Exception as e:
                print(f"{TEXTOS['error_procesando']} {archivo}: {e}")
    
    # Generar contenido LaTeX
    # Crear título dinámico con todas las titulaciones en tamaño \small
    titulo_completo = f"{TEXTOS['titulo_informe']}\\\\\n"
    for titulacion in TITULACIONES:
        titulo_completo += f"\\small {titulacion}\\\\\n"
    titulo_completo = titulo_completo.rstrip("\\\\\n")  # Remover última línea
    
    latex_content = f"""\\documentclass[{LATEX_CONFIG['fontsize']},{LATEX_CONFIG['papersize']}]{{{LATEX_CONFIG['documentclass']}}}
\\usepackage[{LATEX_CONFIG['encoding']}]{{inputenc}}
\\usepackage[{LATEX_CONFIG['language']}]{{babel}}
\\usepackage{{geometry}}
\\usepackage{{graphicx}}
\\usepackage{{float}}
\\usepackage{{array}}
\\usepackage{{booktabs}}
\\usepackage{{longtable}}

\\geometry{{margin={LATEX_CONFIG['margins']}}}

\\title{{{titulo_completo}}}
\\author{{{AUTOR_INFORME}}}
\\date{{\\today}}

\\begin{{document}}

\\maketitle
\\tableofcontents
\\newpage

"""
    
    # Generar secciones por carpeta
    for carpeta, info in todas_las_asignaturas.items():
        if info["asignaturas"]:  # Solo si hay asignaturas
            latex_content += f"""
\\section{{{info['nombre']}}}

"""
            
            for asignatura in info["asignaturas"]:
                latex_content += f"""
\\subsection{{{asignatura['codigo']} - {asignatura['nombre']} - {TEXTOS['grupo']} {asignatura['grupo']}}}

"""
                
                # Añadir tabla
                latex_content += generar_tabla_latex(asignatura)
                
                # Añadir gráfico
                # Convertir ruta absoluta a relativa desde la carpeta output
                grafico_relativo = os.path.relpath(asignatura['grafico_path'], DIRECTORIO_OUTPUT).replace('\\', '/')
                latex_content += f"""
\\begin{{figure}}[H]
\\centering
\\includegraphics[width=0.8\\textwidth]{{{grafico_relativo}}}
\\caption{{{asignatura['titulo']}}}
\\end{{figure}}

\\newpage

"""
    
    # Cerrar documento
    latex_content += """
\\end{document}
"""
    
    # Crear directorio output si no existe
    if not os.path.exists(DIRECTORIO_OUTPUT):
        os.makedirs(DIRECTORIO_OUTPUT)
    
    # Guardar archivo LaTeX
    archivo_completo = os.path.join(DIRECTORIO_OUTPUT, ARCHIVO_LATEX)
    with open(archivo_completo, "w", encoding="utf-8") as f:
        f.write(latex_content)
    
    print(f"{TEXTOS['archivo_generado']}: {archivo_completo}")
    print(f"{TEXTOS['comando_compilar']}")
    
    # Generar resumen
    print(f"\n=== {TEXTOS['resumen']} ===")
    for carpeta, info in todas_las_asignaturas.items():
        print(f"\n{info['nombre']}: {len(info['asignaturas'])} {TEXTOS['assignatures']}")
        for asignatura in info['asignaturas']:
            print(f"  - {asignatura['codigo']} - {TEXTOS['grupo']} {asignatura['grupo']}: {asignatura['total_matriculados']} {TEXTOS['estudiants']}")

if __name__ == "__main__":
    generar_latex_completo()
