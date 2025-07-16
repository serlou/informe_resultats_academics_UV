#!/usr/bin/env python3
"""
Generador de Informe con Barras Apiladas Horizontales
====================================================

Este script genera un informe donde:
- Por cada convocatoria (1Q1, 1Q2, etc.) hay una única tabla con todas las asignaturas
- Un único gráfico de barras apiladas horizontales por convocatoria
- Mucho más compacto que el informe con diagramas de sectores

Autor: Sergio López Ureña - Coordinació 2o curs
"""

import os
import glob
import matplotlib.pyplot as plt
import numpy as np
from collections import defaultdict
import shutil

from extraer_resultado_de_excel import (
    extraer_resultado_de_excel, 
    obtener_info_asignatura
)
from config import (
    TIPOS_CONVOCATORIAS, ETIQUETAS_RESULTADOS, COLORES_RESULTADOS,
    DIRECTORIO_EXCELS, DIRECTORIO_OUTPUT, SUBDIRECTORIO_GRAFICOS, 
    ARCHIVO_LATEX_BARRAS, LATEX_CONFIG, TEXTOS, CURSO, AUTOR_INFORME, 
    TITULACIONES, ASIGNATURAS
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

def obtener_datos_por_convocatoria():
    """
    Obtiene todos los datos organizados por convocatoria.
    
    Returns:
        dict: Diccionario con estructura:
        {
            "1Q1": {
                "asignaturas": {
                    "34154": {
                        "nombre": "Programació matemàtica",
                        "grupos": {
                            "A": {"NP": 5, "SU": 3, ...},
                            "B": {"NP": 2, "SU": 4, ...}
                        }
                    },
                    ...
                }
            },
            ...
        }
    """
    datos_convocatorias = {}
    
    for carpeta, info_conv in TIPOS_CONVOCATORIAS.items():
        print(f"📂 Procesando convocatoria: {info_conv['nombre']}")
        
        # Buscar archivos Excel en esta carpeta
        patron = os.path.join(DIRECTORIO_EXCELS, carpeta, "*.xls")
        archivos = glob.glob(patron)
        
        if not archivos:
            print(f"  ⚠️  No se encontraron archivos en {carpeta}")
            continue
        
        datos_convocatorias[carpeta] = {
            "nombre": info_conv["nombre"],
            "asignaturas": defaultdict(lambda: {"nombre": "", "grupos": {}})
        }
        
        for archivo in archivos:
            try:
                # Extraer información del archivo
                codigo, nombre, grupo, _ = obtener_info_asignatura(archivo)
                
                if codigo not in ASIGNATURAS:
                    print(f"  ⚠️  Código {codigo} no encontrado en configuración")
                    continue
                
                # Extraer resultados
                resultados = extraer_resultado_de_excel(archivo)
                
                # Almacenar datos
                datos_convocatorias[carpeta]["asignaturas"][codigo]["nombre"] = ASIGNATURAS[codigo]
                datos_convocatorias[carpeta]["asignaturas"][codigo]["grupos"][grupo] = resultados
                
                print(f"  ✓ {codigo}_{grupo}: {sum(resultados.values())} estudiantes")
                
            except Exception as e:
                print(f"  ❌ Error procesando {archivo}: {e}")
        
        print(f"  📊 Total asignaturas en {carpeta}: {len(datos_convocatorias[carpeta]['asignaturas'])}\n")
    
    return datos_convocatorias

def generar_grafico_barras_apiladas(datos_convocatoria, titulo, archivo_salida):
    """
    Genera un gráfico de barras apiladas horizontales para una convocatoria.
    
    Args:
        datos_convocatoria (dict): Datos de asignaturas y grupos para una convocatoria
        titulo (str): Título del gráfico
        archivo_salida (str): Ruta donde guardar el gráfico
    """
    if not datos_convocatoria["asignaturas"]:
        print(f"  ⚠️  No hay datos para generar gráfico: {titulo}")
        return
    
    # Preparar datos
    asignaturas_ordenadas = []
    datos_grupos = defaultdict(list)  # {grupo: [valores_por_asignatura]}
    
    # Obtener todas las asignaturas ordenadas por código
    codigos_ordenados = sorted(datos_convocatoria["asignaturas"].keys())
    
    for codigo in codigos_ordenados:
        info_asignatura = datos_convocatoria["asignaturas"][codigo]
        nombre_corto = f"{codigo}\n{info_asignatura['nombre'][:20]}..."
        asignaturas_ordenadas.append(nombre_corto)
        
        # Agregar datos de cada grupo, combinando todos los grupos de la asignatura
        totales_asignatura = {cat: 0 for cat in ETIQUETAS_RESULTADOS.keys()}
        
        for grupo, resultados in info_asignatura["grupos"].items():
            for categoria, valor in resultados.items():
                totales_asignatura[categoria] += valor
        
        # Convertir a porcentajes
        total_estudiantes = sum(totales_asignatura.values())
        if total_estudiantes > 0:
            porcentajes_asignatura = {cat: (valor / total_estudiantes * 100) for cat, valor in totales_asignatura.items()}
        else:
            porcentajes_asignatura = {cat: 0 for cat in ETIQUETAS_RESULTADOS.keys()}
        
        # Agregar porcentajes a las listas
        for categoria in ETIQUETAS_RESULTADOS.keys():
            datos_grupos[categoria].append(porcentajes_asignatura[categoria])
    
    # Crear el gráfico con altura fija y compacta
    fig, ax = plt.subplots(figsize=(12, 2.5))  # Altura reducida a 2.5 pulgadas para mayor compactación

    # Posiciones de las barras
    y_pos = np.arange(len(asignaturas_ordenadas))
    
    # Crear barras apiladas
    left = np.zeros(len(asignaturas_ordenadas))
    
    # Usar las categorías en el mismo orden que las tablas (NP, SU, AP, NO, EX, MH)
    categorias_ordenadas = list(ETIQUETAS_RESULTADOS.keys())
    
    for categoria in categorias_ordenadas:
        valores = datos_grupos[categoria]
        if any(v > 0 for v in valores):  # Solo mostrar si hay datos
            ax.barh(y_pos, valores, left=left, 
                   label=ETIQUETAS_RESULTADOS[categoria],
                   color=COLORES_RESULTADOS[categoria],
                   alpha=0.8, edgecolor='white', linewidth=0.5)
            left += np.array(valores)
    
    # Configurar el gráfico
    ax.set_yticks(y_pos)
    ax.set_yticklabels(asignaturas_ordenadas, fontsize=8)
    ax.set_xlabel('Percentatge d\'estudiants (%)', fontsize=10)
    ax.set_title(titulo, fontsize=12, fontweight='bold', pad=20)
    
    # Configurar eje X para mostrar porcentajes del 0 al 100%
    ax.set_xlim(0, 100)
    ax.set_xticks(range(0, 101, 20))  # Marcas cada 20%
    
    # Leyenda
    ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=9)
    
    # Ajustar layout
    plt.tight_layout()
    
    # Guardar
    plt.savefig(archivo_salida, dpi=300, bbox_inches='tight')
    plt.close()
    
    print(f"  ✓ Gráfico guardado: {os.path.basename(archivo_salida)}")

def generar_tabla_latex_convocatoria(datos_convocatoria, titulo):
    """
    Genera una tabla LaTeX para una convocatoria con todas sus asignaturas.
    
    Args:
        datos_convocatoria (dict): Datos de la convocatoria
        titulo (str): Título de la convocatoria
        
    Returns:
        str: Código LaTeX de la tabla
    """
    if not datos_convocatoria["asignaturas"]:
        return ""
    
    latex = f"""
\\section{{{titulo}}}

\\begin{{table}}[H]
\\centering
\\small
\\begin{{tabular}}{{|p{{4cm}}|c|c|c|c|c|c|c|}}
\\hline
\\textbf{{{TEXTOS["tabla_asignatura"]}}} & """ + " & ".join([f"\\textbf{{{categoria}}}" for categoria in ETIQUETAS_RESULTADOS.keys()]) + f""" & \\textbf{{{TEXTOS["tabla_total"]}}} \\\\
\\hline
"""
    
    # Ordenar asignaturas por código
    codigos_ordenados = sorted(datos_convocatoria["asignaturas"].keys())
    
    totales_generales = {cat: 0 for cat in ETIQUETAS_RESULTADOS.keys()}
    
    for i, codigo in enumerate(codigos_ordenados):
        info_asignatura = datos_convocatoria["asignaturas"][codigo]
        nombre = info_asignatura["nombre"]
        
        # Sumar todos los grupos de esta asignatura
        totales_asignatura = {cat: 0 for cat in ETIQUETAS_RESULTADOS.keys()}
        grupos_str = ""
        
        for grupo, resultados in sorted(info_asignatura["grupos"].items()):
            grupos_str += grupo
            for categoria, valor in resultados.items():
                totales_asignatura[categoria] += valor
                totales_generales[categoria] += valor
        
        total_asignatura = sum(totales_asignatura.values())
        
        # Agregar fila a la tabla con porcentajes
        nombre_completo = f"{codigo} - {nombre}"
        if len(grupos_str) > 1:
            nombre_completo += f" ({grupos_str})"
        
        latex += f"{nombre_completo} & "
        
        # Calcular porcentajes y formatear
        if total_asignatura > 0:
            porcentajes = [f"{(totales_asignatura[cat] / total_asignatura * 100):.1f}\\%" for cat in ETIQUETAS_RESULTADOS.keys()]
        else:
            porcentajes = ["0.0\\%" for cat in ETIQUETAS_RESULTADOS.keys()]
        
        latex += " & ".join(porcentajes)
        latex += f" & {total_asignatura} \\\\\n"
        
        # Agregar separador entre asignaturas (excepto después de la última)
        if i < len(codigos_ordenados) - 1:
            latex += "\\hline\n"
    
    latex += """\\hline
\\end{tabular}
\\caption{Resultats en percentatges}
\\end{table}

"""
    
    return latex

def generar_latex_completo(datos_convocatorias):
    """
    Genera el documento LaTeX completo con todas las convocatorias.
    
    Args:
        datos_convocatorias (dict): Todos los datos organizados por convocatoria
    """
    print("📝 Generando documento LaTeX...")
    
    # Crear directorio de gráficos si no existe
    graficos_dir = os.path.join(DIRECTORIO_OUTPUT, SUBDIRECTORIO_GRAFICOS)
    if not os.path.exists(graficos_dir):
        os.makedirs(graficos_dir)
    
    # Preámbulo LaTeX
    latex_content = f"""\\documentclass[{LATEX_CONFIG["fontsize"]},{LATEX_CONFIG["papersize"]}]{{{LATEX_CONFIG["documentclass"]}}}
\\usepackage[utf8]{{inputenc}}
\\usepackage[catalan]{{babel}}
\\usepackage[margin={LATEX_CONFIG["margins"]}]{{geometry}}
\\usepackage{{graphicx}}
\\usepackage{{float}}
\\usepackage{{booktabs}}
\\usepackage{{array}}
\\usepackage{{longtable}}

\\title{{{TEXTOS["titulo_informe"]} \\\\ {CURSO} \\\\
"""
    
    # Agregar titulaciones al título
    for i, titulacion in enumerate(TITULACIONES):
        latex_content += f"\\small {titulacion}"
        if i < len(TITULACIONES) - 1:
            latex_content += " \\\\\n"
        else:
            latex_content += "}\n"
    
    latex_content += f"""\\author{{{AUTOR_INFORME}}}
\\date{{\\today}}

\\begin{{document}}

\\maketitle

"""
    
    # Procesar cada convocatoria
    for carpeta, datos_conv in datos_convocatorias.items():
        if not datos_conv["asignaturas"]:
            continue
        
        print(f"  📊 Generando contenido para {carpeta}...")
        
        # Generar gráfico de barras apiladas
        nombre_grafico = f"barras_{carpeta}.png"
        archivo_grafico = os.path.join(graficos_dir, nombre_grafico)
        generar_grafico_barras_apiladas(datos_conv, datos_conv["nombre"], archivo_grafico)
        
        # Generar tabla LaTeX
        tabla_latex = generar_tabla_latex_convocatoria(datos_conv, datos_conv["nombre"])
        latex_content += tabla_latex
        
        # Agregar gráfico al LaTeX
        latex_content += f"""
\\begin{{figure}}[H]
\\centering
\\includegraphics[width=0.9\\textwidth]{{{SUBDIRECTORIO_GRAFICOS}/{nombre_grafico}}}
\\caption{{Distribució de resultats - {datos_conv["nombre"]}}}
\\end{{figure}}

\\clearpage

"""
    
    latex_content += "\\end{document}"
    
    # Guardar archivo LaTeX
    archivo_latex = os.path.join(DIRECTORIO_OUTPUT, ARCHIVO_LATEX_BARRAS)
    with open(archivo_latex, 'w', encoding='utf-8') as f:
        f.write(latex_content)
    
    print(f"✅ {TEXTOS['archivo_generado']}: {archivo_latex}")
    print(f"📄 {TEXTOS['comando_compilar_barras']}")

def main():
    """
    Función principal del generador de informe con barras apiladas.
    """
    print("🚀 Iniciando generación de informe con barras apiladas...\n")
    
    # Limpiar outputs anteriores
    limpiar_outputs_anteriores()
    
    # Obtener datos organizados por convocatoria
    datos_convocatorias = obtener_datos_por_convocatoria()
    
    if not datos_convocatorias:
        print("❌ No se encontraron datos para procesar")
        return
    
    # Generar documento LaTeX completo
    generar_latex_completo(datos_convocatorias)
    
    print("\n🎉 ¡Informe con barras apiladas generado exitosamente!")
    print(f"📁 Revisa la carpeta '{DIRECTORIO_OUTPUT}' para ver los resultados")

if __name__ == "__main__":
    main()
