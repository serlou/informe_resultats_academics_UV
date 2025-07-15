import pandas as pd
import matplotlib.pyplot as plt
import os
import warnings
from config import (
    ASIGNATURAS, ETIQUETAS_RESULTADOS, COLORES_RESULTADOS, 
    TIPOS_CONVOCATORIAS, PATRON_CODIGO_ASIGNATURA, PATRON_GRUPO, 
    PATRON_CARPETA, TEXTOS
)
warnings.filterwarnings("ignore")

# A partir de un archivo de excel, extrae los resultados. Los resultados es una estructura con los campos
# "NP", "SU", "AP", "NO", "EX", "MH".
# Esta información está en la columna M. Cada fila desde la 2 hasta la última fila tiene las etiquetas:
# "No presentat", "Suspès", "Aprovat", "Notable"
# Si se encuentra una etiqueta que no sea esta, se lanza una excepción.
def extraer_resultado_de_excel(filename):
    if not os.path.exists(filename):
        raise FileNotFoundError(f"El archivo {filename} no existe.")
    
    # Inicializar el diccionario de resultados
    resultados = {"NP": 0, "SU": 0, "AP": 0, "NO": 0, "EX": 0, "MH": 0}
    
    try:
        # Usar pandas para leer el archivo Excel sin interpretar cabeceras automáticamente
        # header=None evita que pandas use la primera fila como cabeceras
        df = pd.read_excel(filename, engine=None, header=None)  
        
        # Verificar si existe la columna M (índice 12)
        if len(df.columns) <= 12:
            raise ValueError(f"El archivo {filename} no tiene suficientes columnas (necesita al menos columna M)")
        
        # Obtener la columna M (índice 12)
        column_m = df.iloc[:, 12]  # Columna M es índice 12
        
        # Buscar la fila que contiene "DSP_NOMID1" en la columna M
        fila_inicio = None
        for i, valor in enumerate(column_m):
            if pd.notna(valor) and str(valor).strip() == "DSP_NOMID1":
                fila_inicio = i + 1  # Las calificaciones empiezan en la siguiente fila
                # print(f"Encontrado 'DSP_NOMID1' en fila {i+1} de {filename}")
                break
        
        if fila_inicio is None:
            # Si no se encuentra "DSP_NOMID1", intentar buscar otras variantes o usar comportamiento por defecto
            print(f"Advertencia: No se encontró 'DSP_NOMID1' en {filename}")
            # Mostrar algunas filas de la columna M para debug
            print("Primeras 10 filas de la columna M:")
            for i in range(min(10, len(column_m))):
                valor = column_m.iloc[i]
                print(f"  Fila {i+1}: '{valor}' (tipo: {type(valor)})")
            fila_inicio = 0
        
        # Recorrer los valores de la columna M desde la fila de inicio
        for i in range(fila_inicio, len(column_m)):
            resultado = column_m.iloc[i]
            # Convertir a string y limpiar espacios
            resultado_str = str(resultado).strip() if pd.notna(resultado) else ""
            
            if resultado_str == "No presentat":
                resultados["NP"] += 1
            elif resultado_str == "Suspès":
                resultados["SU"] += 1
            elif resultado_str == "Aprovat":
                resultados["AP"] += 1
            elif resultado_str == "Notable":
                resultados["NO"] += 1
            elif resultado_str == "Excel·lent":
                resultados["EX"] += 1
            elif resultado_str == "Matrícula d'honor" or resultado_str == "Matrícula d'Honor":
                resultados["MH"] += 1
            elif resultado_str == "" or resultado_str == "nan":
                # Ignorar celdas vacías o NaN
                continue
            else:
                # Si encontramos algo que no es una calificación conocida, podría ser el final de los datos
                # Solo mostrar advertencia si no es una celda vacía
                if resultado_str and resultado_str != "nan":
                    print(f"Advertencia en {filename}: Etiqueta desconocida '{resultado_str}' en fila {i+1}")
                
    except Exception as e:
        if "No module named" in str(e):
            raise RuntimeError(f"Error al leer el archivo: {e}. "
                             "Puede que necesites instalar el paquete xlrd para archivos .xls")
        else:
            raise
    
    return resultados

def generar_diagrama_sectores(resultados, titulo="Distribución de Resultados", mostrar=True, guardar_archivo=None):
    """
    Genera un diagrama de sectores a partir de los resultados.
    
    Args:
        resultados (dict): Diccionario con las claves "NP", "SU", "AP", "NO", "EX", "MH"
        titulo (str): Título del gráfico
        mostrar (bool): Si True, muestra el gráfico en pantalla
        guardar_archivo (str): Si se proporciona, guarda el gráfico en este archivo
    """
    # Filtrar solo los resultados que tienen valores > 0
    resultados_filtrados = {k: v for k, v in resultados.items() if v > 0}
    
    if not resultados_filtrados:
        print("No hay datos para mostrar en el diagrama.")
        return
    
    # Etiquetas más descriptivas
    etiquetas_completas = ETIQUETAS_RESULTADOS
    
    # Colores para cada categoría
    colores = COLORES_RESULTADOS
    
    # Preparar datos para el gráfico
    etiquetas = [etiquetas_completas[k] for k in resultados_filtrados.keys()]
    valores = list(resultados_filtrados.values())
    colores_graf = [colores[k] for k in resultados_filtrados.keys()]
    
    # Crear el gráfico
    plt.figure(figsize=(10, 8))
    
    # Crear el diagrama de sectores
    wedges, texts, autotexts = plt.pie(valores, labels=etiquetas, colors=colores_graf, 
                                       autopct='%1.1f%%', startangle=90)
    
    # Mejorar la apariencia
    plt.title(titulo, fontsize=16, fontweight='bold', pad=20)
    
    # Añadir leyenda con conteos
    leyenda_labels = [f"{etiquetas_completas[k]}: {v}" for k, v in resultados_filtrados.items()]
    total_matriculados = sum(resultados_filtrados.values())
    plt.legend(wedges, leyenda_labels, title=f"{TEXTOS['leyenda_estudiantes']}: {total_matriculados}", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    
    # Asegurar que el gráfico sea circular
    plt.axis('equal')
    
    # Ajustar el layout para que quepa la leyenda
    plt.tight_layout()
    
    # Guardar archivo si se especifica
    if guardar_archivo:
        plt.savefig(guardar_archivo, dpi=300, bbox_inches='tight')
        print(f"Gráfico guardado en: {guardar_archivo}")
    
    # Mostrar gráfico si se solicita
    if mostrar:
        plt.show()
    
    return plt.gcf()  # Retornar la figura para uso posterior

def obtener_info_asignatura(filename):
    """
    Extrae información de la asignatura a partir del nombre del archivo.
    
    Args:
        filename (str): Nombre del archivo Excel (ej: "excels/1Q2/34154_A_1Q2.xls")
        
    Returns:
        tuple: (codigo_asignatura, nombre_asignatura, grupo, convocatoria)
    """
    import re
    
    # Extraer código de asignatura del nombre del archivo
    codigo_match = re.search(PATRON_CODIGO_ASIGNATURA, filename)
    if not codigo_match:
        raise ValueError(f"No se pudo extraer el código de asignatura de {filename}")
    
    codigo_asignatura = codigo_match.group(1)
    nombre_asignatura = ASIGNATURAS.get(codigo_asignatura, "Asignatura desconocida")
    
    # Extraer grupo (A, B, etc.) del nombre del archivo
    grupo_match = re.search(PATRON_GRUPO, filename)
    if grupo_match:
        grupo = grupo_match.group(1)
    else:
        grupo = "?"  # Si no se encuentra, usar ?
    
    # Extraer número de convocatoria del nombre de la carpeta, no del archivo
    # Los archivos están en carpetas como: 1Q2, 2Q1, 2Q2, A2
    carpeta_match = re.search(PATRON_CARPETA, filename)
    if carpeta_match:
        carpeta = carpeta_match.group(1)
        convocatoria = TIPOS_CONVOCATORIAS.get(carpeta, {}).get("convocatoria", "1")
    else:
        convocatoria = "1"  # Por defecto si no se puede determinar
    
    return codigo_asignatura, nombre_asignatura, grupo, convocatoria

def generar_titulo_completo(filename):
    """
    Genera un título completo para el gráfico basado en el nombre del archivo.
    
    Args:
        filename (str): Nombre del archivo Excel
        
    Returns:
        str: Título formateado como "CODIGO - NOMBRE ASIGNATURA - Grup X - Convocatoria Y"
    """
    codigo, nombre, grupo, convocatoria = obtener_info_asignatura(filename)
    return f"{codigo} - {nombre} - {TEXTOS['grupo']} {grupo} - {TEXTOS['convocatoria']} {convocatoria}"
