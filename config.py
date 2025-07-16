# Configuración del Informe de Resultados Académicos
# =====================================================
# Este archivo contiene toda la configuración específica del curso y titulaciones.
# Modifica los valores según tus necesidades.

# INFORMACIÓN GENERAL
# ===================
CURSO = "2o curs"
AUTOR_INFORME = "Sergio López Ureña - Coordinació 2o curs"

# TITULACIONES
# ============
# Lista de todas las titulaciones (todas se muestran con el mismo tamaño \small)
TITULACIONES = [
    "Grau en Matemàtiques",
    "Doble Grau en Matemàtiques i en Enginyeria Telemàtica",
    "Doble Grau en Matemàtiques i en Enginyeria Informàtica", 
    "Doble Grau en Física i Matemàtiques"
]

# ASIGNATURAS
# ===========
# Diccionario con código de asignatura como clave y nombre como valor
ASIGNATURAS = {
    "34154": "Programació matemàtica",
    "34155": "Àlgebra lineal i geometria II", 
    "34156": "Anàlisi matemàtica II",
    "34161": "Mètodes numèrics per a l'àlgebra lineal",
    "34164": "Topologia",
    "34168": "Estructures algebraiques",
    "34170": "Equacions diferencials ordinàries",
    "34242": "Mecànica I",
    "34245": "Termodinàmica",
    "34251": "Laboratori de termodinàmica",
    "34651": "Ètica. Legislació i professió",
    "34670": "Estructures de dades i algorismes",
    "34879": "Empresa",
    "34885": "Arquitectura de xarxes de computadors",
    "36586": "Anàlisi Matemàtica II F-M",
    "36587": "Àlgebra Lineal i Geometria II F-M",
    "36588": "Equacions Diferencials Ordinàries F-M",
    "36589": "Mètodes Numèrics"
}

# TIPOS DE CONVOCATORIAS
# ======================
# Configuración de las carpetas de archivos Excel y sus descripciones
TIPOS_CONVOCATORIAS = {
    "1Q1": {
        "nombre": "Primer Quadrimestre - Primera Convocatòria",
        "convocatoria": "1"
    },
    "1Q2": {
        "nombre": "Primer Quadrimestre - Segona Convocatòria",
        "convocatoria": "2"
    },
    "2Q1": {
        "nombre": "Segon Quadrimestre - Primera Convocatòria", 
        "convocatoria": "1"
    },
    "2Q2": {
        "nombre": "Segon Quadrimestre - Segona Convocatòria",
        "convocatoria": "2"
    },
    "A1": {
        "nombre": "Assignatures Anuals - Primera Convocatòria",
        "convocatoria": "1"
    },
    "A2": {
        "nombre": "Assignatures Anuals - Segona Convocatòria",
        "convocatoria": "2"
    }
}

# ETIQUETAS DE RESULTADOS
# =======================
# Mapeo de códigos internos a etiquetas completas en catalán
ETIQUETAS_RESULTADOS = {
    "NP": "No presentat",
    "SU": "Suspès", 
    "AP": "Aprovat",
    "NO": "Notable",
    "EX": "Excel·lent",
    "MH": "Matrícula d'Honor"
}

# MAPEO DE CALIFICACIONES
# =======================
# Mapeo de todas las posibles variantes de calificaciones encontradas en los archivos Excel
# hacia los códigos internos del sistema
MAPEO_CALIFICACIONES = {
    "NP": [
        "No presentat",
        "No presentado"
    ],
    "SU": [
        "Suspès", 
        "Suspenso"
    ],
    "AP": [
        "Aprovat",
        "Aprobado"
    ],
    "NO": [
        "Notable"
    ],
    "EX": [
        "Excel·lent",
        "Excelente", 
        "Sobresaliente"
    ],
    "MH": [
        "Matrícula d'honor",
        "Matrícula d'Honor", 
        "Matrícula de Honor",
        "Matrícula de honor"
    ]
}

# COLORES PARA GRÁFICOS
# =====================
# Colores para cada tipo de resultado en los diagramas de sectores
COLORES_RESULTADOS = {
    "NP": "#ff99cc",  # Rosa
    "SU": "#ff0000",  # Rojo
    "AP": "#ffff00",  # Amarillo
    "NO": "#00ff00",  # Verde
    "EX": "#0000ff",  # Azul
    "MH": "#800080"   # Morado
}

# CONFIGURACIÓN DE ARCHIVOS
# =========================
# Directorio base donde están los archivos Excel
DIRECTORIO_EXCELS = "excels"

# Directorio de salida para gráficos e informe
DIRECTORIO_OUTPUT = "output"
SUBDIRECTORIO_GRAFICOS = "graficos"

# Nombre del archivo LaTeX de salida para informe con diagrama de sectores
ARCHIVO_LATEX_SECTORES = "informe_sectores.tex"

# Nombre del archivo LaTeX de salida para informe compacto con barras apiladas
ARCHIVO_LATEX_BARRAS = "informe_barras.tex"

# CONFIGURACIÓN REGEX
# ===================
# Patrones para extraer información de los nombres de archivos
PATRON_CODIGO_ASIGNATURA = r'(\d{5})'       # 5 dígitos para el código
PATRON_GRUPO = r'_([AB])_[^.]*'                # Grupos A, B, etc. (patrón flexible para sufijos)
PATRON_CARPETA = r'excels/([^/]+)/'         # Nombre de la carpeta

# CONFIGURACIÓN LATEX
# ===================
# Configuración del documento LaTeX
LATEX_CONFIG = {
    "documentclass": "article",
    "fontsize": "12pt",
    "papersize": "a4paper",
    "encoding": "utf8",
    "language": "catalan",
    "margins": "2.5cm"
}

# MENSAJES DE TEXTO
# =================
# Textos utilizados en el informe y mensajes
TEXTOS = {
    "titulo_informe": "Informe de Resultats Acadèmics",
    "tabla_resultado": "Resultat",
    "tabla_estudiantes": "Estudiants", 
    "tabla_porcentaje": "Percentatge",
    "tabla_total": "Total",
    "leyenda_estudiantes": "Num. estudiantes",
    "grupo": "Grup",
    "convocatoria": "Convocatòria",
    "archivo_generado": "Arxiu LaTeX generat",
    "comando_compilar": "Per compilar executa: cd output && pdflatex informe_sectores.tex",
    "comando_compilar_barras": "Per compilar executa: cd output && pdflatex informe_barras.tex",
    "resumen": "RESUM",
    "procesado": "Processat",
    "error_procesando": "Error processant",
    "assignatures": "assignatures", 
    "estudiants": "estudiants",
    "tabla_asignatura": "Assignatura",
    "tabla_grupos": "Grups"
}
