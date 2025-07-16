# Sistema de Informes de Resultados Académicos

Este sistema permite generar automáticamente **dos tipos de informes** en PDF con tablas y gráficos de resultados académicos a partir de archivos Excel:

1. **Informe con Diagrama de Sectores**: Detallado con un gráfico circular por cada asignatura y grupo
2. **Informe Compacto con Barras Apiladas**: Compacto con una tabla y un gráfico de barras horizontales por convocatoria

## ¿Qué Tipo de Informe Elegir?

### Informe con Diagrama de Sectores (Detallado)
- ✅ **Ideal para**: Análisis detallado por asignatura
- ✅ **Contiene**: Un gráfico circular por cada asignatura-grupo 
- ✅ **Ventajas**: Mucha información visual, fácil de ver distribuciones específicas
- ❌ **Desventajas**: Documento muy largo (puede ser +50 páginas)

### Informe Compacto con Barras Apiladas (Recomendado)
- ✅ **Ideal para**: Reuniones, presentaciones, vista general
- ✅ **Contiene**: Una tabla por convocatoria + un gráfico de barras horizontales
- ✅ **Ventajas**: Documento corto (~6 páginas), fácil comparación entre asignaturas
- ✅ **Mejor para**: Coordinación, reuniones de departamento

## Configuración para Otros Coordinadores

### 1. Archivo de Configuración (`config.py`)

Para adaptar el sistema a tu curso/titulación, solo necesitas modificar el archivo `config.py`. Aquí están los principales parámetros que debes cambiar:

#### Información General
```python
CURSO = "Tu curso"  # Ej: "1er curs", "3er curs"
AUTOR_INFORME = "Tu nombre de coordinación"  # Ej: "Coordinació 1er curs"
```

#### Titulaciones
```python
# Todas las titulaciones se muestran con el mismo tamaño (\small)
TITULACIONES = [
    "Tu primera titulación",
    "Tu segunda titulación",
    "Tu tercera titulación"
]
```

#### Asignaturas
Modifica el diccionario `ASIGNATURAS` con los códigos y nombres de tus asignaturas:
```python
ASIGNATURAS = {
    "12345": "Nombre de asignatura 1",
    "67890": "Nombre de asignatura 2",
    # Añade todas tus asignaturas aquí
}
```

#### Tipos de Convocatorias
Configura las carpetas donde están tus archivos Excel:
```python
TIPOS_CONVOCATORIAS = {
    "carpeta1": {
        "nombre": "Descripción de la convocatoria",
        "convocatoria": "1"  # o "2"
    },
    # Añade más según tu estructura
}
```

### 2. Estructura de Archivos

Organiza tus archivos Excel en la siguiente estructura:
```
excels/
├── carpeta1/
│   ├── codigo_A_periodo.xls
│   ├── codigo_B_periodo.xls
│   └── ...
├── carpeta2/
│   └── ...
```

Los archivos deben seguir el patrón: `CODIGO_GRUPO_PERIODO.xls`
- `CODIGO`: Código de 5 dígitos de la asignatura
- `GRUPO`: A, B, C, etc.
- `PERIODO`: Identificador del período

### 3. Formato de Archivos Excel

Los archivos Excel deben tener:
- Columna M con las calificaciones
- Una fila que contenga "DSP_NOMID1" en la columna M (marca el inicio de datos)
- Las calificaciones deben usar exactamente estas etiquetas (puedes cambiarlas en `config.py`):
  - "No presentat"
  - "Suspès"
  - "Aprovat"
  - "Notable"
  - "Excel·lent"
  - "Matrícula d'Honor"

### 4. Personalización Adicional

#### Colores de Gráficos
Puedes cambiar los colores en `COLORES_RESULTADOS`:
```python
COLORES_RESULTADOS = {
    "NP": "#ff99cc",  # Rosa
    "SU": "#ff0000",  # Rojo
    # ... modifica según tu preferencia
}
```

#### Textos en Otro Idioma
Si necesitas cambiar el idioma, modifica las etiquetas en `ETIQUETAS_RESULTADOS` y `TEXTOS`.

#### Patrones de Archivos
Si tus archivos siguen un patrón diferente, modifica las expresiones regulares:
```python
PATRON_CODIGO_ASIGNATURA = r'(\d{5})'  # Para códigos de 5 dígitos
PATRON_GRUPO = r'_([AB])_'             # Para grupos A, B
```

## Uso del Sistema

### 1. Instalación de Dependencias
```bash
pip install pandas matplotlib openpyxl
```

**Dependencias necesarias:**
- `pandas`: Para leer y procesar archivos Excel
- `matplotlib`: Para generar los gráficos de sectores
- `openpyxl`: Motor para escribir archivos Excel modernos

**Nota:** El sistema es compatible con archivos `.xls` y `.xlsx` automáticamente.

### 2. Generar Informe con Diagrama de Sectores (Detallado)
```bash
python generar_informe_sectores.py
```

### 3. Generar Informe Compacto con Barras Apiladas
```bash
python generar_informe_barras.py
```

### 4. Procesar un Archivo Individual
```python
from extraer_resultado_de_excel import extraer_resultado_de_excel, generar_diagrama_sectores

# Extraer resultados
resultados = extraer_resultado_de_excel("excels/carpeta/archivo.xls")
print(resultados)

# Generar gráfico
generar_diagrama_sectores(resultados, titulo="Mi Gráfico")
```

## Archivos de Salida

### Informe con Diagrama de Sectores
- `output/informe_sectores.tex`: Archivo LaTeX completo
- `output/graficos/`: Gráficos individuales en PNG (uno por asignatura-grupo)
- `output/informe_sectores.pdf`: PDF final (después de compilar LaTeX)

### Informe Compacto con Barras Apiladas  
- `output/informe_barras.tex`: Archivo LaTeX compacto
- `output/graficos/barras_*.png`: Gráficos de barras por convocatoria
- `output/informe_barras.pdf`: PDF final (después de compilar LaTeX)
- `output/informe_sectores.pdf`: PDF final (después de compilar LaTeX)

## Compilación del PDF

Para generar el PDF final:
```bash
cd output
pdflatex informe_resultados.tex
```

## Solución de Problemas

### Error: "No se pudo extraer el código de asignatura"
- Verifica que el nombre del archivo siga el patrón correcto
- Ajusta `PATRON_CODIGO_ASIGNATURA` en `config.py` si es necesario

### Error: "No se encontró DSP_NOMID1"
- Verifica que la columna M del Excel tenga esta etiqueta
- Los datos de calificaciones deben estar después de esta fila

### Colores o etiquetas incorrectas
- Revisa `COLORES_RESULTADOS` y `ETIQUETAS_RESULTADOS` en `config.py`

### Archivos no encontrados
- Verifica que la estructura de carpetas coincida con `TIPOS_CONVOCATORIAS`
- Asegúrate de que `DIRECTORIO_EXCELS` apunte a la carpeta correcta

## Personalización Avanzada

Si necesitas cambios más profundos:

1. **Modificar `extraer_resultado_de_excel.py`**: Para cambiar cómo se procesan los archivos Excel
2. **Modificar `generar_informe_sectores.py`**: Para cambiar la estructura del informe LaTeX
3. **Añadir nuevos tipos de gráficos**: Extender la función `generar_diagrama_sectores()`

## Contacto

Si tienes problemas con la configuración, contacta con el autor original del sistema:
sergio.lopez-urena@uv.es

Copyright © 2025 Sergio López Ureña