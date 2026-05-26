# DesignStormDuration-CVC v1.0.0

**Duración de tormenta de diseño y coeficiente de avance a partir de registros horarios de precipitación en un clima andino bimodal**
  
  Red hidrometeorológica CVC — Valle del Cauca, Colombia

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![R ≥ 4.0](https://img.shields.io/badge/R-%E2%89%A54.0-blue.svg)](https://www.r-project.org/)
[![Version](https://img.shields.io/badge/version-1.0.0-green.svg)](https://github.com/MauricioVictoriaN/DesignStormDuration-CVC/releases)
[![DOI](https://img.shields.io/badge/DOI-10.31224%2F7062-blue.svg)](https://doi.org/10.31224/7062)

---
  
  ## Descripción
  
  DesignStormDuration-CVC v1.0.0 es el material suplementario de reproducibilidad del preprint Versión 2:
  
  > Victoria Niño, M. J. (2026). *Design storm duration from hourly rainfall records in a bimodal Andean climate.* engrXiv. https://doi.org/10.31224/7062

El script implementa un **marco metodológico completo** para el procesamiento de series horarias de precipitación de la red CVC y la estimación de duraciones de tormenta de diseño con intervalos de confianza al 95 %, integrando **cinco** etapas:
  
  1. **Control de calidad y relleno en cascada** — interpolación lineal para brechas ≤ 6 h, seguida de climatología horaria × mes (Paulhus & Kohler 1952) para brechas mayores.
2. **Depuración anual OMM N°168** — criterios C1–C2 adaptados al ciclo de trabajo variable (8–16 h/día) de los registradores CVC, con umbral C1 = 60 % justificado normativamente.
3. **Identificación de tormentas independientes (IETD)** — criterio Cv = 1 de Restrepo-Posada & Eagleson (1982) con cap físico de 6 h para convección tropical andina.
4. **Coeficiente de avance (r)** — derivación del parámetro de asimetría temporal para cada tormenta independiente, con estadísticos (mediana, media, SD, percentiles) para uso en hietogramas sintéticos.
5. **Análisis de frecuencia poblacional** — ajuste de cuatro distribuciones (Exponencial, Gamma, Log-Normal, Weibull), evaluadas por KS + Anderson-Darling + coherencia IETD, con bootstrap IC 95 % (B = 1 000).

Aplicado a la estación **La Primavera** (Buga, Valle del Cauca, 1 644 m s.n.m., 2016–2025), el marco identifica **1 027 tormentas independientes** y produce:
  - **Duraciones de diseño**: 3.86 h (Tr = 2 a), 10.0 h (Tr = 10 a) y 21.7 h (Tr = 100 a) mediante la distribución Log-Normal (μ_ln = 1.352, σ_ln = 0.743).
- **Coeficiente de avance**: mediana r = 0.167 (media = 0.248, SD = 0.195), indicando que el pico de intensidad ocurre típicamente en el primer 17 % de la duración de la tormenta.

---
  
  ## Novedades en v1.0.0 respecto a la versión preliminar
  
  | Componente | Estado |
  |---|---|
  | Coeficiente de avance (r) | 🆕 **Nuevo** — Sección 10b del script, Tablas 3–4 del manuscrito |
  | Gráfico G15 | 🆕 **Nuevo** — Distribución del coeficiente de avance |
  | Duración física de tormenta | 🆕 **Corregido** — `duracion_h_real = difftime(fin,inicio)+1h` para r |
  | Archivo IETD_*.xlsx | 🆕 **Ampliado** — 10 hojas (nuevas: Estadísticas r, Percentiles r, r vs Duración) |
  | 15 gráficos en un solo PDF | 🆕 **Integrado** — G1–G15 en PDF único + PNGs individuales |
  | Guardado robusto de PNGs | 🆕 **Mejorado** — verificación con tryCatch + file.exists |
  | Versión del manuscrito | **v2** — DOI original preservado (10.31224/7062) |
  
  ---
  
  ## Instalación rápida
  
  ### Prerrequisitos
  
  - **R ≥ 4.0** (recomendado 4.3) — [Descargar R](https://cloud.r-project.org/)
- **RStudio** (recomendado) — [Descargar RStudio](https://posit.co/download/rstudio-desktop/)
- Los paquetes se instalan automáticamente en la primera ejecución.

### Paquetes R requeridos

| Paquete | Uso |
  |---|---|
  | `readxl` | Lectura del archivo Excel de entrada (formato nativo CVC) |
  | `dplyr`, `tidyr` | Manipulación y transformación de datos |
  | `lubridate` | Manejo de fechas y horas |
  | `ggplot2` | Generación de gráficos (G1–G15) |
  | `writexl` | Exportación inicial a Excel |
  | `zoo` | Interpolación lineal de series temporales (`na.approx`) |
  | `scales` | Etiquetas de ejes en gráficos |
  | `openxlsx` | Exportación Excel con formato y estilos |
  | `stringr` | Extracción de mes y año de etiquetas de bloque |
  | `patchwork` | Composición de gráficos múltiples |
  | `gridExtra` | Layouts de gráficos auxiliares |
  
  ### Pasos de ejecución
  
  ```r
# 1. Clonar el repositorio
#    git clone https://github.com/MauricioVictoriaN/DesignStormDuration-CVC.git

# 2. Abrir analisis_precipitacion_horaria_v1.0.0.R en RStudio

# 3. Modificar ÚNICAMENTE la Sección 1 (líneas ~65–110):
RUTA_ARCHIVO <- "D:\\R\\IETD\\Datos_precipitacion_crudos.xls"   # ruta al archivo CVC
RUTA_SALIDA  <- "D:\\R\\IETD\\"                                   # carpeta de resultados
ESTACION     <- "La Primavera"                                    # nombre de la estación

# 4. Ejecutar todo el script (Ctrl+Shift+Enter en RStudio)
#    Los paquetes faltantes se instalan automáticamente.
#    Tiempo estimado de ejecución: 3–8 min (según longitud de la serie).
