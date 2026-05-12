# DesignStormDuration-CVC v1.0.0

**Duración de tormenta de diseño a partir de registros horarios de precipitación en un clima andino bimodal**

Red hidrometeorológica CVC — Valle del Cauca, Colombia

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![R ≥ 4.0](https://img.shields.io/badge/R-%E2%89%A54.0-blue.svg)](https://www.r-project.org/)
[![Version](https://img.shields.io/badge/version-1.0.0-green.svg)](https://github.com/MauricioVictoriaN/DesignStormDuration-CVC/releases)

---

## Descripción

DesignStormDuration-CVC v1.0.0 es el material suplementario de reproducibilidad del preprint:

> Victoria Niño, M. J. (2026). *Design storm duration from hourly rainfall records in a bimodal Andean climate.* engrXiv. https://github.com/MauricioVictoriaN/DesignStormDuration-CVC

El script implementa un **marco metodológico completo** para el procesamiento de series horarias de precipitación de la red CVC y la estimación de duraciones de tormenta de diseño con intervalos de confianza al 95 %, integrando cuatro etapas:

1. **Control de calidad y relleno en cascada** — interpolación lineal para brechas ≤ 6 h, seguida de climatología horaria × mes (Paulhus & Kohler 1952) para brechas mayores.
2. **Depuración anual OMM N°168** — criterios C1–C2 adaptados al ciclo de trabajo variable (8–16 h/día) de los registradores CVC, con umbral C1 = 60 % justificado normativamente.
3. **Identificación de tormentas independientes (IETD)** — criterio Cv = 1 de Restrepo-Posada & Eagleson (1982) con cap físico de 6 h para convección tropical andina.
4. **Análisis de frecuencia poblacional** — ajuste de cuatro distribuciones (Exponencial, Gamma, Log-Normal, Weibull), evaluadas por KS + Anderson-Darling + coherencia IETD, con bootstrap IC 95 % (B = 1 000).

Aplicado a la estación La Primavera (Buga, Valle del Cauca, 1 644 m s.n.m., 2016–2025), el marco identifica **1 027 tormentas independientes** y produce duraciones de diseño de **3.86 h (Tr = 2 a)**, **10.0 h (Tr = 10 a)** y **21.7 h (Tr = 100 a)** mediante la distribución Log-Normal (μ_ln = 1.352, σ_ln = 0.743).

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
| `ggplot2` | Generación de gráficos (G1–G14) |
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
```

---

## Archivo de datos de entrada

El script acepta el **formato nativo de exportación CVC**: archivo `.xls` o `.xlsx` con una hoja llamada `VarHorarioTotal`, organizada en bloques mensuales. Cada bloque contiene una fila de etiqueta (mes año), una fila de encabezado con la celda `DÍA\HORA` y 24 columnas de hora (00–23).

Los valores faltantes se reconocen automáticamente en los formatos: `***<`, `NA`, celda vacía y `Nulo`.

> **No se requiere preprocesamiento manual** del archivo de la CVC. El script detecta la estructura automáticamente.

---

## Parámetros de configuración clave (Sección 1)

| Parámetro | Valor caso de estudio | Descripción |
|---|---|---|
| `UMBRAL_LLUVIA` | 0.1 mm/h | Intensidad mínima para clasificar una hora como lluviosa |
| `MAX_PRECIP` | 130 mm/h | Límite físico superior para el Valle del Cauca (IDEAM 2014) |
| `MAX_GAP_RELLENO` | 6 h | Máximo gap para interpolación lineal (Stage 1) |
| `UMBRAL_ANUAL` | 0.60 (60 %) | Disponibilidad anual mínima — criterio C1 OMM 168 §5.4 |
| `UMBRAL_HUM` | 0.50 (50 %) | Disponibilidad temporada húmeda — criterio C2 OMM 168 §5.7 |
| `APLICAR_C3` | `FALSE` | Criterio C3 desactivado (gaps estructurales del registrador CVC) |
| `MESES_HUMEDA` | c(3,4,5,9,10,11) | Temporadas húmedas bimodales Mar–May y Sep–Nov |
| `ANIO_INICIO_UTIL` | 2014 | Primer año con datos reales (excluye período sin cobertura) |
| `IETD_CANDIDATOS` | c(1,2,3,4,5,6,8,10,12,16,20,24) | Valores candidatos IETD en horas |
| `IETD_MAX_FISICO` | 6 h | Cap físico para convección tropical andina (Poveda et al. 2002) |
| `IETD_MIN_PROFUNDIDAD` | 1.0 mm | Profundidad mínima por tormenta (descarta trazas) |

---

## Flujo de procesamiento

```
Archivo Excel CVC (.xls/.xlsx)
        │
        ▼
[§0]  Librerías (instalación automática)
        │
        ▼
[§1]  Configuración  ◄── ÚNICO bloque a modificar por estación
        │
        ▼
[§2]  Carga robusta (detecta formato XLS con interior XLSX)
        │
        ▼
[§3]  Detección automática de estructura
        │   • Localiza columna DÍA\HORA (normalización ASCII/Latin-1)
        │   • Detecta 24 columnas de hora en posición variable
        ▼
[§4]  Parseo de bloques mensuales (algoritmo O(n), vectores pre-asignados)
        │
        ▼
[§5]  Control de calidad
        │   • Clasifica: Válido / Faltante / Negativo / Extremo (> 130 mm/h)
        ▼
[§6]  Relleno en cascada
        │   • Stage 1: interpolación lineal ≤ 6 h
        │   • Stage 2: climatología horaria × mes (192 combinaciones)
        │   • Stage 3: cero residual (gaps largos sin vecinos)
        ▼
[§7]  Depuración OMM N°168
        │   • C1: disponibilidad anual ≥ 60 %  →  VALIDO_C1
        │   • C2: disponibilidad temp. húmeda ≥ 50 %  →  VALIDO_C2
        │   • C3 desactivado (ciclo de trabajo variable del registrador)
        ▼
[§8]  Análisis exploratorio
        │   • Ciclos anual, mensual y diurno
        ▼
[§9]  Validación cruzada del relleno (5 % retiro, set.seed(42))
        │   • Métricas: NSE, PBIAS, RMSE
        ▼
[§10] Módulo IETD — Restrepo-Posada & Eagleson (1982)
        │   • Evalúa 12 candidatos IETD (1–24 h)
        │   • Selecciona h* = argmin |Cv_TBT(h) − 1|, h ≤ 6 h
        │   • Cataloga tormentas: inicio, duración, profundidad, TBT
        ▼
[§11] Gráficos G1–G10 (PDF + PNG 150 dpi)
        ▼
[§12] Excel general (Resultados_*.xlsx, 9 hojas)
        ▼
[§13] Excel IETD (IETD_*.xlsx, 7 hojas)
        ▼
[§14] Reporte final en consola
        ▼
[§15] Análisis de frecuencia de duraciones (set.seed(42))
        │   • Probabilidades empíricas (posición de Weibull, OMM)
        │   • Ajuste: Exponencial, Gamma, Log-Normal, Weibull
        │   • Bondad de ajuste: KS + Anderson-Darling + coherencia IETD
        │   • Cuantiles de diseño Tr = 2–200 años
        │   • Bootstrap IC 95 % (B = 1 000 remuestras)
        │   • Gráficos G11–G14 + hojas Excel adicionales
        ▼
Archivos de salida en RUTA_SALIDA/
```

---

## Salidas del análisis

### Excel de resultados generales (`Resultados_[estacion].xlsx`, 9 hojas)

| Hoja | Contenido |
|---|---|
| `Metadatos` | Parámetros de configuración adoptados y estadísticas clave |
| `QC_Calidad` | Resumen de categorías QC (válido / faltante / negativo / extremo) |
| `Gaps_Serie` | Inventario de brechas: inicio, fin y duración en horas |
| `Estadisticas` | Estadísticas descriptivas de la serie original válida |
| `Totales_Anuales` | Precipitación total anual, máximo horario y disponibilidad |
| `Ciclo_Mensual` | Precipitación media mensual multianual (mm/mes) |
| `Ciclo_Diurno` | Fracción de lluvia e intensidad media por hora del día |
| `Validacion_Relleno` | Métricas de validación cruzada (NSE, PBIAS, RMSE) |
| `Auditoria_OMM` | Decisión C1/C2/Rechazado por año con disponibilidades |

### Excel IETD (`IETD_[estacion].xlsx`, 10 hojas)

| Hoja | Contenido |
|---|---|
| `Metadatos_IETD` | IETD adoptado, Cv, cap físico, tasa de tormentas, TBT medio |
| `Evaluacion_Cv_IETD` | Cv_TBT para cada IETD candidato (fila óptima resaltada) |
| `Catalogo_Tormentas` | Catálogo completo: inicio, fin, duración, profundidad, intensidad |
| `TBT_Tiempos_Entre_Tormentas` | Tiempos entre tormentas en horas, días y ln(TBT) |
| `Estadisticas_Anuales_IETD` | N tormentas, profundidad y duración media por año |
| `Serie_Horaria_Depurada` | Serie horaria depurada con etiqueta de evento por hora |
| `Maximos_Anuales_Duracion` | Máximos anuales de precipitación acumulada para duraciones 1–24 h |
| `Duraciones_Empiricas` | Duraciones ordenadas con probabilidades de Weibull |
| `Bondad_Ajuste_KS` | Estadísticos KS y AD para las cuatro distribuciones |
| `Duraciones_Diseno` | Duraciones de diseño Tr = 2–200 a con IC 95 % y nivel Tier I/II |

### Gráficos

| Archivo | Contenido |
|---|---|
| `G1_Serie_Anual.png` | Precipitación anual con tendencia lineal |
| `G2_Ciclo_Mensual.png` | Ciclo anual de precipitación (mm/mes) |
| `G3_Ciclo_Diurno.png` | Ciclo diurno: fracción de lluvia e intensidad media |
| `G4_Disponibilidad.png` | Disponibilidad anual vs umbral C1 (verde/rojo) |
| `G5_Distribucion.png` | Histograma log-log de intensidades horarias |
| `G6_Heatmap.png` | Mapa de calor precipitación mensual año × mes |
| `G7_Depuracion.png` | Depuración OMM 168: decisión C1/C2/Rechazado por año |
| `G8_IETD_Cv.png` | Curva Cv_TBT vs IETD candidato con punto óptimo |
| `G9_Profundidades.png` | Distribución de profundidades por tormenta (log-x) |
| `G10_TBT.png` | Distribución de tiempos entre tormentas |
| `G11_Papel_Probabilidad.png` | Papel de probabilidad empírico vs distribuciones ajustadas |
| `G12_Curva_Tr.png` | Curva duración–período de retorno con IC 95 % bootstrap |
| `G13_Histograma_Densidad.png` | Histograma de duraciones con densidades ajustadas |
| `G14_QQplot.png` | Diagrama QQ de la distribución seleccionada |

---

## Resultados del caso de estudio

**Estación La Primavera** · Guadalajara de Buga · Valle del Cauca · 1 644 m s.n.m. · 2016–2025

| Variable | Valor |
|---|---|
| Años válidos (OMM 168) | 9 (2016, 2018–2025; excluye 2014, 2015, 2017) |
| Disponibilidad media | 64 % |
| N registros válidos | 56 384 horas |
| Intensidad máxima observada | 123.2 mm/h |
| Precipitación media anual | 2 189 mm |
| Mes más lluvioso | Abril (88 mm/mes) |
| Hora de mayor actividad convectiva | 15:00–16:00 h local |
| IETD adoptado | 6 h (Cv_TBT = 4.47) |
| N tormentas independientes | 1 027 (114 tormentas/año) |
| Profundidad media por tormenta | 19.0 mm |
| Duración media por tormenta | 5.1 h |
| Distribución seleccionada | Log-Normal (μ_ln = 1.352, σ_ln = 0.743) |
| D_KS (Log-Normal) | 0.093 (mínimo entre candidatos) |
| A² Anderson-Darling | 7.1 (mínimo entre candidatos) |

### Duraciones de diseño

| Tr (años) | P | Duración (h) | IC 95 % (h) | Nivel |
|---|---|---|---|---|
| 2 | 0.50 | 3.86 | [3.70, 4.05] | **Tier I** ✓ |
| 5 | 0.80 | 7.22 | [6.86, 7.63] | **Tier I** ✓ |
| 10 | 0.90 | 10.01 | [9.43, 10.68] | **Tier I** ✓ |
| 25 | 0.96 | 13.97 | [13.10, 14.92] | Tier II † |
| 50 | 0.98 | 17.76 | [16.51, 19.24] | Tier II † |
| 100 | 0.99 | 21.74 | [20.08, 23.70] | Tier II † |
| 200 | 0.995 | 26.17 | [24.02, 28.70] | Tier II † |

**Tier I** (Tr ≤ 2 · n_años = 18 a): resultados estadísticamente confiables, dentro del rango observado (Cunnane 1978).

**† Tier II** (Tr > 18 a): extrapolación. El IC 95 % bootstrap captura solo variabilidad muestral paramétrica; no incluye incertidumbre de modelo, no-estacionariedad ni saturación instrumental. Usar como límite inferior pendiente de análisis regional multi-estación.

---

## Rango de aplicabilidad

El marco es directamente transferible a cualquier estación CVC con:

- Formato de archivo: exportación nativa CVC (hoja `VarHorarioTotal`)
- Régimen climático: **bimodal** (Valle del Cauca, Mar–May / Sep–Nov)
- Disponibilidad anual real: **≥ 60 %** (ciclo de trabajo variable del registrador)
- N años válidos: **≥ 5** (resultados Tier I requieren ≥ 9 años)

Para otras zonas climáticas de Colombia, ajustar `IETD_MAX_FISICO` según el mecanismo dominante:

| Zona | IETD_MAX_FISICO |
|---|---|
| Valle del Cauca (bimodal, convectivo) | 6 h |
| Pacífico colombiano (persistente, frontal) | 12 h |
| Cordillera (mixto convectivo-frontal) | 8 h |
| Llanos Orientales (unimodal, convectivo) | 6 h |

---

## Citar este trabajo

**Cita recomendada (formato APA):**

> Victoria Niño, M. J. (2026). *Design storm duration from hourly rainfall records in a bimodal Andean climate* (v1.0.0) [Material suplementario]. GitHub. https://github.com/MauricioVictoriaN/DesignStormDuration-CVC

**Entrada BibTeX:**

```bibtex
@misc{victoriaNino2026design,
  author    = {Victoria Niño, Mauricio Javier},
  title     = {{DesignStormDuration-CVC v1.0.0}: Design storm duration
               from hourly rainfall records in a bimodal Andean climate},
  year      = {2026},
  publisher = {GitHub},
  note      = {Supplementary material. \url{https://github.com/MauricioVictoriaN/DesignStormDuration-CVC}},
  url       = {https://github.com/MauricioVictoriaN/DesignStormDuration-CVC}
}
```

---

## Referencias clave

- Chow, V.T., Maidment, D.R. & Mays, L.W. (1988). *Applied Hydrology*. McGraw-Hill.
- Cunnane, C. (1978). Unbiased plotting positions — a review. *Journal of Hydrology*, 37(3–4), 205–222. https://doi.org/10.1016/0022-1694(78)90017-3
- Efron, B. & Tibshirani, R.J. (1993). *An Introduction to the Bootstrap*. Chapman & Hall.
- IDEAM (2014). *Protocolo para el procesamiento y análisis de información hidrometeorológica*. Bogotá.
- Justus, C.G., Hargraves, W.R., Mikhail, A. & Graber, D. (1978). Methods for estimating wind speed frequency distributions. *Journal of Applied Meteorology*, 17(3), 350–353. https://doi.org/10.1175/1520-0450(1978)017<0350:MFEWSF>2.0.CO;2
- Moriasi, D.N. et al. (2007). Model evaluation guidelines for systematic quantification of accuracy in watershed simulations. *Transactions of the ASABE*, 50(3), 885–900. https://doi.org/10.13031/2013.23153
- OMM (2008). *Guía de Prácticas Hidrológicas*, Vol. I, WMO-No. 168, 6ª ed. §5.4, §5.7, §6.2.
- Paulhus, J.L.H. & Kohler, M.A. (1952). Interpolation of missing precipitation records. *Monthly Weather Review*, 80(8), 129–133.
- Poveda, G. et al. (2002). Seasonality in ENSO-related precipitation, river discharges, soil moisture, and vegetation index in Colombia. *Water Resources Research*, 37(8). https://doi.org/10.1029/2000WR900371
- Restrepo-Posada, P.J. & Eagleson, P.S. (1982). Identification of independent rainstorms. *Journal of Hydrology*, 55(1–4), 303–319. https://doi.org/10.1016/0022-1694(82)90136-6
- Rodríguez-Iturbe, I., Cox, D.R. & Isham, V. (1987). Some models for rainfall based on stochastic point processes. *Proc. R. Soc. London A*, 410, 269–288. https://doi.org/10.1098/rspa.1987.0039
- Scholz, F.W. & Stephens, M.A. (1987). K-sample Anderson-Darling tests. *JASA*, 82(399), 918–924.
- Teegavarapu, R.S.V. & Chandramouli, V. (2005). Improved weighting methods for estimation of missing precipitation records. *Journal of Hydrology*, 312(1–4), 191–206. https://doi.org/10.1016/j.jhydrol.2005.02.015

---

## Licencia

Este proyecto está bajo la licencia **MIT** — ver el archivo [LICENSE](LICENSE) para más detalles.

---

## Contacto

**Mauricio Javier Victoria Niño**  
Investigador Independiente — Cali, Colombia  
📧 hidratecsa@gmail.com  
🔗 ORCID: [0009-0003-4328-5691](https://orcid.org/0009-0003-4328-5691)

---

<p align="center">
  <em>Red hidrometeorológica CVC · Valle del Cauca · Andes colombianos · DesignStormDuration-CVC v1.0.0</em>
</p>
