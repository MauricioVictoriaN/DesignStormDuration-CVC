# ==============================================================================
# analisis_precipitacion_horaria.R  —  v1.0.0
# ------------------------------------------------------------------------------
# Título:   Design storm duration from hourly rainfall records
#           in a bimodal Andean climate
#           Integrating WMO-168 data screening, cascade infilling,
#           and IETD-based frequency analysis for the CVC hydrometeorological
#           network, Valle del Cauca, Colombia
# ------------------------------------------------------------------------------
# Autor:    Mauricio Javier Victoria Niño
#           Investigador independiente · Cali, Colombia
#           hidratecsa@gmail.com
#           ORCID: 0009-0003-4328-5691
# ------------------------------------------------------------------------------
# Referencia del manuscrito:
#   Victoria Niño, M.J. (2025). Design storm duration from hourly rainfall
#   records in a bimodal Andean climate. Preprint, engrXiv.
#   Material suplementario: https://github.com/MauricioVictoriaN/
#                           DesignStormDuration-CVC
# ------------------------------------------------------------------------------
# Descripción:
#   Marco metodológico para el procesamiento completo de series de precipitación
#   horaria de la red CVC (Valle del Cauca, Colombia). Incluye:
#     0.  Librerías (instalación automática)
#     1.  Configuración  ← ÚNICO bloque a modificar por estación
#     2.  Carga del archivo (formato nativo CVC, .xls/.xlsx)
#     3.  Detección automática de estructura (bloques mes×hora)
#     4.  Parseo de bloques mensuales
#     5.  Control de calidad (válido / faltante / negativo / extremo)
#     6.  Relleno en cascada (interpolación lineal → climatología → cero)
#     7.  Depuración OMM N°168 (criterios C1–C2 adaptados a red CVC)
#     8.  Análisis exploratorio (ciclos anual, mensual, diurno)
#     9.  Validación cruzada del relleno (NSE, PBIAS, RMSE)
#    10.  Módulo IETD — Restrepo-Posada & Eagleson (1982)
#    11.  Gráficos G1–G10 (PDF + PNG, 150 dpi)
#    12.  Exportación Excel general (Resultados_*.xlsx, 9 hojas)
#    13.  Exportación Excel IETD   (IETD_*.xlsx, 7 hojas)
#    14.  Reporte final en consola
#    15.  Análisis de frecuencia de duraciones (distribuciones, KS, AD,
#         bootstrap IC95%, gráficos G11–G14, hojas Excel adicionales)
# ------------------------------------------------------------------------------
# Referencias principales:
#   Restrepo-Posada & Eagleson (1982). J. Hydrology 55:303–319.
#   WMO N°168 (2008). Guide to Hydrological Practices, 6th ed. §5.4, §5.7.
#   Paulhus & Kohler (1952). Monthly Weather Review 80:129–133.
#   Chow, Maidment & Mays (1988). Applied Hydrology. McGraw-Hill.
#   Poveda et al. (2002). Water Resources Research 37(8):2169–2178.
#   Justus et al. (1978). J. Applied Meteorology 17(3):350–353.
# ------------------------------------------------------------------------------
# Licencia:  MIT
# Versión:   1.0.0  |  2025
# Reproducibilidad: set.seed(42) fijado en secciones 9 y 15
# ==============================================================================

# ─────────────────────────────────────────────────────────────────────────────
# 0. LIBRERÍAS
# ─────────────────────────────────────────────────────────────────────────────
paquetes <- c("readxl","dplyr","tidyr","lubridate","ggplot2","writexl",
              "zoo","scales","openxlsx","stringr","patchwork","gridExtra")
nuevos <- paquetes[!paquetes %in% installed.packages()[,"Package"]]
if (length(nuevos)) install.packages(nuevos, dependencies = TRUE)
suppressPackageStartupMessages(lapply(paquetes, library, character.only = TRUE))
cat("OK Librerias cargadas.\n")

# ─────────────────────────────────────────────────────────────────────────────
# 1. CONFIGURACIÓN – ÚNICO BLOQUE A MODIFICAR POR ESTACIÓN
# ─────────────────────────────────────────────────────────────────────────────

# --- Rutas -------------------------------------------------------------------
RUTA_ARCHIVO <- "D:\\R\\IETD\\Datos_precipitacion_crudos.xls"
RUTA_SALIDA  <- "D:\\R\\IETD\\"          # carpeta de resultados (se crea si no existe)

# --- Metadatos de estación ---------------------------------------------------
ESTACION     <- "La Primavera"           # nombre corto para archivos y gráficos
CODIGO       <- "NA"                     # código CVC (si aplica)
MUNICIPIO    <- "Guadalajara de Buga"
DEPARTAMENTO <- "Valle del Cauca"
FUENTE       <- "CVC"
COORD_N      <- NA_real_                 # latitud decimal (positivo Norte)
COORD_E      <- NA_real_                 # longitud decimal (negativo Oeste)
ALTURA_msnm  <- 1644

# --- Parámetros hidrológicos -------------------------------------------------
UMBRAL_LLUVIA   <- 0.1    # mm/h mínimo para considerar lluvia (traza)
MAX_PRECIP      <- 130    # mm/h límite físico para Valle del Cauca
MAX_GAP_RELLENO <- 6      # horas máx a rellenar con interpolación lineal

# --- Depuración OMM 168 (ajustar según disponibilidad real de la estación) ---
UMBRAL_ANUAL     <- 0.60  # C1: disponibilidad anual mínima
UMBRAL_HUM       <- 0.50  # C2: disponibilidad en temporada húmeda mínima
APLICAR_C3       <- FALSE # C3: TRUE activa veto por gap > MAX_GAP_HUM_DIAS
MAX_GAP_HUM_DIAS <- 15L   # C3: solo si APLICAR_C3 = TRUE
MESES_HUMEDA     <- c(3L,4L,5L,9L,10L,11L)  # Mar-May y Sep-Nov (bimodal CVC)
ANIO_INICIO_UTIL <- 2014L # primer año con datos reales (0 = usar todos)

# --- Parámetros IETD ---------------------------------------------------------
# Restrepo-Posada & Eagleson (1982): IETD = tiempo mínimo de período seco
# entre tormentas independientes. Criterio Cv = 1 (tiempos entre llegadas
# siguen distribución exponencial → proceso de Poisson).
# Se prueba un rango de valores candidatos y se selecciona el que minimiza |Cv-1|.
IETD_CANDIDATOS  <- c(1,2,3,4,5,6,8,10,12,16,20,24) # horas a evaluar
IETD_MAX_FISICO  <- 6L   # cap fisico: limite superior para conveccion
                          # tropical (Valle del Cauca). Fuente: literatura
                          # de tormentas convectivas tropicales; Poveda et al.
                          # (2002) ENSO y ciclo diurno en Colombia.
IETD_MIN_LLUVIA <- UMBRAL_LLUVIA  # mm/h mínimo para que una hora "cuente"
IETD_MIN_PROFUNDIDAD <- 1.0       # mm mínimos por tormenta (descartar trazas)

dir.create(RUTA_SALIDA, showWarnings = FALSE, recursive = TRUE)
cat(sprintf("OK Configuracion: %s | %s\n", ESTACION, RUTA_SALIDA))

# ─────────────────────────────────────────────────────────────────────────────
# 2. CARGA DEL ARCHIVO
# ─────────────────────────────────────────────────────────────────────────────
cat("Leyendo archivo Excel...\n")

leer_excel_robusto <- function(ruta, hoja = "VarHorarioTotal") {
  tryCatch(
    read_excel(ruta, sheet = hoja, col_names = FALSE),
    error = function(e) {
      ruta_tmp <- sub("\\.[Xx][Ll][Ss][Xx]?$", "_tmp.xlsx", ruta)
      file.copy(ruta, ruta_tmp, overwrite = TRUE)
      on.exit(unlink(ruta_tmp))
      tryCatch(
        read_excel(ruta_tmp, sheet = hoja, col_names = FALSE),
        error = function(e2) stop(sprintf(
          "No se pudo leer el archivo.\n  Ruta: %s\n  Error: %s\n
  Solucion: abra en Excel y guarde como .xlsx", ruta, e2$message))
      )
    }
  )
}

df_raw <- leer_excel_robusto(RUTA_ARCHIVO)
cat(sprintf("OK Archivo leido: %d filas x %d columnas.\n", nrow(df_raw), ncol(df_raw)))

# ─────────────────────────────────────────────────────────────────────────────
# 3. DETECCIÓN AUTOMÁTICA DE ESTRUCTURA
# ─────────────────────────────────────────────────────────────────────────────
# Convierte a character para evitar problemas de codificación Windows
df_chr <- as.data.frame(
  lapply(df_raw, function(col)
    vapply(col, function(v)
      if (is.null(v) || (length(v)==1L && is.na(v))) NA_character_
      else as.character(v[[1]]), character(1))),
  stringsAsFactors = FALSE)

# Detectar "DIA\HORA" sin regex con tildes (falla en locales Latin-1)
es_dia_hora <- function(x) {
  if (is.na(x) || !nzchar(x)) return(FALSE)
  grepl("DIA", toupper(iconv(x, to = "ASCII//TRANSLIT", sub = "?")), fixed = TRUE)
}

col_dia <- NA_integer_; fila_h1 <- NA_integer_
for (ci in seq_len(ncol(df_chr))) {
  hits <- which(vapply(df_chr[[ci]], es_dia_hora, logical(1)))
  if (length(hits) > 0) { col_dia <- ci; fila_h1 <- hits[1]; break }
}
if (is.na(col_dia)) {
  cat("DIAGNOSTICO primeras 30 filas x 10 cols:\n")
  print(df_chr[1:min(30,nrow(df_chr)), 1:min(10,ncol(df_chr))])
  stop("Encabezado DIA\\HORA no encontrado. Verifique la hoja del archivo.")
}
cat(sprintf("OK Encabezado en col %d, fila %d\n", col_dia, fila_h1))

# Detectar columnas de hora (busca en hasta 4 filas desde el encabezado)
hora_cols <- list()
for (fb in fila_h1:min(nrow(df_chr), fila_h1+3L)) {
  tmp <- list()
  for (ci in seq_len(ncol(df_chr))) {
    h <- suppressWarnings(as.integer(trimws(as.character(df_chr[[ci]][fb]))))
    if (!is.na(h) && h >= 0L && h <= 23L) {
      hs <- sprintf("%02d", h)
      if (!hs %in% names(tmp)) tmp[[hs]] <- ci
    }
  }
  if (length(tmp) >= 20L) { hora_cols <- tmp; cat(sprintf("OK Horas en fila %d\n",fb)); break }
}
hora_cols <- hora_cols[order(as.integer(names(hora_cols)))]
if (length(hora_cols) < 20) stop(sprintf("Solo %d horas detectadas (esperadas 24).", length(hora_cols)))
cat(sprintf("OK %d columnas de hora detectadas.\n", length(hora_cols)))

meses_num <- setNames(1:12,
  c("January","February","March","April","May","June",
    "July","August","September","October","November","December"))

idx_headers <- which(vapply(df_chr[[col_dia]], es_dia_hora, logical(1)))
cat(sprintf("OK %d bloques mensuales detectados.\n", length(idx_headers)))

# ─────────────────────────────────────────────────────────────────────────────
# 4. PARSEO DE BLOQUES MENSUALES
# ─────────────────────────────────────────────────────────────────────────────
parsear_bloque <- function(ib) {
  ih   <- idx_headers[ib]
  ilbl <- max(1L, ih - 1L)
  ifin <- if (ib < length(idx_headers)) idx_headers[ib+1L]-1L else nrow(df_chr)
  if (ih >= ifin) return(NULL)

  txt <- paste(unlist(df_chr[ilbl,], use.names=FALSE), collapse=" ")
  mm  <- stringr::str_extract(txt,
    "January|February|March|April|May|June|July|August|September|October|November|December")
  ay  <- stringr::str_extract(txt, "\\b(19|20)\\d{2}\\b")
  if (is.na(mm)) return(NULL)
  mes  <- meses_num[[mm]]
  anio <- if (!is.na(ay)) as.integer(ay) else NA_integer_

  hn  <- names(hora_cols); nh <- length(hn)
  fd  <- (ih+1L):ifin
  cap <- length(fd)*nh
  va <- integer(cap); vm <- integer(cap); vd <- integer(cap)
  vh <- integer(cap); vv <- character(cap); ptr <- 0L

  for (fi in fd) {
    dr <- trimws(df_chr[[col_dia]][fi])
    if (is.na(dr)) next
    di <- suppressWarnings(as.integer(dr))
    if (is.na(di) || di<1L || di>31L) next
    for (k in seq_len(nh)) {
      ci <- hora_cols[[k]]
      if (ci > ncol(df_chr)) next
      ptr <- ptr+1L
      va[ptr] <- anio; vm[ptr] <- mes; vd[ptr] <- di
      vh[ptr] <- as.integer(hn[k])
      rv <- df_chr[[ci]][fi]
      vv[ptr] <- if (is.na(rv)) NA_character_ else rv
    }
  }
  if (ptr == 0L) return(NULL)
  data.frame(anio=va[1:ptr], mes=vm[1:ptr], dia=vd[1:ptr],
             hora=vh[1:ptr], valor_raw=vv[1:ptr], stringsAsFactors=FALSE)
}

cat("Parseando bloques (puede tardar 1-2 min)...\n")
bloques <- Filter(Negate(is.null), lapply(seq_along(idx_headers), parsear_bloque))
if (!length(bloques)) stop("Ningún bloque extraído. Verifique la hoja 'VarHorarioTotal'.")
df_parsed <- do.call(rbind, bloques); rownames(df_parsed) <- NULL
cat(sprintf("OK %d registros extraídos.\n", nrow(df_parsed)))

# ─────────────────────────────────────────────────────────────────────────────
# 5. CONSTRUCCIÓN Y CONTROL DE CALIDAD DE LA SERIE
# ─────────────────────────────────────────────────────────────────────────────
df_parsed <- df_parsed %>%
  mutate(
    es_faltante = is.na(valor_raw) | valor_raw %in% c("***<","NA","","Nulo"),
    precip_raw  = suppressWarnings(as.numeric(valor_raw))
  ) %>%
  filter(!is.na(anio)) %>%
  rowwise() %>%
  mutate(fecha_posible = tryCatch(
    { make_datetime(anio,mes,dia,hora); TRUE }, error=function(e) FALSE, warning=function(w) FALSE
  )) %>%
  ungroup() %>%
  filter(fecha_posible) %>%
  mutate(fecha_hora = make_datetime(anio,mes,dia,hora)) %>%
  arrange(fecha_hora) %>%
  distinct(fecha_hora, .keep_all = TRUE)

cat(sprintf("OK %d registros tras parseo de fechas.\n", nrow(df_parsed)))

fecha_min <- min(df_parsed$fecha_hora)
fecha_max <- max(df_parsed$fecha_hora)

df_serie <- data.frame(fecha_hora = seq(fecha_min, fecha_max, by="1 hour")) %>%
  left_join(df_parsed %>% select(fecha_hora, precip_raw, es_faltante), by="fecha_hora") %>%
  mutate(
    tipo_dato = case_when(
      is.na(precip_raw) & (is.na(es_faltante)|es_faltante) ~ "Faltante",
      !is.na(precip_raw) & precip_raw < 0                  ~ "Negativo",
      !is.na(precip_raw) & precip_raw > MAX_PRECIP         ~ "Extremo",
      !is.na(precip_raw)                                   ~ "Valido",
      TRUE                                                  ~ "Faltante"
    ),
    precip_qc = if_else(tipo_dato == "Valido", precip_raw, NA_real_)
  )

resumen_qc <- df_serie %>% count(tipo_dato) %>%
  mutate(pct = round(100*n/nrow(df_serie),2))
cat(sprintf("\nPeriodo: %s a %s | Total horas: %d\n",
    format(fecha_min,"%Y-%m-%d"), format(fecha_max,"%Y-%m-%d"), nrow(df_serie)))
print(resumen_qc)

# Análisis de gaps
df_serie <- df_serie %>%
  mutate(es_vacio = is.na(precip_qc),
         grp_gap  = cumsum(!es_vacio | row_number()==1))
gaps_info <- df_serie %>% filter(es_vacio) %>%
  group_by(grp_gap) %>%
  summarise(inicio=first(fecha_hora), fin=last(fecha_hora), dur_h=n(), .groups="drop") %>%
  arrange(desc(dur_h))
cat(sprintf("Gaps: %d | Max: %.1f dias | Mediana: %.0f h\n",
    nrow(gaps_info), max(gaps_info$dur_h)/24, median(gaps_info$dur_h)))

# ─────────────────────────────────────────────────────────────────────────────
# 6. RELLENO DE DATOS
# ─────────────────────────────────────────────────────────────────────────────
cat("\nAplicando métodos de relleno...\n")

# M1: Interpolación lineal (gaps cortos <= MAX_GAP_RELLENO horas)
precip_lineal <- pmax(
  na.approx(df_serie$precip_qc, maxgap=MAX_GAP_RELLENO, na.rm=FALSE), 0, na.rm=TRUE)

# M2: Climatología horaria mes×hora (Paulhus & Kohler 1952)
# Más apropiado que media móvil local para lluvia con muchos ceros
clim_h <- df_serie %>% filter(!is.na(precip_qc)) %>%
  mutate(mc=month(fecha_hora), hc=hour(fecha_hora)) %>%
  group_by(mc,hc) %>% summarise(clim=mean(precip_qc,na.rm=TRUE), .groups="drop")

precip_clim <- df_serie$precip_qc
na_idx <- which(is.na(df_serie$precip_qc))
if (length(na_idx) > 0 && nrow(clim_h) > 0) {
  mc_v <- month(df_serie$fecha_hora[na_idx])
  hc_v <- hour(df_serie$fecha_hora[na_idx])
  for (j in seq_along(na_idx)) {
    v <- clim_h$clim[clim_h$mc==mc_v[j] & clim_h$hc==hc_v[j]]
    if (length(v)==1L && !is.na(v)) precip_clim[na_idx[j]] <- v
  }
}
precip_clim <- pmax(precip_clim, 0, na.rm=TRUE)
cat(sprintf("OK Climatologia: %d combinaciones mes×hora.\n", nrow(clim_h)))

# Serie final: lineal para gaps cortos → climatología → cero residual
df_serie <- df_serie %>%
  mutate(
    precip_lineal  = precip_lineal,
    precip_clim    = as.numeric(precip_clim),
    precip_relleno = case_when(
      !is.na(precip_qc)                              ~ precip_qc,
      !is.na(precip_lineal) & is.na(precip_qc)      ~ precip_lineal,
      !is.na(precip_clim)   & is.na(precip_lineal)  ~ precip_clim,
      TRUE                                           ~ 0
    ),
    metodo_relleno = case_when(
      !is.na(precip_qc)                              ~ "Original",
      !is.na(precip_lineal) & is.na(precip_qc)      ~ "Interp_lineal",
      !is.na(precip_clim)   & is.na(precip_lineal)  ~ "Climatologia",
      TRUE                                           ~ "Cero_gap_largo"
    )
  )

print(df_serie %>% count(metodo_relleno) %>%
  mutate(pct=round(100*n/nrow(df_serie),1)))

# ─────────────────────────────────────────────────────────────────────────────
# 7. DEPURACIÓN HIDROLÓGICA DE AÑOS (OMM 168 / IETD)
# ─────────────────────────────────────────────────────────────────────────────
cat("\n======================================================================\n")
cat("  DEPURACION HIDROLOGICA – OMM N.168 / IDEAM 2014\n")
cat("======================================================================\n")

gap_max_h <- function(x) {
  if (!any(x, na.rm=TRUE)) return(0L)
  r <- rle(x); max(r$lengths[r$values], na.rm=TRUE)
}

anios_eval <- sort(unique(year(df_serie$fecha_hora)))
if (ANIO_INICIO_UTIL > 0) anios_eval <- anios_eval[anios_eval >= ANIO_INICIO_UTIL]
cat(sprintf("Anos a evaluar desde %d: %d\n", ANIO_INICIO_UTIL, length(anios_eval)))

tabla_dep <- do.call(rbind, Filter(Negate(is.null), lapply(anios_eval, function(a) {
  dfa  <- df_serie[year(df_serie$fecha_hora)==a, ]
  n_tot <- nrow(dfa); if (n_tot==0L) return(NULL)
  n_val <- sum(!is.na(dfa$precip_qc))
  disp  <- n_val/n_tot

  dfh  <- dfa[month(dfa$fecha_hora) %in% MESES_HUMEDA, ]
  n_ht <- nrow(dfh); n_hv <- sum(!is.na(dfh$precip_qc))
  disp_h <- if (n_ht>0) n_hv/n_ht else NA_real_

  # Gap máximo dentro de cada mes húmedo (excluye gaps de temporada seca)
  gap_hd <- 0L
  for (mh in MESES_HUMEDA) {
    dm <- dfa[month(dfa$fecha_hora)==mh, ]
    if (nrow(dm)==0) next
    g <- ceiling(gap_max_h(is.na(dm$precip_qc))/24)
    if (g > gap_hd) gap_hd <- g
  }
  gap_ad <- ceiling(gap_max_h(is.na(dfa$precip_qc))/24)

  c1 <- disp  >= UMBRAL_ANUAL
  c2 <- !is.na(disp_h) && disp_h >= UMBRAL_HUM
  c3 <- if (APLICAR_C3) gap_hd > MAX_GAP_HUM_DIAS else FALSE

  dec <- if (c3) "RECHAZADO" else if (c1) "VALIDO_C1" else if (c2) "VALIDO_C2" else "RECHAZADO"

  causa <- if (dec=="RECHAZADO") {
    if (c3) sprintf("Gap %dd en mes humedo > %dd", gap_hd, MAX_GAP_HUM_DIAS)
    else sprintf("Disp %.1f%% < %.0f%% anual y Disp_hum %.1f%% < %.0f%%",
                 100*disp,100*UMBRAL_ANUAL,100*ifelse(is.na(disp_h),0,disp_h),100*UMBRAL_HUM)
  } else ""

  norma <- switch(dec,
    VALIDO_C1 = sprintf("OMM168 par.5.4 (anual>=%.0f%%)", 100*UMBRAL_ANUAL),
    VALIDO_C2 = sprintf("OMM168 par.5.7 (hum>=%.0f%%)",  100*UMBRAL_HUM),
    RECHAZADO = "OMM168 par.5.4/5.7")

  data.frame(Anio=a, N_total=n_tot, N_validos=n_val,
             Disp_anual_pct=round(100*disp,1),
             N_hum_total=n_ht, N_hum_validos=n_hv,
             Disp_hum_pct=round(100*ifelse(is.na(disp_h),0,disp_h),1),
             Gap_max_anual_d=gap_ad, Gap_max_hum_d=gap_hd,
             C1=ifelse(c1,"OK","FALLA"), C2=ifelse(c2,"OK","FALLA"),
             C3=ifelse(c3,"FALLA","OK"),
             Decision=dec, Norma=norma, Causa=causa,
             stringsAsFactors=FALSE)
})))

anios_ok  <- tabla_dep$Anio[tabla_dep$Decision %in% c("VALIDO_C1","VALIDO_C2")]
anios_no  <- tabla_dep$Anio[tabla_dep$Decision=="RECHAZADO"]
n_c1 <- sum(tabla_dep$Decision=="VALIDO_C1")
n_c2 <- sum(tabla_dep$Decision=="VALIDO_C2")
n_no <- sum(tabla_dep$Decision=="RECHAZADO")

cat(sprintf("Validos C1 (>=%.0f%% anual): %d | C2 (>=%.0f%% hum): %d | Rechazados: %d\n",
    100*UMBRAL_ANUAL, n_c1, 100*UMBRAL_HUM, n_c2, n_no))
cat(sprintf("Serie IETD: %d anos [%s]\n",
    length(anios_ok), paste(anios_ok, collapse=", ")))
cat(sprintf("Rechazados: [%s]\n", paste(anios_no, collapse=", ")))

if (length(anios_ok)==0) {
  cat("\nDIAGNOSTICO top disponibilidad:\n")
  print(tabla_dep[order(-tabla_dep$Disp_anual_pct),
        c("Anio","Disp_anual_pct","Disp_hum_pct","Gap_max_hum_d","Decision")])
  stop("Sin anos validos. Ajuste UMBRAL_ANUAL, UMBRAL_HUM o ANIO_INICIO_UTIL.")
}
if (length(anios_ok)<10)
  cat("ADVERTENCIA: <10 anos validos. Confianza estadistica reducida (OMM168 par.6.2)\n")

# Serie depurada
df_dep <- df_serie %>%
  filter(year(fecha_hora) %in% anios_ok) %>%
  mutate(anio=year(fecha_hora), mes=month(fecha_hora),
         dia=day(fecha_hora),   hora=hour(fecha_hora),
         Decision_anio = if_else(anio %in% tabla_dep$Anio[tabla_dep$Decision=="VALIDO_C1"],
                                 "VALIDO_C1","VALIDO_C2"))

cat(sprintf("Registros serie depurada: %d horas (%.1f%% del total)\n",
    nrow(df_dep), 100*nrow(df_dep)/nrow(df_serie)))

# ─────────────────────────────────────────────────────────────────────────────
# 8. ANÁLISIS EXPLORATORIO
# ─────────────────────────────────────────────────────────────────────────────
df_an <- df_serie %>%
  mutate(anio=year(fecha_hora), mes=month(fecha_hora),
         dia=day(fecha_hora),   hora=hour(fecha_hora))

prec_orig <- df_an$precip_relleno[df_an$metodo_relleno=="Original"]

stats <- data.frame(
  Estadistica = c("N validos","N total horas","Media (mm/h)","Mediana (mm/h)",
                  "SD (mm/h)","Maximo (mm/h)","P95 (mm/h)","P99 (mm/h)",
                  "Total acumulado (mm)","Horas con lluvia","Frac lluvia (%)"),
  Valor = c(length(prec_orig), nrow(df_an),
            round(mean(prec_orig,na.rm=TRUE),4),
            round(median(prec_orig,na.rm=TRUE),4),
            round(sd(prec_orig,na.rm=TRUE),4),
            round(max(prec_orig,na.rm=TRUE),2),
            round(quantile(prec_orig,.95,na.rm=TRUE),2),
            round(quantile(prec_orig,.99,na.rm=TRUE),2),
            round(sum(df_an$precip_relleno,na.rm=TRUE),1),
            sum(prec_orig>UMBRAL_LLUVIA,na.rm=TRUE),
            round(100*mean(prec_orig>UMBRAL_LLUVIA,na.rm=TRUE),2)))
cat("\nEstadisticas:\n"); print(stats)

tot_an <- df_an %>% group_by(anio) %>%
  summarise(total_mm=round(sum(precip_relleno,na.rm=TRUE),1),
            max_mm=round(max(precip_relleno,na.rm=TRUE),1),
            horas_lluvia=sum(precip_relleno>UMBRAL_LLUVIA,na.rm=TRUE),
            disp_pct=round(100*mean(metodo_relleno=="Original"),1), .groups="drop")

ciclo_m <- df_an %>% group_by(mes) %>%
  summarise(media_mm_mes=round(mean(precip_relleno,na.rm=TRUE)*24*30.4,1),
            intens=round(mean(precip_relleno[precip_relleno>UMBRAL_LLUVIA],na.rm=TRUE),2),
            frac_pct=round(100*mean(precip_relleno>UMBRAL_LLUVIA,na.rm=TRUE),1),
            .groups="drop") %>% mutate(mes_nom=month.abb[mes])

ciclo_h <- df_an %>% group_by(hora) %>%
  summarise(media=round(mean(precip_relleno,na.rm=TRUE),4),
            frac_pct=round(100*mean(precip_relleno>UMBRAL_LLUVIA,na.rm=TRUE),1),
            intens=round(mean(precip_relleno[precip_relleno>UMBRAL_LLUVIA],na.rm=TRUE),2),
            .groups="drop")

# ─────────────────────────────────────────────────────────────────────────────
# 9. VALIDACIÓN CRUZADA DEL RELLENO
# ─────────────────────────────────────────────────────────────────────────────
set.seed(42)
df_v  <- df_an %>% filter(metodo_relleno=="Original", !is.na(precip_relleno))
idx_t <- sample(nrow(df_v), round(.05*nrow(df_v)))
pos_t <- match(df_v$fecha_hora[idx_t], df_an$fecha_hora)
pqc_v <- df_an$precip_relleno; pqc_v[pos_t] <- NA
obs_t <- df_an$precip_relleno[pos_t]

p_lin <- pmax(na.approx(pqc_v, maxgap=MAX_GAP_RELLENO, na.rm=FALSE), 0, na.rm=TRUE)

p_clim <- pqc_v
nai <- which(is.na(pqc_v))
if (length(nai)>0 && nrow(clim_h)>0) {
  mc2 <- month(df_an$fecha_hora[nai]); hc2 <- hour(df_an$fecha_hora[nai])
  for (j in seq_along(nai)) {
    vc <- clim_h$clim[clim_h$mc==mc2[j] & clim_h$hc==hc2[j]]
    if (length(vc)==1L && !is.na(vc)) p_clim[nai[j]] <- vc
  }
}
p_clim <- pmax(as.numeric(p_clim), 0, na.rm=TRUE)

metricas_fn <- function(obs, pred, nm) {
  ok <- !is.na(pred) & !is.na(obs)
  o  <- obs[ok]; p <- pred[ok]; n <- length(o)
  if (n<10) return(data.frame(Metodo=nm,N=n,RMSE=NA,MAE=NA,PBIAS=NA,NSE=NA,R2=NA))
  rmse  <- sqrt(mean((o-p)^2))
  mae   <- mean(abs(o-p))
  pbias <- 100*sum(p-o)/sum(o)
  nse   <- 1 - sum((o-p)^2)/sum((o-mean(o))^2)
  r2    <- suppressWarnings(cor(o,p)^2)
  data.frame(Metodo=nm, N=n, RMSE=round(rmse,4), MAE=round(mae,4),
             PBIAS_pct=round(pbias,2), NSE=round(nse,4), R2=round(r2,4))
}

metricas <- bind_rows(
  metricas_fn(obs_t, p_lin[pos_t],   "Interp_lineal"),
  metricas_fn(obs_t, p_clim[pos_t],  "Climatologia_horaria"),
  metricas_fn(obs_t, rep(0,length(pos_t)), "Cero"))
cat("\nMetricas validacion cruzada (5%):\n"); print(metricas)
cat("NSE: >0.75 muy bueno | 0.65-0.75 bueno | <0.5 insatisfactorio\n")
cat("PBIAS: |<10%| muy bueno | |<25%| bueno\n")
cat("Nota: NSE negativo es esperado en precipitacion horaria con muchos ceros.\n")
cat("      La climatologia reduce el sesgo sistematico (PBIAS) respecto a interpolacion.\n")

# ─────────────────────────────────────────────────────────────────────────────
# 10. MÓDULO IETD – RESTREPO-POSADA & EAGLESON (1982)
# ─────────────────────────────────────────────────────────────────────────────
cat("\n======================================================================\n")
cat("  ANALISIS IETD – RESTREPO-POSADA & EAGLESON (1982)\n")
cat("  Criterio: Cv = 1 (proceso Poisson de tormentas independientes)\n")
cat("======================================================================\n")

# Función principal: identifica tormentas dado un IETD candidato
identificar_tormentas <- function(precip_h, fechas, ietd_h) {
  # Una hora es "lluviosa" si supera el umbral
  lluvia <- !is.na(precip_h) & precip_h >= IETD_MIN_LLUVIA

  # Agrupar horas lluviosas: nuevo evento si gap seco >= ietd_h
  evento <- integer(length(lluvia))
  id_ev  <- 0L
  en_ev  <- FALSE
  ult_lluvia <- -Inf

  for (i in seq_along(lluvia)) {
    if (lluvia[i]) {
      if (!en_ev || (i - ult_lluvia) >= ietd_h) {
        id_ev <- id_ev + 1L
        en_ev <- TRUE
      }
      evento[i] <- id_ev
      ult_lluvia <- i
    } else {
      if (en_ev && (i - ult_lluvia) >= ietd_h) en_ev <- FALSE
    }
  }

  # Resumir tormentas
  df_ev <- data.frame(
    idx       = seq_along(lluvia),
    fecha     = fechas,
    precip    = precip_h,
    evento_id = evento
  ) %>% filter(evento_id > 0) %>%
    group_by(evento_id) %>%
    summarise(
      inicio       = first(fecha),
      fin          = last(fecha),
      duracion_h   = n(),
      profundidad_mm = sum(precip, na.rm=TRUE),
      intensidad_max = max(precip, na.rm=TRUE),
      .groups = "drop"
    ) %>%
    filter(profundidad_mm >= IETD_MIN_PROFUNDIDAD)

  # Tiempos entre inicios de tormentas (horas)
  if (nrow(df_ev) < 2) return(list(tormentas=df_ev, cv=NA_real_, n=nrow(df_ev)))
  tbt <- as.numeric(difftime(df_ev$inicio[-1], df_ev$inicio[-nrow(df_ev)], units="hours"))
  tbt <- tbt[tbt > 0]
  cv  <- if (length(tbt)>=2) sd(tbt)/mean(tbt) else NA_real_
  list(tormentas=df_ev, tbt=tbt, cv=cv, n=nrow(df_ev),
       media_tbt=mean(tbt), sd_tbt=sd(tbt))
}

# Aplicar sobre serie depurada únicamente
precip_dep <- df_dep$precip_relleno
fechas_dep <- df_dep$fecha_hora

# Evaluar todos los IETD candidatos
cat(sprintf("Evaluando %d valores IETD sobre %d horas de datos...\n",
    length(IETD_CANDIDATOS), length(precip_dep)))

res_ietd <- lapply(IETD_CANDIDATOS, function(ietd_h) {
  r <- identificar_tormentas(precip_dep, fechas_dep, ietd_h)
  data.frame(
    IETD_h       = ietd_h,
    N_tormentas  = r$n,
    Media_TBT_h  = round(ifelse(is.null(r$media_tbt), NA, r$media_tbt), 2),
    SD_TBT_h     = round(ifelse(is.null(r$sd_tbt),    NA, r$sd_tbt),    2),
    Cv           = round(ifelse(is.na(r$cv), NA, r$cv), 4),
    Dif_Cv_1     = round(ifelse(is.na(r$cv), NA, abs(r$cv - 1)), 4)
  )
})
tabla_ietd <- do.call(rbind, res_ietd)
cat("\nResultados por IETD candidato:\n"); print(tabla_ietd)

# Seleccionar IETD óptimo: el que minimiza |Cv - 1|
# con restriccion fisica: IETD <= IETD_MAX_FISICO (conveccion tropical)
tabla_ietd_fis <- tabla_ietd[tabla_ietd$IETD_h <= IETD_MAX_FISICO, ]

if (nrow(tabla_ietd_fis) > 0 &&
    any(!is.na(tabla_ietd_fis$Dif_Cv_1))) {
  # Hay candidatos dentro del rango fisico: elegir el mejor Cv
  idx_opt  <- which.min(tabla_ietd_fis$Dif_Cv_1)
  IETD_OPT <- tabla_ietd_fis$IETD_h[idx_opt]
  CV_OPT   <- tabla_ietd_fis$Cv[idx_opt]
  cat(sprintf("OK IETD dentro del rango fisico (<= %dh)\n", IETD_MAX_FISICO))
} else {
  # Sin candidatos validos en rango fisico: usar el menor IETD disponible
  idx_opt  <- which.min(tabla_ietd$IETD_h)
  IETD_OPT <- tabla_ietd$IETD_h[idx_opt]
  CV_OPT   <- tabla_ietd$Cv[idx_opt]
  cat(sprintf("AVISO: ningun IETD candidato <= %dh. Usando IETD=%dh.\n",
      IETD_MAX_FISICO, IETD_OPT))
}

cat(sprintf("\nTabla Cv por IETD (rango fisico marcado con *):\n"))
tabla_ietd_print <- tabla_ietd %>%
  mutate(Fisico = ifelse(IETD_h <= IETD_MAX_FISICO, "*", " "),
         Optimo = ifelse(IETD_h == IETD_OPT, "<<< OPTIMO", ""))
print(tabla_ietd_print)

cat(sprintf("\n>>> IETD OPTIMO (con cap fisico %dh): %d horas (Cv=%.4f)\n",
    IETD_MAX_FISICO, IETD_OPT, CV_OPT))
cat(sprintf("    Cv sin cap fisico seria: %dh (Cv=%.4f)\n",
    tabla_ietd$IETD_h[which.min(tabla_ietd$Dif_Cv_1)],
    tabla_ietd$Cv[which.min(tabla_ietd$Dif_Cv_1)]))

# Advertencia si Cv se aleja significativamente de 1
if (CV_OPT > 1.5) {
  cat(sprintf("\nADVERTENCIA: Cv=%.3f >> 1. El proceso de tormentas NO es Poisson.\n", CV_OPT))
  cat("  Causa probable: agrupamiento estacional (regimen bimodal).\n")
  cat("  El IETD se adopta por restriccion fisica, no por criterio estadistico.\n")
  cat("  Documente esta limitacion en el informe.\n")
} else if (CV_OPT < 0.7) {
  cat(sprintf("\nADVERTENCIA: Cv=%.3f << 1. Proceso mas regular que Poisson.\n", CV_OPT))
  cat("  Posible sub-separacion de tormentas (IETD muy corto).\n")
} else {
  cat(sprintf("\nOK: Cv=%.3f cerca de 1. Proceso compatible con Poisson.\n", CV_OPT))
}

# Extraer tormentas con el IETD óptimo
res_opt  <- identificar_tormentas(precip_dep, fechas_dep, IETD_OPT)
df_torm  <- res_opt$tormentas
tbt_opt  <- res_opt$tbt

# Agregar columnas útiles para análisis posterior
df_torm <- df_torm %>%
  mutate(
    anio             = year(inicio),
    mes              = month(inicio),
    intensidad_media = round(profundidad_mm / duracion_h, 4),
    TBT_horas        = c(NA, round(tbt_opt, 2))  # tiempo desde tormenta anterior
  )

cat(sprintf("\nEstadisticas de tormentas (IETD = %d h):\n", IETD_OPT))
cat(sprintf("  N tormentas:           %d\n", nrow(df_torm)))
cat(sprintf("  Duracion media:        %.1f h\n", mean(df_torm$duracion_h)))
cat(sprintf("  Profundidad media:     %.1f mm\n", mean(df_torm$profundidad_mm)))
cat(sprintf("  Intensidad max global: %.1f mm/h\n", max(df_torm$intensidad_max)))
cat(sprintf("  TBT medio:             %.1f h (%.1f dias)\n",
    mean(tbt_opt,na.rm=TRUE), mean(tbt_opt,na.rm=TRUE)/24))
cat(sprintf("  Tasa media:            %.2f tormentas/año\n",
    nrow(df_torm)/length(anios_ok)))

# Serie horaria con etiqueta de evento para IETD
df_serie_ietd <- data.frame(
  fecha_hora     = fechas_dep,
  precip_mm_h    = precip_dep,
  metodo_relleno = df_dep$metodo_relleno,
  anio           = year(fechas_dep),
  mes            = month(fechas_dep),
  dia            = day(fechas_dep),
  hora           = hour(fechas_dep)
)
# Asignar evento_id a cada hora
ev_vec <- integer(length(precip_dep))
for (i in seq_len(nrow(df_torm))) {
  idx_ev <- which(fechas_dep >= df_torm$inicio[i] & fechas_dep <= df_torm$fin[i])
  ev_vec[idx_ev] <- df_torm$evento_id[i]
}
df_serie_ietd$evento_id   <- ev_vec
df_serie_ietd$en_tormenta <- ev_vec > 0

# ─────────────────────────────────────────────────────────────────────────────
# 11. GRÁFICOS
# ─────────────────────────────────────────────────────────────────────────────
cat("\nGenerando graficos...\n")
tema <- theme_minimal(base_size=11) +
  theme(plot.title=element_text(face="bold",hjust=.5),
        plot.subtitle=element_text(hjust=.5,color="grey40"),
        panel.grid.minor=element_blank(), axis.title=element_text(size=10))

# G1: Serie anual
g1 <- ggplot(tot_an, aes(anio, total_mm)) +
  geom_col(fill="#2196F3", alpha=.75) +
  geom_smooth(method="lm", se=TRUE, color="#F44336", linewidth=.8, formula=y~x) +
  scale_x_continuous(breaks=seq(1970,2030,5)) +
  labs(title=paste("Precipitación Anual –",ESTACION),
       subtitle="Totales acumulados con tendencia lineal",
       x="Año", y="mm/año") + tema

# G2: Ciclo mensual
g2 <- ggplot(ciclo_m, aes(factor(mes,labels=month.abb), media_mm_mes)) +
  geom_col(fill="#4CAF50", alpha=.8) +
  labs(title="Ciclo Anual de Precipitación",
       subtitle="Promedio mensual multianual (mm/mes)",
       x="Mes", y="mm/mes") + tema

# G3: Ciclo diurno
g3 <- ggplot(ciclo_h, aes(hora)) +
  geom_col(aes(y=frac_pct, fill="Fraccion lluvia (%)"), alpha=.4) +
  geom_line(aes(y=intens*5, color="Intensidad media (×5)"), linewidth=1) +
  scale_x_continuous(breaks=0:23) +
  scale_fill_manual(values=c("Fraccion lluvia (%)"="#3F51B5")) +
  scale_color_manual(values=c("Intensidad media (×5)"="#E91E63")) +
  labs(title="Ciclo Diurno de Precipitación", x="Hora", y="%  |  mm/h×5",
       fill=NULL, color=NULL) + tema + theme(legend.position="bottom")

# G4: Disponibilidad
g4 <- ggplot(tot_an, aes(anio, disp_pct)) +
  geom_col(aes(fill=disp_pct<(100*UMBRAL_ANUAL)), show.legend=FALSE) +
  scale_fill_manual(values=c("FALSE"="#66BB6A","TRUE"="#EF5350")) +
  geom_hline(yintercept=100*UMBRAL_ANUAL, linetype="dashed") +
  annotate("text", x=min(tot_an$anio)+1,
           y=100*UMBRAL_ANUAL+2, label=sprintf("%.0f%% OMM C1",100*UMBRAL_ANUAL),
           size=3, hjust=0) +
  scale_x_continuous(breaks=seq(1970,2030,5)) +
  scale_y_continuous(labels=function(x) paste0(x,"%"), limits=c(0,105)) +
  labs(title="Disponibilidad de Datos por Año", x="Año", y="%") + tema

# G5: Distribución log-log
p_pos <- prec_orig[prec_orig > 0 & !is.na(prec_orig)]
g5 <- ggplot(data.frame(p=p_pos), aes(p)) +
  geom_histogram(bins=50, fill="#FF9800", color="white", alpha=.8) +
  scale_x_log10(labels=label_number()) +
  scale_y_log10(labels=label_number()) +
  labs(title="Distribución de Intensidades (horas con lluvia)",
       subtitle=sprintf("Escala log-log | umbral: %.1f mm/h",UMBRAL_LLUVIA),
       x="Intensidad (mm/h)", y="Frecuencia") + tema

# G6: Heatmap año×mes
hm <- df_an %>% group_by(anio,mes) %>%
  summarise(tot=sum(precip_relleno,na.rm=TRUE), .groups="drop")
g6 <- ggplot(hm, aes(mes,anio,fill=tot)) +
  geom_tile(color="white", linewidth=.2) +
  scale_fill_gradientn(colours=c("#E3F2FD","#64B5F6","#1976D2","#0D47A1"),
                       name="mm/mes") +
  scale_x_continuous(breaks=1:12, labels=month.abb) +
  labs(title="Precipitación Mensual", x="Mes", y="Año") + tema

# G7: Depuración OMM
g7 <- ggplot(tabla_dep %>%
  mutate(Col=case_when(Decision=="VALIDO_C1"~"Valido C1",
                       Decision=="VALIDO_C2"~"Valido C2", TRUE~"Rechazado")),
  aes(Anio, Disp_anual_pct, fill=Col)) +
  geom_col(width=.8) +
  geom_hline(yintercept=100*UMBRAL_ANUAL, linetype="dashed", color="#1565C0") +
  scale_fill_manual(values=c("Valido C1"="#2E7D32","Valido C2"="#F9A825","Rechazado"="#C62828")) +
  scale_y_continuous(limits=c(0,105), labels=function(x) paste0(x,"%")) +
  labs(title="Depuración Hidrológica – OMM N°168",
       subtitle=sprintf("C1>=%.0f%% | C2>=%.0f%% hum | C3 %s",
         100*UMBRAL_ANUAL, 100*UMBRAL_HUM,
         ifelse(APLICAR_C3,"activo","desactivado (gaps estructurales)")),
       x="Año", y="Disponibilidad (%)", fill="Decisión") +
  tema + theme(legend.position="bottom")

# G8: Cv vs IETD
g8 <- ggplot(tabla_ietd, aes(IETD_h, Cv)) +
  geom_line(color="#1976D2", linewidth=1) +
  geom_point(aes(color=Dif_Cv_1==min(Dif_Cv_1,na.rm=TRUE)), size=3, show.legend=FALSE) +
  scale_color_manual(values=c("FALSE"="#1976D2","TRUE"="#E53935")) +
  geom_hline(yintercept=1, linetype="dashed", color="#E53935") +
  geom_vline(xintercept=IETD_OPT, linetype="dotted", color="#E53935") +
  annotate("text", x=IETD_OPT+.5, y=max(tabla_ietd$Cv,na.rm=TRUE)*.95,
           label=sprintf("IETD=%dh\nCv=%.3f",IETD_OPT,CV_OPT), hjust=0, size=3.5) +
  scale_x_continuous(breaks=IETD_CANDIDATOS) +
  labs(title="Criterio Cv = 1 para Selección de IETD",
       subtitle="Restrepo-Posada & Eagleson (1982) | Punto rojo = IETD óptimo",
       x="IETD candidato (horas)", y="Coeficiente de Variación (Cv)") + tema

# G9: Distribución de profundidades por tormenta
g9 <- ggplot(df_torm, aes(profundidad_mm)) +
  geom_histogram(bins=40, fill="#7B1FA2", color="white", alpha=.8) +
  scale_x_log10(labels=label_number()) +
  labs(title=sprintf("Distribución de Profundidades por Tormenta (IETD=%dh)",IETD_OPT),
       subtitle=sprintf("N=%d tormentas | P_min=%.1f mm",nrow(df_torm),IETD_MIN_PROFUNDIDAD),
       x="Profundidad (mm)", y="Frecuencia") + tema

# G10: Tiempos entre tormentas
g10 <- ggplot(data.frame(tbt=tbt_opt), aes(tbt)) +
  geom_histogram(bins=40, fill="#00796B", color="white", alpha=.8) +
  geom_vline(xintercept=mean(tbt_opt), linetype="dashed", color="#E53935") +
  annotate("text", x=mean(tbt_opt)+2, y=Inf, vjust=2,
           label=sprintf("Media=%.1fh\nCv=%.3f",mean(tbt_opt),CV_OPT), size=3.5) +
  labs(title=sprintf("Distribución de Tiempos Entre Tormentas (IETD=%dh)",IETD_OPT),
       subtitle="Línea roja = media | Cv~1 valida independencia de eventos",
       x="Tiempo entre tormentas (horas)", y="Frecuencia") + tema

# Guardar gráficos
ruta_pdf <- paste0(RUTA_SALIDA,"Graficos_Precipitacion_",ESTACION,".pdf")
pdf(ruta_pdf, width=12, height=7, onefile=TRUE)
for (g in list(g1,g2,g3,g4,g5,g6,g7,g8,g9,g10)) suppressWarnings(print(g))
dev.off()
for (nm in c("G1_Serie_Anual","G2_Ciclo_Mensual","G3_Ciclo_Diurno","G4_Disponibilidad",
             "G5_Distribucion","G6_Heatmap","G7_Depuracion","G8_IETD_Cv",
             "G9_Profundidades","G10_TBT")) {
  idx_g <- which(c("G1_Serie_Anual","G2_Ciclo_Mensual","G3_Ciclo_Diurno","G4_Disponibilidad",
                   "G5_Distribucion","G6_Heatmap","G7_Depuracion","G8_IETD_Cv",
                   "G9_Profundidades","G10_TBT")==nm)
  suppressWarnings(ggsave(paste0(RUTA_SALIDA,nm,".png"),
    plot=list(g1,g2,g3,g4,g5,g6,g7,g8,g9,g10)[[idx_g]],
    width=12, height=6, dpi=150))
}
cat(sprintf("OK %s\n", ruta_pdf))

# ─────────────────────────────────────────────────────────────────────────────
# 12. EXPORTACIÓN EXCEL GENERAL
# ─────────────────────────────────────────────────────────────────────────────
eh <- createStyle(fgFill="#1976D2",fontColour="#FFFFFF",textDecoration="BOLD",
                  fontName="Arial",fontSize=10,halign="CENTER")
ew <- createStyle(fgFill="#FFCDD2",fontName="Arial",fontSize=9,halign="CENTER")
eg <- createStyle(fgFill="#C8E6C9",fontName="Arial",fontSize=9,halign="CENTER")
ey <- createStyle(fgFill="#FFF9C4",fontName="Arial",fontSize=9,halign="CENTER")

add_sheet <- function(wb, nm, df, hdr_style=eh, ...) {
  addWorksheet(wb, nm)
  writeData(wb, nm, df, ...)
  nr <- if (hasArg(startRow)) list(...)$startRow else 1
  addStyle(wb, nm, hdr_style, rows=nr, cols=1:ncol(df), gridExpand=TRUE)
  setColWidths(wb, nm, cols=1:ncol(df), widths="auto")
}

wb <- createWorkbook()

# Metadatos
meta_df <- data.frame(
  Campo=c("Estacion","Codigo","Municipio","Departamento","Fuente",
          "Coord_N","Coord_E","Altura_msnm","Periodo_inicio","Periodo_fin",
          "N_horas_total","N_horas_validas","Disp_total_pct",
          "UMBRAL_ANUAL_C1","UMBRAL_HUM_C2","C3_activo",
          "IETD_optimo_h","Cv_IETD","N_tormentas","Fecha_proceso"),
  Valor=c(ESTACION,CODIGO,MUNICIPIO,DEPARTAMENTO,FUENTE,
          as.character(COORD_N),as.character(COORD_E),as.character(ALTURA_msnm),
          format(fecha_min,"%Y-%m-%d"),format(fecha_max,"%Y-%m-%d"),
          nrow(df_serie),sum(!is.na(df_serie$precip_qc)),
          round(100*mean(!is.na(df_serie$precip_qc)),1),
          sprintf("%.0f%%",100*UMBRAL_ANUAL),sprintf("%.0f%%",100*UMBRAL_HUM),
          ifelse(APLICAR_C3,"SI","NO"),
          as.character(IETD_OPT),as.character(round(CV_OPT,4)),
          as.character(nrow(df_torm)),format(Sys.Date(),"%Y-%m-%d")))
add_sheet(wb,"Metadatos",meta_df)

add_sheet(wb,"QC_Calidad",resumen_qc)
add_sheet(wb,"Gaps_Serie",gaps_info %>% mutate(inicio=format(inicio),fin=format(fin)))
add_sheet(wb,"Estadisticas",stats)
add_sheet(wb,"Totales_Anuales",tot_an)
add_sheet(wb,"Ciclo_Mensual",ciclo_m)
add_sheet(wb,"Ciclo_Diurno",ciclo_h)
add_sheet(wb,"Validacion_Relleno",metricas)

# Auditoria depuración
add_sheet(wb,"Auditoria_OMM",tabla_dep)
for (r in seq_len(nrow(tabla_dep))) {
  est <- if (tabla_dep$Decision[r]=="VALIDO_C1") eg
         else if (tabla_dep$Decision[r]=="VALIDO_C2") ey else ew
  addStyle(wb,"Auditoria_OMM",est,rows=r+1,cols=1:ncol(tabla_dep),gridExpand=TRUE)
}

ruta_xl <- paste0(RUTA_SALIDA,"Resultados_",ESTACION,".xlsx")
saveWorkbook(wb, ruta_xl, overwrite=TRUE)
cat(sprintf("OK %s\n", ruta_xl))

# ─────────────────────────────────────────────────────────────────────────────
# 13. EXPORTACIÓN EXCEL IETD (archivo específico para análisis posterior)
# ─────────────────────────────────────────────────────────────────────────────
cat("\nGenerando archivo IETD...\n")

wb_ietd <- createWorkbook()

# --- Hoja 1: Metadatos y criterio IETD --------------------------------------
meta_ietd <- data.frame(
  Parametro=c("Estacion","Codigo","Municipio","Periodo_inicio","Periodo_fin",
              "Anos_validos","N_anos","IETD_optimo_horas","Cv_obtenido",
              "IETD_max_fisico_horas","Criterio_seleccion","Referencia",
              "IETD_min_lluvia_mm_h","Profundidad_min_tormenta_mm",
              "N_tormentas_identificadas","Tasa_tormentas_anio",
              "TBT_medio_horas","TBT_medio_dias","Fecha_proceso"),
  Valor=c(ESTACION,CODIGO,MUNICIPIO,
          format(min(df_torm$inicio),"%Y-%m-%d"),
          format(max(df_torm$fin),"%Y-%m-%d"),
          paste(anios_ok,collapse=", "),
          length(anios_ok),
          IETD_OPT, round(CV_OPT,4),
          sprintf("%dh (cap fisico conveccion tropical, Poveda et al. 2002)",
                  IETD_MAX_FISICO),
          "Cv = 1 con restriccion fisica IETD <= IETD_MAX_FISICO",
          "Restrepo-Posada & Eagleson (1982), WMO-168, Poveda et al. (2002)",
          IETD_MIN_LLUVIA, IETD_MIN_PROFUNDIDAD,
          nrow(df_torm),
          round(nrow(df_torm)/length(anios_ok),2),
          round(mean(tbt_opt,na.rm=TRUE),2),
          round(mean(tbt_opt,na.rm=TRUE)/24,2),
          format(Sys.Date(),"%Y-%m-%d")))

addWorksheet(wb_ietd,"Metadatos_IETD")
writeData(wb_ietd,"Metadatos_IETD",meta_ietd)
addStyle(wb_ietd,"Metadatos_IETD",
  createStyle(fgFill="#0D47A1",fontColour="#FFFFFF",textDecoration="BOLD",
              fontName="Arial",fontSize=11,halign="CENTER"),
  rows=1,cols=1:2,gridExpand=TRUE)
setColWidths(wb_ietd,"Metadatos_IETD",cols=1:2,widths=c(38,50))

# --- Hoja 2: Evaluación Cv por IETD candidato --------------------------------
addWorksheet(wb_ietd,"Evaluacion_Cv_IETD")
writeData(wb_ietd,"Evaluacion_Cv_IETD",tabla_ietd)
addStyle(wb_ietd,"Evaluacion_Cv_IETD",eh,rows=1,cols=1:ncol(tabla_ietd),gridExpand=TRUE)
# Resaltar fila óptima
addStyle(wb_ietd,"Evaluacion_Cv_IETD",
  createStyle(fgFill="#A5D6A7",fontName="Arial",textDecoration="BOLD",halign="CENTER"),
  rows=idx_opt+1, cols=1:ncol(tabla_ietd), gridExpand=TRUE)
setColWidths(wb_ietd,"Evaluacion_Cv_IETD",cols=1:ncol(tabla_ietd),widths=18)

# --- Hoja 3: Catálogo de tormentas independientes ----------------------------
df_torm_exp <- df_torm %>%
  mutate(inicio=format(inicio,"%Y-%m-%d %H:%M"),
         fin=format(fin,"%Y-%m-%d %H:%M")) %>%
  rename(Evento_ID=evento_id, Inicio=inicio, Fin=fin,
         Duracion_h=duracion_h, Profundidad_mm=profundidad_mm,
         Intensidad_max_mm_h=intensidad_max,
         Intensidad_media_mm_h=intensidad_media,
         TBT_h_desde_anterior=TBT_horas,
         Anio=anio, Mes=mes)

addWorksheet(wb_ietd,"Catalogo_Tormentas")
writeData(wb_ietd,"Catalogo_Tormentas",df_torm_exp)
addStyle(wb_ietd,"Catalogo_Tormentas",eh,rows=1,cols=1:ncol(df_torm_exp),gridExpand=TRUE)
setColWidths(wb_ietd,"Catalogo_Tormentas",cols=1:ncol(df_torm_exp),widths="auto")

# --- Hoja 4: Tiempos entre tormentas (TBT) para ajuste distribucional --------
df_tbt <- data.frame(
  N_par       = seq_along(tbt_opt),
  Tormenta_i  = df_torm_exp$Evento_ID[-1],
  Tormenta_j  = df_torm_exp$Evento_ID[-nrow(df_torm_exp)],
  TBT_horas   = round(tbt_opt, 3),
  TBT_dias    = round(tbt_opt/24, 4),
  ln_TBT      = round(log(tbt_opt), 4)   # útil para ajuste exponencial/Weibull
)
addWorksheet(wb_ietd,"TBT_Tiempos_Entre_Tormentas")
writeData(wb_ietd,"TBT_Tiempos_Entre_Tormentas",df_tbt)
addStyle(wb_ietd,"TBT_Tiempos_Entre_Tormentas",eh,rows=1,cols=1:ncol(df_tbt),gridExpand=TRUE)
setColWidths(wb_ietd,"TBT_Tiempos_Entre_Tormentas",cols=1:ncol(df_tbt),widths=18)

# --- Hoja 5: Estadísticas por año (tormentas/año para verificar estacionariedad)
est_anual_ietd <- df_torm %>%
  group_by(anio) %>%
  summarise(N_tormentas=n(),
            Precip_total_mm=round(sum(profundidad_mm),1),
            Duracion_media_h=round(mean(duracion_h),1),
            Profund_media_mm=round(mean(profundidad_mm),1),
            Intens_max_mm_h=round(max(intensidad_max),1),
            TBT_medio_h=round(mean(TBT_horas,na.rm=TRUE),1),
            .groups="drop") %>%
  left_join(tabla_dep %>% select(Anio,Disp_anual_pct,Decision),
            by=c("anio"="Anio"))

addWorksheet(wb_ietd,"Estadisticas_Anuales_IETD")
writeData(wb_ietd,"Estadisticas_Anuales_IETD",est_anual_ietd)
addStyle(wb_ietd,"Estadisticas_Anuales_IETD",eh,rows=1,cols=1:ncol(est_anual_ietd),gridExpand=TRUE)
setColWidths(wb_ietd,"Estadisticas_Anuales_IETD",cols=1:ncol(est_anual_ietd),widths=20)

# --- Hoja 6: Serie horaria depurada con etiquetas de evento ------------------
df_sh_exp <- df_serie_ietd %>%
  mutate(fecha_hora=format(fecha_hora,"%Y-%m-%d %H:%M")) %>%
  rename(Fecha_Hora=fecha_hora, Precip_mm_h=precip_mm_h,
         Metodo=metodo_relleno, Anio=anio, Mes=mes, Dia=dia, Hora=hora,
         Evento_ID=evento_id, En_tormenta=en_tormenta)

addWorksheet(wb_ietd,"Serie_Horaria_Depurada")
writeData(wb_ietd,"Serie_Horaria_Depurada",df_sh_exp)
addStyle(wb_ietd,"Serie_Horaria_Depurada",eh,rows=1,cols=1:ncol(df_sh_exp),gridExpand=TRUE)
setColWidths(wb_ietd,"Serie_Horaria_Depurada",cols=1:ncol(df_sh_exp),widths="auto")

# --- Hoja 7: Máximos anuales por duración (insumo IDF) -----------------------
duraciones_idf <- c(1,2,3,4,6,8,10,12,16,20,24)
max_anual_dur <- lapply(anios_ok, function(a) {
  p_a <- df_dep$precip_relleno[year(df_dep$fecha_hora)==a]
  row <- c(Anio=a)
  for (d in duraciones_idf) {
    if (length(p_a) >= d)
      row[paste0("P",d,"h_mm")] <- round(max(zoo::rollsum(p_a,d,fill=NA,align="right"),na.rm=TRUE),1)
    else
      row[paste0("P",d,"h_mm")] <- NA_real_
  }
  as.data.frame(t(row))
}) %>% do.call(rbind, .) %>%
  mutate(across(everything(), as.numeric))

addWorksheet(wb_ietd,"Maximos_Anuales_Duracion")
writeData(wb_ietd,"Maximos_Anuales_Duracion",max_anual_dur)
addStyle(wb_ietd,"Maximos_Anuales_Duracion",eh,rows=1,cols=1:ncol(max_anual_dur),gridExpand=TRUE)
setColWidths(wb_ietd,"Maximos_Anuales_Duracion",cols=1:ncol(max_anual_dur),widths=14)

ruta_ietd <- paste0(RUTA_SALIDA,"IETD_",ESTACION,".xlsx")
saveWorkbook(wb_ietd, ruta_ietd, overwrite=TRUE)
cat(sprintf("OK %s\n", ruta_ietd))

# ─────────────────────────────────────────────────────────────────────────────
# 14. REPORTE FINAL EN CONSOLA
# ─────────────────────────────────────────────────────────────────────────────
# Precipitación media anual sobre años válidos (depurada)
prom_dep <- mean(tot_an$total_mm[tot_an$anio %in% anios_ok], na.rm=TRUE)

cat("\n", strrep("=",70), "\n")
cat(sprintf("  REPORTE FINAL – %s | %s\n", ESTACION, FUENTE))
cat(strrep("-",70), "\n")
cat(sprintf("  Período serie original:  %s a %s\n",
    format(fecha_min,"%Y-%m-%d"), format(fecha_max,"%Y-%m-%d")))
cat(sprintf("  Serie depurada (OMM168): %d años [%s]\n",
    length(anios_ok), paste(anios_ok, collapse=", ")))
cat(sprintf("  Disponibilidad original: %.1f%%\n",
    100*mean(!is.na(df_serie$precip_qc))))
cat(sprintf("  Promedio anual (dep.):   %.0f mm/año\n", prom_dep))
cat(sprintf("  Máximo horario:          %.1f mm/h\n", max(prec_orig,na.rm=TRUE)))
cat(sprintf("  Mes más lluvioso:        %s\n",
    month.name[ciclo_m$mes[which.max(ciclo_m$media_mm_mes)]]))
cat(sprintf("  Hora de mayor actividad: %02d:00\n",
    ciclo_h$hora[which.max(ciclo_h$frac_pct)]))
cat(strrep("-",70), "\n")
cat(sprintf("  IETD óptimo:             %d horas (Cv = %.4f)\n", IETD_OPT, CV_OPT))
cat(sprintf("  N tormentas:             %d (%.1f/año)\n",
    nrow(df_torm), nrow(df_torm)/length(anios_ok)))
cat(sprintf("  TBT medio:               %.1f h (%.1f días)\n",
    mean(tbt_opt), mean(tbt_opt)/24))
cat(sprintf("  Profundidad media:       %.1f mm | Duración media: %.1f h\n",
    mean(df_torm$profundidad_mm), mean(df_torm$duracion_h)))
cat(strrep("-",70), "\n")
cat("  ARCHIVOS GENERADOS:\n")
cat(sprintf("  Resultados generales:  %s\n", ruta_xl))
cat(sprintf("  Archivo IETD:          %s\n", ruta_ietd))
cat(sprintf("  PDF graficos:          %s\n", ruta_pdf))
cat(sprintf("  PNGs:                  %sG*.png\n", RUTA_SALIDA))
cat(strrep("=",70), "\n")
cat("OK Analisis completado.\n")

# =============================================================================
# 15. ANÁLISIS DE FRECUENCIA DE DURACIONES DE TORMENTAS INDEPENDIENTES
# Objetivo: duración de tormenta de diseño para una probabilidad dada
# Enfoque: población completa de tormentas independientes (Restrepo-Posada &
#          Eagleson 1982) → análisis de frecuencia → distribución ajustada
# Fundamento: las duraciones de tormentas convectivas siguen distribuciones
#   sesgadas (Exponencial, Log-Normal, Gamma). Se ajustan varias y se
#   selecciona la de mejor bondad de ajuste (KS, AD).
# Referencia: Restrepo-Posada & Eagleson (1982); Chow et al. (1988) Cap.12
# =============================================================================
cat("\n", strrep("=",70), "\n")
cat("  15. FRECUENCIA DE DURACIONES DE TORMENTAS INDEPENDIENTES\n")
cat("  Restrepo-Posada & Eagleson (1982)\n")
cat(strrep("=",70), "\n")

# ── 15.0 Verificación de la muestra ─────────────────────────────────────────
x <- df_torm$duracion_h   # población completa de tormentas independientes
n <- length(x)

cat(sprintf("\nPoblacion: %d tormentas independientes (IETD = %dh)\n", n, IETD_OPT))
cat(sprintf("Min: %.1f h | Media: %.2f h | Mediana: %.1f h\n", min(x), mean(x), median(x)))
cat(sprintf("Max: %.1f h | SD: %.2f h | CV: %.3f\n", max(x), sd(x), sd(x)/mean(x)))

# Coeficiente de asimetría muestral
Cs <- (n * sum((x - mean(x))^3)) / ((n-1)*(n-2)*sd(x)^3)
cat(sprintf("Cs (asimetria): %.3f\n", Cs))
cat(sprintf("Percentiles: P10=%.1fh P25=%.1fh P50=%.1fh P75=%.1fh P90=%.1fh P95=%.1fh\n",
    quantile(x,.10), quantile(x,.25), quantile(x,.50),
    quantile(x,.75), quantile(x,.90), quantile(x,.95)))

# Nota sobre el rango esperado
cat(sprintf("\nNOTA: En zonas convectivas (Valle del Cauca) las tormentas tipicas\n"))
cat(sprintf("      duran 1-6h. Valores mayores corresponden a sistemas de meso-\n"))
cat(sprintf("      escala o frentes. Todos se incluyen en el ajuste.\n\n"))

# ── 15.1 Probabilidades empíricas (Weibull) ──────────────────────────────────
# Fórmula de Weibull: P(X <= x) = m/(n+1), recomendada OMM por ser insesgada
df_emp <- data.frame(duracion_h = sort(x)) %>%
  mutate(
    rango   = seq_len(n),
    P_emp   = rango / (n + 1),          # prob. no excedencia
    Tr_emp  = 1 / (1 - P_emp),          # periodo de retorno
    P_exc   = 1 - P_emp                 # prob. excedencia
  )

# ── 15.2 Ajuste de distribuciones ────────────────────────────────────────────
# Distribuciones apropiadas para duraciones (variable positiva, sesgada):
#   Exponencial  – caso especial Gamma(1), implica proceso Poisson
#   Gamma        – flexible para variables positivas sesgadas
#   Log-Normal   – muy usada en hidrología para duraciones
#   Weibull      – común en análisis de confiabilidad y duraciones

# --- Exponencial (1 parámetro: lambda = 1/media) ----------------------------
lambda_exp <- 1 / mean(x)
q_exp_fn   <- function(p) qexp(p, rate = lambda_exp)
p_exp_fn   <- function(q) pexp(q, rate = lambda_exp)
ks_exp     <- suppressWarnings(ks.test(x, p_exp_fn))

# --- Gamma (2 parámetros: forma k, escala theta) por método de momentos ------
k_gam    <- (mean(x) / sd(x))^2
th_gam   <- sd(x)^2 / mean(x)
q_gam_fn <- function(p) qgamma(p, shape = k_gam, scale = th_gam)
p_gam_fn <- function(q) pgamma(q, shape = k_gam, scale = th_gam)
ks_gam   <- suppressWarnings(ks.test(x, p_gam_fn))

# --- Log-Normal (2 parámetros: mu_ln, sigma_ln) ------------------------------
lx       <- log(x)
mu_ln    <- mean(lx)
sd_ln    <- sd(lx)
q_ln_fn  <- function(p) qlnorm(p, meanlog = mu_ln, sdlog = sd_ln)
p_ln_fn  <- function(q) plnorm(q, meanlog = mu_ln, sdlog = sd_ln)
ks_ln    <- suppressWarnings(ks.test(x, p_ln_fn))

# --- Weibull (2 parámetros: forma c, escala u) – MLE iterativo ---------------
# MLE para Weibull: resolver iterativamente (Newton-Raphson sobre log-likelihood)
weibull_mle <- function(x) {
  # Estimación inicial: método de momentos
  cv   <- sd(x)/mean(x)
  c0   <- (cv)^(-1.086)   # aproximación empírica de Justus (1978)
  for (iter in 1:100) {
    sc <- sum(x^c0 * log(x)) / sum(x^c0)
    g  <- 1/c0 + mean(log(x)) - sc
    dg <- -1/c0^2 - (sum(x^c0*(log(x))^2)*sum(x^c0) - (sum(x^c0*log(x)))^2) /
               sum(x^c0)^2
    c1 <- c0 - g/dg
    if (abs(c1 - c0) < 1e-8) break
    c0 <- max(c1, 0.01)
  }
  u0 <- (mean(x^c0))^(1/c0)
  list(shape=c0, scale=u0)
}
wb_par   <- weibull_mle(x)
c_wb     <- wb_par$shape; u_wb <- wb_par$scale
q_wb_fn  <- function(p) qweibull(p, shape=c_wb, scale=u_wb)
p_wb_fn  <- function(q) pweibull(q, shape=c_wb, scale=u_wb)
ks_wb    <- suppressWarnings(ks.test(x, p_wb_fn))

# ── 15.3 Tabla de bondad de ajuste ───────────────────────────────────────────
tabla_ks <- data.frame(
  Distribucion   = c("Exponencial","Gamma","Log-Normal","Weibull"),
  Parametros     = c(
    sprintf("lambda=%.4f (media=%.2fh)", lambda_exp, 1/lambda_exp),
    sprintf("k=%.3f, theta=%.3f", k_gam, th_gam),
    sprintf("mu_ln=%.3f, sd_ln=%.3f", mu_ln, sd_ln),
    sprintf("c=%.3f, u=%.3f", c_wb, u_wb)),
  KS_D           = round(c(ks_exp$statistic, ks_gam$statistic,
                            ks_ln$statistic,  ks_wb$statistic), 5),
  p_valor        = round(c(ks_exp$p.value,   ks_gam$p.value,
                            ks_ln$p.value,    ks_wb$p.value),  4),
  Resultado      = ifelse(c(ks_exp$p.value, ks_gam$p.value,
                             ks_ln$p.value,  ks_wb$p.value) >= 0.05,
                          "ACEPTADA", "RECHAZADA")
)
cat("Bondad de ajuste (Kolmogorov-Smirnov, alpha=5%):\n")
print(tabla_ks)

# Seleccionar la de mayor p-valor entre las aceptadas; si ninguna, la menor D
aceptadas <- tabla_ks[tabla_ks$Resultado == "ACEPTADA", ]
if (nrow(aceptadas) > 0) {
  dist_rec <- aceptadas$Distribucion[which.max(aceptadas$p_valor)]
} else {
  dist_rec <- tabla_ks$Distribucion[which.min(tabla_ks$KS_D)]
  cat("AVISO: ninguna distribucion aceptada al 5% (frecuente con n>200).\n")
  cat("  Con muestras grandes el KS detecta diferencias minimas.\n")
  cat("  Se usa la distribucion de menor estadistico D (mejor ajuste relativo).\n")
}
cat(sprintf("\nDistribucion recomendada: %s\n", dist_rec))

q_rec_fn <- switch(dist_rec,
  "Exponencial" = q_exp_fn,
  "Gamma"       = q_gam_fn,
  "Log-Normal"  = q_ln_fn,
  "Weibull"     = q_wb_fn)

# ── 15.4 Cuantiles de diseño ─────────────────────────────────────────────────
# Probabilidades de no excedencia de interés para diseño
P_diseño  <- c(0.50, 0.75, 0.80, 0.90, 0.95, 0.98, 0.99, 0.995)
Tr_diseño <- round(1 / (1 - P_diseño), 0)

# Bootstrap IC 95% (B=1000 remuestras)
set.seed(42); B <- 1000
boot_q <- matrix(NA_real_, nrow=B, ncol=length(P_diseño))
for (b in seq_len(B)) {
  xb <- sample(x, n, replace=TRUE)
  qfn <- switch(dist_rec,
    "Exponencial" = { lb <- 1/mean(xb); function(p) qexp(p, rate=lb) },
    "Gamma"       = { kb <- (mean(xb)/sd(xb))^2; tb <- sd(xb)^2/mean(xb)
                      function(p) qgamma(p, shape=kb, scale=tb) },
    "Log-Normal"  = { mlb <- mean(log(xb)); slb <- sd(log(xb))
                      function(p) qlnorm(p, meanlog=mlb, sdlog=slb) },
    "Weibull"     = { pb <- weibull_mle(xb)
                      function(p) qweibull(p, shape=pb$shape, scale=pb$scale) }
  )
  boot_q[b,] <- sapply(P_diseño, qfn)
}
ic_low <- apply(boot_q, 2, quantile, 0.025, na.rm=TRUE)
ic_hig <- apply(boot_q, 2, quantile, 0.975, na.rm=TRUE)
q_pts  <- sapply(P_diseño, q_rec_fn)

tabla_diseño <- data.frame(
  P_no_excedencia   = P_diseño,
  Tr_años           = Tr_diseño,
  Duracion_diseño_h = round(q_pts,  2),
  IC95_inf_h        = round(ic_low, 2),
  IC95_sup_h        = round(ic_hig, 2),
  Distribucion      = dist_rec
)
cat("\nDuraciones de diseño:\n"); print(tabla_diseño)

# ── 15.5 Gráficos ────────────────────────────────────────────────────────────

# Rango de probabilidades para las curvas teóricas
p_curva  <- seq(0.01, 0.995, length.out=400)
x_curva  <- seq(min(x)*0.5, max(x)*1.15, length.out=400)

# G11: Papel de probabilidad (Weibull empírico vs distribuciones ajustadas)
df_teoricas <- bind_rows(
  data.frame(P=p_curva, dur=q_exp_fn(p_curva), Dist="Exponencial"),
  data.frame(P=p_curva, dur=q_gam_fn(p_curva), Dist="Gamma"),
  data.frame(P=p_curva, dur=q_ln_fn(p_curva),  Dist="Log-Normal"),
  data.frame(P=p_curva, dur=q_wb_fn(p_curva),  Dist="Weibull")
) %>% filter(is.finite(dur), dur > 0)

breaks_Tr <- c(1.05, 1.25, 2, 5, 10, 25, 50, 100)
breaks_P  <- 1 - 1/breaks_Tr

g11 <- ggplot() +
  geom_line(data=df_teoricas,
            aes(P, dur, color=Dist, linetype=Dist), linewidth=0.9) +
  geom_point(data=df_emp,
             aes(P_emp, duracion_h, shape="Datos (Weibull)"),
             color="black", size=2.5, alpha=0.7) +
  geom_point(data=tabla_diseño,
             aes(P_no_excedencia, Duracion_diseño_h),
             color="#E53935", size=4, shape=18,
             show.legend=FALSE) +
  geom_errorbar(data=tabla_diseño,
                aes(x=P_no_excedencia, ymin=IC95_inf_h, ymax=IC95_sup_h),
                color="#E53935", width=0.01, linewidth=0.7) +
  scale_x_continuous(
    breaks = breaks_P,
    labels = paste0("Tr=", breaks_Tr, "a\n(P=", round(breaks_P,2), ")"),
    limits = c(0, 0.998)
  ) +
  scale_color_manual(values=c(
    "Exponencial"="#1976D2","Gamma"="#388E3C",
    "Log-Normal"="#F57C00","Weibull"="#7B1FA2")) +
  scale_linetype_manual(values=c(
    "Exponencial"="solid","Gamma"="dashed",
    "Log-Normal"="dotdash","Weibull"="dotted")) +
  scale_shape_manual(values=c("Datos (Weibull)"=16)) +
  annotate("text", x=0.02, y=max(df_torm$duracion_h)*0.9,
           label=sprintf("Recomendada: %s\nn=%d tormentas",dist_rec,n),
           hjust=0, size=3.5, color="#E53935") +
  labs(
    title    = "Análisis de Frecuencia – Duración de Tormentas Independientes",
    subtitle = sprintf("Estación %s | IETD=%dh | n=%d tormentas | Cv=%.3f | Diamantes rojos: diseño",
                       ESTACION, IETD_OPT, n, sd(x)/mean(x)),
    x        = "Probabilidad de no excedencia (Período de retorno)",
    y        = "Duración de tormenta (horas)",
    color="Distribución", linetype="Distribución", shape=NULL
  ) + tema + theme(legend.position="bottom")

# G12: Curva duración–Tr con IC 95% (distribución recomendada)
p_curva2  <- seq(0.01, 0.995, length.out=300)
Tr_curva2 <- 1/(1-p_curva2)
df_curva2 <- data.frame(
  Tr  = Tr_curva2,
  dur = sapply(p_curva2, q_rec_fn)
) %>% filter(is.finite(dur), dur > 0)

g12 <- ggplot() +
  geom_ribbon(data=data.frame(Tr=Tr_diseño, lo=ic_low, hi=ic_hig),
              aes(Tr, ymin=lo, ymax=hi), fill="#1976D2", alpha=0.15) +
  geom_line(data=df_curva2, aes(Tr, dur), color="#1976D2", linewidth=1.2) +
  geom_point(data=tabla_diseño,
             aes(Tr_años, Duracion_diseño_h), color="#1976D2", size=3.5) +
  geom_point(data=df_emp,
             aes(Tr_emp, duracion_h), shape=1, color="grey40",
             size=2, alpha=0.6) +
  geom_text(data=tabla_diseño,
            aes(Tr_años, Duracion_diseño_h,
                label=sprintf("%.1fh", Duracion_diseño_h)),
            vjust=-0.9, size=3.2, color="#1976D2") +
  scale_x_log10(breaks=c(1,2,5,10,25,50,100,200),
                labels=c("1","2","5","10","25","50","100","200")) +
  labs(
    title    = sprintf("Curva Duración–Período de Retorno (%s)", dist_rec),
    subtitle = sprintf("IC 95%% bootstrap (B=%d) | Círculos: datos empíricos | Estación %s",
                       B, ESTACION),
    x = "Período de retorno (años) – escala log",
    y = "Duración de tormenta de diseño (horas)"
  ) + tema

# G13: Histograma con densidades ajustadas
df_dens <- bind_rows(
  data.frame(x=x_curva, d=dexp(x_curva, rate=lambda_exp),  Dist="Exponencial"),
  data.frame(x=x_curva, d=dgamma(x_curva, shape=k_gam, scale=th_gam), Dist="Gamma"),
  data.frame(x=x_curva, d=dlnorm(x_curva, meanlog=mu_ln, sdlog=sd_ln),Dist="Log-Normal"),
  data.frame(x=x_curva, d=dweibull(x_curva, shape=c_wb, scale=u_wb),  Dist="Weibull")
) %>% filter(is.finite(d), d >= 0)

g13 <- ggplot(data.frame(x=x), aes(x)) +
  geom_histogram(aes(y=after_stat(density)),
                 bins=max(8, round(sqrt(n))),
                 fill="#B0BEC5", color="white", alpha=0.85) +
  geom_line(data=df_dens,
            aes(x, d, color=Dist, linetype=Dist), linewidth=0.9) +
  geom_rug(aes(x), sides="b", color="grey50", alpha=0.5, linewidth=0.4) +
  scale_color_manual(values=c(
    "Exponencial"="#1976D2","Gamma"="#388E3C",
    "Log-Normal"="#F57C00","Weibull"="#7B1FA2")) +
  scale_linetype_manual(values=c(
    "Exponencial"="solid","Gamma"="dashed",
    "Log-Normal"="dotdash","Weibull"="dotted")) +
  labs(title="Histograma y Densidades Ajustadas – Duración de Tormentas",
       subtitle=sprintf("n=%d | Media=%.1fh | Cv=%.3f | Recomendada: %s",
                        n, mean(x), sd(x)/mean(x), dist_rec),
       x="Duración de tormenta (horas)", y="Densidad",
       color="Distribución", linetype="Distribución") +
  tema + theme(legend.position="bottom")

# G14: QQ-plot distribución recomendada
q_teo <- sapply(df_emp$P_emp, q_rec_fn)
g14 <- ggplot(data.frame(obs=df_emp$duracion_h, teo=q_teo), aes(teo, obs)) +
  geom_abline(slope=1, intercept=0, linetype="dashed", color="grey50", linewidth=0.8) +
  geom_point(size=2.8, color="#1976D2", alpha=0.8) +
  labs(title=sprintf("QQ-Plot – %s", dist_rec),
       subtitle=sprintf("KS p-valor = %.4f | %s | n=%d",
                        tabla_ks$p_valor[tabla_ks$Distribucion==dist_rec],
                        tabla_ks$Resultado[tabla_ks$Distribucion==dist_rec], n),
       x="Duración teórica (horas)", y="Duración observada (horas)") +
  tema

# Guardar gráficos
ruta_pdf_freq <- paste0(RUTA_SALIDA,"Frecuencia_Duraciones_",ESTACION,".pdf")
pdf(ruta_pdf_freq, width=12, height=7, onefile=TRUE)
suppressWarnings({ print(g11); print(g12); print(g13); print(g14) })
dev.off()
for (nm in c("G11_Papel_Probabilidad","G12_Curva_Tr",
             "G13_Histograma_Densidad","G14_QQplot")) {
  gi <- list(g11,g12,g13,g14)[[match(nm, c("G11_Papel_Probabilidad","G12_Curva_Tr",
                                             "G13_Histograma_Densidad","G14_QQplot"))]]
  suppressWarnings(ggsave(paste0(RUTA_SALIDA,nm,".png"), gi, width=12, height=6, dpi=150))
}
cat(sprintf("OK %s\n", ruta_pdf_freq))

# ── 15.6 Exportar al Excel IETD ──────────────────────────────────────────────
wb_ietd2 <- loadWorkbook(ruta_ietd)

# Hoja 1: Muestra completa con probabilidades empíricas
addWorksheet(wb_ietd2, "Duraciones_Empiricas")
writeData(wb_ietd2, "Duraciones_Empiricas",
  df_emp %>% rename(Duracion_h=duracion_h, Rango=rango,
                    P_Weibull=P_emp, Tr_Weibull=Tr_emp, P_Excedencia=P_exc) %>%
    mutate(across(where(is.numeric), ~round(.x,4))))
addStyle(wb_ietd2,"Duraciones_Empiricas", eh, rows=1, cols=1:5, gridExpand=TRUE)
setColWidths(wb_ietd2,"Duraciones_Empiricas", cols=1:5, widths=18)

# Hoja 2: Bondad de ajuste
addWorksheet(wb_ietd2, "Bondad_Ajuste_KS")
meta_ks <- data.frame(
  Campo=c("N tormentas","Media (h)","Mediana (h)","SD (h)","CV","Cs (asimetria)",
          "Distribucion recomendada","Criterio","Nota"),
  Valor=c(n, round(mean(x),2), round(median(x),2), round(sd(x),2),
          round(sd(x)/mean(x),3), round(Cs,3), dist_rec,
          "Mayor p-valor KS entre distribuciones aceptadas (p>0.05)",
          "Si CV~1 confirma proceso Poisson (Restrepo-Posada & Eagleson 1982)"))
writeData(wb_ietd2,"Bondad_Ajuste_KS", meta_ks, startRow=1)
writeData(wb_ietd2,"Bondad_Ajuste_KS", tabla_ks, startRow=nrow(meta_ks)+3)
hr <- nrow(meta_ks)+3
addStyle(wb_ietd2,"Bondad_Ajuste_KS", eh, rows=c(1,hr),
         cols=1:max(ncol(meta_ks),ncol(tabla_ks)), gridExpand=TRUE)
for (r in seq_len(nrow(tabla_ks))) {
  est <- if (tabla_ks$p_valor[r]>=0.05)
    createStyle(fgFill="#C8E6C9",fontName="Arial",fontSize=9,halign="CENTER")
  else
    createStyle(fgFill="#FFCDD2",fontName="Arial",fontSize=9,halign="CENTER")
  addStyle(wb_ietd2,"Bondad_Ajuste_KS", est,
           rows=hr+r, cols=1:ncol(tabla_ks), gridExpand=TRUE)
}
setColWidths(wb_ietd2,"Bondad_Ajuste_KS", cols=1:ncol(tabla_ks), widths=c(16,45,12,10,14))

# Hoja 3: Duraciones de diseño
addWorksheet(wb_ietd2, "Duraciones_Diseno")
meta_dd <- data.frame(
  Campo=c("Estacion","Variable analizada","Distribucion recomendada",
          "Metodo parametros","Bondad de ajuste","IC bootstrap",
          "N remuestras bootstrap","Unidad","Interpretacion"),
  Valor=c(ESTACION,
          "Duracion de tormentas independientes (horas)",
          dist_rec,"Momentos muestrales",
          sprintf("KS p-valor = %.4f (%s)",
                  tabla_ks$p_valor[tabla_ks$Distribucion==dist_rec],
                  tabla_ks$Resultado[tabla_ks$Distribucion==dist_rec]),
          "Percentiles 2.5% y 97.5%", B, "HORAS",
          paste0("Para una probabilidad P dada, la duracion de diseno es el cuantil ",
                 "de la distribucion ajustada. Ejemplo: P=0.90 (Tr=10a) -> ",
                 round(q_rec_fn(0.90),1)," h es la duracion que solo se supera ",
                 "el 10% de las tormentas.")))
writeData(wb_ietd2,"Duraciones_Diseno", meta_dd, startRow=1)
addStyle(wb_ietd2,"Duraciones_Diseno", eh, rows=1, cols=1:2, gridExpand=TRUE)
dd_row <- nrow(meta_dd)+3
writeData(wb_ietd2,"Duraciones_Diseno", tabla_diseño, startRow=dd_row)
addStyle(wb_ietd2,"Duraciones_Diseno", eh,
         rows=dd_row, cols=1:ncol(tabla_diseño), gridExpand=TRUE)
# Resaltar columna de duración recomendada
addStyle(wb_ietd2,"Duraciones_Diseno",
  createStyle(fgFill="#E3F2FD",fontName="Arial",fontSize=9,
              halign="CENTER",textDecoration="BOLD"),
  rows=(dd_row+1):(dd_row+nrow(tabla_diseño)), cols=3, gridExpand=TRUE)
setColWidths(wb_ietd2,"Duraciones_Diseno", cols=1:ncol(tabla_diseño), widths=20)

saveWorkbook(wb_ietd2, ruta_ietd, overwrite=TRUE)
cat(sprintf("OK Hojas de frecuencia actualizadas: %s\n", ruta_ietd))

# ── 15.7 Reporte consola ─────────────────────────────────────────────────────
cat("\n", strrep("=",70), "\n")
cat("  DURACIONES DE DISEÑO – RESTREPO-POSADA & EAGLESON (1982)\n")
cat(sprintf("  Distribución: %-20s | KS p-valor: %.4f\n",
    dist_rec, tabla_ks$p_valor[tabla_ks$Distribucion==dist_rec]))
cat(sprintf("  N tormentas:  %-6d | CV=%.3f | IETD=%dh\n", n, sd(x)/mean(x), IETD_OPT))
cat(strrep("-",70), "\n")
cat(sprintf("  %-8s  %-18s  %-18s  %s\n",
    "Tr (años)","P no-excedencia","Duración diseño (h)","IC 95% (h)"))
cat(strrep("-",70), "\n")
for (i in seq_along(P_diseño))
  cat(sprintf("  %-8.0f  %-18.4f  %-18.2f  [%.2f – %.2f]\n",
      Tr_diseño[i], P_diseño[i], q_pts[i], ic_low[i], ic_hig[i]))
cat(strrep("-",70), "\n")
cat(sprintf("  Interpretacion: una tormenta de diseno con P=%.2f (Tr=%.0f a)\n",
    P_diseño[which(Tr_diseño==10)], 10))
cat(sprintf("  tiene duracion de %.2f h (IC95: %.2f-%.2f h)\n",
    q_pts[which(Tr_diseño==10)],
    ic_low[which(Tr_diseño==10)],
    ic_hig[which(Tr_diseño==10)]))
cat(strrep("=",70), "\n")
cat("OK Seccion 15 completada.\n")
