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
#   Victoria Niño, M.J. (2026). Design storm duration from hourly rainfall
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
#    10b. Coeficiente de avance (r)
#    11.  Gráficos G1–G15 (todos en un solo PDF + PNGs individuales)
#    12.  Exportación Excel general (Resultados_*.xlsx)
#    13.  Exportación Excel IETD (IETD_*.xlsx) — incluye r
#    14.  Reporte final en consola
#    15.  Análisis de frecuencia de duraciones + r en tabla_diseño
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
# Versión:   1.0.0  |  2026
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

RUTA_ARCHIVO <- "D:\\R\\IETD\\Datos_precipitacion_crudos.xls"
RUTA_SALIDA  <- "D:\\R\\IETD\\"

ESTACION     <- "La Primavera"
CODIGO       <- "NA"
MUNICIPIO    <- "Guadalajara de Buga"
DEPARTAMENTO <- "Valle del Cauca"
FUENTE       <- "CVC"
COORD_N      <- NA_real_
COORD_E      <- NA_real_
ALTURA_msnm  <- 1644

UMBRAL_LLUVIA   <- 0.1
MAX_PRECIP      <- 130
MAX_GAP_RELLENO <- 6

UMBRAL_ANUAL     <- 0.60
UMBRAL_HUM       <- 0.50
APLICAR_C3       <- FALSE
MAX_GAP_HUM_DIAS <- 15L
MESES_HUMEDA     <- c(3L,4L,5L,9L,10L,11L)
ANIO_INICIO_UTIL <- 2014L

IETD_CANDIDATOS  <- c(1,2,3,4,5,6,8,10,12,16,20,24)
IETD_MAX_FISICO  <- 6L
IETD_MIN_LLUVIA  <- UMBRAL_LLUVIA
IETD_MIN_PROFUNDIDAD <- 1.0

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
df_chr <- as.data.frame(
  lapply(df_raw, function(col)
    vapply(col, function(v)
      if (is.null(v) || (length(v)==1L && is.na(v))) NA_character_
      else as.character(v[[1]]), character(1))),
  stringsAsFactors = FALSE)

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

precip_lineal <- pmax(
  na.approx(df_serie$precip_qc, maxgap=MAX_GAP_RELLENO, na.rm=FALSE), 0, na.rm=TRUE)

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
# 7. DEPURACIÓN HIDROLÓGICA DE AÑOS (OMM 168)
# ─────────────────────────────────────────────────────────────────────────────
cat("\n======================================================================\n")
cat("  DEPURACION HIDROLOGICA – OMM N.168\n")
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
  
  data.frame(Anio=a, N_total=n_tot, N_validos=n_val,
             Disp_anual_pct=round(100*disp,1),
             N_hum_total=n_ht, N_hum_validos=n_hv,
             Disp_hum_pct=round(100*ifelse(is.na(disp_h),0,disp_h),1),
             Gap_max_anual_d=gap_ad, Gap_max_hum_d=gap_hd,
             C1=ifelse(c1,"OK","FALLA"), C2=ifelse(c2,"OK","FALLA"),
             C3=ifelse(c3,"FALLA","OK"),
             Decision=dec, Norma="OMM168", Causa=causa,
             stringsAsFactors=FALSE)
})))

anios_ok  <- tabla_dep$Anio[tabla_dep$Decision %in% c("VALIDO_C1","VALIDO_C2")]
anios_no  <- tabla_dep$Anio[tabla_dep$Decision=="RECHAZADO"]
n_c1 <- sum(tabla_dep$Decision=="VALIDO_C1")
n_c2 <- sum(tabla_dep$Decision=="VALIDO_C2")
n_no <- sum(tabla_dep$Decision=="RECHAZADO")

cat(sprintf("Validos C1: %d | C2: %d | Rechazados: %d\n", n_c1, n_c2, n_no))
cat(sprintf("Serie IETD: %d anos [%s]\n", length(anios_ok), paste(anios_ok, collapse=", ")))

if (length(anios_ok)==0) stop("Sin anos validos.")

df_dep <- df_serie %>%
  filter(year(fecha_hora) %in% anios_ok) %>%
  mutate(anio=year(fecha_hora), mes=month(fecha_hora),
         dia=day(fecha_hora), hora=hour(fecha_hora),
         Decision_anio = if_else(anio %in% tabla_dep$Anio[tabla_dep$Decision=="VALIDO_C1"],
                                 "VALIDO_C1","VALIDO_C2"))

cat(sprintf("Registros serie depurada: %d horas\n", nrow(df_dep)))

# ─────────────────────────────────────────────────────────────────────────────
# 8. ANÁLISIS EXPLORATORIO
# ─────────────────────────────────────────────────────────────────────────────
df_an <- df_serie %>%
  mutate(anio=year(fecha_hora), mes=month(fecha_hora),
         dia=day(fecha_hora), hora=hour(fecha_hora))

prec_orig <- df_an$precip_relleno[df_an$metodo_relleno=="Original"]

stats <- data.frame(
  Estadistica = c("N validos","N total horas","Media (mm/h)","Maximo (mm/h)","Total acumulado (mm)"),
  Valor = c(length(prec_orig), nrow(df_an),
            round(mean(prec_orig,na.rm=TRUE),4),
            round(max(prec_orig,na.rm=TRUE),2),
            round(sum(df_an$precip_relleno,na.rm=TRUE),1)))
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
            .groups="drop")

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
cat("\nMetricas validacion:\n"); print(metricas)

# ─────────────────────────────────────────────────────────────────────────────
# 10. MÓDULO IETD – RESTREPO-POSADA & EAGLESON (1982)
# ─────────────────────────────────────────────────────────────────────────────
cat("\n======================================================================\n")
cat("  ANALISIS IETD – RESTREPO-POSADA & EAGLESON (1982)\n")
cat("======================================================================\n")

identificar_tormentas <- function(precip_h, fechas, ietd_h) {
  lluvia <- !is.na(precip_h) & precip_h >= IETD_MIN_LLUVIA
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
  df_ev <- data.frame(
    idx = seq_along(lluvia), fecha = fechas, precip = precip_h, evento_id = evento
  ) %>% filter(evento_id > 0) %>%
    group_by(evento_id) %>%
    summarise(
      inicio            = first(fecha),
      fin               = last(fecha),
      duracion_h        = n(),
      duracion_h_real   = as.numeric(difftime(last(fecha), first(fecha), units = "hours")) + 1,
      profundidad_mm    = sum(precip, na.rm = TRUE),
      intensidad_max    = max(precip, na.rm = TRUE),
      hora_pico         = first(fecha[precip == max(precip, na.rm = TRUE)]),
      .groups = "drop"
    ) %>%
    filter(profundidad_mm >= IETD_MIN_PROFUNDIDAD)
  if (nrow(df_ev) > 0) {
    fuera <- with(df_ev, hora_pico < inicio | hora_pico > fin)
    if (any(fuera, na.rm = TRUE))
      df_ev$hora_pico[which(fuera)] <- df_ev$inicio[which(fuera)]
  }
  if (nrow(df_ev) < 2) return(list(tormentas = df_ev, cv = NA_real_, n = nrow(df_ev)))
  tbt <- as.numeric(difftime(df_ev$inicio[-1], df_ev$inicio[-nrow(df_ev)], units = "hours"))
  tbt <- tbt[tbt > 0]
  cv <- if (length(tbt) >= 2) sd(tbt) / mean(tbt) else NA_real_
  list(tormentas = df_ev, tbt = tbt, cv = cv, n = nrow(df_ev), media_tbt = mean(tbt), sd_tbt = sd(tbt))
}

precip_dep <- df_dep$precip_relleno
fechas_dep <- df_dep$fecha_hora

cat(sprintf("Evaluando %d valores IETD sobre %d horas...\n",
            length(IETD_CANDIDATOS), length(precip_dep)))

res_ietd <- lapply(IETD_CANDIDATOS, function(ietd_h) {
  r <- identificar_tormentas(precip_dep, fechas_dep, ietd_h)
  data.frame(IETD_h = ietd_h, N_tormentas = r$n,
             Media_TBT_h = round(ifelse(is.null(r$media_tbt), NA, r$media_tbt), 2),
             SD_TBT_h = round(ifelse(is.null(r$sd_tbt), NA, r$sd_tbt), 2),
             Cv = round(ifelse(is.na(r$cv), NA, r$cv), 4),
             Dif_Cv_1 = round(ifelse(is.na(r$cv), NA, abs(r$cv - 1)), 4))
})
tabla_ietd <- do.call(rbind, res_ietd)
cat("\nResultados por IETD candidato:\n")
print(tabla_ietd)

tabla_ietd_fis <- tabla_ietd[tabla_ietd$IETD_h <= IETD_MAX_FISICO, ]
if (nrow(tabla_ietd_fis) > 0 && any(!is.na(tabla_ietd_fis$Dif_Cv_1))) {
  idx_opt <- which.min(tabla_ietd_fis$Dif_Cv_1)
  IETD_OPT <- tabla_ietd_fis$IETD_h[idx_opt]
  CV_OPT <- tabla_ietd_fis$Cv[idx_opt]
} else {
  idx_opt <- which.min(tabla_ietd$IETD_h)
  IETD_OPT <- tabla_ietd$IETD_h[idx_opt]
  CV_OPT <- tabla_ietd$Cv[idx_opt]
}
cat(sprintf("\n>>> IETD OPTIMO: %d horas (Cv = %.4f)\n", IETD_OPT, CV_OPT))

res_opt <- identificar_tormentas(precip_dep, fechas_dep, IETD_OPT)
df_torm <- res_opt$tormentas
tbt_opt <- res_opt$tbt

df_torm <- df_torm %>%
  mutate(
    anio             = year(inicio),
    mes              = month(inicio),
    intensidad_media = round(profundidad_mm / duracion_h, 4),
    TBT_horas        = c(NA, round(tbt_opt, 2))
  )

# =============================================================================
# SECCIÓN 10b – COEFICIENTE DE AVANCE (r)
# =============================================================================
cat("\n======================================================================\n")
cat("  10b. COEFICIENTE DE AVANCE (r) — t_pico / D\n")
cat("======================================================================\n")

df_torm <- df_torm %>%
  mutate(r = as.numeric(difftime(hora_pico, inicio, units = "hours")) / duracion_h_real)

n_fuera <- sum(df_torm$r < 0 | df_torm$r > 1, na.rm = TRUE)
if (n_fuera > 0) {
  cat(sprintf("  ⚠️  %d tormentas con r fuera de [0,1]. Forzando a 0.\n", n_fuera))
  df_torm$r[df_torm$r < 0 | df_torm$r > 1] <- 0
} else {
  cat("  ✅  Todas las tormentas: r dentro de [0,1]\n")
}

r_mean   <- mean(df_torm$r, na.rm = TRUE)
r_median <- median(df_torm$r, na.rm = TRUE)
r_sd     <- sd(df_torm$r, na.rm = TRUE)
r_p25    <- quantile(df_torm$r, 0.25, na.rm = TRUE)
r_p75    <- quantile(df_torm$r, 0.75, na.rm = TRUE)
r_min    <- min(df_torm$r, na.rm = TRUE)
r_max    <- max(df_torm$r, na.rm = TRUE)
r_Cv     <- r_sd / r_mean
n_r      <- sum(!is.na(df_torm$r))
r_Cs     <- (n_r * sum((df_torm$r - r_mean)^3, na.rm = TRUE)) /
  ((n_r - 1) * (n_r - 2) * r_sd^3)

cat(sprintf("\n  --- Estadisticos de r (v1.0.0) ---\n"))
cat(sprintf("  N tormentas:    %d\n", nrow(df_torm)))
cat(sprintf("  Media r:        %.4f\n", r_mean))
cat(sprintf("  Mediana r:      %.4f\n", r_median))
cat(sprintf("  SD r:           %.4f\n", r_sd))
cat(sprintf("  Cv r:           %.4f\n", r_Cv))
cat(sprintf("  Asimetria Cs r: %.4f\n", r_Cs))
cat(sprintf("  P25 r:          %.4f\n", r_p25))
cat(sprintf("  P75 r:          %.4f\n", r_p75))
cat(sprintf("  Rango r:        [%.4f, %.4f]\n", r_min, r_max))
cat(sprintf("\n>>> Coef. avance recomendado (mediana): r = %.4f\n", r_median))
cat(sprintf("    Pico ocurre en el %.1f%% de la duracion\n", 100 * r_median))

r_percentiles <- data.frame(
  Percentil = c(0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95),
  r = round(quantile(df_torm$r, c(0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95), na.rm = TRUE), 4)
)
cat("\nPercentiles de r:\n")
print(r_percentiles)

cat("\nDistribucion de r (histograma):\n")
r_hist <- hist(df_torm$r, breaks = seq(0, 1, 0.1), plot = FALSE)
for (i in seq_along(r_hist$mids)) {
  barra <- paste(rep("█", round(50 * r_hist$density[i] / max(r_hist$density))), collapse = "")
  cat(sprintf("  [%.1f-%.1f): %s (n=%d)\n", r_hist$breaks[i], r_hist$breaks[i+1], barra, r_hist$counts[i]))
}

r_por_duracion <- df_torm %>%
  mutate(grupo_duracion = cut(duracion_h_real,
                              breaks = c(0, 2, 4, 6, 8, 12, 24, 100),
                              labels = c("1-2h","3-4h","5-6h","7-8h","9-12h","13-24h",">24h"))) %>%
  group_by(grupo_duracion) %>%
  summarise(n = n(), r_media = round(mean(r, na.rm = TRUE), 4),
            r_mediana = round(median(r, na.rm = TRUE), 4), .groups = "drop") %>%
  filter(!is.na(grupo_duracion))
cat("\n  r vs Duracion real:\n")
print(r_por_duracion)

df_serie_ietd <- data.frame(
  fecha_hora = fechas_dep, precip_mm_h = precip_dep,
  metodo_relleno = df_dep$metodo_relleno,
  anio = year(fechas_dep), mes = month(fechas_dep),
  dia = day(fechas_dep), hora = hour(fechas_dep))
ev_vec <- integer(length(precip_dep))
for (i in seq_len(nrow(df_torm))) {
  idx_ev <- which(fechas_dep >= df_torm$inicio[i] & fechas_dep <= df_torm$fin[i])
  ev_vec[idx_ev] <- df_torm$evento_id[i]
}
df_serie_ietd$evento_id <- ev_vec
df_serie_ietd$en_tormenta <- ev_vec > 0

# ─────────────────────────────────────────────────────────────────────────────
# 11. GRÁFICOS G1–G15
# ─────────────────────────────────────────────────────────────────────────────
cat("\nGenerando G1-G15...\n")
tema <- theme_minimal(base_size = 11) +
  theme(plot.title = element_text(face = "bold", hjust = .5),
        plot.subtitle = element_text(hjust = .5, color = "grey40"),
        panel.grid.minor = element_blank(), axis.title = element_text(size = 10))

g1 <- ggplot(tot_an, aes(anio, total_mm)) +
  geom_col(fill = "#2196F3", alpha = .75) +
  geom_smooth(method = "lm", se = TRUE, color = "#F44336", linewidth = .8, formula = y ~ x) +
  scale_x_continuous(breaks = seq(1970, 2030, 5)) +
  labs(title = paste("Precipitación Anual –", ESTACION),
       subtitle = "Totales acumulados con tendencia lineal",
       x = "Año", y = "mm/año") + tema

g2 <- ggplot(ciclo_m, aes(factor(mes, labels = month.abb), media_mm_mes)) +
  geom_col(fill = "#4CAF50", alpha = .8) +
  labs(title = "Ciclo Anual de Precipitación",
       subtitle = "Promedio mensual multianual (mm/mes)",
       x = "Mes", y = "mm/mes") + tema

g3 <- ggplot(ciclo_h, aes(hora)) +
  geom_col(aes(y = frac_pct), fill = "#3F51B5", alpha = .4) +
  geom_line(aes(y = intens * 5), color = "#E91E63", linewidth = 1) +
  scale_x_continuous(breaks = 0:23) +
  labs(title = "Ciclo Diurno de Precipitación", x = "Hora", y = "% / mm/h×5") + tema

g4 <- ggplot(tot_an, aes(anio, disp_pct)) +
  geom_col(aes(fill = disp_pct < (100 * UMBRAL_ANUAL)), show.legend = FALSE) +
  scale_fill_manual(values = c("FALSE" = "#66BB6A", "TRUE" = "#EF5350")) +
  geom_hline(yintercept = 100 * UMBRAL_ANUAL, linetype = "dashed") +
  scale_x_continuous(breaks = seq(1970, 2030, 5)) +
  scale_y_continuous(labels = function(x) paste0(x, "%"), limits = c(0, 105)) +
  labs(title = "Disponibilidad de Datos por Año", x = "Año", y = "%") + tema

p_pos <- prec_orig[prec_orig > 0 & !is.na(prec_orig)]
g5 <- ggplot(data.frame(p = p_pos), aes(p)) +
  geom_histogram(bins = 50, fill = "#FF9800", color = "white", alpha = .8) +
  scale_x_log10(labels = label_number()) +
  scale_y_log10(labels = label_number()) +
  labs(title = "Distribución de Intensidades",
       subtitle = sprintf("Umbral: %.1f mm/h", UMBRAL_LLUVIA),
       x = "Intensidad (mm/h)", y = "Frecuencia") + tema

hm <- df_an %>% group_by(anio, mes) %>%
  summarise(tot = sum(precip_relleno, na.rm = TRUE), .groups = "drop")
g6 <- ggplot(hm, aes(mes, anio, fill = tot)) +
  geom_tile(color = "white", linewidth = .2) +
  scale_fill_gradientn(colours = c("#E3F2FD", "#64B5F6", "#1976D2", "#0D47A1"), name = "mm/mes") +
  scale_x_continuous(breaks = 1:12, labels = month.abb) +
  labs(title = "Precipitación Mensual", x = "Mes", y = "Año") + tema

g7 <- ggplot(tabla_dep %>%
               mutate(Col = case_when(Decision == "VALIDO_C1" ~ "Valido C1",
                                      Decision == "VALIDO_C2" ~ "Valido C2", TRUE ~ "Rechazado")),
             aes(Anio, Disp_anual_pct, fill = Col)) +
  geom_col(width = .8) +
  geom_hline(yintercept = 100 * UMBRAL_ANUAL, linetype = "dashed", color = "#1565C0") +
  scale_fill_manual(values = c("Valido C1" = "#2E7D32", "Valido C2" = "#F9A825", "Rechazado" = "#C62828")) +
  scale_y_continuous(limits = c(0, 105), labels = function(x) paste0(x, "%")) +
  labs(title = "Depuración OMM N°168", x = "Año", y = "Disponibilidad (%)", fill = "Decisión") +
  tema + theme(legend.position = "bottom")

g8 <- ggplot(tabla_ietd, aes(IETD_h, Cv)) +
  geom_line(color = "#1976D2", linewidth = 1) +
  geom_point(size = 3, color = "#1976D2") +
  geom_point(data = tabla_ietd[which.min(tabla_ietd$Dif_Cv_1), ], size = 4, color = "#E53935") +
  geom_hline(yintercept = 1, linetype = "dashed", color = "#E53935") +
  geom_vline(xintercept = IETD_OPT, linetype = "dotted", color = "#E53935") +
  scale_x_continuous(breaks = IETD_CANDIDATOS) +
  labs(title = "Selección de IETD (Cv = 1)",
       subtitle = sprintf("IETD=%dh, Cv=%.3f", IETD_OPT, CV_OPT),
       x = "IETD candidato (horas)", y = "Cv") + tema

g9 <- ggplot(df_torm, aes(profundidad_mm)) +
  geom_histogram(bins = 40, fill = "#7B1FA2", color = "white", alpha = .8) +
  scale_x_log10(labels = label_number()) +
  labs(title = "Profundidades por Tormenta",
       subtitle = sprintf("N=%d | IETD=%dh", nrow(df_torm), IETD_OPT),
       x = "Profundidad (mm)", y = "Frecuencia") + tema

g10 <- ggplot(data.frame(tbt = tbt_opt), aes(tbt)) +
  geom_histogram(bins = 40, fill = "#00796B", color = "white", alpha = .8) +
  geom_vline(xintercept = mean(tbt_opt), linetype = "dashed", color = "#E53935") +
  labs(title = "Tiempos Entre Tormentas",
       subtitle = sprintf("Media=%.1fh | Cv=%.3f", mean(tbt_opt), CV_OPT),
       x = "TBT (horas)", y = "Frecuencia") + tema

g15 <- ggplot(df_torm, aes(r)) +
  geom_histogram(aes(y = after_stat(density)),
                 bins = max(8, round(sqrt(nrow(df_torm)))),
                 fill = "#E91E63", color = "white", alpha = 0.75) +
  geom_density(color = "#880E4F", linewidth = 1.1, bw = 0.05) +
  geom_vline(xintercept = r_median, linetype = "dashed", color = "#F44336", linewidth = 0.9) +
  annotate("text", x = min(r_median + 0.04, 0.95), y = Inf, vjust = 2,
           label = sprintf("Mediana r = %.4f", r_median),
           size = 3.5, color = "#F44336", hjust = 0) +
  scale_x_continuous(limits = c(0, 1), breaks = seq(0, 1, 0.1)) +
  labs(title = "Coeficiente de Avance (r)",
       subtitle = sprintf("n=%d | r_med=%.4f | SD=%.4f | Cv=%.4f",
                          nrow(df_torm), r_median, r_sd, r_Cv),
       x = "r = t_pico / D", y = "Densidad") + tema

# ─────────────────────────────────────────────────────────────────────────────
# 12. EXPORTACIÓN EXCEL GENERAL (Resultados_*.xlsx)
# ─────────────────────────────────────────────────────────────────────────────
cat("\nGenerando archivo Resultados...\n")
eh <- createStyle(fgFill="#1976D2", fontColour="#FFFFFF", textDecoration="BOLD")
wb <- createWorkbook()

meta_df <- data.frame(
  Campo = c("Estacion", "Version", "Periodo", "IETD_optimo", "Cv", "N_tormentas",
            "r_mediana", "r_media", "r_sd"),
  Valor = c(ESTACION, "v1.0.0",
            sprintf("%s a %s", format(fecha_min,"%Y-%m-%d"), format(fecha_max,"%Y-%m-%d")),
            IETD_OPT, round(CV_OPT,4), nrow(df_torm),
            round(r_median,4), round(r_mean,4), round(r_sd,4)))

addWorksheet(wb, "Metadatos")
writeData(wb, "Metadatos", meta_df)
addStyle(wb, "Metadatos", eh, rows = 1, cols = 1:2, gridExpand = TRUE)

addWorksheet(wb, "Evaluacion_Cv_IETD")
writeData(wb, "Evaluacion_Cv_IETD", tabla_ietd)
addStyle(wb, "Evaluacion_Cv_IETD", eh, rows = 1, cols = 1:ncol(tabla_ietd), gridExpand = TRUE)

df_torm_exp <- df_torm %>%
  mutate(inicio = format(inicio,"%Y-%m-%d %H:%M"),
         fin = format(fin,"%Y-%m-%d %H:%M"),
         hora_pico = format(hora_pico,"%Y-%m-%d %H:%M")) %>%
  rename(Evento_ID = evento_id, Inicio = inicio, Fin = fin,
         Hora_Pico = hora_pico, Duracion_h_lluvia = duracion_h,
         Duracion_h_real = duracion_h_real, Profundidad_mm = profundidad_mm,
         Intensidad_max = intensidad_max, Intensidad_media = intensidad_media,
         Coef_Avance_r = r, TBT_h = TBT_horas)

addWorksheet(wb, "Catalogo_Tormentas")
writeData(wb, "Catalogo_Tormentas", df_torm_exp)
addStyle(wb, "Catalogo_Tormentas", eh, rows = 1, cols = 1:ncol(df_torm_exp), gridExpand = TRUE)

addWorksheet(wb, "Estadisticas_r")
df_r_stats <- data.frame(
  Estadistica = c("N tormentas","Media r","Mediana r","SD r","Min","Max","P25","P75"),
  Valor = c(nrow(df_torm), round(r_mean,4), round(r_median,4), round(r_sd,4),
            round(r_min,4), round(r_max,4), round(r_p25,4), round(r_p75,4)))
writeData(wb, "Estadisticas_r", df_r_stats)
addStyle(wb, "Estadisticas_r", eh, rows = 1, cols = 1:2, gridExpand = TRUE)

addWorksheet(wb, "Percentiles_r")
writeData(wb, "Percentiles_r", r_percentiles)
addStyle(wb, "Percentiles_r", eh, rows = 1, cols = 1:2, gridExpand = TRUE)

addWorksheet(wb, "r_vs_Duracion")
writeData(wb, "r_vs_Duracion", r_por_duracion)
addStyle(wb, "r_vs_Duracion", eh, rows = 1, cols = 1:ncol(r_por_duracion), gridExpand = TRUE)

addWorksheet(wb, "Auditoria_OMM")
writeData(wb, "Auditoria_OMM", tabla_dep)
addStyle(wb, "Auditoria_OMM", eh, rows = 1, cols = 1:ncol(tabla_dep), gridExpand = TRUE)

ruta_xl <- paste0(RUTA_SALIDA, "Resultados_", ESTACION, ".xlsx")
saveWorkbook(wb, ruta_xl, overwrite = TRUE)
cat(sprintf("OK Resultados: %s\n", ruta_xl))

# =============================================================================
# 13. EXPORTACIÓN EXCEL IETD (IETD_*.xlsx)
# =============================================================================
cat("\nGenerando archivo IETD...\n")

wb_ietd <- createWorkbook()

# Hoja 1: Metadatos IETD
meta_ietd <- data.frame(
  Parametro = c("Estacion", "Version", "Municipio", "Departamento",
                "Periodo_inicio", "Periodo_fin", "Anos_validos",
                "IETD_optimo_horas", "Cv_obtenido", "N_tormentas",
                "Tasa_tormentas_anio", "TBT_medio_horas",
                "Duracion_media_h_lluvia", "Duracion_media_h_real",
                "r_mediana_v1.0.0", "r_media_v1.0.0", "r_sd_v1.0.0",
                "r_min", "r_max", "r_P25", "r_P75",
                "Fecha_proceso"),
  Valor = c(ESTACION, "v1.0.0", MUNICIPIO, DEPARTAMENTO,
            format(min(df_torm$inicio),"%Y-%m-%d"),
            format(max(df_torm$fin),"%Y-%m-%d"),
            paste(anios_ok, collapse=", "),
            IETD_OPT, round(CV_OPT,4), nrow(df_torm),
            round(nrow(df_torm)/length(anios_ok),2),
            round(mean(tbt_opt, na.rm=TRUE),2),
            round(mean(df_torm$duracion_h),2),
            round(mean(df_torm$duracion_h_real),2),
            round(r_median,4), round(r_mean,4), round(r_sd,4),
            round(r_min,4), round(r_max,4), round(r_p25,4), round(r_p75,4),
            format(Sys.Date(),"%Y-%m-%d")))

addWorksheet(wb_ietd, "Metadatos_IETD")
writeData(wb_ietd, "Metadatos_IETD", meta_ietd)
addStyle(wb_ietd, "Metadatos_IETD",
         createStyle(fgFill="#0D47A1", fontColour="#FFFFFF", textDecoration="BOLD",
                     fontName="Arial", fontSize=11, halign="CENTER"),
         rows=1, cols=1:2, gridExpand=TRUE)
setColWidths(wb_ietd, "Metadatos_IETD", cols=1:2, widths=c(38, 50))

# Hoja 2: Evaluacion Cv
addWorksheet(wb_ietd, "Evaluacion_Cv_IETD")
writeData(wb_ietd, "Evaluacion_Cv_IETD", tabla_ietd)
addStyle(wb_ietd, "Evaluacion_Cv_IETD", eh, rows=1, cols=1:ncol(tabla_ietd), gridExpand=TRUE)
addStyle(wb_ietd, "Evaluacion_Cv_IETD",
         createStyle(fgFill="#A5D6A7", fontName="Arial", textDecoration="BOLD", halign="CENTER"),
         rows=idx_opt+1, cols=1:ncol(tabla_ietd), gridExpand=TRUE)
setColWidths(wb_ietd, "Evaluacion_Cv_IETD", cols=1:ncol(tabla_ietd), widths=18)

# Hoja 3: Catalogo con r
df_torm_ietd <- df_torm %>%
  mutate(inicio = format(inicio,"%Y-%m-%d %H:%M"),
         fin = format(fin,"%Y-%m-%d %H:%M"),
         hora_pico = format(hora_pico,"%Y-%m-%d %H:%M")) %>%
  rename(Evento_ID = evento_id, Inicio = inicio, Fin = fin,
         Hora_Pico = hora_pico, Duracion_h_lluvia = duracion_h,
         Duracion_h_real = duracion_h_real, Profundidad_mm = profundidad_mm,
         Intensidad_max = intensidad_max, Intensidad_media = intensidad_media,
         Coef_Avance_r = r, TBT_h = TBT_horas)

addWorksheet(wb_ietd, "Catalogo_Tormentas")
writeData(wb_ietd, "Catalogo_Tormentas", df_torm_ietd)
addStyle(wb_ietd, "Catalogo_Tormentas", eh, rows=1, cols=1:ncol(df_torm_ietd), gridExpand=TRUE)
setColWidths(wb_ietd, "Catalogo_Tormentas", cols=1:ncol(df_torm_ietd), widths="auto")

# Hoja 4: TBT
df_tbt <- data.frame(
  N_par = seq_along(tbt_opt),
  Tormenta_i = df_torm_ietd$Evento_ID[-1],
  Tormenta_j = df_torm_ietd$Evento_ID[-nrow(df_torm_ietd)],
  TBT_horas = round(tbt_opt, 3),
  TBT_dias = round(tbt_opt/24, 4))
addWorksheet(wb_ietd, "TBT_Tiempos_Entre_Tormentas")
writeData(wb_ietd, "TBT_Tiempos_Entre_Tormentas", df_tbt)
addStyle(wb_ietd, "TBT_Tiempos_Entre_Tormentas", eh, rows=1, cols=1:ncol(df_tbt), gridExpand=TRUE)
setColWidths(wb_ietd, "TBT_Tiempos_Entre_Tormentas", cols=1:ncol(df_tbt), widths=18)

# Hoja 5: Estadisticas anuales con r
est_anual_ietd <- df_torm %>%
  group_by(anio) %>%
  summarise(N_tormentas = n(),
            Precip_total_mm = round(sum(profundidad_mm), 1),
            Duracion_media_h = round(mean(duracion_h), 1),
            Duracion_media_h_real = round(mean(duracion_h_real), 1),
            Profund_media_mm = round(mean(profundidad_mm), 1),
            Intens_max_mm_h = round(max(intensidad_max), 1),
            r_medio = round(mean(r, na.rm=TRUE), 4),
            r_mediana = round(median(r, na.rm=TRUE), 4),
            .groups = "drop") %>%
  left_join(tabla_dep %>% select(Anio, Disp_anual_pct, Decision),
            by = c("anio" = "Anio"))

addWorksheet(wb_ietd, "Estadisticas_Anuales_IETD")
writeData(wb_ietd, "Estadisticas_Anuales_IETD", est_anual_ietd)
addStyle(wb_ietd, "Estadisticas_Anuales_IETD", eh, rows=1, cols=1:ncol(est_anual_ietd), gridExpand=TRUE)
setColWidths(wb_ietd, "Estadisticas_Anuales_IETD", cols=1:ncol(est_anual_ietd), widths=20)

# Hoja 6: Estadisticas de r
df_r_stats_ietd <- data.frame(
  Estadistica = c("N tormentas", "Media r", "Mediana r", "SD r", "Cv r",
                  "Minimo r", "Maximo r", "P25 r", "P75 r",
                  "Asimetria Cs", "Duracion media (h lluvia)",
                  "Duracion media (h real)", "r recomendado (mediana)"),
  Valor = c(nrow(df_torm), round(r_mean,4), round(r_median,4), round(r_sd,4),
            round(r_Cv,4), round(r_min,4), round(r_max,4),
            round(r_p25,4), round(r_p75,4), round(r_Cs,4),
            round(mean(df_torm$duracion_h),2),
            round(mean(df_torm$duracion_h_real),2),
            round(r_median,4)))

addWorksheet(wb_ietd, "Estadisticas_r")
writeData(wb_ietd, "Estadisticas_r", df_r_stats_ietd)
addStyle(wb_ietd, "Estadisticas_r", eh, rows=1, cols=1:2, gridExpand=TRUE)
setColWidths(wb_ietd, "Estadisticas_r", cols=1:2, widths=c(35, 15))

# Hoja 7: Percentiles de r
addWorksheet(wb_ietd, "Percentiles_r")
writeData(wb_ietd, "Percentiles_r", r_percentiles)
addStyle(wb_ietd, "Percentiles_r", eh, rows=1, cols=1:2, gridExpand=TRUE)
setColWidths(wb_ietd, "Percentiles_r", cols=1:2, widths=18)

# Hoja 8: r vs Duracion
addWorksheet(wb_ietd, "r_vs_Duracion")
writeData(wb_ietd, "r_vs_Duracion", r_por_duracion)
addStyle(wb_ietd, "r_vs_Duracion", eh, rows=1, cols=1:ncol(r_por_duracion), gridExpand=TRUE)
setColWidths(wb_ietd, "r_vs_Duracion", cols=1:ncol(r_por_duracion), widths=18)

ruta_ietd <- paste0(RUTA_SALIDA, "IETD_", ESTACION, ".xlsx")
saveWorkbook(wb_ietd, ruta_ietd, overwrite = TRUE)
cat(sprintf("OK IETD: %s\n", ruta_ietd))
# =============================================================================

# ─────────────────────────────────────────────────────────────────────────────
# 14. REPORTE FINAL
# ─────────────────────────────────────────────────────────────────────────────
cat("\n", strrep("=", 70), "\n")
cat(sprintf("  REPORTE FINAL v1.0.0 – %s\n", ESTACION))
cat(strrep("-", 70), "\n")
cat(sprintf("  IETD: %dh (Cv=%.4f) | Tormentas: %d (%.1f/año)\n",
            IETD_OPT, CV_OPT, nrow(df_torm), nrow(df_torm)/length(anios_ok)))
cat(sprintf("  Duracion media: %.1f h (lluvia) / %.1f h (real)\n",
            mean(df_torm$duracion_h), mean(df_torm$duracion_h_real)))
cat(sprintf("  Profundidad media: %.1f mm\n", mean(df_torm$profundidad_mm)))
cat(sprintf("  TBT medio: %.1f h (%.1f dias)\n", mean(tbt_opt), mean(tbt_opt)/24))
cat(strrep("-", 70), "\n")
cat(sprintf("  >>> r: mediana=%.4f | media=%.4f | SD=%.4f\n", r_median, r_mean, r_sd))
cat(sprintf("  >>> Cv r=%.4f | Cs r=%.4f\n", r_Cv, r_Cs))
cat(sprintf("  >>> Pico en %.1f%% de la duracion\n", 100 * r_median))
cat(strrep("-", 70), "\n")
cat("  ARCHIVOS GENERADOS:\n")
cat(sprintf("  Resultados: %s\n", ruta_xl))
cat(sprintf("  IETD:       %s\n", ruta_ietd))
cat(sprintf("  PDF:        %s\n", paste0(RUTA_SALIDA, "Graficos_Precipitacion_", ESTACION, ".pdf")))
cat(sprintf("  PNGs:       %sG*.png\n", RUTA_SALIDA))
cat(strrep("=", 70), "\n")

# =============================================================================
# 15. FRECUENCIA DE DURACIONES + G11-G14
# =============================================================================
cat("\n", strrep("=", 70), "\n")
cat("  15. FRECUENCIA DE DURACIONES\n")
cat(strrep("=", 70), "\n")

x <- df_torm$duracion_h
n <- length(x)

lambda_exp <- 1 / mean(x)
q_exp_fn <- function(p) qexp(p, rate = lambda_exp)
p_exp_fn <- function(q) pexp(q, rate = lambda_exp)
ks_exp <- suppressWarnings(ks.test(x, p_exp_fn))

k_gam <- (mean(x) / sd(x))^2
th_gam <- sd(x)^2 / mean(x)
q_gam_fn <- function(p) qgamma(p, shape = k_gam, scale = th_gam)
p_gam_fn <- function(q) pgamma(q, shape = k_gam, scale = th_gam)
ks_gam <- suppressWarnings(ks.test(x, p_gam_fn))

lx <- log(x)
mu_ln <- mean(lx)
sd_ln <- sd(lx)
q_ln_fn <- function(p) qlnorm(p, meanlog = mu_ln, sdlog = sd_ln)
p_ln_fn <- function(q) plnorm(q, meanlog = mu_ln, sdlog = sd_ln)
ks_ln <- suppressWarnings(ks.test(x, p_ln_fn))

weibull_mle <- function(x) {
  cv <- sd(x) / mean(x); c0 <- (cv)^(-1.086)
  for (iter in 1:100) {
    sc <- sum(x^c0 * log(x)) / sum(x^c0)
    g <- 1 / c0 + mean(log(x)) - sc
    dg <- -1 / c0^2 - (sum(x^c0 * (log(x))^2) * sum(x^c0) - (sum(x^c0 * log(x))^2)) / sum(x^c0)^2
    c1 <- c0 - g / dg
    if (abs(c1 - c0) < 1e-8) break
    c0 <- max(c1, 0.01)
  }
  list(shape = c0, scale = (mean(x^c0))^(1 / c0))
}
wb_par <- weibull_mle(x)
c_wb <- wb_par$shape; u_wb <- wb_par$scale
q_wb_fn <- function(p) qweibull(p, shape = c_wb, scale = u_wb)
p_wb_fn <- function(q) pweibull(q, shape = c_wb, scale = u_wb)
ks_wb <- suppressWarnings(ks.test(x, p_wb_fn))

tabla_ks <- data.frame(
  Distribucion = c("Exponencial", "Gamma", "Log-Normal", "Weibull"),
  KS_D = round(c(ks_exp$statistic, ks_gam$statistic, ks_ln$statistic, ks_wb$statistic), 5),
  p_valor = round(c(ks_exp$p.value, ks_gam$p.value, ks_ln$p.value, ks_wb$p.value), 4))
cat("\nBondad de ajuste:\n"); print(tabla_ks)

dist_rec <- tabla_ks$Distribucion[which.min(tabla_ks$KS_D)]
cat(sprintf("Recomendada: %s (menor D de KS)\n", dist_rec))
q_rec_fn <- switch(dist_rec, "Exponencial" = q_exp_fn, "Gamma" = q_gam_fn,
                   "Log-Normal" = q_ln_fn, "Weibull" = q_wb_fn)

# Cuantiles de diseño
P_diseño <- c(0.50, 0.75, 0.80, 0.90, 0.95, 0.98, 0.99, 0.995)
Tr_diseño <- round(1 / (1 - P_diseño), 0)

set.seed(42); B <- 1000
boot_q <- matrix(NA_real_, nrow = B, ncol = length(P_diseño))
for (b in seq_len(B)) {
  xb <- sample(x, n, replace = TRUE)
  qfn <- switch(dist_rec,
                "Exponencial" = { lb <- 1/mean(xb); function(p) qexp(p, rate=lb) },
                "Gamma" = { kb <- (mean(xb)/sd(xb))^2; tb <- sd(xb)^2/mean(xb); function(p) qgamma(p, shape=kb, scale=tb) },
                "Log-Normal" = { mlb <- mean(log(xb)); slb <- sd(log(xb)); function(p) qlnorm(p, meanlog=mlb, sdlog=slb) },
                "Weibull" = { pb <- weibull_mle(xb); function(p) qweibull(p, shape=pb$shape, scale=pb$scale) })
  boot_q[b, ] <- sapply(P_diseño, qfn)
}
ic_low <- apply(boot_q, 2, quantile, 0.025, na.rm = TRUE)
ic_hig <- apply(boot_q, 2, quantile, 0.975, na.rm = TRUE)
q_pts <- sapply(P_diseño, q_rec_fn)

R_DISENO <- round(r_median, 4)
tabla_diseño <- data.frame(
  P_no_excedencia = P_diseño, Tr_años = Tr_diseño,
  Duracion_diseño_h = round(q_pts, 2),
  IC95_inf = round(ic_low, 2), IC95_sup = round(ic_hig, 2),
  Coef_Avance_r = R_DISENO, Distribucion = dist_rec)
cat("\nDuraciones de diseño:\n"); print(tabla_diseño)

# G11-G14
df_emp <- data.frame(duracion_h = sort(x)) %>%
  mutate(P_emp = seq_len(n) / (n + 1), Tr_emp = 1 / (1 - P_emp))

p_curva <- seq(0.01, 0.995, length.out = 400)
df_teoricas <- bind_rows(
  data.frame(P = p_curva, dur = q_exp_fn(p_curva), Dist = "Exponencial"),
  data.frame(P = p_curva, dur = q_gam_fn(p_curva), Dist = "Gamma"),
  data.frame(P = p_curva, dur = q_ln_fn(p_curva), Dist = "Log-Normal"),
  data.frame(P = p_curva, dur = q_wb_fn(p_curva), Dist = "Weibull")) %>%
  filter(is.finite(dur), dur > 0)

breaks_Tr <- c(1.05, 1.25, 2, 5, 10, 25, 50, 100)
breaks_P <- 1 - 1 / breaks_Tr

g11 <- ggplot() +
  geom_line(data = df_teoricas, aes(P, dur, color = Dist, linetype = Dist), linewidth = 0.9) +
  geom_point(data = df_emp, aes(P_emp, duracion_h), color = "black", size = 2, alpha = 0.5) +
  geom_point(data = tabla_diseño, aes(P_no_excedencia, Duracion_diseño_h),
             color = "#E53935", size = 4, shape = 18) +
  geom_errorbar(data = tabla_diseño, aes(x = P_no_excedencia, ymin = IC95_inf, ymax = IC95_sup),
                color = "#E53935", width = 0.01) +
  scale_x_continuous(breaks = breaks_P,
                     labels = paste0("Tr=", breaks_Tr, "a\n(P=", round(breaks_P,2), ")"), limits = c(0, 0.998)) +
  scale_color_manual(values = c("Exponencial"="#1976D2","Gamma"="#388E3C",
                                "Log-Normal"="#F57C00","Weibull"="#7B1FA2")) +
  scale_linetype_manual(values = c("Exponencial"="solid","Gamma"="dashed",
                                   "Log-Normal"="dotdash","Weibull"="dotted")) +
  annotate("text", x = 0.02, y = max(x) * 0.9,
           label = sprintf("Recom: %s\nn=%d | r=%.4f", dist_rec, n, R_DISENO),
           hjust = 0, size = 3.5, color = "#E53935") +
  labs(title = "Frecuencia de Duraciones", x = "Probabilidad no excedencia", y = "Duracion (h)",
       color = "Dist.", linetype = "Dist.") + tema + theme(legend.position = "bottom")

x_curva <- seq(min(x)*0.5, max(x)*1.15, length.out = 400)
df_dens <- bind_rows(
  data.frame(x = x_curva, d = dexp(x_curva, rate = lambda_exp), Dist = "Exponencial"),
  data.frame(x = x_curva, d = dgamma(x_curva, shape = k_gam, scale = th_gam), Dist = "Gamma"),
  data.frame(x = x_curva, d = dlnorm(x_curva, meanlog = mu_ln, sdlog = sd_ln), Dist = "Log-Normal"),
  data.frame(x = x_curva, d = dweibull(x_curva, shape = c_wb, scale = u_wb), Dist = "Weibull")) %>%
  filter(is.finite(d), d >= 0)

g13 <- ggplot(data.frame(x = x), aes(x)) +
  geom_histogram(aes(y = after_stat(density)), bins = max(8, round(sqrt(n))),
                 fill = "#B0BEC5", color = "white", alpha = 0.85) +
  geom_line(data = df_dens, aes(x, d, color = Dist, linetype = Dist), linewidth = 0.9) +
  scale_color_manual(values = c("Exponencial"="#1976D2","Gamma"="#388E3C",
                                "Log-Normal"="#F57C00","Weibull"="#7B1FA2")) +
  scale_linetype_manual(values = c("Exponencial"="solid","Gamma"="dashed",
                                   "Log-Normal"="dotdash","Weibull"="dotted")) +
  labs(title = "Histograma y Densidades", x = "Duracion (h)", y = "Densidad",
       color = "Dist.", linetype = "Dist.") + tema + theme(legend.position = "bottom")

q_teo <- sapply(df_emp$P_emp, q_rec_fn)
g14 <- ggplot(data.frame(obs = df_emp$duracion_h, teo = q_teo), aes(teo, obs)) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "grey50") +
  geom_point(size = 2.5, color = "#1976D2", alpha = 0.8) +
  labs(title = sprintf("QQ-Plot – %s", dist_rec),
       subtitle = sprintf("n=%d | r=%.4f", n, R_DISENO),
       x = "Teorico (h)", y = "Observado (h)") + tema

Tr_curva <- 1 / (1 - seq(0.01, 0.995, length.out = 300))
df_curva <- data.frame(Tr = Tr_curva, dur = sapply(seq(0.01, 0.995, length.out = 300), q_rec_fn)) %>%
  filter(is.finite(dur), dur > 0)
g12 <- ggplot() +
  geom_ribbon(data = data.frame(Tr = Tr_diseño, lo = ic_low, hi = ic_hig),
              aes(Tr, ymin = lo, ymax = hi), fill = "#1976D2", alpha = 0.15) +
  geom_line(data = df_curva, aes(Tr, dur), color = "#1976D2", linewidth = 1.2) +
  geom_point(data = tabla_diseño, aes(Tr_años, Duracion_diseño_h), color = "#1976D2", size = 3) +
  scale_x_log10(breaks = c(1,2,5,10,25,50,100,200),
                labels = c("1","2","5","10","25","50","100","200")) +
  labs(title = sprintf("Curva Tr (%s)", dist_rec),
       subtitle = sprintf("Bootstrap B=%d | r=%.4f", B, R_DISENO),
       x = "Tr (años)", y = "Duracion (h)") + tema

# =============================================================================
# GUARDAR PDF ÚNICO CON G1-G15
# =============================================================================
cat("\n--- Guardando PDF (G1-G15) ---\n")
ruta_pdf <- paste0(RUTA_SALIDA, "Graficos_Precipitacion_", ESTACION, ".pdf")
pdf(ruta_pdf, width = 12, height = 7, onefile = TRUE)
suppressWarnings({
  print(g1); print(g2); print(g3); print(g4); print(g5)
  print(g6); print(g7); print(g8); print(g9); print(g10)
  print(g11); print(g12); print(g13); print(g14); print(g15)
})
dev.off()
cat(sprintf("OK PDF con 15 graficos: %s\n", ruta_pdf))

# =============================================================================
# GUARDAR PNGs G1-G15
# =============================================================================
cat("\n--- Guardando PNGs (G1-G15) ---\n")
nombres_png <- c("G1_Serie_Anual","G2_Ciclo_Mensual","G3_Ciclo_Diurno","G4_Disponibilidad",
                 "G5_Distribucion","G6_Heatmap","G7_Depuracion","G8_IETD_Cv",
                 "G9_Profundidades","G10_TBT",
                 "G11_Papel_Probabilidad","G12_Curva_Tr","G13_Histograma_Densidad","G14_QQplot",
                 "G15_Coef_Avance_r")
lista_graf <- list(g1,g2,g3,g4,g5,g6,g7,g8,g9,g10,g11,g12,g13,g14,g15)
png_ok <- 0; png_err <- 0

for (i in seq_along(nombres_png)) {
  ruta_png <- paste0(RUTA_SALIDA, nombres_png[i], ".png")
  tryCatch({
    ggsave(ruta_png, plot = lista_graf[[i]], width = 12, height = 6, dpi = 150)
    if (file.exists(ruta_png) && file.info(ruta_png)$size > 0) {
      cat(sprintf("  OK %s (%d KB)\n", nombres_png[i],
                  round(file.info(ruta_png)$size / 1024)))
      png_ok <- png_ok + 1
    } else {
      cat(sprintf("  ERROR %s: archivo vacio\n", nombres_png[i]))
      png_err <- png_err + 1
    }
  }, error = function(e) {
    cat(sprintf("  ERROR %s: %s\n", nombres_png[i], e$message))
    png_err <<- png_err + 1
  })
}
cat(sprintf("\n  PNGs: %d OK, %d errores de 15\n", png_ok, png_err))

# ─────────────────────────────────────────────────────────────────────────────
# FIN DEL SCRIPT v1.0.0
# ─────────────────────────────────────────────────────────────────────────────
cat("\nOK v1.0.0 completado.\n")
