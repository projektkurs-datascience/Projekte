#!/usr/bin/env Rscript

suppressPackageStartupMessages({
  if (!requireNamespace("yaml", quietly = TRUE)) {
    install.packages("yaml", repos = "https://cloud.r-project.org")
  }
  if (!requireNamespace("readxl", quietly = TRUE)) {
    install.packages("readxl", repos = "https://cloud.r-project.org")
  }
})

`%||%` <- function(a, b) if (is.null(a)) b else a

# ------------------------------------------------------------
# Basisverzeichnisse
# ------------------------------------------------------------
ROOT <- normalizePath(".", winslash = "/")
PROJECTS_ROOT <- file.path(ROOT, "projects")  # ggf. auf "quarto_pages/projects" ändern

# ------------------------------------------------------------
# Defaults
# ------------------------------------------------------------
DEFAULTS <- list(
  project = list(
    type = "descriptive",
    semester = "WS 25-26",
    categories = character(0),
    description = "Kurze Beschreibung des Projekts und Motivation"
  ),
  student = list(
    authors = character(0)
  )
)

# ------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------
slug_to_title <- function(s) {
  s <- gsub("[-_]+", " ", s)
  s <- trimws(s)
  tools::toTitleCase(s)
}

read_meta <- function(path) {
  if (file.exists(path)) {
    m <- yaml::yaml.load_file(path)
    if (is.null(m)) list() else m
  } else list()
}

yaml_vec <- function(x) {
  x <- as.character(x)
  x <- x[nzchar(trimws(x))]
  if (length(x) == 0) "[]"
  else paste0("[", paste(sprintf('"%s"', gsub('"','\\"', x, fixed = TRUE)), collapse = ", "), "]")
}

yaml_scalar <- function(x) {
  if (is.null(x) || is.na(x) || !nzchar(trimws(as.character(x)))) return(NA_character_)
  xn <- suppressWarnings(as.numeric(x))
  if (!is.na(xn)) as.character(xn) else sprintf('"%s"', gsub('"','\\"', as.character(x), fixed = TRUE))
}

clean_authors <- function(...) {
  x <- unlist(list(...), use.names = FALSE)
  x <- as.character(x)
  x <- x[!is.na(x)]
  x <- trimws(x)
  x <- x[nzchar(x)]
  x <- unlist(strsplit(x, "\\s*,\\s*"))
  x <- trimws(x)
  x <- x[nzchar(x)]
  unique(x)
}

# ------------------------------------------------------------
# HTML & MP4 finden
# ------------------------------------------------------------
find_student_html <- function(sdir, project_name = NULL) {
  files <- list.files(sdir, full.names = FALSE)
  if (!length(files)) return(NULL)
  hits <- grep("^projekt_.*\\.html?$", tolower(files))
  if (length(hits)) files[hits[1]] else NULL
}

find_student_mp4 <- function(sdir) {
  mp4s <- list.files(sdir, pattern = "\\.mp4$", full.names = FALSE, ignore.case = TRUE)
  if (length(mp4s) == 0) return(NULL)
  if (length(mp4s) == 1) return(mp4s[1])
  pref <- mp4s[tolower(mp4s) == "screencast.mp4"]
  if (length(pref)) pref[1] else mp4s[1]
}

assets_section <- function(html_rel = NULL, mp4_rel = NULL) {
  parts <- character()
  if (!is.null(mp4_rel)) {
    parts <- c(
      parts,
      "**Screencast**", "",
      sprintf('<video controls width="100%%" style="margin-bottom: 2em;"><source src="%s" type="video/mp4"></video>', mp4_rel)
    )
  }
  if (!is.null(html_rel)) {
    parts <- c(
      parts,
      "**Ausarbeitung**", "",
      sprintf('<iframe src="%s" loading="lazy" width="100%%" height="1200" style="border:0"></iframe>', html_rel)
    )
  }
  if (length(parts) == 0) {
    parts <- c("> TODO: Lege `Projekt_<Projektname>.html` und/oder genau **eine** `.mp4` in diesen Ordner, dann wird die Einbettung automatisch angezeigt.")
  }
  paste(parts, collapse = "\n")
}

write_if_missing <- function(path, content) {
  if (!file.exists(path)) {
    dir.create(dirname(path), recursive = TRUE, showWarnings = FALSE)
    con <- file(path, open = "w+", encoding = "UTF-8")
    on.exit(close(con), add = TRUE)
    writeLines(content, con, useBytes = TRUE)
    message("created: ", path)
  } else {
    message("exists:  ", path)
  }
}

# ------------------------------------------------------------
# Excel pro Projektordner lesen
# Erwartet 'metadata.xlsx' oder 'metadata.xls' mit Spalten:
# student_folder, author1, author2, author3, rank
# ------------------------------------------------------------
read_project_metadata <- function(pdir) {
  candidates <- list.files(pdir, pattern = "^metadata\\.(xlsx|xls)$", ignore.case = TRUE, full.names = TRUE)
  if (!length(candidates)) return(NULL)
  f <- candidates[1]
  df <- readxl::read_excel(f)
  names(df) <- tolower(gsub("\\s+", "_", names(df)))
  needed <- c("student_folder", "author1", "author2", "author3", "rank")
  missing <- setdiff(needed, names(df))
  if (length(missing)) {
    stop("In '", basename(f), "' fehlen Spalten: ", paste(missing, collapse = ", "), " (im Ordner: ", pdir, ")")
  }
  df$student_folder <- as.character(df$student_folder)
  df$author1 <- as.character(df$author1)
  df$author2 <- as.character(df$author2)
  df$author3 <- as.character(df$author3)
  df$rank <- as.character(df$rank)
  df
}

# ---- Top-3 aus Excel bestimmen ---------------------------------------------
top3_student_folders <- function(md_df) {
  if (is.null(md_df) || !nrow(md_df)) return(character(0))
  rnk <- suppressWarnings(as.numeric(md_df$rank))
  rnk[is.na(rnk)] <- Inf
  ord <- order(rnk, tolower(md_df$student_folder), na.last = TRUE)
  head(md_df$student_folder[ord][is.finite(rnk[ord])], 3)
}
top3_filenames <- function(folders) {
  if (!length(folders)) return(character(0))
  paste0(folders, ".qmd")
}
top3_paths <- function(folders) {
  if (!length(folders)) return(character(0))
  paste0("student_projects/", folders, "/", folders, ".qmd")
}

# ------------------------------------------------------------
# Seiten-Generatoren
# ------------------------------------------------------------
generate_project_page <- function(pdir, md_df) {
  pname <- basename(pdir)
  meta <- DEFAULTS$project
  m <- read_meta(file.path(pdir, "meta.yml"))
  for (k in names(m)) meta[[k]] <- m[[k]]
  
  # Top-3 bestimmen
  t3_folders <- top3_student_folders(md_df)
  t3_files   <- top3_filenames(t3_folders)  # nur Dateinamen, z.B. "gruppe_a.qmd"
  t3_paths   <- top3_paths(t3_folders)      # Pfade für das Grid (optional nicht nötig)
  
  # YAML für die zwei Listings:
  # 1) Grid der Top-3: sortiere nach rank und zeige max. 3 Einträge
  # 2) Tabelle der restlichen: schließe Top-3 über 'exclude: filename: [...]' aus
  top3_block <- paste(
    "  - id: top3",
    '    contents: ["student_projects/*/*.qmd"]',
    "    type: grid",
    "    grid-columns: 3",
    "    sort: rank",
    "    max-items: 3",
    sep = "\n"
  )
  
  rest_block <- paste(
    "  - id: rest",
    '    contents: ["student_projects/*/*.qmd"]',
    "    type: table",
    "    sort: rank",
    "    fields: [rank, title, authors]",
    "    field-display-names:",
    '      rank: "Rang"',
    '      title: "Ausarbeitung"',
    '      authors: "Autor:innen"',
    if (length(t3_files)) {
      paste0("    exclude:\n",
             "      filename: [",
             paste(sprintf('"%s"', t3_files), collapse = ", "),
             "]")
    } else NULL,
    sep = "\n"
  )
  
  front <- paste(c(
    "---",
    sprintf('title: "%s"', (m$title %||% slug_to_title(pname))),
    sprintf('type: "%s"', meta$type),
    sprintf('semester: "%s"', meta$semester),
    sprintf("categories: %s", yaml_vec(meta$categories)),
    "listing:",
    top3_block,
    rest_block,
    "---"
  ), collapse = "\n")
  
  body <- paste(
    meta$description %||% "", "",
    "## Top 3 Ausarbeitungen",
    "",
    "::: {#top3}",
    ":::",
    "",
    "## Weitere Ausarbeitungen",
    "",
    "::: {#rest}",
    ":::",
    "",
    sep = "\n"
  )
  
  out_path <- file.path(pdir, sprintf("%s_page.qmd", pname))
  write_if_missing(out_path, paste(front, body, sep = "\n"))
}

generate_student_page <- function(sdir, project_name, md_df) {
  sname <- basename(sdir)
  m <- read_meta(file.path(sdir, "meta.yml"))
  title <- m$title %||% slug_to_title(sname)
  
  # Excel-Zeile für diesen student_folder
  authors <- character(0)
  rank_val <- NA_character_
  if (!is.null(md_df)) {
    row <- md_df[tolower(md_df$student_folder) == tolower(sname), , drop = FALSE]
    if (nrow(row) >= 1) {
      authors <- clean_authors(row$author1[1], row$author2[1], row$author3[1])
      rank_val <- as.character(row$rank[1])
    } else {
      message("Warnung: Kein Eintrag in metadata für student_folder='", sname, "'.")
    }
  }
  
  html_rel <- find_student_html(sdir, project_name)
  mp4_rel  <- find_student_mp4(sdir)
  
  front_lines <- c(
    "---",
    sprintf('title: "%s"', title),
    sprintf("authors: %s", yaml_vec(authors))
  )
  if (!is.null(rank_val) && nzchar(trimws(rank_val))) {
    front_lines <- c(front_lines, paste0("rank: ", yaml_scalar(rank_val)))
  }
  front_lines <- c(front_lines, "---")
  front <- paste(front_lines, collapse = "\n")
  
  authors_line <- if (length(authors)) paste0("Autoren: ", paste(authors, collapse = ", "), "\n") else ""
  
  body <- paste(
    #authors_line, "",
    assets_section(html_rel, mp4_rel),
    "",
    sep = "\n"
  )
  
  out_path <- file.path(sdir, sprintf("%s.qmd", sname))
  write_if_missing(out_path, paste(front, body, sep = "\n"))
}

# ------------------------------------------------------------
# Main
# ------------------------------------------------------------
main <- function() {
  if (!dir.exists(PROJECTS_ROOT)) {
    message("Verzeichnis nicht gefunden: ", PROJECTS_ROOT)
    return(invisible())
  }
  
  projects <- list.dirs(PROJECTS_ROOT, full.names = TRUE, recursive = FALSE)
  for (p in sort(projects)) {
    pname <- basename(p)
    
    # 1) Excel des Projekts lesen (für Top-3 & Team-Metadaten)
    md_df <- read_project_metadata(p)  # NULL, falls nicht vorhanden
    
    # 2) Projektseite mit Grid(Top3) + Table(Rest) erzeugen
    generate_project_page(p, md_df)
    
    # 3) Team-Seiten erzeugen/ergänzen
    sp_root <- file.path(p, "student_projects")
    if (dir.exists(sp_root)) {
      students <- list.dirs(sp_root, full.names = TRUE, recursive = FALSE)
      for (s in sort(students)) {
        generate_student_page(s, pname, md_df)
      }
    }
  }
}

main()
