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

# Names that should NOT be linked to a GitHub profile.
UNKNOWN_NAMES <- c("unknown", "unbekannt", "anonym", "")

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
  else paste0("[", paste(sprintf('"%s"', gsub('"', '\\"', x, fixed = TRUE)), collapse = ", "), "]")
}

yaml_scalar <- function(x) {
  if (is.null(x) || is.na(x) || !nzchar(trimws(as.character(x)))) return(NA_character_)
  xn <- suppressWarnings(as.numeric(x))
  if (!is.na(xn)) as.character(xn) else sprintf('"%s"', gsub('"', '\\"', as.character(x), fixed = TRUE))
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

is_unknown_author <- function(a) {
  tolower(trimws(a)) %in% UNKNOWN_NAMES
}

# Find a thumbnail image inside `dir`.
# Prefers a file literally called image.* over any other image.
find_thumbnail <- function(dir) {
  imgs <- list.files(
    dir,
    pattern = "\\.(png|jpe?g|gif|webp|svg)$",
    ignore.case = TRUE,
    full.names = FALSE
  )
  if (!length(imgs)) return(NULL)
  pref <- imgs[tolower(tools::file_path_sans_ext(imgs)) == "image"]
  if (length(pref)) pref[1] else imgs[1]
}

# Quarto `authors:` block. Real names get a GitHub URL,
# UNKNOWN_NAMES stay plain text (no link).
yaml_authors_block <- function(authors) {
  authors <- authors[nzchar(trimws(authors))]
  if (!length(authors)) return("authors: []")
  esc <- function(s) gsub('"', '\\"', s, fixed = TRUE)
  lines <- "authors:"
  for (a in authors) {
    a <- trimws(a)
    lines <- c(lines, sprintf('  - name: "%s"', esc(a)))
    if (!is_unknown_author(a)) {
      lines <- c(lines, sprintf('    url: "https://github.com/%s"', a))
    }
  }
  paste(lines, collapse = "\n")
}

# Pre-built Markdown for the listing column on the project page.
# Real names become [name](https://github.com/name); unknown names stay plain.
authors_to_markdown_links <- function(authors) {
  authors <- authors[nzchar(trimws(authors))]
  if (!length(authors)) return("")
  parts <- vapply(authors, function(a) {
    a <- trimws(a)
    if (is_unknown_author(a)) a
    else sprintf("[%s](https://github.com/%s)", a, a)
  }, character(1))
  paste(parts, collapse = ", ")
}

# ------------------------------------------------------------
# HTML & MP4 finden
# ------------------------------------------------------------
# Picks any "Projekt_*.html" (or .htm) in the student folder,
# regardless of how the parent project folder is named.
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
  
  # Project-level thumbnail (shown in project_overview grid)
  img_rel <- find_thumbnail(pdir)
  
  # Top-3 bestimmen
  t3_folders <- top3_student_folders(md_df)
  t3_files   <- top3_filenames(t3_folders)
  
  # Listing-Blöcke ---------------------------------------------------------
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
    "    fields: [rank, title, authors_links]",
    "    field-display-names:",
    '      rank: "Rang"',
    '      title: "Ausarbeitung"',
    '      authors_links: "Autor:innen"',
    if (length(t3_files)) {
      paste0("    exclude:\n",
             "      filename: [",
             paste(sprintf('"%s"', t3_files), collapse = ", "),
             "]")
    } else NULL,
    sep = "\n"
  )
  
  # Front matter -----------------------------------------------------------
  front_lines <- c(
    "---",
    sprintf('title: "%s"', (m$title %||% slug_to_title(pname))),
    sprintf('type: "%s"', meta$type),
    sprintf('semester: "%s"', meta$semester),
    sprintf("categories: %s", yaml_vec(meta$categories))
  )
  if (!is.null(img_rel)) {
    front_lines <- c(front_lines, sprintf('image: "%s"', img_rel))
  }
  if (!is.null(meta$description) && nzchar(trimws(meta$description))) {
    desc_escaped <- gsub('"', '\\"', meta$description, fixed = TRUE)
    front_lines <- c(front_lines, sprintf('description: "%s"', desc_escaped))
  }
  front_lines <- c(front_lines, "listing:", top3_block, rest_block, "---")
  front <- paste(front_lines, collapse = "\n")
  
  # Body -------------------------------------------------------------------
  # Description lebt jetzt im YAML (description:) und wird von Quarto
  # automatisch unter dem Titel und in Listings angezeigt.
  body <- paste(
    "## Top 3 Ausarbeitungen", "",
    "::: {#top3}", ":::", "",
    "## Weitere Ausarbeitungen", "",
    "::: {#rest}", ":::", "",
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
  img_rel  <- find_thumbnail(sdir)
  
  # Front matter -----------------------------------------------------------
  front_lines <- c(
    "---",
    sprintf('title: "%s"', title),
    yaml_authors_block(authors)
  )
  if (length(authors)) {
    md_links <- authors_to_markdown_links(authors)
    md_links_escaped <- gsub('"', '\\"', md_links, fixed = TRUE)
    # Field used by the "Weitere Ausarbeitungen" table on the project page,
    # so the column is clickable. Non-GitHub names ("Unknown") stay plain.
    front_lines <- c(front_lines, sprintf('authors_links: "%s"', md_links_escaped))
  }
  if (!is.null(rank_val) && nzchar(trimws(rank_val))) {
    front_lines <- c(front_lines, paste0("rank: ", yaml_scalar(rank_val)))
  }
  if (!is.null(img_rel)) {
    front_lines <- c(front_lines, sprintf('image: "%s"', img_rel))
  }
  front_lines <- c(front_lines, "---")
  front <- paste(front_lines, collapse = "\n")
  
  # Body -------------------------------------------------------------------
  body <- paste(
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
    md_df <- read_project_metadata(p)               # NULL falls nicht vorhanden
    generate_project_page(p, md_df)                 # Projektseite (Grid+Tabelle)
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