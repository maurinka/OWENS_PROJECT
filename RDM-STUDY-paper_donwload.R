
# == Automate RDM-STUDY paper download and knowledge base ======================



# Install required packages ----

required_packages <- c(
  "readxl",
  "openxlsx",
  "dplyr",
  "stringr",
  "stringi",
  "purrr",
  "tibble",
  "httr2"
)

packages_to_install <- required_packages[
  !required_packages %in% rownames(installed.packages())
]

if (length(packages_to_install) > 0) {
  install.packages(packages_to_install)
}

invisible(
  lapply(
    required_packages,
    library,
    character.only = TRUE
  )
)


# User settings ----

# Main project folder.
# Forward slashes work correctly in R on Windows.
root_dir <- Sys.getenv("RDM_ROOT_DIR", unset = "/Users/maura/Desktop/KnowldegeBase")

# Input files.
knowledge_base_file <- file.path(
  root_dir,
  "KnowldegeBase.xlsx"
)

openalex_file <- file.path(
  root_dir,
  "open_alex_full_new.xlsx"
)

# Output files.
master_file <- file.path(
  root_dir,
  "reconciled_master_corrected.xlsx"
)

download_log_file <- file.path(
  root_dir,
  "download_log.csv"
)

# Folder in which study folders will be stored.
studies_dir <- root_dir

# IMPORTANT:
# Keep FALSE for the first run.
# Inspect reconciled_master.xlsx first.
# Change to TRUE only when you are satisfied with the matching.
DOWNLOAD_PDFS <- FALSE

# Optional safety limit for the first download test.
#
# Example:
# MAX_DOWNLOADS <- 5
#
# Set to Inf once you are comfortable with the workflow.
MAX_DOWNLOADS <- 5


# OpenAlex API key ----

# Recommended:
# Store your API key in an environment variable called
# OPENALEX_API_KEY.
#
# For a temporary R session you can run:
#
# Sys.setenv(OPENALEX_API_KEY = "your-key-here")
#
# Do not save your real API key in a script that may be shared.

openalex_api_key <- Sys.getenv("OPENALEX_API_KEY")

if (DOWNLOAD_PDFS && openalex_api_key == "") {
  stop(
    paste0(
      "No OpenAlex API key found.\n",
      "Run:\n",
      'Sys.setenv(OPENALEX_API_KEY = "your-key-here")'
    )
  )
}


# Check input files ----

if (!file.exists(knowledge_base_file)) {
  stop(
    paste(
      "Knowledge Base file not found:",
      knowledge_base_file
    )
  )
}

if (!file.exists(openalex_file)) {
  stop(
    paste(
      "OpenAlex file not found:",
      openalex_file
    )
  )
}

# Reconciliation does not create or move study folders.


# Helper function: clean column names ----

clean_names_simple <- function(x) {
  
  x |>
    stringi::stri_trans_general("Latin-ASCII") |>
    stringr::str_to_lower() |>
    stringr::str_replace_all("[^a-z0-9]+", "_") |>
    stringr::str_replace_all("^_|_$", "")
}


# Helper function: find a column by possible names ----

find_column <- function(
    df,
    candidates,
    required = FALSE,
    label = NULL
) {
  
  names_clean <- clean_names_simple(names(df))
  candidates_clean <- clean_names_simple(candidates)
  
  match_position <- match(
    candidates_clean,
    names_clean,
    nomatch = 0
  )
  
  match_position <- match_position[
    match_position > 0
  ]
  
  if (length(match_position) == 0) {
    
    if (required) {
      
      if (is.null(label)) {
        label <- paste(candidates, collapse = ", ")
      }
      
      stop(
        paste0(
          "Could not identify the column for: ",
          label,
          "\n\nAvailable columns are:\n",
          paste(names(df), collapse = "\n")
        )
      )
    }
    
    return(NULL)
  }
  
  names(df)[match_position[1]]
}


# Helper function: safely extract a column ----

get_column <- function(
    df,
    candidates,
    default = NA_character_,
    required = FALSE,
    label = NULL
) {
  
  column_name <- find_column(
    df = df,
    candidates = candidates,
    required = required,
    label = label
  )
  
  if (is.null(column_name)) {
    return(rep(default, nrow(df)))
  }
  
  as.character(df[[column_name]])
}


# Helper function: normalize DOI ----

normalize_doi <- function(x) {
  
  x <- stringr::str_trim(as.character(x))
  
  x <- stringr::str_to_lower(x)
  
  x <- stringr::str_replace(
    x,
    "^https?://(dx\\.)?doi\\.org/",
    ""
  )
  
  x <- stringr::str_replace(
    x,
    "^doi:\\s*",
    ""
  )
  
  x <- stringr::str_trim(x)
  
  x[x %in% c(
    "",
    "na",
    "n/a",
    "null"
  )] <- NA_character_
  
  x
}


# Helper function: normalize titles for matching ----

normalize_title <- function(x) {
  
  x <- as.character(x)
  
  x <- stringi::stri_trans_general(
    x,
    "Latin-ASCII"
  )
  
  x <- stringr::str_to_lower(x)
  
  x <- stringr::str_replace_all(
    x,
    "[^a-z0-9]+",
    " "
  )
  
  x <- stringr::str_squish(x)
  
  x[x == ""] <- NA_character_
  
  x
}


# Helper function: normalize years ----

normalize_year <- function(x) {
  
  x <- stringr::str_extract(
    as.character(x),
    "\\b(18|19|20|21)[0-9]{2}\\b"
  )
  
  suppressWarnings(
    as.integer(x)
  )
}


# Helper function: normalize OpenAlex work IDs ----

normalize_openalex_id <- function(x) {
  
  x <- as.character(x)
  
  x <- stringr::str_extract(
    x,
    "W[0-9]+"
  )
  
  x
}


# Helper function: create safe filename text ----

make_slug <- function(x, max_length = 60) {
  
  x <- as.character(x)
  
  x <- stringi::stri_trans_general(
    x,
    "Latin-ASCII"
  )
  
  x <- stringr::str_to_lower(x)
  
  x <- stringr::str_replace_all(
    x,
    "[^a-z0-9]+",
    "-"
  )
  
  x <- stringr::str_replace_all(
    x,
    "^-|-$",
    ""
  )
  
  x <- stringr::str_sub(
    x,
    1,
    max_length
  )
  
  x <- stringr::str_replace(
    x,
    "-$",
    ""
  )
  
  x
}


# Helper function: first author surname ----

first_author_surname <- function(author_string) {
  
  x <- as.character(author_string)
  
  x <- stringr::str_split(
    x,
    pattern = ";|\\||,| and ",
    simplify = TRUE
  )[, 1]
  
  x <- stringr::str_trim(x)
  
  # If format is "First Last", keep last word.
  words <- stringr::str_split(
    x,
    "\\s+",
    simplify = TRUE
  )
  
  result <- apply(
    words,
    1,
    function(z) {
      
      z <- z[z != ""]
      
      if (length(z) == 0) {
        return("unknown")
      }
      
      z[length(z)]
    }
  )
  
  result
}


# Read both Excel files ----

knowledge_raw <- readxl::read_excel(
  knowledge_base_file
)

openalex_raw <- readxl::read_excel(
  openalex_file
)

message(
  "Knowledge Base rows: ",
  nrow(knowledge_raw)
)

message(
  "OpenAlex rows: ",
  nrow(openalex_raw)
)


# Inspect original column names ----

message("\nKnowledge Base columns:")
print(names(knowledge_raw))

message("\nOpenAlex columns:")
print(names(openalex_raw))


# Standardize Knowledge Base fields ----

# The aliases below allow the script to tolerate small changes
# in column naming.

knowledge <- tibble::tibble(
  
  Author = get_column(
    knowledge_raw,
    c(
      "Author",
      "Authors",
      "author",
      "authors"
    )
  ),
  
  Title = get_column(
    knowledge_raw,
    c(
      "Title",
      "title",
      "paper_title",
      "publication_title"
    ),
    required = TRUE,
    label = "Knowledge Base title"
  ),
  
  Year = get_column(
    knowledge_raw,
    c(
      "Year",
      "year",
      "publication_year"
    )
  ),
  
  ID = get_column(
    knowledge_raw,
    c(
      "ID",
      "study_id",
      "folder_name",
      "folder"
    )
  ),
  
  ID_long_existing = get_column(
    knowledge_raw,
    c(
      "ID-long",
      "ID_long",
      "id_long",
      "long_id"
    )
  ),
  
  DOI = get_column(
    knowledge_raw,
    c(
      "DOI",
      "doi"
    )
  ),
  
  existing_status = get_column(
    knowledge_raw,
    c(
      "status",
      "Status",
      "download",
      "Download"
    )
  )
)


knowledge$source_row <- seq_len(nrow(knowledge))
knowledge$Type <- get_column(knowledge_raw, c("Type", "work_type", "publication_type"))

# Remove completely empty Knowledge Base rows ----

knowledge <- knowledge |>
  dplyr::filter(
    !(
      is.na(Title) |
        stringr::str_trim(Title) == ""
    )
  )


# Normalize Knowledge Base fields ----

knowledge <- knowledge |>
  dplyr::mutate(
    
    Year = normalize_year(Year),
    
    doi_key = normalize_doi(DOI),
    
    title_key = normalize_title(Title),
    
    title_year_key = dplyr::if_else(
      !is.na(title_key) & !is.na(Year),
      paste(title_key, Year, sep = "||"),
      NA_character_
    ),
    
    existing_status = stringr::str_to_lower(
      stringr::str_trim(existing_status)
    )
  )


# Map existing manual downloads ----

# IMPORTANT:
#
# Only done_MG means the paper has already been downloaded.
#
# done_OP is NOT treated as downloaded.
# to_do is NOT treated as downloaded.
#
# done_MG -> manual_GS

knowledge <- knowledge |>
  dplyr::mutate(
    
    Download = dplyr::case_when(
      
      existing_status == "done_mg" ~ "manual_GS",
      
      TRUE ~ NA_character_
    )
  )


# Standardize OpenAlex fields ----

openalex <- tibble::tibble(
  
  Author = get_column(
    openalex_raw,
    c(
      "Author",
      "Authors",
      "authors",
      "authorships",
      "author_display_name"
    )
  ),
  
  Title = get_column(
    openalex_raw,
    c(
      "Title",
      "title",
      "display_name",
      "work_title"
    ),
    required = TRUE,
    label = "OpenAlex title"
  ),
  
  Year = get_column(
    openalex_raw,
    c(
      "Year",
      "year",
      "publication_year"
    )
  ),
  
  DOI = get_column(
    openalex_raw,
    c(
      "DOI",
      "doi"
    )
  ),
  
  OpenAlex_ID = get_column(
    openalex_raw,
    c(
      "OpenAlex_ID",
      "openalex_id",
      "work_id",
      "openalex",
      "id"
    )
  ),
  
  open_access = get_column(
    openalex_raw,
    c(
      "open_access",
      "Open Access",
      "is_oa",
      "oa_status"
    )
  ),
  
  pdf_url_existing = get_column(
    openalex_raw,
    c(
      "pdf_url",
      "PDF URL",
      "pdf",
      "best_oa_location_pdf_url"
    )
  )
)


openalex$Publication_date <- get_column(openalex_raw, c("Date", "publication_date"))
openalex$OA_status <- get_column(openalex_raw, c("Open Access status", "oa_status"))
openalex$source_row <- seq_len(nrow(openalex))
openalex$Topic <- get_column(openalex_raw, c("Topic", "primary_topic", "primary_topic.display_name"))
openalex$Type <- get_column(openalex_raw, c("Type", "work_type", "publication_type"))
openalex <- dplyr::filter(openalex, !is.na(Title), stringr::str_trim(Title) != "")

# Normalize OpenAlex fields ----

openalex <- openalex |>
  dplyr::mutate(
    
    Year = normalize_year(Year),
    
    doi_key = normalize_doi(DOI),
    
    title_key = normalize_title(Title),
    
    title_year_key = dplyr::if_else(
      !is.na(title_key) & !is.na(Year),
      paste(title_key, Year, sep = "||"),
      NA_character_
    ),
    
    OpenAlex_ID = normalize_openalex_id(
      OpenAlex_ID
    )
  )


# Reconcile all source records BEFORE assigning study IDs. ----
# DOI is evidence, not an exclusive fallback key. Different-DOI title
# matches are retained for review; a matching title alone is not proof.
blank_to_na <- function(x) {
  x <- trimws(as.character(x))
  x[is.na(x) | x == ""] <- NA_character_
  x
}
first_present <- function(x) {
  x <- x[!is.na(x) & as.character(x) != ""]
  if (length(x)) x[[1]] else NA
}
collapse_values <- function(x) paste(sort(unique(x[!is.na(x) & x != ""])), collapse = " | ")
author_surnames <- function(x) {
  if (is.na(x) || !nzchar(x)) return(character())
  people <- strsplit(x, "[;|]| and ")[[1]]
  people <- vapply(people, function(person) {
    person <- trimws(person)
    if (grepl(",", person, fixed = TRUE)) person <- strsplit(person, ",", fixed = TRUE)[[1]][1]
    else person <- tail(strsplit(person, "[[:space:]]+")[[1]], 1)
    normalize_title(person)
  }, character(1))
  unique(people[!is.na(people) & nzchar(people)])
}
authors_agree <- function(a, b) {
  a <- author_surnames(a); b <- author_surnames(b)
  length(a) > 0 && length(b) > 0 &&
    length(intersect(a, b)) / max(length(a), length(b)) >= 0.8
}
material_kind <- function(title) {
  dplyr::case_when(
    stringr::str_detect(dplyr::coalesce(title, ""),
      "^(review for|reviewer [0-9]|author response|decision letter|evaluation summary|assessment)") ~ "peer_review",
    stringr::str_detect(dplyr::coalesce(title, ""),
      "^(data |data and |data r |code |supplement|appendi|site coordinates|traits from)|data deposit$") ~ "associated_material",
    TRUE ~ "publication"
  )
}
pdf_files <- function(id) {
  if (is.na(id) || !nzchar(id)) return(character())
  list.files(file.path(studies_dir, id), pattern = "[.]pdf$",
             recursive = TRUE, full.names = TRUE, ignore.case = TRUE)
}

knowledge$ID <- blank_to_na(knowledge$ID)
knowledge$source <- "KnowledgeBase"
openalex$source <- "OpenAlex"
records <- dplyr::bind_rows(knowledge, openalex)
records$record_id <- seq_len(nrow(records))
records$material_kind <- material_kind(records$title_key)
type_key <- normalize_title(records$Type)
records$material_kind[type_key %in% c("dataset", "software", "supplementary material", "supplementary materials", "paratext", "erratum")] <- "associated_material"
records$material_kind[type_key %in% c("peer review", "peer review report")] <- "peer_review"
records$doi_key <- blank_to_na(records$doi_key)
records$DOI <- records$doi_key

# Preserve IDs assigned by previous runs, including generated IDs, without
# treating previous output as another source publication.
previous_file <- if (file.exists(master_file)) master_file else file.path(root_dir, "reconciled_master.xlsx")
previous <- if (file.exists(previous_file)) readxl::read_excel(previous_file, sheet = "Master") else tibble::tibble()
prior_ids <- rep(list(character()), nrow(records))
if (nrow(previous)) {
  previous$doi_key <- normalize_doi(previous$DOI)
  previous$title_key <- normalize_title(previous$Title)
  for (i in seq_len(nrow(records))) {
    hit <- !is.na(records$doi_key[i]) & !is.na(previous$doi_key) & previous$doi_key == records$doi_key[i]
    # Missing-DOI identities require matching title, authors and year.
    fallback <- is.na(records$doi_key[i]) & is.na(previous$doi_key) &
      !is.na(previous$title_key) & previous$title_key == records$title_key[i]
    fallback[is.na(fallback)] <- FALSE
    for (j in which(fallback)) {
      fallback[j] <- isTRUE(previous$Year[j] == records$Year[i]) &&
        authors_agree(previous$Author[j], records$Author[i])
    }
    hit[is.na(hit)] <- FALSE
    prior_ids[[i]] <- unique(blank_to_na(previous$ID[hit | fallback]))
    prior_ids[[i]] <- stats::na.omit(prior_ids[[i]])
  }
}

parent <- seq_len(nrow(records))
find_root <- function(i) { while (parent[i] != i) i <- parent[i]; i }
join_records <- function(i, j) {
  a <- find_root(i); b <- find_root(j)
  if (a != b) parent[max(a, b)] <<- min(a, b)
}
edges <- list()
add_edge <- function(i, j, decision, reason) {
  edges[[length(edges) + 1L]] <<- tibble::tibble(
    record_a = i, record_b = j, decision = decision, reason = reason)
}
# Strong identity matches apply within and across both inputs.
for (key in c("doi_key", "OpenAlex_ID")) {
  groups <- split(seq_len(nrow(records)), records[[key]], drop = TRUE)
  for (ix in groups) if (length(ix) > 1) {
    for (j in ix[-1]) {
      join_records(ix[1], j)
      add_edge(ix[1], j, "merge", paste0("same_", key))
    }
  }
}

# Consider title candidates even if both records have identifiers.
# A component-level DOI check prevents a DOI-less bridge from joining
# two incompatible DOI-bearing publications through transitivity.
groups <- split(seq_len(nrow(records)), records$title_key, drop = TRUE)
for (ix in groups) if (length(ix) > 1) {
  pairs <- utils::combn(ix, 2)
  for (k in seq_len(ncol(pairs))) {
    i <- pairs[1, k]; j <- pairs[2, k]
    if (find_root(i) == find_root(j)) next
    roots <- vapply(seq_len(nrow(records)), find_root, integer(1))
    members <- which(roots %in% c(find_root(i), find_root(j)))
    dois <- unique(stats::na.omit(records$doi_key[members]))
    kinds_ok <- all(records$material_kind[members] == "publication")
    metadata_ok <- authors_agree(records$Author[i], records$Author[j]) &&
      !is.na(records$Year[i]) && !is.na(records$Year[j]) &&
      abs(records$Year[i] - records$Year[j]) <= 1
    if (length(dois) <= 1 && kinds_ok && metadata_ok) {
      join_records(i, j)
      add_edge(i, j, "merge", "same_title_compatible_authors_year_no_DOI_conflict")
    } else {
      reason <- if (!kinds_ok) "distinct_materials_or_versions_need_parent_link" else
        if (length(dois) > 1) "different_DOIs_verify_version_relationship" else
          "title_match_insufficient_author_or_year_evidence"
      add_edge(i, j, "review", reason)
    }
  }
}
records$group <- vapply(seq_len(nrow(records)), find_root, integer(1))
match_edges <- dplyr::bind_rows(edges)
if (!nrow(match_edges)) match_edges <- tibble::tibble(record_a=integer(), record_b=integer(), decision=character(), reason=character())
review_pairs <- match_edges |>
  dplyr::filter(decision == "review", records$group[record_a] != records$group[record_b])

# Reserve every old/folder ID, including empty reserved Knowledge Base rows.
reserved <- c(get_column(knowledge_raw, c("ID", "study_id", "folder_name", "folder")),
              if (nrow(previous)) previous$ID else character(),
              basename(list.dirs(studies_dir, recursive = FALSE)))
numbers <- suppressWarnings(as.integer(stringr::str_extract(reserved, "[0-9]+$")))
next_number <- max(c(0L, numbers), na.rm = TRUE) + 1L
assigned <- character()
records$canonical_ID <- NA_character_
id_map_rows <- list(); master_rows <- list()
for (ix in split(seq_len(nrow(records)), records$group)) {
  g <- records[ix, ]
  ids <- unique(c(stats::na.omit(g$ID), unlist(prior_ids[ix], use.names = FALSE)))
  ids <- ids[!is.na(ids) & nzchar(ids)]
  # Recover zero-padding typos only when the correctly padded folder exists.
  padded <- vapply(ids, function(id) {
    n <- suppressWarnings(as.integer(sub("^RDM-STUDY-", "", id)))
    if (is.na(n)) return(NA_character_)
    candidate <- sprintf("RDM-STUDY-%04d", n)
    if (dir.exists(file.path(studies_dir, candidate))) candidate else NA_character_
  }, character(1))
  ids <- unique(c(ids, stats::na.omit(padded)))
  files_by_id <- lapply(ids, pdf_files)
  counts <- lengths(files_by_id)
  kb_ids <- stats::na.omit(g$ID[g$source == "KnowledgeBase"])
  if (length(ids)) {
    order_ids <- order(-counts, !(ids %in% kb_ids), ids)
    ids <- ids[order_ids]
    canonical <- ids[1]
    if (any(ids %in% assigned)) stop("One legacy ID maps to conflicting publications. Resolve before exporting: ", paste(ids, collapse=", "))
  } else {
    canonical <- sprintf("RDM-STUDY-%04d", next_number)
    next_number <- next_number + 1L
  }
  assigned <- c(assigned, ids, canonical)
  files <- unique(unlist(lapply(ids, pdf_files), use.names = FALSE))
  has_manual_gs <- any(g$existing_status == "done_mg", na.rm = TRUE) && length(files) > 0
  order_rows <- order(!(g$ID %in% canonical), g$source != "KnowledgeBase", is.na(g$doi_key), g$record_id)
  g <- g[order_rows, ]
  has_oa <- any(g$source == "OpenAlex")
  row <- tibble::tibble(
    Author = as.character(first_present(g$Author)), Title = as.character(first_present(g$Title)),
    Year = as.integer(first_present(g$Year)), ID = canonical,
    ID_long_existing = as.character(first_present(g$ID_long_existing[g$ID %in% canonical])),
    Download = if (has_manual_gs) "manual_GS" else if (length(files)) "manual_OA" else NA_character_,
    DOI = as.character(first_present(g$doi_key)),
    open_access = as.character(first_present(g$open_access)),
    source_query = if (has_oa) "OpenAlex" else "Google Scholar, not OpenAlex",
    OpenAlex_ID = as.character(first_present(g$OpenAlex_ID)),
    pdf_url_existing = as.character(first_present(g$pdf_url_existing)),
    match_method = if (nrow(g) > 1) "global_publication_group" else "single_source_record",
    all_DOIs = collapse_values(g$doi_key), all_OpenAlex_IDs = collapse_values(g$OpenAlex_ID),
    legacy_IDs = collapse_values(ids), existing_pdf_count = length(files),
    existing_files = collapse_values(files), source_record_count = nrow(g),
    has_knowledge_base = any(g$source == "KnowledgeBase"), material_kind = first_present(g$material_kind),
    all_types = collapse_values(g$Type),
    Topic = as.character(first_present(collapse_values(g$Topic))),
    review_flag = if (any(ix %in% c(review_pairs$record_a, review_pairs$record_b))) "CHECK_MATCH_CANDIDATES" else "OK"
  )
  master_rows[[length(master_rows)+1L]] <- row
  id_map_rows[[length(id_map_rows)+1L]] <- tibble::tibble(
    legacy_ID = if (length(ids)) ids else canonical, canonical_ID = canonical,
    action = ifelse((if (length(ids)) ids else canonical) == canonical, "keep", "alias_no_files_moved"))
  records$canonical_ID[ix] <- canonical
}
master <- dplyr::bind_rows(master_rows)
id_mapping <- dplyr::bind_rows(id_map_rows)
existing_master <- dplyr::filter(master, has_knowledge_base)
new_openalex <- dplyr::filter(master, !has_knowledge_base)
review_candidates <- review_pairs |>
  dplyr::mutate(ID_a = records$canonical_ID[record_a], ID_b = records$canonical_ID[record_b],
                Title_a = records$Title[record_a], Title_b = records$Title[record_b],
                DOI_a = records$DOI[record_a], DOI_b = records$DOI[record_b]) |>
  dplyr::distinct(ID_a, ID_b, reason, .keep_all = TRUE)
source_records <- records |>
  dplyr::select(canonical_ID, source, source_row, record_id, Author, Title, Year, DOI, OpenAlex_ID, ID, Type, Topic, Publication_date, OA_status, material_kind)
stopifnot(nrow(source_records) == nrow(knowledge) + nrow(openalex),
          !anyDuplicated(master$ID), !anyDuplicated(id_mapping$legacy_ID))
doi_check <- source_records |>
  dplyr::filter(!is.na(DOI)) |>
  dplyr::distinct(DOI, canonical_ID) |>
  dplyr::count(DOI) |>
  dplyr::filter(n > 1)
stopifnot(nrow(doi_check) == 0)


# Create ID-long ----

# Existing ID-long values are preserved.
#
# Missing ID-long values are generated using:
#
# RDM-STUDY-XXXX_author_year_short-title

generated_id_long <- paste(
  
  master$ID,
  
  make_slug(
    first_author_surname(
      master$Author
    ),
    max_length = 25
  ),
  
  ifelse(
    is.na(master$Year),
    "year-unknown",
    master$Year
  ),
  
  make_slug(
    master$Title,
    max_length = 60
  ),
  
  sep = "_"
)

master <- master |>
  dplyr::mutate(
    
    `ID-long` = dplyr::if_else(
      
      !is.na(ID_long_existing) &
        stringr::str_trim(ID_long_existing) != "",
      
      ID_long_existing,
      
      generated_id_long
    )
  )


# Normalize final DOI values ----

# Keep DOI in a consistent canonical form.
#
# Example:
# 10.1111/example.12345
#
# rather than:
# https://doi.org/10.1111/example.12345

master <- master |>
  dplyr::mutate(
    DOI = normalize_doi(DOI)
  )


# Verify unique study IDs ----

duplicate_ids <- master |>
  dplyr::filter(
    !is.na(ID)
  ) |>
  dplyr::count(ID) |>
  dplyr::filter(n > 1)

if (nrow(duplicate_ids) > 0) {
  
  print(duplicate_ids)
  
  stop(
    "Duplicate RDM-STUDY IDs detected. Downloads have been stopped."
  )
}


# Verify remaining duplicate DOIs ----

duplicate_dois <- master |>
  dplyr::filter(
    !is.na(DOI)
  ) |>
  dplyr::count(DOI) |>
  dplyr::filter(n > 1)

if (nrow(duplicate_dois) > 0) {
  
  message(
    "\nWARNING: duplicate DOIs remain in the master table:"
  )
  
  print(duplicate_dois)
}


# Verify remaining duplicate normalized titles ----

title_check <- master |>
  dplyr::mutate(
    title_check = normalize_title(Title)
  ) |>
  dplyr::filter(
    !is.na(title_check)
  ) |>
  dplyr::count(title_check) |>
  dplyr::filter(n > 1)

if (nrow(title_check) > 0) {
  
  message(
    "\nWARNING: duplicate normalized titles remain."
  )
  
  print(title_check)
}


# Create study folders ----

# Each paper receives a folder corresponding exactly to ID.
#
# Example:
#
# studies/RDM-STUDY-0185/

# Folders are created only when an actual download is attempted.

# Detect previously automated downloads ----

# This makes the script resumable.
#
# If an automated PDF already exists at:
#
# studies/RDM-STUDY-XXXX/RDM-STUDY-XXXX.pdf
#
# it is not downloaded again.

expected_pdf <- file.path(
  studies_dir,
  master$ID,
  paste0(
    master$ID,
    ".pdf"
  )
)

master <- master |>
  dplyr::mutate(
    
    automated_file_exists = file.exists(
      expected_pdf
    ),
    
    Download = dplyr::case_when(
      
      Download == "manual_GS" ~ "manual_GS",
      
      automated_file_exists ~ "automated",
      
      TRUE ~ Download
    )
  )


# Create reconciled master table ----

# User-facing final columns:
#
# Author
# Title
# Year
# ID
# ID-long
# Download
# DOI
# open_access
# source_query
#
# Additional audit columns are placed on a second worksheet.

final_table <- master |>
  dplyr::select(
    
    Author,
    Title,
    Year,
    ID,
    `ID-long`,
    Download,
    DOI,
    open_access,
    source_query,
    Topic
  ) |>
  dplyr::arrange(ID)


# Create audit table ----

audit_table <- master |>
  dplyr::select(ID, Title, Topic, DOI, OpenAlex_ID, match_method, review_flag,
                all_DOIs, all_OpenAlex_IDs, legacy_IDs, source_record_count,
                existing_pdf_count, existing_files, material_kind,
                pdf_url_existing, automated_file_exists) |>
  dplyr::arrange(ID)

add_reconciliation_sheets <- function(workbook) {
  sheets <- list(ID_mapping = id_mapping, Source_records = source_records,
                 Match_candidates = review_candidates, Match_decisions = match_edges)
  for (name in names(sheets)) {
    openxlsx::addWorksheet(workbook, name)
    openxlsx::writeData(workbook, name, sheets[[name]])
    openxlsx::freezePane(workbook, name, firstRow = TRUE)
  }
}

# Create summary table ----

summary_table <- tibble::tibble(
  
  measure = c(
    "Total studies",
    "Existing Knowledge Base studies",
    "New OpenAlex studies",
    "Google Scholar only",
    "OpenAlex source",
    "Already downloaded manually from Google Scholar",
    "Automated PDFs already present",
    "Pending"
  ),
  
  value = c(
    
    nrow(master),
    
    nrow(existing_master),
    
    nrow(new_openalex),
    
    sum(
      master$source_query ==
        "Google Scholar, not OpenAlex",
      na.rm = TRUE
    ),
    
    sum(
      master$source_query ==
        "OpenAlex",
      na.rm = TRUE
    ),
    
    sum(
      master$Download ==
        "manual_GS",
      na.rm = TRUE
    ),
    
    sum(
      master$Download ==
        "automated",
      na.rm = TRUE
    ),
    
    sum(
      is.na(master$Download) |
        master$Download == "",
      na.rm = TRUE
    )
  )
)


# Read-only validation mode: expose tables to a test harness, then stop
# before any Excel export, folder creation, or download.
if (identical(Sys.getenv("RDM_VALIDATE_ONLY"), "1")) {
  message("Validation complete: ", nrow(master), " publication groups; ",
          nrow(review_candidates), " unresolved match candidates.")
  stop(structure(list(message = "RDM_VALIDATION_COMPLETE", call = NULL),
                 class = c("rdm_validation_complete", "error", "condition")))
}

# Export reconciled master table before downloading ----

workbook <- openxlsx::createWorkbook()
  add_reconciliation_sheets(workbook)

openxlsx::addWorksheet(
  workbook,
  "Master"
)

openxlsx::writeData(
  workbook,
  "Master",
  final_table
)

openxlsx::addWorksheet(
  workbook,
  "Audit"
)

openxlsx::writeData(
  workbook,
  "Audit",
  audit_table
)

openxlsx::addWorksheet(
  workbook,
  "Summary"
)

openxlsx::writeData(
  workbook,
  "Summary",
  summary_table
)


# Format Excel output ----

header_style <- openxlsx::createStyle(
  textDecoration = "bold",
  halign = "center",
  valign = "center"
)

openxlsx::addStyle(
  workbook,
  "Master",
  header_style,
  rows = 1,
  cols = seq_len(ncol(final_table)),
  gridExpand = TRUE
)

openxlsx::freezePane(
  workbook,
  "Master",
  firstRow = TRUE
)

openxlsx::addFilter(
  workbook,
  "Master",
  rows = 1,
  cols = seq_len(ncol(final_table))
)

openxlsx::setColWidths(
  workbook,
  "Master",
  cols = 1:ncol(final_table),
  widths = "auto"
)

# Prevent extremely wide title columns.
openxlsx::setColWidths(
  workbook,
  "Master",
  cols = which(
    names(final_table) == "Title"
  ),
  widths = 60
)

openxlsx::setColWidths(
  workbook,
  "Master",
  cols = which(
    names(final_table) == "ID-long"
  ),
  widths = 60
)

if (file.exists(master_file)) {
  backup <- paste0(master_file, ".", format(Sys.time(), "%Y%m%d-%H%M%S"), ".bak")
  if (!file.copy(master_file, backup, overwrite = FALSE)) stop("Could not back up existing corrected workbook")
}
openxlsx::saveWorkbook(
  workbook,
  master_file,
  overwrite = TRUE
)

message(
  "\nReconciled master written to:\n",
  master_file
)


# Print reconciliation summary ----

print(summary_table)

message(
  "\nPlease inspect reconciled_master_corrected.xlsx and Match_candidates before enabling downloads."
)


# Stop here if downloads are disabled ----

if (!DOWNLOAD_PDFS) {
  
  message(
    paste0(
      "\nDOWNLOAD_PDFS is FALSE.\n",
      "No PDFs were downloaded.\n",
      "If the reconciliation looks correct, change:\n\n",
      "DOWNLOAD_PDFS <- TRUE\n\n",
      "and run the script again."
    )
  )
  
} else {
  
  
  if (nrow(review_candidates) > 0) stop("Resolve Match_candidates before enabling downloads.")
  if (any(master$source_query == "OpenAlex" & is.na(master$OpenAlex_ID))) {
    stop("OpenAlex Work IDs are missing. Resolve identifiers before enabling downloads.")
  }

  # Select papers eligible for automated download ----
  
  download_candidates <- master |>
    dplyr::filter(
      
      # Do not overwrite Google Scholar downloads.
      is.na(Download),
      
      # Only OpenAlex-query papers can use the
      # OpenAlex content API.
      source_query == "OpenAlex",
      
      # We need an OpenAlex work ID.
      !is.na(OpenAlex_ID),
      
      # Skip files already downloaded.
      !automated_file_exists
    )
  
  
  # Apply optional download limit ----
  
  if (is.finite(MAX_DOWNLOADS)) {
    
    download_candidates <- download_candidates |>
      dplyr::slice_head(
        n = MAX_DOWNLOADS
      )
  }
  
  message(
    "\nPapers selected for this download run: ",
    nrow(download_candidates)
  )
  
  
  # Function to download one OpenAlex PDF ----
  
  download_openalex_pdf <- function(
    study_id,
    openalex_id
  ) {
    
    destination <- file.path(
      studies_dir,
      study_id,
      paste0(
        study_id,
        ".pdf"
      )
    )
    
    # Do not download again if already present.
    if (file.exists(destination)) {
      
      return(
        tibble::tibble(
          ID = study_id,
          OpenAlex_ID = openalex_id,
          result = "already_exists",
          http_status = NA_integer_,
          file = destination
        )
      )
    }
    
    url <- paste0(
      "https://content.openalex.org/works/",
      openalex_id,
      ".pdf"
    )
    
    request <- httr2::request(url) |>
      httr2::req_url_query(
        api_key = openalex_api_key
      ) |>
      httr2::req_user_agent(
        "RDM-STUDY-0023 literature downloader"
      ) |>
      httr2::req_error(
        is_error = function(resp) FALSE
      ) |>
      httr2::req_timeout(120)
    
    response <- tryCatch(
      
      httr2::req_perform(request),
      
      error = function(e) e
    )
    
    
    if (inherits(response, "error")) {
      
      return(
        tibble::tibble(
          ID = study_id,
          OpenAlex_ID = openalex_id,
          result = "request_error",
          http_status = NA_integer_,
          file = NA_character_
        )
      )
    }
    
    
    status <- httr2::resp_status(
      response
    )
    
    
    if (status != 200) {
      
      return(
        tibble::tibble(
          ID = study_id,
          OpenAlex_ID = openalex_id,
          result = paste0(
            "http_",
            status
          ),
          http_status = status,
          file = NA_character_
        )
      )
    }
    
    
    body <- httr2::resp_body_raw(
      response
    )
    
    
    # Check for PDF signature:
    # PDF files normally begin with "%PDF".
    pdf_signature <- rawToChar(
      body[1:min(4, length(body))]
    )
    
    
    if (!startsWith(
      pdf_signature,
      "%PDF"
    )) {
      
      return(
        tibble::tibble(
          ID = study_id,
          OpenAlex_ID = openalex_id,
          result = "response_not_pdf",
          http_status = status,
          file = NA_character_
        )
      )
    }
    
    
    dir.create(file.path(studies_dir, study_id), recursive = TRUE, showWarnings = FALSE)
    writeBin(
      body,
      destination
    )
    
    
    tibble::tibble(
      ID = study_id,
      OpenAlex_ID = openalex_id,
      result = "downloaded",
      http_status = status,
      file = destination
    )
  }
  
  
  # Download available OpenAlex PDFs ----
  
  download_results <- purrr::map2_dfr(
    
    download_candidates$ID,
    
    download_candidates$OpenAlex_ID,
    
    function(study_id, openalex_id) {
      
      message(
        "Trying ",
        study_id,
        " - ",
        openalex_id
      )
      
      result <- download_openalex_pdf(
        study_id = study_id,
        openalex_id = openalex_id
      )
      
      # Brief pause avoids rapid-fire requests.
      Sys.sleep(0.2)
      
      result
    }
  )
  
  
  # Save download log ----
  
  if (file.exists(download_log_file)) {
    
    old_log <- read.csv(
      download_log_file,
      stringsAsFactors = FALSE
    )
    
    combined_log <- dplyr::bind_rows(
      old_log,
      download_results
    )
    
  } else {
    
    combined_log <- download_results
  }
  
  
  write.csv(
    combined_log,
    download_log_file,
    row.names = FALSE
  )
  
  
  # Update Download status after successful downloads ----
  
  downloaded_ids <- download_results |>
    dplyr::filter(
      result %in% c(
        "downloaded",
        "already_exists"
      )
    ) |>
    dplyr::pull(ID)
  
  
  master <- master |>
    dplyr::mutate(
      
      Download = dplyr::case_when(
        
        Download == "manual_GS" ~
          "manual_GS",
        
        ID %in% downloaded_ids ~
          "automated",
        
        TRUE ~
          Download
      )
    )
  
  
  # Recreate final table after downloads ----
  
  final_table <- master |>
    dplyr::select(
      
      Author,
      Title,
      Year,
      ID,
      `ID-long`,
      Download,
      DOI,
      open_access,
      source_query,
    Topic
      
    ) |>
    dplyr::arrange(ID)
  
  
  # Identify OpenAlex papers requiring manual retrieval ----
  
  manual_oa_candidates <- download_results |>
    dplyr::filter(
      !result %in% c(
        "downloaded",
        "already_exists"
      )
    ) |>
    dplyr::left_join(
      master |>
        dplyr::select(
          ID,
          Author,
          Title,
          Year,
          DOI,
          OpenAlex_ID,
          open_access
        ),
      by = c(
        "ID",
        "OpenAlex_ID"
      )
    )
  
  
  # Export updated workbook ----
  
  workbook <- openxlsx::createWorkbook()
  add_reconciliation_sheets(workbook)
  
  
  openxlsx::addWorksheet(
    workbook,
    "Master"
  )
  
  openxlsx::writeData(
    workbook,
    "Master",
    final_table
  )
  
  
  openxlsx::addWorksheet(
    workbook,
    "Manual_OA_needed"
  )
  
  openxlsx::writeData(
    workbook,
    "Manual_OA_needed",
    manual_oa_candidates
  )
  
  
  openxlsx::addWorksheet(
    workbook,
    "Download_log"
  )
  
  openxlsx::writeData(
    workbook,
    "Download_log",
    download_results
  )
  
  
  openxlsx::addWorksheet(
    workbook,
    "Audit"
  )
  
  openxlsx::writeData(
    workbook,
    "Audit",
    audit_table
  )
  
  
  openxlsx::freezePane(
    workbook,
    "Master",
    firstRow = TRUE
  )
  
  openxlsx::addFilter(
    workbook,
    "Master",
    rows = 1,
    cols = 1:ncol(final_table)
  )
  
  openxlsx::setColWidths(
    workbook,
    "Master",
    cols = 1:ncol(final_table),
    widths = "auto"
  )
  
  openxlsx::setColWidths(
    workbook,
    "Master",
    cols = which(
      names(final_table) == "Title"
    ),
    widths = 60
  )
  
  openxlsx::setColWidths(
    workbook,
    "Master",
    cols = which(
      names(final_table) == "ID-long"
    ),
    widths = 60
  )
  
  
  openxlsx::saveWorkbook(
    workbook,
    master_file,
    overwrite = TRUE
  )
  
  
  # Report download results ----
  
  message(
    "\nDownload results:"
  )
  
  print(
    download_results |>
      dplyr::count(
        result,
        sort = TRUE
      )
  )
  
  message(
    "\nUpdated master table written to:\n",
    master_file
  )
  
  message(
    "\nDownload log written to:\n",
    download_log_file
  )
}


# End ----