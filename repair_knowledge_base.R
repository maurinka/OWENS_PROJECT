# Input directory containing the three supplied files. This edits table data only;
# it does not rewrite Excel formatting or the other worksheets.
library(dplyr)
library(readr)
library(readxl)
library(stringr)
library(stringi)
library(stringdist)

data_dir <- '.'
out_dir <- if (length(commandArgs(TRUE)) > 1) commandArgs(TRUE)[2] else "."
dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

# Undo UTF-8 bytes that Excel interpreted as MacRoman. Leave already valid
# accented text alone when the reverse conversion is not valid UTF-8.
repair <- function(x) coalesce(iconv(iconv(x, "UTF-8", "macintosh"), "UTF-8", "UTF-8"), x)
text <- function(x) na_if(str_squish(str_replace_all(str_replace_all(x, fixed("\\n"), " "), "\\p{Cf}", "")), "")
key <- function(x) na_if(str_squish(str_replace_all(stri_trans_general(str_to_lower(x), "Latin-ASCII"), "[^a-z0-9]+", " ")), "")
doi <- function(x) {
  x <- str_remove(str_to_lower(str_trim(x)), "^(https?://(dx\\.)?doi\\.org/|doi:\\s*)")
  na_if(x, "")
}
prepare <- function(d) d |> mutate(across(c(Title, Author), text), doi_key = doi(DOI), title_key = key(Title), author_key = key(Author))

orig <- read_excel(file.path(data_dir, "KnowldegeBase.xlsx"), sheet = "02_10_26")
csv <- read_csv(file.path(data_dir, "orig_csv.csv"), locale = locale(encoding = "UTF-8"), show_col_types = FALSE) |> prepare()
works <- read_csv(file.path(data_dir, "works.csv"), locale = locale(encoding = "UTF-8"), show_col_types = FALSE) |> prepare()
kb <- orig |> mutate(across(c(Author, Title), repair)) |> prepare()

# The CSV's one repeated DOI differs only by a literal \n in the title.
# Remove equivalent reference rows, then reject any remaining DOI ambiguity.
csv <- csv |> distinct(doi_key, title_key, author_key, Year, .keep_all = TRUE)
works <- works |> distinct(doi_key, title_key, author_key, Year, .keep_all = TRUE)
stopifnot(!anyDuplicated(na.omit(csv$doi_key)), !anyDuplicated(na.omit(works$doi_key)))

# Exact DOI first. Otherwise exact normalized title + year, using authors
# to resolve multiple candidates. Never automatically merge different DOIs.
# Missing years only match when both are missing AND authors agree.
match_rows <- function(a, b) vapply(seq_len(nrow(a)), function(i) {
  j <- which(!is.na(a$doi_key[i]) & b$doi_key == a$doi_key[i])
  if (!length(j) && !is.na(a$title_key[i])) {
    same_year <- !is.na(a$Year[i]) & !is.na(b$Year) & b$Year == a$Year[i]
    both_missing <- is.na(a$Year[i]) & is.na(b$Year) & !is.na(a$author_key[i]) & b$author_key == a$author_key[i]
    j <- which(b$title_key == a$title_key[i] & (same_year | both_missing) &
                 (is.na(a$doi_key[i]) | is.na(b$doi_key) | b$doi_key == a$doi_key[i]))
    same_author <- j[which(!is.na(a$author_key[i]) & b$author_key[j] == a$author_key[i])]
    if (length(same_author)) j <- same_author
  }
  if (length(j) == 1L) j else NA_integer_
}, integer(1))

update_rows <- function(a, b, j) {
  a$Author <- coalesce(b$Author[j], a$Author)
  a$Title <- coalesce(b$Title[j], a$Title)
  # Keep the existing DOI; fill it when a safe fallback match supplies one.
  a$DOI <- coalesce(a$doi_key, b$doi_key[j])
  if (!"Type" %in% names(a)) a$Type <- NA_character_
  a$Type <- coalesce(a$Type, b$Type[j])
  prepare(a)
}
j_csv <- match_rows(kb, csv)
kb <- update_rows(kb, csv, j_csv)
j_works <- match_rows(kb, works)
updated <- update_rows(kb, works, j_works)
stopifnot(nrow(updated) == nrow(orig))

# Inspect unused works against ALL existing rows, including rows matched by DOI.
# Same-title alternate DOIs, incomplete records and near matches go to review.
unused <- which(!seq_len(nrow(works)) %in% na.omit(j_works))
review_pairs <- bind_rows(lapply(unused, function(j) {
  if (is.na(works$title_key[j])) return(tibble())
  same_title <- !is.na(updated$title_key) & updated$title_key == works$title_key[j]
  same_year <- !is.na(works$Year[j]) & !is.na(updated$Year) & updated$Year == works$Year[j]
  dist <- stringdist(works$title_key[j], updated$title_key, method = "lv") /
    pmax(nchar(works$title_key[j]), nchar(updated$title_key))
  author_dist <- stringdist(works$author_key[j], updated$author_key, method = "lv") /
    pmax(nchar(works$author_key[j]), nchar(updated$author_key))
  i <- which(same_title | (same_year & ((!is.na(dist) & dist <= .25) |
                  (!is.na(author_dist) & author_dist <= .10))))
  if (!length(i)) return(tibble())
  tibble(works_row = j, excel_row = i + 1L, Title_works = works$Title[j],
         Author_works = works$Author[j], DOI_works = works$DOI[j],
         Title_orig = updated$Title[i], Author_orig = updated$Author[i],
         DOI_orig = updated$DOI[i], title_distance = dist[i])
}))
hold <- union(review_pairs$works_row, which(is.na(works$Title) | is.na(works$Year)))
append_rows <- setdiff(unused, hold)
new_rows <- works[append_rows, ] |> mutate(Download = "to_do")

finish <- function(d) {
  # Fix only these observed publisher-route artefacts; preserve all DOI suffixes,
  # including version, dataset and review identifiers.
  d |> mutate(DOI = str_replace(doi(DOI), "^10\\.1108/ijm-06-2021-0353(\\.html|/full/html)\\?utm_.*$", "10.1108/ijm-06-2021-0353"),
              DOI = if_else(is.na(DOI), NA_character_, paste0("https://doi.org/", DOI))) |>
    select(-any_of(c("doi_key", "title_key", "author_key")))
}
full <- bind_rows(updated, new_rows) |> finish()
stopifnot(nrow(full) == nrow(orig) + length(append_rows))
write_csv(full, file.path(out_dir, "KnowledgeBase_clean.csv"))
write_csv(review_pairs, file.path(out_dir, "match_review.csv"))
write_csv(works[intersect(unused, hold), ] |> mutate(works_row = intersect(unused, hold)) |> finish(), file.path(out_dir, "works_held_for_review.csv"))
write_csv(tibble(excel_row = seq_len(nrow(orig)) + 1L, csv_row = j_csv, works_row = j_works), file.path(out_dir, "match_audit.csv"))
cat(sprintf("Original: %d; CSV matches: %d; works matches: %d; appended: %d; held: %d; final: %d\n",
            nrow(orig), sum(!is.na(j_csv)), sum(!is.na(j_works)), length(append_rows), length(intersect(unused, hold)), nrow(full)))
