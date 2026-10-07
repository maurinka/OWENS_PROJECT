# Download a bounded batch from the reviewed workbook without rebuilding IDs.
library(readxl)
library(httr2)
root <- Sys.getenv("RDM_ROOT_DIR", "/Users/maura/Desktop/KnowldegeBase")
key <- Sys.getenv("OPENALEX_API_KEY")
if (!nzchar(key)) stop("OPENALEX_API_KEY is not configured")
book <- file.path(root, "reconciled_master_corrected.xlsx")
m <- read_excel(book, "Master")
a <- read_excel(book, "Audit")
d <- merge(m, a[,c("ID","OpenAlex_ID","review_flag","material_kind","existing_pdf_count")], by="ID")
d <- subset(d, is.na(Download) & review_flag=="OK" & material_kind=="publication" & !is.na(OpenAlex_ID) & existing_pdf_count==0)
limit <- as.integer(Sys.getenv("RDM_BATCH_SIZE", "5"))
stopifnot(!is.na(limit), limit > 0)
d <- head(d, limit)
logfile <- file.path(root,"download_batch_log.csv")
results <- list()
for (i in seq_len(nrow(d))) {
  id <- d$ID[i]; work <- d$OpenAlex_ID[i]
  folder <- file.path(root,id)
  destination <- file.path(folder,paste0(id,"_",work,"_openalex_main.pdf"))
  result <- "request_error"; status <- NA_integer_; bytes <- 0L
  # Any existing PDF means this first batch should leave the folder alone.
  existing <- list.files(folder,pattern="[.]pdf$",recursive=TRUE,ignore.case=TRUE)
  if (length(existing)) {
    result <- "skipped_existing_pdf"
  } else {
    response <- tryCatch(request(paste0("https://content.openalex.org/works/",work,".pdf")) |>
      req_url_query(api_key=key) |>
      req_error(is_error=function(resp) FALSE) |>
      req_timeout(60) |> req_perform(), error=function(e) NULL)
    if (!is.null(response)) {
      status <- resp_status(response)
      if (status==200L) {
        body <- resp_body_raw(response)
        if (length(body)>=5 && identical(body[1:5],charToRaw("%PDF-"))) {
          dir.create(folder,recursive=TRUE,showWarnings=FALSE)
          part <- tempfile("download-",tmpdir=folder,fileext=".part")
          writeBin(body,part)
          if (file.exists(destination) || !file.rename(part,destination)) stop("Could not safely finalize PDF")
          bytes <- length(body); result <- "downloaded"
        } else result <- "response_not_pdf"
      } else result <- paste0("http_",status)
    }
  }
  row <- data.frame(ID=id,OpenAlex_ID=work,result=result,http_status=status,
    bytes=bytes,file=if(result=="downloaded") destination else NA_character_,
    timestamp=as.character(Sys.time()))
  results[[length(results)+1L]] <- row
  write.table(row,logfile,sep=",",row.names=FALSE,col.names=!file.exists(logfile),
    append=file.exists(logfile),qmethod="double")
  message(id,": ",result)
  if (!is.na(status) && status %in% c(401L,402L,403L,429L)) break
  Sys.sleep(0.3)
}
if (length(results)) {
  results <- do.call(rbind,results)
  write.csv(results,file.path(root,"download_batch_latest.csv"),row.names=FALSE)
  print(results[,c("ID","result","http_status","bytes")],row.names=FALSE)
} else message("No eligible papers in this batch")
