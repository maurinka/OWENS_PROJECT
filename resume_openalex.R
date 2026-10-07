# Resume OpenAlex cached-PDF retrieval. NEVER writes the source workbook.
# Run in OWENS_PROJECT: source("resume_openalex.R")
# Uses the locally configured OPENALEX_API_KEY; never logs the key.
for(p in c("readxl","httr2","jsonlite","digest")) if(!requireNamespace(p,quietly=TRUE)) stop("Missing package: ",p)
root<-Sys.getenv("RDM_ROOT_DIR","/Users/maura/Desktop/KnowldegeBase")
book<-file.path(root,"reconciled_master_corrected.xlsx")
key<-Sys.getenv("OPENALEX_API_KEY");stopifnot(nzchar(key))
run<-format(Sys.time(),"%Y%m%d-%H%M%S")
state<-file.path(root,"download_records",paste0("openalex-",run));dir.create(state,recursive=TRUE)
before<-digest::digest(file=book,algo="sha256")
writeLines(before,file.path(state,"workbook-sha256-before.txt"))
m<-as.data.frame(readxl::read_excel(book,"Master"));s<-as.data.frame(readxl::read_excel(book,"Source_records"));mapping<-as.data.frame(readxl::read_excel(book,"ID_mapping"))
stopifnot(!anyDuplicated(m$ID))
pdfs<-function(id) {ids<-unique(c(id,mapping$legacy_ID[mapping$canonical_ID==id]));unlist(lapply(ids,function(x) list.files(file.path(root,x),pattern="[.]pdf$",recursive=TRUE,full.names=TRUE,ignore.case=TRUE)))}
# Recognize user annotations conservatively. Retry missing/no-access records,
# but preserve every manual or other unrecognized nonblank status.
allowed<-c("","to_do","no_access","not_there")
status<-ifelse(is.na(m$Download),"",tolower(trimws(m$Download)))
ids<-m$ID[status %in% allowed & !vapply(m$ID,function(id) length(pdfs(id))>0,logical(1))]
write.csv(m[m$ID %in% ids,],file.path(state,"selected_studies.csv"),row.names=FALSE)
logfile<-file.path(state,"results.csv");blocked<-FALSE;saved<-0L
log<-function(id,work,result,http=NA_integer_,file="",sha="") {
 row<-data.frame(ID=id,OpenAlex_ID=work,result=result,http_status=http,file=file,sha256=sha,time=as.character(Sys.time()))
 write.table(row,logfile,sep=",",row.names=FALSE,col.names=!file.exists(logfile),append=file.exists(logfile),qmethod="double")
}
fetch<-function(url) {
 Sys.sleep(.15)
 r<-tryCatch(httr2::req_perform(httr2::req_timeout(httr2::req_error(httr2::req_url_query(httr2::request(url),api_key=key),is_error=function(r) FALSE),60)),error=function(e) NULL)
 if(is.null(r))return(list(status=NA_integer_,body=NULL))
 code<-httr2::resp_status(r)
 if(code %in% c(401L,402L,403L,429L))blocked<<-TRUE
 list(status=code,body=if(code==200L)httr2::resp_body_raw(r) else NULL)
}
pdfinfo<-Sys.getenv("RDM_PDFINFO",Sys.which("pdfinfo"))
if(!nzchar(pdfinfo))pdfinfo<-path.expand("~/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdfinfo")
stopifnot(file.exists(pdfinfo))
message("Missing studies selected: ",length(ids),". Workbook is read-only.")
for(id in ids) {
 if(blocked)break
 if(length(pdfs(id)))next
 rows<-s[s$canonical_ID==id & !is.na(s$canonical_ID),]
 works<-unique(na.omit(rows$OpenAlex_ID));works<-sub("^https://openalex.org/","",works);works<-works[grepl("^W[0-9]+$",works)]
 if(!length(works)) {
  dois<-unique(na.omit(rows$DOI));dois<-dois[grepl("^10[.]",dois)]
  if(length(dois)) {
   r<-fetch(paste0("https://api.openalex.org/works/https://doi.org/",dois[1]))
   if(!is.null(r$body)) {meta<-jsonlite::fromJSON(rawToChar(r$body),simplifyVector=FALSE);works<-sub("https://openalex.org/","",meta$id)}
   else log(id,"","metadata_unavailable",r$status)
  }
 }
 if(!length(works)){log(id,"","no_work_id");next}
 for(work in works) {
  if(blocked || length(pdfs(id)))break
  cache<-file.path(root,"download_records",paste0(digest::digest(work,algo="sha256",serialize=FALSE),".json"))
  meta<-if(file.exists(cache))tryCatch(jsonlite::fromJSON(cache,simplifyVector=FALSE),error=function(e)NULL) else NULL
  if(is.null(meta)) {
   r<-fetch(paste0("https://api.openalex.org/works/",work))
   if(is.null(r$body)){log(id,work,"metadata_unavailable",r$status);next}
   meta<-jsonlite::fromJSON(rawToChar(r$body),simplifyVector=FALSE)
   writeBin(r$body,cache)
  }
  if(!isTRUE(meta$has_content$pdf)){log(id,work,"no_cached_pdf");next}
  r<-fetch(paste0("https://content.openalex.org/works/",work,".pdf"))
  if(is.null(r$body)){log(id,work,"download_unavailable",r$status);next}
  if(length(r$body)<5 || !identical(r$body[1:5],charToRaw("%PDF-"))){log(id,work,"not_pdf",r$status);next}
  folder<-file.path(root,id);dir.create(folder,showWarnings=FALSE)
  part<-tempfile("openalex-",tmpdir=folder,fileext=".part");writeBin(r$body,part)
  info<-suppressWarnings(system2(pdfinfo,shQuote(part),stdout=TRUE,stderr=TRUE))
  if(!any(grepl("^Pages:[[:space:]]+[1-9][0-9]*",info))){unlink(part);log(id,work,"invalid_pdf");next}
  # Recheck folders just before saving in case a manual download arrived meanwhile.
  if(length(pdfs(id))){unlink(part);log(id,work,"preserved_new_manual_pdf");next}
  dest<-file.path(folder,paste0(id,"_",work,"_openalex_main.pdf"))
  stopifnot(!file.exists(dest),file.rename(part,dest))
  log(id,work,"downloaded",200L,dest,digest::digest(file=dest,algo="sha256"));saved<-saved+1L
  message(id,": downloaded (",saved," new PDFs)")
 }
}
after<-digest::digest(file=book,algo="sha256");writeLines(after,file.path(state,"workbook-sha256-after.txt"))
writeLines(c(paste("New PDFs:",saved),paste("API stopped on access/allowance response:",blocked),paste("Workbook unchanged:",identical(before,after))),file.path(state,"summary.txt"))
message("Finished: ",saved," PDFs. Workbook unchanged: ",identical(before,after),". Logs: ",state)
