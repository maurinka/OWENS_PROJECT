# Reproducible second-pass retrieval for the EXISTING RDM workbook only.
# Run in RStudio: source("download_other_sources.R")
# No OpenAlex API requests, reference/citation expansion, login automation,
# CAPTCHA handling, paywall bypass, ID reassignment, or deletion of old PDFs.
# Sources (checked 2026-10-01):
# https://www.crossref.org/documentation/retrieve-metadata/rest-api/
# https://pmc.ncbi.nlm.nih.gov/tools/id-converter-api/
# https://pmc.ncbi.nlm.nih.gov/tools/pmcaws/
# https://pmc-oa-opendata.s3.amazonaws.com/README.txt
# PMC files supplied by NIH NLM NCBI PubMed Central Article Datasets on AWS.
# Cached responses, per-study logs, and sessionInfo make decisions auditable.

required <- c("readxl","openxlsx","httr2","jsonlite","xml2","digest","zip")
missing <- required[!vapply(required,requireNamespace,logical(1),quietly=TRUE)]
if(length(missing)) stop("Install required packages: ",paste(missing,collapse=", "))
ROOT <- Sys.getenv("RDM_ROOT_DIR","/Users/maura/Desktop/KnowldegeBase")
BOOK <- file.path(ROOT,"reconciled_master_corrected.xlsx")
STATE <- file.path(ROOT,"download_records","other_sources")
MAX_STUDIES <- as.numeric(Sys.getenv("RDM_OTHER_LIMIT","Inf"))
WORKERS <- as.integer(Sys.getenv("RDM_OTHER_WORKERS","2"))
RETRY_COMPLETED <- identical(Sys.getenv("RDM_OTHER_RETRY"),"1")
UPDATE_WORKBOOK <- !identical(Sys.getenv("RDM_OTHER_UPDATE"),"0")
REPORT_ONLY <- identical(Sys.getenv("RDM_OTHER_REPORT_ONLY"),"1")
CONTACT_EMAIL <- Sys.getenv("RDM_CONTACT_EMAIL","") # optional API contact; never inferred
PDFINFO <- Sys.getenv("RDM_PDFINFO",Sys.which("pdfinfo"))
if(!nzchar(PDFINFO)) {
  bundled <- path.expand("~/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/pdfinfo")
  if(file.exists(bundled)) PDFINFO <- bundled
}
if(!nzchar(PDFINFO)) stop("Install Poppler/pdfinfo or set RDM_PDFINFO to its executable path.")
dir.create(STATE,recursive=TRUE,showWarnings=FALSE)
dir.create(file.path(STATE,"cache"),showWarnings=FALSE)
RUN <- format(Sys.time(),"%Y%m%d-%H%M%S")
script_args <- sub("^--file=","",grep("^--file=",commandArgs(),value=TRUE))
if(length(script_args) && file.exists(script_args[1])) file.copy(script_args[1],file.path(STATE,paste0("script-",RUN,".R")),overwrite=FALSE)
writeLines(capture.output(sessionInfo()),file.path(STATE,paste0("session-",RUN,".txt")))
jsonlite::write_json(list(run=RUN,root=ROOT,max_studies=as.character(MAX_STUDIES),workers=WORKERS,
  retry_completed=RETRY_COMPLETED,update_workbook=UPDATE_WORKBOOK,
  scope="existing workbook IDs only; one verified PDF per missing study; no OpenAlex API requests"),
  file.path(STATE,paste0("settings-",RUN,".json")),auto_unbox=TRUE,pretty=TRUE)
`%or%` <- function(x,y) if(is.null(x) || !length(x) || all(is.na(x))) y else x
norm <- function(x) trimws(gsub("[^a-z0-9]+"," ",tolower(iconv(x,to="ASCII//TRANSLIT"))))
doi_norm <- function(x) sub("^(https?://(dx[.])?doi[.]org/|doi:[[:space:]]*)","",tolower(trimws(x)))
verified_page_links <- function(body,base,dois) {
  doc<-tryCatch(xml2::read_html(body),error=function(e) NULL)
  if(is.null(doc)) return(character())
  tags<-xml2::xml_find_all(doc,"//meta[@name or @property]")
  keys<-tolower(xml2::xml_attr(tags,"name"));vals<-xml2::xml_attr(tags,"content")
  page_doi<-doi_norm(vals[keys %in% c("citation_doi","dc.identifier","dc.identifier.doi","dc.identifier.uri")])
  if(!any(page_doi %in% dois)) return(character())
  links<-unique(vals[keys %in% c("citation_pdf_url","wkhealth_pdf_url")])
  xml2::url_absolute(links[!is.na(links)],base)
}
hash <- function(x) digest::digest(x,algo="sha256",serialize=FALSE)
read_sheet <- function(n) as.data.frame(readxl::read_excel(BOOK,n),stringsAsFactors=FALSE)
m <- read_sheet("Master"); src <- read_sheet("Source_records"); audit <- read_sheet("Audit")
mapping <- read_sheet("ID_mapping"); status <- read_sheet("Download_status")
stopifnot(!anyDuplicated(m$ID))
pdfs <- function(id) {
  ids <- unique(c(id,mapping$legacy_ID[mapping$canonical_ID==id]))
  unique(unlist(lapply(ids,function(x) list.files(file.path(ROOT,x),pattern="[.]pdf$",full.names=TRUE,recursive=TRUE,ignore.case=TRUE))))
}
download_state <- ifelse(is.na(m$Download), "", tolower(trimws(m$Download)))
eligible <- which(download_state %in% c("", "to_do", "no_access", "not_there") & vapply(m$ID,function(id) !length(pdfs(id)),logical(1)))
if(is.finite(MAX_STUDIES)) eligible <- head(eligible,MAX_STUDIES)

# Each worker writes its own checkpoint; no concurrent workbook modifications.
run_one <- function(i) {
  id <- m$ID[i]; checkpoint <- file.path(STATE,paste0(id,".rds"))
  if(file.exists(checkpoint) && !RETRY_COMPLETED) {
    old <- readRDS(checkpoint)
    transient <- any(old$events$result %in% c("http_429","network_error","host_deferred_after_rate_limit","worker_error"))
    if((identical(old$result,"downloaded") && file.exists(old$file)) || (!identical(old$result,"downloaded") && !transient)) return(old)
  }
  events <- list(); blocked <- character(); saved <- NA_character_
  record <- function(stage,url,result,http=NA_integer_,evidence="",file=NA_character_) {
    events[[length(events)+1L]] <<- data.frame(ID=id,stage=stage,url=url,result=result,
      http_status=as.character(http),evidence=evidence,file=file,time=format(Sys.time(),"%Y-%m-%d %H:%M:%S"),stringsAsFactors=FALSE)
  }
  get <- function(url,stage,cache=TRUE) {
    if(!grepl("^https?://",url)) return(NULL)
    # PMC content is retrieved through its cloud service, never page scraping.
    host <- httr2::url_parse(url)$hostname
    cp <- file.path(STATE,"cache",paste0(hash(url),".rds"))
    if(cache && file.exists(cp)) {record(stage,url,"cached_response");return(readRDS(cp))}
    cooldown <- file.path(STATE,paste0("cooldown-",hash(host),".rds"))
    if(file.exists(cooldown) && isTRUE(tryCatch(readRDS(cooldown)>as.numeric(Sys.time()),error=function(e) FALSE))) {
      record(stage,url,"host_deferred_after_rate_limit");return(NULL)
    }
    if(host %in% blocked) {record(stage,url,"host_deferred_after_rate_limit");return(NULL)}
    Sys.sleep(0.6)
    r <- tryCatch(httr2::request(url) |> httr2::req_user_agent("RDM-Literature-Retrieval/1.0") |>
      httr2::req_error(is_error=function(r) FALSE) |> httr2::req_timeout(25) |>
      httr2::req_options(maxfilesize=80000000) |> httr2::req_perform(),error=function(e) NULL)
    if(is.null(r)) {record(stage,url,"network_error");return(NULL)}
    code <- httr2::resp_status(r)
    if(code!=200) {
      record(stage,url,paste0("http_",code),code)
      if(code==429) {
        blocked <<- unique(c(blocked,host))
        delay <- suppressWarnings(as.numeric(httr2::resp_header(r,"retry-after") %or% "60"))
        if(is.na(delay)) delay <- 60
        saveRDS(as.numeric(Sys.time())+max(60,delay),cooldown)
      }
      return(NULL)
    }
    result <- list(body=httr2::resp_body_raw(r),url=httr2::resp_url(r),type=httr2::resp_header(r,"content-type") %or% "")
    if(cache) saveRDS(result,cp)
    record(stage,url,"response_received",200)
    result
  }
  json <- function(r) if(is.null(r)) NULL else tryCatch(jsonlite::fromJSON(rawToChar(r$body),simplifyVector=FALSE),error=function(e) NULL)
  accept_pdf <- function(r,url,evidence) {
    if(is.null(r)) return(FALSE)
    if(length(r$body)<5 || !identical(r$body[1:5],charToRaw("%PDF-"))) {record("pdf",url,"not_pdf",evidence=evidence);return(FALSE)}
    temp <- tempfile(fileext=".pdf");on.exit(unlink(temp))
    writeBin(r$body,temp)
    info <- suppressWarnings(system2(PDFINFO,shQuote(temp),stdout=TRUE,stderr=TRUE))
    if(!is.null(attr(info,"status")) || !any(grepl("^Pages:[[:space:]]+[1-9]",info))) {record("pdf",url,"invalid_pdf",evidence=evidence);return(FALSE)}
    digest <- hash(r$body); folder <- file.path(ROOT,id)
    destination <- file.path(folder,paste0(id,"_other_",substr(digest,1,16),".pdf"))
    dir.create(folder,showWarnings=FALSE)
    if(!file.exists(destination) && !file.copy(temp,destination,overwrite=FALSE)) stop("Cannot save PDF")
    saved <<- destination
    record("pdf",url,"downloaded",200,paste(evidence,"SHA256",digest),destination)
    TRUE
  }
  finish <- function(result,best_link="") {
    log <- if(length(events)) do.call(rbind,events) else data.frame()
    if(file.exists(checkpoint)) {
      old <- readRDS(checkpoint)
      file.copy(checkpoint,paste0(checkpoint,".",RUN,".bak"),overwrite=FALSE)
    }
    prior <- status$Download_result[match(id,status$ID)]
    transient <- any(log$result %in% c("http_429","network_error","host_deferred_after_rate_limit"))
    review <- if(result=="downloaded") "downloaded" else if(isTRUE(prior=="pending_API_allowance_reset")) "pending_OpenAlex_allowance_reset" else if(transient) "retry_other_sources_later" else "manual_retrieval_needed"
    answer <- list(ID=id,result=result,review=review,file=saved,best_link=best_link,events=log,run=RUN)
    saveRDS(answer,checkpoint)
    message(id,": ",result)
    answer
  }
  if(length(pdfs(id))) return(finish("existing_files_preserved"))
  rows <- src[src$canonical_ID==id,,drop=FALSE]
  dois <- unique(na.omit(doi_norm(rows$DOI)));dois <- dois[grepl("^10[.][0-9]{4,9}/",dois)]
  if(!length(dois)) {record("identity","","missing_valid_DOI");return(finish("missing_valid_DOI"))}
  best <- paste0("https://doi.org/",dois[1]); pages <- best
  # Reuse OpenAlex metadata already obtained. No new OpenAlex calls/credits.
  for(wid in unique(na.omit(rows$OpenAlex_ID))) {
    cp <- file.path(ROOT,"download_records",paste0(hash(wid),".json"))
    if(file.exists(cp)) {
      oa <- tryCatch(jsonlite::fromJSON(cp,simplifyVector=FALSE),error=function(e) NULL)
      for(loc in oa$locations) if(isTRUE(loc$is_oa) && !is.null(loc$landing_page_url)) pages <- c(pages,loc$landing_page_url)
    }
  }
  for(doi in dois) {
    # Exact DOI lookup, not a free-text result chosen by rank.
    crurl <- paste0("https://api.crossref.org/works/",utils::URLencode(doi,reserved=TRUE))
    cr <- json(get(crurl,"crossref"))$message
    if(!is.null(cr) && identical(doi_norm(cr$DOI),doi)) {
      record("identity",crurl,"verified_DOI",evidence=doi)
      for(link in cr$link) if(grepl("pdf",link$`content-type` %or% "",ignore.case=TRUE)) {
        u <- link$URL
        if(!is.null(u) && accept_pdf(get(u,"crossref_pdf",FALSE),u,paste("Crossref exact DOI",doi))) return(finish("downloaded",u))
      }
      pages <- unique(c(pages,cr$URL,cr$resource$primary$URL))
    }
    # Official PMC ID mapping followed by public S3 version metadata.
    query <- list(ids=doi,idtype="doi",format="json",tool="RDM_Literature_Retrieval")
    if(nzchar(CONTACT_EMAIL)) query$email <- CONTACT_EMAIL
    converter <- do.call(httr2::req_url_query,c(list(httr2::request("https://pmc.ncbi.nlm.nih.gov/tools/idconv/api/v1/articles/")),query))$url
    pm <- json(get(converter,"pmc_id"))
    for(rec in pm$records) if(!is.null(rec$pmcid) && identical(doi_norm(rec$doi %or% ""),doi)) {
      bucket <- "https://pmc-oa-opendata.s3.amazonaws.com/"
      listing <- get(paste0(bucket,"?list-type=2&prefix=",rec$pmcid,".&delimiter=%2F"),"pmc_versions")
      if(is.null(listing)) next
      doc <- tryCatch(xml2::read_xml(listing$body),error=function(e) NULL)
      if(is.null(doc)) next
      prefixes <- xml2::xml_text(xml2::xml_find_all(doc,"//*[local-name()='CommonPrefixes']/*[local-name()='Prefix']"))
      for(prefix in prefixes) {
        metaurl <- paste0(bucket,prefix,sub("/$","",prefix),".json")
        meta <- json(get(metaurl,"pmc_metadata"))
        if(is.null(meta) || !identical(doi_norm(meta$doi %or% ""),doi)) next
        u <- meta$pdf_url
        if(!is.null(u) && nzchar(u)) {
          u <- sub("^s3://pmc-oa-opendata/",bucket,u)
          if(accept_pdf(get(u,"pmc_pdf",FALSE),u,paste("PMC exact DOI",doi,"license",meta$license_code %or% "unspecified"))) return(finish("downloaded",u))
        }
      }
    }
  }
  # Only citation_pdf_url from an identity-verified article page is followed.
  # Generic PDF links, references and supplementary links are not guessed.
  for(u in unique(pages)) {
    if(is.na(u) || !grepl("^https?://",u)) next
    if(grepl("(pmc[.]ncbi[.]nlm[.]nih[.]gov|ncbi[.]nlm[.]nih[.]gov/pmc)",u)) next
    r <- get(u,"landing_page")
    if(is.null(r)) next
    if(length(r$body)>=5 && identical(r$body[1:5],charToRaw("%PDF-"))) {
      if(startsWith(u,"https://doi.org/") && accept_pdf(r,u,"PDF returned by exact DOI resolver")) return(finish("downloaded",u))
      next
    }
    links <- verified_page_links(r$body,r$url,dois)
    if(!length(links)) {record("identity",u,"no_DOI_verified_PDF_link");next}
    for(pdfurl in links) {
      if(accept_pdf(get(pdfurl,"page_pdf",FALSE),pdfurl,"citation PDF link on DOI-verified page")) return(finish("downloaded",pdfurl))
    }
    record("landing_page",u,"no_verified_PDF_retrieved")
  }
  finish("no_PDF_retrieved",best)
}

if(!identical(Sys.getenv("RDM_OTHER_DEFINE_ONLY"),"1")) {
  message("Eligible missing studies: ",length(eligible),". Existing files and IDs are preserved.")
  safe_run <- function(i) tryCatch(run_one(i),error=function(e) {
    id<-m$ID[i];message(id,": worker error (see checkpoint)")
    x<-list(ID=id,result="worker_error",review="retry_needed",file=NA_character_,best_link="",
      events=data.frame(ID=id,stage="worker",url="",result="worker_error",http_status="",evidence=conditionMessage(e),file=NA_character_,time=as.character(Sys.time())),run=RUN)
    saveRDS(x,file.path(STATE,paste0(id,".rds")));x
  })
  if(REPORT_ONLY) {
    message("Rebuilding workbook/report from checkpoints; no network requests.")
  } else if(WORKERS>1 && length(eligible)) {
    # Separate R sessions are portable and avoid macOS fork/network crashes.
    cluster <- parallel::makePSOCKcluster(WORKERS,outfile="")
    parallel::clusterExport(cluster,c("ROOT","BOOK","STATE","RUN","RETRY_COMPLETED","CONTACT_EMAIL","PDFINFO",
      "m","src","audit","mapping","status","%or%","norm","doi_norm","verified_page_links","hash","pdfs","run_one","safe_run"),envir=environment())
    tryCatch(invisible(parallel::parLapplyLB(cluster,eligible,safe_run)),finally=parallel::stopCluster(cluster))
  } else invisible(lapply(eligible,safe_run))
  all <- lapply(list.files(STATE,pattern="^RDM-STUDY-.*[.]rds$",full.names=TRUE),readRDS)
  logs <- do.call(rbind,lapply(all,function(x) x$events))
  report <- do.call(rbind,lapply(all,function(x) data.frame(ID=x$ID,result=x$result,review=x$review,file=x$file,best_link=x$best_link,
    reason=paste(unique(x$events$result[!x$events$result %in% c("response_received","cached_response","verified_DOI")]),collapse=" | "),run=x$run)))
  write.csv(logs,file.path(STATE,"other_sources_log.csv"),row.names=FALSE)
  write.csv(report,file.path(STATE,"other_sources_status.csv"),row.names=FALSE)
  outstanding <- merge(m,report,by="ID",all=FALSE)
  outstanding <- outstanding[outstanding$result!="downloaded" & (is.na(outstanding$Download) | outstanding$Download==""),]
  manual <- outstanding[outstanding$review=="manual_retrieval_needed",]
  retry <- outstanding[outstanding$review!="manual_retrieval_needed",]
  write.csv(manual,file.path(STATE,"manual_retrieval.csv"),row.names=FALSE)
  write.csv(retry,file.path(STATE,"retry_later.csv"),row.names=FALSE)
  if(UPDATE_WORKBOOK) {
    wb<-openxlsx::loadWorkbook(BOOK)
    ok<-report$ID[report$result=="downloaded" & file.exists(ifelse(is.na(report$file),"",report$file))]
    for(id in ok) {
      i<-match(id,m$ID);m$Download[i]<-"automated"
      openxlsx::writeData(wb,"Master","automated",startRow=i+1,startCol=match("Download",names(m)),colNames=FALSE)
      j<-match(id,audit$ID); files<-pdfs(id)
      for(n in c("existing_pdf_count","existing_files","automated_file_exists")) {
        value<-switch(n,existing_pdf_count=length(files),existing_files=paste(files,collapse=" | "),automated_file_exists=TRUE)
        openxlsx::writeData(wb,"Audit",value,startRow=j+1,startCol=match(n,names(audit)),colNames=FALSE)
      }
      j<-match(id,status$ID);status$Download_result[j]<-"PDF_available_locally";status$PDF_count[j]<-length(files)
    }
    openxlsx::writeData(wb,"Download_status",status)
    summary<-read_sheet("Summary")
    summary$value[summary$measure=="Automated PDFs already present"]<-sum(m$Download=="automated",na.rm=TRUE)
    summary$value[summary$measure=="Pending"]<-sum(is.na(m$Download)|m$Download=="")
    openxlsx::writeData(wb,"Summary",summary)
    for(name in c("Other_sources_log","Other_sources_status","Manual_retrieval","Retry_later")) {
      if(name %in% names(wb)) openxlsx::removeWorksheet(wb,name)
      openxlsx::addWorksheet(wb,name)
      data<-switch(name,Other_sources_log=logs,Other_sources_status=report,Manual_retrieval=manual,Retry_later=retry)
      openxlsx::writeData(wb,name,data);openxlsx::freezePane(wb,name,firstRow=TRUE)
      openxlsx::setColWidths(wb,name,cols=seq_len(ncol(data)),widths=24)
      openxlsx::addFilter(wb,name,rows=1,cols=seq_len(ncol(data)))
    }
    # Back up and atomically replace the workbook; leave source spreadsheets alone.
    backup<-file.path(STATE,paste0("master-before-",RUN,".xlsx.bak"))
    stopifnot(file.copy(BOOK,backup,overwrite=FALSE))
    temp<-file.path(ROOT,paste0(".master-",RUN,".xlsx"))
    openxlsx::saveWorkbook(wb,temp,overwrite=TRUE)
    # Some openxlsx versions emit unused relationships to absent drawings.
    # Remove only absent targets and expand filters to all populated rows.
    unpack <- tempfile("xlsx-");dir.create(unpack);utils::unzip(temp,exdir=unpack)
    for(p in list.files(unpack,pattern="[.]rels$",recursive=TRUE,full.names=TRUE)) {
      doc<-xml2::read_xml(p);base<-dirname(dirname(p))
      for(rel in xml2::xml_children(doc)) {
        if(identical(xml2::xml_attr(rel,"TargetMode"),"External")) next
        target<-xml2::xml_attr(rel,"Target")
        target_path<-if(startsWith(target,"/")) file.path(unpack,sub("^/","",target)) else file.path(base,target)
        if(!file.exists(target_path)) xml2::xml_remove(rel)
      }
      xml2::write_xml(doc,p)
    }
    for(p in list.files(file.path(unpack,"xl","worksheets"),pattern="^sheet[0-9]+[.]xml$",full.names=TRUE)) {
      doc<-xml2::read_xml(p);filter<-xml2::xml_find_first(doc,"//*[local-name()='autoFilter']")
      if(!inherits(filter,"xml_missing")) {
        rows<-xml2::xml_find_all(doc,"//*[local-name()='sheetData']/*[local-name()='row']")
        last<-max(as.integer(xml2::xml_attr(rows,"r")))
        col<-gsub("[0-9]","",tail(strsplit(xml2::xml_attr(filter,"ref"),":")[[1]],1))
        xml2::xml_set_attr(filter,"ref",paste0("A1:",col,last))
      }
      xml2::write_xml(doc,p)
    }
    cleaned<-paste0(temp,".clean.xlsx")
    zip::zipr(cleaned,list.files(unpack,recursive=TRUE,all.files=TRUE,no..=TRUE),root=unpack,mode="mirror")
    unlink(unpack,recursive=TRUE);unlink(temp);stopifnot(file.rename(cleaned,temp))
    check<-readxl::read_excel(temp,"Master")
    stopifnot(identical(check$ID,m$ID),"Topic" %in% names(check),sum(check$Download=="automated",na.rm=TRUE)==sum(m$Download=="automated",na.rm=TRUE))
    stopifnot(file.rename(temp,BOOK))
  }
  print(table(report$result)); message("Done. Logs: ",STATE)
}
