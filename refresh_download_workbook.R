# Reconcile local PDFs and all download logs without changing manual Download values.
library(openxlsx);library(readxl);library(dplyr);library(jsonlite)
root<-Sys.getenv("RDM_ROOT_DIR","/Users/maura/Desktop/KnowldegeBase")
book<-file.path(root,"reconciled_master_corrected.xlsx");before_hash<-digest::digest(file=book,algo="sha256")
read<-function(n)as.data.frame(read_excel(book,n))
wb<-loadWorkbook(book);m<-read("Master");original<-m;a<-read("Audit");mapping<-read("ID_mapping");summary<-read("Summary")
manual<-!is.na(m$Download) & nzchar(m$Download) & m$Download!="automated"
state<-file.path(root,"download_records")
logs<-list(read.csv(file.path(root,"download_batch_log.csv")),stream_in(file(file.path(state,"attempts.jsonl")),verbose=FALSE))
for(p in list.files(state,pattern="^results[.]csv$",recursive=TRUE,full.names=TRUE))logs[[length(logs)+1]]<-read.csv(p)
other<-file.path(state,"other_sources")
checkpoints<-lapply(list.files(other,pattern="^RDM-STUDY-.*[.]rds$",full.names=TRUE),readRDS)
ol<-bind_rows(lapply(checkpoints,function(x)x$events));logs[[length(logs)+1]]<-ol
combined<-bind_rows(lapply(logs,function(z){z[]<-lapply(z,as.character);z}))
report<-bind_rows(lapply(checkpoints,function(x)data.frame(ID=x$ID,result=x$result,review=x$review,file=x$file,best_link=x$best_link,run=x$run)))
statuses<-list()
for(i in seq_len(nrow(m))){
 id<-m$ID[i];aliases<-unique(c(id,mapping$legacy_ID[mapping$canonical_ID==id]))
 files<-unique(unlist(lapply(aliases,function(z)list.files(file.path(root,z),pattern="[.]pdf$",recursive=TRUE,full.names=TRUE,ignore.case=TRUE))))
 z<-combined[!is.na(combined$ID)&combined$ID==id,,drop=FALSE]
 auto<-any(z$result %in% c("downloaded","duplicate_content") & !is.na(z$file) & file.exists(ifelse(is.na(z$file),"",z$file)))
 if(auto && !manual[i])m$Download[i]<-"automated"
 j<-match(id,a$ID);a$existing_pdf_count[j]<-length(files);a$existing_files[j]<-paste(files,collapse=" | ");a$automated_file_exists[j]<-auto
 # Current local evidence supersedes all historical failure/quota messages.
 latest<-report[report$ID==id,,drop=FALSE]
 result<-if(length(files))"PDF_available_locally" else if(grepl("^manual",ifelse(is.na(m$Download[i]),"",m$Download[i])))"manual_status_no_local_PDF" else if(nrow(latest)&&latest$review[1] %in% c("retry_other_sources_later","retry_needed"))"retry_temporary_failure" else "no_PDF_retrieved"
 statuses[[i]]<-data.frame(ID=id,Download=m$Download[i],Download_result=result,PDF_count=length(files),Attempts=nrow(z),Outcomes=paste(sort(unique(z$result)),collapse=" | "))
}
ds<-bind_rows(statuses)
summary$value[summary$measure=="Automated PDFs already present"]<-sum(m$Download=="automated",na.rm=TRUE)
summary$value[summary$measure=="Pending"]<-sum(ds$PDF_count==0)
summary<-summary[!summary$measure %in% c("Studies with local PDFs","Manual Download statuses","to_do annotations retained"),]
summary<-bind_rows(summary,data.frame(measure=c("Studies with local PDFs","Manual Download statuses","to_do annotations retained"),value=c(sum(ds$PDF_count>0),sum(grepl("^manual",m$Download)),sum(m$Download=="to_do",na.rm=TRUE))))
queue<-merge(m,ds[,c("ID","Download_result","PDF_count")],by="ID",sort=FALSE)
queue<-merge(queue,report[,c("ID","review","best_link")],by="ID",all.x=TRUE,sort=FALSE)
queue<-queue[queue$PDF_count==0,]
retry<-queue[queue$Download_result=="retry_temporary_failure",];todo<-queue[queue$Download_result!="retry_temporary_failure",]
replace<-function(name,data){if(name %in% names(wb))removeWorksheet(wb,name);addWorksheet(wb,name);writeData(wb,name,data);freezePane(wb,name,firstRow=TRUE);addFilter(wb,name,rows=1,cols=seq_len(ncol(data)));setColWidths(wb,name,cols=seq_len(ncol(data)),widths=24);addStyle(wb,name,createStyle(textDecoration="bold",fgFill="#E8EEF4"),rows=1,cols=seq_len(ncol(data)),gridExpand=TRUE)}
# Retain Master formatting and edit only automatically inferred Download cells.
writeData(wb,"Master",m$Download,startRow=2,startCol=match("Download",names(m)),colNames=FALSE)
replace("Audit",a);replace("Summary",summary);replace("Download_status",ds);replace("Download_log",combined);replace("Other_sources_log",ol);replace("Other_sources_status",report);replace("Manual_retrieval",todo);replace("Retry_later",retry)
run<-format(Sys.time(),"%Y%m%d-%H%M%S");backup<-file.path(state,paste0("master-before-refresh-",run,".xlsx.bak"));stopifnot(file.copy(book,backup))
temp<-file.path(root,paste0(".refresh-",run,".xlsx"));saveWorkbook(wb,temp,overwrite=TRUE)
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
check<-as.data.frame(read_excel(temp,"Master"))
stopifnot(identical(check$Download,m$Download),identical(check$Download[manual],original$Download[manual]),identical(check$ID,original$ID),identical(check$Topic,original$Topic))
stopifnot(identical(check[,names(check)!="Download"],original[,names(original)!="Download"]))
stopifnot(identical(before_hash,digest::digest(file=book,algo="sha256")))
stopifnot(file.rename(temp,book))
print(table(m$Download,useNA="ifany"));print(table(ds$Download_result));cat("Manual queue:",nrow(todo),"Retry queue:",nrow(retry),"\n")
