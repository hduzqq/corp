
# 将表格按企业拆分

rm(list=ls())

threetab <- c("profit","balance","cash")
for(tabnum in 1:3){
  fname <- paste("E:/work/corp/corp_class/three_table/qm_",threetab[tabnum],".csv",sep="")
  sheet <- read.csv(fname,header = T)
  path <- paste("E:/work/corp/corp_class/three_table/",threetab[tabnum],sep="")
  dir.create(path)
  setwd(path)
  startRow <- 1
  for( i in 2:dim(sheet)[1]){
    cat(i-1,"--")
    if (sheet[i,2] != sheet[i-1,2]){
      
      file_name <- paste("cid",sheet[i-1,2],".txt",sep="")
      endRow <- i-1
      
      data <- sheet[startRow:endRow,]
      
      write.table(data,file=file_name,col.names = F,
                  sep = "\t",append = T)
      startRow <- i
    }
  }
}
