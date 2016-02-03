rm(list=ls())
library(xlsx)
library(dplyr)
setwd("E:/work/corp/corp_class/corp_infor")
corp <- as.data.frame(read.csv("qm_corp.csv",header=T,encoding="utf-8"))
head(corp);dim(corp)
corp1 <- select(corp,id,name,LegalPerson,Address,SignDate,ContractDate,industry_class)
head(corp1);dim(corp1)
write.csv(corp1,"corp1.csv")
###############与上次结果合并##########
rm(list=ls())
library(xlsx)
library(dplyr)
setwd("E:/work/corp/corp_class/corp_infor")
corp1 <- read.csv("corp1.csv",header = T,encoding="utf-8")[,-1]
head(corp1);dim(corp1);names(corp1)
corp2 <- read.csv("m13_after.csv",header = T,encoding="utf-8")[,-1]
head(corp2);dim(corp2);names(corp2)
##按照cid合并merge###
corp3 <- merge(corp1,corp2,by.x="id",by.y="cid",all.x = T)
head(corp3);dim(corp3);names(corp3)
corp4 <- select(corp3,id,name,corp,LegalPerson,Address,SignDate,ContractDate,industry_class,class1)
address <- paste(corp3$address1,corp3$address2,corp3$address3,sep="")
corp5 <- cbind(corp4[,1:5],address,corp4[,6:9])
head(corp5);dim(corp5);names(corp5)
write.csv(corp5,"corp5.csv")



##########删除公司名中的数字############
corp6 <- read.csv("corp5_after.csv",header = T,encoding="utf-8")[,-1]
head(corp6);dim(corp6);names(corp6)
name <- as.character(corp6[,2])

library(stringr)
for(i in 1:length(name)){
 # i <- 1375
  ##数字
  d <- unlist(str_extract_all(name[i], "\\d"))
  if(length(d)!=0){
    for(j in 1:length(d)) name[i] <- sub(d[j],"",name[i])
  }
}
for(i in 1:length(name)){
  #i <- 1375
  ##"_"
  d <- unlist(str_extract_all(name[i], "_"))
  if(length(d)!=0){
    for(j in 1:length(d)) name[i] <- sub(d[j],"",name[i])
  }
}

for(i in 1:length(name)){
  d <- substr(name[i],1,2)
  if( d == "[]") name[i] <- substr(name[i],3,nchar(name[i]))
  
}
corp7  <- cbind(corp6[,1],name,corp6[,3:9])
names(corp7) <- names(corp6)
head(corp7,20)
write.csv(corp7,"corp7.csv")


          
####################缺失行业#######
rm(list=ls())
library(dplyr)
library(xlsx)
setwd("E:/work/corp/corp_class/corp_infor")
corp8 <- read.csv("corp7_after.csv",header = T,encoding="utf-8")[,-1]
head(corp8);dim(corp8);names(corp8)

in_class <- read.xlsx("qm_industry_class.xlsx",1,header=T,encoding="UTF-8")
names(in_class)
corp9 <- merge(corp8,in_class,by.x= "industry_class",by.y="code",all.x=T,sort=F,all.y = T)
head(corp9);dim(corp9);names(corp9)
write.csv(corp9,"corp10.csv")

##################添加其他信息#########

rm(list=ls())
library(dplyr)
setwd("E:/work/corp/corp_class/corp_infor")
corp11 <- read.csv("corp11.csv",header = T,encoding="utf-8")[,-1]
head(corp11);dim(corp11);names(corp11)
name <- paste("cid",corp11[,1],".txt",sep="")
operating_income <- accounts_receivable <- accounts_payable <- as.data.frame(matrix(NA,ncol=4,nrow=length(name)))

for(i in 1:length(name)){
 #i <- 37
  cat(i,"-",name[i],"-",sep="")
  if( file.exists(paste("E:/work/corp/corp_class/three_table/profit/",name[i],sep="")) ){ 
    setwd("E:/work/corp/corp_class/three_table/profit/")
    data.p <- as.data.frame(read.table(name[i],header=T)[,-c(1,2)])
    names(data.p) <- c("cid","rownum","value","total","date")
    if( dim(data.p)[1] != 0){
      
      
      #累计营业收入
      data.p1 <- arrange(filter(data.p, rownum == 1),desc(date))
      if( dim(data.p1)[1] != 0 ){
        ic2014 <-  filter(data.p1, date < 201500 & date > 201400) 
        if( dim(ic2014)[1] != 0 ){
          ic1 <- filter(ic2014,date == max(ic2014[,5]))
          operating_income[i,1] <- ic1[4]
          operating_income[i,2] <- ic1[5]
          
        }
        cat("operating_income2014 done!")
        
        ic2015 <-  filter(data.p1, date < 201600 & date > 201500)
        if(  dim(ic2015)[1] != 0 ){
          ic2 <- filter(ic2015,date == max(ic2015[,5]))
          operating_income[i,3] <- ic2[4]
          operating_income[i,4] <- ic2[5]
        }
        cat("operating_income2015 done!")
      }
    }
    
  }
  if( file.exists(paste("E:/work/corp/corp_class/three_table/balance/",name[i],sep="")) ){
    setwd("E:/work/corp/corp_class/three_table/balance")
    data.b <- as.data.frame(read.table(name[i],header=T)[,-c(1,2)])
    names(data.b) <- c("cid","rownum","value","total","date")
    if(dim(data.b)[1] != 0){
      #累计应收账款
      data.b1 <- arrange(filter(data.b, rownum == 4),desc(date))
      if( dim(data.b1)[1] != 0){
        rc2014 <-  filter(data.b1, date < 201500 & date > 201400) 
        if( dim(rc2014)[1] != 0 ){
          rc1 <- filter(rc2014,date == max(rc2014[,5]))
          accounts_receivable[i,1] <- rc1[3]
          accounts_receivable[i,2] <- rc1[5]
          
        }
        cat("accounts_receivable2014 done!")
        
        rc2015 <-  filter(data.b1, date < 201600 & date > 201500)
        if(  dim(rc2015)[1] != 0 ){
          rc2 <- filter(rc2015,date == max(rc2015[,5]))
          accounts_receivable[i,3] <- rc2[3]
          accounts_receivable[i,4] <- rc2[5]
        }
        cat("accounts_receivable2015 done!")

      }
     
      ###累计应付账款
      data.b2 <- arrange(filter(data.b, rownum == 33),desc(date))
      if( dim(data.b2)[1] != 0){
        pa2014 <-  filter(data.b2, date < 201500 & date > 201400) 
        if( dim(pa2014)[1] != 0 ){
          pa1 <- filter(pa2014,date == max(pa2014[,5]))
          accounts_payable[i,1] <- pa1[3]
          accounts_payable[i,2] <- pa1[5]
        }
        cat("accounts_payable2014 done!")
        pa2015 <-  filter(data.b2, date < 201600 & date > 201500)
        if(  dim(pa2015)[1] != 0 ){
          pa2 <- filter(pa2015,date == max(pa2015[,5]))
          accounts_payable[i,3] <- pa2[3]
          accounts_payable[i,4] <- pa2[5]
        }
        cat("accounts_payable2015 done!")
      }

    }

  }
  cat("\n")
}


setwd("E:/work/corp/corp_class")
dt <- cbind(corp11,operating_income,accounts_receivable,accounts_payable)
head(dt)
names(dt)<- c("Id","Name","LegalPerson" ,"Address","SignDate","ContractDate","industry_class",
              "operating_income2014","OIM2014","operating_income2015","OIM2015",
              "accounts_receivable2014","ARM2014","accounts_receivable2015","ARM2015",
              "accounts_payable2014","APM2014","accounts_payable2015","APM2015")
write.csv(dt,"corp12.csv")


################按行业分别保存####
rm(list=ls())
library(dplyr)
library(xlsx)
setwd("E:/work/corp/corp_class")
corp12 <- read.csv("corp12.csv",header = T,encoding="utf-8")[,-1]
introduction  <- read.xlsx("introduction.xlsx",1,header = T,encoding="UTF-8")
head(corp12);dim(corp12);names(corp12)
class <- names(table(corp12$industry_class))
setwd("E:/work/corp/corp_class/result")

for(i in 1:length(class)){
  dt.c1 <- filter(corp12,industry_class == class[i])
  dt.c2 <- dt.c1[,-7]
  fname <- paste(class[i],".xlsx",sep="")
  write.xlsx(dt.c2,fname,sheetName = class[i],row.names=F,showNA=F,append=T)
  write.xlsx(introduction,fname,sheetName = "说明",row.names=F,showNA=F,append=T)
  cat(i,"-")
}



exist <- numeric()

for(i in 1:length(class)){
#  i <- 1
  dt.c1 <- filter(corp12,industry_class == class[i])
  dt.c2 <- dt.c1[,-7]
  dt.c3 <- dt.c2[,7:18]
  ex <- 0
  for(j in 1:dim(dt.c3)[1]){
    #j <- 5
    ex <- ex + as.numeric( sum(is.na(dt.c3[j,])) != 12)
  }
  exist[i] <- ex
}

names(exist) <- class
write.xlsx(exist,"ex.xlsx")
