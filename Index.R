rm(list=ls())
library(xlsx)
setwd("F:/work/index/table")
balancestate <- read.xlsx("qm_balancesheet.xlsx",1,colIndex = c(3,5),
                          header = T,encoding="UTF-8")

profitstate <- read.xlsx("qm_profitstatement.xlsx",1,colIndex = c(3,4),
                          header = T,encoding="UTF-8")

cashstate <- read.xlsx("qm_cashflows.xlsx",1,colIndex = c(2,5),
                          header = T,encoding="UTF-8")

table.name <- c("balance","cash","profit")

path <- paste("F:/work/index/table/",table.name,sep="")

rd <- function(cid,current_period){

  name_txt <- paste("cid",cid,".txt",sep="")
 
  setwd(path[1])
  d <- read.table(name_txt,header = F)[,-(1:2)]
  names(d) <- c("truerownum","value","total","month")
  d.balance <- subset(merge(d,balancestate,by="truerownum",all = FALSE, sort = F),month==current_period)
  
  setwd(path[2])
  d <- read.table(name_txt,header = F)[,-(1:2)]
  names(d) <- c("truerownum","value","total","month")
  d.cash <- subset(merge(d,cashstate,by="truerownum",all = FALSE, sort = F),month==current_period)
  
  setwd(path[3])
  d <- read.table(name_txt,header = F)[,-(1:2)]
  names(d) <- c("rownum","value","total","month")
  d.profit <- subset(merge(d,profitstate,by="rownum",all = FALSE, sort = F),month==current_period)
  
  
  
  return(list(b=d.balance,c=d.cash,p=d.profit))
  
}

index <- function(cid,current_period,prior_period){

  
  #��������
  dc.b <- rd(cid,current_period)$b;dc.b
  dc.c <- rd(cid,current_period)$c;dc.c
  dc.p <- rd(cid,current_period)$p;dc.p
  #��������
  dp.b <- rd(cid,prior_period)$b
  dp.c <- rd(cid,prior_period)$c
  dp.p <- rd(cid,prior_period)$p

 
  
  #��ծ��ָ��:
  # 1�ʲ���ծ��=��ծ�ܶ�/�ʲ��ܶ�
  DAR <- dc.b[which(dc.b$truerownum == 47),2] / dc.b[which(dc.b$truerownum == 30),2]
  # 2���������������ʲ���������ծ�ı���
  CR <-  dc.b[which(dc.b$truerownum == 15),2] / dc.b[which(dc.b$truerownum == 41),2]
  # 3�ٶ�����=�ٶ��ʲ�/������ծ��100%�����У��ٶ��ʲ�=�����ʽ�+�����Խ����ʲ�+Ӧ���˿�+Ӧ��Ʊ��
  ATR <- sum(dc.b[which(dc.b$truerownum == 1),2],dc.b[which(dc.b$truerownum == 2),2],
             dc.b[which(dc.b$truerownum == 3),2],dc.b[which(dc.b$truerownum == 4),2]
             ) / dc.b[which(dc.b$truerownum == 41),2]
  
  
  #Ӫ����ָ�꣺
  # 4�ʲ���ת�ʣ��¶ȣ� = Ӫҵ����/ƽ���ʲ��ܶ� ,ƽ���ʲ��ܶ�=(�����ʲ���ծ����ĩ��+�����ʲ���ծ����ĩ��) /2
  AT <- dc.p[which(dc.p$rownum == 1),2] / mean(c(dc.b[which(dc.b$truerownum == 30),2],dp.b[which(dc.b$truerownum == 30),2]))
  # 5�����ʲ���ת�ʣ��¶ȣ�=Ӫҵ����/ƽ�������ʲ��ܶ�
  CAT <- dc.p[which(dc.p$rownum == 1),2] / mean(c(dc.b[which(dc.b$truerownum == 15),2],dp.b[which(dc.b$truerownum == 15),2]))
  # # 6�����ת�ʣ��£�=�����ɱ���ƽ�������� ���У������ɱ�ָӪҵ�ɱ���ƽ��������=�����ڴ��+���ڴ����/2 ,
  # IT <-  dc.p[which(dc.p$rownum == 2),2] / mean(c(dc.b[which(dc.b$truerownum == 9),2],dp.b[which(dc.b$truerownum == 9),2]))
  # # 7Ӧ���˿���ת��=��������/ƽ��Ӧ���˿�=��������\(0.5 x(Ӧ���˿�������ĩ+������ĩ)) 
  # ART <- dc.c[which(dc.c$truerownum == 1),2] / mean(c(dc.b[which(dc.b$truerownum == 4),2],dp.b[which(dc.b$truerownum == 4),2]))

    
  #ӯ����ָ��:  
  # 8Ӫҵ������(����������)=Ӫҵ����/Ӫҵ�����100%
  OPR <- dc.p[which(dc.p$rownum == 21),2] / dc.p[which(dc.p$rownum == 1),2]
  # 9�ɱ�����������=�����ܶ�/�ɱ������ܶ��100%  �ɱ�����һ��ָ��Ӫҵ��ɱ������Ӻ������ڼ����(���۷��á��������� ���������)��
  CPR <- dc.p[which(dc.p$rownum == 30),2] / sum(c(dc.p[which(dc.p$rownum == 2),2],dc.p[which(dc.p$rownum == 3),2],
                                                  dc.p[which(dc.p$rownum == 11),2],dc.p[which(dc.p$rownum == 14),2],
                                                  dc.p[which(dc.p$rownum == 18),2]))
  # 10�ʲ�������=������/�ʲ��ܶ� 
  ROA <- dc.p[which(dc.p$rownum == 32),2]/ dc.b[which(dc.b$truerownum == 30),2]
  # 11���ʲ��������ֳƹɶ�Ȩ��������,�Ǿ�������ƽ���ɶ�Ȩ��İٷֱ�
  ROE <- dc.p[which(dc.p$rownum == 32),2]/ mean(c(dc.b[which(dc.b$truerownum == 52),2],dp.b[which(dc.b$truerownum == 52),2]))
  
  
  
  #��չ��ָ��
  # 12Ӫҵ����������=������Ӫҵ����-����Ӫ�����룩/����Ӫҵ����*100% 
  MBIGR <-  (dc.p[which(dc.p$rownum == 1),2] - dp.p[which(dp.p$rownum == 1),2])/dp.p[which(dp.p$rownum == 1),2]
  # 13������������=(���ھ�����-���ھ�����)/���ھ�����*100%
  NPGR <- (dc.p[which(dc.p$rownum == 32),2] - dp.p[which(dp.p$rownum == 32),2])/dp.p[which(dp.p$rownum == 32),2]
  # 14�ʲ�������=(�����ʲ�-�����ʲ�)/�����ʲ���100%
  AGR <- (dc.b[which(dc.b$truerownum == 30),2] - dp.b[which(dp.b$truerownum == 30),2])/dp.b[which(dp.b$truerownum == 30),2]
  # 15Ӫҵ���������ʣ��������������ʣ�=������Ӫҵ�����ܶ�-����Ӫҵ�����ܶ/����Ӫҵ�����ܶ��100%
  OPGR <- (dc.p[which(dc.p$rownum == 30),2] - dp.p[which(dp.p$rownum == 30),2])/dp.p[which(dp.p$rownum == 30),2]
  
  
  #�ֽ�����
  # 16�ֽ�������ծ��=��Ӫ��ֽ����� / ��ĩ������ծ
  CLDR <- dc.c[which(dc.c$truerownum == 7),2] / dc.b[which(dc.b$truerownum == 41),2]
  # 17�����ֽ����=��Ӫ��ֽ�����/��������
  SCR <- dc.c[which(dc.c$truerownum == 7),2] / dc.c[which(dc.c$truerownum == 1),2]
  
  indexs <- list("DAR"=DAR,"CR"=CR,"ATR"=ATR,
                 "AT"=AT,"CAT"=CAT,
                 "OPR"=OPR,"CPR"=CPR,"ROA"=ROA,"ROE"=ROE,
                 "MBIGR"=MBIGR,"NPGR"=NPGR,"AGR"=AGR,"OPGR"=OPGR,
                 "CLDR"=CLDR,"SCR"=SCR)
  return(indexs)
}

###########################������ҵ��ĳ��ʱ��������#############################
#�ҳ������ݵĹ�˾id
existcorp <- numeric(0)
existnum <- 0
cid <- c(1:3721,300001:300293)
for(c.num in 1:length(cid) ){
  #c.num <- 1
  cat( cid[c.num],"-",sep="")
  if( file.exists(paste("F:/work/index/table/cash/cid",cid[c.num],".txt",sep="")) & 
      file.exists(paste("F:/work/index/table/balance/cid",cid[c.num],".txt",sep="")) &
      file.exists(paste("F:/work/index/table/profit/cid",cid[c.num],".txt",sep="")))
  {
    cat( cid[c.num],"-",sep="")
    existnum <- existnum + 1
    existcorp[existnum] <- cid[c.num]
  }
}
existcorp
##################ĳ��ʱ���ȫ����ҵ������###
time <- c(c(201405+0:7),c(201500+1:12))
cidexist <- existcorp
for( t.num in  2:length(time)){
  #t.num <- 2
  data.t <- matrix(NA,ncol=15,nrow=length(cidexist))
  colnames(data.t) <- c("DAR","CR","ATR",
                        "AT","CAT",
                        "OPR","CPR","ROA","ROE",
                        "MBIGR","NPGR","AGR","OPGR",
                        "CLDR","SCR")
  rownames(data.t) <- paste("cidexist=",cidexist,sep="")
  
  for(c.num in 1:length(cidexist) ){
    #c.num <- 13
    indexdata <- index(cidexist[c.num],time[t.num],time[t.num-1])
    for(i.row in 1:length(indexdata)){
      if(length( unlist(indexdata[i.row])  ) != 0)   
        data.t[c.num,i.row] <- round(unlist(indexdata[i.row]),4)} 
    cat(cidexist[c.num],"-",sep="")
  }
  

  setwd("F:/work/index/table/data.t")
  head(data.t)
  names <- paste("time",time[t.num],".csv",sep="")
  write.csv(data.t,names)
} 

#############ĳ����ҵȫ��ʱ��������#############################################
time <- c(201405+0:7,201500+1:10)
cidexist <- existcorp
for(c.num in 1:length(cidexist) ){
  #c.num <- 1
  data.c <- matrix(NA,ncol=17,nrow=length(time)-1)
  colnames(data.c) <- c("DAR","CR","ATR",
                        "AT","CAT","IT","ART",
                        "OPR","CPR","ROA","ROE",
                        "MBIGR","NPGR","AGR","OPGR",
                        "CLDR","SCR")
  rownames(data.c) <- paste("t=",time[-1],sep="")
  for(t.num in 2:length(time)){
    # t.num <- 2
    indexdata <- index(cidexist[c.num],time[t.num],time[t.num-1])
    for(i.row in 1:length(indexdata)) 
      if(length( unlist(indexdata[i.row])  ) != 0)   
        data.c[t.num-1,i.row] <- round(unlist(indexdata[i.row]),4)
      
  }
  cat(cidexist[c.num],"-",sep="")
  setwd("F:/table/test1/result/data.c")
  names <- paste("cid",cidexist[c.num],".xlsx",sep="")
  write.xlsx(data.c,names)
}


