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

  
  #本期数据
  dc.b <- rd(cid,current_period)$b;dc.b
  dc.c <- rd(cid,current_period)$c;dc.c
  dc.p <- rd(cid,current_period)$p;dc.p
  #上期数据
  dp.b <- rd(cid,prior_period)$b
  dp.c <- rd(cid,prior_period)$c
  dp.p <- rd(cid,prior_period)$p

 
  
  #偿债性指标:
  # 1资产负债率=负债总额/资产总额
  DAR <- dc.b[which(dc.b$truerownum == 47),2] / dc.b[which(dc.b$truerownum == 30),2]
  # 2流动比率是流动资产对流动负债的比率
  CR <-  dc.b[which(dc.b$truerownum == 15),2] / dc.b[which(dc.b$truerownum == 41),2]
  # 3速动比率=速动资产/流动负债×100%，其中：速动资产=货币资金+交易性金融资产+应收账款+应收票据
  ATR <- sum(dc.b[which(dc.b$truerownum == 1),2],dc.b[which(dc.b$truerownum == 2),2],
             dc.b[which(dc.b$truerownum == 3),2],dc.b[which(dc.b$truerownum == 4),2]
             ) / dc.b[which(dc.b$truerownum == 41),2]
  
  
  #营运性指标：
  # 4资产周转率（月度） = 营业收入/平均资产总额 ,平均资产总额=(上月资产负债表期末数+当月资产负债表期末数) /2
  AT <- dc.p[which(dc.p$rownum == 1),2] / mean(c(dc.b[which(dc.b$truerownum == 30),2],dp.b[which(dc.b$truerownum == 30),2]))
  # 5流动资产周转率（月度）=营业收入/平均流动资产总额
  CAT <- dc.p[which(dc.p$rownum == 1),2] / mean(c(dc.b[which(dc.b$truerownum == 15),2],dp.b[which(dc.b$truerownum == 15),2]))
  # # 6存货周转率（月）=销货成本÷平均存货余额 其中：销货成本指营业成本，平均存货余额=（上期存货+本期存货）/2 ,
  # IT <-  dc.p[which(dc.p$rownum == 2),2] / mean(c(dc.b[which(dc.b$truerownum == 9),2],dp.b[which(dc.b$truerownum == 9),2]))
  # # 7应收账款周转率=销售收入/平均应收账款=销售收入\(0.5 x(应收账款上月期末+当月期末)) 
  # ART <- dc.c[which(dc.c$truerownum == 1),2] / mean(c(dc.b[which(dc.b$truerownum == 4),2],dp.b[which(dc.b$truerownum == 4),2]))

    
  #盈利性指标:  
  # 8营业利润率(销售利润率)=营业利润/营业收入×100%
  OPR <- dc.p[which(dc.p$rownum == 21),2] / dc.p[which(dc.p$rownum == 1),2]
  # 9成本费用利润率=利润总额/成本费用总额×100%  成本费用一般指主营业务成本及附加和三项期间费用(销售费用、管理费用 、财务费用)。
  CPR <- dc.p[which(dc.p$rownum == 30),2] / sum(c(dc.p[which(dc.p$rownum == 2),2],dc.p[which(dc.p$rownum == 3),2],
                                                  dc.p[which(dc.p$rownum == 11),2],dc.p[which(dc.p$rownum == 14),2],
                                                  dc.p[which(dc.p$rownum == 18),2]))
  # 10资产报酬率=净利润/资产总额 
  ROA <- dc.p[which(dc.p$rownum == 32),2]/ dc.b[which(dc.b$truerownum == 30),2]
  # 11净资产收益率又称股东权益收益率,是净利润与平均股东权益的百分比
  ROE <- dc.p[which(dc.p$rownum == 32),2]/ mean(c(dc.b[which(dc.b$truerownum == 52),2],dp.b[which(dc.b$truerownum == 52),2]))
  
  
  
  #发展性指标
  # 12营业收入增长率=（本期营业收入-上期营务收入）/上期营业收入*100% 
  MBIGR <-  (dc.p[which(dc.p$rownum == 1),2] - dp.p[which(dp.p$rownum == 1),2])/dp.p[which(dp.p$rownum == 1),2]
  # 13净利润增长率=(当期净利润-上期净利润)/上期净利润*100%
  NPGR <- (dc.p[which(dc.p$rownum == 32),2] - dp.p[which(dp.p$rownum == 32),2])/dp.p[which(dp.p$rownum == 32),2]
  # 14资产增长率=(当期资产-上期资产)/上期资产×100%
  AGR <- (dc.b[which(dc.b$truerownum == 30),2] - dp.b[which(dp.b$truerownum == 30),2])/dp.b[which(dp.b$truerownum == 30),2]
  # 15营业利润增长率（销售利润增长率）=（本期营业利润总额-上期营业利润总额）/上期营业利润总额×100%
  OPGR <- (dc.p[which(dc.p$rownum == 30),2] - dp.p[which(dp.p$rownum == 30),2])/dp.p[which(dp.p$rownum == 30),2]
  
  
  #现金流量
  # 16现金流动负债比=经营活动现金净流量 / 期末流动负债
  CLDR <- dc.c[which(dc.c$truerownum == 7),2] / dc.b[which(dc.b$truerownum == 41),2]
  # 17销售现金比率=经营活动现金净流量/销售收入
  SCR <- dc.c[which(dc.c$truerownum == 7),2] / dc.c[which(dc.c$truerownum == 1),2]
  
  indexs <- list("DAR"=DAR,"CR"=CR,"ATR"=ATR,
                 "AT"=AT,"CAT"=CAT,
                 "OPR"=OPR,"CPR"=CPR,"ROA"=ROA,"ROE"=ROE,
                 "MBIGR"=MBIGR,"NPGR"=NPGR,"AGR"=AGR,"OPGR"=OPGR,
                 "CLDR"=CLDR,"SCR"=SCR)
  return(indexs)
}

###########################所有企业在某个时间点的数据#############################
#找出有数据的公司id
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
##################某个时间点全部企业的数据###
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

#############某个企业全部时间点的数据#############################################
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



