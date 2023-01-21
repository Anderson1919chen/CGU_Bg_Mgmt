library(RODBC)
library(readxl)
library(dplyr)
library(xlsx)
library(lubridate) ##ymd()
library(stringr)
library(tidyr)
library(readr)
library(stringi)
library(splitstackshape)
library(xlsx)
library(tibble)
library(data.table)
library(openxlsx)

##載入需要的資料表##
conn <- odbcConnect("SQL_IR", uid="sa", pwd="Ir5829iro@661")
UP_OSAClubActivity = sqlQuery(conn, 'select * from UP_OSAClubActivity' ,as.is = TRUE) 
UP_OSAClubCode = sqlQuery(conn, 'select * from UP_OSAClubCode' ,as.is = TRUE) 
UP_OSAClubTYPEc = sqlQuery(conn, 'select * from UP_OSAClubTYPEc' ,as.is = TRUE)
##原始檔案備份##
require(openxlsx)
write.xlsx(UP_OSAClubCode, file="UP_OSAClubCode_.xlsx", showNA=F, sheetName="UP_OSAClubCode", row.names=FALSE)
write.xlsx(UP_OSAClubTYPEc, file="UP_OSAClubTYPEc_.xlsx", showNA=F, sheetName="UP_OSAClubTYPEc", row.names=FALSE)
close(conn)
UP_OSAClubCode = UP_OSAClubCode %>% mutate(S=1)

#####################社團六類代碼對照表_索取新資料匯入，學年度欄位需先行新增#########################################
fileaddress = "C:/Users/user/Desktop/"
De_OSAClubCode <- read_excel(paste0(fileaddress,"社團六類代碼對照表.xlsx") , col_types = "text")
De_OSAClubCode = De_OSAClubCode %>% rename(Code=編號,ClubName=社團,Activity=活動名稱,ACADMYEAR=學年度)%>%
  mutate(Code_ot=NA) %>% mutate(S=0)

##UP_OSAClubCode欄位新與舊資料串接，若有重複取最新資料
UP_OSAClubCode = rbind(UP_OSAClubCode,De_OSAClubCode) 
UP_OSAClubCode_ = UP_OSAClubCode %>% select(ACADMYEAR,Code,ClubName,S) ##因Code不是唯一值，將其資料變成唯一值
UP_OSAClubCode_ <- unique( UP_OSAClubCode_ )  ##去除重複值

UP_OSAClubCode = UP_OSAClubCode_ %>% arrange(ACADMYEAR,Code,ClubName,S) %>% group_by(ACADMYEAR,Code,ClubName) %>% slice (1) %>%
  merge(UP_OSAClubCode , by=c("ACADMYEAR","Code","ClubName","S")) %>% select(-S)
##去除重複列
UP_OSAClubCode <- unique( UP_OSAClubCode ) 

##讀取UP_OSAClubCode資料表且將Code欄位擷取前2位
UP_OSAClubCode = UP_OSAClubCode %>% mutate(Code1=Code) 
UP_OSAClubCode = concat.split(UP_OSAClubCode, "Code1", sep = "-", drop = TRUE)
UP_OSAClubCode = UP_OSAClubCode %>% 
  mutate(Code_ot = ifelse(Code!='' & nchar(as.character(Code1_2))<=1,paste0(as.character(Code1_1),"-0",as.character(Code1_2)),
                          ifelse(Code!='' & nchar(as.character(Code1_2))==2,paste0((Code1_1),"-",as.character(Code1_2)),''))) %>%
  select(-Code1_1,-Code1_2,-Code1_3)

####匯入資料庫
conn <- odbcConnect("SQL_IR", uid="sa", pwd="Ir5829iro@661")
##清除資料
sqlQuery(conn, 'truncate table UP_OSAClubCode')
##上傳
sqlSave(conn, UP_OSAClubCode , tablename = "UP_OSAClubCode", rownames=FALSE ,append = T )
odbcClose(conn)
#####################################################################################################################
#####################################################################################################################
#####################################################################################################################
##############學生參與社團活動(UP_OSAClubActivity)與六類(UP_OSAClubTYPEc)之社團名稱是否有空值####################
##UP_OSAClubActivity篩選資料且新增六類型態_供參考
UP_OSAAc = UP_OSAClubActivity %>% select(ClubName,ActivityTerm) %>% mutate(START = NA)
UP_OSAAc$START <- str_locate(pattern ="[A-Za-z]+", UP_OSAAc$ActivityTerm)[,1]  ##取得六類代碼位置(起始位置)
UP_OSAAc = UP_OSAAc %>%
  mutate(END = substring(ActivityTerm,START+1) ) 
UP_OSAAc$END <- str_locate(pattern ="[A-Za-z]+", UP_OSAAc$END)[,1]  ##取得六類代碼位置(結束位置)
UP_OSAAc = UP_OSAAc %>% mutate(END = as.integer(END)) %>%
  mutate(ClubTYPE = ifelse(!is.na(START) & !is.na(END),substr(ActivityTerm,START,START+END),
                           ifelse(!is.na(START) & is.na(END),substr(ActivityTerm,START,START),NA))) %>%
  select(-START,-END)
##去除重複列
UP_OSAAc <- unique( UP_OSAAc ) 
##UP_OSAClubActivity串接UP_OSAClubTYPEc_確認是否有筆數_若有需自行再UP_OSAClubTYPEc 資料表補上
UP_OSAAc %>% merge(UP_OSAClubTYPEc,by=c("ClubName"),all=TRUE) %>% filter(is.na(Club_TYPE)) %>% 
  select(-ClubName_Consolidate,-Code_ot,-Club_TYPE)
##注意!!若有筆數新增UP_OSAClubTYPEc欄位說明:ClubName→輸入/複製產生之ClubName、Club_TYPE→參考產生之ClubTYPE，對應碼參考如*、
## ClubName_Consolidate→重新定義社團名稱，必須與原本ClubName_Consolidate欄位相同，除非社團是今年度新增，絕對不可自行定義名稱、
## Code_ot→以UP_OSAClubCode之最新學年度為參考基準，先找到UP_OSAClubCode中對應的ClubName再看看其Code_ot，若最新年度Code_ot無值可為NULL
#####*L:公共類社團、O:宗教暨服務類社團、P:地區類社團,、Q:學藝類社團、R:音樂類社團、S:體育類社團、其他:自行判斷

