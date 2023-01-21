library(RODBC)##載入已安裝的套件
library(readxl)
library(dplyr)
library(xlsx)
library(lubridate) ##ymd()##將年月日格式的文字轉為日期物件
library(stringr)
library(tidyr)
library(readr)
library(stringi)
library(splitstackshape)
library(xlsx)
library(tibble)
library(data.table)


####至onedrive網址→上一層目錄→整包資料夾檔案下載放置以下指定資料夾
setwd("C:/Users/user/Desktop/123")##設置工作目錄
getwd()##顯示目前工作目錄
path <- "C:/Users/user/Desktop/123/"
files <- list.files(path=path, pattern="*.csv")##列出路徑中的文件
##讀取檔案後將檔案更名另存檔至另外一個指定資料夾
i=1
while(i<=5){
for (file in files) {

 bindtemp <- read.csv(paste(path,file,sep="" ),colClasses = "character" )##paste拼接字串，資料型態為字串，csv檔案讀取##read.csv("路徑",header=T)
 bindtemp = bindtemp %>% mutate(無 = ifelse(names(bindtemp)[5]=='','',names(bindtemp)[5])) %>%##ifelse(邏輯判斷,判斷為真執行,判斷為偽執行)##mutate(資料框,變數)，mutate(iris,abc=1:150)為資料框新增欄位
   select(c(1:5)) %>%
   mutate(a = names(bindtemp)[1],
          b = names(bindtemp)[2],
          c = names(bindtemp)[3],
          d = names(bindtemp)[4],
          f = names(bindtemp)[5])
 
 bindtemp = setNames(bindtemp , c("v1","v2","v3","v4","v5","a","b","c","d","f"))##使用 setNames() 函數對bindtemp資料框的欄位重新命名 
 bindtemp = bindtemp %>% filter(v1!='')##對 bindtemp 資料框中 v1 欄位為空值的資料進行篩選，只保留 v1 欄位有值的資料。
 
 setwd("C:/Users/user/Desktop/788")
 write.table(bindtemp, file=paste0(i,".csv"), sep=",", row.names=FALSE,na="")##輸出wirte.table(資料框,"路徑")##使用 paste0() 函數將 i 變數和 ".csv" 拼接成新的檔名， sep="," 是指定欄位之間的分隔符為 ","，row.names=FALSE 是指不要顯示資料框中的 row names，na="" 是指空值要顯示為 ""
 
 setwd("C:/Users/user/Desktop/123")
 i=i+1
  }
}
##資料夾中所有檔案合併，並篩選與新增欄位
setwd("C:/Users/user/Desktop/788")
getwd()
path <- "C:/Users/user/Desktop/788/"
files <- list.files(path=path, pattern="*.csv")
bindtemp = data.frame()##資料框
bindtemp1= data.frame()
for (file in files) {##如果 files 陣列是 ["file1.csv", "file2.csv", "file3.csv"]，那麼在第一次迴圈執行時，會將 "file1.csv" 賦值給 file 變數
  
  bindtemp <- read.csv(paste(path,file,sep="" ),colClasses = "character" )
 
 bindtemp1 <- rbind(bindtemp1,bindtemp)##將兩個資料框上下結合
}
bindtemp1$c = gsub('X', '', bindtemp1$c)##替換字符串gsub(要匹配的,要替換的,字符串向量)，bindtemp1$c代表取出bindtemp1資料集中的c欄位
bindtemp1$d = gsub('X', '', bindtemp1$d)
bindtemp1$d =gsub('\\.', '', bindtemp1$d)
bindtemp1$d = gsub('/', '', bindtemp1$d)
bindtemp1$v1 = gsub('/', '0', bindtemp1$v1)
bindtemp1$v1 = gsub('&', '0', bindtemp1$v1)
`%notlike%` <- Negate(`%like%`)##我們可以使用 %notlike% 來判斷一個字符串是否不包含另一個字符串。例如，"abc" %notlike% "a" 返回 FALSE，因為 "abc" 字符串中包含 "a"。
years = c("2017","2018","2019","2020","2021","2022","2023","2024","2025")
bindtemp1 = bindtemp1 %>% mutate(d = substr(d,1,8),v1=ifelse(substr(v1,1,1)=='b' | substr(v1,1,1)=='B' |   ##|或，&且
                                                               substr(v1,1,1)=='d' | substr(v1,1,1)=='D' |
                                                               substr(v1,1,1)=='m' | substr(v1,1,1)=='M',substr(v1,1,8),v1)) %>%
  rename(STDNO=v1 , Name=v2 , TYPE=v3 , ClubName = a , Activity = b , ActivityTerm=c , Activity_Date=d , TYPE2=f) %>%
  select(-v4,-v5) %>%##將 v4, v5 欄位從資料集中刪除
  mutate(for(i in 1:length(years)){
           Activity_Date = ifelse(substr(Activity_Date,1,4)==years[i],paste0((years[i]-1911),substr(Activity_Date,5,8)),Activity_Date)
         } ,
  Activity_Date = 
           ifelse(substr(Activity_Date,4,4)=='1' | substr(Activity_Date,4,4)=='0',Activity_Date,
                  ifelse(Activity_Date=='','',paste0( substr(Activity_Date,1,3),'0',substr(Activity_Date,4,6)))),## Activity_Date 欄位進行修改，如果第4位為0或1，則直接返回原值，否則就把第1-3位,0,第4-6位 拼接起來
  Activity_Date = substr(Activity_Date,1,7),
  ACADMYEAR = substr(ActivityTerm,1,3),
  ##ACADMYEAR = '108',
  Term = ifelse((substr(Activity_Date,4,7)>='0801' | substr(Activity_Date,4,7)<='0131') & Activity_Date !='','1',##ACADMYEAR = '108'時，Term=1
                       ifelse(Activity_Date =='','','2')),
  STDNO = ifelse(substr(STDNO,1,1)=='b' & str_length(STDNO)>=8,paste0('B',substr(STDNO,2,10)),
                 ifelse(substr(STDNO,1,1)=='d' & str_length(STDNO)>=8,paste0('D',substr(STDNO,2,10)),
                        ifelse(substr(STDNO,1,1)=='m' & str_length(STDNO)>=8,paste0('M',substr(STDNO,2,10)),STDNO))),
##STDNO欄位與Name欄位是否對調
  STDNO_1=STDNO,  Name_1=Name,
  STDNO = ifelse(Name %like% '^[a-zA-Z_0-9]' &
                   (substr(STDNO,1,1)!='B' & substr(STDNO,1,1)!='D' & substr(STDNO,1,1)!='M') ,Name_1,STDNO),##如果 Name 欄位的值以英文字母或數字開頭且 STDNO 欄位的第一個字符不是 B、D、M，則返回 Name_1 的值，否則返回 STDNO 的值。
  Name = ifelse(Name %like% '^[a-zA-Z_0-9]' &
    (substr(STDNO,1,1)!='B' & substr(STDNO,1,1)!='D' & substr(STDNO,1,1)!='M') ,STDNO_1,Name),
  STDNO = ifelse(Name %like% '老師'  ,Name_1,STDNO),
  Name = ifelse(Name %like% '老師'  ,STDNO_1,Name),
  A = ifelse(substr(STDNO,1,1)=='B' | substr(STDNO,1,1)=='D' |substr(STDNO,1,1)=='M' ,1,NA))
  bindtemp1_ = bindtemp1 %>% filter(!is.na(A)) %>% select(-A,-STDNO_1,-Name_1)##用 filter() 篩選出 A 這個欄位不是 NA 的資料，再用 select() 選擇所有除了 A、STDNO_1、Name_1 這三個欄位的資料，並將篩選後的資料存入 bindtemp1_ 這個變數裡
  bindtemp1_3 = bindtemp1 %>% filter(is.na(A) & str_length(STDNO)>=2 & STDNO!='學號' & substr(STDNO,1,1)!='#') %>% 
    select(-A,-STDNO_1,-Name_1)##只保留 A 欄位為 NA，且學號長度大於等於 2，且STDNO不等於 "學號" 且學號第一個字元不等於 "#" 的資料，並且移除 A、STDNO_1、Name_1 欄位，結果存入 bindtemp1_3 變數中。
  
##去除重複值，若Activity欄位為"一般志工"與"服務活動"無須剃除重複值
  bindtemp1_1 = bindtemp1_ %>% filter(Activity=='一般志工' | Activity=='服務活動')
  bindtemp1_2 = bindtemp1_ %>% filter(!Activity=='一般志工' & !Activity=='服務活動')
  bindtemp1_2 <- unique( bindtemp1_2 ) 
##合併本校生(剔除重複值)、校外生(不剔除重複值)  
  bindtemp1 = rbind(bindtemp1_1,bindtemp1_2,bindtemp1_3)

##連接資料庫
conn <- odbcConnect("SQL_IR", uid="sa", pwd="Ir5829iro@661")
UP_OSAClubActivity = sqlQuery(conn, 'select * from UP_OSAClubActivity' ,as.is = TRUE)
##原始檔案備份
write.table(SAClUP_OubActivity, file="UP_OSAClubActivity.csv", sep=",", row.names=FALSE,na="")
##匯入資料庫
sqlSave(conn, bindtemp1 , tablename = "UP_OSAClubActivity", rownames=FALSE ,append = T)##表示將 R 的 dataframe bindtemp1 儲存到資料庫中名為 UP_OSAClubActivity 的資料表中，由於 append=T 表示是追加資料而非覆蓋，rownames=FALSE 則表示不將 R dataframe 中的 row.names 儲存到資料庫中
close(conn)
####匯入後再去跑R檔案"SA_data_ClubAc_code.R"

