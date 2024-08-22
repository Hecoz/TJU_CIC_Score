#学业奖计算程序
#作者：zhaoyignquan
#最后修改时间：2022-9-6
library(readxl)
library(xlsx)
source("utils.R")

#main
allStuData <- readData("../01input", "2022学生成绩/2022全体成绩.xls", "Sheet0")
allStuData <- dataProcessing(allStuData)

allStuNum <-  unique(allStuData[,c("学号")])

newStuData <- NULL
for (stuNum in allStuNum$学号) {
  
  currentStu <- allStuData[which(allStuData$学号 == stuNum),]
  currentStu <- recalcScoreAndRemoveDupFailedClasses(currentStu)
  currentStu <- removeFailedClasses(currentStu)
  if(nrow(currentStu) > 0){
    
    newStuData <- rbind(newStuData, currentStu)
  }
}

masters <- getDataByStuType(newStuData, c("博士生","同等学力博","工程博士","留学博士","硕士生","双证工程硕","工程硕士","留学硕士"))
allMasterStuNum <- unique(masters[,c("学号")])

allAllowStuData <- readData("../01input","05-奖学金发放对象.xlsx", "Sheet0")
allAllowStuNum <-  unique(allAllowStuData[,c("学号")])

classesInfo <- computeClassesInfo(newStuData)

FinalScore <- NULL
for(stuNum in allMasterStuNum$学号){
  
#  if(stuNum %in% allAllowStuNum$学号){
    
    currentStu <- masters[which(masters$学号 == stuNum),]
    currentStuName <- currentStu$姓名[1]
    print(paste(stuNum, currentStuName,sep = "-"))
    
    stuInfo <- calcCurrentStuFinalScore(classesInfo, currentStu)
    FinalScore <- rbind(FinalScore, stuInfo)
#  }
}

FinalScore <- data.frame(FinalScore)
colnames(FinalScore) <- c("学号","姓名", "院系", "专业", "修读类别", "学籍状态", "考试类别", "学生类别", "考试情况", "导师","总分")
write.xlsx(FinalScore, paste("../03output/最终成绩2022",".xls",sep = ""),row.names = FALSE,sheetName = "Sheet0")



