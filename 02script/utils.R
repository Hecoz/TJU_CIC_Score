#功能函数
#作者：zhaoyignquan
#最后修改时间：2022-9-6

#读取学生数据
#filepath: 学生数据所在文件夹
#filename：学生成绩文件名
#sheetname：学生成绩数据所在sheet
readData <- function(filepath, filename, sheetname){
  
  data <- read_excel(file.path(filepath,filename,fsep = .Platform$file.sep) ,sheet = sheetname) 
  return(data)
}

#对学生数据进行处理
#1.将分数为空的数据填充为->0
#2.将考试情况为空数据填充为->正常
#data: 学生原始数据
dataProcessing <- function(data){
  
  data$分数[is.na(data$分数)] <- 0
  data$考试情况[is.na(data$考试情况)] <- "正常"
  return(data)
}

#重新计算补考课程分数，同时删除重复无效的课程数据
#1.课程重复，但分数都高于60，保留最高分
#2.缓考，且其中一次高于60分，保留60分，删除小于60分
#3.缓考，且每次都低于60分，删除其考试数据
#currentStu: 当前学生所有数据(考试科目及分数等)
recalcScoreAndRemoveDupFailedClasses <- function(currentStu){
  
  dupClasses <- currentStu[duplicated(currentStu$课程代码),]
  
  if(nrow(dupClasses) > 0){
    
    failedClasses <- unique(dupClasses$课程代码)
    for (c in failedClasses) {
      
      currentFailedClass <- currentStu[which(currentStu$课程代码==c), ]
      maxScore <- max(currentFailedClass[currentFailedClass$课程代码 == c, ]$分数)
      minScore <- min(currentFailedClass[currentFailedClass$课程代码 == c, ]$分数)
      if(maxScore < 60){
        
        currentStu <- currentStu[-which(currentStu$课程序号 %in% currentFailedClass$课程序号),]
      } else {
        
        delayTimes <- nrow(currentFailedClass[which(currentFailedClass$考试情况=="缓考"),])
        saveItem <- currentFailedClass[which.max(currentFailedClass[currentFailedClass$课程代码 == c, ]$分数),]
        currentStu <- currentStu[-which(currentStu$课程代码 == c),]
        currentStu <- rbind(currentStu, saveItem)
        if(minScore < 60){
          if(delayTimes < 1){
            currentStu[which(currentStu$课程代码 == c), "分数"] <- 60
          }else{
            if(delayTimes > 1){
              currentStu[which(currentStu$课程代码 == c), "分数"] <- 60
            }
          }
        }
      }
    }
  }
  return(currentStu)
}

#删除该学生所有不及格的科目
#currentStu: 当前学生所有数据(考试科目及分数等)
removeFailedClasses <- function(currentStu){
  
  ExamFailedStu <- NULL
  #挂科同学
  FinalFailedClass <- currentStu[which(currentStu$分数 < 60), ]
  if(nrow(FinalFailedClass) > 0){
    
    ExamFailedStu <- rbind(ExamFailedStu, currentStu)
    failClass <- unique(FinalFailedClass$课程代码)
    for (c in failClass) {
      currentStu <- currentStu[-which(currentStu$课程代码 == c),]
    }
  }
  return(currentStu)
}

#计算获取课程信息，包括每门课程的平均分，课程标准差
#newStuData: 处理后的学生数据（删除重复，挂科数据）
computeClassesInfo <- function(newStuData){
  
  #课程信息
  classesInfo <- NULL
  allClassNum <- unique(newStuData[, c("课程代码")])
  for(classNum in allClassNum$课程代码){
    
    currentClass <- newStuData[which(newStuData$课程代码 == classNum), ]
    currentClass$分数[is.na(currentClass$分数)] <- 0
    #课程标准差
    SI <- sqrt(sum((currentClass$分数 - mean(currentClass$分数))^2)/nrow(currentClass)) + 0.0001
    classItem <- c(currentClass$课程序号[1], currentClass$课程代码[1], currentClass$课程名称[1], currentClass$课程学分[1], mean(currentClass$分数),SI , nrow(currentClass))
    classesInfo <- rbind(classesInfo, classItem)
  }
  
  classesInfo <- data.frame(classesInfo,stringsAsFactors = FALSE)
  colnames(classesInfo) <- c("课程序号","课程代码","课程名称","课程学分" ,"课程平均分","课程标准差", "选修人数")
  classesInfo$课程学分 <- as.numeric(classesInfo$课程学分)
  classesInfo$课程平均分 <- as.numeric(classesInfo$课程平均分)
  classesInfo$课程标准差 <- as.numeric(classesInfo$课程标准差)
  classesInfo$选修人数 <- as.numeric(classesInfo$选修人数)
  
  #write.xlsx(classesInfo, "../03output/classesInfo.xls",row.names = FALSE)
  return(classesInfo)
}

#根据学生类别获取学生数据
#stuData: 学生数据
#stuType：学生类别
getDataByStuType <- function(stuData, stuType){
  
  data <- stuData[which(stuData$学生类别 %in% stuType),]
  data$分数[is.na(data$分数)] <- 0
  
  return(data)
}


#计算当前学生最终学业成绩
#classesInfo: 课程信息，包含要使用到的课程标准差，课程平均分
#currentStu：当前要计算的学生数据
calcCurrentStuFinalScore <- function(classesInfo, currentStu){
  
  #学生所修课程总学分
  TC <- sum(currentStu$课程学分)
  calcData <- currentStu[, c("课程代码","分数")]
  stuClass <- classesInfo[which(classesInfo$课程代码 %in% currentStu$课程代码), ]
  allData <- merge(stuClass, calcData, by.y = "课程代码")
  
  RS <- sum(allData$课程学分* (allData$分数 - allData$课程平均分)/allData$课程标准差)
  #学生K的标准得分
  KS <- RS/TC
  #学生K的百分制得分
  ##所选课程平均分总分
  MU <- sum(allData$课程平均分)
  ##SIGMA
  SIGMA <- sqrt(sum(allData$课程学分 * allData$课程标准差^2)/sum(allData$课程学分))
  #学生K的百分制得分
  BS <- KS*SIGMA + MU/nrow(allData)
  
  stuInfo <- c(stuNum, currentStuName, 
               currentStu$院系[1], currentStu$专业[1], 
               currentStu$修读类别[1], currentStu$学籍状态[1],
               currentStu$考试类别[1],currentStu$学生类别[1],
               currentStu$考试情况[1],currentStu$导师[1], BS)
  return(stuInfo)
}
