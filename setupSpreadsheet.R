library("readxl")
library("writexl")
library("stringr")

# Call this function just to run everything

runThis <- function(){
# First get people data
  allPeople <- setAllPeopleInfo()
# Now organisational support
  allOrgs <- setAllSupportInfo()
# Get currency
  currency <- setCurrencyTitle()
# Get room and per diems
  livingCosts <- setLivingCosts()
# Get other costs
  otherFoodCosts <- setOtherCosts()
  
  fullDF <- buildDataFrame( allPeople, allOrgs, currency, otherFoodCosts, livingCosts)  
  fileName <- readline("Enter the name of the spreadsheet")
   
  write_xlsx(fullDF, fileName)
  return(fullDF)
  
}

#Function just to declare a set of global variables
setUpBudgetBuild <- function(){
  assign("allPeople", list(), envir = .GlobalEnv)
  assign("allOrgs", list(), envir = .GlobalEnv)
  assign("currency", list(), envir = .GlobalEnv)
  assign("livingCosts", list(), envir = .GlobalEnv)
  assign("otherFoodCosts", list(), envir = .GlobalEnv)
}

#xl_formula("=B2+B3")


#df <- data.frame(x=c("x+1"),y=c("y+1"),z=c("z+1"))
#write_xlsx(df, "test.xlsx")
#df <- data.frame(x=c(1,2,3),y=c(4,5,6),z=c("","",""))

#df[2:4,3] <- xl_formula(c("=A2+B2","=A3+B3","=A4+B4"))
#df[,3]    <- xl_formula(c("=A2+B2","=A3+B3","=A4+B4"))
#df
#write_xlsx(df, "test.xlsx",append=TRUE)
#df[1,3] <- "z+1"
#xl_formula(c("=A1+B1","=A2+B2","=A3+B3"))
#df

#names(df)

# Base function to set up information about an instructor/helper/other

instructorInfo <- function(name,location, nights,travel){
  list(name=name,location=location,travel=travel,nights=nights)
}

# Setter function for instructorInfo
setInstructorInfo <- function(){
  
  name <- readline("What is the name of the person? ")
  if (str_length(name) == 0){
    return -1 
  }
  else{
    location <- readline("What institution are they coming from? ")
    n <- readline("How many nights are they staying? ")
    nights <- as.integer(n) 
    t <- readline("What is the estimate of their travel expenses? ")
    travel <- as.integer(t)
    return(instructorInfo(name=name,location=location, nights=nights,travel=travel))
  }
}
  
# Menu option to enter in data for instructors/helpers/etc.
# This creates a list of lists of the form
# Roles -> lisr(course, name1data, name2data, ...)
setAllPeopleInfo <- function(){
  roles <- c("Open Science", "Carpentry", "Computational Infrastructures", 
             "Information Security", "Research Data Management",
             "Analysis","Visualisation","Author Carpentry",
             "Helpers","Organisers", "Photographer","Other","Finish")
  
  courses <- c("Open Science", "Carpentry", "Computational Infrastructures", 
               "Information Security", "Research Data Management",
               "Analysis","Visualisation","Author Carpentry")
  allPeople <- list()
  option <- 0
  while ( option != length(roles) ){
    option <- menu(roles, title = "Type number for a role or to finish")
    if ( option < length(roles) ){
      thisRole <- roles[option] 
      reply <- readline(paste("How many people are going to be in the role",
                              thisRole," "))
      while(is.na(as.integer(reply))){
        reply <- readline("Please enter the _number of people_ who are going to be in this role.")
      }
      nPeople <- as.integer(reply)
      cat(paste("Getting information for",nPeople," individuals on",thisRole))
      x <- list()
      x$Course <- thisRole %in% courses
      for ( i in c(1:nPeople) ){
        cat(paste("Enter data for individual",i,"for",thisRole))
        thisPerson <- setInstructorInfo()
        if ( is.integer(thisPerson)){
          break
        }
        x <- addAndNameToList(x,thisPerson,thisPerson$name)
      }
      allPeople <- addAndNameToList(allPeople,x,thisRole)

    }
  }
  return(allPeople)
}

# Short function to determine if one of the roles is a course.

isThisACourse <- function(thisRole){
  courses <- c("Open Science", "Carpentry", "Computational Infrastructures", 
               "Information Security", "Research Data Management",
               "Analysis","Visualisation","Author Carpentry")
  answer <- thisRole %in% courses
  if ( length(answer) > 1 ){
    warning("isThisACourse expects one string as input returning vector",
            call. = TRUE)
  }
  
  return(answer)

}  

# Short function to add an element to a list and give it a name
addAndNameToList <- function(myList,newElement,newName){
  myList[[length(myList) + 1]] <- newElement
  names(myList)[length(myList)] <- newName
  return(myList)
}

# Base function to create information about support from an institution
supportInfo <- function(name,amount=0){
  list(name=name,amount=amount)
}

#setter for supportInfo
setSupportInfo <- function(){

  name <- readline("What is the name of the supporting organisation? ")
  
  a <- readline("How much have they committed (just put in 0 if you don't know yet) ")
  amount <- as.integer(a)
  return(supportInfo(name,amount))
}

# Get all supporting organisations
setAllSupportInfo <- function(){
  allSupport <- list()
  reply <- readline("How many organisations are contributing? ")
  while(is.na(as.integer(reply))){
    reply <- readline("Please enter the _number of organisations_ that are going to be contributing.")
  }
  n <- as.integer(reply)
  for ( i in c(1:n)){
    x <- setSupportInfo()
    allSupport[[length(allSupport)+1]] <- x
  }
  return(allSupport)
}

#Base function for other Costs
otherCosts <- function(refreshments,dinner){
  list(refreshments=refreshments,dinner=dinner)
}

# Setter for above
setOtherCosts <- function(){
  
  refreshments <- as.integer(readline("What is the estimated refreshments cost? "))
  dinner <- as.integer(readline("What is the esimated costs of the school dinner? "))
  return(otherCosts(refreshments,dinner))

}


# Change currency (if needs be)
currencyTitle <- function(curr="Euro"){
  paste("All costs in",curr)
}

# Setter for above
setCurrencyTitle <- function(){
  curr <- readline("What is the currency being used in this spreadhseet?\n(If it's Euro just press enter)")
  if ( str_length(curr) == 0){
    currencyTitle()
  }
  else{
    currencyTitle(curr)
  }
}

# Information on basic room and per diem costs
livingCosts <- function(room,dailyFood){
  list(room=room,dailyFood=dailyFood)
}

# Setter for above
setLivingCosts <- function(){
  room <- as.integer(readline("How much does a single room cost?"))
  dailyFood <- as.integer(readline("What is the per diem rate?"))
  livingCosts(room,dailyFood)
}

# Create data frame based on all of the data. 
buildDataFrame <- function( allPeople, allSupport, currency, oCosts, lCosts){
  peopleDF <- buildPeopleDataFrame( allPeople, currency)
  nRows <- computeNumRows(allPeople)
  nCols <- computeNumCols(allSupport)
  supportDF <- buildSupportDataFrame(allSupport,nRows) 
  peopleSupportDF <- cbind.data.frame(peopleDF,supportDF)
  psDF <- addExcelMacros(peopleSupportDF,nRows)
  summaryDF <- buildSummaryDF(lCosts,oCosts,nRows,dim(psDF)[2],allSupport)
  colnames(summaryDF) <- colnames(psDF)
  finalDFNoMacros <- rbind(psDF,summaryDF)
  colsAddMacros  <- c(5:dim(psDF)[2])
  finalDF <- addMacros(finalDFNoMacros,colsAddMacros)
  return(finalDF)
}

# Function to replace a column with equivalent formula entries
# df is the data frame and n a vector of columns to replace
addMacros <- function(df,n){
  
  for ( col in n){
    thisCol <- df[,col]
    updatedCol <- sapply(thisCol,function(x){
        if ( is.na(x) ){
          '=" "'
        }
        else{
          paste("=",x,sep="")}
      })
    df[,col] <- xl_formula(updatedCol)
  }
  
  return(df)
}

# Add columns of people+support DF relevant macros (total food, accomodation)
addExcelMacros <- function(psDF,nRows){
  roomData <- nRows + 13
  foodData <- nRows + 14

  for ( i in 1:nRows){
    if (!is.na(psDF[i,2])){
      psDF[i,7] <- multiplyCells(LETTERS[4],foodData,LETTERS[5],i+1)
      psDF[i,8] <- multiplyCells(LETTERS[4],roomData,LETTERS[5],i+1)
    }
  }
  
  return(psDF)
}

# Create data frame template with people data in it. 
# It doesn't have macro information in it

buildPeopleDataFrame <- function(allPeople,currency){
  nRows <- computeNumRows(allPeople)
  emptyStrings <- rep("",nRows)
# Nine colums of data  
#  templateDF <- data.frame(V1=emptyStrings,V2=emptyStrings,
#                           V3=emptyStrings,V4=emptyStrings,
#                           V5=emptyStrings,V6=emptyStrings,
#                           V7=emptyStrings,V8=emptyStrings,
#                           V9=emptyStrings,stringsAsFactors = FALSE)
  templateDF <- data.frame(matrix(NA,nrow=nRows,ncol=9))
  colnames(templateDF) <- c(currency, "Instructors", "Location", "Support required",
                            "Nights staying","Travel","Food","Accomodation","Honorarium")
  
# Find all the course-related roles
  roles <- sapply(allPeople,function(r){
    return(r$Course)
  })
  
  courses <- which(roles)
  nonCourses <- which(!roles)
  orderedRoles <- c(courses,nonCourses)
  thisRow <- 1
  
  templateDF[thisRow,1] <- "Funding committed"
  thisRow <- increment(thisRow) 
  for ( role in names(orderedRoles) ){
    templateDF[thisRow,1] <- role
    for ( personName in setdiff(names(allPeople[[role]]),"Course")){
      x <- allPeople[[role]][[personName]]
      templateDF[thisRow,2] <- personName
      templateDF[thisRow,3] <- x$location
      templateDF[thisRow,5] <- as.integer(x$nights)
      templateDF[thisRow,6] <- as.integer(x$travel)
      thisRow <- increment(thisRow)
    }
    thisRow <- increment(thisRow)
  }
  return(templateDF)
  
}

increment <- function(n){
  return(n+1)
}

# Create data frame template with support data in it. 
# It doesn't have macro information in it
buildSupportDataFrame <- function(allSupport,nRows){
  nCols <- computeNumCols(allSupport)
  supportDF <- data.frame(matrix(NA, nrow=nRows, ncol=nCols))
#  for ( i in c(1:nRows)){
#    for ( j in c(1:nCols)){
#      supportDF[i,j] <- ""
#    }
#  }
  colnames(supportDF) <- sapply(allSupport,function(a){a$name})
  supportDF[1,] <- sapply(allSupport,function(a){a$amount})
  
  return(supportDF)
}

# Compute total number of rows required for all the people data
# Each role gets a row for each person plus an additional space for readability
computeNumRows <- function(allPeople){
  nRoles <- length(allPeople) + 1
# + 1 to allow for row for committed funds
  nRows <- nRoles + sum(sapply(allPeople,function(a){length(a) - 1}))
  return(nRows)
}

# Return number of columns for support
computeNumCols <- function(allSupport){
  return(length(allSupport))
}


# This creates a data frame with all the summary data at the bottom of the spreadsheet
# living is a list with the information about per diems etc.
# other has dinner etc expenses
# nRows is the total number of rows in the total data frame (people+support) so far
# nCols is the _total_ number of columsn in the total data frame (people+support) so far

buildSummaryDF <- function(living,other,nRows,nCols,allSupportInfo){
  nR <- 13
  summaryDF <- data.frame(matrix(NA,nrow=nR,ncol=nCols)) 
#  for ( i in 1:dim(summaryDF)[1]){
#    for ( j in 1:dim(summaryDF)[2]){
#      summaryDF[i,j] <- " "
#    }
#  }
  if (nCols > length(LETTERS)){
    stop("Code cannot deal with columns that are of the form AA etc. talk to Hugh about this!")
  }
  nCP <- nCols - computeNumCols(allSupportInfo)
  
  # Fixed items
  summaryDF[2,1] <- "Sub totals"
  summaryDF[5,1] <- "Other expenses"
  summaryDF[5,2] <- "Refreshments"
  summaryDF[6,2] <- "Dinner"
  summaryDF[12,3] <- "Single Room Accomodation"
  summaryDF[13,3] <- "Per Diem"
  summaryDF[12,4] <- as.integer(living$room)
  summaryDF[13,4] <- as.integer(living$dailyFood)
  summaryDF[8,4] <- "Sub-totals = "
  summaryDF[9,4] <- "Total Expense"
  summaryDF[10,4] <- "Total Budget"
  summaryDF[11,4] <- "Budget Result"
  
  # travel
  summaryDF[8,6] <- sumVerticalRange(LETTERS[6],3,nRows+1)
  # food (including extras)
  summaryDF[8,7] <- sumVerticalRange(LETTERS[7],3,nRows+7)
  summaryDF[5,7] <- other$refreshments
  summaryDF[6,7] <- other$dinner
  
  # Accomodation
  summaryDF[8,8] <- sumVerticalRange(LETTERS[8],3,nRows+1)
  # Honorarium
  summaryDF[8,9] <- sumVerticalRange(LETTERS[9],3,nRows+1)
  
  # Total expenses
  summaryDF[9,9] <- paste("0-",sumHorizontalRange(LETTERS[6],LETTERS[9],nRows+9),sep="")
  # Total Budget
  summaryDF[10,9] <- sumHorizontalRange(LETTERS[nCP+1],LETTERS[nCols],2)
  # Budget remainder
  summaryDF[11,9] <- sumVerticalRange(LETTERS[9],nRows+10,nRows+11)

  # Running total of where contributions are being spent
  for ( i in (nCP+1):nCols ) {
    summaryDF[2,i] <- sumVerticalRange(LETTERS[i],3,nRows+1)
  }
  
  return(summaryDF)
  
}

# This creates a string for an excel range along column
createVerticalRange <- function(letter,low,high){
  paste(letter,low,":",letter,high,sep="")
}

# Sum along column range
sumVerticalRange <- function(letter,low,high){
  paste("sum(",createVerticalRange(letter,low,high),")",sep="")
}

# This creates a string for an excel range along column
createHorizontalRange <- function(lowLetter,highLetter,column){
  paste(lowLetter,column,":",highLetter,column,sep="")
}

# Sum along row range
sumHorizontalRange <- function(lowLetter,highLetter,column){
  paste("sum(",createHorizontalRange(lowLetter,highLetter,column),")",sep="")
}

# Return formula to compute product of two cells
multiplyCells <- function(letter1,column1,letter2,column2){
  paste(letter1,column1,"*",letter2,column2,sep="")
}
