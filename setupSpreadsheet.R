library("readxl")
library("writexl")
library("stringr")


#xl_formula("=B2+B3")
#df <- data.frame(x=c("x+1", 1,2,3),y=c("y+1",4,5,6),z=c("z+1","","",""))

#df[,3]    <- xl_formula(c("=A1+B1","=A2+B2","=A3+B3","=A4+B4")


#df[1,3] <- "z+1"
#xl_formula(c("=A1+B1","=A2+B2","=A3+B3"))
#write_xlsx(df,col_names=FALSE, "test.xlsx")
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
      nPeople <- as.integer(readline(paste("How many people are going to be in the role",
                                           thisRole," ")))
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
  n <- as.integer(readline("How many organisations are contributing? "))
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
  psDF <- addExcelMacros(peopleSupportDF,nRows,nCols)
  summaryDF <- buildSummaryDF(living,other,nRows,nCols)
  colnames(summaryDF) <- colnames(peopleSupportDF)
  finalDF <- rbind(peopleSupportDF,summaryDF)
  return(finalDF)
}

# Create data frame template with people data in it. 
# It doesn't have macro information in it

buildPeopleDataFrame <- function(allPeople,currency){
  nRows <- computeNumRows(allPeople)
  emptyStrings <- rep("",nRows)
# Nine colums of data  
  templateDF <- data.frame(V1=emptyStrings,V2=emptyStrings,
                           V3=emptyStrings,V4=emptyStrings,
                           V5=emptyStrings,V6=emptyStrings,
                           V7=emptyStrings,V8=emptyStrings,
                           V9=emptyStrings,stringsAsFactors = FALSE)
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
      templateDF[thisRow,5] <- x$nights
      templateDF[thisRow,6] <- x$travel
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
  for ( i in c(1:nRows)){
    for ( j in c(1:nCols)){
      supportDF[i,j] <- ""
    }
  }
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
  for ( i in 1:dim(summaryDF)[1]){
    for ( j in 1:dim(summaryDF)[2]){
      summaryDF[i,j] <- " "
    }
  }
  if (nCols > length(LETTERS)){
    stop("Code cannot deal with columns that are of the form AA etc. talk to Hugh about this!")
  }
  nCP <- nCols - computeNumCols(allSupportInfo)
  
  # Fixed items
  summaryDF[2,1] <- "Sub totals"
  summaryDF[5,1] <- "Other expenses"
  summaryDF[5,2] <- "Refreshments"
  summaryDF[6,2] <- "Dinner"
  summaryDF[12,1] <- "Single Room Accomodation"
  summaryDF[13,1] <- "Per Diem"
  summaryDF[12,2] <- living$room
  summaryDF[13,2] <- living$dailyFood
  summaryDF[8,5] <- "Sub-totals = "
  
  # formulae (even when they aren't)
  blanks <- rep("=",nR)
  travelCol <- blanks
  travelCol[8] <- sumVerticalRange(LETTERS[6],2,nRows)
  summaryDF[,6] <- xl_formula(travelCol)
  
  foodCol <- blanks
  foodCol[8] <- sumVerticalRange(LETTERS[7],2,nRows+6)
  foodCol[5] <- paste("=",other$refreshments)
  foodCol[6] <- paste("=",other$dinner)
  summaryDF[,7] <- xl_formula(foodCol)
  
  accomodationCol <- blanks
  accomodationCol[8] <- sumVerticalRange(LETTERS[8],2,nRows)
  accomodationCol[9] <- '="Total Expense"'
  accomodationCol[10] <- '="Total Budget"'
  accomodationCol[11] <- '="Budget Result"'
  summaryDF[,8] <- xl_formula(accomodationCol)
  
  totalsCol <- blanks
  totalsCol[8] <- sumVerticalRange(LETTERS[9],2,nRows)
  totalsCol[9] <- sumHorizontalRange(LETTERS[6],LETTERS[9],nRows+8)
  totalsCol[10] <- sumHorizontalRange(LETTERS[nCP+1],LETTERS[nCols],2)
  totalsCol[11] <- sumVerticalRange(LETTERS[10],nRows+9,nRows+10)
  summaryDF[,9] <- xl_formula(totalsCol)
  
  for ( i in (nCP+1):nCols ) {
    x <- blanks
    x[2] <- sumVerticalRange(LETTERS[i],2,nRows)
    summaryDF[,i] <- xl_formula(x)
  }
  
  return(summaryDF)
  
}

# This creates a string for an excel range along column
createVerticalRange <- function(letter,low,high){
  paste(letter,low,":",letter,high,sep="")
}

# Sum along column range
sumVerticalRange <- function(letter,low,high){
  paste("=sum(",createVerticalRange(letter,low,high),")")
}

# This creates a string for an excel range along column
createHorizontalRange <- function(lowLetter,highLetter,column){
  paste(lowLetter,column,":",highLetter,column,sep="")
}

# Sum along row range
sumHorizontalRange <- function(lowLetter,highLetter,column){
  paste("=sum(",createHorizontalRange(lowLetter,highLetter,column),")")
}
