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

  name <- readline("What is the name of the supporting organisation? \nPress enter to finish ")
  if (str_length(name) == 0){
    return -1 
  }
  else{
    a <- readline("How much have they committed (just put in 0 if you don't know yet) ")
    amount <- as.integer(a)
    return(supportInfo(name,amount))
  }
}

# Get all supporting organisations
setAllSupportInfo <- function(){
  allSupport <- list()
  x <- setSupportInfo()
  while( !is.integer(x) ){
    allSupport[[length(allSupport)+1]] <- x
    x <- setSupportInfo()
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
  otherCosts(room,dailyFood)
}

# Create data frame based on all of the data. 
buildDataFrame <- function( allPeople, allSupport, currency, oCosts, lCosts){
  peopleDF <- buildPeopleDataFrame( allPeople)
  nRows <- computeNumRows(allPeople)
  nCols <- computeNumCols(allSupport)
  supportDF <- buildSupportDataFrame(allSupport,nRows) 
  peopleSupportDF <- cbind.data.frame(peopleDF,supportDF)
  summaryDF <- buildSummaryDF(living,other,nRows,nCols)
  templateDF <- rbind(peopleSupportDF,summaryDF)
  finalDF <- addExcelMacros(templateDF,nRows,nCols)
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
  
  for ( role in names(orderedRoles) ){
    templateDF[thisRow,1] <- role
    for ( personName in setdiff(names(allPeople[[role]]),"Course")){
      x <- allPeople[[role]][[personName]]
      templateDF[thisRow,2] <- personName
      templateDF[thisRow,3] <- x$location
      templateDF[thisRow,5] <- x$nights
      templateDF[thisRow,6] <- x$travel
      thisRow <- thisRow + 1
    }
    thisRow <- thisRow + 1
  }
  return(templateDF)
  
}

# Create data frame template with support data in it. 
# It doesn't have macro information in it
buildSupportDataFrame <- function(allSupport,nRows){
  nCols <- computeNumCols(allSupport)
  
}

# Compute total number of rows required for all the people data
# Each role gets a row for each person plus an additional space for readability
computeNumRows <- function(allPeople){
  nRoles <- length(allPeople)
  nRows <- nRoles + sum(sapply(allPeople,function(a){length(a) - 1}))
  return(nRows)
}


