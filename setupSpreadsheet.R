library("readxl")
library("writexl")
library("stringr")


xl_formula("=B2+B3")
df <- data.frame(x=c(1,2,3),y=c(4,5,6),z=xl_formula(c("=A1+B1","=A2+B2","=A3+B3")))
getwd()
xl_formula(c("=A1+B1","=A2+B2","=A3+B3"))
write_xlsx(df,"test.xlsx")


# Base function to set up information about an instructor/helper/other

instructorInfo <- function(name, supportRequired=TRUE, 
                           location, nights,travel, honorarium=0){
  list(name=name,supportRequired=supportRequired,
            location=location,travel=travel,nights=nights,honorarium=honorarium)
}

# Setter function for instructorInfo
setInstructorInfo <- function(){
  
  name <- readline("What is the name of the person?\n(If you are done just press enter)")
  if (str_length(name) == 0){
    return -1 
  }
  else{
    location <- readline("What institution are they coming from?")
    n <- readline("How many nights are they staying?")
    nights <- as.integer(n) 
    t <- readline("What is the estimate of their travel expenses?")
    travel <- as.integer(t)
    support <- readline("We assume that they will need support - if not, type N here")
    if ( str_detect(support,"N")  ){
      supportRequired = FALSE
    }
    else{
      supportRequired = TRUE
    }
      
    honorarium <- readline("If they require an honorarium put the number here.")
    if ( str_length(honorarium) == 0  ){
      honorarium <- 0
    }
    else{
      honorarium <- as.integer(honorarium)
    }
    instructorInfo(name=name,supportRequired=supportRequired, 
                   location=location, nights=nights,travel=travel, honorarium=honorarium)
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
      print(paste("Getting information for",nPeople," individuals on",thisRole))
      x <- list()
      x$Course <- thisRole %in% courses
      for ( i in c(1:nPeople) ){
        print(paste("Enter data for individual",i,"for",thisRole))
        thisPerson <- setInstructorInfo()
        if ( is.integer(thisPerson)){
          break
        }
        x[thisPerson$name] <- thisPerson
      }
      allPeople[thisRole] <- x
    }
  }
  allPeople
}
  

supportInfo <- function(name,amount=0){
  list(name=name,amount=amount)
}

setSupportInfo <- function(){

  name <- readline("What is the name of the supportng organisation?\n(If you are done just press enter)")
  if (str_length(name) == 0){
    return -1 
  }
  else{
    a <- readline("How much have they committed (just put in 0 if you don't know yet)")
    amount <- as.integer(a)
    supportInfo(name,amount)
  }
}

otherCosts <- function(refreshments,dinner){
  list(refreshments=refreshments,dinner=dinner)
}

setOtherCosts <- function(){
  
  refreshments <- as.integer(readline("What is the estimated refreshments cost?"))
  dinner <- as.integer(readline("What is the esimated costs of the school dinner?"))
  otherCosts(refreshments,dinner)

}



currency <- function(curr="Euro"){
  paste("All costs in",curr)
}

setCurrency <- function(){
  curr <- readline("What is the currency being used in this spreadhseet?\n(If it's Euro just press enter)")
  if ( str_length(curr) == 0){
    currency()
  }
  else{
    currency(curr)
  }
}

livingCosts <- function(room,dailyFood){
  list(room=room,dailyFood=dailyFood)
}

setLivingCosts <- function(){
  room <- as.integer(readline("How much does a single room cost?"))
  dailyFood <- as.integer(readline("What is the per diem rate?"))
  otherCosts(room,dailyFood)
}


