library("readxl")
library("writexl")
library("lubridate")
library("stringr")


xl_formula("=B2+B3")
df <- data.frame(x=c(1,2,3),y=c(4,5,6),z=xl_formula(c("=A1+B1","=A2+B2","=A3+B3")))
getwd()
xl_formula(c("=A1+B1","=A2+B2","=A3+B3"))
write_xlsx(df,"test.xlsx")


instructorInfo <- function(name,course, supportRequired=TRUE, 
                           location, nights,honorarium=FALSE){
  list(name=name,course=course,supportRequired=supportRequired,
            location=location,nights=nights,honorarium=honorarium)
}

supportInfo <- function(name,amount=0){
  list(name=name,amount=amount)
}

otherCosts <- function(refreshments,dinner){
  list(refreshments=refreshments,dinner=dinner)
}

currency <- function(curr="Euro"){
  paste("All costs in",curr)
}

livingCosts <- function(room,dailyFood){
  list(room=room,dailyFood)
}