#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

library(shiny)
myDir <- getwd()
source(paste(myDir,"setupSpreadsheet.R",sep="/"))

# Define UI for application that draws a histogram
ui <- fluidPage(
  titlePanel("budgetBuild"),
  
  sidebarLayout(
    sidebarPanel(
      helpText("Enter role details."),
      
      selectInput("role", 
                  label = "Choose a course/role",
                  choices = c("Open Science", "Carpentry", "Computational Infrastructures", 
                              "Information Security", "Research Data Management",
                              "Analysis","Visualisation","Author Carpentry",
                              "Helpers","Organisers", "Photographer","Other"),
                  selected = "Other"),
      
      textInput("name", "Name of person", 
                value = ""),
      
      sliderInput("range", 
                  label = "Range of interest:",
                  min = 0, max = 100, value = c(0, 100))
      ),
    
    mainPanel(
      textOutput("selected_role")
    )
  )  
)

# Define server logic required to draw a histogram
server <- function(input, output) {

  
  # if ( !exists(allPeople) ){
  #   allPeople <- list()
  # }
  
  output$selected_role <- renderText({ 
    paste("You have selected", getwd())
          #input$role)
  }) 
}

# Run the application 
shinyApp(ui = ui, server = server)

