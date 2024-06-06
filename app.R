library(shiny)
library(shinyjs)
library(DT)
library(data.table)
library(lubridate)
library(shinyalert)
library(readxl)
library(dplyr) 
library(shinyWidgets)
library(openxlsx)
library(shinymanager)
library(shinythemes)
library(bslib)
library(knitr) 


inactivity <- "function idleTimer() {
var t = setTimeout(logout, 18000000);
window.onmousemove = resetTimer; // catches mouse movements
window.onmousedown = resetTimer; // catches mouse movements
window.onclick = resetTimer;     // catches mouse clicks
window.onscroll = resetTimer;    // catches scrolling
window.onkeypress = resetTimer;  //catches keyboard actions

function logout() {
window.close();  //close the window
}

function resetTimer() {
clearTimeout(t);
t = setTimeout(logout, 18000000);  // time is in milliseconds (1000 is 1 second)
}
}
idleTimer();"


# data.frame with credentials info
credentials <- data.frame(
  user = c("gluon","vivian", "paris"),
  password = c(scrypt::hashPassword("gluon1235!!"), scrypt::hashPassword("vivian99"), scrypt::hashPassword("Paris5682!")),
  is_hashed_password = TRUE,
  stringsAsFactors = FALSE
)


maker_order_page <- tabPanel("廠商訂購單",fluidRow(
        tags$style("body {
    -moz-transform: scale(0.75, 0.75); /* Moz-browsers */
    zoom: 0.75; /* Other non-webkit browsers */
    zoom: 75%; /* Webkit browsers */
}
              "),
  column(3, 
         style = "margin-bottom: 20px;",
         selectizeInput("no", "輸入產品編號", choices = NULL, multiple = FALSE, options = list(placeholder = "Type to search...", create  = TRUE)),
         actionButton("reset_po_data","Reset")),
  #actionButton("add_adjust_po_data","新增或修改"))
  
  column(12,uiOutput("no_data"))))

quotaion_page <- tabPanel("客戶報價單")
  



ui <- secure_app(head_auth = tags$script(inactivity),
                 navbarPage(
                   theme = shinytheme("cerulean"),
                   title=div(img(src="gluon_logo.png", "系統"),
                             height = "10px%",
                             width = "10px%",
                             style = "position: relative; margin:-15px 0px"),
                   #navbarMenu(title = "成本報價預估單")
                   #tabPanel("","conten1"),
                   #tabPanel("", "content2")
                   #id = "navbar",
                   
                   maker_order_page,
                   
                   quotaion_page
                   
                 ))

server <- function(input, output, session) {
  
  result_auth <- secure_server(check_credentials = check_credentials(credentials))
  
  output$res_auth <- renderPrint({
    reactiveValuesToList(result_auth)
  })
  
  source("helper/maker_po_server.R", local = TRUE)
}






shinyApp(ui = ui, server = server)


















