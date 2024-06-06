# Set options to prevent scientific notation

options(scipen = 999)



load_maker_data <- function(){
  suppressWarnings({ 
    
    # maker PO data
    filename <- read_excel("Data/Maker_PO.xlsx", col_types = c("text", "text", "text", "numeric", "numeric",  "text","numeric","text", "text","text","text", "text", "text", "text","text", "text"))
 
    
    # Define the origin date for Excel's numeric date system
    origin_date <- as.Date("1899-12-30")
    
    # Loop through each element in the "Shipment_date" column
    for (i in seq_along(filename$Shipment_Date)) {
      elem <- filename$Shipment_Date[i]
      
      if (!is.na(as.numeric(elem))) {
        filename$Shipment_Date[i] <- gsub("^0", "", gsub("/0", "/", format(origin_date + as.numeric(elem), "%Y/%m/%d")), perl = TRUE)
      } else if(is.na(as.numeric(elem))){
        filename$Shipment_Date[i] <- elem
      }
    }
    
    for (i in seq_along(filename$日期)) {
      elem <- filename$日期[i]
      
      if (!is.na(as.numeric(elem))) {
        filename$日期[i] <- gsub("^0", "", gsub("/0", "/", format(origin_date + as.numeric(elem), "%Y/%m/%d")), perl = TRUE)
      } else if(is.na(as.numeric(elem))){
        filename$日期[i] <- elem
      }
    }
    filename[is.na(filename)] <- ""
    
    #maker_po_data <- maker_po_data[, c(1:15)]
    filename <- filename[order(filename$maker_order_number), ]
    return(filename)
  }) 
  
  
  
}
 


maker_name <- read_excel("Data/廠商資料.xlsx", sheet = 1)
# Select only the 2nd and 3rd columns
maker_name <- maker_name[, c(2, 3)]


# maker po
maker_po_data <- reactiveValues(data = NULL)
  
# selected data
selected_data<-reactiveValues(data = NULL)
  
observeEvent(input$no, {
  maker_po_data$data <- load_maker_data()
  original_choices_no <- unique(maker_po_data$data$maker_order_number)
    
  if (input$no ==""){
    updateSelectizeInput(session, "no", choices = c("", original_choices_no))
      
      
    output$no_data<-renderUI({
      fluidPage(
        column(12,dataTableOutput("no_data_dt"))
      ) 
    })
      
      
      
    #### render DataTable part ####
    output$no_data_dt<-renderDataTable({
      DT=maker_po_data$data
      datatable(DT,selection = 'single',
                escape=F, 
                options = list(autoWidth = TRUE,
                               scrollX = TRUE,
                               scrollY = "500px",
                               columnDefs = list(list(width = '200px', targets = c(1,2,11,15)),
                                                 list(width = '150px', targets = c(3,12,13)),
                                                 list(width = '90px', targets = c(4,8, 14)),
                                                 list(width = '50px', targets = c(6,7))),
                               lengthChange = FALSE,
                               paging = FALSE)) })
      
      
      
      
  }
  else if (input$no %in% maker_po_data$data$maker_order_number) {
    updateSelectizeInput(session, "no", selected = input$no)
    no <- input$no
    selected_data$data<- maker_po_data$data[grepl(no,maker_po_data$data$maker_order_number), ]
    output$no_data<-renderUI({
      fluidPage(
        column(6,
               tags$head(
                 tags$style(HTML("
      .modal-dialog {
        overflow-y: auto;
      }"))
               ),
               HTML('<div class="btn-group" role="group" aria-label="Basic example" style = "padding:10px">'),
               div(style="display:inline-block;",actionButton(inputId = "Add_row_head",label = "Add")),
               div(style="display:inline-block;",actionButton(inputId = "mod_row_head",label = "Edit") ),
               div(style="display:inline-block;",actionButton(inputId = "Del_row_head",label = "Delete") ),
               ### Optional: a html button 
               # HTML('<input type="submit" name="Add_row_head" value="Add">'),
               HTML('</div>')),
        
        column(12,dataTableOutput("no_data_dt")),
        tags$script("$(document).on('click', '#no_data_dt button', function () {
                 Shiny.onInputChange('lastClickId',this.id);
                 Shiny.onInputChange('lastClick', Math.random()) });"),
        column(3, actionButton(inputId = "update_maker_po",label = "Save")),
        column(12, 
                 
               div(style = "text-align: center;", # Align to the right
                   uiOutput("download_button_maker_po")
               )
        )
          
          
          
      ) 
    })
      
      
      
    #### render DataTable part ####
    output$no_data_dt<-renderDataTable({
      DT=selected_data$data
      datatable(DT,selection = 'single',
                escape=F, 
                options = list(autoWidth = TRUE,
                               scrollX = TRUE,
                               columnDefs = list(list(width = '200px', targets = c(1,2,11,15)),
                                                 list(width = '150px', targets = c(3,12,13)),
                                                 list(width = '90px', targets = c(4,8, 14)),
                                                 list(width = '50px', targets = c(6,7))),
                               lengthChange = FALSE,
                               paging = FALSE
                )) })
      
  }
    
    
    
    
  else {
    selected_data$data<- data.frame(matrix(nrow = 0, ncol = length(colnames(maker_po_data$data))))
    # Set column names
    colnames(selected_data$data) <- colnames(maker_po_data$data)
    output$no_data<-renderUI({
      fluidPage(
        column(6,
               tags$head(
                 tags$style(HTML("
      .modal-dialog {
        overflow-y: auto;
      }
    "))
               ),
               HTML('<div class="btn-group" role="group" aria-label="Basic example" style = "padding:10px">'),
               div(style="display:inline-block;",actionButton(inputId = "Add_row_head",label = "Add")),
               div(style="display:inline-block;",actionButton(inputId = "mod_row_head",label = "Edit") ),
               div(style="display:inline-block;",actionButton(inputId = "Del_row_head",label = "Delete") ),
               ### Optional: a html button 
               # HTML('<input type="submit" name="Add_row_head" value="Add">'),
               HTML('</div>')),
          
        column(12,dataTableOutput("no_data_dt")),
        tags$script("$(document).on('click', '#no_data_dt button', function () {
                 Shiny.onInputChange('lastClickId',this.id);
                 Shiny.onInputChange('lastClick', Math.random()) });"),
        column(3, actionButton(inputId = "update_maker_po",label = "Save")),
        column(12, 
               
               div(style = "text-align: center;", # Align to the right
                   uiOutput("download_button_maker_po")
               )
        )
        
       
          
      ) 
    })
      
      
      
    #### render DataTable part ####
    output$no_data_dt<-renderDataTable({
      DT=selected_data$data
      datatable(DT,selection = 'single',
                escape=F, 
                options = list(autoWidth = TRUE,
                               scrollX = TRUE,
                               columnDefs = list(list(width = '200px', targets = c(1,2,11,15)),
                                                 list(width = '150px', targets = c(3,12,13)),
                                                 list(width = '90px', targets = c(4,8, 14)),
                                                 list(width = '50px', targets = c(6,7))),
                               lengthChange = FALSE,
                               paging = FALSE
               )) })
      
      
   }
    
}) 
  
  
  
observeEvent(input$Add_row_head, ignoreInit = TRUE, {
  ### This is the pop up board for input a new row
    
  if(nrow(selected_data$data) != 0){
    showModal(modalDialog(title = "新增內容",
                          selectizeInput(paste0("name_add", input$Add_row_head), "名稱",choices= c("",unique(maker_po_data$data$名稱)), multiple = TRUE ,options = list(create  = TRUE, delimiter = '?')),
                          selectizeInput(paste0("model_add", input$Add_row_head), "Model",choices= c("",unique(maker_po_data$data$Model)), options = list(create  = TRUE)),
                          numericInput(paste0("price_add", input$Add_row_head), "價格",value=""),
                          numericInput(paste0("quan_add", input$Add_row_head), "數量",value=""),
                          selectizeInput(paste0("unit_add", input$Add_row_head), "單位",choices= c("",unique(maker_po_data$data$單位)), options = list(create  = TRUE)),
                          helpText(HTML("<span style='color: blue;'>如果是折扣，請寫負數價格，數量請寫1，單位空白</span>")),
                          br(),
                          actionButton("go", "Add") ))
    
      
  }
  else{
      
    showModal(modalDialog(title = "新增內容",
                          selectizeInput(paste0("name_add", input$Add_row_head), "名稱",choices= c("",unique(maker_po_data$data$名稱)),  multiple = TRUE,options = list(create = TRUE, delimiter = '?')),
                          selectizeInput(paste0("model_add", input$Add_row_head), "Model",choices= c("",unique(maker_po_data$data$Model)), options = list(create  = TRUE)),
                          numericInput(paste0("price_add", input$Add_row_head), "價格",value=''),
                          numericInput(paste0("quan_add", input$Add_row_head), "數量",value=''),
                          selectizeInput(paste0("unit_add", input$Add_row_head), "單位",choices= c("",unique(maker_po_data$data$單位)), options = list(create  = TRUE)),
                          helpText(HTML("<span style='color: blue;'>如果是折扣，請寫負數價格，數量請寫1，單位空白</span>")),
                          br(),
                          numericInput(paste0("tax_add", input$Add_row_head), "Tax",value=0, min =0 , max =1),
                          helpText(HTML("<span style='color: blue;'>範例：5%，請填0.05</span>")),
                          br(),
                          selectizeInput(paste0("cur_add", input$Add_row_head), "幣別",choices= c("",unique(maker_po_data$data$幣別)), options = list(create  = TRUE)),
                          selectizeInput(paste0("company_add", input$Add_row_head), "我方公司",choices= c("",unique(maker_po_data$data$我方公司)), options = list(create  = TRUE)),
                          selectizeInput(paste0("maker_add", input$Add_row_head), "Maker", choices= c("", unique(c(maker_po_data$data$Maker, maker_name$廠商簡稱))), options = list(create  = TRUE)),
                          selectizeInput(paste0("end_add", input$Add_row_head), "End User", choices= c("", unique(maker_po_data$data$End_User)) ,options = list(create  = TRUE)),
                          selectizeInput(paste0("pt_add", input$Add_row_head), "Payment Term", choices= c("", unique(maker_po_data$data$Payment_Terms)), selected ="", multiple = TRUE, options = list(create  = TRUE, delimiter = '?')),
                          selectizeInput(paste0("des_add", input$Add_row_head), "Destination", choices= c("", unique(maker_po_data$data$Destination)), selected = "", options = list(create  = TRUE)),
                          selectizeInput(paste0("ship_add", input$Add_row_head), "Shipment Date", choices= c("", unique(maker_po_data$data$Shipment_Date)), selected = "",options = list(create  = TRUE)),
                          dateInput(paste0("date_add", input$Add_row_head), "日期", value = Sys.Date(), format = "yyyy/mm/dd"),
                          textAreaInput(paste0("remark_add", input$Add_row_head), "Remark"), 
                          helpText(HTML("<span style='color: blue;'>範例：1. 型式：OG-250H(RoHS) 2. 客先：PSMC <br>(數字後面請加小數點及空格)</span>")),
                          actionButton("go", "Add"),
                          style = "overflow-y: auto;  vertical-align: middle; height: 800px;"))
  
      
  }
    
    
 })
  
  
### Add a new row to DT  
observeEvent(input$go, {
    
  if (!is.numeric(input[[paste0("price_add", input$Add_row_head)]]) || 
      is.na(input[[paste0("price_add", input$Add_row_head)]]) || 
      !is.numeric(input[[paste0("quan_add", input$Add_row_head)]]) || 
      is.na(input[[paste0("quan_add", input$Add_row_head)]])) {
        sendSweetAlert(
        session = session,
        title = "Notice",
        text = "請填寫價格與數量",
        type = "error"
        )
       } 
    
    
    
  else{
      
    if(nrow(selected_data$data) != 0){
      
       name_add <- input[[paste0("name_add", input$Add_row_head)]]
      if (!is.null(name_add) && length(name_add) > 0) {
        merged_names <- paste(name_add, collapse = " / ")
      } else {
        merged_names <- ""
      }
    
      new_row = data.frame(
        maker_order_number = input$no,
        名稱 = merged_names,
        Model = input[[paste0("model_add", input$Add_row_head)]],
        價格 = input[[paste0("price_add", input$Add_row_head)]],
        Quantity = input[[paste0("quan_add", input$Add_row_head)]],
        單位 = input[[paste0("unit_add", input$Add_row_head)]],
        Tax = selected_data$data$Tax[1],

          
        # hide
        幣別 = selected_data$data$幣別[1],
        我方公司 = selected_data$data$我方公司[1],
        Maker = selected_data$data$Maker[1],
        End_User = selected_data$data$End_User[1],
        Payment_Terms = selected_data$data$Payment_Terms[1],
        Destination = selected_data$data$Destination[1],
        Shipment_Date = selected_data$data$Shipment_Date[1],
        日期 = selected_data$data$日期[1],
        Remark = selected_data$data$Remark[1]
          
          
      )
      selected_row=input$no_data_dt_rows_selected
      
      

      if(!is.null(selected_row)&&nrow(selected_data$data)!=selected_row){
        upper_row <- selected_data$data[1:selected_row,]
        lower_row <- selected_data$data[-c(1:selected_row),]
        selected_data$data<-rbind(upper_row,
                                  new_row,
                                  lower_row)
        

      }
    
        
    else{
      selected_data$data<-rbind(selected_data$data,new_row )
    }
    }
      
    else{
      if(!is.numeric(input[[paste0("tax_add", input$Add_row_head)]]) || 
         is.na(input[[paste0("tax_add", input$Add_row_head)]])||
         input[[paste0("tax_add", input$Add_row_head)]]<0 || 
         input[[paste0("tax_add", input$Add_row_head)]]>1){
           sendSweetAlert(
             session = session,
             title = "Notice",
             text = "請填寫Tax，並且介於0至1之間",
             type = "error"
            )
          }
      else{
        
        name_add <- input[[paste0("name_add", input$Add_row_head)]]
        if (!is.null(name_add) && length(name_add) > 0) {
          merged_names <- paste(name_add, collapse = " / ")
        } else {
          merged_names <- ""
        }
        
        pt_add <- input[[paste0("pt_add", input$Add_row_head)]]
        if (!is.null(pt_add) && length(pt_add) > 0) {
          merged_pt <- paste(pt_add, collapse = " / ")
        } else {
          merged_pt <- ""
        }
        
        
        
        new_row = data.frame(
          maker_order_number = input$no,
          名稱 = merged_names,  
          Model = input[[paste0("model_add", input$Add_row_head)]],
          價格 = input[[paste0("price_add", input$Add_row_head)]],
          Quantity = input[[paste0("quan_add", input$Add_row_head)]],
          單位 = input[[paste0("unit_add", input$Add_row_head)]],
          Tax =  input[[paste0("tax_add", input$Add_row_head)]],
          幣別 = input[[paste0("cur_add", input$Add_row_head)]],
          
            
          我方公司 = input[[paste0("company_add", input$Add_row_head)]],
          Maker = input[[paste0("maker_add", input$Add_row_head)]],
          End_User = input[[paste0("end_add", input$Add_row_head)]],
          Payment_Terms = merged_pt,
          Destination = input[[paste0("des_add", input$Add_row_head)]],
          Shipment_Date = input[[paste0("ship_add", input$Add_row_head)]],
          日期 = gsub("-", "/", gsub("-0", "/", input[[paste0("date_add", input$Add_row_head)]]), perl = TRUE),
          Remark = input[[paste0("remark_add", input$Add_row_head)]]
         
        )
        selected_data$data<-rbind(selected_data$data,new_row )
        
      }
        
        
        
    }
   
    
    removeModal()
      
  }
    
})
  
  
  
### delete selected rows part
### this is warning messge for deleting
observeEvent(input$Del_row_head,{
    
  showModal(
    if(length(input$no_data_dt_rows_selected)>=1 ){
      modalDialog(
        title = "Warning",
        paste("Are you sure delete this row?" ),
        footer = tagList(
          modalButton("Cancel"),
          actionButton("ok", "Yes")
        ), easyClose = TRUE)
    }else{
      modalDialog(
        title = "Warning",
        paste("Please select row(s) that you want to delete!" ),easyClose = TRUE
      )
    }
      
  )
})
  
### If user say OK, then delete the selected rows
observeEvent(input$ok, {
  selected_data$data=selected_data$data[-input$no_data_dt_rows_selected,]
  removeModal()
})
  
  
### edit button
observeEvent(input$mod_row_head,{
  showModal(
    if(length(input$no_data_dt_rows_selected)>=1 ){
      modalDialog(
        fluidPage(
          h3(strong("Modification"),align="center"),
          hr(),
          uiOutput("row_edit"),
          actionButton("save_changes","Save changes")),
        style = "overflow-y: auto; vertical-align: middle; height: 800px;")
    }
      
    else{
      modalDialog(
        title = "Warning",
        paste("Please select the row that you want to edit!" ),easyClose = TRUE
      )
      }
      
   )
})
  
#### modify part
output$row_edit <- renderUI({
  selected_row=input$no_data_dt_rows_selected
  multiple_name <- unlist(strsplit(selected_data$data$名稱[selected_row], " / "))  
  multiple_pt <- unlist(strsplit(selected_data$data$Payment_Terms[selected_row], " / "))  
  fluidRow(
    textInput(paste0("number_edit", input$mod_row_head), "maker_order_number", value = input$no),
    selectizeInput(paste0("name_edit", input$mod_row_head), "名稱",choices= c(unique(c(maker_po_data$data$名稱, multiple_name))), selected = c(multiple_name), multiple = TRUE,options = list(create  = TRUE, delimiter = '?')),
    selectizeInput(paste0("model_edit", input$mod_row_head), "Model",choices= c(unique(c(maker_po_data$data$Model, selected_data$data$Model[selected_row]))), selected = selected_data$data$Model[selected_row], options = list(create  = TRUE)),
    numericInput(paste0("price_edit", input$mod_row_head), "價格",value=as.numeric(selected_data$data$價格[selected_row])),
    numericInput(paste0("quan_edit", input$mod_row_head), "數量",value=as.numeric(selected_data$data$Quantity[selected_row])),
    numericInput(paste0("tax_edit", input$mod_row_head), "Tax",value=as.numeric(selected_data$data$Tax[selected_row]), min =0, max=1),
    helpText(HTML("<span style='color: blue;'>範例：5%，請填0.05</span>")),
    br(),
    selectizeInput(paste0("unit_edit", input$mod_row_head), "單位", choices= c(unique(c(maker_po_data$data$單位, selected_data$data$單位[selected_row]))), selected = selected_data$data$單位[selected_row], options = list(create  = TRUE)),
    selectizeInput(paste0("cur_edit", input$mod_row_head), "幣別",choices= c(unique(c(maker_po_data$data$幣別, selected_data$data$幣別[selected_row]))), selected = selected_data$data$幣別[selected_row], options = list(create  = TRUE)),
    selectizeInput(paste0("company_edit", input$mod_row_head), "我方公司",choices= c(unique(c(maker_po_data$data$我方公司, selected_data$data$我方公司[selected_row]))), selected = selected_data$data$我方公司[selected_row], options = list(create  = TRUE)),
    selectizeInput(paste0("maker_edit", input$mod_row_head), "Maker", choices= c(unique(c(maker_po_data$data$Maker, maker_name$廠商簡稱,selected_data$data$Maker[selected_row]))), selected = selected_data$data$Maker[selected_row], options = list(create  = TRUE)),
    selectizeInput(paste0("end_edit", input$mod_row_head), "End User", choices= c(unique(c(maker_po_data$data$End_User, selected_data$data$End_User[selected_row]))), selected = selected_data$data$End_User[selected_row], options = list(create  = TRUE)),
    selectizeInput(paste0("pt_edit", input$mod_row_head), "Payment Term", choices= c(unique(c(maker_po_data$data$Payment_Terms, multiple_pt))), selected = c(multiple_pt), multiple = TRUE, options = list(create  = TRUE, delimiter = '?')),

    selectizeInput(paste0("des_edit", input$mod_row_head), "Destination", choices= c(unique(c(maker_po_data$data$Destination, selected_data$data$Destination[selected_row]))), selected = selected_data$data$Destination[selected_row], options = list(create  = TRUE)),
    selectizeInput(paste0("ship_edit", input$mod_row_head), "Shipment Date", choices= c(unique(c(maker_po_data$data$Shipment_Date, selected_data$data$Shipment_Date[selected_row]))), selected = selected_data$data$Shipment_Date[selected_row], options = list(create  = TRUE)),
    dateInput(paste0("date_edit", input$mod_row_head), "日期", value = selected_data$data$日期[selected_row], format = "yyyy/mm/dd"),
    textAreaInput(paste0("remark_edit", input$mod_row_head), "Remark", value =  selected_data$data$Remark[selected_row],rows = 3),
    helpText(HTML("<span style='color: blue;'>範例：1. 型式：OG-250H(RoHS) 2. 客先：PSMC <br>(數字後面請加小數點及空格)</span>")))
  
})
  
### This is to replace the modified row to existing row
  
observeEvent(input$save_changes, {
  if (!is.numeric(input[[paste0("price_edit", input$mod_row_head)]]) || 
      is.na(input[[paste0("price_edit", input$mod_row_head)]] ) || 
      !is.numeric(input[[paste0("quan_edit", input$mod_row_head)]]) || 
      is.na(input[[paste0("quan_edit", input$mod_row_head)]] )) {
        sendSweetAlert(
          session = session,
          title = "Notice",
          text = "請填寫價格與數量",
          type = "error"
         )
      
       }
    
  else if(!is.numeric(input[[paste0("tax_edit", input$mod_row_head)]]) || 
          is.na(input[[paste0("tax_edit", input$mod_row_head)]])||
          input[[paste0("tax_edit", input$mod_row_head)]]<0 || 
          input[[paste0("tax_edit", input$mod_row_head)]]>1){
             sendSweetAlert(
               session = session,
               title = "Notice",
               text = "請填寫Tax，並且介於0至1之間",
               type = "error"
              )
           }
    
  else{
    # Initialize a list to store the edited values
    edited_values <- list()
    selected_row=input$no_data_dt_rows_selected
    
    name_edit <- input[[paste0("name_edit", input$mod_row_head)]]
    if (!is.null(name_edit) && length(name_edit) > 0) {
      merged_names_edit <- paste(name_edit, collapse = " / ")
    } else {
      merged_names_edit <- ""
    }
    
    pt_edit <- input[[paste0("pt_edit", input$mod_row_head)]]
    if (!is.null(pt_edit) && length(pt_edit) > 0) {
      merged_pt_edit <- paste(pt_edit, collapse = " / ")
    } else {
      merged_pt_edit <- ""
    }
    
    
    # Add each edited value to the list
    edited_values$maker_order_number <- input[[paste0("number_edit", input$mod_row_head)]]
    edited_values$名稱 <- merged_names_edit
    edited_values$Model <- input[[paste0("model_edit", input$mod_row_head)]]
    edited_values$價格 <- input[[paste0("price_edit", input$mod_row_head)]]
    edited_values$數量 <- input[[paste0("quan_edit", input$mod_row_head)]]
    edited_values$單位 <- input[[paste0("unit_edit", input$mod_row_head)]]
    edited_values$Tax <- input[[paste0("tax_edit", input$mod_row_head)]]
    edited_values$幣別 <- input[[paste0("cur_edit", input$mod_row_head)]]
    edited_values$我方公司 <- input[[paste0("company_edit", input$mod_row_head)]]
    edited_values$Maker <- input[[paste0("maker_edit", input$mod_row_head)]]
    edited_values$End_User <- input[[paste0("end_edit", input$mod_row_head)]]
    edited_values$Payment_Terms <- merged_pt_edit
    edited_values$Destination <- input[[paste0("des_edit", input$mod_row_head)]]
    edited_values$Shipment_Date <- input[[paste0("ship_edit", input$mod_row_head)]]
    edited_values$日期 <- gsub("-", "/", gsub("-0", "/", input[[paste0("date_edit", input$mod_row_head)]]), perl = TRUE)
    edited_values$Remark <- input[[paste0("remark_edit", input$mod_row_head)]]
    # Convert the list to a data frame
    edited_df <- as.data.frame(edited_values)
    # Replace the corresponding row in selected_data$data with the edited values
    selected_data$data[selected_row, ] <- edited_df
    
    for (i in seq_len(nrow(selected_data$data))) {
      selected_data$data$maker_order_number[i] <- edited_values$maker_order_number
      selected_data$data$Tax[i] <- edited_values$Tax
      selected_data$data$幣別[i] <- edited_values$幣別
      selected_data$data$我方公司[i] <- edited_values$我方公司
      selected_data$data$Maker[i] <- edited_values$Maker
      selected_data$data$End_User[i] <- edited_values$End_User
      selected_data$data$Payment_Terms[i] <- edited_values$Payment_Terms
      selected_data$data$Destination[i] <- edited_values$Destination
      selected_data$data$Shipment_Date[i] <- edited_values$Shipment_Date
      selected_data$data$日期[i] <-  edited_values$日期
      selected_data$data$Remark[i] <- edited_values$Remark
    }
    removeModal()
    # Now you can use `selected_data$data` with the edited values
  }
  
    
})
  
  
  
  
  
  
  
  
  
observeEvent(input$reset_po_data, {
  original_choices_no <- unique(maker_po_data$data$maker_order_number)
  updateSelectizeInput(session, "no", choices = c("", original_choices_no))
  # Reset
  
    
})
  
excel_helpers <- source("helper/format_excel.R", local = TRUE)$value
  
output$download_button_maker_po <- renderUI({
    
  if(nrow(selected_data$data)!=0){
    actionButton("download_maker","下載訂購單")
  }
    
  else {
    actionButton("dummybutton_maker","下載訂購單")
  }
    
})
  
  
observeEvent(input$download_maker,{
  if(selected_data$data$我方公司[1]=="Gluon Taiwan"){
     showModal(modalDialog(
      title = "下載格式",
      downloadButton("download_maker_tj", "外購日語訂購單"),
      downloadButton("download_maker_te", "外購英語訂購單"),
      downloadButton("download_maker_tt", "內購訂購單"),
      align = "center",
      tags$head(
        tags$style(HTML("
          .modal {
            display: flex;
            align-items: center;
            justify-content: center;
          }
        "))),
      easyClose = TRUE
    ))
  }
  else if (selected_data$data$我方公司[1]=="Gluon Japan"){
    showModal(modalDialog(
      title = "下載格式",
      downloadButton("download_maker_jj", "日語訂購單"),
      align = "center",
      tags$head(
        tags$style(HTML("
          .modal {
            display: flex;
            align-items: center;
            justify-content: center;
          }
        "))),
      easyClose = TRUE
    ))
  }
  else if (selected_data$data$我方公司[1]=="無錫群唐信"){
    showModal(modalDialog(
      title = "下載格式",
      downloadButton("download_maker_ce", "英語訂購單"),
      align = "center",
      tags$head(
        tags$style(HTML("
          .modal {
            display: flex;
            align-items: center;
            justify-content: center;
          }
        "))),
      easyClose = TRUE
    ))
  }
  else{
    showModal(modalDialog(
      title = "下載格式",
      downloadButton("download_maker_tj", "日語訂購單"),
      downloadButton("download_maker_te", "英語訂購單"),
      downloadButton("download_maker_tt", "內購訂購單"),
      align = "center",
      tags$head(
      tags$style(HTML("
          .modal {  
          display: flex;
          align-items: center;
          justify-content: center;
          }
        "))),
      easyClose = TRUE
    ))
  }
    
    
  })
  
  
 observeEvent(input$dummybutton_maker,{
  showModal(modalDialog(
    title = "Warning",
    "沒有資料下載"
  ))
})
  
  
output$download_maker_tj <- downloadHandler(
    
  filename = function() {
    paste0(as.character(input$no), "訂購單.xlsx")
  },
  content = function(file){
    my_workbook <- createWorkbook()
    
    addWorksheet(
      wb = my_workbook,
      sheetName = paste0(as.character(input$no))
    )
      
    setColWidths(
      my_workbook,
      1,
      cols = 1:7,
      widths =  c(1.67, 4.5, 46, 6, 6,16.5, 18.5)
      
    )
      
    setRowHeights(
      my_workbook,
      sheet = 1,  # Specify the sheet number or name
      rows = 1:6,  # Specify the range of rows
      heights = c(20,20, 20, 20, 20, 30)  # Convert your desired heights to Excel units
    )
    
    insertImage(
      my_workbook, 
      sheet = 1,
      file = "www/gluon_title.png",
      width = 8.4,
      height = 1.32,
      startRow = 1,
      startCol = 2,
      units = "in",
      dpi = 300
    )
      
    # purchase order
      
    excel_helpers$create_header_row(
      my_workbook,
      1,
      list(
        list("Purchase Order", 7)
      ),
      startRow = 6,
      startCol = 1
    )
    
      
    # po style
    addStyle(
      my_workbook,
      sheet = 1,
      cols = 1:7,
        rows = 6,
        style = createStyle(
          textDecoration = "Bold",
          fontSize = 24,
          halign = "center", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # No
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 9,
        startCol = 2,
        x = paste0("NO: ", as.character(input$no))
      )
      
      # No style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 9,
        cols = 2,
        style = createStyle(
          
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      # date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 10,
        startCol = 7,
        x = paste0(selected_data$data$日期[1])
      )
      
      # date style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 10,
        cols = 7,
        style = createStyle(
          halign = "right",
          fontSize = 14,
          fontName = "Arial"
        )
      )
      
      
      # maker 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 11,
        startCol = 2,
        x = paste0("客先 : ")
      )
      
      
      # maker name
      if (selected_data$data$Maker[1] %in% maker_name$廠商簡稱) {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 12,
          startCol = 2,
          x = paste0("   ",maker_name$廠商名稱[maker_name$廠商簡稱==selected_data$data$Maker[1]])
        )
      } else {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 12,
          startCol = 2,
          x = paste0("   ",selected_data$data$Maker[1])
        )
      }
      
      # style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 11:12,
        cols = 2,
        style = createStyle(
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      paymentT <- strsplit(selected_data$data$Payment_Terms[1], " / ")[[1]]
      PTlength <- length(paymentT) -1
     
      # payment 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 14,
        startCol = 2,
        x = paste0("支払条件：")
      )
      
      
      # payment terms
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 15,
        startCol = 2,
        x = paste0("   ", c(paymentT))
      )
      
      # destination 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 17+PTlength,
        startCol = 2,
        x = paste0("受渡場所：")
      )
      
      # destination name
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 18+PTlength,
        startCol = 2,
        x = paste0("   ", selected_data$data$Destination[1])
      )
      
      # Shipment date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 20+PTlength,
        startCol = 2,
        x = paste0("納期：")
      )
      
      # Shipment_Date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 21+PTlength,
        startCol = 2,
        x = paste0("   ", selected_data$data$Shipment_Date[1])
      )
      
      
      # B style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 13:(21+PTlength),
        cols = 2,
        style = createStyle(
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 2,
        x = "No. ",
        
      )
      
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 3,
        x = "Description",
        
      )
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Q'ty", 2)
        ),
        startRow = 24+PTlength,
        startCol = 4
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 6,
        x = paste0("Unit Price (", selected_data$data$幣別[1], ")"),
        
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 7,
        x = paste0("Amount (", selected_data$data$幣別[1], ")"),
        
      )
      
      # table style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 24+PTlength,
        cols = 2:7,
        style = createStyle(
          fontSize = 12,
          fontName = "Arial",
          border = "TopBottomLeftRight", 
          halign = "center"
        ),
        gridExpand = TRUE
      )
      
      # Initialize a list to store the rows of the final data frame
      maker_table <- data.frame(
        No = integer(),
        Description = character(),
        Quantity = numeric(),
        Unit = character(),
        UP = numeric(),
        Am = numeric(),
        stringsAsFactors = FALSE  # Ensure character columns are not converted to factors
      )
      
      # Loop through selected_data$data
      for (i in seq_len(nrow(selected_data$data))) {
        # Get names after splitting
        names <- strsplit(selected_data$data$名稱[i], " / ")[[1]]
        # Check if Model is not empty
        if (selected_data$data$Model[i] != "") {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if (j == length(names)) {
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
              maker_table <- rbind(maker_table, c("", selected_data$data$Model[i], "", "", "", ""))
            } else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        } else {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if (j == length(names)) {
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
            } else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        }
      } 
      
      # Set column names
      colnames(maker_table) <- c("No", "Description", "Quantity", "Unit", "UnitP", "TotalP")
      
      # Verify the class and formatting
      class(maker_table$UnitP) <- "comma"
      class(maker_table$TotalP) <- "comma"
      
    
      # Create a style
      table_style <- createStyle(
        fontSize = 12,
        fontName = "Arial",
        border = "TopBottomLeftRight"
      )
      
      # Apply style to each cell in maker_table
      for (row_index in 1:nrow(maker_table)) {
        for (col_index in 1:ncol(maker_table)) {
          addStyle(my_workbook, sheet = 1, 
                   style = table_style, 
                   rows = 24 + row_index+PTlength, 
                   cols = 1 + col_index)
        }
      }
      
      
      
      # Write data to the worksheet
      writeData(
        my_workbook,
        sheet = 1,
        x = maker_table,
        startRow = 25+PTlength,
        startCol = 2,
        colNames = FALSE  # Suppress column headers
      )
      
     
  
      # Q'ty
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:5,
        rows = (25+PTlength):(24+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "TopBottom"
          
        ),
        gridExpand = TRUE
      )
      
      # No.
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2,
        rows = (25+PTlength):(24+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
        ),
        gridExpand = TRUE
      )
      
      
      # sub total
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Sub Total ", 5)
        ),
        startRow = 24+nrow(maker_table)+1+PTlength,
        startCol = 2
      )
      
      st <- sum(selected_data$data$Quantity*selected_data$data$價格)
      # sub total
      writeData(
        my_workbook,
        sheet = 1,
        x = paste0(format(st, scientific = FALSE, big.mark = ",")),
        startRow = 24+nrow(maker_table)+1+PTlength,
        startCol = 7,
        colNames = FALSE  # Suppress column headers
      )
      
      # sub total style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:7,
        rows = 24+nrow(maker_table)+1+PTlength,
        style = createStyle(
          fontSize = 12,
          halign = "right", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      if(selected_data$data$Tax[1]!=0){
        # tax
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(paste0("Tax(", selected_data$data$Tax[1]*100,"%)"), 5)
          ),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 2
        )
        
        # tax amount
        writeData(
          my_workbook,
          sheet = 1,
          x = sum(selected_data$data$Quantity*selected_data$data$價格*(selected_data$data$Tax[1])),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 7,
          colNames = FALSE  # Suppress column headers
        )
        
        # tax style
        addStyle(
          my_workbook,
          sheet = 1,
          cols = 2:7,
          rows = 24+nrow(maker_table)+1+1+PTlength,
          style = createStyle(
            fontSize = 12,
            halign = "right", 
            fontName = "Arial",
            border = "TopBottomLeftRight"
            
          ),
          gridExpand = TRUE
        )
      }
      
      else{
        # tax
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(paste0("Tax(税抜)"), 5)
          ),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 2
        )
        
        # tax amount
        writeData(
          my_workbook,
          sheet = 1,
          x = "0",
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 7,
          colNames = FALSE  # Suppress column headers
        )
        
        # tax style
        addStyle(
          my_workbook,
          sheet = 1,
          cols = 2:7,
          rows = 24+nrow(maker_table)+1+1+PTlength,
          style = createStyle(
            fontSize = 12,
            halign = "right", 
            fontName = "Arial",
            border = "TopBottomLeftRight"
            
          ),
          gridExpand = TRUE
        )
      }
      
      
      # grand total
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list(paste0("Grand Total"), 5)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+PTlength,
        startCol = 2
      )
      gt <- sum(selected_data$data$Quantity*selected_data$data$價格*(1+selected_data$data$Tax[1]))
      # grand total amount
      writeData(
        my_workbook,
        sheet = 1,
        x = paste0(format(gt, scientific = FALSE, big.mark = ",")),
        startRow = 24+nrow(maker_table)+1+1+1+PTlength,
        startCol = 7,
        colNames = FALSE  # Suppress column headers
      )
      
      # grand total style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:7,
        rows = 24+nrow(maker_table)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 12,
          halign = "right", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      
      # remark
      
      if (is.na(selected_data$data$Remark[1])||selected_data$data$Remark[1] == ""){
        split_maker_remark <- ""
        
      }
      else{
        
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("Remark:"),
          startRow = 24+nrow(maker_table)+1+1+1+1+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        
        
        split_maker_remark <- unlist(strsplit(selected_data$data$Remark[1], "\\s(?=\\d+\\. )", perl = TRUE))
        # remark content
        writeData(
          my_workbook,
          sheet = 1,
          x = paste("    ", split_maker_remark),
          startRow = 24+nrow(maker_table)+1+1+1+1+1+PTlength,
          startCol = 2,
        )
        
      }
      
      
      
      # signature
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Gluon Technology Service Corp.", 4)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+1+1+PTlength,
        startCol = 4
      )
      
      
      # signature style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:7,
        rows = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 18,
          halign = "left", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      
      # underline style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:7,
        rows = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+6+1+1+PTlength,
        style = createStyle(
          border = "Bottom"
        ),
        gridExpand = TRUE
      )
      
      # shinga chen
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Shinga Chen President", 3)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+6+1+1+1+PTlength,
        startCol = 4
      )
      
      # shinga chen style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:7,
        rows = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+6+1+1+1+PTlength,
        style = createStyle(
          fontSize = 11,
          halign = "left"
        ),
        gridExpand = TRUE
      )
      
      
      saveWorkbook(my_workbook, file)
    },

    
  )
  output$download_maker_te <- downloadHandler(
    
    filename = function() {
      paste0(as.character(input$no), "訂購單.xlsx")
    },
    content = function(file){
      my_workbook <- createWorkbook()
      
      addWorksheet(
        wb = my_workbook,
        sheetName = paste0(as.character(input$no))
      )
      
      setColWidths(
        my_workbook,
        1,
        cols = 1:7,
        widths =  c(1.67, 4.5, 46, 6, 6,16.5, 18.5)
        
      )
      
      setRowHeights(
        my_workbook,
        sheet = 1,  # Specify the sheet number or name
        rows = 1:6,  # Specify the range of rows
        heights = c(20,20, 20, 20, 20, 30)  # Convert your desired heights to Excel units
      )
      
      insertImage(
        my_workbook, 
        sheet = 1,
        file = "www/gluon_title.png",
        width = 8.4,
        height = 1.32,
        startRow = 1,
        startCol = 2,
        units = "in",
        dpi = 300
      )
      
      # purchase order
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Purchase Order", 7)
        ),
        startRow = 6,
        startCol = 1
      )
      
      
      # po style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 1:7,
        rows = 6,
        style = createStyle(
          textDecoration = "Bold",
          fontSize = 24,
          halign = "center", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # No
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 9,
        startCol = 2,
        x = paste0("NO: ", as.character(input$no))
      )
      
      # No style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 9,
        cols = 2,
        style = createStyle(
          
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      # date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 10,
        startCol = 7,
        x = paste0(selected_data$data$日期[1])
      )
      
      # date style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 10,
        cols = 7,
        style = createStyle(
          halign = "right",
          fontSize = 14,
          fontName = "Arial"
        )
      )
      
      
      # maker 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 11,
        startCol = 2,
        x = paste0("TO:")
      )
    
      # maker name
      if (selected_data$data$Maker[1] %in% maker_name$廠商簡稱) {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 12,
          startCol = 2,
          x = paste0("   ",maker_name$廠商名稱[maker_name$廠商簡稱==selected_data$data$Maker[1]])
        )
      } else {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 12,
          startCol = 2,
          x = paste0("   ",selected_data$data$Maker[1])
        )
      }
      
      
      # style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 11:12,
        cols = 2,
        style = createStyle(
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      

      # payment 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 14,
        startCol = 2,
        x = paste0("Payment Terms:")
      )
      
      # payment terms
      paymentT <- strsplit(selected_data$data$Payment_Terms[1], " / ")[[1]]
      PTlength <- length(paymentT) -1
      # payment terms
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 15,
        startCol = 2,
        x = paste0("   ", c(paymentT))
      )
      
      # destination 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 17+PTlength,
        startCol = 2,
        x = paste0("Destination: ")
      )
      
      # destination name
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 18+PTlength,
        startCol = 2,
        x = paste0("   ", selected_data$data$Destination[1])
      )
      
      # Shipment date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 20+PTlength,
        startCol = 2,
        x = paste0("Shipment Date:")
      )
      
      # Shipment_Date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 21+PTlength,
        startCol = 2,
        x = paste0("   ", selected_data$data$Shipment_Date[1])
      )
      
      # style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 14:(21+PTlength),
        cols = 2,
        style = createStyle(
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 2,
        x = "No. ",
        
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 3,
        x = "Description",
        
      )
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Q'ty", 2)
        ),
        startRow = 24+PTlength,
        startCol = 4
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 6,
        x = paste0("Unit Price (", selected_data$data$幣別[1], ")"),
        
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 7,
        x = paste0("Amount (", selected_data$data$幣別[1], ")"),
        
      )
      
      # table style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 24+PTlength,
        cols = 2:7,
        style = createStyle(
          fontSize = 12,
          fontName = "Arial",
          border = "TopBottomLeftRight", 
          halign = "center"
        ),
        gridExpand = TRUE
      )
      
      # Initialize a list to store the rows of the final data frame
      maker_table <- data.frame(
        No = integer(),
        Description = character(),
        Quantity = numeric(),
        Unit = character(),
        UP = numeric(),
        Am = numeric(),
        stringsAsFactors = FALSE  # Ensure character columns are not converted to factors
      )
      
      # Loop through selected_data$data
      for (i in seq_len(nrow(selected_data$data))) {
        # Get names after splitting
        names <- strsplit(selected_data$data$名稱[i], " / ")[[1]]
        # Check if Model is not empty
        if (selected_data$data$Model[i] != "") {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if (j == length(names)) {
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
              maker_table <- rbind(maker_table, c("", selected_data$data$Model[i], "", "", "", ""))
            } else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        } else {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if (j == length(names)) {
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
            } else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        }
      } 
      
      # Set column names
      colnames(maker_table) <- c("No", "Description", "Quantity", "Unit", "UnitP", "TotalP")
      
      # Verify the class and formatting
      class(maker_table$UnitP) <- "comma"
      class(maker_table$TotalP) <- "comma"
      
      
      # Create a style
      table_style <- createStyle(
        fontSize = 12,
        fontName = "Arial",
        border = "TopBottomLeftRight"
      )
      
      # Apply style to each cell in maker_table
      for (row_index in 1:nrow(maker_table)) {
        for (col_index in 1:ncol(maker_table)) {
          addStyle(my_workbook, sheet = 1, 
                   style = table_style, 
                   rows = 24 + row_index+PTlength, 
                   cols = 1 + col_index)
        }
      }
      
      
      
      # Write data to the worksheet
      writeData(
        my_workbook,
        sheet = 1,
        x = maker_table,
        startRow = 25+PTlength,
        startCol = 2,
        colNames = FALSE  # Suppress column headers
      )
      
      # Q'ty
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:5,
        rows = (25+PTlength):(24+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "TopBottom"
          
        ),
        gridExpand = TRUE
      )
      
      # No.
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2,
        rows = (25+PTlength):(24+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      # sub total
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Sub Total ", 5)
        ),
        startRow = 24+nrow(maker_table)+1+PTlength,
        startCol = 2
      )
      
      st <- sum(selected_data$data$Quantity*selected_data$data$價格)
      # sub total
      writeData(
        my_workbook,
        sheet = 1,
        x = paste0(format(st, scientific = FALSE, big.mark = ",")),
        startRow = 24+nrow(maker_table)+1+PTlength,
        startCol = 7,
        colNames = FALSE  # Suppress column headers
      )
      
      # sub total style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:7,
        rows = 24+nrow(maker_table)+1+PTlength,
        style = createStyle(
          fontSize = 12,
          halign = "right", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      if(selected_data$data$Tax[1]!=0){
        # tax
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(paste0("Tax(", selected_data$data$Tax[1]*100,"%)"), 5)
          ),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 2
        )
        
        # tax amount
        writeData(
          my_workbook,
          sheet = 1,
          x = sum(selected_data$data$Quantity*selected_data$data$價格*(selected_data$data$Tax[1])),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 7,
          colNames = FALSE  # Suppress column headers
        )
        
        # tax style
        addStyle(
          my_workbook,
          sheet = 1,
          cols = 2:7,
          rows = 24+nrow(maker_table)+1+1+PTlength,
          style = createStyle(
            fontSize = 12,
            halign = "right", 
            fontName = "Arial",
            border = "TopBottomLeftRight"
            
          ),
          gridExpand = TRUE
        )
      }
      
      else{
        # tax
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(paste0("Tax(excluded)"), 5)
          ),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 2
        )
        
        # tax amount
        writeData(
          my_workbook,
          sheet = 1,
          x = "0",
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 7,
          colNames = FALSE  # Suppress column headers
        )
        
        # tax style
        addStyle(
          my_workbook,
          sheet = 1,
          cols = 2:7,
          rows = 24+nrow(maker_table)+1+1+PTlength,
          style = createStyle(
            fontSize = 12,
            halign = "right", 
            fontName = "Arial",
            border = "TopBottomLeftRight"
            
          ),
          gridExpand = TRUE
        )
      }
      
      
      # grand total
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list(paste0("Grand Total"), 5)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+PTlength,
        startCol = 2
      )
      gt <- sum(selected_data$data$Quantity*selected_data$data$價格*(1+selected_data$data$Tax[1]))
      # grand total amount
      writeData(
        my_workbook,
        sheet = 1,
        x = paste0(format(gt, scientific = FALSE, big.mark = ",")),
        startRow = 24+nrow(maker_table)+1+1+1+PTlength,
        startCol = 7,
        colNames = FALSE  # Suppress column headers
      )
      
      # grand total style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:7,
        rows = 24+nrow(maker_table)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 12,
          halign = "right", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      
      # remark
      
      if (is.na(selected_data$data$Remark[1])||selected_data$data$Remark[1] == ""){
        split_maker_remark <- ""
        
      }
      else{
        
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("Remark:"),
          startRow = 24+nrow(maker_table)+1+1+1+1+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        
        
        split_maker_remark <- unlist(strsplit(selected_data$data$Remark[1], "\\s(?=\\d+\\. )", perl = TRUE))
        # remark content
        writeData(
          my_workbook,
          sheet = 1,
          x = paste("    ", split_maker_remark),
          startRow = 24+nrow(maker_table)+1+1+1+1+1+PTlength,
          startCol = 2,
        )
        
      }
      
      
      
      # signature
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Gluon Technology Service Corp.", 4)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+1+1+PTlength,
        startCol = 4
      )
      
      
      # signature style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:7,
        rows = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 18,
          halign = "left", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      
      # underline style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:7,
        rows = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+6+1+1+PTlength,
        style = createStyle(
          border = "Bottom"
        ),
        gridExpand = TRUE
      )
      
      # shinga chen
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Shinga Chen President", 3)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+6+1+1+1+PTlength,
        startCol = 4
      )
      
      # shinga chen style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:7,
        rows = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+6+1+1+1+PTlength,
        style = createStyle(
          fontSize = 11,
          halign = "left"
        ),
        gridExpand = TRUE
      )
      
      
      saveWorkbook(my_workbook, file)
    }
    
  )
  
  output$download_maker_tt <- downloadHandler(
    
    filename = function() {
      paste0(as.character(input$no), "訂購單.xlsx")
    },
    content = function(file){
      my_workbook <- createWorkbook()
      
      addWorksheet(
        wb = my_workbook,
        sheetName = paste0(as.character(input$no))
      )
      
      setColWidths(
        my_workbook,
        1,
        cols = 1:7,
        widths =  c(1.67, 4.5, 46, 6, 6,16.5, 18.5)
        
      )
      
      setRowHeights(
        my_workbook,
        sheet = 1,  # Specify the sheet number or name
        rows = 1:6,  # Specify the range of rows
        heights = c(20,20, 20, 20, 20, 30)  # Convert your desired heights to Excel units
      )
      
      insertImage(
        my_workbook, 
        sheet = 1,
        file = "www/gluon_title.png",
        width = 8.4,
        height = 1.32,
        startRow = 1,
        startCol = 2,
        units = "in",
        dpi = 300
      )
      
      # purchase order
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Purchase Order", 7)
        ),
        startRow = 6,
        startCol = 1
      )
      
      
      # po style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 1:7,
        rows = 6,
        style = createStyle(
          textDecoration = "Bold",
          fontSize = 24,
          halign = "center", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # No
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 9,
        startCol = 2,
        x = paste0("NO: ", as.character(input$no))
      )
      
      # No style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 9,
        cols = 2,
        style = createStyle(
          
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      # date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 10,
        startCol = 7,
        x = paste0(selected_data$data$日期[1])
      )
      
      # date style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 10,
        cols = 7,
        style = createStyle(
          halign = "right",
          fontSize = 14,
          fontName = "Arial"
        )
      )
      
      
      # maker 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 11,
        startCol = 2,
        x = paste0("TO:")
      )
      
    
      # maker name
      if (selected_data$data$Maker[1] %in% maker_name$廠商簡稱) {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 12,
          startCol = 2,
          x = paste0("   ",maker_name$廠商名稱[maker_name$廠商簡稱==selected_data$data$Maker[1]])
        )
      } else {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 12,
          startCol = 2,
          x = paste0("   ",selected_data$data$Maker[1])
        )
      }
      
      
      # style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 11:12,
        cols = 2,
        style = createStyle(
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # payment 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 14,
        startCol = 2,
        x = paste0("Payment Terms:")
      )
    
     
      # payment terms
      paymentT <- strsplit(selected_data$data$Payment_Terms[1], " / ")[[1]]
      PTlength <- length(paymentT) -1
      # payment terms
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 15,
        startCol = 2,
        x = paste0("   ", c(paymentT))
      )
      
      # destination 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 17+PTlength,
        startCol = 2,
        x = paste0("Destination: ")
      )
      
      # destination name
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 18+PTlength,
        startCol = 2,
        x = paste0("   ", selected_data$data$Destination[1])
      )
      
      # Shipment date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 20+PTlength,
        startCol = 2,
        x = paste0("Shipment Date:")
      )
      
      # Shipment_Date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 21+PTlength,
        startCol = 2,
        x = paste0("   ", selected_data$data$Shipment_Date[1])
      )
      
      
      # B style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 13:(21+PTlength),
        cols = 2,
        style = createStyle(
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 2,
        x = "No. ",
        
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 3,
        x = "Description",
        
      )
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Q'ty", 2)
        ),
        startRow = 24+PTlength,
        startCol = 4
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 6,
        x = paste0("Unit Price (", selected_data$data$幣別[1], ")"),
        
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 7,
        x = paste0("Amount (", selected_data$data$幣別[1], ")"),
        
      )
      
      # table style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 24+PTlength,
        cols = 2:7,
        style = createStyle(
          fontSize = 12,
          fontName = "Arial",
          border = "TopBottomLeftRight", 
          halign = "center"
        ),
        gridExpand = TRUE
      )
      
      # Initialize a list to store the rows of the final data frame
      maker_table <- data.frame(
        No = integer(),
        Description = character(),
        Quantity = numeric(),
        Unit = character(),
        UP = numeric(),
        Am = numeric(),
        stringsAsFactors = FALSE  # Ensure character columns are not converted to factors
      )
      
      # Loop through selected_data$data
      for (i in seq_len(nrow(selected_data$data))) {
        # Get names after splitting
        names <- strsplit(selected_data$data$名稱[i], " / ")[[1]]
        # Check if Model is not empty
        if (selected_data$data$Model[i] != "") {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if(j == 1){
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
            }
            else if (j == length(names)) {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
              maker_table <- rbind(maker_table, c("", selected_data$data$Model[i], "", "", "", ""))
            } else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        } else {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if (j ==1) {
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
            } else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        }
      } 
      
      # Set column names
      colnames(maker_table) <- c("No", "Description", "Quantity", "Unit", "UnitP", "TotalP")
      
      # Verify the class and formatting
      class(maker_table$UnitP) <- "comma"
      class(maker_table$TotalP) <- "comma"
      
      
      # Create a style
      table_style <- createStyle(
        fontSize = 12,
        fontName = "Arial",
        border = "TopBottomLeftRight"
      )
      
      # Apply style to each cell in maker_table
      for (row_index in 1:nrow(maker_table)) {
        for (col_index in 1:ncol(maker_table)) {
          addStyle(my_workbook, sheet = 1, 
                   style = table_style, 
                   rows = 24 + row_index+PTlength, 
                   cols = 1 + col_index)
        }
      }
      
      
      
      # Write data to the worksheet
      writeData(
        my_workbook,
        sheet = 1,
        x = maker_table,
        startRow = 25+PTlength,
        startCol = 2,
        colNames = FALSE  # Suppress column headers
      )
      
      # Q'ty
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:5,
        rows = (25+PTlength):(24+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "TopBottom"
          
        ),
        gridExpand = TRUE
      )
      
      # No.
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2,
        rows = (25+PTlength):(24+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
        ),
        gridExpand = TRUE
      )
      
      # sub total
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Sub Total ", 5)
        ),
        startRow = 24+nrow(maker_table)+1+PTlength,
        startCol = 2
      )
      
      st <- sum(selected_data$data$Quantity*selected_data$data$價格)
      # sub total
      writeData(
        my_workbook,
        sheet = 1,
        x = paste0(format(st, scientific = FALSE, big.mark = ",")),
        startRow = 24+nrow(maker_table)+1+PTlength,
        startCol = 7,
        colNames = FALSE  # Suppress column headers
      )
      
      # sub total style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:7,
        rows = 24+nrow(maker_table)+1+PTlength,
        style = createStyle(
          fontSize = 12,
          halign = "right", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      if(selected_data$data$Tax[1]!=0){
        # tax
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(paste0("Tax(", selected_data$data$Tax[1]*100,"%)"), 5)
          ),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 2
        )
        
        # tax amount
        writeData(
          my_workbook,
          sheet = 1,
          x = sum(selected_data$data$Quantity*selected_data$data$價格*(selected_data$data$Tax[1])),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 7,
          colNames = FALSE  # Suppress column headers
        )
        
        # tax style
        addStyle(
          my_workbook,
          sheet = 1,
          cols = 2:7,
          rows = 24+nrow(maker_table)+1+1+PTlength,
          style = createStyle(
            fontSize = 12,
            halign = "right", 
            fontName = "Arial",
            border = "TopBottomLeftRight"
            
          ),
          gridExpand = TRUE
        )
      }
      
      else{
        # tax
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(paste0("Tax(excluded)"), 5)
          ),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 2
        )
        
        # tax amount
        writeData(
          my_workbook,
          sheet = 1,
          x = "0",
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 7,
          colNames = FALSE  # Suppress column headers
        )
        
        # tax style
        addStyle(
          my_workbook,
          sheet = 1,
          cols = 2:7,
          rows = 24+nrow(maker_table)+1+1+PTlength,
          style = createStyle(
            fontSize = 12,
            halign = "right", 
            fontName = "Arial",
            border = "TopBottomLeftRight"
            
          ),
          gridExpand = TRUE
        )
      }
      
      
      # grand total
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list(paste0("Grand Total"), 5)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+PTlength,
        startCol = 2
      )
      gt <- sum(selected_data$data$Quantity*selected_data$data$價格*(1+selected_data$data$Tax[1]))
      # grand total amount
      writeData(
        my_workbook,
        sheet = 1,
        x = paste0(format(gt, scientific = FALSE, big.mark = ",")),
        startRow = 24+nrow(maker_table)+1+1+1+PTlength,
        startCol = 7,
        colNames = FALSE  # Suppress column headers
      )
      
      # grand total style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:7,
        rows = 24+nrow(maker_table)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 12,
          halign = "right", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      
      # remark
      
      if (is.na(selected_data$data$Remark[1])||selected_data$data$Remark[1] == ""){
        split_maker_remark <- ""
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("Remark:"),
          startRow = 24+nrow(maker_table)+4+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","1. 發票抬頭&寄送地址:義達科技股份有限公司(統編:28622325)"),
          startRow = 24+nrow(maker_table)+5+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","30262 新竹縣竹北市溪州路250-2號"),
          startRow = 24+nrow(maker_table)+6+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","TEL:03-5520-107"),
          startRow = 24+nrow(maker_table)+7+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","FAX:03-5520-323"),
          startRow = 24+nrow(maker_table)+8+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","2. 若因天災、事變或其他不可抗力之情事所致無法如期交貨"),
          startRow = 24+nrow(maker_table)+9+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","，請於24小時內Mail通報回覆。"),
          startRow = 24+nrow(maker_table)+10+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        
      }
      else{
        
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("Remark:"),
          startRow = 24+nrow(maker_table)+4+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","1. 發票抬頭&寄送地址:義達科技股份有限公司(統編:28622325)"),
          startRow = 24+nrow(maker_table)+5+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","30262 新竹縣竹北市溪州路250-2號"),
          startRow = 24+nrow(maker_table)+6+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","TEL:03-5520-107"),
          startRow = 24+nrow(maker_table)+7+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","FAX:03-5520-323"),
          startRow = 24+nrow(maker_table)+8+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","2. 若因天災、事變或其他不可抗力之情事所致無法如期交貨"),
          startRow = 24+nrow(maker_table)+9+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","，請於24小時內Mail通報回覆。"),
          startRow = 24+nrow(maker_table)+10+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        
        
        
        split_maker_remark <- unlist(strsplit(selected_data$data$Remark[1], "(\\d+\\. )", perl = TRUE))[-1]
        for (i in seq_along(split_maker_remark)) {
          split_maker_remark[i] <- paste(i + 2, ". ", split_maker_remark[i], sep = "")
        }
        
        
        # remark content
        writeData(
          my_workbook,
          sheet = 1,
          x = paste("  ", split_maker_remark),
          startRow = 24+nrow(maker_table)+11+PTlength,
          startCol = 2,
        )
        
      }
      
      
      
      # signature
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Gluon Technology Service Corp.", 3)
        ),
        startRow = 24+nrow(maker_table)+11+length(split_maker_remark)+1+1+1+PTlength,
        startCol = 5
      )
      
      
      # signature style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 5:7,
        rows = 24+nrow(maker_table)+11+length(split_maker_remark)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 15,
          halign = "left", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      
      # underline style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 5:7,
        rows = 24+nrow(maker_table)+11+length(split_maker_remark)+1+6+1+1+PTlength,
        style = createStyle(
          border = "Bottom"
        ),
        gridExpand = TRUE
      )
      
      # shinga chen
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Shinga Chen President", 3)
        ),
        startRow = 24+nrow(maker_table)+11+length(split_maker_remark)+1+6+1+1+1+PTlength,
        startCol = 5
      )
      
      # shinga chen style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 5:7,
        rows = 24+nrow(maker_table)+11+length(split_maker_remark)+1+6+1+1+1+PTlength,
        style = createStyle(
          fontSize = 11,
          halign = "left"
        ),
        gridExpand = TRUE
      )
      
      
      
      # maker signature
      if (selected_data$data$Maker[1] %in% maker_name$廠商簡稱) {
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(maker_name$廠商名稱[maker_name$廠商簡稱==selected_data$data$Maker[1]], 2)
          ),
          startRow = 24+nrow(maker_table)+11+length(split_maker_remark)+1+1+1+PTlength,
          startCol = 2
        )
        
      } else {
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(selected_data$data$Maker[1], 2)
          ),
          startRow = 24+nrow(maker_table)+11+length(split_maker_remark)+1+1+1+PTlength,
          startCol = 2
        )
        
      }
      
      
      
      
      # maker signature style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:3,
        rows = 24+nrow(maker_table)+11+length(split_maker_remark)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 15,
          halign = "left", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      
      # maker underline style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:3,
        rows = 24+nrow(maker_table)+11+length(split_maker_remark)+1+6+1+1+PTlength,
        style = createStyle(
          border = "Bottom"
        ),
        gridExpand = TRUE
      )
      
      
      
      
      saveWorkbook(my_workbook, file)
    }
    
  )
  
  
  output$download_maker_jj <- downloadHandler(
    
    filename = function() {
      paste0(as.character(input$no), "訂購單.xlsx")
    },
    content = function(file){
      my_workbook <- createWorkbook()
      
      addWorksheet(
        wb = my_workbook,
        sheetName = paste0(as.character(input$no))
      )
      
      setColWidths(
        my_workbook,
        1,
        cols = 1:7,
        widths =  c(1.67, 4.5, 46, 6, 6,16.5, 18.5)
        
      )
      
      # Gluon Japan title
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Gluon Japan Corporation", 7)
        ),
        startRow = 1,
        startCol = 1
      )
      
      # Gluon Japan title style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 1:7,
        rows = 1,
        style = createStyle(
          textDecoration = "Bold",
          fontSize = 28,
          halign = "center", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # Chinese title
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("義達科技股份有限公司", 7)
        ),
        startRow = 2,
        startCol = 1
      )
      
      # Chinese title style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 1:7,
        rows = 2,
        style = createStyle(
          fontSize = 22,
          halign = "center", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # address1 title
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("新竹縣竹北市溪州路250-2號 Phone: 886-3-552-0107  Fax : 886-3-552-0323", 7)
        ),
        startRow = 3,
        startCol = 1
      )
      
      # address1 style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 1:7,
        rows = 3,
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # address1 title
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("稅籍地址：新竹市竹光路139號8F-1", 7)
        ),
        startRow = 4,
        startCol = 1
      )
      
      # address1 style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 1:7,
        rows = 4,
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      
      # purchase order
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Purchase Order", 7)
        ),
        startRow = 6,
        startCol = 1
      )
      
      
      # po style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 1:7,
        rows = 6,
        style = createStyle(
          textDecoration = "Bold",
          fontSize = 24,
          halign = "center", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # No
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 9,
        startCol = 2,
        x = paste0("NO: ", as.character(input$no))
      )
      
      # No style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 9,
        cols = 2,
        style = createStyle(
          
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      # date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 10,
        startCol = 7,
        x = paste0(selected_data$data$日期[1])
      )
      
      # date style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 10,
        cols = 7,
        style = createStyle(
          halign = "right",
          fontSize = 14,
          fontName = "Arial"
        )
      )
      
      
      # maker 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 11,
        startCol = 2,
        x = paste0("客先 : ")
      )
      
      # maker name
      if (selected_data$data$Maker[1] %in% maker_name$廠商簡稱) {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 12,
          startCol = 2,
          x = paste0("   ",maker_name$廠商名稱[maker_name$廠商簡稱==selected_data$data$Maker[1]])
        )
      } else {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 12,
          startCol = 2,
          x = paste0("   ",selected_data$data$Maker[1])
        )
      }
      
      
      # style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 11:12,
        cols = 2,
        style = createStyle(
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
     
      
      # payment 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 14,
        startCol = 2,
        x = paste0("支払条件：")
      )
      
     
      
      # payment terms
      paymentT <- strsplit(selected_data$data$Payment_Terms[1], " / ")[[1]]
      PTlength <- length(paymentT) -1
      # payment terms
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 15,
        startCol = 2,
        x = paste0("   ", c(paymentT))
      )
      
      
      # destination 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 17+PTlength,
        startCol = 2,
        x = paste0("受渡場所：")
      )
      
      # destination name
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 18+PTlength,
        startCol = 2,
        x = paste0("   ", selected_data$data$Destination[1])
      )
      
      # Shipment date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 20+PTlength,
        startCol = 2,
        x = paste0("納期：")
      )
      
      # Shipment_Date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 21+PTlength,
        startCol = 2,
        x = paste0("   ", selected_data$data$Shipment_Date[1])
      )
      
      # style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 13:(21+PTlength),
        cols = 2,
        style = createStyle(
          fontSize = 14,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 2,
        x = "No. ",
        
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 3,
        x = "Description",
        
      )
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Q'ty", 2)
        ),
        startRow = 24+PTlength,
        startCol = 4
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 6,
        x = paste0("Unit Price (", selected_data$data$幣別[1], ")"),
        
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 24+PTlength,
        startCol = 7,
        x = paste0("Amount (", selected_data$data$幣別[1], ")"),
        
      )
      
      # table style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 24+PTlength,
        cols = 2:7,
        style = createStyle(
          fontSize = 12,
          fontName = "Arial",
          border = "TopBottomLeftRight", 
          halign = "center"
        ),
        gridExpand = TRUE
      )
      
      # Initialize a list to store the rows of the final data frame
      maker_table <- data.frame(
        No = integer(),
        Description = character(),
        Quantity = numeric(),
        Unit = character(),
        UP = numeric(),
        Am = numeric(),
        stringsAsFactors = FALSE  # Ensure character columns are not converted to factors
      )
      
      # Loop through selected_data$data
      for (i in seq_len(nrow(selected_data$data))) {
        # Get names after splitting
        names <- strsplit(selected_data$data$名稱[i], " / ")[[1]]
        # Check if Model is not empty
        if (selected_data$data$Model[i] != "") {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if (j == length(names)) {
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
              maker_table <- rbind(maker_table, c("", selected_data$data$Model[i], "", "", "", ""))
            } else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        } else {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if (j == length(names)) {
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
            } else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        }
      } 
      
      # Set column names
      colnames(maker_table) <- c("No", "Description", "Quantity", "Unit", "UnitP", "TotalP")
      
      # Verify the class and formatting
      class(maker_table$UnitP) <- "comma"
      class(maker_table$TotalP) <- "comma"
      
      
      # Create a style
      table_style <- createStyle(
        fontSize = 12,
        fontName = "Arial",
        border = "TopBottomLeftRight"
      )
      
      # Apply style to each cell in maker_table
      for (row_index in 1:nrow(maker_table)) {
        for (col_index in 1:ncol(maker_table)) {
          addStyle(my_workbook, sheet = 1, 
                   style = table_style, 
                   rows = 24 + row_index+PTlength, 
                   cols = 1 + col_index)
        }
      }
      
      
      
      # Write data to the worksheet
      writeData(
        my_workbook,
        sheet = 1,
        x = maker_table,
        startRow = 25+PTlength,
        startCol = 2,
        colNames = FALSE  # Suppress column headers
      )
      
      # Q'ty
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:5,
        rows = (25+PTlength):(24+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "TopBottom"
          
        ),
        gridExpand = TRUE
      )
      
      # No.
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2,
        rows = (25+PTlength):(24+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
        ),
        gridExpand = TRUE
      )
      
      
      # sub total
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Sub Total ", 5)
        ),
        startRow = 24+nrow(maker_table)+1+PTlength,
        startCol = 2
      )
      
      st <- sum(selected_data$data$Quantity*selected_data$data$價格)
      # sub total
      writeData(
        my_workbook,
        sheet = 1,
        x = paste0(format(st, scientific = FALSE, big.mark = ",")),
        startRow = 24+nrow(maker_table)+1+PTlength,
        startCol = 7,
        colNames = FALSE  # Suppress column headers
      )
      
      # sub total style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:7,
        rows = 24+nrow(maker_table)+1+PTlength,
        style = createStyle(
          fontSize = 12,
          halign = "right", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      if(selected_data$data$Tax[1]!=0){
        # tax
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(paste0("Tax(", selected_data$data$Tax[1]*100,"%)"), 5)
          ),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 2
        )
        
        # tax amount
        writeData(
          my_workbook,
          sheet = 1,
          x = sum(selected_data$data$Quantity*selected_data$data$價格*(selected_data$data$Tax[1])),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 7,
          colNames = FALSE  # Suppress column headers
        )
        
        # tax style
        addStyle(
          my_workbook,
          sheet = 1,
          cols = 2:7,
          rows = 24+nrow(maker_table)+1+1+PTlength,
          style = createStyle(
            fontSize = 12,
            halign = "right", 
            fontName = "Arial",
            border = "TopBottomLeftRight"
            
          ),
          gridExpand = TRUE
        )
      }
      
      else{
        # tax
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(paste0("Tax(税抜)"), 5)
          ),
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 2
        )
        
        # tax amount
        writeData(
          my_workbook,
          sheet = 1,
          x = "0",
          startRow = 24+nrow(maker_table)+1+1+PTlength,
          startCol = 7,
          colNames = FALSE  # Suppress column headers
        )
        
        # tax style
        addStyle(
          my_workbook,
          sheet = 1,
          cols = 2:7,
          rows = 24+nrow(maker_table)+1+1+PTlength,
          style = createStyle(
            fontSize = 12,
            halign = "right", 
            fontName = "Arial",
            border = "TopBottomLeftRight"
            
          ),
          gridExpand = TRUE
        )
      }
      
      
      # grand total
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list(paste0("Grand Total"), 5)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+PTlength,
        startCol = 2
      )
      gt <- sum(selected_data$data$Quantity*selected_data$data$價格*(1+selected_data$data$Tax[1]))
      # grand total amount
      writeData(
        my_workbook,
        sheet = 1,
        x = paste0(format(gt, scientific = FALSE, big.mark = ",")),
        startRow = 24+nrow(maker_table)+1+1+1+PTlength,
        startCol = 7,
        colNames = FALSE  # Suppress column headers
      )
      
      # grand total style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:7,
        rows = 24+nrow(maker_table)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 12,
          halign = "right", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      
      # remark
      
      if (is.na(selected_data$data$Remark[1])||selected_data$data$Remark[1] == ""){
        split_maker_remark <- ""
        
      }
      else{
        
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("Remark:"),
          startRow = 24+nrow(maker_table)+1+1+1+1+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        
        
        split_maker_remark <- unlist(strsplit(selected_data$data$Remark[1], "\\s(?=\\d+\\. )", perl = TRUE))
        # remark content
        writeData(
          my_workbook,
          sheet = 1,
          x = paste("    ", split_maker_remark),
          startRow = 24+nrow(maker_table)+1+1+1+1+1+PTlength,
          startCol = 2,
        )
        
      }
      
      
      
      # signature
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Gluon Japan Corporation", 4)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+1+1+PTlength,
        startCol = 4
      )
      
      
      # signature style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:7,
        rows = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 18,
          halign = "left", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      
      # underline style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:7,
        rows = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+6+1+1+PTlength,
        style = createStyle(
          border = "Bottom"
        ),
        gridExpand = TRUE
      )
      
      # shinga chen
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Shinya luchi President", 3)
        ),
        startRow = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+6+1+1+1+PTlength,
        startCol = 4
      )
      
      # shinga chen style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:7,
        rows = 24+nrow(maker_table)+1+1+1+length(split_maker_remark)+1+6+1+1+1+PTlength,
        style = createStyle(
          fontSize = 11,
          halign = "left"
        ),
        gridExpand = TRUE
      )
      
      
      saveWorkbook(my_workbook, file)
    }
    
  )
  
  
  output$download_maker_ce <- downloadHandler(
    
    filename = function() {
      paste0(as.character(input$no), "訂購單.xlsx")
    },
    content = function(file){
      my_workbook <- createWorkbook()
      
      addWorksheet(
        wb = my_workbook,
        sheetName = paste0(as.character(input$no))
      )
      
      setColWidths(
        my_workbook,
        1,
        cols = 1:7,
        widths =  c(1.67, 4.5, 46, 6, 6,16.5, 18.5)
        
      )
      
      setRowHeights(
        my_workbook,
        sheet = 1,  # Specify the sheet number or name
        rows = 1:6,  # Specify the range of rows
        heights = c(20,35, 25, 20, 20, 30)  # Convert your desired heights to Excel units
      )
      
      insertImage(
        my_workbook, 
        sheet = 1,
        file = "www/CTS_logo.png",
        width = 2.1,
        height = 1.04,
        startRow = 1,
        startCol = 2,
        units = "in",
        dpi = 500
      )
      
      # company name CH
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("无锡群唐信电子科技有限公司", 5)
        ),
        startRow = 2,
        startCol = 3
      )
      
      
      # company name CH style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 3:7,
        rows = 2,
        style = createStyle(
          fontSize = 22,
          halign = "right", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # company name EN
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Wuxi Chun Tang Shin Electronics Technology Co. Ltd.", 5)
        ),
        startRow = 3,
        startCol = 3
      )
      
      
      # company name EN style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 3:7,
        rows = 3,
        style = createStyle(
          fontSize = 18,
          halign = "right", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      # purchase order
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Purchase Order", 7)
        ),
        startRow = 8,
        startCol = 1
      )
      
      
      # po style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 1:7,
        rows = 8,
        style = createStyle(
          textDecoration = "Bold",
          fontSize = 24,
          halign = "center", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # No
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 7,
        startCol = 2,
        x = paste0("NO: ", as.character(input$no))
      )
      
      # No style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 7,
        cols = 2,
        style = createStyle(
          
          fontSize = 12,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      # date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 10,
        startCol = 7,
        x = paste0(selected_data$data$日期[1])
      )
      
      # date style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 10,
        cols = 7,
        style = createStyle(
          halign = "right",
          fontSize = 12,
          fontName = "Arial"
        )
      )
      
      
      
      # maker name
      if (selected_data$data$Maker[1] %in% maker_name$廠商簡稱) {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 11,
          startCol = 2,
          x = paste0("TO: ",maker_name$廠商名稱[maker_name$廠商簡稱==selected_data$data$Maker[1]])
        )
      } else {
        writeData(
          my_workbook,
          sheet = 1,
          startRow = 11,
          startCol = 2,
          x = paste0("TO: ",selected_data$data$Maker[1])
        )
      }
      
      
      # style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 11,
        cols = 2,
        style = createStyle(
          fontSize = 12,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      # payment 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 12,
        startCol = 2,
        x = paste0("Payment Terms: ")
      )
      
      
      # payment terms
      paymentT <- strsplit(selected_data$data$Payment_Terms[1], " / ")[[1]]
      PTlength <- length(paymentT) -1
      # payment terms
      for(i in seq_len(length(paymentT))){
        if(i==1){
          writeData(
            my_workbook,
            sheet = 1,
            startRow = 12,
            startCol = 2,
            x = paste0("Payment Terms: ", paymentT[1])
          )
        }
        else{
          writeData(
            my_workbook,
            sheet = 1,
            startRow = 12+i-1,
            startCol = 2,
            x = paste0("                              ", paymentT[i])
          )
        }
      }
      
     
      
      # Shipment date
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 13+PTlength,
        startCol = 2,
        x = paste0("Shipment Date: ", selected_data$data$Shipment_Date[1])
      )
      
      
      # destination 
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 14+PTlength,
        startCol = 2,
        x = paste0("Destination: ", selected_data$data$Destination[1])
      )
      
      
      
      # B style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 12:(14+PTlength),
        cols = 2:3,
        style = createStyle(
          fontSize = 12,
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 16+PTlength,
        startCol = 2,
        x = "No. ",
        
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 16+PTlength,
        startCol = 3,
        x = "Description",
        
      )
      
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("Q'ty", 2)
        ),
        startRow = 16+PTlength,
        startCol = 4
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 16+PTlength,
        startCol = 6,
        x = paste0("Unit Price (", selected_data$data$幣別[1], ")"),
        
      )
      
      writeData(
        my_workbook,
        sheet = 1,
        startRow = 16+PTlength,
        startCol = 7,
        x = paste0("Amount (", selected_data$data$幣別[1], ")"),
        
      )
      
      # table style
      addStyle(
        my_workbook,
        sheet = 1,
        rows = 16+PTlength,
        cols = 2:7,
        style = createStyle(
          fontSize = 12,
          fontName = "Arial",
          halign = "center",
          border = "TopBottomLeftRight"
        ),
        gridExpand = TRUE
      )
      
      # Initialize a list to store the rows of the final data frame
      maker_table <- data.frame(
        No = integer(),
        Description = character(),
        Quantity = numeric(),
        Unit = character(),
        UP = numeric(),
        Am = numeric(),
        stringsAsFactors = FALSE  # Ensure character columns are not converted to factors
      )
      
      # Loop through selected_data$data
      for (i in seq_len(nrow(selected_data$data))) {
        # Get names after splitting
        names <- strsplit(selected_data$data$名稱[i], " / ")[[1]]
        # Check if Model is not empty
        if (selected_data$data$Model[i] != "") {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if(j == 1){
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
            }
            else if (j == length(names)) {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
              maker_table <- rbind(maker_table, c("", selected_data$data$Model[i], "", "", "", ""))
              maker_table <- rbind(maker_table, c("", "", "", "", "", ""))
            } else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        } else {
          # Loop through each element after splitting
          for (j in seq_along(names)) {
            # Append row for each element after splitting
            if (j ==1) {
              maker_table <- rbind(maker_table, c(i, names[j], selected_data$data$Quantity[i], selected_data$data$單位[i], selected_data$data$價格[i], selected_data$data$價格[i] * selected_data$data$Quantity[i]))
            } 
            else if(j ==length(names)){
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
              maker_table <- rbind(maker_table, c("", "", "", "", "", ""))
            }
            else {
              maker_table <- rbind(maker_table, c("", names[j], "", "", "", ""))
            }
          }
        }
      } 
      
      # Set column names
      colnames(maker_table) <- c("No", "Description", "Quantity", "Unit", "UnitP", "TotalP")
      
      # Verify the class and formatting
      class(maker_table$UnitP) <- "comma"
      class(maker_table$TotalP) <- "comma"
      
      
      # Create a style
      table_style <- createStyle(
        fontSize = 12,
        fontName = "Arial",
        border = "TopBottomLeftRight"
      )
      
      # Apply style to each cell in maker_table
      for (row_index in 1:nrow(maker_table)) {
        for (col_index in 1:ncol(maker_table)) {
          addStyle(my_workbook, sheet = 1, 
                   style = table_style, 
                   rows = 16 + row_index+PTlength, 
                   cols = 1 + col_index)
        }
      }
      
      
      
      # Write data to the worksheet
      writeData(
        my_workbook,
        sheet = 1,
        x = maker_table,
        startRow = 17+PTlength,
        startCol = 2,
        colNames = FALSE  # Suppress column headers
      )
      
      # Q'ty
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 4:5,
        rows = (17+PTlength):(16+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "Top"
          
        ),
        gridExpand = TRUE
      )
      
      # No.
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2,
        rows = (17+PTlength):(16+nrow(maker_table)+PTlength),
        style = createStyle(
          fontSize = 12,
          halign = "center", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
        ),
        gridExpand = TRUE
      )
      
     
     # grand total without tax 
     
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list(paste0("Grand Total"), 5)
        ),
        startRow = 16+nrow(maker_table)+1+PTlength,
        startCol = 2
      )
      gt <- sum(selected_data$data$Quantity*selected_data$data$價格)
      # grand total amount
      writeData(
        my_workbook,
        sheet = 1,
        x = paste0(format(gt, scientific = FALSE, big.mark = ",")),
        startRow = 16+nrow(maker_table)+1+PTlength,
        startCol = 7,
        colNames = FALSE  # Suppress column headers
      )
      
      # grand total style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:7,
        rows = 16+nrow(maker_table)+1+PTlength,
        style = createStyle(
          fontSize = 12,
          halign = "right", 
          fontName = "Arial",
          border = "TopBottomLeftRight"
          
        ),
        gridExpand = TRUE
      )
      
      
      # remark
      
      if (is.na(selected_data$data$Remark[1])||selected_data$data$Remark[1] == ""){
        split_maker_remark <- ""
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("Remark:"),
          startRow = 16+nrow(maker_table)+2+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("1. 发票寄送地址：无锡市梁溪区广南路311号1008  ，  TEL:13348109589，闵会计（收）"),
          startRow = 16+nrow(maker_table)+3+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("2. 我方公司开票资料: "),
          startRow = 16+nrow(maker_table)+4+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","公司名称: 无锡群唐信电子科技有限公司"),
          startRow = 16+nrow(maker_table)+5+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","纳税识别号: 91320211MA1X2HY882"),
          startRow = 16+nrow(maker_table)+6+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","公司地址、电话: 无锡市滨湖区隐秀路813-602-2 18201951337"),
          startRow = 16+nrow(maker_table)+7+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","开户行及帐号: 招商银行股份有限公司无锡学前支行   510904290910101"),
          startRow = 16+nrow(maker_table)+8+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        
      }
      else{
        
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("Remark:"),
          startRow = 16+nrow(maker_table)+2+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("1. 发票寄送地址：无锡市梁溪区广南路311号1008  ，  TEL:13348109589，闵会计（收）"),
          startRow = 16+nrow(maker_table)+3+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("2. 我方公司开票资料: "),
          startRow = 16+nrow(maker_table)+4+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","公司名称: 无锡群唐信电子科技有限公司"),
          startRow = 16+nrow(maker_table)+5+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","纳税识别号: 91320211MA1X2HY882"),
          startRow = 16+nrow(maker_table)+6+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","公司地址、电话: 无锡市滨湖区隐秀路813-602-2 18201951337"),
          startRow = 16+nrow(maker_table)+7+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        writeData(
          my_workbook,
          sheet = 1,
          x = paste0("   ","开户行及帐号: 招商银行股份有限公司无锡学前支行   510904290910101"),
          startRow = 16+nrow(maker_table)+8+PTlength,
          startCol = 2,
          colNames = FALSE  # Suppress column headers
        )
        
        
        
        split_maker_remark <- unlist(strsplit(selected_data$data$Remark[1], "(\\d+\\. )", perl = TRUE))[-1]
        for (i in seq_along(split_maker_remark)) {
          split_maker_remark[i] <- paste(i + 2, ". ", split_maker_remark[i], sep = "")
        }
        
        
        # remark content
        writeData(
          my_workbook,
          sheet = 1,
          x = paste(split_maker_remark),
          startRow = 16+nrow(maker_table)+9+PTlength,
          startCol = 2,
        )
        
      }
      
      
      
      
      
      # signature style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 5:7,
        rows = 16+nrow(maker_table)+9+length(split_maker_remark)+1+1+1+PTlength,
        style = createStyle(
          fontSize = 15,
          halign = "left", 
          fontName = "Arial"
        ),
        gridExpand = TRUE
      )
      
      
      
      # underline style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 5:7,
        rows = 16+nrow(maker_table)+9+length(split_maker_remark)+1+6+1+1+PTlength,
        style = createStyle(
          border = "Bottom"
        ),
        gridExpand = TRUE
      )
      
      # 無錫
      excel_helpers$create_header_row(
        my_workbook,
        1,
        list(
          list("无锡群唐信电子科技有限公司", 3)
        ),
        startRow = 16+nrow(maker_table)+9+length(split_maker_remark)+1+6+1+1+1+PTlength,
        startCol = 5
      )
      
      # 無錫 style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 5:7,
        rows = 16+nrow(maker_table)+9+length(split_maker_remark)+1+6+1+1+1+PTlength,
        style = createStyle(
          fontSize = 10,
          halign = "left"
        ),
        gridExpand = TRUE
      )
      
      # 無錫logo
      insertImage(
        my_workbook, 
        sheet = 1,
        file = "www/CTS_stamp.png",
        width = 2.43,
        height = 2.83,
        startRow = 16+nrow(maker_table)+9+length(split_maker_remark)+1+1+1+1+PTlength,
        startCol = 5,
        units = "in",
        dpi = 300
      )
      
      
      # maker underline style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:3,
        rows = 16+nrow(maker_table)+9+length(split_maker_remark)+1+6+1+1+PTlength,
        style = createStyle(
          border = "Bottom"
        ),
        gridExpand = TRUE
      )
      
      
      # maker signature
      if (selected_data$data$Maker[1] %in% maker_name$廠商簡稱) {
        
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(maker_name$廠商名稱[maker_name$廠商簡稱==selected_data$data$Maker[1]], 2)
          ),
          startRow = 16+nrow(maker_table)+9+length(split_maker_remark)+1+6+1+1+1+PTlength,
          startCol = 2
        )
        
        
      } else {
        excel_helpers$create_header_row(
          my_workbook,
          1,
          list(
            list(selected_data$data$Maker[1], 2)
          ),
          startRow = 16+nrow(maker_table)+9+length(split_maker_remark)+1+6+1+1+1+PTlength,
          startCol = 2
        )
        
      }
      
      # maker name style
      addStyle(
        my_workbook,
        sheet = 1,
        cols = 2:3,
        rows = 16+nrow(maker_table)+9+length(split_maker_remark)+1+6+1+1+1+PTlength,
        style = createStyle(
          fontSize = 10,
          halign = "left"
        ),
        gridExpand = TRUE
      )
      
      
    
      
      
      
      saveWorkbook(my_workbook, file)
    }
    
  )
  
  
  observeEvent(input$update_maker_po, {
    if (nrow(selected_data$data)!=0){
      maker_po_data$data <- maker_po_data$data[maker_po_data$data$maker_order_number != input$no, ]
      maker_po_data$data <- rbind(maker_po_data$data, selected_data$data)
      write.xlsx(maker_po_data$data,"Data/Maker_PO.xlsx")
      sendSweetAlert(
        session = session,
        title = "Saved!!",
        text = "儲存成功",
        type = "success"
      )
      
      #maker_po_data$data <- load_maker_data()
      #original_choices_no <- unique(maker_po_data$data$maker_order_number)
      #updateSelectizeInput(session, "no", choices = c(original_choices_no), selected = selected_data$data$maker_order_number[1])

    }
    else{
      maker_po_data$data <- maker_po_data$data[maker_po_data$data$maker_order_number != input$no, ]
      write.xlsx(maker_po_data$data,"Data/Maker_PO.xlsx")
      sendSweetAlert(
        session = session,
        title = "Saved!!",
        text = "儲存成功",
        type = "success"
      )
      maker_po_data$data <- load_maker_data()
      original_choices_no <- unique(maker_po_data$data$maker_order_number)
      updateSelectizeInput(session, "no", choices = c("", original_choices_no))
      
    }
    
    
    
    
    
    
  })
  
  
  
  
  





