library(shiny)
library(shinyjs)
library(shinyalert)
library(readxl)
source("inmsp_fac.R")

email_input <- NULL
email_output <- NULL
email_file <- NULL
coiin_dir_input <- NULL
render_email_text <- FALSE

ui <- tagList(fluidPage(
               #theme = "bootstrap.css",
               #includeScript("./www/text.js"),
                useShinyalert(),
  sidebarLayout(
    sidebarPanel(
      fileInput("file1", "Choose facility data file",
                accept = c(
                  "application/vnd.ms-excel",
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  ".xls",
                  ".xlsx")
      ),
      dateRangeInput('dateRange',
                     label = 'Enter date range',
                     start = "", end = ""),
      actionButton(inputId = "genReports",
                   label = "Generate reports"),
      tags$hr(),
      fileInput("file2", "Choose facility email file",
                accept = c(
                  "application/vnd.ms-excel",
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  ".xls",
                  ".xlsx")),
      # tags$div(class="form-group shiny-input-container",
      #          tags$div(tags$label("File input")),
      #          tags$div(tags$label("Choose folder", class="btn btn-primary",
      #                              tags$input(id = "fileIn", webkitdirectory = TRUE, type = "file", style="display: none;", onchange="pressed()"))),
      #          tags$label("No folder choosen", id = "noFile"),
      #          tags$div(id="fileIn_progress", class="progress progress-striped active shiny-file-input-progress",
      #                   tags$div(class="progress-bar")
      #          )
      # ),
      #verbatimTextOutput("coiin_dir"),
      textInput("coiin_dir", "Specify COIIN directory"),
      textInput("emailId", "Specify 'from' email address"),
      passwordInput("password", "Password"),
      actionButton(inputId = "sendEmails", label = "Send emails")
    ),
    mainPanel(
      #dataTableOutput("tbl2"),
      tags$h3(textOutput("gen_reports")),
      tags$h3(textOutput("sent_emails"))
    )
  )
)
#,
#HTML("<script type='text/javascript' src='getFolders.js'></script>")
)

# validates the facility data input fields
validateFacilityData <- function(input) {
  if (is.null(input$file1)) {
    "Please choose a facility data file"
  } else if (is.null(input$dateRange)) {
    "Please enter a date range"
  } else {
    range <- as.Date(input$dateRange)
    if (length(range) != 2 || is.na(range[1]) || is.na(range[2]))
      "Please enter a date range"
    else
      NULL
  }
}

# validates the facility email input fields
validateEmailData <- function(input) {
  if (is.null(input$file2)) {
    "Please choose a facility email file"
  } else if (is.null(input$emailId) || trimws(input$emailId) == "") {
    "Please enter 'from' email address"
  } else if (is.null(input$password) || trimws(input$password) == "") {
    "Please enter password"
  } else {
      NULL
  }
}

showProgSendEmail <- function(send_email) {
  if (send_email) {
    # Create a Progress object
    progress <- shiny::Progress$new()
    progress$set(message = "Sending emails...", value = 0)
    # Close the progress when this reactive exits (even if there's an error)
    on.exit(progress$close())
  
    # Create a callback function to update progress.
    # Each time this is called:
    # - If `value` is NULL, it will move the progress bar 1/5 of the remaining
    #   distance. If non-NULL, it will set the progress to that value.
    # - It also accepts optional detail text.
    updateProgress <- function(value = NULL, detail = NULL) {
      if (is.null(value)) {
        value <- progress$getValue()
        value <- value + (progress$getMax() - value) / 5
      }
      progress$set(value = value, detail = detail)
    }
    email_message <- sendEmail(email_file$datapath, email_input$emailId, email_input$password,
                               coiin_dir = coiin_dir_input, updateProgress = updateProgress)
    # email_message <- sendEmail(email_file$datapath, email_input$emailId, email_input$password,
    #                            coiin_files = email_input$fileIn, updateProgress = updateProgress)
    if (render_email_text) {
      email_output$sent_emails <- renderText({
        email_message
      })
    } else {
      email_message
    }
  }
}

# Define server logic required to draw a histogram
server <- function(input, output, session) {
  reports <- eventReactive(input$genReports, {
    validate(validateFacilityData(input))
    # input$file1 will be NULL initially. After the user selects
    # and uploads a file, it will be a data frame with 'name',
    # 'size', 'type', and 'datapath' columns. The 'datapath'
    # column will contain the local filenames where the data can
    # be found.
    inFile <- input$file1
    if (is.null(inFile))
      return(NULL)
    
    # Create a Progress object
    progress <- shiny::Progress$new()
    progress$set(message = "Generating reports...", value = 0)
    # Close the progress when this reactive exits (even if there's an error)
    on.exit(progress$close())
    
    # Create a callback function to update progress.
    # Each time this is called:
    # - If `value` is NULL, it will move the progress bar 1/5 of the remaining
    #   distance. If non-NULL, it will set the progress to that value.
    # - It also accepts optional detail text.
    updateProgress <- function(value = NULL, detail = NULL) {
      if (is.null(value)) {
        value <- progress$getValue()
        value <- value + (progress$getMax() - value) / 5
      }
      progress$set(value = value, detail = detail)
    }
    
    generateReports(inFile$datapath, input$dateRange, updateProgress)
  })
  
  output$gen_reports <- renderText({
    reports()
  })
  
  emails <- eventReactive(input$sendEmails, {
    validate(validateFacilityData(input))
    validate(validateEmailData(input))

    # inFile <- input$file2
    # if (is.null(inFile))
    #   return(NULL)
    email_input <<- input
    email_file <<- input$file2
    coiin_dir_input <<- input$coiin_dir
    render_email_text <<- FALSE
    if (trimws(coiin_dir_input) == "") {
    #if (is.null(input$fileIn)) {
      coiin_dir_input <<- NULL
      render_email_text <<- TRUE
      email_output <<- output
      shinyalert(
        title = "No COIIN directory specified",
        text = "Are you sure you'd like to send emails without COIIN reports?",
        closeOnEsc = FALSE,
        closeOnClickOutside = FALSE,
        html = FALSE,
        type = "success",
        showConfirmButton = TRUE,
        showCancelButton = TRUE,
        confirmButtonText = "OK",
        cancelButtonText = "CANCEL",
        confirmButtonCol = "#AEDEF4",
        timer = 0,
        imageUrl = "",
        animation = FALSE,
        callbackR = showProgSendEmail
      )
      return(NULL)
    } else {
      showProgSendEmail(TRUE)
    }
  })
  
  output$gen_reports <- renderText({
    reports()
  })
  
  output$sent_emails <- renderText({
    emails()
  })
  
   output$coiin_dir <- renderPrint({
     input$mydata
   })
   
   # output$tbl2 <- DT::renderDataTable(
   #   input$fileIn
   # )
}

# Run the application 
shinyApp(ui = ui, server = server)

