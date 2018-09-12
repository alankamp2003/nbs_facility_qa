library(readxl)
library(rmarkdown)
library(knitr)
library(dplyr)
library(mailR)
library(futile.logger)
library(properties)

facility_dir <- "facility_files"
properties_file <- "facility.properties"

generateReports <- function(excel_file, dateRange, updateProgress = NULL) {
  df <- read_excel(excel_file)
  df <- data.frame(df)
  
  # replace one or more "."s in each column name with a single "_"
  # rename the columns with new names
  names <- lapply(colnames(df), mygrep)
  names <- unlist(names)
  colnames(df) <- names
  
  # get columns that start with "T_" and ones that don't;
  # the former are for all facilities and the latter for individual
  # facilities
  cols <- startsWith(names, 'T_')
  fac_cols <- names[!cols][3:20]
  tot_cols <- names[cols][1:18]
  
  # get the facility's values and the total values in a vector
  output_colnames <- c("Facility Totals", "Iowa Totals")
  output_rownames <- c("Initial", "Repeat", "Total", "Layered/clotted", 
                       "Quantity not sufficient", "Didn't soak through",
                       "Applied both sides", "Paper scratched", "Specimen age",
                       "Serum separated", "Contaminated", "Other", "Total Unsatisfactory",
                       "Rate (%) (goal is <1%)", "Unknown weight", "Unknown transfusion",
                       "Early collection <24 hours [1] ", "Transfused before collection [2]")
  
  # create the output directory if it isn't present; otherwise remove all files in it
  dir_name <- facility_dir
  if (!dir.exists(dir_name)) {
    dir.create(dir_name)
  } else {
    files <- dir(dir_name)
    for (i in 1 : length(files)) {
      file.remove(paste(dir_name,"/",files[i], sep = ""))
    }
  }
  for (i in 1:nrow(df)) {
  #for (i in 1:4) {
    # generate the output data frame for each facility
    fac_vals <- get_col_vals(df[i, fac_cols])
    tot_vals <- get_col_vals(df[i, tot_cols])
    output_df <- data.frame(fac_vals, tot_vals)
    colnames(output_df) <- output_colnames
    rownames(output_df) <- output_rownames
    
    org_id <- df[i,1]
    name <- df[i,2]

    # generate the output report file
    range <- as.Date(dateRange)
    beg <- format(range[1], format="%B %d %Y")
    end <- format(range[2], format="%B %d %Y")
    
    rmarkdown::render("inmsp_fac_rmd.Rmd", 
                      output_file = paste(org_id,"_qa.pdf", sep = ""),
                      output_dir = dir_name,
                      params = list(data = output_df,
                                    heading = paste("Quality Report for ", name, "(", org_id, ")"),
                                    caption = paste(beg, "through", end)))
    # If we were passed a progress update function, call it
    if (is.function(updateProgress)) {
      text <- paste0("Facility #", org_id)
      updateProgress(detail = text)
    }
  }
  # show the full path to the directory containing facility reports
  dir_name <- gsub('\\\\', '\\', normalizePath(dir_name), fixed = TRUE)
  return(paste("Reports generated in", dir_name))
}

# extracts the facility id from the file name at the passed index in the passed list of files;
# also removes split_str from the id 
getId <- function(files, index, split_str) {
  id <- strsplit(files[index], "\\.")[[1]][1]
  id <- strsplit(id, split_str)[[1]]
}

# This function sends emails. It can optionally take a function,
# updateProgress, which will be called as each email is sent.
sendEmail <- function(email_file, from_email, password, coiin_dir = NULL, updateProgress = NULL) {
  qual_dir_name <- facility_dir
  quality_files <- list.files(qual_dir_name)
  if (is.null(quality_files) || length(quality_files) == 0) {
    p <- normalizePath(getwd())
    error <- sprintf("Facility reports not found. They must be in a folder called '%s', under %s", qual_dir_name, p)
    return(error)
  }
  coiin_files <- NULL
  if (! is.null(coiin_dir)) {
    coiin_dir <- file.path(coiin_dir)
    coiin_files <- list.files(coiin_dir)
  }
  # validate email address to be of the right format;
  # currently only gmail.com and uiowa.edu are supported
  email_parts <- strsplit(from_email, "@")
  if (length (email_parts[[1]]) != 2)
    return("Please enter a valid email address")
  user_name <- email_parts[[1]][1]
  domain <- email_parts[[1]][2]
  ssl_arg <- FALSE
  tls_arg <- FALSE
  server <- NULL
  out_port <- NULL
  if (domain == "gmail.com") {
    server <- "smtp.gmail.com"
    out_port <- 465
    ssl_arg <- TRUE
  } else if (domain == "uiowa.edu") {
    server <- "smtp.office365.com"
    out_port <- 587
    user_name <- from_email
    tls_arg <- TRUE
  } else {
    return("The email address must have either gmail.com or uiowa.edu")
  }
  all_emails <- read_excel(email_file) %>% arrange(`Facility ID`)
  # go through the directories for quality and coiin reports and send emails to all addresses;
  # the names in the list of files are sorted alphabetically
  attach_files <- c()
  log_file <- "facility.log"
  if (file.exists(log_file))
    file.remove(log_file)
  else
    file.create(log_file)
  flog.appender(appender.file(log_file))
  if (!file.exists(properties_file))
    return(sprintf("File containing email subject and body not found. It must be under %s and called '%s'", normalizePath(getwd()), properties_file))
  props <- read.properties(properties_file)
  #print(props$subject)
  #print(props$body)
  for (i in 1:length(quality_files)) {
    id <- getId(quality_files, i, "_qa")
    fac_emails <- filter(all_emails, `Facility ID` == id) %>% select(Email)
    fac_emails <- fac_emails[[1]]
    # each facility can have multiple email addresses; send the quality and coiin reports
    # for a facility to all of them; a facility may or may not have a coiin report
    if (length(fac_emails) == 0) {
      flog.warn("No email address found for facility id %s", id)
      next
    }
    add_coiin_warn <- TRUE
    for (j in 1:length(fac_emails)) {
      attach_files <- c(paste(qual_dir_name, "/",quality_files[i], sep = ""))
      if (!is.null(coiin_files)) {
        found <- FALSE
        # go through the coiin directory and find the report that has the same facility id
        # as the quality report; if found, include it in the list of attachments; the names
        # in the list of files are sorted alphabetically
        ci <- 1
        while (ci <= length(coiin_files)) {
          if (id == getId(coiin_files, ci, "_coiin")) {
            found <- TRUE
            break
          } else {
            ci = ci + 1
          }
        }
        if (found) {
          attach_files <- c(attach_files, paste(coiin_dir,"/",coiin_files[ci], sep = ""))
        } else if (add_coiin_warn) {
          flog.warn("No COIIN report found for facility id %s", id)
          add_coiin_warn <- FALSE
        }
      } else if (add_coiin_warn) {
        flog.warn("No COIIN report found for facility id %s", id)
        add_coiin_warn <- FALSE
      }
      send.mail(from = from_email,
                to = c(fac_emails[j]),
                subject = props$subject,
                body = paste(props$line1, props$line2, props$line3, props$line4, sep = "\n\n"),
                smtp = list(host.name = server, port = out_port, user.name = user_name,
                            passwd = password, ssl = ssl_arg, tls = tls_arg),
                authenticate = TRUE,
                attach.files = attach_files,
                send = TRUE)
    }
    # If we were passed a progress update function, call it
    if (is.function(updateProgress)) {
      text <- paste0("Facility #", id)
      updateProgress(detail = text)
    }
  }
  return(sprintf("Sent emails. See %s for any warnings or errors.", normalizePath(properties_file)))
}

# This function sends emails. It can optionally take a function,
# updateProgress, which will be called as each email is sent.
sendEmail1 <- function(email_file, from_email, password, coiin_files = NULL, updateProgress = NULL) {
  qual_dir_name = "fac_files"
  quality_files <- list.files(qual_dir_name)
  # validate email address to be of the right format;
  # currently only gmail.com and uiowa.edu are supported
  email_parts <- strsplit(from_email, "@")
  if (length (email_parts[[1]]) != 2)
    return("Please enter a valid email address")
  user_name <- email_parts[[1]][1]
  domain <- email_parts[[1]][2]
  if (domain == "gmail.com") {
    smtp <- "smtp.gmail.com"
    port <- 465
  } else if (domain == "uiowa.edu") {
    server <- "smtp.office365.com"
    out_port <- 587
  } else {
    return("The email address must have either gmail.com or uiowa.edu")
  }
  all_emails <- read_excel(email_file) %>% arrange(`Facility ID`)
  # go through the directories for quality and coiin reports and send emails to all addresses;
  # the names in the list of files are sorted alphabetically
  attach_files <- c()
  log_file <- "facility.log"
  if (file.exists(log_file))
    file.remove(log_file)
  else
    file.create(log_file)
  flog.appender(appender.file(log_file))
  for (i in 1:length(quality_files)) {
    id <- getId(quality_files, i, "_qa")
    fac_emails <- filter(all_emails, `Facility ID` == id) %>% select(Email)
    fac_emails <- fac_emails[[1]]
    #print(email)
    # each facility can have multiple email addresses; send the quality and coiin reports
    # for a facility to all of them; a facility may or may not have a coiin report
    if (length(fac_emails) == 0)
      flog.warn("No email address found for facility id %s", id)
    add_coiin_warn <- TRUE
    for (j in 1:length(fac_emails)) {
      attach_files <- c(paste(qual_dir_name, "/",quality_files[i], sep = ""))
      if (!is.null(coiin_files)) {
        found <- FALSE
        # go through the coiin directory and find the report that has the same facility id
        # as the quality report; if found, include it in the list of attachments; the names
        # in the list of files are sorted alphabetically
        ci <- 1
        while (ci <= length(coiin_files)) {
          if (id == getId(coiin_files, ci, "_coiin")) {
            found <- TRUE
            break
          } else {
            ci = ci + 1  
          }
        }
        if (found) {
          attach_files <- c(attach_files, paste(coiin_dir,"/",coiin_files[ci], sep = ""))
        } else if (add_coiin_warn) {
          flog.warn("No COIIN report found for facility id %s", id)
          add_coiin_warn <- FALSE
        }
      } else if (add_coiin_warn) {
        flog.warn("No COIIN report found for facility id %s", id)
        add_coiin_warn <- FALSE
      }
      #       send.mail(from = from_email,
      #                 to = c(fac_emails[j]),
      #                 subject = "Newborn Screening Reports",
      #                 body = "Hello-
      #                 
      # Attached are your newborn screening reports.  Contact Ashley Comer by phone 515-725-1525 or email ashley-comer@uiowa.edu with any questions.
      #                 
      #                 
      # Thank you,
      #                 
      # Newborn Screening Staff",
      #                 #smtp = list(host.name = server, port = out_port, user.name = user_name, passwd = password, tls = TRUE)
      #                 smtp = list(host.name = server, port = out_port, user.name = from_email, passwd = password, tls = TRUE),
      #                 authenticate = TRUE,
      #                 attach.files = attach_files,
      #                 send = TRUE)
    }
    # If we were passed a progress update function, call it
    if (is.function(updateProgress)) {
      text <- paste0("Facility #", id)
      updateProgress(detail = text)
    }
  }
  return("Sent emails. See facility.log for any warnings or errors.")
}

# replaces one or more "."s in the argument with a single "_"
mygrep <- function(name) {
  return (gsub("\\.{1,}", "_", name))
}

#returns a vector populated with values in the passed data frame;
#this method is needed because the speadsheet's columns are not 
#in the same order as the rows in the report
get_col_vals <- function(sheet_vals) {
  vals <- c()
  vals[1] <- format(sheet_vals[1,2], nsmall = 0)
  vals[2] <- format(sheet_vals[1,3], nsmall = 0)
  vals[3] <- format(sheet_vals[1,1], nsmall = 0)
  vals[4] <- format(sheet_vals[1,8], nsmall = 0)
  vals[5] <- format(sheet_vals[1,9], nsmall = 0)
  vals[6] <- format(sheet_vals[1,7], nsmall = 0)
  vals[7] <- format(sheet_vals[1,5], nsmall = 0)
  vals[8] <- format(sheet_vals[1,11], nsmall = 0)
  vals[9] <- format(sheet_vals[1,4], nsmall = 0)
  vals[10] <- format(sheet_vals[1,10], nsmall = 0)
  vals[11] <- format(sheet_vals[1,6], nsmall = 0)
  vals[12] <- format(sheet_vals[1,12], nsmall = 0)
  vals[13] <- format(sheet_vals[1,13], nsmall = 0)
  vals[14] <- format(sheet_vals[1,14], nsmall = 2)
  vals[15] <- format(sheet_vals[1,18], nsmall = 0)
  vals[16] <- format(sheet_vals[1,17], nsmall = 0)
  vals[17] <- format(sheet_vals[1,15], nsmall = 0)
  vals[18] <- format(sheet_vals[1,16], nsmall = 0)
  
  return(vals)
}

#create a data frame from the passed row of the passed data frame
#getRowDF <- function(full_df, row, fac_cols, tot_cols, output_colnames, output_rownames) {
 
  
 # return(output_df)
#}