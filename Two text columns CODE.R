library(officer)
library(ggplot2)
library(tidyverse)
library(rvg)

library(magrittr)

# layout = "Two text columns"
# Template for title + text1 + text2 + date_label

# Define a function to add a slide and populate the content
add_custom_slide <- function(ppt, title, text1, text2, date_label = Sys.Date()) {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\Template.pptx")
  
  ppt %>%
    add_slide(layout = "Two text columns", master = "HQ Master Style Slide") %>%
    ph_with(value = title, location = ph_location_label(ph_label = "BigTitle")) %>%
    ph_with(value = text1, location = ph_location_label(ph_label = "Text1")) %>%
    ph_with(value = text2, location = ph_location_label(ph_label = "Text2")) %>%
    ph_with(value = format(date_label, "%B %d, %Y"), location = ph_location_label(ph_label = "Date"))
  
  
  # Save the updated PowerPoint
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\Two text columns.pptx") 
  
  
}

# Example usage
ppt <- read_pptx() # Reading the ppt that I want to add texts
ppt <- add_custom_slide(ppt, 
                        title = "Sexual Health Key Messages", 
                        text1 = "Highest weekly average last week
178 sexual health visits / day
Highest number of daily service users on Monday July 8, 2024 at 220 users
New visits have stabilized.
Our testing is significantly less expensive than Public Health. ", 
                        text2 = "Successful  Code for Canada Hackathon
Idea 1:
Improve kiosk check-in process
Idea 2:
Automate mass sending of all negative results (80% of patients) ")
