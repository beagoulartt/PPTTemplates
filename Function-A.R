library(officer)
library(ggplot2)
library(tidyverse)
library(rvg)

library(magrittr)


# Template for title + text1 + text2 + date_label

# Define a function to add a slide and populate the content
add_custom_slide <- function(ppt, title, text1, text2, date_label = Sys.Date()) {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx")
  
  ppt %>%
    add_slide(layout = "Two text columns", master = "HQ Master Style Slide") %>%
    ph_with(value = title, location = ph_location_label(ph_label = "BigTitle")) %>%
    ph_with(value = text1, location = ph_location_label(ph_label = "Text1")) %>%
    ph_with(value = text2, location = ph_location_label(ph_label = "Text2")) %>%
    ph_with(value = format(date_label, "%B %d, %Y"), location = ph_location_label(ph_label = "Date"))
  
  
  # Save the updated PowerPoint
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
  
  
}

# Example usage
ppt <- read_pptx() # Assuming you have already created a PowerPoint object
ppt <- add_custom_slide(ppt, 
                        title = "Testing this new object thing", 
                        text1 = "It looks greats let's see if it works!!!", 
                        text2 = "Blablablablabla22222")
