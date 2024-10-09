library(officer)
library(ggplot2)
library(tidyverse)
library(rvg)



# Function for the template Title + Big Text

# Define a function to add a slide and populate the content
add_Title_Text <- function(ppt, title, text1 = Sys.Date()) {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
  
  
  ppt %>%
    add_slide(layout = "TitleandText", master = "HQ Master Style Slide") %>%
    ph_with(value = title, location = ph_location_label(ph_label = "Text Placeholder 8")) %>%
    ph_with(value = text1, location = ph_location_label(ph_label = "Text Placeholder 10"))
    
    
# Save the updated PowerPoint
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 

    
}

# Example usage

ppt <- add_Title_Text(ppt, 
                      title = "Second test being done", 
                      text1 = "This is a second slide template")
