library(officer)
library(ggplot2)
library(tidyverse)
library(rvg)

# Open the PowerPoint presentation
ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Template.pptx")

# Select the 3rd slide
ppt <- on_slide(ppt, index = 3)

#slide_summary(ppt, index = 3)

#View(layout_properties(ppt, layout = "Two text columns"))



# Clear and modify the title using its type "BigTitle"
ppt <- ph_remove(ppt, type = "title")
ppt <- ph_with(ppt, value = "Testing Template", location = ph_location_label(ph_label = "BigTitle"))

# Clear and modify text in text box 1 using its label "Text1"
ppt <- ph_remove(ppt, ph_label = "Text1")
ppt <- ph_with(ppt, value = "This is the first box of text, testing it right now!", location = ph_location_label(ph_label = "Text1"))

# Clear and modify text in text box 2 using its label "Text2"
ppt <- ph_remove(ppt, ph_label = "Text2")
ppt <- ph_with(ppt, value = "Updated text 2", location = ph_location_label(ph_label = "Text2"))

# Clear and modify the date using its label "Date"
ppt <- ph_remove(ppt, ph_label = "Date")
ppt <- ph_with(ppt, value = format(Sys.Date(), "%B %d, %Y"), location = ph_location_label(ph_label = "Date"))

# Save the updated PowerPoint
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx")
 
 

 
# Add a new slide with a specific layout
ppt <- add_slide(ppt, layout = "Two text columns", master = "HQ Master Style Slide")

# Select the 39 slide
ppt <- on_slide(ppt, index = 41)

# Add title using its type "BigTitle"
ppt <- ph_with(ppt, value = "Testing new slide added", location = ph_location_label(ph_label = "BigTitle"))


# Add text in text box 1 using its label "Text1"
ppt <- ph_with(ppt, value = "This is exciting!!!", location = ph_location_label(ph_label = "Text1"))

# Add text in text box 2 using its label "Text2"
ppt <- ph_with(ppt, value = "Blablablablabla", location = ph_location_label(ph_label = "Text2"))

# Clear and modify the date using its label "Date"
ppt <- ph_with(ppt, value = format(Sys.Date(), "%B %d, %Y"), location = ph_location_label(ph_label = "Date"))

# Save the updated PowerPoint
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx")




# **************************************
# Hommood: Added code below to simplify adding a slide
# Add a new slide with a specific layout
ppt <- add_slide(ppt, layout = "Two text columns", master = "HQ Master Style Slide") %>% 
  ph_with(ppt, value = "Testing new slide added", location = ph_location_label(ph_label = "BigTitle")) %>% 
  ph_with(ppt, value = "This is exciting!!!", location = ph_location_label(ph_label = "Text1")) %>% # Add text in text box 2 using its label "Text2"
  ph_with(ppt, value = "Blablablablabla", location = ph_location_label(ph_label = "Text2")) %>% # Clear and modify the date using its label "Date"
  ph_with(ppt, value = format(Sys.Date(), "%B %d, %Y"), location = ph_location_label(ph_label = "Date"))

# Create a ggplot chart
gg <- ggplot(mtcars, aes(x = wt, y = mpg)) + geom_point()

# Add the ggplot chart as a vector graphic (editable in PowerPoint)
ppt <- ppt %>% add_slide(layout = "Title and Content", master = "Office Theme") %>%
  ph_with(value = dml(ggobj = gg), location = ph_location_type(type = "body"))


# **************************************

# Save the updated PowerPoint
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx")



#library(officer)
#library(magrittr)


# Define a function to add a slide and populate the content
add_custom_slide <- function(ppt, title, text1, text2, date_label = Sys.Date()) {
  
  # Open the PowerPoint presentation
ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Template.pptx")
  
  ppt %>%
    add_slide(layout = "Two text columns", master = "HQ Master Style Slide") %>%
    ph_with(value = title, location = ph_location_label(ph_label = "BigTitle")) %>%
    ph_with(value = text1, location = ph_location_label(ph_label = "Text1")) %>%
    ph_with(value = text2, location = ph_location_label(ph_label = "Text2")) %>%
    ph_with(value = format(date_label, "%B %d, %Y"), location = ph_location_label(ph_label = "Date"))
}

# Example usage
ppt <- read_pptx() # Assuming you have already created a PowerPoint object
ppt <- add_custom_slide(ppt, 
                        title = "Testing this new object thing", 
                        text1 = "It looks greats let's see if it works!!!", 
                        text2 = "Blablablablabla22222")

# Save the updated PowerPoint
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 



#layout_summary()


#This is the right one function template

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

ppt <- add_custom_slide(ppt, 
                        title = "Testing this new object thing", 
                        text1 = "It looks greats let's see if it works!!!", 
                        text2 = "Blablablablabla22222")



#Next function for another type of slide


# Define a function to add a slide and populate the content
add_Title_Text <- function(ppt, title, text1 = Sys.Date()) {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
  
  
  ppt %>%
    add_slide(layout = "TitleandText", master = "HQ Master Style Slide") %>%
    ph_with(value = title, location = ph_location_label(ph_label = "Text Placeholder 8")) %>%
    ph_with(value = text1, location = ph_location_label(ph_label = "Text Placeholder 10")) %>%
    
  # Save the updated PowerPoint
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
  
}

# Example usage

ppt <- add_Title_Text(ppt, 
                        title = "Second test being done", 
                        text1 = "This is a second slide template")






#layout_summary()