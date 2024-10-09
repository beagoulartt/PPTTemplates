library(officer)
library(ggplot2)
library(tidyverse)
library(rvg)


#Ignore this one because it is the same as the Function A

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

# Save the updated PowerPoint
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx")

