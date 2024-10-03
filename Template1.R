library(officer)
library(ggplot2)

# Open the PowerPoint presentation
ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Template.pptx")

# Select the 3rd slide
ppt <- on_slide(ppt, index = 3)

slide_summary(ppt, index = 3)

View(layout_properties(ppt, layout = "Two text columns"))



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
 
 
 
 # Add a new slide with a specific layout
 ppt <- add_slide(ppt, layout = "Two text columns", master = "HQ Master Style Slide")

# Save the updated PowerPoint
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx")




# Select the 39 slide
ppt <- on_slide(ppt, index = 39)

# Clear and modify the title using its type "BigTitle"
ppt <- ph_with(ppt, value = "Testing new slide added", location = ph_location_label(ph_label = "BigTitle"))


# Save the updated PowerPoint
print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx")


slide_summary(ppt)
