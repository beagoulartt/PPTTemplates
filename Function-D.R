library(officer)
library(magrittr)  # For the pipe operator %>%


# Function for the template Title + Graphic Image + Legend

# Define a function to add a slide and populate the content
add_Title_Graphic_Legend <- function(ppt, title, graphic_path, legend = Sys.Date()) {

  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
  
  # Add a new slide and populate placeholders
  ppt <- ppt %>%
    add_slide(layout = "Title/Graphic/Legend", master = "HQ Master Style Slide") %>%
    
    # Add title text to the placeholder labeled "Title"
    ph_with(value = title, location = ph_location_label(ph_label = "Title")) %>%
    
    # Add Graphic to the placeholder labeled "Legend"
    ph_with(value = legend, location = ph_location_label(ph_label = "Legend")) %>%
    
    # Add the picture to the placeholder labeled "Picture"
    ph_with(value = external_img(graphic_path), location = ph_location_label(ph_label = "Graphicpicture"))
  
  # Save the updated PowerPoint
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
  
  return(ppt)
}

# Specify the path to the graphic file
graphic_file_path <- "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\thisPlot_transparent.png"

# Call the function with the title, text, and picture
ppt <- add_Title_Graphic_Legend(
  ppt, 
  title = "First vs Return
 July 2023 - June 2024", 
  legend = "04,725 visits in 2022
27,025 visits in 2023
18,986 visits in 2024
",
  graphic_path = graphic_file_path
)