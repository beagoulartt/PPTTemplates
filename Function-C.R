library(officer)
library(magrittr)  # For the pipe operator %>%

# Function for the template Title + Big Text + Picture on the right side
# Define a function to add a slide and populate the content
add_Title_Text_Picture <- function(title, text1 = Sys.Date(), picture_path) {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\Template.pptx")
  
  # Add a new slide and populate placeholders
  ppt <- ppt %>%
    add_slide(layout = "Title/Text/Picture", master = "HQ Master Style Slide") %>%
    
    # Add title text to the placeholder labeled "Title"
    ph_with(value = title, location = ph_location_label(ph_label = "Title")) %>%
    
    # Add text to the placeholder labeled "BigText1"
    ph_with(value = text1, location = ph_location_label(ph_label = "BigText1")) %>%
    
    # Add the picture to the placeholder labeled "Picture"
    ph_with(value = external_img(picture_path), location = ph_location_label(ph_label = "Picture"))
  
  # Save the updated PowerPoint
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\TitleTextPicture.pptx")
  
  return(ppt)
}

# Specify the path to the picture file
picture_file_path <- "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\sidepic.png"

# Call the function with the title, text, and picture
ppt <- add_Title_Text_Picture(
  title = "The problem", 
  text1 = "What is the problem?
Tasks nurses spend a lot of time sending negative test results to patients, which could be better used for delivering care to patients.,

Why is this important?
Lead to increased employee burnout
Reduced patient care
Results not shared quickly
Reduced level of patient care",
  picture_path = picture_file_path  # Pass the picture path as an argument
)
