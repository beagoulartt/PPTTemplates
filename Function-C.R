library(officer)
library(magrittr)  # For the pipe operator %>%


# Function for the template Title + Big Text + Picture

# Define a function to add a slide and populate the content
add_Title_Text_Picture <- function(ppt, title, text1 = Sys.Date(), picture_path) {
  
  # Open the PowerPoint presentation
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
  
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
  print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
  
  return(ppt)
}

# Specify the path to the picture file
picture_file_path <- "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\HQ-image.jpg"

# Call the function with the title, text, and picture
ppt <- add_Title_Text_Picture(
  ppt, 
  title = "Third test being done", 
  text1 = "This is a function that uses Title, Text and a side image",
  picture_path = picture_file_path  # Pass the picture path as an argument
)