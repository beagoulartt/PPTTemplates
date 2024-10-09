# Function for the template Title + Graphic Image + Legend
add_Title_Table_Legend <- function(ppt, title, table_path, legend = Sys.Date()) {
  # Function for the template Title + Graphic Image + Legend
  add_Title_Table_Legend <- function(ppt, title, table_path, legend = Sys.Date()) {
    
    # Open the PowerPoint presentation
    ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
    
    # Add a new slide and populate placeholders
    ppt <- ppt %>%
      add_slide(layout = "Title/Table/Legend", master = "HQ Master Style Slide") %>%
      
      # Add title text to the placeholder labeled "Title"
      ph_with(value = title, location = ph_location_label(ph_label = "Title")) %>%
      
      # Add legend to the placeholder labeled "Legend"
      ph_with(value = legend, location = ph_location_label(ph_label = "Legend")) %>%
      
      # Add the picture (table image) to the placeholder labeled "Picture"
      ph_with(value = external_img(table_path), location = ph_location_label(ph_label = "Picture"))
    
    # Save the updated PowerPoint
    print(ppt, target = "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx") 
    
    return(ppt)
  }
  
  # Initialize the PowerPoint object
  ppt <- read_pptx("C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\SH Board Meeting Slide Deck - Updated.pptx")
  
  # Specify the path to the table file
  table_file_path <- "C:\\Users\\DataIntern\\HQToronto\\Shared Docs - General\\Clinical Reporting\\ReportingProjects\\DataIntern\\PPTTemplates\\TemplateProject\\Table.png"
  
  # Call the function with the title, text, and picture
  ppt <- add_Title_Table_Legend(
    ppt, 
    title = "Name of the Table", 
    legend = "4,725 visits in 2022\n27,025 visits in 2023\n7,699 visits",  # Added newline character for formatting
    table_path = table_file_path  # Use 'table_path' instead of 'graphic_path'
  )