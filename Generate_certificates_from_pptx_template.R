#install.packages("officer")
#install.packages("docxtractr")
library(officer)
library(docxtractr)

setwd("/Users/shouben/Documents/R/Medium")

# Load the PowerPoint template
template <- read_pptx("Diploma.pptx")

# Load the participants CSV file
participants <- read.csv("participants_list.csv")

# Create a new output folder for PDF files
output_folder <- "output"
dir.create(output_folder, showWarnings = FALSE)
output_folder_pdf <- file.path(output_folder, "pdf") # Create pdf_output as a subfolder
dir.create(output_folder_pdf, showWarnings = FALSE)

# Iterate over each row in the participants data frame
for (i in 1:nrow(participants)) {
  participant <- participants[i, ]
  
  slide_layout <- "layout_gis" # Name of the pptx layout (in Slide master in office, see the name)
  
  presentation <- read_pptx("Diploma.pptx") # So each pptx is composed of only one slide (recall of the blank template)
  
  slide <- add_slide(presentation, layout = slide_layout) # It is important that the base pptx is empty. Only layout created in Slide master)
  
  view <- layout_properties(slide) # See the placeholder's names and other layout properties. Here, important to be sure the placeholder label are adapted and different for each location (see in the pptx layout and adapt)
  
  ph <- ph_with(slide, type = "body", value = participant$Name, location = ph_location_label(ph_label = "Name"))
  ph <- ph_with(slide, type = "body", value = participant$title_diploma, location = ph_location_label(ph_label = "Title_diploma"))
  ph <- ph_with(slide, type = "body", value = participant$institution, location = ph_location_label(ph_label = "institution"))
  ph <- ph_with(slide, type = "body", value = participant$date, location = ph_location_label(ph_label = "date"))

  
  # Save the PowerPoint presentation with a unique name
  filename <- paste0("output/Diploma_", participant$Name,".pptx")
  print(ph, target = filename)
  
  # Convert the PPTX file to PDF using docxtractr and directly send them in the output folder
  pdf_filename <- paste0("output/pdf/Diploma_", participant$Name, ".pdf")
  convert_to_pdf(filename,pdf_filename)
}

