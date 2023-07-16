### Bulk certificates from pptx in R ###
# Creation of bulk certificates/diplomas with pptx template in R

First, install "officer" and "docxtractr" packages
```{r}
install.packages("officer")
install.packages("docxtractr")
```
Load the pptx template
```{r}
template <- read_pptx("Diploma.pptx")
```
Load the csv data
```{r}
participants <- read.csv("participants_list.csv")
```
Create output folders
  One for the pptx that will be generated:
```{r}
output_folder <- "output"
dir.create(output_folder, showWarnings = FALSE)
```
  One for the pdf that will be generated after the conversion of the pptx files
```{r}
output_folder_pdf <- file.path(output_folder, "pdf") 
dir.create(output_folder_pdf, showWarnings = FALSE)
```
This will create a subfolder in the output folder

We will now use a "for" loop to iterate over each row of the csv file to create a unique certificate for each participants
```{r}
for (i in 1:nrow(participants)) {
  participant <- participants[i, ]
  
  slide_layout <- "layout_diploma" 
  
  presentation <- read_pptx("Diploma.pptx") 
  
  slide <- add_slide(presentation, layout = slide_layout) 
  
  view <- layout_properties(slide) 
  
  ph <- ph_with(slide, type = "body", value = participant$Name, location = ph_location_label(ph_label = "Name"))
  ph <- ph_with(slide, type = "body", value = participant$title_diploma, location = ph_location_label(ph_label = "Title_diploma"))
  ph <- ph_with(slide, type = "body", value = participant$institution, location = ph_location_label(ph_label = "institution"))
  ph <- ph_with(slide, type = "body", value = participant$date, location = ph_location_label(ph_label = "date"))

  filename <- paste0("output/Diploma_", participant$Name,".pptx")
  print(ph, target = filename)
  
  pdf_filename <- paste0("output/pdf/Diploma_", participant$Name, ".pdf")
  convert_to_pdf(filename,pdf_filename)
}
```
