### Bulk certificates from pptx in R ###
# Creation of bulk certificates/diplomas with pptx template in R

#First, install "officer" and "docxtractr" packages

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
