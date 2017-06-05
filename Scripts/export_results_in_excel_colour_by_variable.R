library(openxlsx)
library(reshape2)
load("/media/development/DataCATz/Datasets/blog_RM2_data.RData")
head(limma_results_multiple_comparisons)

# Identifing probes which are significant in at least one comparison (adjusted p-value<0.05)
significant<-dcast(subset(limma_results_multiple_comparisons,adj.P.Val<0.001),probe~.,length,value.var="comparison")
colnames(significant)<-c("probe","num_significant_comparisons")
head(significant)
significant_limma_results_multiple_comparisons<-subset(limma_results_multiple_comparisons,probe %in% significant$probe)

# Generating data frame containing p-values (rows correspond to probes, columns correspond to comparisons)
pvals<-dcast(data=significant_limma_results_multiple_comparisons,formula=probe~comparison,value.var="adj.P.Val")
rownames(pvals)<-pvals$probe
pvals$probe<-NULL
head(pvals)

# Generating data frame containing fold changes (rows correspond to probes, columns correspond to comparisons in the same order as above)
fold_changes<-dcast(data=significant_limma_results_multiple_comparisons,formula=probe~comparison,value.var="logFC")
rownames(fold_changes)<-fold_changes$probe
fold_changes$probe<-NULL
head(fold_changes)

xlsx_filename<-"/media/development/DataCATz/Tables/results.xlsx"

# Creating workbook
wb<-openxlsx::createWorkbook("Results")

# Creating and filling the p-values worksheet
openxlsx::addWorksheet(wb,"pvalues",gridLines=TRUE)
writeData(wb,sheet=1,pvals,rowNames=TRUE)

# Creating and filling the fold changes worksheet
openxlsx::addWorksheet(wb,"fold_changes",gridLines=TRUE)
writeData(wb,sheet=2,fold_changes,rowNames=TRUE)
openxlsx::saveWorkbook(wb,xlsx_filename,overwrite=TRUE)

# Generating gradient of colours
logfc_intervals<-c(-Inf,seq(-8,8,1),Inf)
blues<-colorRampPalette(c("blue", "white"))
reds<-colorRampPalette(c("white","red"))
list_of_colours<-blues(floor((length(logfc_intervals)-1)/2))
list_of_colours<-c(list_of_colours,reds(floor((length(logfc_intervals)-1)/2)))

# Applying colours
for(i in seq(1,length(logfc_intervals)-1)){
  cells_to_colour<-which(pvals<0.05 & fold_changes>logfc_intervals[i] & fold_changes<logfc_intervals[i+1],arr.ind=TRUE)
  style<-createStyle(fgFill=list_of_colours[i])
  openxlsx::addStyle(wb,sheet=1,rows=cells_to_colour[,1]+1,col=cells_to_colour[,2]+1,style=style)
}

# Saving the workbook in an Excel file
openxlsx::saveWorkbook(wb,xlsx_filename,overwrite=TRUE)

# Printing package version
sessionInfo()

