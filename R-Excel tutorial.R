#Welcome to How to Incorporate (or Replace) Your Excel Tasks with R
#Author: Mitchell D. Shuey
#------
install.packages("tidyverse")
install.packages("xlsx")
#The following function returns where your directory is. setwd("path") will naturally set it.
getwd()
#PHASE ONE: Importing data. Easiest way is to have the file, with a short name, already in your working directory.
#shift+right click your file in Windows Explorer, then choose "Copy as path". 
#Replace backslashes with forward ones via Ctrl-R
#Don't forget to select sheet and skip!
library(readxl)
WinePrd = read_excel("C:/Users/LENOVO/Documents/GitHub/Tutorials/Megafile_of_global_wine_data_1835_to_2016_1217.xlsx", 
                      sheet = "T6 Wine production", skip=1)
WineCon = read_excel("C:/Users/LENOVO/Documents/GitHub/Tutorials/Megafile_of_global_wine_data_1835_to_2016_1217.xlsx", 
                    sheet = "T34 Wine consumption vol", skip = 1)
View(WinePrd)
colnames(WinePrd)[1]="Year"
View(WineCon)
#Scroll to the rightmost edge of the sheet. Notice Norway. 
#This will be a recurring piece of the data to account for as we transform. 
#It's a great example of quirks in datasets that are best accounted for early.
#For now, we rearrange:
ncol(WineCon)
WineCon=WineCon[,c(1,ncol(WineCon), 2:57)]
colnames(WineCon)[1]="Year"
#------
#PHASE TWO: Slicing data into cross-sections
library(tidyr)
#converting from wide to tall
tidy.prd = gather(WinePrd, key=Country, value=Vol, -Year, -X__2, -ncol(WinePrd))
#removing extra column
tidy.prd = tidy.prd[,-2]

#repeating for wine consumption.
tidy.con = gather(WineCon, key=Country, value= Vol, -Year, -X__2, -`Coeff. of variation`)
tidy.con = tidy.con[,-2]

library(dplyr)
tidy.wine=merge(tidy.prd, tidy.con, by=c("Year", "Country"), suffixes=c(".prd",".con"), all=T)
View(tidy.wine)
#Hold on, we're missing something! (all=T)

#Now we want better factors to organize our countries with.
tidy.wine[,7] = tidy.wine$Region
tidy.wine$Region[tidy.wine$Country %in% colnames(WineCon[,c(2:27)])] = "Europe"
tidy.wine$Region[tidy.wine$Country %in% colnames(WineCon[,c(28:29)])] = "Oceania"
tidy.wine$Region[tidy.wine$Country %in% colnames(WineCon[,c(30,31)])] = "NorthAm"
tidy.wine$Region[tidy.wine$Country %in% colnames(WineCon[,c(32:37)])] = "LatAm"
tidy.wine$Region[tidy.wine$Country %in% colnames(WineCon[,c(38:43)])] = "AME"
tidy.wine$Region[tidy.wine$Country %in% colnames(WineCon[,c(44:54)])] = "AsiaPac"
tidy.wine$Region[tidy.wine$Country %in% colnames(WineCon[,55])] ="Other"
tidy.wine$Region[tidy.wine$Country %in% colnames(WineCon[,56])] = "WORLD"
tidy.wine$Region=as.factor(tidy.wine$Region)
#Did we factor everything as we wanted to?
summary(tidy.wine$Region)
#We're ready to move on!
#----
#PHASE THREE: Graphs and subset analysis
library(ggplot2)
tidy.wine = tidy.wine %>% mutate(`net`=`Vol.prd`-`Vol.con`)
Region.Prd.Year=spread(tidy.wine, Region, -2)
summarise(group_by(tidy.wine, Region), sum.prd=sum(Vol.prd))
r.y=ggplot(tidy.wine[tidy.wine$Country=="Portugal",], aes(Year, `Vol.prd`))
r.y+geom_line()
