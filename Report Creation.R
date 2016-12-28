# Many of the packages below meant for enabling R to handle xls and xlsx serve the same-
# -purpose. All are loaded just for options

library(colorspace)
library(ggplot2)
library(ggrepel)
library(devtools)
library(readxl)
library(XLConnect)
library(installr)
library(rJava)
library(xlsx)
library(gdata)
library(gtools)
library(dplyr)
library(openxlsx)
library(rtools)

############################ Raw Data Input #####################################

rawname <- paste0("C:/Users/abhijithasok/Documents/Purchase Daily Dashboards/Raw Data/OMX_PO_ITEM_RATE_ALT_",toupper(format(Sys.Date(),"%d-%b-%Y")),".csv") #The date format flexibility in the name ensures that the particular day's data is correctly picked up

rawdata <- read.csv(file = rawname,header=T,stringsAsFactors = F)

############################# Creating Inc/Dec ##################################

pastdata <- rawdata[rawdata$PO.date != format(Sys.Date()-1,"%d-%b-%y"),] #extracting all purchases from 2 days before and back, from the present date
presentdata <- rawdata[rawdata$PO.date == format(Sys.Date()-1,"%d-%b-%y"),] #extracting all purchases from 1 day before, from the present date

#From this point, 'past' refers to an element(s) in 'pastdata' and 'present' refers to an element(s) in 'presentdata'

uomcon <- read.csv("C:/Users/abhijithasok/Documents/Purchase Daily Dashboards/UOM Conversion list.csv",header=T) #List of unit combos between past & present and their conversion factors
uomuncon <- read.csv("C:/Users/abhijithasok/Documents/Purchase Daily Dashboards/UOM Unconversion.csv",header=T) #List of unit combos that are to be treated as different items, in spite of other parameters being same

### Major operations on pastdata ###

pastdata$pastID <- ""
pastdata$pastID <- 1 : nrow(pastdata) #Unique ID for past instances

presentdata$presentID <- ""
presentdata$presentID <- 1 : nrow(presentdata) #Unique ID for present instances

m <- merge(pastdata,presentdata[,c("Item.desc","Extra.Desc","UOM")],by = c("Item.desc","Extra.Desc"),all.x=T) #Matching every purchase entry in the past to the unit of purchase in the present(if it exists), based on descriptions of the item of purchase
m <- m[!duplicated(m),]

colnames(m)[which(names(m) == "UOM.x")] <- "UOM"
colnames(m)[which(names(m) == "UOM.y")] <- "Present.UOM"

m$pastID<-NULL
m$Different.UOM<-""

m$Different.UOM <- ifelse(m$UOM == m$Present.UOM,0,1) #Checking the rows where past and present units are differing

m$Unconversion.Flag <- ""

uomuncon$Unconversion.Flag <- c(1,1)

m <- merge(m,uomuncon,by.x=c("UOM","Present.UOM"),by.y=c("Unconversion.Past","Unconversion.Present"),all.x=TRUE) #Matching every purchase entry in the past with a flag that indicates whether the entry should be left as is, in spite of past and present units differing
m$Unconversion.Flag.x<-NULL
colnames(m)[which(names(m) == "Unconversion.Flag.y")] <- "Unconversion.Flag"

m$Unit.Price.in.Present.Unit<-""

m <- merge(m,uomcon,by.x=c("UOM","Present.UOM"),by.y=c("UOM.Original","UOM.Convert"),all.x=TRUE)

m$Unit.Price.in.Present.Unit <- ifelse(m$Different.UOM==1 & m$Unconversion.Flag==0,m$Unit.Price/m$Dividing.factor,m$Unit.Price) #Using dividing factor from conversion table to alter unit price if the units are different, and carrying the same value over otherwise

m$Converted.UOM <- ""

m$Converted.UOM <- ifelse(m$UOM == m$Present.UOM,m$UOM,ifelse(m$Unconversion.Flag == 0,m$Present.UOM,m$UOM)) #Final units after all conversions

### Major operations on presentdata ###

presentdata$presentID <-NULL
n <- presentdata  

aggmin <- aggregate(Unit.Price.in.Present.Unit~Item.desc+Extra.Desc+Converted.UOM,m,function(x)min(x)) #Computing minimum unit price by item descriptions from the past
aggmax <- aggregate(Unit.Price.in.Present.Unit~Item.desc+Extra.Desc+Converted.UOM,m,function(x)max(x)) #Computing maximum unit price by item descriptions from the past

minmerge <- merge(n,aggmin,by.x=c("Item.desc","Extra.Desc","UOM"),by.y=c("Item.desc","Extra.Desc","Converted.UOM"),all.x=T)
maxmerge <- merge(minmerge,aggmax,by.x=c("Item.desc","Extra.Desc","UOM"),by.y=c("Item.desc","Extra.Desc","Converted.UOM"),all.x=T) #Matching minimum and maximum prices between past and present across item descriptions and units

n <- maxmerge

colnames(n)[which(names(n) == "Unit.Price.in.Present.Unit.x")] <- "Past.Min.Price"
colnames(n)[which(names(n) == "Unit.Price.in.Present.Unit.y")] <- "Past.Max.Price"

n$Past.Avg.Price <- ""

n$Past.Avg.Price <- (n$Past.Min.Price + n$Past.Max.Price)/2

n$Inc.Dec.Prev.Avg <- ""

n$Inc.Dec.Prev.Avg <- ifelse(n$Past.Avg.Price != 0,ifelse(n$Unit.Price>n$Past.Avg.Price,"INCREASE",
                                                          ifelse(n$Unit.Price<n$Past.Avg.Price,"DECREASE",
                                                                 ifelse(n$Unit.Price == n$Past.Avg.Price,"NO CHANGE","NEW ITEM"))),"NEW ITEM") #Marking whether unit price increased, decreased or stayed constant compared to past average or whether the item is being purchased for the first time

n$Change.Amount <- ""

n$Change.Amount <- abs(n$Unit.Price - n$Past.Avg.Price) #Change amount

n$Change.Percentage <- ""

n$Change.Percentage <- ifelse(!is.na((n$Change.Amount/n$Past.Avg.Price)),paste0(round(((n$Change.Amount/n$Past.Avg.Price)*100),digits=2),"%"),"NA") #Change amount in %

n <- n[,c(4:12,1:3,13:27)] #Rearranging columns to match original data variable order

####### Creating Increase and Decrease lists from overall purchased items ########

inclist <- na.omit(n[n$Inc.Dec.Prev.Avg == "INCREASE",])
declist <- na.omit(n[n$Inc.Dec.Prev.Avg == "DECREASE",])

#### Creating workbook with increase and decrease lists (Flexible naming) #######

write.xlsx2(inclist,file = paste0("C:/Users/abhijithasok/Documents/Purchase Daily Dashboards/Inc_Dec/Increase_Decrease workbook - ",format(Sys.Date(),"%b %d, %Y"),".xlsx"),sheetName = "Increase in Price from Past Avg", append = FALSE, row.names = FALSE)
write.xlsx2(declist,file = paste0("C:/Users/abhijithasok/Documents/Purchase Daily Dashboards/Inc_Dec/Increase_Decrease workbook - ",format(Sys.Date(),"%b %d, %Y"),".xlsx"),sheetName = "Decrease in Price from Past Avg", append=TRUE, row.names = FALSE)

####################### Folder preparation for storing generated plots #######################

mainDir <- "C:/Users/abhijithasok/Documents/Purchase Daily Dashboards/Figures"
subDir <- paste0(format(Sys.Date(),"%b %d")," generated figures")

dir.create(file.path(mainDir, subDir))
setwd(file.path(mainDir, subDir))

subDir1 <- paste0("Increase - ",format(Sys.Date(),"%b %d"))
subDir2 <- paste0("Decrease - ",format(Sys.Date(),"%b %d"))

dir.create(file.path(mainDir, subDir, subDir1))
dir.create(file.path(mainDir, subDir, subDir2))

datainc <- inclist
datadec <- declist


################# Increase Figures ###################

for (i in 1:nrow(datainc))
{
  itemdesc<-datainc[i,10]
  extradesc<-datainc[i,11]
  
  tsdata <- rawdata[which(rawdata$Item.desc==itemdesc & rawdata$Extra.Desc==extradesc), ]
  tsdata$PO.date<-as.Date(tsdata$PO.date, "%d-%b-%y")
  tsdata$Item.desc<-gsub("/","-",tsdata$Item.desc)
  tsdata$Extra.Desc<-gsub("/","-",tsdata$Extra.Desc)
  name<-paste0("C:/Users/abhijithasok/Documents/Purchase Daily Dashboards/Figures/",substring(format(Sys.Date(),"%b %d, %Y"),1,6)," generated figures/Increase - ",substring(format(Sys.Date(),"%b %d, %Y"),1,6),"/",i," - ",gsub( " .*$", "", itemdesc),gsub( " .*$", "", extradesc),".jpg") #Plot save destination (flexible naming based on current date, item description, extra description)
  tryCatch({
    p <- ggplot(tsdata, aes(y=tsdata$Unit.Price, x=tsdata$PO.date, color=gsub( " .*$", "",tsdata$Company.Name)), type="n", xlab="Date", ylab="Unit Price") +
      geom_point() + geom_line() + geom_text_repel(aes(label=tsdata$Unit.Price), size=3) + ggtitle(paste(itemdesc,extradesc, sep=" ")) +
      labs(x="Date",y="Unit Price") + scale_colour_discrete(name = "Company Name")
    ggsave(filename = name, plot=p, width = 25, height = 10, units = "cm") #saving the generated plots
  }, error=function(e){cat("ERROR :",conditionMessage(e), "\n")})
}


################# Decrease Figures ###################

for (i in 1:nrow(datadec))
{
  itemdesc<-datadec[i,10]
  extradesc<-datadec[i,11]
  
  tsdata <- rawdata[which(rawdata$Item.desc==itemdesc & rawdata$Extra.Desc==extradesc), ]
  tsdata$PO.date<-as.Date(as.character(tsdata$PO.date), "%d-%b-%y")
  tsdata$Item.desc<-gsub("/","-",tsdata$Item.desc)
  tsdata$Extra.Desc<-gsub("/","-",tsdata$Extra.Desc)
  name<-paste0("C:/Users/abhijithasok/Documents/Purchase Daily Dashboards/Figures/",substring(format(Sys.Date(),"%b %d, %Y"),1,6)," generated figures/Decrease - ",substring(format(Sys.Date(),"%b %d, %Y"),1,6),"/",i," - ",gsub( " .*$", "", itemdesc),gsub( " .*$", "", extradesc),".jpg") #Plot save destination (flexible naming based on current date, item description, extra description)
  tryCatch({
    p <- ggplot(tsdata, aes(y=tsdata$Unit.Price, x=tsdata$PO.date, color=gsub( " .*$", "",tsdata$Company.Name)), type="n", xlab="Date", ylab="Unit Price") +
      geom_point() + geom_line() + geom_text_repel(aes(label=tsdata$Unit.Price), size=3) + ggtitle(paste(itemdesc,extradesc, sep=" ")) +
      labs(x="Date",y="Unit Price") + scale_colour_discrete(name = "Company Name")
    ggsave(filename = name, plot=p, width = 25, height = 10, units = "cm") #saving the generated plots
  }, error=function(e){cat("ERROR :",conditionMessage(e), "\n")})
}
