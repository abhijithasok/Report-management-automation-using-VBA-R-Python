library(RDCOMClient)

################# Outlook Mail ###################

OutApp <- COMCreate("Outlook.Application") #Creating an email object

outMail = OutApp$CreateItem(0) #Configuration of email parameters
 
outMail[["To"]] = "abc@xyz.com"
outMail[["CC"]] = paste("abc1@xyz.com","abc2@xyz.com","abc3@xyz.com","abc4@xyz.com",sep=";")
outMail[["subject"]] = paste0("Daily Purchase statistics - ",format(Sys.Date(),"%b %d, %Y"))
outMail[["body"]] = paste0("Hi,
                           
                           PFA the data, report and figures for ",format(Sys.Date(),"%B %d")," (",format(Sys.Date()-1,"%B %d")," purchases vs past).", 
                           
                           "Thanks!
                           ")
outMail[["Attachments"]]$Add(paste0("C:\\Users\\abhijithasok\\Documents\\Purchase Daily Dashboards\\Raw Data\\OMX_PO_ITEM_RATE_ALT_",toupper(format(Sys.Date(),"%d-%b-%Y")),".zip")) #Attaching compressed raw data
outMail[["Attachments"]]$Add(paste0("C:\\Users\\abhijithasok\\Documents\\Purchase Daily Dashboards\\Inc_Dec\\Increase_Decrease workbook - ",format(Sys.Date(),"%b %d, %Y"),".xlsx")) #Attaching created workbook
outMail[["Attachments"]]$Add(paste0("C:\\Users\\abhijithasok\\Documents\\Purchase Daily Dashboards\\Figures\\",substring(format(Sys.Date(),"%b %d, %Y"),1,6)," generated figures",".zip")) #Attaching compressed folders of generated figures
                    
outMail$Send() #Sending mail