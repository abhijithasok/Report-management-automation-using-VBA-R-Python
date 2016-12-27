# Report-management-automation-using-VBA-R-Python
This structure integrates operations on VBA, R and Python to fully automate a daily data report creation and transfer solution.

Scenario:
There is a daily receipt as an email, of a dataset that contains the data of all organizational purchases(real estate) over the past 1 year, starting from the previous day. What is of interest to be gained out of this data are those items that were purchased on the previous day whose unit price in that particular purchase is:
1. More than
2. Less than, the average unit price for the same item over the past 1 year. 

These reports would later be used in Time Series implementations and predictions for optimizing purchases over time.

The workbook that contains all such purchases, separated into different tabs(Increase/Decrease) as well as individual plotted charts of unit prices of all these items, over all purchases made over the previous year and the original data are to be sent to a set of people via email every day.

The data-based challenges in creating the workbook can be roughly outlined as follows:
1. The same item could have been purchased in different units(eg: kg, quintal) in different purchases over the past year. The unit prices would be altered accordingly in the incoming data, thereby making a direct filter and plot error-prone.
2. Some items could be purchased as different forms of quantity(eg: cement, purchased by weight and purchased by volume). Even though they may have different units in the incoming data, the purchases of different forms of quantity are to be treated differently.

The workflow for this automation is as follows:
----------------- Daily loop at around a specific time ----------------------------------------------------
1. Retrieve the xls file from the corresponding e-mail
2. Convert the xls file to csv (optional step, but it eases the compatibility with R)
3. Perform operations and generate the output workbook and figures in the required format
4. Compress the figures and the original data(so that they fit well into an e-mail
5. Send the e-mail
-----------------------------------------------------------------------------------------------------------

The tool usage is as follows:
1. VB code on Microsoft Outlook that identifies an e-mail with a specific header(that also contains the date of that particular day) and downloads the attachment along with it into a specified folder.
2. VB code on Microsoft Excel that picks up this downloaded xls file and converts it into a csv file of the same name and saves it in the same folder
3. R script that accepts the following inputs:
   a. This converted csv file of raw data
   b. A list of unit combinations between past and previous day and their appropriate dividing factor to convert past purchases to the
      previous day's purchase unit 
   c. A list of unit combinations between past and previous day purchases that have to be treated differently, due to differences in 
      forms of quantity
      
   The R script generates a workbook of 2 tabs - one containing the list of items purchased on the previous day with an average past  
   price higher then the past 1 year average(now, error-free) and the second, the list of items in the same context whose price is lower    than the past one year average.
   
   The R script also generates time-based line plots for each of the items in these tabs of the workbook.
4. Python script that takes in the original raw data downloaded from the e-mail by the VB Script, as well as the folders that contain
   the figures generated through the R script and compresses both, saving them into the same folders.
5. R script that creates an e-mail Outlook object and sends an e-mail to specified recipients with the compressed original data,  
   compressed figures(plots) as well as the Increase/Decrease workbook as attachments.
