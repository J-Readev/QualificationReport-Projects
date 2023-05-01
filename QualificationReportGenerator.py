#Hello. This is a python file in order to sort out the qualification report that iTrent generates monthly. 

#Pandas is needed for this.
import pandas as pd
from datetime import datetime


#Accessing Spreadsheet.Specifically require the Qualifications page, but only needs select columns to determine who requires a new qualification.
#Read_Excel will need to locate the qualification report in whatever directory it is saved in.
Whole_Spreadsheet = pd.read_excel(r [LOCATION OF FILE], 
                    sheet_name = "Qualifications",
                    usecols = ["Full Name","Position", "Subject","Qualification Valid Until","Status"])

#Dropping data that isn't needed.
Whole_Spreadsheet.dropna(subset=["Subject"], inplace= True)
Whole_Spreadsheet.dropna(subset=["Qualification Valid Until"], inplace= True)

#Ensuring all data is displayed in csv file.
pd.set_option("display.max_rows",None)

#Generating CSV with data needed.
Needed_Data = Whole_Spreadsheet.to_csv("Qualification Report Needed Data.csv",header = True, index = False)

#Creating title for file to correspond to current month & year using datetime.
current_Month = datetime.now().month
current_Year = datetime.now().year

#If statement checking current months numerical output and changing it to text before adding to file name.
if current_Month == 1:
    current_Month = "January"
elif current_Month == 2:
    current_Month = "February"
elif current_Month == 3:
    current_Month = "March"
elif current_Month == 4:
    current_Month = "April"
elif current_Month == 5:
    current_Month = "May"
elif current_Month == 6:
    current_Month = "June"
elif current_Month == 7:
    current_Month = "July"
elif current_Month == 8:
    current_Month = "August"
elif current_Month == 9:
    current_Month = "September"
elif current_Month == 10:
    current_Month = "October"
elif current_Month == 11:
    current_Month = "November"
elif current_Month == 12:
    current_Month = "December"
    
#Connecting everything together.
file_Title = str(current_Month) + " " + str(current_Year) + " " +"Qualification Report Needed Data.csv"

#Generating CSV with data needed with updated title.
Needed_Data = Whole_Spreadsheet.to_csv(file_Title,header = True, index = False)
