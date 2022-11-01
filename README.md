# Homework_Repo_Week2_VBA

The code runs through each line of ticker data. It pulls out each individual ticker symbol and creates a separate list. 
The opening price for the year is identified and subtracted from the closing price to give the yearly change, which is placed in the corresponding column for the ticker. 
Conditional formatting is applied: green for stocks which have increased and red for those which have decreased. 

The yearly change is divided by the opening price for the year to give the percentage change over the year.

The volume of stocks which are sold on each date is added up and placed into the row matching the ticker symbol. 

The code then runs through the calculated data and pulls out the information for the Greatest % increase, the greatest % decrease and the greatest total volume. 
