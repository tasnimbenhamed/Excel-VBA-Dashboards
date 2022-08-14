# Excel-VBA-Dashboard
Dynamic and interactive dashboard based on financial data 

There's some work-around that's behind the first worksheet ( The dataset is not collected directly like that ) . The process of preparing data and merging it the first time or 
each time I want to update the dataset is tidious ( I have 11  indices thus 11 separate worksheets ). So I created some queries to automate the process: The only manual job will be downloading the price series of each 
index . Once downloaded , here's how the query works : there are many indices that have the same date index series( MSCI indices) . These are automatically merged together
. However , those familiar with power query know that many queries cannot be merged all at a time ( you should merge pair by pair ) . The idea is therefore , to append 
these indices all at a time ( this is indeed possible with power query ) then pivot them on the source column , to obtain a merge instead of an append --> Problem solved! 
However , this only works for series that have a common column in all values ( case of the MSCI indices) .  For the rest of the indices ( which aren't many ) , the merge 
is done pair by pair . However , this is going to be done once and for all . The next time I update the dataset , as long as I'm working with queries , I can just change
the source and the steps of the query will automatically apply . 

Another thing to note is that the dashboard updates automatically when I update the dataset , hence the 'dynamic' character . The problem that I encountered with this is
that you CANNOT make a table expand when another table does , without VBA code or power query . Using queries here does not solve the problem because the tables I'm working
on have formulas that depend on formulas in the other tables . So I added VBA code : Note that this code consists of only adding dates to each table and the other columns ( of each table ) 
will automatically update , without any code required ( because we're working with TABLES ) . 
