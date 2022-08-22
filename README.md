# Excel-VBA-Dashboard
This is a dynamic dashboard that analyzes the performance of the stock market in many countries of the world , since the begining of the year 2022 , comparing it to the year 2021 .  

There's some work-around that's behind the first worksheet ( The dataset is not collected directly like that ) . The process of preparing data and merging it the first time or 
each time I want to update the dataset is tidious ( I have 11  indices thus 11 separate worksheets ). So I created some <b>queries</b> to automate the process: The only manual job will be downloading the price series of each 
index . Once downloaded , here's how the query works : there are many indices that have the same date index series( MSCI indices) . These are automatically merged together
. However , those familiar with power query know that many queries cannot be merged all at a time ( you should merge pair by pair ) . The idea is therefore , to <b>append</b> 
these indices all at a time ( this is indeed possible with power query ) then <b>pivot them on the source column</b> , to obtain a merge instead of an append --> Problem solved! 
However , this only works for series that have a common column in all values ( case of the MSCI indices) .  For the rest of the indices ( which aren't many ) , the merge 
is done pair by pair . However , this is going to be done once and for all . The next time I update the dataset , as long as I'm working with queries , I can <b>just change
the source</b> and the steps of the query will automatically apply . 

Another thing to note is that the dashboard updates automatically when I update the dataset ( by clicking on the Activate button ) ,  hence the 'dynamic' character . The problem that I encountered with this is
that you <b>CANNOT make a table expand when another table does</b> , without VBA code or power query . Using queries here does not solve the problem because the tables I'm working
on have formulas that depend on formulas in the other tables . So I added VBA code : Note that this code consists of only adding dates to each table and the other columns ( of each table ) 
will automatically update , <b>without any code required</b>( because we're working with TABLES ) . 

One final thing is that when you download the xlsm file and open it , you'll get a message of this kind : '
security risk microsoft has blocked macros from running because the source of this file is untrusted' ( this message pops up whenever you open an external file that contains VBA code )  , and the functionalities of the dashboard will obviously not work . To fix this , close the file --> right click on the file --> select properties --> tick unblocked ( option that shows at the bottom of the box ) , then reopen the file and click enable editing . 
Below is just a screenshot of the dashhboard worksheet : 
![Dashboard screenshot](https://user-images.githubusercontent.com/69468586/185928063-ddff1abd-15fe-495a-9bd3-d69108d3c354.PNG)

