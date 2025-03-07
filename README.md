# Project Introduction
Hey, everybody! It's me back again doing a documentation of my project to showcase my skills. This project will mainly utilize **Excel** to reach our objective of making a spreadsheet that simulates the emergency response times of a certain  Disaster and Risk Management Office here in my locality. This was actually a key component for my graduate studies, and since my team just recently went to Singapore to present our study, it made me revisit a couple of things, hence me recreating this project. Without wasting any more time, let's get right into it. 
# Project Overview
As the title suggest, I will be creating an emergency response simulator using Excel. But the simulation technique that we will be using in this project is a classic simulation technique called ***Monte Carlo Simulation Technique***, a renowned technique in my major due to its simplicity and plain effectivity in forecasting results. The project will mainly focus on simulating the long term average response time of the DRMO (short for the Disaster and Risk Management Office) I mentioned earlier to get actual numbers on how long they would actually perform in the long term based on the historical data we have gathered; if it is still within the accepted range of 8 minutes or so. 
# Project Creation
## Headers
Headers are important in spreadsheets. Which is why it is important that we foresee what headers we would actually need to meet our objective. Good for you, I already did the hard part of going back and forth in my trial and error sessions upon creating the project. So, the headers that we will be having are these:

![image](https://github.com/user-attachments/assets/47bc30dc-8519-4e9e-8490-ccbdb5a2e28a)

Let me explain briefly what each header is for.

***Day*** - This will determine what day we are at.

***Date Value*** - This is where I assigned actual date values for the ***Day*** column.

***Number of Incidents*** - This will determine the number of incidents that will happen in a day.

***Random Generated Incident Time*** - This will be the random generated time of when an incident will occur.

***Sorted Time*** - This is basically the sorted values of the previous column.

***Date-Time of Incident*** - This is where I merge the date values to the random generated incident times values.

***Random Number*** - Just a column of randonm numbers that will be used for the next column.

***Incident Location*** - This is where the incident will happen, determined using the random number generated in the previous column.

***Status*** - This will be the indicator whether the previous response has elapsed 30 minutes, solely be for the purpose of checking whether an alternative post is needed.

***Assigned Post*** - This will be the assigned post based on the data.

***RT (Response Time)*** - The response time of the assigned post.

***Alternative Post*** - This will be for the sake if the main assigned post is still busy with an ongoing case.

***RT*** - The response time of the alternative post.

## VLOOKUP and Reference Tables
### VLOOKUP Tables
This will be where we will be converting our historical data into tables where the VLOOKUP function can perform well. Below is a picture of the tables:

![image](https://github.com/user-attachments/assets/44a1706f-d829-4320-93a6-b5cae4dc5eb3)
*Various VLOOKUP tables that will be used for the project*

As you can see, the tables are made in such a way where each random number, post, or incident locations has their own corresponding values, all of which are perfect to be used in conjunction with the VLOOKUP functionality.

To explain better, refer to the first table of ***Incident Per Day Probability Table***. On the left is the random number, and on the right is its correspnding number of incidents that will occur in a day. Say that the random number generated is 1, thus, based on the table, it would equate to having no incident for that day.

Again, these are all based on historical data. If according to the historical data that most likely number of incidents per day is 2, then that would mean more random numbers will be assigned to 2 incidents per day. But I did not just guessed and feel how much random numbers each possibility gets, I actually calculated the probability of all of those possibilities and made a range for each one of those based on their cummulative probabilities. This is an oversimplicafication of the explanation of how I came up with this but what I want to tell you is that these are all accurate and faithful to the historical data that I used. It's just that I am more focused with showing what formulas I made to get to the objective rather than the explanation of the simulation technique itself.

### Reference Table
This table will be for referencing how many incidents will happen for all the days in the simulatiom. Below is a picture of the table:

![image](https://github.com/user-attachments/assets/b6e8cee6-9947-4558-9e97-73109f64b60d)  
*The Reference Table that shows how many incidents will occur for each day based on random generated numbers*

In such a short while, I already used the first table in our ***VLOOKUP Table Worksheet***. I generated random numbers on Column C using the `RAND` function, and then used VLOOKUP by matching those random generated numbers to the number of incidents that will occur for all the days based on the ***VLOOKUP Table Worksheet***. ALso, since this is a simulation, I did a total of 10000 runs or 10000 days to account for the variability of the data.

This is the formula that I made in Column B:
```
=VLOOKUP(C2,Table7,2,FALSE)
```
A simple formula that utilized the VLOOKUP function.

## Simulation Table
Now, this is where simulation itself will take place. The headers that I discussed earlier is also located here. Since the headers are basically the steps for the simulation, I will be dividing each of them and explain them one by one. Let's get started.
### Day
As said earlier, this is just a numerical representation of what day we are at. It would just be simple to drag down until day 10000 but it's not that easy as we need to free up some rows based for the number of incidents that will occur for that day. It would be a tedious task to manually repeat so I decided to automate this step. A simple script will not do so we need a create actually code using VBA for this automation to consider all scenarios.

Below is the VBA code for the script:

```VBA
Sub Step1_InsertValuesAndShift()
    Dim ws As Worksheet
    Dim refWs As Worksheet
    Dim refTable As Range
    Dim startCell As Range
    Dim rowCount As Long
    Dim i As Long
    Dim value1 As Double
    Dim value2 As Long
    
    Set ws = ThisWorkbook.Sheets("Simulation")
    Set refWs = ThisWorkbook.Sheets("Reference Table")
    Set refTable = refWs.Range("ReferenceTable")
    Set startCell = ws.Range("A2")
    
    rowCount = refTable.Rows.Count
    For i = 1 To rowCount
        value1 = refTable.Cells(i, 1).Value
        value2 = refTable.Cells(i, 2).Value
        
        startCell.Value = value1
        
        If value2 = 0 Then
            Set startCell = startCell.Offset(1, 0)
        Else
            Set startCell = startCell.Offset(value2, 0)
        End If
    Next i
End Sub
```
To explain simply, this code will be refering to the reference table and do these sequence of steps:
1. Put the day number.
2. Shift down based on how many incidents will occur in that day.
3. Repeat.

### Date Value
This column is where we will assign actual dates for the previous column. Again, it would be a tedious task to manually so we created a VBA code to automate this step.

Below is the VBA code for the script:

```VBA
Sub Step2_InsertValuesAndShiftForDateValueColumn()
    Dim ws As Worksheet
    Dim refWs As Worksheet
    Dim refTable As Range
    Dim startCell As Range
    Dim rowCount As Long
    Dim i As Long
    Dim value1 As Double
    Dim value2 As Long
    
    Set ws = ThisWorkbook.Sheets("Simulation")
    Set refWs = ThisWorkbook.Sheets("Reference Table")
    Set refTable = refWs.Range("ReferenceTable")
    Set startCell = ws.Range("B2")
    
    rowCount = refTable.Rows.Count
    For i = 1 To rowCount
        value1 = refTable.Cells(i, 1).Value
        value2 = refTable.Cells(i, 2).Value
        
        startCell.Value = value1
        
        If value2 = 0 Then
            Set startCell = startCell.Offset(1, 0)
        Else
            Set startCell = startCell.Offset(value2, 0)
        End If
    Next i
End Sub

```
This code will do the following sequence of steps:
1. Check if there is a value in the left adjacent cell.
2. If there is, put a date in there. (Starts at 1/1/1900)
3. If there is nothing, shift down.
4. If there is a value again on the left adjacent cell, put a date again but +1 more day than the previous date used.
5. Repeat.

This code however does not yet suffice for this column as there are still blank cells that must be filled. Thus, we need a script that fills in the blank values.

Below is the VBA code for the script:
```VBA
Sub Step3_FillEmptyCells()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Simulation")
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Set rng = ws.Range("B2:B" & lastRow)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = "" Then
            ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```

This script will do the following sequence of steps:
1. Check the first cell and copy its value.
2. Paste the value if the cell, when shifted down, is empty.
3. If stumbled upon a cell that is not empty, copy its value.
4. Repeat.

### Number of Incidents
This column would just be the same as the first column. We just need to show how many incidents will happen that day based on the reference table. Thus, to automate this step, we will just be copying the code of our first script and do minor edits as needed.

Below is the VBA code for the script:
```VBA
Sub Step4_InsertValuesAndShiftForNoOfIncidents()
    Dim ws As Worksheet
    Dim refWs As Worksheet
    Dim refTable As Range
    Dim startCell As Range
    Dim rowCount As Long
    Dim i As Long
    Dim value1 As Double
    Dim value2 As Long

    Set ws = ThisWorkbook.Sheets("Simulation")
    Set refWs = ThisWorkbook.Sheets("Reference Table")
    Set refTable = refWs.Range("ReferenceTable")
    Set startCell = ws.Range("C2")
    
    rowCount = refTable.Rows.Count
    For i = 1 To rowCount
        value1 = refTable.Cells(i, 2).Value
        value2 = refTable.Cells(i, 2).Value
        
        startCell.Value = value1
        
        If value2 = 0 Then
            Set startCell = startCell.Offset(1, 0)
        Else
            Set startCell = startCell.Offset(value2, 0)
        End If
    Next i
End Sub
```

### Random Generated Incident Time
This column will determine the time of when the incident will occur. Good for us, we have Excel's built in functioanlity of generating random time values. It was simply a matter of using the `RAND` function and changing the number format to ***Time***.

Formula used for Column D:
```
=RAND()
```
We then copy the whole column and paste as values to prevent the recalculation of the `RAND` function. 

### Sorted Time
This column is a little bit tricky as we need to sort the time for the whole column, but also taking into account that those time values did not occur in just 1 day, they have their own corresponding days. Therefore, we needed to create yet another script to automate this task.

Below is the VBA code for the script:
```VBA
Sub Step5_AutoSortIncidentTime()
    Dim ws As Worksheet
    Dim refWs As Worksheet
    Dim incidentCount As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim refRow As Long
    
    Set ws = ThisWorkbook.Sheets("Simulation")
    Set refWs = ThisWorkbook.Sheets("Reference Table")

    startRow = 2

    For refRow = 2 To refWs.Cells(refWs.Rows.Count, "D").End(xlUp).Row
        incidentCount = refWs.Cells(refRow, 2).Value
        
        If incidentCount > 0 Then
            endRow = startRow + incidentCount - 1
            ws.Cells(startRow, 4).Formula = "=SORT(E" & startRow & ":E" & endRow & ")"
            startRow = endRow + 1
        Else
            ws.Cells(startRow, 4).Value = "-"
            startRow = startRow + 1
        End If
    Next refRow
End Sub
```

This script will do the following sequence of steps:
1. Check the number of incidents for that day.
2. Insert in the blanks in the formula `=SORT(D_:D_)` based on the number of incidents for that day. Example output for a day with 2 incidents: `@=SORT(D2:D3)`
3. Shift down based on how many incidents occured.
4. If no incident will occur that day, return `-`.
5. Repeat.

Since the script is returning formulas that have `@` at the beginning, we needed a script that removes all of those for the whole column to properly execute the spill formula. 

Below is the VBA code for the script:
```VBA
Sub Step5_FixFormulasInColumnE()
    Dim ws As Worksheet
    Dim cell As Range
    Dim lastRow As Long
    Dim formulaText As String

    Set ws = ThisWorkbook.Sheets("Simulation")

    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    For Each cell In ws.Range("E2:E" & lastRow)
        If cell.HasFormula Then
            formulaText = cell.Formula

            If Left(formulaText, 2) = "=@" Then
                formulaText = "=" & Mid(formulaText, 3)
            End If

            cell.Formula2 = formulaText
        End If
    Next cell

End Sub
```
This just checks whether the formula has `@` at the beginning and rewrites the formula without it.

### Date-Time of Incident
This column will just merge the date values and the sorted time values. This is solely for the purpose the ***Status*** column because there will be situations where the difference in time of 12:00 AM and 1:00 AM is another thing compared to the difference in time of 1/1/1900 12:00 AM and 1/2/1900 1:00 AM. Also, if the left adjacent cell has `-` in it, we would also return the same for the cells in this column for easier understanding.

Formula used for Column F:
```
=IF(E2<>"-",B2+E2,"-")
```

### Random Number
This column will just basically generate random numbers if there is a value in the left adjacent cell. And again, if the left adjacent cell has a value of `-`, return the same.

Formula used for Column G:
```
=IF(F2<>"-",RANDBETWEEN(1,10000),"-")
```

### Incident Location
This is where we determine where the incident location will occur. This is simply a matter of using the VLOOKUP function in conjunction with the VLOOKUP Tables that we made earlier to return the incident location.

Formula used for Column H:
```
=IF(COUNT(G2)=1,VLOOKUP(G2,Table4,2,FALSE),"-")
```

### Status
Starting from here are the tricky parts of the whole project. It is possible that the main responders assigned to respond to a certain location is still busy with an ongoing case, therefore, we need to assign the alternative responder for that place given the situation. I assumed that a response will take atmost 30 minutes to fully finish (going there, tending the needs of the victims, and going back the headquarters). We need to do this in such a way where we tell that the next case will occur in less than 30 minutes, so we should be checking if there will be a need to deploy the alternative responder or not.

So, to tell whether the previous case has not exceed 30 minutes as this case occurs, we use the following formula for Column I:
```
=ABS(F2-F1)<TIME(,30,)
```

It was simply a matter of substracting two time values and getting its absolute values. There will be three possible outcomes, `TRUE`,`FALSE`, and `#VALUE!`. To put it simply, A returned value of `TRUE` says that this case will happen before 30 minutes have passed since the last, and values of `FALSE` and `#VALUE!` says otherwise. Therefore, we need to be wary in those instances where the `TRUE` value is returned for the next columns.

### Assigned Post
As we start assigning posts to the incidents, we must first check the value in the ***Status*** column and if whether a responder will be assigned again within a span of 30 minutes. We need a formula that checks the following:
1. If the ***Status*** says `FALSE` or `#VALUE!`.
2. If the ***Status*** says `TRUE`.
3. If the responder that will respond, when the returned value in the ***Status*** column is `TRUE`, is the same as the previous responder.
4. If the responder that will respond, when the returned value in the ***Status*** column is `TRUE`, is not the same as the previous responder.

Taking those into account, we then need the formula to return the following in these given scenarios:
1. If 1. is `TRUE`, return the assigned responder.
2. If 3. is `TRUE`, return `-` and assign an alternative responder in the ***Alternative Post*** column.
3. If 4. is `TRUE`, return the assigned responder.

By incorporating all those scenarios, we are left with a formula that has nested `IF`s within an `IFS` functionality. Below is the formula used for Column J:
```
=IFS(ISERROR(I2),IF(H2<>"-",VLOOKUP(H2,Table2,2,FALSE),"-"),I2=TRUE,IF(VLOOKUP(H2,Table2,2,FALSE)=J1,"-",VLOOKUP(H2,Table2,2,FALSE)),I2=FALSE,VLOOKUP(H2,Table2,2,FALSE))
```

### RT (Response Time)
This will be a short break after what just happened in the previous column. This column will just utilize a `VLOOKUP` function within an `IF` function to display the response time of the responders going to the incident location, and returning `-` if the left adjacent cell has that value as well. 

Formula used for Column K:
```
=IF(J2<>"-",VLOOKUP(H2,Table5,2,FALSE),"-")
```

### Alternative Post
This column just basically picks up where we left off in the ***Assigned Post***column. Before assigning alternative responders, we must first determine when should we assign one based on the structure of our worksheet.

To determine whether we must assign an alternative responder, the following conditions must be met:  
1. The value in the ***Status*** column is `TRUE`.
2. The value in the ***Assigned Post*** column is `-`.
3. The supposed responder of this case is the same as the previous case.

Once all of those are met, we should be returning the alternative responder, else, return `-`.

Thus, below is the formula used for Column L:
```
=IFS(ISERROR(I2),"-",AND(I2=TRUE,J2="-"),VLOOKUP(H2,Table3,2,FALSE),AND(I2=TRUE,NOT(J2="-")),"-",I2=FALSE,"-")
```

### RT
This column is basically the same for the other ***RT*** column, but this time, we use the `VLOOKUP` Table for the response time of the alternative responders. So it again a matter of using the `IF` and `VLOOKUP` function, and also returning `-` if the left adjacent cell has that value as well.

Formula used for Column M:
```
=IF(L2<>"-",VLOOKUP(H2,Table6,2,FALSE),"-")
```

## Simulation Summary Table
This is the table where we get the final results of our simulation and get valuable information and insights. Below is a picture of that table:

![image](https://github.com/user-attachments/assets/ee5b42fb-4df8-49d2-8578-350b0e41676f)  
*The Simulation Summary Table showing the results of the simulation*

Making this table is easy. I just used functions like `CONCAT`,`AVERAGE`,`COUNTIFS`,`MAX`, and `SUM`, and doing conditional formattings here and there for better emphasis. I will still show below the list of formulas that I used for this table.

```
=CONCAT("The simulation did a total of ",MAX(A:A)," runs.")
=CONCAT("The simulation had a total of ",SUM(C:C)," incidents.")
=CONCAT(COUNTIFS(J:J, "<>-", J:J, "<>")-1, " cases used their main post for respose.")
=CONCAT(COUNTIFS(L:L,"<>-",L:L,"<>")-1," cases used their alternative post for response.")
=CONCAT("The average response time is ",ROUND(AVERAGE(K:K,M:M),2)," minutes.")
=COUNTIF(J:J,O11)+COUNTIF(L:L,O11)
=P11/MAX(B:B)
=P11/SUM(C:C)
```

# Ending Remarks
It's been a while since I last did a project here. I hope some people get to see this as this is one of the works I am most proud of as of this day despite still needing to polish my skills in writing formulas. I would also like to thank my research mate as she was actually the one who made the starting concept of the simulator. I just happened to continue this work of hers. Hope you all realized a thing or two in the potential of Excel to create this type of project. Wishing you all had fun reading this as I had fun making it. Thanks!
