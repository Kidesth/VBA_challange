VBA_ challenge stocks Analysis

Overview of the project:
The purpose of this project was to refactor code to analyze stokes data from any year within the given workbook. For the first time, I used nested statements to analyze the given data. The refactor code upgrades the efficiency by looping through the data set once

 Results:
My refactor code used a nested loop to find the stock return of change for a specified given year in my data
  I created any arrays of the ticker index volume 12 strings of the data. Then I used them as an index for my original loop while the inner loop went through the data looking for the information method the index thicker. 


Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearvalue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearvalue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
     tickerindex = 0
    
    '1b) Create three output arrays
    Dim tickervolumes(12) As Long
    Dim tickerstartingprice(12) As Single
    Dim tickerendingprice(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickervolumes(i) = 0
    Next i
        
    
     'output the data for the current ticker
      ' Worksheets("DQ Analysis").Activate
      '      Cells(4 + i, 1).Value = ticker
       '     Cells(4 + i, 2).Value = totalvolume
       '      cells(4 + i, 3).value = tickerendingprices
        
    
       
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickervolumes(tickerindex) = tickervolumes(tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
         If Cells(i - 1, 1).Value <> tickers(tickerindex) Then
          tickerstartingprice(tickerindex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
         
         If Cells(i + 1, 1).Value <> tickers(tickerindex) Then
          tickerendingprice(tickerindex) = Cells(i, 6).Value
          
          '3d Increase the tickerIndex
          tickerindex = tickerindex + 1
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Sheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickervolumes(i)
        Cells(4 + i, 3).Value = tickerendingprice(i) / tickerstartingprice(i) - 1
    Next i
    
    'Formatting
    Sheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)

End Sub

	This returned the information I was running 0.3632813 seconds for the year 2017 and 0.0546875 for the year 2018.
	Insert picture

Summary: 
•	The advantage and disadvantages of refactoring code.
Refactoring is a key part of the coding process. When refactoring code is more efficient by using less memory or improving the logic of the code to make it easier for future users to read, Refactoring code does have a disadvantage as well. Code refactoring and restructuring are time-consuming. It’s not a project you can complete in a short time. Depending on the size and complexity of your app, code refactoring can take months to complete.  If the code was properly annotated, then it could be hard to decode the purpose of a certain line of the section of code.
•	Pros and cons of refactoring code
 Refactoring code is a great way to explore alternative methods it also allowed further opportunities to debug different types of coding issues. Your code will be better organized for refactoring It won’t be any more function than it was before the change. Refactoring is just changing the structure of the code. 
