##Purpose

The purpose of this was to collect stock information in the year 2017 and 2018 To give Steve and his parents the results to see what stocks are worth investing it. WE want to simplify all the info so with a click of a button Steve and his parents can get the information they need

##The Data

The data has two charts on 12 different stocks. The stock information contains a ticker value, the volume of the stock, Stock opening, closing and closing price, the high and low price. 

##Results

2017 was a great year for the stock, in 2018 the only stock that ended up having a better return was "TERP". Every other stock in. 2018 went down and had far less returns that its previous year. The Screentshots are in the repo. Here is the Code

Sub AllStocksAnalysis()



    Dim startTime As Single
    
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
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
    
     Worksheets(yearValue).Activate
    
  
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerindex As Integer
    
    tickerindex = 0
    
     '1b) Create three output arrays
     
        Dim tickerVolumes(12) As Long
        
        Dim tickerStartingPrices(12) As Single
        
        Dim tickerEndingPrices(12) As Single

     
        
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    
    
    For i = 0 To 11
    
    tickerVolumes(i) = 0
    
    tickerStartingPrices(i) = 0
    
    tickerEndingPrices(i) = 0
    
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    
            For i = 2 To RowCount
            
            Next i
            

    '3a) Increase volume for current ticker
    
             tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
             
    '3b) Check if the current row is the first row with the selected tickerIndex.
    

            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
    
        tickerStartingPrices(tickerindex) = Cells(i, 6).Value
        
    End If
    
     '3c) check if the current row is the last row with the selected ticker
     

     If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
     
        tickerEndingPrices(tickerindex) = Cells(i, 6).Value
        
     End If
     
     '3d Increase the tickerIndex.
     
         If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
         
            tickerindex = tickerindex + 1
            
        End If
        
        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
        For i = 0 To 11
        
        Next i
        
                Worksheets("All Stocks Analysis").Activate
        
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
   
    
##Summary

##Pros and Cons of Refactoring Code

Refactoring code has many benefits not only for us but for other users that use it. Refactoring can make the code perform quicker which was one of our tasks in this project, and it cleans up the code. We are also able to help others who use it and simplify it with comments and makes it very easy for them to follow and not get confused. The main disadvantage for Refactoring will be on projects that are way too big to be able to refactor and code


##Pros and Cons of Refactor vs Original

The pros of the refactor was decreasing the code time and leaving more clear instructions for others to use. I dont really think there is any cons to the refactor. We were able to make it code faster and even make it more clear and easier to use than before.
