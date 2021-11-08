# Green Stock Analysis

## Object Overview

### Background

#### Steve knew how to use Excel, but he needed  assistance  analyzing data to help his parents invest in renewable energy. He had all the information necessary and wanted a more efficient	way to view the data. By using Visual Basic Application (VBA), I was able to create an useable interactive code within Excel to present each stocks return on investment and annual volume. By doing so, I was able to help Steve view 12 renewable stocks with a click of a button.

### Purpose

#### Steve requires a higher number of stocks to be analyzed and I will help him make it easier. Even though it may take a longer to present the data, I’m here to improve process and make it faster. By taking the workbooks that were provided, I refactored the VBA code to make it more effecitve.

## Results

### Refactoring the Code
	
#### In order to make this more efficient, I generated three new arrays for the code: -tickersVolumes(12) to hold volume - tickerStartingPrices(12) to hold starting price -tickerEnding(12) to hold ending price. These arrays hold the data for the loop. By creating the ticker array in the original code establishes a ticker symbol that can be called for each stock. Then  by genernating a variable named ticker index, I was able to match the three arrays with the tickers. With the arrays  constructed, I’m able to operate the Nested For Loops and variables to loop through the data and complete the analysis.

#### Refactored

    Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Title Analysis
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Count the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker index to reference proper ticker in the arrays.
    Dim tickerIndex As Integer
    'Initiate tickerIndex at zero.
    tickerIndex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    '2a) Create for loop to analyze each ticker in the array.
    For tickerIndex = 0 To 11
    'Initiate each ticker's volume at zero.
    tickerVolumes(tickerIndex) = 0
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
        
        '2b) Loop over all the rows in the spreadsheet.
        For I = 2 To RowCount
        
            '3a) Increase volume for current ticker.
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(I, 8).Value
    
        
            '3b) Check if the current row is the first row with the current ticker.
                    
            If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I - 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the first row for current ticker, set starting price.
                tickerStartingPrices(tickerIndex) = Cells(I, 6).Value
            
            'End If
            End If
            
            
        '3c) Check if the current row is the last row with the current ticker.

            If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the last row for current ticker, set ending price.
                tickerEndingPrices(tickerIndex) = Cells(I, 6).Value
            
            'End if
            End If
            
        '3d) Check if the current row is the last row with the current ticker.
            If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
                
                'if it is, increase tickerIndex to move on to next ticker in array.
                tickerIndex = tickerIndex + 1
            
            'End If
            End If
    
        Next I
        
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For I = 0 To 11
        
        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker Row Label
        Cells(4 + I, 1).Value = tickers(I)
        
        'Sum of Volume
        Cells(4 + I, 2).Value = tickerVolumes(I)
        
        'ReturnValue
        Cells(4 + I, 3).Value = tickerEndingPrices(I) / tickerStartingPrices(I) - 1
        

    
    Next I

#### Original

     Sub AllStocksAnalysis()
      
     Dim startTime As Single
     Dim endTime As Single
      
    yearValue = InputBox("What year would you like to run the analysis on?")
       
        startTime = Timer
        
 	'1)  Format the output sheet on the "All Stocks Analysis" worksheet.
    
    'Activate "All Stocks Analysis" Worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'TitleAnalysis
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        
               'Create header row
                Cells(3, 1).Value = "Ticker"
                Cells(3, 2).Value = "Total Daily Volue"
                Cells(3, 3).Value = "Return"
        
	'2) Initialize an array of all tickers.
    
    'Declare an array with 12 string elements
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
    
	'3) Preparre for the analysis of all stickers

	'3a) Initialize variables for the starting price and ending price.
       
       'Creating a Variable for Starting & Ending Price
      Dim startingPrice As Single
      Dim endingPrice As Single

	'3b) Activate the data worksheet.
    
    Worksheets(yearValue).Activate
    
	'3c) Find the number of rows to loop over.
    
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
 
 
 	'4)  Loop through the tickers.
       
    For I = 0 To 11
    
        ticker = tickers(I)
        totalVolume = 0
        
       
       '5)  Loop through rows in the data.
    
        'Activate Data Worksheet
        Worksheets(yearValue).Activate
        
        For j = 2 To RowCount

      '5a) Find the total volume for the current ticker.
      
      'Identify ticker
      If Cells(j, 1).Value = ticker Then
    
        totalVolume = totalVolume + Cells(j, 8).Value
        
        End If
        
        
        '5b) Find the starting price for the current ticker.
    
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            startingPrice = Cells(j, 6).Value

        End If

        '5c) Find the ending price for the current ticker.
    
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            endingPrice = Cells(j, 6).Value

        End If

    Next j
    
    '6)  Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + I, 1).Value = ticker
        Cells(4 + I, 2).Value = totalVolume
        Cells(4 + I, 3).Value = endingPrice / startingPrice - 1
        

    Next I

    'Formatting
     Worksheets("All Stocks Analysis").Activate
     Range("A3:C3").Font.Bold = True
     Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
     Range("B4:B15").NumberFormat = "#,#0"
     Range("C4:C15").NumberFormat = "0.0%"
     Columns("B").AutoFit
    
    If Cells(4, 3) > 0 Then
        'Color the cell green
        Cells(4, 3).Interior.Color = vbGreen
    ElseIf Cells(4, 3) < 0 Then
    
        'Color the cell red
        Cells(4, 3).Interior.Color = vbRed
    
    Else
        'Clear the cell color
        Cells(4, 3).Interior.Color = xlNone
    
    End If
    
        dataRowStart = 4
        dataRowEnd = 15
        For I = dataRowStart To dataRowEnd
    
            If Cells(I, 3) > 0 Then
    
                'Color the cell green
                Cells(I, 3).Interior.Color = vbGreen
    
            ElseIf Cells(I, 3) < 0 Then
    
                'Color the cell red
                Cells(I, 3).Interior.Color = vbRed
    
            Else
    
                'Clear the cell color
                Cells(I, 3).Interior.Color = xlNone
    
            End If

    Next I
    
    endTime = Timer
    MsgBox " This code ran in  " & (endTime - startTime) & " seconds for the year " & (yearValue)

#### 2017 vs 2018 Stock Performance

##### There is a great difference between the performance of the renewable energy stocks from 2017 to 2018. ENPH and RUN made a positive ROI in both years. Majority of the  stock had a decline in volume in the year 2018.

![This is an image](https://github.com/daryld2239/stock-analysis/blob/main/Resources/Stocks%202017.png)
![This is an image](https://github.com/daryld2239/stock-analysis/blob/main/Resources/Stocks%202018.png)

##### Before Steven produces this information to his parents, he should observe the industry related and investment reasons they should invest. By observing the data, it would be better for his parents to invest an other industries.

#### Execution Time

##### The changes I made to the code proved to be successful. The time improved from 0.6484375 seconds to 0.1328125 for 2017, and, 0.65625 to 0.2328125 for 2018.

![This is an image](https://github.com/daryld2239/stock-analysis/blob/main/Resources/Stocks%202018.png)

![This is an image](https://github.com/daryld2239/stock-analysis/blob/main/Resources/Refactored%202018%20Time.png)

![This is an image](https://github.com/daryld2239/stock-analysis/blob/main/Resources/2017%20Time.png)

![This is an image](https://github.com/daryld2239/stock-analysis/blob/main/Resources/Refactored%202018%20Time.png)

## Summary

### Advantages of refactoring code

#### The advantage of refracting code it make be more efficient, as long as it done carefully and everything is correct. By reducing the time it took to produce data can make a big difference in analyzing data.

### Disadvantages of refactoring code

#### While refactoring code has its advantages, it also has many disadvantages. When refactoring code, you must make sure that you are keeping the data you use intact, any slight change could corrupt the code. Any minor mistake may take away time for analyzing data. By an error being in a code, that may be lengthy could take a lot of time to find. By testing the outputs, I was able to find the mistake but it was time consuming. 
