# Stock-Analysis
## Overview
### Originally my friend Steve asked me to complete an analysis of the DQ stock for his parents. After the data showed that stock to have negative outcomes, Steve then asked
me to dive farther into other stocks and see what some other idea might be. Further, Steve wanted the ability to look up the stocks for 2017 and 2018 and to be able to run the
analyses at the push of a button.
## Results
### I was able to compile data sets for 2017 and 2018 where Steve is able to see how the stocks fared for the years. The results show only two stocks had positive outcomes for
both years, ENPH and RUN. The others were in the negative one or both years.
## Summary
### This data will be easy for Steve to see and understand, and even easier for him to explain to his parents. The entire formatting section of code:    
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
Allows Steve to see the positive and negative values for the stocks just by glancing. Steve now has all the tools to he needs to help his parents make an informed dedision
about stocks.
## Refactored VBA
### Advantages Refactoring VBA allows it to run faster. Even if there is a lot of data, it can run quickly through everything, as if condensing everything for efficiency. 
### Disadvatages I found the biggest disadvantage to be preciseness. You have to be even more exact with your understading in refactoring and really understand what you are
doing. This was definitely a struggle. 

