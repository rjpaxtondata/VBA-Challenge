Attribute VB_Name = "Module3"
Sub stock()

 ' --------------------------------------------
'   loop through each sheet &worksheet-related steps
 ' --------------------------------------------
 
For Each ws In Worksheets

'-----create variable for worksheets
    Dim Wksheetname As String

'------grab the worksheet name
    Wksheetname = ws.Name

'------find last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'----------find last column - may not need
   ' LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
'----------activate the worksheets
    ws.Activate
    

 ' --------------------------------------------
'   create variables and initialize them
 ' --------------------------------------------
 
    Dim ticker As String
    ticker = ""
    
    Dim TotStockVal As Double
    TotStockVal = 0
'Yearly Change column header
    Dim yrchange As Double
        yrchange = 0

'Percent Change column header
    Dim perchange As Double
        perchange = 0

'Greatest % Increase header
    Dim Bigincrease As Double
        Bigincrease = 0
    
'Greatest % Decrease header
    Dim Bigdecrease As Double
        Bigdecrease = 0

'Greatest Total volume header
    Dim Bigvolume As Integer
        Bigvolume = 0
  Dim open_value As Double
  Dim close_value As Double
  open_value = 0
  close_value = 0
    
'----------------------------------------------
  ' Summary table results
'---------------------------------------------

'--------Sum table location
Dim SumTable_Row As Integer
    SumTable_Row = 2
  

'----------------------------
'Loop through stock days
'----------------------------
For i = 2 To lastrow


    

        'find ticker location
            ticker = ws.Cells(i, 1).Value

        'Calculate estimates for key variables
    TotStockVal = TotStockVal + ws.Cells(i, 7).Value
   
        'find location for opening price
            open_value = ws.Cells(i, 3).Value
            
       ' Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        
        'find location for closing price
            close_value = ws.Cells(i, 6).Value
            
            'create yearly change variable
                yrchange = close_value - open_value
        
        'To fix division problem
            If open_value = 0 Then
                perchange = close_value
                Else
        'calculation for %change
                perchange = (yrchange / open_value) * 100
                
                     
        End If
        
   
        'set headers for summary table
            ws.Range("I" & SumTable_Row).Value = ticker
            ws.Range("J" & SumTable_Row).Value = yrchange
            ws.Range("K" & SumTable_Row).Value = perchange
            ws.Range("L" & SumTable_Row).Value = TotStockVal
           'format cell
            ws.Range("K" & SumTable_Row).NumberFormat = "0.00"
          
        ' Add one to the summary table row
            SumTable_Row = SumTable_Row + 1

              
      
        
        'Reset values back to zeroe
            TotStockVal = 0
            yrchange = 0
            perchance = 0
           


Else
    
    
    TotStockVal = TotStockVal + ws.Cells(i, 7).Value
 '     yrchange = close_value - open_value
'    perchange = yrchange / open_value
    
 End If
            'conditional formating gain vs loss
  
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
               
Next i


Next ws


End Sub

