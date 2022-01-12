Attribute VB_Name = "Module1"
Sub Stocks():

Dim continue As Boolean
continue = True


Dim ticker As String
Dim next_ticker As String


Dim open_value As Double
Dim close_value As Double
Dim volume As Double


Dim row As Double
row = 2


Dim first_open As Boolean
first_open = True


Dim row_results As Double
row_results = 2

Do While continue = True

    ticker = Cells(row, 1).Value
    next_ticker = Cells(row + 1, 1).Value


    volume = volume + Cells(row, 7).Value

    
    
    If first_open = True Then
        open_value = Cells(row, 3).Value
        first_open = False
    
    End If
    
    
    
    If ticker <> next_ticker Then
        first_open = True
        close_value = Cells(row, 6)
        
        Cells(row_results, 9).Value = ticker
        
        Cells(row_results, 10).Value = (close_value) - (open_value)
        
            If Cells(row_results, 10).Value > 0 Then
            Cells(row_results, 10).Interior.ColorIndex = 4
            
            ElseIf Cells(row_results, 10).Value = 0 Then
            Cells(row_results, 10).Interior.ColorIndex = 6
            
            Else
            Cells(row_results, 10).Interior.ColorIndex = 3
            
            End If
        
        
        If open_value = 0 Then
        Cells(row_results, 11).Value = "N/A"
        
        Else
        Cells(row_results, 11).Value = ((Cells(row_results, 10).Value) / (open_value) * 100)
        
        End If
        
        Cells(row_results, 12).Value = volume
        volume = 0
        
        row_results = row_results + 1
        
        

    End If

    If next_ticker = "" Then
    continue = False
    
    End If


row = row + 1



Loop

End Sub

