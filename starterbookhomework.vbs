
Sub starterbook():
    Dim orgRow As Integer
    Dim stateColumn As Integer
    For i=2 to 4115
    State= Cells(i,6).Value
        If (State = "successful") Then  
            Cells(i,6).Interior.ColorIndex = 4
        
        ElseIf (State = "failed") Then
            Cells(i,6).Interior.ColorIndex = 3
        
        ElseIf (State = "canceled") Then
            Cells(i,6).Interior.ColorIndex = 5
        Else
            Cells(i,6).Interior.ColorIndex = 6
        End If
    next i 
End Sub 

Sub PercentFunded():
    Dim PercentFunded As String
    Dim funding As Double
    Dim pledged As Double
    Dim goal As Double
    Dim total As Double
    SummaryRow = 2
    Range("O1").EntireColumn.Insert
    Cells(1,15).Value= "Percent Funded"
    For i = 2 to 4115
        pledged = cells(i,5).Value
        goal = cells(i,4).Value
        total = cells(i,15).Value
        If i >1 Then 
            cells(i,15).Value = ((cells(i,5).Value)/(cells(i,4).Value))*100 
        Else
        End If 
    Next i
End Sub 

Sub PercentFunded1():
    Dim column As Integer
    Dim funding As Double
    For i = 2 to 4115
        If i >= 200 Then
            cells(i,15).Interior.ColorIndex = 5 
        ElseIf i >= 100 Then
            cells(i,15).Interior.ColorIndex = 4
        ElseIf i >= 0 Then
            cells(i,15).Interior.ColorIndex = 9
        Else
        End If
    Next i
End Sub 


Sub AverageDonation():
    Dim AverageDonation As String
    Range("P1").EntireColumn.Insert
    Cells(1,16).Value= "Average Donation"
End Sub

Sub Category():
    Dim Category As String
    Range("Q1").EntireColumn.Insert
    Cells(1,17).Value="Category"
End Sub

Sub subCategory():
    Dim subCategory As String
    Range("R1").EntireColumn.Insert
    Cells(1,18).Value="Sub-Category"
End Sub

Sub DateCreatedConversion():
    Dim DateCreatedConversion As String
    Range("S1").EntireColumn.Insert
    Cells(1,19).Value="Date Created Conversion"
End Sub

Sub DateEndedConversion():
    Dim DateEndedConversion As String
    Range("T1").EntireColumn.Insert
    Cells(1,20).Value= "Date Ended Conversion"
End Sub


Set xlsWorksheet = workbook.easy_getSheet("Second tab")