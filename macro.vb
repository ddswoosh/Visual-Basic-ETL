Option Explicit

Dim numberOfRows As Integer
Dim currentRange As Range
Dim currentCase As Range
Dim currentVisitDurationCell As Range
Dim currentDrivingDurationCell As Range

Dim minutesRow As Integer
Dim hoursRow As Integer
Dim combinedTimeRow As Integer

Const checkInColumn = "F"
Const checkOutColumn = "G"
Const visitDurationColumn = "H"
Const drivingDurationColumn = "I"
Const badData = 0
Const missingData = "N/A"


Sub Format_CRM_CSV_File()
'
' Format the crm csv file
'
'
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:I").Select
    Columns("A:I").EntireColumn.AutoFit
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Rows.AutoFit
    Range("A1:G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    numberOfRows = Selection.Rows.Count

    'Calculate_Visit_Durations
    'Calculate_Drive_Durations
    VisitDuration
    DrivingDuration
    Combined_Duration
    Copy_All_Used_Cells
End Sub
Function GetNumberOfRows() As Integer

Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select

GetNumberOfRows = Selection.Rows.Count
End Function

Sub VisitDuration()
'
' Macro4 Macro
'
    Dim counter As Integer
    
    numberOfRows = GetNumberOfRows
    For counter = 1 To numberOfRows
        Set currentVisitDurationCell = Range(Cells(counter, visitDurationColumn), Cells(counter, visitDurationColumn))
        Set currentRange = Range(Cells(counter, checkInColumn), Cells(counter, checkInColumn))
        Set currentCase = Cells(counter, checkOutColumn)
        
        Select Case currentCase.Value
        Case Is = Empty
            currentVisitDurationCell.Value = missingData
        Case Is >= currentRange.Value
            currentVisitDurationCell.Value = WorksheetFunction.RoundUp((currentCase.Value - currentRange.Value) * 1440, 0)
        Case Else
            currentVisitDurationCell.Value = badData
        End Select
        
    Next counter

End Sub
Sub DrivingDuration()

    numberOfRows = GetNumberOfRows
    

    Set currentDrivingDurationCell = Range(Cells(1, drivingDurationColumn), Cells(1, drivingDurationColumn))
    
    If IsEmpty(Cells(1, checkOutColumn).Value) Then currentDrivingDurationCell.Value = missingData Else currentDrivingDurationCell.Value = 0
    
    Dim counter As Integer
        
    For counter = 2 To numberOfRows
       Set currentRange = Range(Cells(counter, drivingDurationColumn), Cells(counter, drivingDurationColumn))
        Set currentCase = Cells(counter - 1, checkOutColumn)
        
        Select Case currentCase.Value
        Case Is = Empty
            currentRange.Value = missingData
        Case Is <= Cells(counter, checkInColumn).Value
            currentRange.Value = WorksheetFunction.RoundUp((Cells(counter, checkInColumn).Value - currentCase.Value) * 1440, 0)
        Case "N/A"
            currentRange.Value = "N/A"
        Case Else
            currentRange.Value = badData
        End Select
    Next counter
    
End Sub
Sub Copy_All_Used_Cells()
    Range("A1", Cells(numberOfRows + 4, drivingDurationColumn)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
End Sub
Sub Combined_Duration()

    minutesRow = numberOfRows + 2
    hoursRow = numberOfRows + 3
    combinedTimeRow = numberOfRows + 4
    
    Dim visitMinutesCell As Range
    Dim drivingMinutesCell As Range
    
    Dim visitHoursCell As Range
    Dim drivingHoursCell As Range
    Dim combinedTimeCell As Range
    
    
    Set visitMinutesCell = Cells(minutesRow, visitDurationColumn)
    Set drivingMinutesCell = Cells(minutesRow, drivingDurationColumn)
    
    Set visitHoursCell = Range(Cells(hoursRow, visitDurationColumn), Cells(hoursRow, visitDurationColumn))
    Set drivingHoursCell = Range(Cells(hoursRow, drivingDurationColumn), Cells(hoursRow, drivingDurationColumn))
    
    
    Set combinedTimeCell = Cells(combinedTimeRow, visitDurationColumn)
    
    combinedTimeCell = checkOutColumn & combinedTimeRow
    
    Range(Cells(minutesRow, "E"), Cells(minutesRow, "E")).Value = "Number of visits: " & numberOfRows
    Range(Cells(minutesRow, checkOutColumn), Cells(minutesRow, checkOutColumn)).Value = "Total (in minutes)"
    visitMinutesCell.Value = WorksheetFunction.Sum(Range(Cells(1, visitDurationColumn), Cells(numberOfRows, visitDurationColumn)))
    drivingMinutesCell.Value = WorksheetFunction.Sum(Range(Cells(1, drivingDurationColumn), Cells(numberOfRows, drivingDurationColumn)))


    Range(Cells(hoursRow, checkOutColumn), Cells(hoursRow, checkOutColumn)).Value = "Total (in hours)"
    visitHoursCell.Value = visitMinutesCell / 60
    drivingHoursCell.Value = drivingMinutesCell / 60

    Range(Cells(combinedTimeRow, checkOutColumn), Cells(combinedTimeRow, checkOutColumn)).Value = "Combined Time"
    combinedTimeCell.Value = visitHoursCell.Value + drivingHoursCell.Value

End Sub


