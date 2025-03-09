Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rngDate As Range
    Dim calendarName As String
    Dim wasVisible As Boolean
    Static oldValue As Variant ' Use Static to preserve the value across different cell selections
    
    ' Set the calendar shape name
    calendarName = "Calendar"
    
    ' Define the date range (Update with actual range reference)
    Set rngDate = Range("D6") ' Change this to your actual date range
    
    ' Store the current value of the merged cell range (from the first cell) **before clearing it** only when it's the first time
    If Not wasVisible And Not Intersect(Target, rngDate) Is Nothing Then
        If rngDate.MergeArea.Cells(1, 1).Value <> "" Then
            oldValue = rngDate.MergeArea.Cells(1, 1).Value
            Debug.Print "Stored Old Value: " & oldValue
        End If
    End If
    
    ' Store the current visibility state of the calendar
    wasVisible = Shapes(calendarName).Visible
    
    ' Show or hide the calendar based on selection
    If Not Intersect(rngDate, Target) Is Nothing Then
        ' Only clear the value if the merged cell is not already empty
        If rngDate.MergeArea.Cells(1, 1).Value <> "" Then
            rngDate.MergeArea.ClearContents
            Debug.Print "Cleared Value"
        End If
        
        If Not wasVisible Then
            ' Only show if it wasn't already visible
            Shapes(calendarName).Visible = True
            Shapes(calendarName).Top = ActiveCell.Top + ActiveCell.Height
            Shapes(calendarName).Left = ActiveCell.Left + ActiveCell.Width
            Debug.Print "Calendar Visible"
        End If
    Else
        If wasVisible Then
            ' Only hide if it was previously visible
            Shapes(calendarName).Visible = False
            Debug.Print "Calendar Hidden"
            
            ' If the cell is empty, restore the old value
            If rngDate.MergeArea.Cells(1, 1).Value = "" Then
                rngDate.MergeArea.Cells(1, 1).Value = oldValue
                Debug.Print "Restoring Old Value: " & oldValue
            End If
        End If
    End If
End Sub
