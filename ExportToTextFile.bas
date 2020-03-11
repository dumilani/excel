Public Sub ExportToTextFile(FName As String, _
    Sep As String, SelectionOnly As Boolean, _
    AppendData As Boolean)

Dim WholeLine As String
Dim FNum As Integer
Dim RowNdx As Long
Dim ColNdx As Integer
Dim StartRow As Long
Dim EndRow As Long
Dim StartCol As Integer
Dim EndCol As Integer
Dim CellValue As String


Application.ScreenUpdating = False
On Error GoTo EndMacro:
FNum = FreeFile

If SelectionOnly = True Then
    With Selection
        StartRow = .Cells(1).Row
        StartCol = .Cells(1).Column
        EndRow = .Cells(.Cells.Count).Row
        EndCol = .Cells(.Cells.Count).Column
    End With
Else
    With ActiveSheet.UsedRange
        StartRow = .Cells(1).Row
        StartCol = .Cells(1).Column
        EndRow = .Cells(.Cells.Count).Row
        EndCol = .Cells(.Cells.Count).Column
    End With
End If

If AppendData = True Then
    Open FName For Append Access Write As #FNum
Else
    Open FName For Output Access Write As #FNum
End If

For RowNdx = StartRow To EndRow
    WholeLine = ""
    For ColNdx = StartCol To EndCol
        If Cells(RowNdx, ColNdx).Value = "" Then
            CellValue = Chr(34) & Chr(34)
        Else
           CellValue = Cells(RowNdx, ColNdx).Value
        End If
        WholeLine = WholeLine & CellValue & Sep
    Next ColNdx
    WholeLine = Left(WholeLine, Len(WholeLine) - Len(Sep))
    Print #FNum, WholeLine
Next RowNdx

EndMacro:
On Error GoTo 0
Application.ScreenUpdating = True
Close #FNum

End Sub
