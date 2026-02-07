Option Explicit

Private gUndoRangeAddress As String
Private gUndoSheetName As String
Private gUndoFormulas As Variant
Private gUndoReady As Boolean

Private Sub CaptureUndo(ByVal target As Range)
    If target Is Nothing Then Exit Sub

    gUndoRangeAddress = target.Address
    gUndoSheetName = target.Worksheet.Name
    gUndoFormulas = target.Formula
    gUndoReady = True

End Sub

Public Sub RegisterUndo(ByVal undoLabel As String)
    If Not gUndoReady Then Exit Sub
    Application.OnUndo undoLabel, "UndoLastMacro"
End Sub

Public Sub UndoLastMacro()
    If Not gUndoReady Then
        MsgBox "Nothing to undo.", vbInformation
        Exit Sub
    End If

    On Error GoTo UndoFailed
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(gUndoSheetName)

    ws.Range(gUndoRangeAddress).Formula = gUndoFormulas
    gUndoReady = False
    Exit Sub

UndoFailed:
    MsgBox "Undo failed. The original range may no longer be available.", vbExclamation
End Sub