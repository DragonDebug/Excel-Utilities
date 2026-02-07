Option Explicit

Private Function GetSelectedRange() As Range
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select one or more cells.", vbInformation
        Exit Function
    End If

    If Selection.Cells.CountLarge = 0 Then
        MsgBox "Please select one or more cells.", vbInformation
        Exit Function
    End If

    Set GetSelectedRange = Selection
End Function

Private Function ShouldProcessTextCell(ByVal cell As Range) As Boolean
    ' Only act on non-empty text constants; skip formulas, arrays, numbers, dates, and errors.
    If cell Is Nothing Then Exit Function
    If cell.HasArray Or cell.HasFormula Then Exit Function

    Dim v As Variant
    v = cell.Value2

    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function

    If VarType(v) = vbString Then
        ShouldProcessTextCell = True
    End If
End Function

Sub UpperMacro(control As IRibbonControl)
    Dim target As Range
    Dim c As Range

    Set target = GetSelectedRange()
    If target Is Nothing Then Exit Sub

    CaptureUndo target

    For Each c In target.Cells
        If ShouldProcessTextCell(c) Then
            c.Value = UCase$(c.Value2)
        End If
    Next c

    RegisterUndo "Undo Uppercase"
End Sub

Sub LowerMacro(control As IRibbonControl)
    Dim target As Range
    Dim c As Range

    Set target = GetSelectedRange()
    If target Is Nothing Then Exit Sub

    CaptureUndo target

    For Each c In target.Cells
        If ShouldProcessTextCell(c) Then
            c.Value = LCase$(c.Value2)
        End If
    Next c

    RegisterUndo "Undo Lowercase"
End Sub

Sub ProperMacro(control As IRibbonControl)
    Dim target As Range
    Dim c As Range

    Set target = GetSelectedRange()
    If target Is Nothing Then Exit Sub

    CaptureUndo target

    For Each c In target.Cells
        If ShouldProcessTextCell(c) Then
            c.Value = WorksheetFunction.Proper(c.Value2)
        End If
    Next c

    RegisterUndo "Undo Proper Case"
End Sub

Sub TransformToFormulaMacro(control As IRibbonControl)
    Dim target As Range
    Dim c As Range
    Dim failedCount As Long

    Set target = GetSelectedRange()
    If target Is Nothing Then Exit Sub

    CaptureUndo target

    For Each c In target.Cells
        If ShouldProcessTextCell(c) Then
            Dim txt As String
            txt = CStr(c.Value2)

            If Len(txt) = 0 Then GoTo NextCell
            If Left$(Trim$(txt), 1) = "=" Then GoTo NextCell

            If Left$(txt, 1) = "'" Then
                txt = Mid$(txt, 2)
            End If

            On Error Resume Next
            c.Formula = "= " & txt
            If Err.Number <> 0 Then
                failedCount = failedCount + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
NextCell:
    Next c

    If failedCount > 0 Then
        MsgBox failedCount & " cell(s) could not be converted to a formula and were skipped.", vbInformation
    End If

    RegisterUndo "Undo Convert to Formula"
End Sub

Sub PasteAsValuesMacro(control As IRibbonControl)
    Dim target As Range
    Dim c As Range
    Dim wasScreenUpdating As Boolean

    Set target = GetSelectedRange()
    If target Is Nothing Then Exit Sub

    CaptureUndo target

    wasScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo RestoreState

    target.Value = target.Value
    RegisterUndo "Undo Paste Values"
    GoTo RestoreState

RestoreState:
    Application.ScreenUpdating = wasScreenUpdating
    On Error GoTo 0
End Sub

Sub TrimMacro(control As IRibbonControl)
    Dim target As Range
    Dim c As Range

    Set target = GetSelectedRange()
    If target Is Nothing Then Exit Sub

    CaptureUndo target

    For Each c In target.Cells
        If ShouldProcessTextCell(c) Then
            c.Value = WorksheetFunction.Trim(c.Value2)
        End If
    Next c

    RegisterUndo "Undo Trim"
End Sub
