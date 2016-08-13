Attribute VB_Name = "Module1"
Sub iterateTable()
'
' full_table_macro Macro
'
'
    Dim oRow As Row
    Dim oCell As Cell
    Dim sCellText As String
    
    
    For Each oRow In ActiveDocument.Tables(1).Rows
        For Each oCell In oRow.Cells
            sCellText = oCell.Range
            sCellText = Left$(sCellText, Len(sCellText) - 2)
            Debug.Print sCellText
        Next
    Next oRow
    
    MsgBox "All done"

End Sub

