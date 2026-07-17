Attribute VB_Name = "Module1"
Sub RenumberSerialNumbers()

    Dim tbl As Table
    Dim i As Long
    Dim Sr As Long

    Set tbl = ActiveDocument.Tables(1)

    Sr = 1

    For i = 3 To tbl.Rows.Count    'Start numbering from Row 3
        tbl.Cell(i, 1).Range.Text = Sr
        Sr = Sr + 1
    Next i

    MsgBox "Serial numbers updated successfully!", vbInformation

End Sub
