Attribute VB_Name = "Module1"
Option Explicit

Sub InsertThermalDigitalImages()

    Dim tbl As Table
    Dim r As Long
    Dim imgFolder As String
    Dim thermalFile As String
    Dim digitalFile As String
    Dim imgPath As String
    Dim rng As Range
    Dim pic As InlineShape
    Dim MissingFiles As String

    'Pick image folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Images"
        If .Show <> -1 Then Exit Sub
        imgFolder = .SelectedItems(1)

        If Right(imgFolder, 1) <> "\" Then
            imgFolder = imgFolder & "\"
        End If
    End With

    Set tbl = ActiveDocument.Tables(1)

    For r = 2 To tbl.Rows.Count

        '================ THERMAL ==================
        thermalFile = CleanCellText(tbl.Cell(r, 4).Range.Text)

        If thermalFile <> "" Then

            If LCase(Right(thermalFile, 4)) <> ".jpg" Then
                thermalFile = thermalFile & ".jpg"
            End If

            imgPath = imgFolder & thermalFile

            If Dir(imgPath) <> "" Then

                Set rng = tbl.Cell(r, 5).Range
                rng.End = rng.End - 1
                rng.Text = ""

                Set pic = rng.InlineShapes.AddPicture( _
                    FileName:=imgPath, _
                    LinkToFile:=False, _
                    SaveWithDocument:=True)

                pic.LockAspectRatio = False
                pic.Width = InchesToPoints(2.5)
                pic.Height = InchesToPoints(1.8)

                tbl.Cell(r, 5).VerticalAlignment = wdCellAlignVerticalCenter
                rng.ParagraphFormat.Alignment = wdAlignParagraphCenter

            Else
                MissingFiles = MissingFiles & thermalFile & vbCrLf
            End If

        End If

        '================ DIGITAL ==================
        digitalFile = CleanCellText(tbl.Cell(r, 6).Range.Text)

        If digitalFile <> "" Then

            If LCase(Right(digitalFile, 4)) <> ".jpg" Then
                digitalFile = digitalFile & ".jpg"
            End If

            imgPath = imgFolder & digitalFile

            If Dir(imgPath) <> "" Then

                Set rng = tbl.Cell(r, 7).Range
                rng.End = rng.End - 1
                rng.Text = ""

                Set pic = rng.InlineShapes.AddPicture( _
                    FileName:=imgPath, _
                    LinkToFile:=False, _
                    SaveWithDocument:=True)

                pic.LockAspectRatio = False
                pic.Width = InchesToPoints(2)
                pic.Height = InchesToPoints(1.8)

                tbl.Cell(r, 7).VerticalAlignment = wdCellAlignVerticalCenter
                rng.ParagraphFormat.Alignment = wdAlignParagraphCenter

            Else
                MissingFiles = MissingFiles & digitalFile & vbCrLf
            End If

        End If

    Next r

    If MissingFiles = "" Then
        MsgBox "Finished successfully!", vbInformation
    Else
        MsgBox "Finished." & vbCrLf & vbCrLf & _
               "These images were not found:" & vbCrLf & _
               MissingFiles, vbExclamation
    End If

End Sub

Function CleanCellText(txt As String) As String

    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(7), "")
    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, vbLf, "")
    txt = Trim(txt)

    CleanCellText = txt

End Function



