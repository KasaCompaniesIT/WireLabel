Attribute VB_Name = "ObsoleteCode"
'===============================
'   BEGIN OBSOLETE CODE
'===============================

'-------------------------------------------------
'   BEGIN   frmFormat code
'-------------------------------------------------
'Private Sub lstFormats_Click()
'' This subroutine opens the label formats file and finds the label
'' format that the user selected and loads the values for that label
'' format into the respective text boxes.
'
'    On Error GoTo ErrorHandler
'
'    Dim strTemp As String
'    Dim intFile As Long
'    Dim blnFound As Boolean
'
'    ' If there are no formats then just exit the subroutine.
'    If lstFormats.ListCount = 0 Then
'        Exit Sub
'    End If
'
'    ' Find an available file location in memory
'    intFile = FreeFile
'
'    ' Open the label formats file for input
'    Open App.Path & "\" & LABEL_FORMATS For Input As #intFile
'
'    ' Set the label found indicator to false
'    blnFound = False
'
'    Do
'        ' Read in a line of the file
'        Input #intFile, strTemp
'        ' If the label format we found matches the one the user
'        ' clicked, then set the label found indicator to true.
'        If Trim$(strTemp) = "~" & Trim$(lstFormats.Text) Then
'            blnFound = True
'        End If
'    Loop Until EOF(intFile) Or blnFound         ' Keep looking until we reach the end of the file or we find the label
'
'    ' If we didn't find the label then let the user know, and close the file.
'    If blnFound = False Then
'        MsgBox "Item Not Found!", , "Error"
'        Close #intFile
'        Exit Sub
'    End If
'
'    ' We must have found the file, so go ahead and read
'    ' in all the values for that label format.
'    Input #intFile, strTemp
'    txtTop.Text = FormatLabelDimension(strTemp)
'    Input #intFile, strTemp
'    txtLeft.Text = FormatLabelDimension(strTemp)
'    Input #intFile, strTemp
'    txtWidth.Text = FormatLabelDimension(strTemp)
'    Input #intFile, strTemp
'    txtHeight.Text = FormatLabelDimension(strTemp)
'    Input #intFile, strTemp
'    txtSpacingTB.Text = FormatLabelDimension(strTemp)
'    Input #intFile, strTemp
'    txtSpacingRL.Text = FormatLabelDimension(strTemp)
'    Input #intFile, strTemp
'    txtLines.Text = Int(Val(strTemp))
'    Input #intFile, strTemp
'    txtLabelsPerRow.Text = Int(Val(strTemp))
'    Input #intFile, strTemp
'    chkOptical.Value = Val(strTemp)
'    Input #intFile, strTemp
'    chkAutoSize.Value = Val(strTemp)
'
'    ' Set the Description text box to the name of the label format.
'    txtDescription.Text = Trim$(lstFormats.Text)
'
'    ' Close the file we had open.
'    Close #intFile
'
'    Exit Sub
'ErrorHandler:
'    MsgBox "Error " & Err.Number & " - " & Err.Description, , "Error"
'End Sub
'-------------------------------------------------
'   END   frmFormat code
'-------------------------------------------------



'-------------------------------------------------
'   BEGIN   frmMain code
'-------------------------------------------------
'
'mnuPrint_Click SNIPPED A
    'CHG 200508 : INTEGRATED OPTICALPRINT ROUTINE INTO THIS ROUTINE, BELOW.
    'If the labels are optically sensed, run OpticalPrint subroutine instead
'    If intOptical = 1 Then
'        OpticalPrint
'        Exit Sub
'    End If
'-------------------------------------------------
'mnuPrint_Click SNIPPED B
'    ' If we are only printing the selected items, the only count
'    ' the quantities of the selected labels.
'    If Not blnPrintAll Then
'        ' Only count the selected labels
'        For Each itmItem In lsvLabels.ListItems
'            If itmItem.Selected = True Then
'                intTQty = intTQty + Val(itmItem.SubItems(1))
'            End If
'            intTemp = intTemp + Val(itmItem.SubItems(1))
'        Next itmItem
'        ' Ask user if they really want to print these labels
'        If intCopies > 1 Then
'            intResult = MsgBox(Trim$(Str$(intTQty)) & " / " & Trim$(Str$(intTemp)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copies." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
'        Else
'            intResult = MsgBox(Trim$(Str$(intTQty)) & " / " & Trim$(Str$(intTemp)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copy." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
'        End If
'    Else
'        ' Count all the labels in the list
'        For Each itmItem In lsvLabels.ListItems
'            intTQty = intTQty + Val(itmItem.SubItems(1))
'        Next itmItem
'        ' Ask user if they really want to print these labels
'        If intCopies > 1 Then
'            intResult = MsgBox(Trim$(Str$(intTQty)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copies." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
'        Else
'            intResult = MsgBox(Trim$(Str$(intTQty)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copy." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
'        End If
'    End If
    
    ' If the user decides not to print by clicking Cancel, just exit the subroutine
'    If intResult = vbCancel Then
'        Exit Sub
'    End If
'-------------------------------------------------
'mnuPrint_Click SNIPPED C
'    intBlank = intLabelsPerRow - (intTQty Mod intLabelsPerRow)
'    If intBlank > 0 Then
'        intTQty = intTQty + intBlank
'    End If
'-------------------------------------------------
'mnuPrint_Click SNIPPED D
'200508 : Obsolete.  not using continuous form tractor feed printers
'    If sngLeftMargin >= 0.5 Then
        ' There is a continuous form tractor feed strip on the side
'        sngPrintLeftX = sngLeftMargin - 0.45
'    Else
        ' Adjust the left margin so that it doesn't cut off
        ' any of the label text on the left side
'        sngPrintLeftX = sngLeftMargin - 0.125
'    End If

'-------------------------------------------------

'-------------------------------------------------
'
''200508 OBS: N.F. Integrated OpticalPrint functionality into the mnuPrint_click routine
'
'Private Sub OpticalPrint()
'' This subroutine prints the labels selected, or all of the labels
'' in the format specified in the label format, and only if they are
'' labels that are optically sensed.
'    Dim ActualTextHeight As Single
'
'    Dim labLabels() As MyLabel
'    Dim strTemp() As String
'    Dim intTQty As Long
'    Dim sngPrintLeftX As Single
'    Dim sngPrintTopY As Single
'    Dim itmItem As ListItem
'    Dim intResult As Long
'    Dim intTemp As Long
'    Dim intCnt As Long
'    Dim intLPR As Long
'    Dim intLPL As Long
'    Dim intBlank As Long
'    Dim blnTemp As Boolean
'    Dim lngLargest As Long
'    Dim blnMultLine2 As Boolean
'    Dim blnMultLine3 As Boolean
'    Dim lngLine As Long
'
'    TempText2 = ""
'
'    ' If there are no labels in the list, then just exit subroutine.
'    If lsvLabels.ListItems.Count = 0 Or frmPrint.Visible Then
'        Exit Sub
'    End If
'
'    intTQty = 0
'    intTemp = 0
'
'    ' If we are only printing the selected items, the only count
'    ' the quantities of the selected labels.
'    If Not blnPrintAll Then
'        ' Only count the selected labels
'        For Each itmItem In lsvLabels.ListItems
'            If itmItem.Selected = True Then
'                intTQty = intTQty + Val(itmItem.SubItems(1))
'            End If
'            intTemp = intTemp + Val(itmItem.SubItems(1))
'        Next itmItem
'        ' Ask user if they really want to print these labels
'        If intCopies > 1 Then
'            intResult = MsgBox(Trim$(Str$(intTQty)) & " / " & Trim$(Str$(intTemp)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copies." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
'        Else
'            intResult = MsgBox(Trim$(Str$(intTQty)) & " / " & Trim$(Str$(intTemp)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copy." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
'        End If
'    Else
'        ' Count all the labels in the list
'        For Each itmItem In lsvLabels.ListItems
'            intTQty = intTQty + Val(itmItem.SubItems(1))
'        Next itmItem
'        ' Ask user if they really want to print these labels
'        If intCopies > 1 Then
'            intResult = MsgBox(Trim$(Str$(intTQty)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copies." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
'        Else
'            intResult = MsgBox(Trim$(Str$(intTQty)) & " labels to print in " & Chr(34) & strLabelFormat & Chr(34) & " format." & vbCrLf & "This will print " & intCopies & " copy." & vbCrLf & "Using " & Printer.DeviceName & " on port " & Printer.Port & vbCrLf & vbCrLf & "Align label sheet in printer and click Ok.", vbOKCancel + vbInformation, "Print")
'        End If
'    End If
'
'    ' If the user decides not to print by clicking Cancel, just exit the subroutine
'    If intResult = vbCancel Then
'        Exit Sub
'    End If
'
'    ' Add room for blank labels to fill out each row
'    intBlank = intLabelsPerRow - (intTQty Mod intLabelsPerRow)
'    If intBlank > 0 Then
'        intTQty = intTQty + intBlank
'    End If
'
'    intTQty = intTQty * intCopies
'
'    ' Re-Diminsion the array to fit all the labels in the list in the array
'    ReDim labLabels(intTQty) As MyLabel
'
'    ' Set the Maximum progress bar value to the number of labels to be printed
'    prgProgress.Max = intTQty
'
'    intTemp = 0
'
'    ' Set the printers margins by adjusting the numbers
'    ' from the currently selected label format
'    If sngLeftMargin >= 0.5 Then
'        ' There is a continuous form tractor feed strip on the side
'        sngPrintLeftX = sngLeftMargin - 0.45
'    Else
'        ' Adjust the left margin so that it doesn't cut off
'        ' any of the label text on the left side
'        sngPrintLeftX = sngLeftMargin - 0.125
'    End If
'
'    ' Adjust the top margin so that no text is cut off
'    sngPrintTopY = sngTopMargin - 0.175
'
'    ' Set the current printer position
'    Printer.CurrentX = sngPrintLeftX * 1440
'    Printer.CurrentY = sngPrintTopY * 1440
'
'    ' Set the printer's font settings
'    With Printer
'        .FontBold = cdlDialog.FontBold
'        .FontItalic = cdlDialog.FontItalic
'        .FontName = cdlDialog.FontName
'        .FontSize = cdlDialog.FontSize
'        .FontStrikethru = cdlDialog.FontStrikethru
'        .FontUnderline = cdlDialog.FontUnderline
'    End With
'
'    ' Set the printed page's height and width
'    Printer.Height = sngSpacingTB * 1440
'    Printer.Width = intLabelsPerRow * sngSpacingRL * 1440 + sngPrintLeftX * 1440
'
'    PrintTemp "A"
'    Call DebugPrinter
'
'    ' Check the label height compared to the height of the
'    ' printed output to make sure the printed output will fit
'    ' on the label format that is currently selected
'    If (Printer.TextHeight("gjqy#^") * intLines) > (sngHeight * 1440) And intAutoSize = 0 Then
'        ' Label text will not fit vertically, so notify the user.
'        MsgBox "Using the current label format and the current font, the text is too tall for the label." & vbCrLf & "Please choose another label format or another font to print this label list.", vbOKOnly + vbInformation, "Text Too Tall!"
'        Exit Sub
'    End If
'
'        PrintTemp "sngWidth, sngHeight " & sngWidth & ", " & sngHeight
'        Debug.Print "sngWidth, sngHeight " & sngWidth & ", " & sngHeight
'
'    ' Add the label text for each label to the array, depending
'    ' on whether the print selection is selected or not.
'    If Not blnPrintAll Then
'        ' Add selected labels to the array
'        For intResult = 1 To intCopies
'            For Each itmItem In lsvLabels.ListItems
'                If itmItem.Selected = True Then
'                    For intCnt = 1 To Val(itmItem.SubItems(1))
'                        On Error Resume Next
'                        strTemp() = Split(itmItem.Text, "|", 3)
'                        With labLabels(intTemp)
'                            For lngLine = 0 To 2
'                                .strLabel(lngLine) = Trim$(strTemp(lngLine))
'
'                                If Len(.strLabel(lngLine)) = 0 Then
'                                    .strLabel(lngLine) = " "
'                                Else
'                                    If lngLine = 1 Then
'                                        blnMultLine2 = True
'                                    ElseIf lngLine = 2 Then
'                                        blnMultLine3 = True
'                                    End If
'                                End If
'                            Next
'                        End With
'                        On Error GoTo 0
'
'
'                        '200507 : NEW SIZING ROUTINE
'                        If Printer.TextWidth(labLabels(intTemp).strLabel(0)) > sngWidth * 1440 Then
'                            labLabels(intTemp).lngSize = SizeToText(sngWidth - 1.5 / 16, sngHeight - 1 / 16, ActualTextHeight, labLabels(intTemp).strLabel(0))
'                            PrintTemp "B"
'                            PrintTemp "labLabels(intTemp).lngSize = " & labLabels(intTemp).lngSize
'                        End If
'
'                        'CHG 200507 : PREVIOUS ROUTINE
'                        ' Check to see if the width of the text will fit horizontally
'                        ' On the printed label, so that no labels overlap the text
''                        For lngLine = 0 To 2
''                            If Printer.TextWidth(labLabels(intTemp).strLabel(lngLine)) > sngWidth * 1440 And intAutoSize = 0 Then
''                                ' Label width isn't wide enough to fit the text onto the label
''                                ' Notify the user about which label has too long of text
''                                If blnMultLine3 Then
''                                    MsgBox "Label " & Chr(34) & Join(labLabels(intTemp).strLabel, "|") & Chr(34) & " is wider than the label!", vbOKOnly + vbInformation, "Label Is Too Wide!"
''                                ElseIf blnMultLine2 Then
''                                    MsgBox "Label " & Chr(34) & labLabels(intTemp).strLabel(0) & "|" & labLabels(intTemp).strLabel(1) & Chr(34) & " is wider than the label!", vbOKOnly + vbInformation, "Label Is Too Wide!"
''                                Else
''                                    MsgBox "Label " & Chr(34) & labLabels(intTemp).strLabel(0) & Chr(34) & " is wider than the label!", vbOKOnly + vbInformation, "Label Is Too Wide!"
''                                End If
''                                prgProgress.Value = 0
''                                Exit Sub
''                            End If
''                        Next lngLine
'
'                        ' Update progress bar to show current status of loading labels into the array
'                        intTemp = intTemp + 1
'                        prgProgress.Value = intTemp
'                    Next intCnt
'                End If
'            Next itmItem
'            ' Fill the rest of the row with blank labels
'            For intCnt = 1 To intBlank
'                For lngLine = 0 To 2
'                    labLabels(intTemp).strLabel(lngLine) = " "
'                Next
'                intTemp = intTemp + 1
'                prgProgress.Value = intTemp
'            Next intCnt
'        Next intResult
'    Else
'        ' Add every label to the array
'        For intResult = 1 To intCopies
'            For Each itmItem In lsvLabels.ListItems
'                For intCnt = 1 To Val(itmItem.SubItems(1))
'                        On Error Resume Next
'                        strTemp() = Split(itmItem.Text, "|", 3)
'                        With labLabels(intTemp)
'                            For lngLine = 0 To 2
'                                .strLabel(lngLine) = Trim$(strTemp(lngLine))
'
'                                If Len(.strLabel(lngLine)) = 0 Then
'                                    .strLabel(lngLine) = " "
'                                Else
'                                    If lngLine = 1 Then
'                                        blnMultLine2 = True
'                                    ElseIf lngLine = 2 Then
'                                        blnMultLine3 = True
'                                    End If
'                                End If
'                            Next
'                        End With
'                        On Error GoTo 0
'
'                        '200507 : NEW SIZING ROUTINE
'                        If Printer.TextWidth(labLabels(intTemp).strLabel(0)) > sngWidth * 1440 Then
'                            labLabels(intTemp).lngSize = SizeToText(sngWidth - 1.5 / 16, sngHeight - 1 / 16, ActualTextHeight, labLabels(intTemp).strLabel(0))
'                            PrintTemp "C"
'                            PrintTemp "labLabels(intTemp).lngSize = " & labLabels(intTemp).lngSize
'                        End If
'
'                        'CHG 200507 : PREVIOUS ROUTINE
'                        ' Check to see if the width of the text will fit horizontally
'                        ' On the printed label, so that no labels overlap the text
''                        For lngLine = 0 To 2
''                            If Printer.TextWidth(labLabels(intTemp).strLabel(lngLine)) > sngWidth * 1440 And intAutoSize = 0 Then
''                                ' Label width isn't wide enough to fit the text onto the label
''                                ' Notify the user about which label has too long of text
''                                If blnMultLine3 Then
''                                    MsgBox "Label " & Chr(34) & Join(labLabels(intTemp).strLabel, "|") & Chr(34) & " is wider than the label!", vbOKOnly + vbInformation, "Label Is Too Wide!"
''                                ElseIf blnMultLine2 Then
''                                    MsgBox "Label " & Chr(34) & labLabels(intTemp).strLabel(0) & "|" & labLabels(intTemp).strLabel(1) & Chr(34) & " is wider than the label!", vbOKOnly + vbInformation, "Label Is Too Wide!"
''                                Else
''                                    MsgBox "Label " & Chr(34) & labLabels(intTemp).strLabel(0) & Chr(34) & " is wider than the label!", vbOKOnly + vbInformation, "Label Is Too Wide!"
''                                End If
''                                prgProgress.Value = 0
''                                Exit Sub
''                            End If
''                        Next lngLine
'
'                    ' Update the progress bar to show the current status of adding the labels to the array
'                    intTemp = intTemp + 1
'                    prgProgress.Value = intTemp
'                Next intCnt
'
'            'Next Label in Listview
'            Next itmItem
'
'            ' Fill the rest of the row with blank labels
'            For intCnt = 1 To intBlank
'                For lngLine = 0 To 2
'                    labLabels(intTemp).strLabel(lngLine) = " "
'                Next
'                intTemp = intTemp + 1
'                prgProgress.Value = intTemp
'            Next intCnt
'        Next intResult
'    End If
'
'    ' Set the largest font size to zero to start with
'    lngLargest = 0
'
''CHG 200507 : USING NEW SIZING ROUTINE ABOVE
''==============
'    ' Check to see if we are automatically sizing the text
'    ' to fit the label
''    If intAutoSize = 1 Then
''        ' Loop through each label in the array
''        For intCnt = LBound(labLabels) To UBound(labLabels)
''            ' Set blnTemp to false so we don't exit the
''            ' loop immediately
''            blnTemp = False
''
''            ' Start the printer's font size at 1
''            Printer.FontSize = 1
''
''            Do
''                ' Check to see if the text still fits
''                ' within the dimensions of the label
''                If blnMultLine3 Then
''                    If (Printer.TextHeight(Join(labLabels(intCnt).strLabel, " ")) * 3) < (sngHeight * 1440) _
''                    And (Printer.TextWidth(labLabels(intCnt).strLabel(0)) < sngWidth * 1440) _
''                    And (Printer.TextWidth(labLabels(intCnt).strLabel(1)) < sngWidth * 1440) _
''                    And (Printer.TextWidth(labLabels(intCnt).strLabel(2)) < sngWidth * 1440) Then
''                        ' If the text still fits, then increase
''                        ' the printer's font size by 1
''                        Printer.FontSize = Printer.FontSize + 1
''                    Else
''                        ' The text doesn't fit any longer, so
''                        ' set blnTemp to True so that we exit
''                        ' the loop, and then set the printer's
''                        ' font size to one less than what it is
''                        blnTemp = True
''                        labLabels(intCnt).lngSize = Printer.FontSize - 1
''                    End If
''                ElseIf blnMultLine2 Then
''                    If (Printer.TextHeight(Join(labLabels(intCnt).strLabel, " ")) * 2) < (sngHeight * 1440) _
''                    And (Printer.TextWidth(labLabels(intCnt).strLabel(0)) < sngWidth * 1440) _
''                    And (Printer.TextWidth(labLabels(intCnt).strLabel(1)) < sngWidth * 1440) Then
''                        ' If the text still fits, then increase
''                        ' the printer's font size by 1
''                        Printer.FontSize = Printer.FontSize + 1
''                    Else
''                        ' The text doesn't fit any longer, so
''                        ' set blnTemp to True so that we exit
''                        ' the loop, and then set the printer's
''                        ' font size to one less than what it is
''                        blnTemp = True
''                        labLabels(intCnt).lngSize = Printer.FontSize - 1
''                    End If
''                Else
''                    If (Printer.TextHeight(labLabels(intCnt).strLabel(0)) * intLines) < (sngHeight * 1440) _
''                    And (Printer.TextWidth(labLabels(intCnt).strLabel(0)) < sngWidth * 1440) Then
''                        ' If the text still fits, then increase
''                        ' the printer's font size by 1
''                        Printer.FontSize = Printer.FontSize + 1
''                    Else
''                        ' The text doesn't fit any longer, so
''                        ' set blnTemp to True so that we exit
''                        ' the loop, and then set the printer's
''                        ' font size to one less than what it is
''                        blnTemp = True
''                        labLabels(intCnt).lngSize = Printer.FontSize - 1
''                    End If
''                End If
''
''                ' If the label is blank, then set blnTemp to
''                ' true so that we exit the loop, and set the
''                ' printer's font size to 8
''                If Len(Trim$(labLabels(intCnt).strLabel(0))) = 0 Then
''                    blnTemp = True
''                    labLabels(intCnt).lngSize = 8
''                    Printer.FontSize = 8
''                End If
''            Loop Until blnTemp = True
''
''            ' Check the largest font used, and if the label's
''            ' font is larger, set the largest font size used
''            ' to the last label's font size
''            If lngLargest < labLabels(intCnt).lngSize Then
''                lngLargest = labLabels(intCnt).lngSize
''            End If
''        Next 'intCnt
''    End If
''=====================
'
'    ' Reset the progress bar value
'    prgProgress.Value = 0
'
'    ' Show the printer status dialog box
'    frmPrint.Show 0, frmMain
'
'    ' Update the printer dialog status to 0 labels printed
'    frmPrint.lblPrint.Caption = "Printing 0/" & Trim$(Str$(intTQty - (intBlank * intCopies)))
'    DoEvents
'
'    intTemp = 0
'
'    Do
'        ' Set the current printer vertical position
'        Printer.CurrentY = sngPrintTopY * 1440
'        PrintTemp "D"
'        DebugPrinter
'
'
''CHG 200507
''==========
''        If blnMultLine3 Then
''            ' Print each label the number of lines it needs to be printed
''            For intLPL = 0 To 2
''                intLPR = 0
''
''                Do
''                    ' Set the current horizontal position of the printer
''                    Printer.CurrentX = intLPR * sngSpacingRL * 1440 + sngPrintLeftX * 1440
''
''                    ' If we are automatically sizing text, then
''                    ' set the font size to the size of the current label
''                    If intAutoSize = 1 And labLabels(intTemp + intLPR).lngSize > 0 Then
''                        Printer.FontSize = labLabels(intTemp + intLPR).lngSize
''                    End If
''
''                    ' Print the next label in the list
''                    Printer.Print labLabels(intTemp + intLPR).strLabel(intLPL);
''                    intLPR = intLPR + 1
''                Loop Until (intLPR = intLabelsPerRow) Or ((intTemp + intLPR) = intTQty)
''
''                ' Print a line feed and carriage return to move
''                ' the printer to the next line for the next line of text
''                If intAutoSize = 1 Then
''                    Printer.FontSize = lngLargest
''                End If
''                Printer.Print
''            Next intLPL
''        ElseIf blnMultLine2 Then
''            ' Print each label the number of lines it needs to be printed
''            For intLPL = 0 To 1
''                intLPR = 0
''
''                Do
''                    ' Set the current horizontal position of the printer
''                    Printer.CurrentX = intLPR * sngSpacingRL * 1440 + sngPrintLeftX * 1440
''
''                    ' If we are automatically sizing text, then
''                    ' set the font size to the size of the current label
''                    If intAutoSize = 1 And labLabels(intTemp + intLPR).lngSize > 0 Then
''                        Printer.FontSize = labLabels(intTemp + intLPR).lngSize
''                    End If
''
''                    ' Print the next label in the list
''                    Printer.Print labLabels(intTemp + intLPR).strLabel(intLPL);
''                    intLPR = intLPR + 1
''                Loop Until (intLPR = intLabelsPerRow) Or ((intTemp + intLPR) = intTQty)
''
''                ' Print a line feed and carriage return to move
''                ' the printer to the next line for the next line of text
''                If intAutoSize = 1 Then
''                    Printer.FontSize = lngLargest
''                End If
''                Printer.Print
''            Next intLPL
''        Else
''==========
'
'            Debug.Assert Not (blnMultLine2 Or blnMultLine3)
'            ' Print each label the number of lines it needs to be printed
'            For intLPL = 1 To intLines
'                intLPR = 0
'
'                lngLargest = 0
'
'                Do
'
'                    If lngLargest < labLabels(intTemp + intLPR).lngSize Then
'                        lngLargest = labLabels(intTemp + intLPR).lngSize
'                    End If
'
'                    ' Set the current horizontal position of the printer
'                    Printer.CurrentX = intLPR * sngSpacingRL * 1440 + sngPrintLeftX * 1440
'                    If Not labLabels(intTemp + intLPR).lngSize = 0 Then
'                        Printer.FontSize = labLabels(intTemp + intLPR).lngSize
'                    Else
'                        'assign minimum size
'                        Printer.FontSize = 8
'                    End If
'
''                    'CHG 200507 : ALWAYS ADJUSTING THE SIZE
''                    ' If we are automatically sizing text, then
''                    ' set the font size to the size of the current label
''                    If intAutoSize = 1 And labLabels(intTemp + intLPR).lngSize > 0 Then
''                        Printer.FontSize = labLabels(intTemp + intLPR).lngSize
''                    End If
'
'                    PrintTemp "E"
'                    DebugPrinter
'
'                    ' Print the next label in the list
'                    Printer.Print labLabels(intTemp + intLPR).strLabel(0);
'
'                    PrintTemp "F"
'                    DebugPrinter
'
'                    intLPR = intLPR + 1
'                Loop Until (intLPR = intLabelsPerRow) Or ((intTemp + intLPR) = intTQty)
'
'                ' Print a line feed and carriage return to move
'                ' the printer to the next line for the next line of text
'                'If intAutoSize = 1 Then
'                    If lngLargest = 0 Then
'                        lngLargest = 8
'                    End If
'                                        Debug.Print lngLargest
'
'                    Printer.FontSize = lngLargest
'                'End If
'                Printer.Print
'
'                PrintTemp "G"
'                PrintTemp "intTemp " & intTemp
'                PrintTemp "intLPR " & intLPR
'                PrintTemp "lngLargest " & lngLargest
'                DebugPrinter
'
'            Next intLPL
''        End If
''====
'        'Update the printer status dialog box
'        intTemp = intTemp + intLabelsPerRow
'        ' Update the printer status dialog box
'        If intTemp > (intTQty - intBlank) Then
'            frmPrint.lblPrint.Caption = "Printing " & Trim$(Str$(intTQty - (intBlank * intCopies))) & " / " & Trim$(Str$(intTQty - (intBlank * intCopies)))
'        Else
'            frmPrint.lblPrint.Caption = "Printing " & Trim$(Str$(intTemp)) & " / " & Trim$(Str$(intTQty - (intBlank * intCopies)))
'        End If
'
'        ' Increment the counter to go to the next row of labels
'
'        ' If we printed a row of labels, then send a new page
'        ' code to the printer.
'
'        'Printer.NewPage
'        intCnt = intCnt + 1
'        DoEvents
'
'    Loop Until intTemp >= intTQty
'
'    ' Done printing, finish sending the data to the printer
'    Printer.KillDoc
'    Printer.EndDoc
'
'    ' Update the printer display dialog to show that we're done
'    frmPrint.lblPrint.Caption = "Done!"
'    DoEvents
'
'    ' Wait for 1 second before closing the printer status dialog box
'    Sleep 1000
'    prgProgress.Value = 0
'
'    ' Close the printer status dialog box
'    Unload frmPrint
'End Sub
'-------------------------------------------------
'   END   frmMain code
'-------------------------------------------------

'===============================
'   END OBSOLETE CODE
'===============================



