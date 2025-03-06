Attribute VB_Name = "Globals"
Option Explicit


'====================
'Label Specifications
'   as written in 'Shop Label Formats.dat' (200510)
'   previously formats.dat
Public Const LABEL_FORMATS As String = "Shop Label Formats.dat"
Public strLabelFormat As String    'Format name
Public sngTopMargin As Single    'Top Margin
Public sngLeftMargin As Single    'Left Margin
Public sngWidth As Single    'Print width
Public sngHeight As Single    'Print height
Public sngSpacingTB As Single    'Vertical spacing
Public sngSpacingRL As Single    'Horizontal spacing
Public intLines As Long    '# Lines repeated on label
Public intLabelsPerRow As Long    'Labels per row
Public intOptical As Long    'Optical labels?  1=Yes
Public intAutoSize As Long    'Autosize text? 1=Yes
'====================

Public intDefaultQty As Long
Public blnSaved As Boolean
Public blnLockFormat As Boolean
Public blnPrintAll As Boolean
Public intCopies As Long

' Declare function for fixing keyboard lock up
' PostMessage is also used by the main form for mouse control
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_KEYDOWN As Long = &H100
Public Const VK_TAB As Long = &H9
Public Enum FeatureModeConstants
    TERMINAL_STRIPS
    WIRE_LABELS
    STOCORD_LABELS
End Enum
Public FeatureMode  As FeatureModeConstants

Public Sub Main()

'200510 N.F.
'   Consolidated source code for terminal labels and wire labels
'   while creating new StoCord features.
#If DebugMode >= 0 And DebugMode <= 1 Then
    #If Feature = 1 Then
        FeatureMode = WIRE_LABELS
    #ElseIf Feature = 2 Then
        FeatureMode = TERMINAL_STRIPS
    #ElseIf Feature = 3 Then
        FeatureMode = STOCORD_LABELS
    #Else
        Debug.Assert False
        'New Feature?
        'Default FeatureMode to Wire Labels
        FeatureMode = WIRE_LABELS
    #End If
    frmMain.Show
#Else
    'frmTests.Show
#End If

End Sub

'SNIPPET
Public Function IsFormLoaded(ByVal FormName As String) As Boolean

    Dim RetVal As Boolean
    Dim f As Form
    RetVal = False
    For Each f In Forms
        If LCase(f.Name) = LCase(FormName) Then
            RetVal = True
            Exit For
        End If
    Next
    IsFormLoaded = RetVal
    
End Function

Public Function GetFormCaption(ByVal FormName As String) As String

    Dim sTemp As String
    If FeatureMode = WIRE_LABELS Then
        sTemp = "Wire Labels"
    ElseIf FeatureMode = TERMINAL_STRIPS Then
        sTemp = "Terminal Labels"
    ElseIf FeatureMode = STOCORD_LABELS Then
        sTemp = "StoCord Labels"
    Else
        'New Feature to caption?
        Debug.Assert False
    End If
    
    Select Case LCase(FormName)
        Case "frmbrowser"
            '"Kasa Wire Labels Help"
            sTemp = "Kasa " & sTemp & " Help"
        Case "frmabout"
            sTemp = "About Kasa " & sTemp
        Case "frmmain"
            sTemp = "Kasa " & sTemp
    End Select
    
    GetFormCaption = sTemp
    
End Function

'200510: N.F.
'   Moved from frmFormat.
Public Function FormatLabelDimension(ByRef s As Variant) As Variant
    FormatLabelDimension = Format(Val(s), "0.000000")
End Function

'-------------------------------------------------
' This subroutine opens the label formats file and finds the label
' format that the user selected and loads the values for that label
' format into the respective text boxes.
'-------------------------------------------------
'200510: N.F.
'   Adapted from lstFormats_click routine.
'-------------------------------------------------
Public Function SelectNewFormat(ByVal FormatName As String, _
 ByRef arrValues() As String) As Boolean

    On Error GoTo ErrorHandler
    
    Dim strTemp As String
    Dim intFile As Long
    Dim blnFound As Boolean
    
    SelectNewFormat = False
    
    ' Find an available file location in memory
    intFile = FreeFile
    
    ' Open the label formats file for input
    Open App.Path & "\" & LABEL_FORMATS For Input As #intFile
    
    ' Set the label found indicator to false
    blnFound = False
    
    Do
        ' Read in a line of the file
        Input #intFile, strTemp
        ' If the label format we found matches the one the user
        ' clicked, then set the label found indicator to true.
        If Trim$(strTemp) = "~" & FormatName Then
            blnFound = True
        End If
    Loop Until EOF(intFile) Or blnFound         ' Keep looking until we reach the end of the file or we find the label
    
    ' If we didn't find the label then let the user know, and close the file.
    If blnFound = False Then
        MsgBox "Item Not Found!", , "Error"
        Close #intFile
        Exit Function
    End If
    
    ' We must have found the file, so go ahead and read
    ' in all the values for that label format.
    Input #intFile, arrValues(0)
    Input #intFile, arrValues(1)
    Input #intFile, arrValues(2)
    Input #intFile, arrValues(3)
    Input #intFile, arrValues(4)
    Input #intFile, arrValues(5)
    Input #intFile, arrValues(6)
    Input #intFile, arrValues(7)
    Input #intFile, arrValues(8)
    Input #intFile, arrValues(9)
    
    ' Close the file we had open.
    Close #intFile
       
    SelectNewFormat = True
       
    Exit Function
ErrorHandler:
    If Err.Number = 53 Then
        MsgBox "Expected label specification file : " & LABEL_FORMATS & vbCr _
         & "In path : " & vbCr & App.Path, vbOKOnly, GetFormCaption("frmmain")
        End
    Else
        MsgBox "Error " & Err.Number & " - " & Err.Description, , "Error"
    End If
End Function


'CHG: 200510 MOVED FROM frmMain
'ASSUMPTIONS:
'   Printer drivers installed and Printer DeviceName exists containing
'       one of the following strings:
'       TDP42H, PTR3, M84 Pro, M-8400RV
'===========================================
Public Function CheckProperPrinter() As Boolean
            
    Dim p As Printer
    Dim RetVal As Boolean
    RetVal = True
    
    'Check if valid printer driver is found
    If InStr(LCase(Printer.DeviceName), LCase("TDP42H")) = 0 And _
     InStr(LCase(Printer.DeviceName), LCase("M84 Pro")) = 0 And _
     InStr(LCase(Printer.DeviceName), LCase("M-8400RV")) = 0 And _
     InStr(LCase(Printer.DeviceName), LCase("ptr3")) = 0 Then
        'Default printer is not the Panduit 3... select the first one that is.
        For Each p In Printers
            If InStr(LCase(p.DeviceName), LCase("TDP42H")) > 0 Or _
             InStr(LCase(p.DeviceName), LCase("M84 Pro")) > 0 Or _
             InStr(LCase(p.DeviceName), LCase("M-8400RV")) > 0 Or _
             InStr(LCase(p.DeviceName), LCase("ptr3")) > 0 Then
                Set Printer = p
            End If
        Next
        'Check one more time if printer is now valid
        If InStr(LCase(Printer.DeviceName), LCase("TDP42H")) = 0 And _
         InStr(LCase(Printer.DeviceName), LCase("M84 Pro")) = 0 And _
         InStr(LCase(Printer.DeviceName), LCase("M-8400RV")) = 0 And _
         InStr(LCase(Printer.DeviceName), LCase("ptr3")) = 0 Then
            MsgBox "Printer Driver is not Installed:  Expected Printer M84 Pro or M-8400RV or TDP42H or PTR3."
            RetVal = False
        End If
    End If

    If RetVal Then
        RetVal = SetTerminalLabelsForm(frmMain.hwnd)
    End If
    
    CheckProperPrinter = RetVal
    
End Function

Public Function SetFieldTagsForm(ByVal hwnd As Long) As Boolean

    Dim RetVal As Boolean
    RetVal = True

    'SAMPD
    'Printer.Orientation = vbPRORLandscape
    If Not FormExists(Printer.DeviceName, "Kasa Field Tag Labels") Then
        'SAMPA, B, C
        If SelectForm("Kasa Field Tag Labels", hwnd, 4, 6) > 0 Then
        'SAMPD
        'If SelectForm("Kasa Field Tag Labels", hwnd, 6, 4) > 0 Then
'            MsgBox "Form " & Chr$(34) & "Kasa Field Tag Labels" & Chr$(34) & " was created!" & vbCrLf & vbCrLf & _
'                   "Please set the " & Printer.DeviceName & " Default Form to this form." & vbCrLf & _
'                   "Kasa Wire Label software will now exit.", vbOKOnly + vbInformation, "Kasa Wire Labels"
'           End
        Else
            MsgBox "Field Tag printing will not work correctly!" & vbCr & "Kasa Field Tag Labels form could not be set or created.", vbOKOnly + vbInformation, "Kasa Field Tag Labels"
            RetVal = False
        End If
    Else
        'SAMPA, B, C
        RetVal = (SelectForm("Kasa Field Tag Labels", hwnd, 4, 6) > 0)
        'SAMPD
        'RetVal = (SelectForm("Kasa Field Tag Labels", hwnd, 6, 4) > 0)
    End If
    
    SetFieldTagsForm = RetVal

End Function

'=================================================
'   This function will set the form to be the
'   maximum size available for the Sato printers.
'   4 x 49
'   Wire Labels will print on this form, because
'   they are optically sensed.
'   Teriminal Labels will print using this form,
'   because the terminal PrintStrips routine
'   assumes a page length of 49"
'   Field Tags have their own form, set elsewhere.
'=================================================
Public Function SetTerminalLabelsForm(ByVal hwnd As Long) As Boolean

    Dim Answer As VbMsgBoxResult
    Dim RetVal As Boolean
    RetVal = True
    
    'chg 200510 N.F.
    '   Changed the form name to Terminal Labels
    '   NOTE:  Only terminal labels form need to be selected, because this
    '       represents the maximum print surface.
    If Not FormExists(Printer.DeviceName, "Kasa Terminal Labels2") Then
        If SelectForm("Kasa Terminal Labels2", hwnd, 4, 49) > 0 Then
            'chg 200510 N.F.
            '   No need to close the program when selecting new form

'            MsgBox "Form " & Chr$(34) & "Kasa Terminal Labels2" & Chr$(34) & " was created!" & vbCrLf & vbCrLf & _
'                   "Please set the Panduit " & Printer.DeviceName & " Default Form to this form." & vbCrLf & _
'                   "Kasa Wire Label software will now exit.", vbOKOnly + vbInformation, "Kasa Wire Labels"
'            End

            RetVal = True
        Else
            'CHG 200510 N.F.
            'Try to select previous value
            If SelectForm("Kasa Wire Labels TDP42H", hwnd, 4, 49) > 0 Then
                RetVal = True
            Else
                'CHG 200510 N.F.
                '   Special message if form creation fails on networked printer
                '   Permit opportunity to continue anyway
                If Left(Printer.DeviceName, 2) = "\\" Then
                    Answer = MsgBox("Could not create form: Kasa Terminal Labels2 (4x49)" & vbCr & "The Printer is " & Printer.DeviceName & vbCr & vbCr & "Do you want to continue?", vbYesNo)
                    RetVal = (Answer = vbYes)
                Else
                    MsgBox "Strip label printing will not work correctly!" & vbCr & "Kasa Terminal Labels2 form could not be set or created.", vbOKOnly + vbInformation, "Kasa Wire Labels"
                    RetVal = False
                End If
            End If
        End If
    Else
        RetVal = (SelectForm("Kasa Terminal Labels2", hwnd, 4, 49) > 0)
    End If


    SetTerminalLabelsForm = RetVal

End Function
'Purpose
'   Displays Printer Dialog Box to prepare and print labels.
'
'Return Values:
'   CheckIsOKtoPrint
'       True if user selected OK
'       False if user selected Cancel
'   PrintAllLabels
'       True if user selected 'All' as print range (or Cancelled)
'       False if user selected 'Selection' as print range and selected OK
Public Function CheckIsOKtoPrint(ByRef ParentFormHandle As Long, _
    ByVal SelectedAllLabels As Boolean, _
    ByRef PrintAllLabels As Boolean) As Boolean
 
    If Not ParentFormHandle > 0 Then
        Debug.Assert False
        'Need to pass Valid form handle : e.g. Me.Hwnd
        Exit Function
    End If
    
    Dim blnTemp As Boolean
    Dim RetVal As Boolean
    
    blnTemp = False
    
    If SelectedAllLabels Then
        RetVal = SelectPrinter(ParentFormHandle, True, intCopies, True)
    Else
        RetVal = SelectPrinter(ParentFormHandle, True, intCopies, False, blnTemp)
    End If
    
    PrintAllLabels = Not blnTemp
    CheckIsOKtoPrint = RetVal
 
End Function


Public Function ConfirmLabelsToPrint(ByRef TotalQty As Long, _
 ByVal PrintAll As Boolean, _
 ByRef oListView As ListView) As Boolean

    Dim itmItem As ListItem
    Dim Answer As VbMsgBoxResult
    Dim intTemp As Long
    Dim msg As String
    
    TotalQty = 0
    
    ' If we are only printing the selected items, the only count
    ' the quantities of the selected labels.
    If Not PrintAll Then
        ' Only count the selected labels
        
        For Each itmItem In oListView.ListItems
            If itmItem.Selected = True Then
                TotalQty = TotalQty + Val(itmItem.SubItems(1))
                'itmItem.Selected = False
            End If
            intTemp = intTemp + Val(itmItem.SubItems(1))
        Next itmItem
        
        msg = Trim$(Str$(TotalQty)) & " / " _
                & Trim$(Str$(intTemp)) & " labels to print in " _
                & Chr(34) & strLabelFormat & Chr(34) & " format." _
                & vbCrLf & "This will print " & intCopies
        
        If intCopies > 1 Then
            msg = msg & " copies."
        Else
            msg = msg & " copy."
        End If
    
    Else
        ' Count all the labels in the list
        
        For Each itmItem In oListView.ListItems
            TotalQty = TotalQty + Val(itmItem.SubItems(1))
        Next itmItem
        
        msg = Trim$(Str$(TotalQty)) & " labels to print in " _
                & Chr(34) & strLabelFormat & Chr(34) & " format." _
                & vbCrLf & "This will print " & intCopies
        
        If intCopies > 1 Then
            msg = msg & " copies."
        Else
            msg = msg & " copy."
        End If
        
    End If

    If TotalQty > 0 Then
        msg = msg & vbCrLf & "Using " & Printer.DeviceName & " on port " _
            & Printer.Port & vbCrLf & vbCrLf _
            & "Align label sheet in printer and click Ok."
    
        ' Ask user if they really want to print these labels
        Answer = MsgBox(msg, vbOKCancel + vbInformation, "Print")
        If Not (Answer = vbOK) Then
            TotalQty = 0
        End If
    End If
    ConfirmLabelsToPrint = Not (TotalQty = 0)
    
End Function


'X : Must be >= 1
'n : Must be >= 1
Public Sub RoundUpToNearestX(ByVal X As Long, ByRef n As Long)

    Dim lRemainder As Long
    
    If X < 1 Or n < 1 Then Err.Raise 5
    
    If X = 1 Then
        Debug.Print n
        Exit Sub
    End If
    
    ' Add room for blank labels to fill out each row
    If (n Mod X) > 0 Then
        n = n + X - (n Mod X)
        
    End If

    Debug.Print n
    
End Sub

Public Sub InitPrinter()

    ' Set the current printer position
'    Printer.CurrentX = sngPrintLeftX * 1440
'    Printer.CurrentY = sngPrintTopY * 1440
    
    ' Set the printer's font settings
    With Printer
        .FontBold = frmMain.cdlDialog.FontBold
        .FontItalic = frmMain.cdlDialog.FontItalic
        .FontName = frmMain.cdlDialog.FontName
        .FontSize = frmMain.cdlDialog.FontSize
        .FontStrikethru = frmMain.cdlDialog.FontStrikethru
        .FontUnderline = frmMain.cdlDialog.FontUnderline
    End With
    
'    ' Set the printed page's height and width
'    Printer.Height = sngSpacingTB * 1440
'    Printer.Width = intLabelsPerRow * sngSpacingRL * 1440 + sngPrintLeftX * 1440

    
'    PrintTemp "A"
'    Call DebugPrinter

End Sub

Public Function AutosizeTallText() As Boolean

    Dim RetVal As Boolean
    
    
    ' Check the label height compared to the height of the
    ' printed output to make sure the printed output will fit
    ' on the label format that is currently selected
    If (Printer.TextHeight("gjqy#^") * intLines) > (sngHeight * 1440) And intAutoSize = 0 Then
        ' Label text will not fit vertically, so notify the user.
        MsgBox "Using the current label format and the current font, the text is too tall for the label." & vbCrLf & "You can try resizing the font, or check 'Automatically Size to Fit' for this label format.", vbOKOnly + vbInformation, "Text Too Tall!"
        RetVal = False
    Else
        RetVal = True
    End If
    AutosizeTallText = RetVal
    
End Function




Public Sub SelectText()
' This subroutine selects the text in the text box it was called from.

    If TypeName(Screen.ActiveControl) = "TextBox" Then
        Screen.ActiveControl.SelStart = 0
        Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
    End If
End Sub

Public Function SplitText(ByVal strLabel As String, ByVal intMaxLength As Integer) As String
' This subroutine will split any text string to a maximum
' width per line of the amount specified by intMaxLength

    Dim intBreak As Integer
    
    strLabel = Trim(strLabel)
    intBreak = Len(strLabel) / 2
    
    'chg 200504
    'Request was for splitting the text evenly
    If Len(strLabel) >= intMaxLength - 1 Then
        SplitText = Left(strLabel, intBreak) & vbCrLf _
         & Right(strLabel, Len(strLabel) - intBreak)
    Else
        SplitText = strLabel
    End If
    
    
    Exit Function
    
    Dim strFinal As String
    Dim intCutMax As Integer
    Dim intCut As Integer
    Dim intPos As Integer
    Dim blnChange As Boolean
    Dim intLast As Integer
    Dim intPrimaryCut As Integer

    ' Trim any spaces off the text
    strLabel = Trim$(strLabel)
    
    ' Set the final split up label to nothing
    strFinal = ""
    
    ' Make the maximum lenght to search for a splitter
    ' character to one extra character
    'intMaxLength = intMaxLength + 1
    
    Do
        intPos = 1
        intCutMax = 0
        intPrimaryCut = 0
        Do
            ' Make sure blnChange is set back to false
            ' so that our current changes are tracked
            blnChange = False
            
            ' Find the first position starting from intPos
            ' of a space
            intCut = InStr(intPos, strLabel, " ")
            ' See if the character that was found is
            ' further from the start of the string
            ' and within the limits of the max length
            If (intCutMax < intCut) And intCut <= intMaxLength And Len(strLabel) > intMaxLength Then
                ' This character is within the limits
                ' so change the position we will split
                ' the text, and make blnChange true
                ' so that we can make sure that this
                ' is the furthest position possible
                intCutMax = intCut
                intPrimaryCut = intCut
                blnChange = True
            End If
            
            ' Find the first position starting from intPos
            ' of a period
            intCut = InStr(intPos, strLabel, ".")
            ' See if the character that was found is
            ' further from the start of the string
            ' and within the limits of the max length
            If (intCutMax < intCut) And intCut <= intMaxLength And Len(strLabel) > intMaxLength Then
                ' This character is within the limits
                ' so change the position we will split
                ' the text, and make blnChange true
                ' so that we can make sure that this
                ' is the furthest position possible
                intCutMax = intCut
                blnChange = True
            End If
            
            ' Find the first position starting from intPos
            ' of a slash
            intCut = InStr(intPos, strLabel, "/")
            ' See if the character that was found is
            ' further from the start of the string
            ' and within the limits of the max length
            If (intCutMax < intCut) And intCut <= intMaxLength And Len(strLabel) > intMaxLength Then
                ' This character is within the limits
                ' so change the position we will split
                ' the text, and make blnChange true
                ' so that we can make sure that this
                ' is the furthest position possible
                intCutMax = intCut
                blnChange = True
            End If
            
            ' Find the first position starting from intPos
            ' of a colon
            intCut = InStr(intPos, strLabel, ":")
            ' See if the character that was found is
            ' further from the start of the string
            ' and within the limits of the max length
            If (intCutMax < intCut) And intCut <= intMaxLength And Len(strLabel) > intMaxLength Then
                ' This character is within the limits
                ' so change the position we will split
                ' the text, and make blnChange true
                ' so that we can make sure that this
                ' is the furthest position possible
                intCutMax = intCut
                blnChange = True
            End If
            
            ' Find the first position starting from intPos
            ' of a backslash
            intCut = InStr(intPos, strLabel, "\")
            ' See if the character that was found is
            ' further from the start of the string
            ' and within the limits of the max length
            If (intCutMax < intCut) And intCut <= intMaxLength And Len(strLabel) > intMaxLength Then
                ' This character is within the limits
                ' so change the position we will split
                ' the text, and make blnChange true
                ' so that we can make sure that this
                ' is the furthest position possible
                intCutMax = intCut
                blnChange = True
            End If
            
            ' Find the first position starting from intPos
            ' of a comma
            intCut = InStr(intPos, strLabel, ",")
            ' See if the character that was found is
            ' further from the start of the string
            ' and within the limits of the max length
            If (intCutMax < intCut) And intCut <= intMaxLength And Len(strLabel) > intMaxLength Then
                ' This character is within the limits
                ' so change the position we will split
                ' the text, and make blnChange true
                ' so that we can make sure that this
                ' is the furthest position possible
                intCutMax = intCut
                blnChange = True
            End If
            
            ' Find the first position starting from intPos
            ' of a dash
            intCut = InStr(intPos, strLabel, "-")
            ' See if the character that was found is
            ' further from the start of the string
            ' and within the limits of the max length
            If (intCutMax < intCut) And intCut <= intMaxLength And Len(strLabel) > intMaxLength Then
                ' This character is within the limits
                ' so change the position we will split
                ' the text, and make blnChange true
                ' so that we can make sure that this
                ' is the furthest position possible
                intCutMax = intCut
                blnChange = True
            End If
            
            ' Find the first position starting from intPos
            ' of an underscore
            intCut = InStr(intPos, strLabel, "_")
            ' See if the character that was found is
            ' further from the start of the string
            ' and within the limits of the max length
            If (intCutMax < intCut) And intCut <= intMaxLength And Len(strLabel) > intMaxLength Then
                ' This character is within the limits
                ' so change the position we will split
                ' the text, and make blnChange true
                ' so that we can make sure that this
                ' is the furthest position possible
                intCutMax = intCut
                blnChange = True
            End If
            
            ' Increase the starting search position
            ' by one character from the last found
            ' split character
            If intPrimaryCut > 0 And Len(strLabel) - intPrimaryCut <= intMaxLength Then
                intCutMax = intPrimaryCut
            End If
            intPos = intCutMax + 1
        ' Keep looping until there are no changes or the
        ' length of the remaining text is within the limits
        ' of the maximum length for a line of text
        Loop Until blnChange = False Or Len(strLabel) <= intMaxLength Or intPrimaryCut > 0
        
        If intCutMax = 0 And Len(strFinal) = 0 Then
        ' If the text is not trimmed and it is the first
        ' text of the loop then make it the final text,
        ' and set the length of it to the Max Length
            If Len(strLabel) > intMaxLength Then
                strFinal = Left$(strLabel, 0.5 * Len(strLabel))
                intLast = 0.5 * Len(strLabel)
                'strFinal = Left$(strLabel, 0.5 * Len(strLabel)) & vbCrLf & Right(strLabel, Len(strLabel) - 0.5 * Len(strLabel))
            Else
                strFinal = Left$(strLabel, intMaxLength)
                intLast = intMaxLength
            End If
            
        ElseIf intCutMax = 0 And Len(strFinal) > 0 Then
        ' If the text is not trimmed, and there is already
        ' text in strFinal then add to the string a return
        ' and the new text, and set the length of the new text
        ' to the Max Length
            strFinal = strFinal & vbCrLf & Left$(strLabel, intMaxLength)
            intLast = intMaxLength
        ElseIf intCutMax > 0 And Len(strFinal) = 0 Then
        ' If the text was trimmed, and it is the first text
        ' of the loop, then make it the final text, and set
        ' the last length added equal to the length of the text
            strFinal = Left$(strLabel, intCutMax)
            intLast = intCutMax
        Else
        ' If the text was trimmed, and there is already text
        ' in strFinal then add to the string a return and the
        ' new text, and set the last length added equal to
        ' the length of the text that was added
            strFinal = strFinal & vbCrLf & Left$(strLabel, intCutMax)
            intLast = intCutMax
        End If
        
        ' Check to see if the remaining text is within
        ' the limits of the maximum length
        If intMaxLength < Len(strLabel) Then
            ' Trim off the text that was just added to
            ' the final text string
            strLabel = Mid$(strLabel, intLast + 1)
        Else
            ' Clear the text of the label, because all of
            ' it was added to the final text string
            strLabel = ""
        End If
    Loop Until Len(strLabel) <= intMaxLength
    
    ' Check any final text that is left to see if it and the
    ' length of the last text added is in total greater than
    ' the maximum lenght allowable, if it is, then add a return
    ' to the final text
    If Len(strLabel) + intLast > intMaxLength Then
        strFinal = strFinal & vbCrLf
    End If
    
    ' Return the combination of the final text and the
    ' remaining text in strLabel
    SplitText = Trim$(strFinal & strLabel)
End Function
