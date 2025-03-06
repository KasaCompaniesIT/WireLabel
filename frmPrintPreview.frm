VERSION 5.00
Begin VB.Form frmPrintPreview 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Label Print Preview"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox picPrinter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2505
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intTQty As Long

'Currently Unused
Dim SelectedPage As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If intTQty = 0 Then Call Command1_Click
End Sub

Private Sub Form_Load()
' This subroutine will preview the label list as strip labels
    Dim NextY As Single
    Dim ActualTextHeight As Single
    
    Dim sFirstLine As String
    Dim sSecondLine As String
    Dim PageHeight As Single
    Dim CurLabel As Long
    Dim PgLastLabel As Long
    Dim Pages As Integer
    Dim CurPage As Integer
    Dim strLabel() As String
    Dim lngSize() As Long
    Dim sngPrintLeftX As Single
    Dim sngPrintTopY As Single
    Dim itmItem As ListItem
    Dim intResult As Long
    Dim intTemp As Long
    Dim intCnt As Long
    Dim intLPR As Long
    Dim intCut As Long
    Dim intPos As Long
    Dim intMax As Long
    Dim intPageOffset As Long
    Dim blnTemp As Boolean
    Dim lngMinSize As Long
    
    intTQty = 0
    intTemp = 0
    
    'chg 200504 (this var may go away if we implement new sizing routine)
    'Max chars per printed line
    intMax = 7
    PageHeight = 49
    
    'Max Page height is 49.2", using 49" (49 * 1440) Twips
    
    'If SelectForm("Kasa Wire Labels PTR2(e)", Me.hwnd, 4, 49) = 0 Then
    '    ' Selection failed!
    '    MsgBox "Error setting form! Cannot print without setting form!", vbOKOnly + vbCritical, "Kasa Wire Labels"
    '    Exit Sub
    'End If
    
    ' If we are only printing the selected items, the only count
    ' the quantities of the selected labels.
    
    'If Not blnPrintAll Then
    If Not False Then
        ' Only count the selected labels
        For Each itmItem In frmMain.lsvLabels.ListItems
            If itmItem.Selected = True Then
                intTQty = intTQty + Val(itmItem.SubItems(1))
            End If
        Next itmItem
    Else
        ' Count all the labels in the list
        For Each itmItem In frmMain.lsvLabels.ListItems
            intTQty = intTQty + Val(itmItem.SubItems(1))
        Next itmItem
    End If
    
    ' Recalculate the total quantity of labels by multiplying
    ' the single list times the number of copies needed
    intTQty = intTQty * intCopies
    
    'chg 200504
    'DOUBLE CHECK QUANTITY to ensure non-zero!
    If intTQty = 0 Then
        MsgBox "No Labels Selected.  Nothing to do!", vbExclamation
        Exit Sub
    End If
    
    ' Re-Dimension the array to hold all the labels it counted
    ReDim strLabel(intTQty - 1)
    ReDim lngSize(intTQty - 1)
    
    
    
    
    
    'Split the text on the wire labels to two lines if necessary
    'If Not blnPrintAll Then
    If Not False Then
        For intResult = 1 To intCopies
            ' Add only the selected label texts to the array
            For Each itmItem In frmMain.lsvLabels.ListItems
                If itmItem.Selected = True Then
                    For intCnt = 1 To Val(itmItem.SubItems(1))
                        strLabel(intTemp) = SplitText(itmItem.Text, intMax)
                        intTemp = intTemp + 1
                    Next intCnt
                End If
            Next itmItem
        Next intResult
    Else
        For intResult = 1 To intCopies
            ' Add all the label texts to the array
            For Each itmItem In frmMain.lsvLabels.ListItems
                For intCnt = 1 To Val(itmItem.SubItems(1))
                    strLabel(intTemp) = SplitText(itmItem.Text, intMax)
                    intTemp = intTemp + 1
                Next intCnt
            Next itmItem
        Next intResult
    End If
    
    
    'Add a dummy labels to provide space at end to cut the labels from spool
    intTQty = intTQty + 3
    ReDim Preserve strLabel(intTQty - 1)
    ReDim Preserve lngSize(intTQty - 1)
    strLabel(intTQty - 3) = "eoj marker"
    strLabel(intTQty - 2) = "eoj marker"
    strLabel(intTQty - 1) = "eoj marker"
    
    intTemp = 0
    
    ' Adjust the printer's margins from what the label format is
'    If sngLeftMargin > 0.125 Then
'        sngPrintLeftX = sngLeftMargin - 0.125
'    Else
'        sngPrintLeftX = 0
'    End If
    
    sngPrintLeftX = sngLeftMargin - 0.125
    sngPrintTopY = sngTopMargin
    
    ' Set the printer's font settings
    With picPrinter
        .FontBold = frmMain.cdlDialog.FontBold
        .FontItalic = frmMain.cdlDialog.FontItalic
        .FontName = frmMain.cdlDialog.FontName
        If intAutoSize = 1 Then
            'do nothing
        Else
            .FontSize = frmMain.cdlDialog.FontSize
        End If
        .FontStrikethru = frmMain.cdlDialog.FontStrikethru
        .FontUnderline = frmMain.cdlDialog.FontUnderline
    End With
    
  
    'chg 200504
    'Print until all 'pages' are finished
    
    picPrinter.Width = ((intLabelsPerRow - 1) * sngSpacingRL + sngWidth + sngLeftMargin) * 1440
    Debug.Print "picprinter width : " & picPrinter.Width / Screen.TwipsPerPixelX / 96
    
    'Set number of pages
    If ((intTQty) * sngSpacingTB) > (PageHeight - sngPrintTopY) Then
        Pages = CInt(((intTQty) * sngSpacingTB) / (PageHeight - sngPrintTopY)) + 1
        If Pages - ((intTQty) * sngSpacingTB) / (PageHeight - sngPrintTopY) > 1 Then
            Pages = Pages - 1
        End If
    Else
        Pages = 1
    End If
    
    CurPage = 1
    CurLabel = 0
    
    
    Do
    
    
    On Error Resume Next
    
    If CurPage = Pages Then
   
        PgLastLabel = UBound(strLabel)
    
        picPrinter.Height = sngPrintTopY * 1440 + (PgLastLabel - CurLabel + 1) * sngSpacingTB * 1440
    
    Else
        

        Debug.Assert (CurPage < Pages)
    
        PgLastLabel = CurLabel + CInt((PageHeight - sngPrintTopY) / sngSpacingTB) - 1
    
        picPrinter.Height = 1440 * PageHeight
    
    End If
    
    'Debug.Print "picprinter height: " & picPrinter.Height / Screen.TwipsPerPixelX / 96
    
    If Err Then
        'bmk todo: comment the stop
        Debug.Print Err.Number & " " & Err.Description
        MsgBox "Unexpected Error: " & vbCr & Err.Number & vbCr & Err.Description, "PrintStrips"
        'Stop
        Err.Clear
    End If
    
    'After the labels have been chosen, text has been split if necessary,
    '   then lines have been sized to proper font, we're finally ready to print!
    
    intPageOffset = CurLabel
    
    For intCnt = CurLabel To PgLastLabel
        
'        ' Set the current printer position
'        If strLabel(intCnt) <> "eoj marker" Then
'            picPrinter.CurrentY = (intCnt - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440
'        Else
'            picPrinter.CurrentY = (intCnt - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440 + 1440 * 0.125
'        End If
        
        
        ' Loop through all lines of text of the label
        Do
            ' Set the horizontal position of the printer
            picPrinter.CurrentX = sngPrintLeftX * 1440
            ' Check the text for a return and line feed
            If InStr(1, strLabel(intCnt), vbCrLf) = 0 Then
                ' There is no line break representing
                '   two lines of text, so just print the text.
                sFirstLine = strLabel(intCnt)
                sSecondLine = ""
                strLabel(intCnt) = ""
            Else
                ' Print all the text up to but not including the line break
                ' Then trim off remaining text w/o line break
                sFirstLine = Left$(strLabel(intCnt), InStr(1, strLabel(intCnt), vbCrLf) - 1)
                sSecondLine = Mid$(strLabel(intCnt), InStr(1, strLabel(intCnt), vbCrLf) + Len(vbCrLf))
                'strLabel(intCnt) = Mid$(strLabel(intCnt), InStr(1, strLabel(intCnt), vbCrLf) + Len(vbCrLf))
                strLabel(intCnt) = ""
            End If
        
        
            ' If we are automatically sizing the text, then set
            ' the picPrinter font size to the size of the text for
            ' this specific label
            If intAutoSize = 1 Then
                If Trim(sFirstLine) <> "" Then
                    If Len(sFirstLine) = 4 Or Len(sFirstLine) = 5 Then
                        'some reason this many chars spills over the allotted width
                        lngSize(intCnt) = frmMain.SizeToText(sngWidth - 1.5 / 16, sngHeight - 1 / 16, ActualTextHeight, sFirstLine, sSecondLine)
                    Else
                        lngSize(intCnt) = frmMain.SizeToText(sngWidth - 1 / 16, sngHeight - 1 / 16, ActualTextHeight, sFirstLine, sSecondLine)
                    End If
                Else
                    sFirstLine = " "
                End If
            
                If lngSize(intCnt) >= 3 Then
                    picPrinter.Font.Size = lngSize(intCnt)
                Else
                    Debug.Assert False
                    picPrinter.Font.Size = 3
                End If
            
            End If
            
            
        ' Set the current printer position
        If strLabel(intCnt) <> "eoj marker" Then
            'chg 20050505
            'The labels should be centered vertically rather than top justified
            picPrinter.CurrentY = (intCnt - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440 + (sngHeight * 1440 - ActualTextHeight) * 0.5
        Else
            picPrinter.CurrentY = (intCnt - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440 + (sngHeight * 1440 - ActualTextHeight) * 0.5 + 1440 * 0.125
        End If
            
            
            picPrinter.Print sFirstLine
            If sSecondLine <> "" Then
                ' Set the horizontal position of the printer
                picPrinter.CurrentX = sngPrintLeftX * 1440
                picPrinter.Print sSecondLine
            End If
            
        ' Continue looping until the entire label is printed
        Loop Until Len(strLabel(intCnt)) = 0
        
        
        '20050505
        If intCnt < PgLastLabel Then
            If strLabel(intCnt + 1) <> "eoj marker" Then
                NextY = (intCnt + 1 - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440
            Else
                NextY = (intCnt + 1 - intPageOffset) * sngSpacingTB * 1440 + sngPrintTopY * 1440 + 1440 * 0.125
            End If
        Else
            NextY = picPrinter.CurrentY
        End If
        
        'Safety valve
        'shouldn't happen if we set PgLastLabel correctly!!
        If NextY > picPrinter.Height Then
            Debug.Assert (intCnt = PgLastLabel)
            Exit For
        End If
        
        DoEvents
    Next 'intCnt
    
    CurPage = CurPage + 1
    
    CurLabel = PgLastLabel + 1
    
    Loop Until CurPage > Pages
    

End Sub
