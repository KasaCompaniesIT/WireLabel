VERSION 5.00
Begin VB.Form frmFieldTagTest 
   AutoRedraw      =   -1  'True
   Caption         =   "Field Tag Print Preview"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFieldTagTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkShowPics 
      Caption         =   "Show Graphics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8280
      TabIndex        =   7
      Top             =   7440
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrintTest 
      Caption         =   "Print Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4560
      TabIndex        =   6
      Top             =   6600
      Width           =   1575
   End
   Begin VB.ComboBox cboFontSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4560
      TabIndex        =   3
      Top             =   7560
      Width           =   1575
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   1080
      ScaleHeight     =   5730
      ScaleWidth      =   8610
      TabIndex        =   0
      Top             =   480
      Width           =   8640
      Begin VB.PictureBox picGraphic2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   6480
         Picture         =   "frmFieldTagTest.frx":0442
         ScaleHeight     =   1140
         ScaleWidth      =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   1560
      End
      Begin VB.PictureBox picGraphic1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         Picture         =   "frmFieldTagTest.frx":7258
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   1
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Label lblFontSize 
      AutoSize        =   -1  'True
      Caption         =   "FontSize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8280
      TabIndex        =   5
      Top             =   6600
      Width           =   615
   End
End
Attribute VB_Name = "frmFieldTagTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VerticalMargin As Single
Dim LeftMargin As Single
Dim RightMargin As Single
Dim CenterMargin As Single
Dim LabelHeight As Single
Dim LabelWidth As Single

Dim p As Printer
Dim printOne As Boolean
Dim Loading As Boolean


Private Sub chkShowPics_Click()
    
    If chkShowPics.Value = 1 Then
        picGraphic1.Visible = True
        picGraphic2.Visible = True
    Else
        picGraphic1.Visible = False
        picGraphic2.Visible = False
    End If
    
    
End Sub

Private Sub Form_Load()

    Loading = True
    
    cmdPrintTest.Enabled = False
    cmdPrint.Enabled = False
    
    ' Populate the font size box
    cboFontSize.AddItem (20)
    cboFontSize.AddItem (22)
    cboFontSize.AddItem (24)
    cboFontSize.AddItem (26)
    cboFontSize.AddItem (28)
    cboFontSize.AddItem (30)
    cboFontSize.AddItem (36)
    cboFontSize.AddItem (42)
    cboFontSize.AddItem (48)

    ' set default
    cboFontSize.Text = 36
       
    ' Set the user specified picture
    picGraphic1.Picture = LoadPicture(frmField.dlgPic.FileName)
    
    ' Set the margins
    VerticalMargin = 0.1 * 1440
    LeftMargin = 0.3 * 1440
    RightMargin = 0.2 * 1440
    CenterMargin = 0.25 * 1440
    LabelWidth = 6 * 1440
    LabelHeight = 4 * 1440
    
    ' Set the location of the pictures
    picGraphic1.Move LeftMargin, VerticalMargin
    picGraphic2.Move picPreview.Width - picGraphic2.Width - RightMargin, VerticalMargin
    Dim CurrentY As Single
    picPreview.CurrentY = Max(picGraphic1.Height, picGraphic2.Height) + VerticalMargin
    
    ' Set up lines array
    Dim cols As Integer
    cols = Val(frmField.txtNumColumns.Text)
'    Dim Lines() As String
'    ReDim Lines(cols)
    Dim prevpos As Integer
    prevpos = 1
    Dim j As Integer
    Dim num As Integer
    
    Dim i As Integer
    Dim start As Integer
    
    
    If frmField.chkColHeaders = 1 Then
        start = 1
    Else
        start = 0
    End If
    
    Call PaintLines
    
    Loading = False

    cmdPrintTest.Enabled = True

End Sub


Private Sub cboFontSize_Click()
    
    If Loading Then Exit Sub
    
    Call PaintLines
    
End Sub


Private Sub cmdPrint_Click()

    Dim cols As Integer
    cols = Val(frmField.txtNumColumns.Text)
    Dim i As Integer
    Dim j As Integer
    Dim num As Integer
    
    Dim oldFont As String
    Dim oldSize As String
    
    Dim Lines() As String
    ReDim Lines(cols)

    Dim oldHeight As Integer
    Dim oldWidth As Integer
    Dim p As Printer

    Dim numLabels As Integer
    Dim start As Integer

    Dim LowestFontSize As Single
    Dim CurrentY As Single

'bmk - Local Printer
'===
#If DebugMode = 1 Then
    For Each p In Printers
    
        If InStr(p.DeviceName, "8000n") > 0 Then
            Set Printer = p
            Exit For
    
        End If
    
    Next p
    '===

#Else
    '===
    'CHG 200510 N.F.
    'This assertion should be true as long as CheckProperPrinter was called.
    Debug.Assert (InStr(LCase(Printer.DeviceName), LCase("TDP42H")) > 0 Or _
     InStr(LCase(Printer.DeviceName), LCase("M84 Pro")) > 0 Or _
     InStr(LCase(Printer.DeviceName), LCase("M-8400RV")) > 0 Or _
     InStr(LCase(Printer.DeviceName), LCase("ptr3")) > 0)
    
    'CHG 200510 N.F.
    'Try to select the proper printer again
    If Not (InStr(LCase(Printer.DeviceName), LCase("TDP42H")) > 0 Or _
     InStr(LCase(Printer.DeviceName), LCase("M84 Pro")) > 0 Or _
     InStr(LCase(Printer.DeviceName), LCase("M-8400RV")) > 0 Or _
     InStr(LCase(Printer.DeviceName), LCase("ptr3")) > 0) Then
        If Not CheckProperPrinter Then End
    End If
    
    Printer.Orientation = vbPRORLandscape

'===
#End If

    Debug.Print Printer.DeviceName
    oldFont = Printer.FontName
    oldSize = Printer.FontSize
    Printer.FontName = picPreview.FontName
    Printer.FontBold = picPreview.FontBold
    Printer.FontItalic = picPreview.FontItalic
    
    Dim PrintTestEnabled As Boolean
    Dim PrintEnabled As Boolean
    
    'Disable the print buttons
    PrintTestEnabled = cmdPrintTest.Enabled
    PrintEnabled = cmdPrint.Enabled
    cmdPrintTest.Enabled = False
    cmdPrint.Enabled = False
    Me.MousePointer = vbHourglass
    
    If printOne = True Then
        start = 0
        numLabels = 0
    Else
        start = 0
        numLabels = frmMain.lsvLabels.ListItems.Count - 1
    
    End If
    
    Printer.FontSize = Val(cboFontSize.Text)
    
    ' Print the labels
    For j = start To numLabels
    
        Call ParseString2(frmMain.lsvLabels.ListItems(j + 1), Lines(), num)
        num = 0
        
         ' Print the graphics
         If picGraphic1.Visible Then
             Printer.PaintPicture picGraphic1.Image, LeftMargin, VerticalMargin
         End If
         If picGraphic2.Visible Then
             Printer.PaintPicture picGraphic2.Image, LabelWidth - picGraphic2.Width - RightMargin, VerticalMargin
         End If
        
        Printer.CurrentY = Max(picGraphic1.Height, picGraphic2.Height) + VerticalMargin
    
        If frmField.chkCombineLastTwo.Value = 1 Then
            'Print first Cols - 2 columns as separate line, centered
            'Print last 2 columns combined on last line, L / R justified
            
            For i = 0 To cols - 3
            
                'If user selects 36, lowestfontsize is that index
                LowestFontSize = cboFontSize.ListIndex
                Printer.FontSize = Val(cboFontSize.List(LowestFontSize))
                
                Do While LabelWidth - Printer.TextWidth(Lines(i)) < 576
                    If LowestFontSize > 0 Then
                        Printer.FontSize = Val(cboFontSize.List(LowestFontSize - 1))
                        LowestFontSize = LowestFontSize - 1
                    Else
                        Printer.FontSize = Printer.FontSize - 2
        '                LowestFontSize = LowestFontSize - 1
                    End If
                Loop
                
                Printer.CurrentX = 0.5 * (LabelWidth - Printer.TextWidth(Lines(i)))
                Printer.Print Lines(i)
            Next i
            
            
            LowestFontSize = cboFontSize.ListIndex
            Printer.FontSize = Val(cboFontSize.List(LowestFontSize))
            
            Do While LabelWidth - Printer.TextWidth(Lines(cols - 1) & Lines(cols - 2)) < 936
                If LowestFontSize > 0 Then
                    Printer.FontSize = Val(cboFontSize.List(LowestFontSize - 1))
                    LowestFontSize = LowestFontSize - 1
                Else
                    Printer.FontSize = Printer.FontSize - 2
                End If
            Loop
            Printer.CurrentX = LeftMargin
            CurrentY = Printer.CurrentY
            Printer.Print Lines(cols - 2)
            Printer.CurrentX = (LabelWidth - RightMargin - Printer.TextWidth(Lines(cols - 1)))
            Printer.CurrentY = CurrentY
            Printer.Print Lines(cols - 1)
             
            If printOne Then
                 Printer.Line (0, 0)-(LabelWidth, LabelHeight), , B
             Else
                 Printer.NewPage
            End If

        Else
            ' Print each column as a separate line, centered
               For i = 0 To cols - 1
                    
                'If user selects 36, lowestfontsize is that index
                LowestFontSize = cboFontSize.ListIndex
                Printer.FontSize = Val(cboFontSize.List(LowestFontSize))
                
                Do While LabelWidth - Printer.TextWidth(Lines(i)) < 576
                    If LowestFontSize > 0 Then
                        Printer.FontSize = Val(cboFontSize.List(LowestFontSize - 1))
                        LowestFontSize = LowestFontSize - 1
                    Else
                        Printer.FontSize = Printer.FontSize - 2
        '                LowestFontSize = LowestFontSize - 1
                    End If
                Loop
                    
                    Printer.CurrentX = 0.5 * (LabelWidth - Printer.TextWidth(Lines(i)))
                    Printer.Print Lines(i)
               Next i
               
               If printOne Then
                    Printer.Line (0, 0)-(LabelWidth, LabelHeight), , B
                Else
                    Printer.NewPage
               End If
        
        End If
    
    Next j

    ' reset old settings
    Printer.EndDoc
    
'SAMPLE D
'    Printer.Orientation = vbPRORPortrait
'   SAMPLE C
'    Printer.Height = oldHeight
'    Printer.Width = oldWidth
    Printer.FontName = oldFont
    Printer.FontSize = oldSize
    
    printOne = False
    
    'Re-enable the print buttons
    cmdPrintTest.Enabled = PrintTestEnabled
    cmdPrint.Enabled = PrintEnabled
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdPrintTest_Click()

    ' Disable the print buttons
    cmdPrint.Enabled = False
    cmdPrintTest.Enabled = False

    ' Send the first label to the printer
    Dim Result As VbMsgBoxResult
    
    'Call the print routine
    printOne = True
    Call cmdPrint_Click
    
    ' Ask the user if they want to continue
    Result = MsgBox("Was the label printed correctly?", vbQuestion + vbYesNo)
    
    'Re-enable the print test button
    cmdPrintTest.Enabled = True
    
    ' Enable the print button if yes
    If Result = vbYes Then
        cmdPrint.Enabled = True
    Else
        MsgBox "Please re-align the labels and try another Print Test.", vbInformation + vbOKOnly
    End If
    

End Sub


Private Function Max(ByRef X As Single, ByRef Y As Single) As Single

If X >= Y Then
    Max = X
Else
    Max = Y
End If
End Function

Private Sub picPreview_Click()

    picPreview.FontSize = cboFontSize.Text

End Sub

Private Sub PaintLines()

    Dim cols As Integer
    Dim Lines() As String
    Dim num As Integer
    Dim i As Integer
    Dim LowestFontSize As Single
    Dim CurrentY As Single
    
    picPreview.FontSize = Val(cboFontSize.Text)
    
    cols = Val(frmField.txtNumColumns.Text)
    ReDim Lines(cols)
    
    Call ParseString2(frmMain.lsvLabels.ListItems(1), Lines(), num)
    num = 0
           
    picPreview.Cls
    picPreview.CurrentY = Max(picGraphic1.Height, picGraphic2.Height) + VerticalMargin

    
    
    If frmField.chkCombineLastTwo.Value = 1 Then
        'Print first Cols - 2 columns as separate line, centered
        'Print last 2 columns combined on last line, L / R justified
        
        For i = 0 To cols - 3
        
            'If user selects 36, lowestfontsize is that index
            LowestFontSize = cboFontSize.ListIndex
            picPreview.FontSize = Val(cboFontSize.List(LowestFontSize))
            
            Do While LabelWidth - picPreview.TextWidth(Lines(i)) < (LeftMargin + RightMargin)
                If LowestFontSize > 0 Then
                    picPreview.FontSize = Val(cboFontSize.List(LowestFontSize - 1))
                    LowestFontSize = LowestFontSize - 1
                Else
                    picPreview.FontSize = picPreview.FontSize - 2
    '                LowestFontSize = LowestFontSize - 1
                End If
            Loop
            
            picPreview.CurrentX = 0.5 * (LabelWidth - picPreview.TextWidth(Lines(i)))
            picPreview.Print Lines(i)
        Next i
        
        
        LowestFontSize = cboFontSize.ListIndex
        picPreview.FontSize = Val(cboFontSize.List(LowestFontSize))
        
        Do While LabelWidth - picPreview.TextWidth(Lines(cols - 1) & Lines(cols - 2)) < (LeftMargin + CenterMargin + RightMargin)
            If LowestFontSize > 0 Then
                picPreview.FontSize = Val(cboFontSize.List(LowestFontSize - 1))
                LowestFontSize = LowestFontSize - 1
            Else
                picPreview.FontSize = picPreview.FontSize - 2
            End If
        Loop
        picPreview.CurrentX = LeftMargin
        CurrentY = picPreview.CurrentY
        picPreview.Print Lines(cols - 2)
        picPreview.CurrentX = (LabelWidth - RightMargin - picPreview.TextWidth(Lines(cols - 1)))
        picPreview.CurrentY = CurrentY
        picPreview.Print Lines(cols - 1)
        
    Else
        ' Print each column as a separate line, centered
        For i = 0 To cols - 1
        
            'If user selects 36, lowestfontsize is that index
            LowestFontSize = cboFontSize.ListIndex
            picPreview.FontSize = Val(cboFontSize.List(LowestFontSize))
            
            Do While LabelWidth - picPreview.TextWidth(Lines(i)) < 576
                If LowestFontSize > 0 Then
                    picPreview.FontSize = Val(cboFontSize.List(LowestFontSize - 1))
                    LowestFontSize = LowestFontSize - 1
                Else
                    picPreview.FontSize = picPreview.FontSize - 2
    '                LowestFontSize = LowestFontSize - 1
                End If
            Loop
            
            picPreview.CurrentX = 0.5 * (LabelWidth - picPreview.TextWidth(Lines(i)))
            picPreview.Print Lines(i)
        Next i
    End If
    
    picPreview.Line (0, 0)-(picPreview.Width, picPreview.Height), , B


    If frmField.lblPicFile.Caption = "" Then
        picGraphic1.Visible = False
    Else
        picGraphic1.Visible = True
    End If
    
End Sub
