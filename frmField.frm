VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmField 
   Caption         =   "Field Tag Wizard"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   Icon            =   "frmField.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboSheets 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CheckBox chkCombineLastTwo 
      Caption         =   "Combine Last Two Columns (4 Line Format)"
      Height          =   375
      Left            =   3690
      TabIndex        =   10
      Top             =   2490
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkColHeaders 
      Caption         =   "Columun Headers in Spreadsheet"
      Height          =   375
      Left            =   630
      TabIndex        =   9
      Top             =   2490
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.TextBox txtNumColumns 
      Height          =   315
      Left            =   1860
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "5"
      Top             =   1950
      Width           =   735
   End
   Begin VB.TextBox txtExcel 
      Height          =   435
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   870
      Visible         =   0   'False
      Width           =   3525
   End
   Begin MSComDlg.CommonDialog dlgPic 
      Left            =   5490
      Top             =   330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdViewPreview 
      Caption         =   "Step 3: Preview Field Tags..."
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   5040
      Width           =   3405
   End
   Begin VB.CommandButton cmdSelectGraphic 
      Caption         =   "Step 2: Select Customer Graphic (Optional)..."
      Height          =   525
      Left            =   600
      TabIndex        =   2
      Top             =   3450
      Width           =   3375
   End
   Begin VB.CommandButton cmdSelectSpreadsheet 
      Caption         =   "Step 1: Select the Excel Workbook..."
      Height          =   525
      Left            =   600
      TabIndex        =   1
      Top             =   150
      Width           =   3525
   End
   Begin MSComDlg.CommonDialog dlgSpreadsheet 
      Left            =   4890
      Top             =   330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblRows 
      AutoSize        =   -1  'True
      Caption         =   "Rows in cboSheets:"
      Height          =   195
      Left            =   2940
      TabIndex        =   12
      Top             =   1950
      Width           =   1425
   End
   Begin VB.Label lblPicFile 
      Height          =   555
      Left            =   600
      TabIndex        =   5
      Top             =   4140
      Width           =   5445
   End
   Begin VB.Label lblSSFile 
      Height          =   435
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   5445
   End
   Begin VB.Label lblStep1 
      AutoSize        =   -1  'True
      Caption         =   "Select the Spreadsheet"
      Height          =   195
      Left            =   630
      TabIndex        =   0
      Top             =   810
      Width           =   1665
   End
   Begin VB.Label lblNumColumns 
      AutoSize        =   -1  'True
      Caption         =   "Columns to Print:"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   1950
      Width           =   1185
   End
End
Attribute VB_Name = "frmField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSheets_Click()
    Static Sheet As String
    
    If cboSheets.ListIndex < 0 Then Exit Sub
    If Sheet <> cboSheets.Text & CStr(chkColHeaders.Value) Then
        Sheet = cboSheets.Text & CStr(chkColHeaders.Value)
        
    Me.MousePointer = vbHourglass
    cmdSelectGraphic.Enabled = False
    cmdViewPreview.Enabled = False
    
    '*******************
    Dim e As Object
    Dim w As Object
    Dim s As Object

    Set e = CreateObject("Excel.Application")
    
    Set w = e.Workbooks.Open(lblSSFile.Caption, ReadOnly:=True)
    Set s = w.sheets(cboSheets.Text)
    Debug.Assert Not (s Is Nothing)
    
    On Error Resume Next
    
    If Not s Is Nothing Then
        With s
        .Range("A1").Select
        .Range(e.selection, e.ActiveCell.SpecialCells(11)).Select
        
    If Err Then
        e.Quit
        Set w = Nothing
        Set e = Nothing
        Me.MousePointer = vbDefault
        MsgBox "There is nothing in the sheet."
        Err.Clear
        Exit Sub
    End If
    
        e.Application.CutCopyMode = False
        e.selection.Copy
        End With
    End If
    
    txtExcel.Text = Clipboard.GetText
    'close workbook w/o saving changes
    e.ActiveWorkbook.Close False
    e.Quit
    Set w = Nothing
    Set e = Nothing
    
    DoEvents
    
    ' Call function to populate list view
    Populate

    Me.MousePointer = vbDefault
    cmdSelectGraphic.Enabled = True
    cmdViewPreview.Enabled = True

    lblRows.Caption = "Rows in " & cboSheets.Text & ": " & frmMain.lsvLabels.ListItems.Count
    
    End If
    
End Sub

Private Sub chkColHeaders_Click()

    Call cboSheets_Click
    
End Sub

Private Sub cmdSelectGraphic_Click()

    dlgPic.ShowOpen
    lblPicFile.Caption = dlgPic.FileName
    lblPicFile.BorderStyle = 1

End Sub

Private Sub cmdSelectSpreadsheet_Click()

    If Dir(lblSSFile.Caption) = "" Then
        lblSSFile.Caption = ""
        cboSheets.Enabled = False
    End If
    
    On Error GoTo NoFileSelected
    dlgSpreadsheet.Filter = "Excel Files|*.xls"
    dlgSpreadsheet.ShowOpen
    
    lblSSFile.Caption = dlgSpreadsheet.FileName
    lblSSFile.BorderStyle = 1
    'datExcel.DatabaseName = lblSSFile.Caption
    
LoadSheet:
    
    Me.MousePointer = vbHourglass
    
    Dim e As Object
    Dim w As Object

    Set e = CreateObject("Excel.Application")
    
    Set w = e.Workbooks.Open(lblSSFile.Caption, ReadOnly:=True)
    Dim i As Integer

    cboSheets.Clear

    For i = 1 To w.sheets.Count
        Debug.Assert Not (w.sheets(i) Is Nothing)
        With cboSheets
            .AddItem w.sheets(i).Name
        End With
    Next i
    
    e.Quit
    Set w = Nothing
    Set e = Nothing

    cboSheets.Enabled = True
    Me.MousePointer = vbDefault
    
    
    Exit Sub
    
NoFileSelected:
    If Err.Number = cdlCancel Then
        'No File selected
        If lblSSFile.Caption = "" Or Dir(lblSSFile.Caption) = "" Then
            cboSheets.Enabled = False
            MsgBox "You must select a spreadsheet to print Field Tags."
        Else
            Resume LoadSheet
        End If
    Else
    
    End If

End Sub

Private Sub cmdViewPreview_Click()


    ' Load frmFieldTagTest
    frmFieldTagTest.Show vbModal, Me

End Sub


Private Sub Populate()
    
    Dim Result() As String
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
    Dim strTemp As String

    ' this function populates the list view based on number of columns user specifies
    Call ParseString(txtExcel.Text, Result(), X)
    
    'Clear the old list
    frmMain.lsvLabels.ListItems.Clear
    
    Dim itm As ListItem
    Dim doFirst As Boolean
        
    ' Test for column headers
    If chkColHeaders = 1 Then
        doFirst = False
    Else
        doFirst = True
    End If


    ' do for each row
    For i = 0 To (X / Val(txtNumColumns.Text)) - 1
    
        strTemp = ""
        
        ' build a row
        For j = 0 To Val(txtNumColumns.Text) - 1
        
            ' build the string
            If j < Val(txtNumColumns.Text) - 1 Then
                strTemp = strTemp & Result((i * Val(txtNumColumns.Text)) + j) & "|"
            Else
                strTemp = strTemp & Result((i * Val(txtNumColumns.Text)) + j)
            End If
        
        Next j
        
        'CHG 200510 N.F.
        ' test for crlf, cr, lf on EITHER END of the string
        If Right(strTemp, 2) = vbCrLf Then
            strTemp = Left(strTemp, Len(strTemp) - 2)
        ElseIf Right(strTemp, 1) = vbCr Or Right(strTemp, 1) = vbLf Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        If Left(strTemp, 2) = vbCrLf Then
            strTemp = Right(strTemp, Len(strTemp) - 2)
        ElseIf Left(strTemp, 1) = vbCr Or Left(strTemp, 1) = vbLf Then
            strTemp = Right(strTemp, Len(strTemp) - 1)
        End If
        
        ' add the label text to the list view
        ' and ignore column headers
        If i > 0 Or doFirst = True Then
            Set itm = frmMain.lsvLabels.ListItems.Add(, , strTemp)
            itm.SubItems(1) = 1
        End If
    
    Next i

End Sub

Private Sub Form_Activate()
    Me.MousePointer = 0

End Sub

Private Sub Form_Load()

    cboSheets.Enabled = False
    lblRows.Caption = ""
    
'For Each p In Printers
'
'    If InStr(p.DeviceName, "8000dn") > 0 Then
'
'        Exit For
'
'    End If
'
'Next p

    cmdSelectGraphic.Enabled = False
    cmdViewPreview.Enabled = False
    Me.MousePointer = 11
    
'EARLY BINDING - ADD REFERENCE TO EXCEL 11.0
'===
'    Dim e As New Excel.Application
'    Dim w As New Excel.Workbook
'    Dim s As New Excel.Worksheet
'===
    
#If DebugMode = 1 Then
'LATE BINDING - REMOVE REFERENCE TO EXCEL 11.0
'===
    Dim e As Object
    Dim w As Object
    Dim s As Object

    Set e = CreateObject("Excel.Application")
'    w = CreateObject("Excel.Workbook")
'    s = CreateObject("Excel.Worksheet")
'===
    
    Set w = e.Workbooks.Open(App.Path & "\FieldTagTest.xls", ReadOnly:=True)
    Set s = w.sheets("Sheet1")
    Debug.Assert Not (s Is Nothing)
    
    If Not s Is Nothing Then
        'Load the fricken data
        With s
        .Range("A1").Select
        '.Range(e.Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        .Range(e.selection, e.ActiveCell.SpecialCells(11)).Select
        e.Application.CutCopyMode = False
        e.selection.Copy
        End With
    End If
    
    txtExcel.Text = Clipboard.GetText
    e.Quit
    Set w = Nothing
    Set e = Nothing
#End If

    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmMain.lsvLabels.ListItems.Clear

End Sub

