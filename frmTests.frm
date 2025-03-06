VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTests 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3836
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Label Text"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdRoundToNearestX 
      Caption         =   "Round To Nearest X"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdConfirmLabelsToPrintTest 
      Caption         =   "Confirm Labels To Print Test"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtResults 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton cmdSelectPrinterTest 
      Caption         =   "SelectPrinter Test"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConfirmLabelsToPrintTest_Click()
Dim TotalQty As Long
Dim PrintAll As Boolean

PrintResult "ConfirmLabelsToPrint(" & TotalQty & ", " & PrintAll & ", ListView1) : " & ConfirmLabelsToPrint(TotalQty, PrintAll, ListView1)
PrintResult "TotalQty = " & TotalQty
PrintResult "ConfirmLabelsToPrint(" & TotalQty & ", " & PrintAll & ", ListView1) : " & ConfirmLabelsToPrint(TotalQty, PrintAll, ListView1)
PrintResult "TotalQty = " & TotalQty
PrintAll = True
PrintResult "ConfirmLabelsToPrint(" & TotalQty & ", " & PrintAll & ", ListView1) : " & ConfirmLabelsToPrint(TotalQty, PrintAll, ListView1)
PrintResult "TotalQty = " & TotalQty

End Sub

Private Sub cmdRoundToNearestX_Click()
Dim n As Long

txtResults = ""

n = 5
Call RoundUpToNearestX(4, n)
PrintAssert (n = 8), "RoundUpToNearestX(4, 5)"
n = 5
Call RoundUpToNearestX(3, n)
PrintAssert (n = 6), "RoundUpToNearestX(3, 5)"

End Sub

Private Sub cmdSelectPrinterTest_Click()
    Dim blnTemp As Boolean

    blnTemp = False
    MsgBox "Click 'All' for Page Range, then Click Cancel on Printer Dialog"
    PrintResult "CheckIsOKtoPrint(Me.Hwnd, True, False) = " & CheckIsOKtoPrint(Me.hwnd, True, blnTemp)
    PrintResult "PrintAllLabels = " & blnTemp
    
    blnTemp = False
    MsgBox "Click 'Selection' for Page Range, then Click Cancel on Printer Dialog"
    PrintResult "CheckIsOKtoPrint(Me.Hwnd, True, False) = " & CheckIsOKtoPrint(Me.hwnd, False, blnTemp)
    PrintResult "PrintAllLabels = " & blnTemp

End Sub

Private Sub PrintAssert(ByVal TheAssertion As Boolean, _
 ByVal Desc As String)

If TheAssertion Then
    PrintResult (Desc & ": Successful")
Else
    PrintResult (Desc & ": Fails")
End If

End Sub

Private Sub PrintResult(ByVal s As String)

Debug.Print s
txtResults = txtResults & s & vbCrLf

End Sub

Private Sub Form_Load()
Dim i As Integer
With ListView1
    For i = 1 To 10
    .ListItems.Add , , "Label" & CStr(i)
    .ListItems(ListView1.ListItems.Count).Selected = True
    .SelectedItem.SubItems(1) = "1"
    .ListItems(ListView1.ListItems.Count).Selected = False
    Next i
End With
End Sub
