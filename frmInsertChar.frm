VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInsertChar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Character(s)"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optRtOff 
      Caption         =   "Characters from Right:"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin MSComCtl2.UpDown updOffset 
      Height          =   285
      Left            =   2176
      TabIndex        =   5
      Top             =   1920
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtOffset"
      BuddyDispid     =   196609
      OrigLeft        =   2400
      OrigTop         =   1920
      OrigRight       =   2640
      OrigBottom      =   2175
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtOffset 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.OptionButton optOffset 
      Caption         =   "Characters from Left:"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.OptionButton optRight 
      Caption         =   "Right"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optLeft 
      Caption         =   "Left"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtInsert 
      Height          =   285
      Left            =   1680
      MaxLength       =   13
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame fraInsert 
      Caption         =   "Insert Location"
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   2535
      Begin VB.TextBox txtRtOff 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Text            =   "0"
         Top             =   1800
         Width           =   615
      End
      Begin MSComCtl2.UpDown updRtOff 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   1800
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtRtOff"
         BuddyDispid     =   196622
         OrigLeft        =   2400
         OrigTop         =   1920
         OrigRight       =   2640
         OrigBottom      =   2175
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
   End
   Begin VB.Label lblInsert 
      Caption         =   "Character(s) to Insert:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblSample 
      Caption         =   "Sample Output:"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmInsertChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strOutput As String

Private Sub cmdCancel_Click()
' Close out the Insert Character dialog box

    Unload frmInsertChar
End Sub

Private Sub cmdOK_Click()
' This subroutine goes through the list of selected labels and
' inserts the character(s) specified at the specified location.

    Dim itmItem As ListItem
    Dim strText As String
    
    ' Perform code on every item in the list
    For Each itmItem In frmMain.lsvLabels.ListItems
        ' Restrict the code to only selected items
        If itmItem.Selected = True Then
            strText = Trim$(txtInsert.Text)
            strOutput = Trim$(itmItem.Text)
    
            If optLeft.Value = True Then
                ' Insert character(s) on left
                itmItem.Text = strText & strOutput
            ElseIf optRight.Value = True Then
                ' Insert character(s) on right
                itmItem.Text = strOutput & strText
            ElseIf optOffset.Value = True Then
                ' Insert character(s) offset from left
                If Val(txtOffset.Text) = 0 Then
                    ' Insert them on the left - position zero
                    itmItem.Text = strText & strOutput
                ElseIf Val(txtOffset.Text) >= Len(strOutput) Then
                    ' Insert them on the right - position greater than or equal to length
                    itmItem.Text = strOutput & strText
                Else
                    ' Insert them somewhere in the middle of the string
                    itmItem.Text = Left$(strOutput, Val(txtOffset.Text)) & strText & Mid$(strOutput, Val(txtOffset.Text) + 1)
                End If
            Else
                ' Insert character(s) offset from right
                If Val(txtRtOff.Text) = 0 Then
                    ' Insert them on the right
                    itmItem.Text = strOutput & strText
                ElseIf Val(txtRtOff.Text) >= Len(strOutput) Then
                    ' Insert them on the left
                    itmItem.Text = strText & strOutput
                Else
                    ' Insert them somewhere in the middle of the string
                    itmItem.Text = Left$(strOutput, (Len(strOutput) - updRtOff.Value)) & strText & Right$(strOutput, updRtOff.Value)
                End If
            End If
        End If
    Next itmItem
    
    ' Label list has changed, and it isn't saved yet.
    blnSaved = False

    ' Close the insert character dialog box
    Unload frmInsertChar
End Sub

Private Sub Form_Load()
' This subroutine finds the longest item that's selected
' and uses it for the example text.

    Dim itmItem As ListItem
    Dim intMax As Long
    
    For Each itmItem In frmMain.lsvLabels.ListItems
        If itmItem.Selected = True Then
            If Len(itmItem.Text) > intMax Then
                intMax = Len(itmItem.Text)
                strOutput = itmItem.Text
            End If
        End If
    Next itmItem
    
    lblOutput.Caption = strOutput
    updOffset.Max = intMax
    updRtOff.Max = intMax
    updRtOff.Value = intMax
End Sub

Private Sub optLeft_Click()
' This subroutine shows a preview with the characters
' inserted on the left.

    Dim strText As String
    
    strText = Trim$(txtInsert.Text)
    
    txtOffset.Enabled = False
    updOffset.Value = 0
    txtRtOff.Enabled = False
    updRtOff.Value = Len(strOutput)
    lblOutput.Caption = strText & strOutput
End Sub

Private Sub optOffset_Click()
' This subroutine shows the preview with the text offset
' the specified number of characters from the left.

    Dim strText As String
    
    strText = Trim$(txtInsert.Text)
    
    txtRtOff.Enabled = False
    txtRtOff.Text = ""
    txtOffset.Enabled = True
    txtOffset.Text = Trim$(Str$(Val(updOffset.Value)))
    
    If Val(txtOffset.Text) = 0 Then
        ' Insert on left
        lblOutput.Caption = strText & strOutput
    ElseIf Val(txtOffset.Text) >= Len(strOutput) Then
        ' Insert on right
        lblOutput.Caption = strOutput & strText
    Else
        ' Insert in the middle somewhere
        lblOutput.Caption = Left$(strOutput, Val(txtOffset.Text)) & strText & Mid$(strOutput, Val(txtOffset.Text) + 1)
    End If
End Sub

Private Sub optRight_Click()
' This subroutine shows the preview with the characters
' inserted on the right.
    
    Dim strText As String
    
    strText = Trim$(txtInsert.Text)
    
    txtOffset.Enabled = False
    updOffset.Value = Len(strOutput)
    txtRtOff.Enabled = False
    updRtOff.Value = 0
    lblOutput.Caption = strOutput & strText
End Sub

Private Sub optRtOff_Click()
' This subroutine shows the preview with the text offset
' the specified number of characters from the right.

    Dim strText As String
    
    strText = Trim$(txtInsert.Text)
    
    txtOffset.Enabled = False
    txtOffset.Text = ""
    txtRtOff.Enabled = True
    txtRtOff.Text = Trim$(Str$(updRtOff.Value))
    
    If Val(txtRtOff.Text) = 0 Then
        ' Insert on right
        lblOutput.Caption = strOutput & strText
    ElseIf Val(txtRtOff.Text) >= Len(strOutput) Then
        ' Insert on left
        lblOutput.Caption = strText & strOutput
    Else
        ' Insert in middle somewhere
        lblOutput.Caption = Left$(strOutput, (Len(strOutput) - updRtOff.Value)) & strText & Right$(strOutput, updRtOff.Value)
    End If
End Sub

Private Sub txtInsert_Change()
' This subroutine updates the preview whenever the insertion
' text is changed, to reflect the new insertion character(s).

    Dim strText As String
    
    strText = Trim$(txtInsert.Text)
    
    If optLeft.Value = True Then
        ' Insert on left
        lblOutput.Caption = strText & strOutput
    ElseIf optRight.Value = True Then
        ' Insert on right
        lblOutput.Caption = strOutput & strText
    ElseIf optOffset.Value = True Then
        ' Insert characters offset from the left
        If Val(txtOffset.Text) = 0 Then
            ' Insert on left
            lblOutput.Caption = strText & strOutput
        ElseIf Val(txtOffset.Text) >= Len(strOutput) Then
            ' Insert on right
            lblOutput.Caption = strOutput & strText
        Else
            ' Insert in middle somewhere
            lblOutput.Caption = Left$(strOutput, Val(txtOffset.Text)) & strText & Mid$(strOutput, Val(txtOffset.Text) + 1)
        End If
    Else
        ' Insert characters offset from the right
        If Val(txtRtOff.Text) = 0 Then
            ' Insert on right
            lblOutput.Caption = strOutput & strText
        ElseIf Val(txtRtOff.Text) >= Len(strOutput) Then
            ' Insert on left
            lblOutput.Caption = strText & strOutput
        Else
            ' Insert in middle somewhere
            lblOutput.Caption = Left$(strOutput, (Len(strOutput) - updRtOff.Value)) & strText & Right$(strOutput, updRtOff.Value)
        End If
    End If
    
End Sub

Private Sub txtInsert_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtOffset_Change()
' This subroutine updates the preview if the left offset is changed.

    Dim strText As String
    
    strText = Trim$(txtInsert.Text)
    
    If Val(txtOffset.Text) = 0 Then
        ' Insert on left
        lblOutput.Caption = strText & strOutput
    ElseIf Val(txtOffset.Text) >= Len(strOutput) Then
        ' Insert on right
        lblOutput.Caption = strOutput & strText
    Else
        ' Insert in middle somewhere
        lblOutput.Caption = Left$(strOutput, Val(txtOffset.Text)) & strText & Mid$(strOutput, Val(txtOffset.Text) + 1)
    End If
End Sub

Private Sub txtOffset_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtOffset_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the offset text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyLeft _
    And KeyAscii <> vbKeyRight And KeyAscii <> vbKeyReturn _
    And KeyAscii <> vbKeyTab Then
        KeyAscii = 0
    End If
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub txtOffset_LostFocus()
' This subroutine converts the text to a correct value.

    txtOffset.Text = Int(Val(txtOffset.Text))
End Sub

Private Sub txtRtOff_Change()
' This subroutine updates the preview if the right offset changes.

    Dim strText As String
    
    strText = Trim$(txtInsert.Text)
    
    If Val(txtRtOff.Text) = 0 Then
        ' Insert on right
        lblOutput.Caption = strOutput & strText
    ElseIf Val(txtRtOff.Text) >= Len(strOutput) Then
        ' Insert on left
        lblOutput.Caption = strText & strOutput
    Else
        ' Insert in middle somewhere
        lblOutput.Caption = Left$(strOutput, (Len(strOutput) - updRtOff.Value)) & strText & Right$(strOutput, updRtOff.Value)
    End If
End Sub

Private Sub txtRtOff_GotFocus()
' This subroutine selects all the text in the text box.

    SelectText
End Sub

Private Sub txtRtOff_KeyPress(KeyAscii As Integer)
' This subroutine keeps the user from entering a non-numeric
' character in the offset text box.  If the user presses
' the return key, it automatically tabs to the next field.

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) _
    And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyLeft _
    And KeyAscii <> vbKeyRight And KeyAscii <> vbKeyReturn _
    And KeyAscii <> vbKeyTab Then
        KeyAscii = 0
    End If
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call PostMessage(Me.hwnd, WM_KEYDOWN, VK_TAB, 0)
    End If
End Sub

Private Sub updOffset_DownClick()
' This subroutine gives the left offset option focus if the
' down button is clicked next to the left offset text.

    optOffset.SetFocus
End Sub

Private Sub updOffset_UpClick()
' This subroutine gives the left offset option focus if the
' up button is clicked next to the left offset text.

    optOffset.SetFocus
End Sub

Private Sub updRtOff_DownClick()
' This subroutine gives the right offset option focus if the
' down button is clicked next to the right offset text.

    optRtOff.SetFocus
End Sub

Private Sub updRtOff_LostFocus()
' This subroutine converts the text to a correct value.

    txtRtOff.Text = Int(Val(txtRtOff.Text))
End Sub

Private Sub updRtOff_UpClick()
' This subroutine gives the right offset option focus if the
' up button is clicked next to the right offset text.

    optRtOff.SetFocus
End Sub
