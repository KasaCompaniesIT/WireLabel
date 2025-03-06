VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2625
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.PictureBox picGraphic2 
         AutoSize        =   -1  'True
         Height          =   405
         Left            =   4200
         Picture         =   "frmFieldTags.frx":0000
         ScaleHeight     =   345
         ScaleWidth      =   360
         TabIndex        =   2
         Top             =   0
         Width           =   420
      End
      Begin VB.PictureBox picGraphic1 
         AutoSize        =   -1  'True
         Height          =   405
         Left            =   0
         Picture         =   "frmFieldTags.frx":066A
         ScaleHeight     =   345
         ScaleWidth      =   360
         TabIndex        =   1
         Top             =   0
         Width           =   420
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VerticalMargin As Single
Dim HorizontalMargin As Single

Private Sub cmdPrint_Click()

Printer.PaintPicture picPreview.Image, 0, 0
Printer.PaintPicture picGraphic1.Image, HorizontalMargin, VerticalMargin
Printer.PaintPicture picGraphic2.Image, picPreview.Width - picGraphic2.Width - HorizontalMargin, VerticalMargin
Printer.EndDoc

End Sub

Private Sub Form_Load()
VerticalMargin = 50
HorizontalMargin = 100
picGraphic1.Move HorizontalMargin, VerticalMargin
picGraphic2.Move picPreview.Width - picGraphic2.Width - HorizontalMargin, VerticalMargin
Dim CurrentY As Single
picPreview.CurrentY = Max(picGraphic1.Height, picGraphic2.Height) + VerticalMargin
Dim Lines(4) As String

Lines(0) = "Var Length Line 1"
Lines(1) = "Other Line 2"
Lines(2) = "New Line 3"
Lines(3) = "Extra Line 4"
Dim i As Integer


For i = 0 To 3
    picPreview.CurrentX = 0.5 * (picPreview.Width - picPreview.TextWidth(Lines(i)))
    picPreview.Print Lines(i)
    
Next i

End Sub

Private Function Max(ByRef x As Single, ByRef y As Single) As Single

If x >= y Then
    Max = x
Else
    Max = y
End If
End Function
