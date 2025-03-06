Attribute VB_Name = "Utilities"
Option Explicit

Public Sub TestParseString()

    Dim A() As String
    Dim X As Integer
    Dim i As Integer
'    Call ParseString("COLUMN1 COLUMN2 COLUMN3 COLUMN4 FIELD NAME  DEVICE NAME LABEL   NUMBER  LOCATION TEXT
'        FIELD NAME1 DEVICE NAME LABEL1  NUMBER1 LOCATION TEXT1
'        FIELD NAME2 DEVICE NAME LABEL3  NUMBER2 LOCATION TEXT2
'        FIELD NAME3 DEVICE NAME LABEL3  NUMBER3 LOCATION TEXT3
'        FIELD NAME4 DEVICE NAME LABEL4  NUMBER4 LOCATION TEXT4
'        ", A(), x)
    
    For i = 0 To X - 1
        Debug.Print A(i)
    Next

    Erase A
    X = 0

    Call ParseString("A, B", A(), X)
    For i = 0 To X - 1
        Debug.Print A(i)
    Next


End Sub


Public Sub ParseString(ByVal myString As String, ByRef arOutput() As String, ByRef numElements As Integer)

    Dim counter As Integer
    counter = InStr(myString, Chr(9))
    
    If counter >= 0 Then
        counter = InStr(myString, Chr(9))
    End If
    
    Dim trimme As Integer
    
    Do While counter > 0
        
        numElements = numElements + 1
        ReDim Preserve arOutput(numElements)
        arOutput(numElements - 1) = Trim(Left(myString, counter - 1))
        ' trim the carriege return off the end of the string
        arOutput(numElements - 1) = Mid(arOutput(numElements - 1), 1, Len(arOutput(numElements - 1))) '- 1)
        myString = Trim(Right(myString, Len(myString) - counter))
        counter = InStr(myString, Chr(9))
    
        ' if a <Enter> was found before a <TAB> then reduce counter size
        If InStr(myString, vbCrLf) < counter Then
            counter = InStr(myString, vbCrLf) + 1
        End If

    Loop
    
        numElements = numElements + 1
        ReDim Preserve arOutput(numElements)
        arOutput(numElements - 1) = Trim(myString)
        Exit Sub

End Sub


Public Sub ParseString2(ByVal myString As String, ByRef arOutput() As String, ByRef numElements As Integer)

    Dim counter As Integer
    counter = InStr(myString, "|")
    
'    If counter >= 0 Then
'        counter = InStr(myString, Chr(9))
'    End If
    
    Dim trimme As Integer
    
    Do While counter > 0
        
        numElements = numElements + 1
        ReDim Preserve arOutput(numElements)
        arOutput(numElements - 1) = Trim(Left(myString, counter - 1))
        ' trim the carriege return off the end of the string
        'arOutput(numElements - 1) = Mid(arOutput(numElements - 1), 1, Len(arOutput(numElements - 1)) - 1)
        'myString = Trim(Right(myString, Len(myString) - counter))
        myString = Trim(Right(myString, Len(myString) - counter))
        counter = InStr(myString, "|")
    
        ' if a <Enter> was found before a <TAB> then reduce counter size
'        If InStr(myString, vbCrLf) < counter Then
'            counter = InStr(myString, vbCrLf) + 1
'        End If

    Loop
    
        numElements = numElements + 1
        ReDim Preserve arOutput(numElements)
        arOutput(numElements - 1) = Trim(myString)
        Exit Sub

End Sub

