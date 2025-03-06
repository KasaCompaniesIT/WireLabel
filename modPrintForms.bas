Attribute VB_Name = "modPrintForms"
Option Explicit

Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" (ByVal hPrinter As Long, ByVal Level As Long, ByRef pForm As Any, ByVal cbBuf As Long, ByRef pcbNeeded As Long, ByRef pcReturned As Long) As Long
Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long
'Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" (ByVal hPrinter As Long, ByVal pFormName As String) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) As Long
Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As Long, lpInitData As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByRef lpString2 As Long) As Long
    
' Optional functions not used, but may be useful.
'Public Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, pcbNeeded As Long) As Long
'Public Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" (ByVal hPrinter As Long, ByVal pFormName As String, ByVal Level As Long, pForm As Byte) As Long

' Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32
Public Const DM_FORMNAME As Long = &H10000
Public Const DM_ORIENTATION = &H1&

' Constants for PRINTER_DEFAULTS.DesiredAccess
'Public Const PRINTER_ACCESS_ADMINISTER = &H4
'Public Const PRINTER_ACCESS_USE = &H8
'Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
'Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

  
' Constants for DocumentProperties() call
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY

' Custom constants for this sample's SelectForm function
Public Const FORM_NOT_SELECTED = 0
Public Const FORM_SELECTED = 1
Public Const FORM_ADDED = 2

Public Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SIZEL
        cx As Long
        cy As Long
End Type

'Public Type SECURITY_DESCRIPTOR
'        Revision As Byte
'        Sbz1 As Byte
'        Control As Long
'        Owner As Long
'        Group As Long
'        Sacl As Long  ' ACL
'        Dacl As Long  ' ACL
'End Type


' The two definitions for FORM_INFO_1 make the coding easier.
Public Type FORM_INFO_1
        Flags As Long
        pName As Long   ' String
        Size As SIZEL
        ImageableArea As RECTL
End Type

Public Type sFORM_INFO_1
        Flags As Long
        pName As String
        Size As SIZEL
        ImageableArea As RECTL
End Type

Public Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

'Public Type PRINTER_DEFAULTS
'        pDatatype As String
'        pDevMode As Long    ' DEVMODE
'        DesiredAccess As Long
'End Type


'Public Type PRINTER_INFO_2
'        pServerName As String
'        pPrinterName As String
'        pShareName As String
'        pPortName As String
'        pDriverName As String
'        pComment As String
'        pLocation As String
'        pDevMode As DEVMODE
'        pSepFile As String
'        pPrintProcessor As String
'        pDatatype As String
'        pParameters As String
'        pSecurityDescriptor As SECURITY_DESCRIPTOR
'        Attributes As Long
'        Priority As Long
'        DefaultPriority As Long
'        StartTime As Long
'        UntilTime As Long
'        Status As Long
'        cJobs As Long
'        AveragePPM As Long
'End Type

Public Function GetFormName(ByVal PrinterHandle As Long, FormSize As SIZEL, FormName As String) As Integer
    ' Finds a form based on the size of the form
    
    Dim lngNumForms As Long
    Dim lngTemp As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1           ' Working FI1 array
    Dim bytTemp() As Byte                  ' Temp FI1 array
    Dim lngBytesNeeded As Long
    Dim lngRetVal As Long
    
    FormName = vbNullString
    GetFormName = 0
    ReDim aFI1(1)
    
    ' First call retrieves the lngBytesNeeded.
    lngRetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, lngBytesNeeded, lngNumForms)
    ReDim bytTemp(lngBytesNeeded)
    ReDim aFI1(lngBytesNeeded / Len(FI1))
    
    ' Second call actually enumerates the supported forms.
    lngRetVal = EnumForms(PrinterHandle, 1, bytTemp(0), lngBytesNeeded, lngBytesNeeded, lngNumForms)
    Call CopyMemory(aFI1(0), bytTemp(0), lngBytesNeeded)
    
    For lngTemp = 0 To lngNumForms - 1
        With aFI1(lngTemp)
            If .Size.cx = FormSize.cx And .Size.cy = FormSize.cy Then
               ' Found the desired form
                FormName = PtrCtoVbString(.pName)
                GetFormName = lngTemp + 1
                Exit For
            End If
        End With
    Next lngTemp
End Function

Public Function AddNewForm(PrinterHandle As Long, FormSize As SIZEL, FormName As String) As String
' This routine will add a form to the printer

    Dim FI1 As sFORM_INFO_1
    Dim aFI1() As Byte
    Dim lngRetVal As Long
    
    With FI1
        .Flags = 0
        .pName = FormName
        
        With .Size
            .cx = FormSize.cx
            .cy = FormSize.cy
        End With
        
        With .ImageableArea
            .Left = 0
            .Top = 0
            .Right = FI1.Size.cx
            .Bottom = FI1.Size.cy
        End With
    End With
    
    ReDim aFI1(Len(FI1))
    Call CopyMemory(aFI1(0), FI1, Len(FI1))
    
    lngRetVal = AddForm(PrinterHandle, 1, aFI1(0))
    
    If lngRetVal = 0 Then
        If Err.LastDllError = 5 Then
            MsgBox "You do not have permissions to add a form to " & _
               Printer.DeviceName, vbExclamation, "Access Denied!"
        Else
            MsgBox "Error: " & CStr(Err.LastDllError) & " Error Adding Form"
        End If
        
        AddNewForm = "None"
    Else
        AddNewForm = FI1.pName
    End If
End Function

Public Function PtrCtoVbString(ByVal Add As Long) As String
    ' Return the string specified by the pointer to the string
    
    Dim strTemp As String * 512
    
    lstrcpy strTemp, ByVal Add
    
'    If InStr(LCase(strTemp), "kasa") > 0 Then
'        Debug.Assert False
'    End If
    
    If (InStr(1, strTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(strTemp, InStr(1, strTemp, Chr(0)) - 1)
    End If
'    Debug.Print PtrCtoVbString
End Function

Public Function SelectForm(FormName As String, ByVal MyhWnd As Long, Optional PaperWidth As Double = 4, Optional PaperHeight As Double = 4) As Integer
    Dim nSize As Long           ' Size of DEVMODE
    Dim pDevMode As DEVMODE
    Dim PrinterHandle As Long   ' Handle to printer
    Dim hPrtDC As Long          ' Handle to Printer DC
    Dim PrinterName As String
    Dim aDevMode() As Byte      ' Working DEVMODE
    Dim FormSize As SIZEL
    
    PrinterName = Printer.DeviceName  ' Current printer
    hPrtDC = Printer.hdc              ' hDC for current Printer
    SelectForm = FORM_NOT_SELECTED    ' Set for failure unless reset in code.
    
    ' Get a handle to the printer.
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ' Retrieve the size of the DEVMODE.
        nSize = DocumentProperties(MyhWnd, PrinterHandle, PrinterName, 0&, _
                0&, 0&)
        ' Reserve memory for the actual size of the DEVMODE.
        ReDim aDevMode(1 To nSize)
    
        ' Fill the DEVMODE from the printer.
        nSize = DocumentProperties(MyhWnd, PrinterHandle, PrinterName, _
                aDevMode(1), 0&, DM_OUT_BUFFER)
        ' Copy the Public (predefined) portion of the DEVMODE.
        Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))
        
        If Not FormExists(PrinterName, FormName) Then
            ' Set the desired size of the form needed.
            With FormSize   ' Given in inches and converted to thousandths of millimeters
                .cx = PaperWidth * 25400    ' width
                .cy = PaperHeight * 25400   ' height
            End With
            
            '200510 N.F.
            '   Handle the case where no form was added
            If AddNewForm(PrinterHandle, FormSize, FormName) = "None" Then
                ClosePrinter (PrinterHandle)
                SelectForm = FORM_NOT_SELECTED   ' Selection Failed!
                Exit Function
            End If
            
            ' Check to make sure new form was added OK
            If GetFormName(PrinterHandle, FormSize, FormName) = 0 Then
                ClosePrinter (PrinterHandle)
                SelectForm = FORM_NOT_SELECTED   ' Selection Failed!
                Exit Function
            Else
                SelectForm = FORM_ADDED  ' Form Added, Selection succeeded!
            End If
        End If
        
        ' Change the appropriate member in the DevMode.
        ' In this case, you want to change the form name.
        pDevMode.dmFormName = FormName & Chr(0)  ' Must be NULL terminated!
        ' Set the dmFields bit flag to indicate what you are changing.
        pDevMode.dmFields = DM_FORMNAME
    
        ' Copy your changes back, then update DEVMODE.
        Call CopyMemory(aDevMode(1), pDevMode, Len(pDevMode))
        nSize = DocumentProperties(MyhWnd, PrinterHandle, PrinterName, _
                aDevMode(1), aDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)
    
        nSize = ResetDC(hPrtDC, aDevMode(1))   ' Reset the DEVMODE for the DC.
    
        ' Close the handle when you are finished with it.
        ClosePrinter (PrinterHandle)
        ' Selection Succeeded! But was Form Added?
        If SelectForm <> FORM_ADDED Then SelectForm = FORM_SELECTED
    Else
        SelectForm = FORM_NOT_SELECTED   ' Selection Failed!
    End If
End Function

' This routine determines if a Printer Form exists by enumerating all Form names
'   for the PrinterName specified.
Public Function FormExists(ByVal PrinterName As String, FormName As String) As Boolean
    
    Dim PrinterHandle As Long   ' Handle to printer
    Dim lngNumForms As Long
    Dim lngTemp As Long
    Dim FI1 As FORM_INFO_1
    
    'Array of form information
    Dim aFI1() As FORM_INFO_1           ' Working FI1 array
    Dim bytTemp() As Byte                  ' Temp FI1 array
    Dim lngBytesNeeded As Long
    Dim lngRetVal As Long
    
    FormExists = False
    
    ' Get a handle to the printer.
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ReDim aFI1(1)
        ' First call retrieves the lngBytesNeeded.
        lngRetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, lngBytesNeeded, lngNumForms)
        ReDim bytTemp(lngBytesNeeded)
        ReDim aFI1(lngBytesNeeded / Len(FI1))
        
        ' Second call actually enumerates the supported forms.
        lngRetVal = EnumForms(PrinterHandle, 1, bytTemp(0), lngBytesNeeded, lngBytesNeeded, lngNumForms)
        CopyMemory aFI1(0), bytTemp(0), lngBytesNeeded
        For lngTemp = 0 To lngNumForms - 1
            With aFI1(lngTemp)
                If PtrCtoVbString(.pName) = FormName Then
                    FormExists = True
                    Debug.Print "Found Form : " & FormName
                    Exit For
                End If
            End With
        Next
        
        ClosePrinter (PrinterHandle)
    End If
End Function

