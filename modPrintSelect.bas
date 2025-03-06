Attribute VB_Name = "modPrintSelect"
Option Explicit

'---------------------------------------------
' Function SelectPrinter
'
' Created by Michael_Flagler / Kasa Industrial Controls, Inc.
' Date: 03-10-2004    Time: 15:55
'
' Requires:  vbprndlg.dll file to be referenced (Microsoft VB Printer Dialog)
'
' Purpose: To display and select the current printer used
'          and all the settings for that printer used by
'          the Printer object built into VB
'
'    blnSoftwareCopies:  True for software controlled copies
'                        False for printer controlled copies
'
'    lngCopies = 1:      Returns number of copies for software controlled copies
'
'    Return value:       Returns True if Print was selected, False otherwise (Cancelled)
'
'---------------------------------------------

Public Function SelectPrinter(frmParentHwnd As Long, blnSoftwareCopies As Boolean, Optional lngCopies As Long = 1, Optional DisableSelection As Boolean = True, Optional PrintSelection As Boolean = False) As Boolean
    Dim pPrinter As Printer
    Dim cPrintDlg As VBPrnDlgLib.PrinterDlg
    
    Set cPrintDlg = New VBPrnDlgLib.PrinterDlg
    
    With cPrintDlg
        .PrinterName = Printer.DeviceName
        .DriverName = Printer.DriverName
        .Port = Printer.Port
        .PaperBin = Printer.PaperBin
        .PaperSize = Printer.PaperSize
        .Flags = cdlPDHidePrintToFile Or (cdlPDNoSelection And DisableSelection) Or IIf(blnSoftwareCopies, 0, cdlPDUseDevModeCopies)
        Printer.TrackDefault = False
        If Not .ShowPrinter(frmParentHwnd) Then
            SelectPrinter = False
            Exit Function
        End If
        
        For Each pPrinter In Printers
            If UCase$(pPrinter.DeviceName) = UCase$(.PrinterName) Then
                Set Printer = pPrinter
            End If
        Next
        
        PrintSelection = (.Flags And cdlPDSelection)
        
        If blnSoftwareCopies Then
            lngCopies = .Copies
        Else
            lngCopies = 1
            Printer.Copies = .Copies
        End If
        Printer.Orientation = .Orientation
        Printer.ColorMode = .ColorMode
        Printer.Duplex = .Duplex
        Printer.PaperBin = .PaperBin
        Printer.PrintQuality = .PrintQuality
    End With
    
    SelectPrinter = True
    
    Set cPrintDlg = Nothing
End Function
