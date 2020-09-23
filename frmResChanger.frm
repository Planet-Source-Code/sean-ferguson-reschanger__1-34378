VERSION 5.00
Begin VB.Form frmResChanger 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ControlBox      =   0   'False
   Icon            =   "frmResChanger.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   150
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "frmResChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H4
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1
Private Const cBitsPerPixel = 12

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
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
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Sub Form_Load()
    On Error Resume Next
    Dim lngScreenHeight As Long
    Dim lngScreenWidth As Long
    Dim lngScreenDepth As Long
    Dim DevM As DEVMODE
    If InStr(1, Command(), " ") < 0 Then MsgBox "In order to use ResChanger, you must specify the command like this:" & vbCrLf & vbCrLf & "reschanger.exe [width] [height] [depth] [path] [parameters]" & vbCrLf & vbCrLf & "Please change the command line parameters and try again.", vbOKOnly + vbExclamation, "ResChanger": Unload Me: End
    cData = Split(Command() & " ", " ", 5)
    If UBound(cData) < 3 Then MsgBox "You must specify the command for ResChanger like this:" & vbCrLf & vbCrLf & "reschanger.exe [width] [height] [depth] [path] [parameters]" & vbCrLf & vbCrLf & "Please change the command line parameters and try again.", vbOKOnly + vbExclamation, "ResChanger": Unload Me: End
    lngScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    lngScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    lngScreenDepth = GetDeviceCaps(Me.hDC, cBitsPerPixel)
    erg& = EnumDisplaySettings(0&, 0&, DevM)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = CLng(cData(0))
    DevM.dmPelsHeight = CLng(cData(1))
    DevM.dmBitsPerPel = CInt(cData(2))
    erg& = ChangeDisplaySettings(DevM, CDS_TEST)
    Select Case erg&
        Case DISP_CHANGE_RESTART
            If MsgBox("In order for the display change to complete successfully, you must restart your computer. Do you wish to do so now?", vbYesNo + vbQuestion, "Restart Required") = vbYes Then
                erg& = ExitWindowsEx(EWX_REBOOT, 0&)
            Else
                DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
                DevM.dmPelsHeight = lngScreenHeight
                DevM.dmPelsWidth = lngScreenWidth
                DevM.dmBitsPerPel = CInt(lngScreenDepth)
                erg& = ChangeDisplaySettings(DevM, CDS_TEST)
            End If
        Case DISP_CHANGE_SUCCESSFUL
            erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
            ShellExecAndWait CStr(cData(3)), CStr(cData(4)), True, Me
            DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            DevM.dmPelsHeight = lngScreenHeight
            DevM.dmPelsWidth = lngScreenWidth
            DevM.dmBitsPerPel = CInt(lngScreenDepth)
            erg& = ChangeDisplaySettings(DevM, CDS_TEST)
            erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
        Case Else
            MsgBox "The display mode " & cData(0) & "x" & cData(1) & "x" & cData(2) & " is not supported by your system.", vbOKOnly + vbExclamation, "Error"
    End Select
    Unload Me
    End
End Sub
