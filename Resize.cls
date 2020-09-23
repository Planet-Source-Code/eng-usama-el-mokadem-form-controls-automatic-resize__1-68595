VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsResize: Resize Class
'
' Name: clsResize
' Description: Resize Class
' Version: 1.00
' Date: 01 May 2007
' Last update: 01 May 2007
' Author: Eng. Usama El-Mokadem: musama@hotmail.com - ©1992-2007
'
' CONTACT INFORMATION:
' Eng. Usama El-Mokadem
' Email: musama@hotmail.com
' Web: http://musama.tripod.com
' Mobile: 0020 10 1289308
' Egypt
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private hOriginal As Long
Private wOriginal As Long
Private ifrm As Form

Public Sub FormInitialize(frm As Form)
    Set ifrm = frm
    Call SetOriginal
End Sub

Public Sub FormResize()
    Dim ctrl As Control
    Dim hFactor As Double
    Dim wFactor As Double

    hFactor = ifrm.Height / hOriginal
    wFactor = ifrm.Width / wOriginal

    On Local Error Resume Next
    For Each ctrl In ifrm.Controls
        ctrl.Height = ctrl.Height * hFactor
        ctrl.Width = ctrl.Width * wFactor

        ctrl.Top = ctrl.Top * hFactor
        ctrl.Left = ctrl.Left * wFactor
    Next
    On Local Error GoTo 0
    Call SetOriginal
End Sub

Private Sub SetOriginal()
    hOriginal = ifrm.Height
    wOriginal = ifrm.Width
End Sub
