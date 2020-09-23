VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "ReSizable Controls"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   5040
      TabIndex        =   12
      Top             =   3840
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   600
      TabIndex        =   11
      Top             =   3360
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   1680
      Width           =   2895
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   435
      Left            =   3240
      TabIndex        =   8
      Top             =   3240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   767
      _Version        =   327682
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   2160
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   3240
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   4683
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3015
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Eng. Usama El-Mokadem"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "http://musama.tripod.com"
      Top             =   6600
      Width           =   5055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Resize: Resizable form
'
' Name: Resize
' Description: Resizable form
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

Private CtrlsReSize As clsResize

Private Sub Form_Initialize()
    Set CtrlsReSize = New clsResize
    Call CtrlsReSize.FormInitialize(Me)
End Sub

Private Sub Form_Terminate()
    Set CtrlsReSize = Nothing
End Sub

Private Sub Form_Resize()
    Call CtrlsReSize.FormResize
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


