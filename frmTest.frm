VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProgressXp Demo"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   70.379
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   149.49
   StartUpPosition =   3  'Windows Default
   Begin Projekt1.XPProgressBar ProgressXp1 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      Value           =   20
      Step_Length     =   8
      BackColor       =   16777215
      BarColor        =   3724597
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show only full items"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   3600
   End
   Begin Projekt1.XPProgressBar ProgressXp1 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
      Value           =   20
      Step_Length     =   8
      Seperator_Width =   0
      BackColor       =   16777215
      BarColor        =   3724597
   End
   Begin Projekt1.XPProgressBar ProgressXp1 
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1720
      Value           =   20
      Step_Length     =   20
      Seperator_Width =   5
      BackColor       =   16777215
      BarColor        =   3724597
   End
   Begin Projekt1.XPProgressBar ProgressXp1 
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      Value           =   20
      Step_Length     =   8
      BackColor       =   16777215
      BarColor        =   12937777
   End
   Begin Projekt1.XPProgressBar ProgressXp1 
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      Value           =   20
      Step_Length     =   8
      BackColor       =   16777215
      BarColor        =   8421631
   End
   Begin Projekt1.XPProgressBar ProgressXp1 
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      Value           =   20
      Step_Length     =   8
      BackColor       =   0
      BarColor        =   65535
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
For i = 0 To ProgressXp1.Count - 1
ProgressXp1(i).DrawOnlyFullItems = Check1.Value ' --> Draw full Items
Next i
End Sub

Private Sub Timer1_Timer()
Static a As Double
For i = 0 To ProgressXp1.Count - 1
ProgressXp1(i).Value = a
Next i

Me.Caption = "ProgressXp Demo  " & CStr(Int(a)) & " %"
a = a + 0.1
If a > 100 Then a = 0
End Sub
