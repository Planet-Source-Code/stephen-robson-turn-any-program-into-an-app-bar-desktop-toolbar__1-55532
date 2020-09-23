VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   2640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPosition 
      Caption         =   "Position"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton optLeft 
         Caption         =   "Left"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1000
      End
      Begin VB.OptionButton optRight 
         Caption         =   "Right"
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1000
      End
      Begin VB.OptionButton optTop 
         Caption         =   "Top"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1000
      End
      Begin VB.OptionButton optBottom 
         Caption         =   "Bottom"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   1000
      End
      Begin VB.OptionButton optFloat 
         Caption         =   "Floating"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1000
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox chkAutoHide 
      Caption         =   "Autohide"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This form allows the user to set the AppBar settings

Private Sub cmdOK_Click()
    If Me.chkAutoHide.Value = 1 Then AppBar.AutoHide = True Else AppBar.AutoHide = False
    If Me.optBottom.Value = True Then
        AppBar.Edge = abeBottom
    ElseIf Me.optTop.Value = True Then
        AppBar.Edge = abeTop
    ElseIf Me.optLeft.Value = True Then
        AppBar.Edge = abeLeft
    ElseIf Me.optRight.Value = True Then
        AppBar.Edge = abeRight
    ElseIf Me.optFloat.Value = True Then
        AppBar.Edge = abeFloat
    End If
    AppBar.UpdateBar
    Unload Me
End Sub

Private Sub Form_Load()
'Display the current settings
    If AppBar.AutoHide = True Then Me.chkAutoHide.Value = 1
    If AppBar.Edge = abeBottom Then
        Me.optBottom.Value = True
    ElseIf AppBar.Edge = abeTop Then
        Me.optTop.Value = True
    ElseIf AppBar.Edge = abeLeft Then
        Me.optLeft.Value = True
    ElseIf AppBar.Edge = abeRight Then
        Me.optRight.Value = True
    ElseIf AppBar.Edge = abeFloat Then
        Me.optFloat.Value = True
    End If
End Sub
