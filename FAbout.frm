VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Us"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FAbout.frx":0000
   ScaleHeight     =   5115
   ScaleWidth      =   7995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CThanks 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6375
      OLEDropMode     =   1  'Manual
      Picture         =   "FAbout.frx":85382
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1245
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CThanks_Click()
    Unload Me
End Sub

