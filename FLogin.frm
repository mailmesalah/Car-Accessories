VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   Icon            =   "FLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FLogin.frx":628A
   ScaleHeight     =   2910
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CLogin 
      Height          =   505
      Left            =   2435
      Picture         =   "FLogin.frx":204ECC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1745
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   4475
      Picture         =   "FLogin.frx":20732E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1745
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      Height          =   2895
      Left            =   15
      Top             =   0
      Width           =   6330
   End
   Begin MSForms.TextBox TUsername 
      Height          =   360
      Left            =   3105
      TabIndex        =   0
      Top             =   480
      Width           =   2850
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5027;635"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TPassword 
      Height          =   345
      Left            =   3105
      TabIndex        =   1
      Top             =   930
      Width           =   2850
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5027;609"
      PasswordChar    =   42
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   1320
      TabIndex        =   5
      Top             =   495
      Width           =   1830
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Username"
      Size            =   "3228;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Left            =   1335
      TabIndex        =   4
      Top             =   915
      Width           =   1830
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Password"
      Size            =   "3228;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CLogin_Click()
Dim rs As Recordset

    Set rs = db.OpenRecordset("Select * From Users Where (Users.Username='" & Trim(TUsername.Text) & "') And (Users.Password='" & TPassword.Text & "')")
    If rs.RecordCount = 1 Then
        sCurrentUserCode = "" & rs!Code
        sCurrentUsername = "" & rs!UserName
        FMain.Show
        rs.Close
        Unload Me
    Else
        TUsername.SetFocus
        rs.Close
    End If
    'CREATES REPORT FOLDER IF NOT EXIST
    arrangeFoldersAndFiles (isLoadingFirstTime)
End Sub

Private Sub Form_Initialize()
    InitCommonControls
    initialisePublicVariables
    setDefaultParentIDAndAccountCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CLogin_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    'loadGeneralDetails
End Sub

Private Sub TPassword_GotFocus()
    TPassword.SelStart = 0
    TPassword.SelLength = Len(TPassword.Text)
End Sub

Private Sub TUsername_GotFocus()
    TUsername.SelStart = 0
    TUsername.SelLength = Len(TUsername.Text)
End Sub


