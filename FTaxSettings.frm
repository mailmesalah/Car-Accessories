VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FTaxSettings 
   Caption         =   "Tax Settings"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   Picture         =   "FTaxSettings.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CClose 
      Height          =   500
      Left            =   3195
      Picture         =   "FTaxSettings.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2115
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   500
      Left            =   1635
      Picture         =   "FTaxSettings.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2115
      Width           =   1365
   End
   Begin MSForms.Label Label2 
      Height          =   405
      Left            =   315
      TabIndex        =   5
      Top             =   780
      Width           =   1230
      VariousPropertyBits=   8388627
      Caption         =   "Purchase Tax"
      Size            =   "2170;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TUnitName 
      Height          =   405
      Left            =   1725
      TabIndex        =   4
      Top             =   735
      Width           =   3000
      VariousPropertyBits=   746604571
      MaxLength       =   8
      BorderStyle     =   1
      Size            =   "5292;714"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TValue 
      Height          =   405
      Left            =   1710
      TabIndex        =   3
      Top             =   1230
      Width           =   3000
      VariousPropertyBits=   746604571
      MaxLength       =   5
      BorderStyle     =   1
      Size            =   "5292;714"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   765
      TabIndex        =   2
      Top             =   1260
      Width           =   780
      VariousPropertyBits=   8388627
      Caption         =   "Sale Tax"
      Size            =   "1376;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FTaxSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
