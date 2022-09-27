VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FPurchase 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Purchase"
   ClientHeight    =   8070
   ClientLeft      =   8295
   ClientTop       =   450
   ClientWidth     =   12540
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FPurchase.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   12540
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CDelete 
      Height          =   500
      Left            =   4800
      Picture         =   "FPurchase.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   500
      Left            =   10695
      Picture         =   "FPurchase.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7425
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   500
      Left            =   9255
      Picture         =   "FPurchase.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7425
      Width           =   1365
   End
   Begin VB.CommandButton CNew 
      Height          =   500
      Left            =   375
      Picture         =   "FPurchase.frx":205968
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7410
      Width           =   1365
   End
   Begin VB.CommandButton CClear 
      Height          =   500
      Left            =   3240
      Picture         =   "FPurchase.frx":207DCA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6735
      Width           =   1365
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   500
      Left            =   1800
      Picture         =   "FPurchase.frx":20A22C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6735
      Width           =   1365
   End
   Begin VB.CommandButton CAddItem 
      Height          =   500
      Left            =   360
      Picture         =   "FPurchase.frx":20C68E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6735
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3240
      Left            =   135
      TabIndex        =   17
      Top             =   1620
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   5715
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   420
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   45678595
      CurrentDate     =   40544
   End
   Begin MSForms.TextBox TDiscount 
      Height          =   375
      Left            =   7395
      TabIndex        =   50
      Top             =   6195
      Width           =   1605
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "2831;661"
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TPayment 
      Height          =   375
      Left            =   7395
      TabIndex        =   51
      Top             =   6555
      Width           =   1605
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "2831;661"
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label11 
      Height          =   330
      Left            =   8070
      TabIndex        =   48
      Top             =   1260
      Width           =   825
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "MRP"
      Size            =   "1455;582"
      BorderColor     =   -2147483641
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TMRP 
      Height          =   435
      Left            =   8190
      TabIndex        =   8
      Top             =   4920
      Width           =   990
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "1746;767"
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label LMFRShortName 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   4590
      TabIndex        =   46
      Top             =   5640
      Width           =   1650
   End
   Begin MSForms.Label LMRP 
      Height          =   345
      Left            =   3030
      TabIndex        =   45
      Top             =   6015
      Width           =   1110
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Size            =   "1958;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   375
      TabIndex        =   44
      Top             =   5640
      Width           =   1470
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Manufacturer"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LManufacturer 
      Height          =   345
      Left            =   1860
      TabIndex        =   43
      Top             =   5640
      Width           =   1800
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Size            =   "3175;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label10 
      Height          =   330
      Left            =   6015
      TabIndex        =   42
      Top             =   1245
      Width           =   960
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Rate"
      Size            =   "1693;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LUnit 
      Height          =   330
      Left            =   1935
      TabIndex        =   41
      Top             =   6390
      Width           =   600
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "1058;582"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TNarration 
      Height          =   420
      Left            =   1500
      TabIndex        =   2
      Top             =   630
      Width           =   3180
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "5609;741"
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   90
      TabIndex        =   40
      Top             =   705
      Width           =   1335
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "2355;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LCurrentStock 
      Height          =   345
      Left            =   1845
      TabIndex        =   39
      Top             =   6030
      Width           =   1365
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Size            =   "2408;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   375
      TabIndex        =   38
      Top             =   6000
      Width           =   1470
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Current Stock"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label19 
      Height          =   330
      Left            =   11070
      TabIndex        =   37
      Top             =   1245
      Width           =   1380
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2434;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label16 
      Height          =   330
      Left            =   6960
      TabIndex        =   36
      Top             =   1230
      Width           =   885
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Qty"
      Size            =   "1561;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1545
      TabIndex        =   35
      Top             =   1245
      Width           =   3480
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Item"
      Size            =   "6138;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   195
      TabIndex        =   34
      Top             =   1245
      Width           =   555
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Sl No"
      Size            =   "979;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LGrandAmount 
      Height          =   555
      Left            =   9255
      TabIndex        =   33
      Top             =   6405
      Width           =   3075
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Grand Amount"
      Size            =   "5424;979"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TAddress 
      Height          =   420
      Left            =   9150
      TabIndex        =   4
      Top             =   645
      Width           =   3210
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "5662;741"
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   420
      Left            =   1485
      TabIndex        =   0
      Top             =   120
      Width           =   1590
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "2805;741"
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   105
      TabIndex        =   32
      Top             =   150
      Width           =   465
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "820;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LPayment 
      Height          =   375
      Left            =   6420
      TabIndex        =   31
      Top             =   6585
      Width           =   855
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Payment"
      Size            =   "1508;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBalancelb 
      Height          =   375
      Left            =   6510
      TabIndex        =   30
      Top             =   6960
      Width           =   720
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "1270;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBalance 
      Height          =   375
      Left            =   7395
      TabIndex        =   29
      Top             =   6900
      Width           =   1605
      BackColor       =   16777215
      Size            =   "2831;661"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   4260
      Left            =   120
      Top             =   1155
      Width           =   12345
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   135
      X2              =   12500
      Y1              =   4845
      Y2              =   4860
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   10005
      TabIndex        =   28
      Top             =   1245
      Width           =   1185
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Tax Amount"
      Size            =   "2090;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   9015
      TabIndex        =   27
      Top             =   1245
      Width           =   1095
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "GrossValue"
      Size            =   "1931;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTaxAmount 
      Height          =   375
      Left            =   10110
      TabIndex        =   26
      Top             =   5670
      Width           =   840
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Total Tax"
      Size            =   "1482;661"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label15 
      Height          =   375
      Left            =   6420
      TabIndex        =   25
      Top             =   6225
      Width           =   855
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Discount"
      Size            =   "1508;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LGrossValue 
      Height          =   375
      Left            =   8325
      TabIndex        =   24
      Top             =   5670
      Width           =   1020
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Total Gross"
      Size            =   "1799;661"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LNetValue 
      Height          =   375
      Left            =   11025
      TabIndex        =   23
      Top             =   5655
      Width           =   1290
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Total Tax"
      Size            =   "2275;661"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LSlNo 
      Height          =   435
      Left            =   255
      TabIndex        =   22
      Top             =   4920
      Width           =   555
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "SLNo"
      Size            =   "979;767"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoItem 
      Height          =   435
      Left            =   1050
      TabIndex        =   5
      Top             =   4920
      Width           =   4350
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "7673;767"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TQuantity 
      Height          =   435
      Left            =   7185
      TabIndex        =   7
      Top             =   4920
      Width           =   915
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "1614;767"
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalAmount 
      Height          =   405
      Left            =   11340
      TabIndex        =   21
      Top             =   4920
      Width           =   1140
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2011;714"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox TRate 
      Height          =   435
      Left            =   6105
      TabIndex        =   6
      Top             =   4920
      Width           =   990
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "1746;767"
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label12 
      Height          =   420
      Left            =   615
      TabIndex        =   20
      Top             =   6345
      Width           =   1470
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTax 
      Height          =   345
      Left            =   5625
      TabIndex        =   19
      Top             =   4920
      Width           =   450
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Tax"
      Size            =   "794;609"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   5520
      TabIndex        =   18
      Top             =   1245
      Width           =   585
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Tax"
      Size            =   "1032;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   7770
      TabIndex        =   10
      Top             =   255
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "Supplier"
      Size            =   "2355;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoSupplier 
      Height          =   420
      Left            =   9150
      TabIndex        =   3
      Top             =   195
      Width           =   3210
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "5662;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   495
      Index           =   120
      Left            =   135
      TabIndex        =   49
      Top             =   1135
      Width           =   12345
      BackColor       =   15724527
      Size            =   "21775;873"
      Picture         =   "FPurchase.frx":20EAF0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label20 
      Height          =   495
      Left            =   120
      TabIndex        =   47
      Top             =   1140
      Width           =   12345
      BackColor       =   128
      Size            =   "21775;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSupplierCode() As String, sSupplierAddress() As String, sAccountCode() As String
Dim sItemCode() As String, sBillingName() As String
Dim gSerialNo As Single, gItem As Single, gQuantity As Single, gTaxAmount As Single, gTax As Single, gBatch As Single, gWarranty As Single, gGrossValue As Single, gUnit As Single, gPurchaseRate As Single, gUnitQuantity As Single, gMRP As Single, gToTalAmount As Single, gBillingName As Single, gItemCode As Single, gMFRShortName As Single
Dim dUnitValue As Double

Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

    If CoItem.ListIndex = -1 Then
        MsgBox "Please Select a Item !", vbInformation
        CoItem.SetFocus
        Exit Sub
    End If
        
    If Val(TQuantity.Text) = 0 Then
        MsgBox "Please Enter Quantity greater than Zero !", vbInformation
        TQuantity.SetFocus
        Exit Sub
    End If
    
    If Val(TRate.Text) = 0 Then
        lYN = MsgBox("Rate given is Zero, Do you want to Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TRate.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(TMRP.Text) = 0 Or Val(TMRP.Text) <= Val(TRate.Text) Then
        lYN = MsgBox("MRP given is Incorrect, Do you want to Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TMRP.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(LSlNo.Caption) > MGrid.Rows Then 'Add
        MGrid.AddItem ""
        MGrid.TextMatrix(MGrid.Rows - 1, gSerialNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(MGrid.Rows - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gTax) = LTax.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gMRP) = Val(TMRP.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gPurchaseRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount) = Format(Val(MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue)) * Val(MGrid.TextMatrix(MGrid.Rows - 1, gTax)) / 100, "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gToTalAmount) = Format((Val(MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue)) + Val(MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount))), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gMFRShortName) = LMFRShortName.Caption & ""
        MGrid.TextMatrix(MGrid.Rows - 1, gUnitQuantity) = dUnitValue
        MGrid.TextMatrix(MGrid.Rows - 1, gBatch) = getUniversaloFor((TRate.Text) & (TMRP.Text))
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(r - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(r - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(r - 1, gBatch) = getUniversaloFor((TRate.Text) & (TMRP.Text))
        MGrid.TextMatrix(r - 1, gTax) = LTax.Caption
        MGrid.TextMatrix(r - 1, gMRP) = Val(TMRP.Text)
        MGrid.TextMatrix(r - 1, gPurchaseRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(r - 1, gGrossValue) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(r - 1, gTaxAmount) = Format(Val(MGrid.TextMatrix(r - 1, gGrossValue)) * Val(MGrid.TextMatrix(r - 1, gTax)) / 100, "0.00")
        MGrid.TextMatrix(r - 1, gToTalAmount) = Format((Val(MGrid.TextMatrix(r - 1, gGrossValue)) + Val(MGrid.TextMatrix(r - 1, gTaxAmount))), "0.00")
        MGrid.TextMatrix(r - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gMFRShortName) = LMFRShortName.Caption & ""
        MGrid.TextMatrix(MGrid.Rows - 1, gUnitQuantity) = dUnitValue
        
    End If
    clearEditControls
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    TPayment.Text = Format(getGrandTotal, "0.00")
    CoItem.SetFocus
End Sub


Private Sub CClear_Click()
    MGrid.Rows = 0
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    TPayment.Text = Format(getGrandTotal, "0.00")
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gItem = 1
    gTax = 2
    gBatch = 3
    gUnit = 4
    gPurchaseRate = 5
    gQuantity = 6
    gMRP = 7
    gGrossValue = 8
    gTaxAmount = 9
    gToTalAmount = 10
    gBillingName = 11
    gItemCode = 12
    gMFRShortName = 13
    gUnitQuantity = 14
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 15
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 890
    MGrid.ColWidth(gItem) = 4200
    MGrid.ColWidth(gQuantity) = 800
    MGrid.ColWidth(gTax) = 600
    MGrid.ColWidth(gBatch) = 0
    MGrid.ColWidth(gMRP) = 1000
    MGrid.ColWidth(gUnit) = 0
    MGrid.ColWidth(gPurchaseRate) = 1160
    MGrid.ColWidth(gGrossValue) = 1160
    MGrid.ColWidth(gToTalAmount) = 1160
    MGrid.ColWidth(gTaxAmount) = 1160
    MGrid.ColWidth(gBillingName) = 0
    MGrid.ColWidth(gItemCode) = 0
    MGrid.ColWidth(gMFRShortName) = 0
    MGrid.ColWidth(gUnitQuantity) = 0
    
    MGrid.ColAlignment(gItem) = vbLeftJustify
    MGrid.ColAlignment(gUnit) = vbLeftJustify
        
    MGrid.RowHeightMin = 350
    
End Sub

Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String
    
    Set rs = db.OpenRecordset("Select Max(Val( Transaction.TransactionNo)) As TNo From Transaction Where ( Transaction.TransactionType = 'P' )")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function

Private Sub getSupplier()

Dim rs As Recordset
    
    CoSupplier.Clear
    
    Set rs = db.OpenRecordset("Select SupplierMaster.SupplierCode,SupplierMaster.AccountCode,SupplierMaster.SupplierName,SupplierMaster.Address1,SupplierMaster.Address2,SupplierMaster.Address3 From SupplierMaster Where (SupplierMaster.Status = True) Order By SupplierMaster.SupplierName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    ReDim sSupplierCode(rs.RecordCount) As String
    ReDim sSupplierAddress(rs.RecordCount) As String
    ReDim sAccountCode(rs.RecordCount) As String
    While rs.EOF = False
        CoSupplier.AddItem "" & rs!SupplierName
        sSupplierCode(CoSupplier.ListCount) = "" & rs!SupplierCode
        sSupplierAddress(CoSupplier.ListCount) = "" & rs!Address1 & " " & rs!Address2 & " " & rs!Address3
        sAccountCode(CoSupplier.ListCount) = "" & rs!AccountCode
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getItem()
Dim rs As Recordset
    
    CoItem.Clear
    
    Set rs = db.OpenRecordset("Select ItemMaster.Code,ItemMaster.ItemName,ItemMaster.BillingName From ItemMaster Where (ItemMaster.Type = 'BItem' ) Order By ItemMaster.ItemName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sItemCode(rs.RecordCount + 1) As String
    ReDim sBillingName(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoItem.AddItem "" & rs!ItemName
        sItemCode(CoItem.ListCount) = "" & rs!Code
        sBillingName(CoItem.ListCount) = "" & rs!BillingName
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getItemDetails()
Dim rs As Recordset, r As Long
    If (CoItem.ListIndex = -1) Then
        LManufacturer.Caption = ""
        LUnit.Caption = ""
        dUnitValue = 0
        LTax.Caption = ""
        LCurrentStock.Caption = ""
        Exit Sub
    End If
    Set rs = db.OpenRecordset("Select ItemMaster.Tax,Manufacturer.ManufacturerName,Units.UnitName,Units.UnitValue,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('O','P','SR','SA') ) And (Transaction.ItemCode = ItemMaster.Code )) As InStock,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('S','PR') ) And (Transaction.ItemCode = ItemMaster.Code )) As OutStock From ItemMaster,Units,Manufacturer Where (ItemMaster.Code = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code = ItemMaster.ManufacturerCode )")
    If rs.RecordCount > 0 Then
        LManufacturer.Caption = "" & rs!ManufacturerName
        LUnit.Caption = "" & rs!UnitName
        dUnitValue = Val("" & rs!UnitValue)
        LCurrentStock.Caption = Val("" & rs!InStock) - Val("" & rs!OutStock)
        LTax.Caption = Val("" & rs!Tax)
    Else
        LManufacturer.Caption = ""
        LUnit.Caption = ""
        dUnitValue = 0
        LCurrentStock.Caption = ""
        LTax.Caption = ""
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("Select Sum(Transaction.Quantity) As Quantity From Transaction Where (Transaction.TransactionNo='" & Trim(TTransactionNo.Text) & "') And (Transaction.TransactionType='P') And (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "')")
    If rs.RecordCount > 0 Then
        LCurrentStock.Caption = Val("" & LCurrentStock.Caption) - Val("" & rs!Quantity)
    End If
    rs.Close
    
    r = 0
    While r < MGrid.Rows
        If (Trim(MGrid.TextMatrix(r, gItemCode)) = sItemCode(CoItem.ListIndex + 1)) Then
            LCurrentStock.Caption = Val("" & LCurrentStock.Caption) + (Val(MGrid.TextMatrix(r, gQuantity)) * Val(MGrid.TextMatrix(r, gUnitQuantity)))
        End If
        r = r + 1
    Wend
End Sub

Private Sub clearControls()
    
    'TTransactionNo.Text = getNewTransactionNo
    'DTPDate.Value = Date
    TNarration.Text = ""
    CoSupplier.Text = ""
    TAddress.Text = ""
    MGrid.Rows = 0
    TPayment.Text = ""
    LBalance.Caption = ""
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    TRate.Text = ""
    TMRP.Text = ""
    LNetValue.Caption = ""
    LGrossValue.Caption = ""
    LTax.Caption = ""
    LTotalAmount.Caption = ""
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    TDiscount.Text = 0#
    TPayment.Text = ""
    LBalance.Caption = ""
End Sub

Private Sub clearEditControls()
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    LTax.Caption = ""
    LTaxAmount.Caption = ""
    LGrossValue.Caption = ""
    TRate.Text = ""
    TMRP.Text = ""
    LTotalAmount.Caption = ""
    TDiscount.Text = 0#
    TPayment.Text = ""
    LBalance.Caption = ""
End Sub




Private Sub TDiscount_Change()
    LBalance.Caption = Format(Trim(LGrandAmount) - Val(TPayment.Text), "0.00")
End Sub

Private Sub TDiscount_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    If Val(LGrandAmount.Caption) < Val(TDiscount.Text) Then
        TDiscount.SetFocus
        Exit Sub
    ElseIf Val(TDiscount.Text) = 0 Then
        TPayment.Text = Format(LGrandAmount, "0.00")
        LBalance.Caption = 0#
    End If
    CSave.SetFocus
  End If
End Sub

Private Sub TDiscount_LostFocus()
LGrandAmount.Caption = Format(getGrandTotal - TDiscount.Text, "0.00")
TPayment.Text = Format(getGrandTotal - TDiscount.Text, "0.00")
End Sub
Private Function getGrandTotal() As Double
Dim dGrandTotal As Double, dTax As Double, dGrossValue As Double, r As Long
    
    r = 0
    dGrandTotal = 0
    dTax = 0
    dGrossValue = 0
    While r < MGrid.Rows
        dGrandTotal = dGrandTotal + Val(MGrid.TextMatrix(r, gToTalAmount))
        dTax = dTax + Val(MGrid.TextMatrix(r, gTaxAmount))
        dGrossValue = dGrossValue + Val(MGrid.TextMatrix(r, gGrossValue))
        r = r + 1
    Wend
    getGrandTotal = dGrandTotal
    LGrossValue.Caption = Format(dGrossValue, "0.00")
    LNetValue.Caption = Format(dGrandTotal, "0.00")
    LTaxAmount.Caption = Format(dTax, "0.00")
End Function

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean
    bFound = False
    If (MsgBox("Do you want to Delete the Bill ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'P' )")
        While rs.EOF = False
            bFound = True
            rs.Delete
            rs.MoveNext
        Wend
        rs.Close
        
        If bFound Then
            MsgBox "Successfully Deleted !", vbInformation
            deleteFromAccountRegister
            clearControls
            TTransactionNo.Text = getNewTransactionNo
        Else
            MsgBox "Bill Not Found !", vbInformation
        End If
    End If
End Sub
Private Sub deleteFromAccountRegister()
Dim rs As Recordset
    Set rs = db.OpenRecordset("Select * From AccountTransaction Where (AccountTransaction.Type = 'PU') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') And (AccountTransaction.InventoryBillNo='" & TTransactionNo.Text & "') And (AccountTransaction.InventoryType='P') ")
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub CNew_Click()
    clearControls
    TTransactionNo.Text = getNewTransactionNo
End Sub

Private Sub CoItem_Change()
    getItemDetails
End Sub

Private Sub CoItem_GotFocus()
    CoItem.SelStart = 0
    CoItem.SelLength = Len(CoItem.Text)
End Sub

Private Sub CoItem_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim r As Long
    If KeyCode = 113 Then
        FItemMaster.Show vbModal
        getItem
    End If
End Sub

Private Sub CoSupplier_Change()
    If CoSupplier.ListIndex <> -1 Then
        TAddress.Text = sSupplierAddress(CoSupplier.ListIndex + 1)
    Else
        TAddress.Text = ""
    End If
End Sub

Private Sub CoSupplier_GotFocus()
    CoSupplier.SelStart = 0
    CoSupplier.SelLength = Len(CoSupplier.Text)
End Sub

Private Sub CoSupplier_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim r As Long
    If KeyCode = 113 Then
        FSupplierRegister.Show vbModal
        getSupplier
    End If
End Sub

Private Sub CPrint_Click()
    'printOpeningStock
End Sub

Private Sub CRemoveItem_Click()
Dim r As Long
    If MGrid.Rows > 0 Then
        If MGrid.Rows = 1 Then
            MGrid.Rows = 0
            clearEditControls
        Else
            MGrid.RemoveItem (MGrid.Row)
            r = 0
            While r < MGrid.Rows
                MGrid.TextMatrix(r, gSerialNo) = r + 1
                r = r + 1
            Wend
            clearEditControls
        End If
        LGrandAmount.Caption = Format(getGrandTotal, "0.00")
        TPayment.Text = Format(getGrandTotal, "0.00")
    Else
    
    End If
End Sub

Private Sub CSave_Click()
Dim rs As Recordset
Dim r As Long, lYN As Long, sStatus As String

    If Val(TTransactionNo.Text) = 0 Then
        MsgBox "Please Enter Valid Transaction No !", vbInformation
        TTransactionNo.SetFocus
        Exit Sub
    End If
    
    If CoSupplier.ListIndex = -1 Then
        'lYN = MsgBox("Do you want to consider General Supplier !", vbDefaultButton2 Or vbYesNo)
        'If lYN = vbYes Then
        'Else
        '    CoSupplier.SetFocus
        '    Exit Sub
        'End If
        MsgBox "Please Select a Supplier !", vbInformation
        CoSupplier.SetFocus
        Exit Sub
    End If
    
    If MGrid.Rows = 0 Then
        MsgBox "No Items Entered !", vbInformation
        CoItem.SetFocus
        Exit Sub
    End If
    
    
    'SAVES DATA TO Transaction TABLE
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'P' )")
    If rs.RecordCount > 0 Then  'Edit
         
        'SAVES DATA TO TransactionRegister ReadyMade
        While rs.EOF = False
            rs.Delete
            rs.MoveNext
        Wend
    End If
    
    r = 0
    While r < MGrid.Rows
        rs.AddNew
        rs!TransactionNo = Trim(TTransactionNo.Text)
        rs!TransactionType = "P"
        rs!TransactionDate = DTPDate.Value
        rs!TransactionTime = Format(Time, "HH:MM AMPM")
        rs!Narration = Trim(TNarration.Text)
        rs!SupplierCode = IIf(CoSupplier.ListIndex = -1, "", sSupplierCode(CoSupplier.ListIndex + 1))
        rs!SupplierName = Trim(CoSupplier.Text)
        rs!SupplierAddress = Trim(TAddress.Text)
        rs!CustomerCode = ""
        rs!CustomerName = ""
        rs!CustomerAddress = ""
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!ItemCode = Trim(MGrid.TextMatrix(r, gItemCode))
        rs!Quantity = Val(MGrid.TextMatrix(r, gQuantity)) * MGrid.TextMatrix(r, gUnitQuantity)
        rs!PurchaseRate = Val(MGrid.TextMatrix(r, gPurchaseRate))
        rs!SaleRate = 0
        rs!Batch = Trim(MGrid.TextMatrix(r, gBatch))
        rs!MRP = Val(MGrid.TextMatrix(r, gMRP)) / (MGrid.TextMatrix(r, gUnitQuantity))
        rs!ReferenceNo = ""
        rs!ReferenceDate = Date
        rs!PurchasePayment = IIf(TPayment.Text = "", LGrandAmount.Caption, Val(TPayment.Text))
        rs!UnitQuantity = Val(MGrid.TextMatrix(r, gUnitQuantity))
        rs!UnitMRP = Val(MGrid.TextMatrix(r, gMRP))
        rs!Tax = Val(MGrid.TextMatrix(r, gTax))
        rs!PurchaseQuantity = Val(MGrid.TextMatrix(r, gQuantity))
        rs!Discount = IIf(Val(TDiscount.Text) = 0, "0", Val(TDiscount.Text))
        rs.Update
        
        r = r + 1
    Wend
    rs.Close
    addToAccountRegister
    MsgBox "Successfully Saved !", vbInformation
    clearControls
    TTransactionNo.Text = getNewTransactionNo
    TTransactionNo.SetFocus
End Sub


Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CClose_Click
    ElseIf (KeyCode = vbKeyP And ((Shift And 7) = 2)) Then
        CPrint_Click
    ElseIf (KeyCode = vbKeyA And ((Shift And 7) = 2)) Then
        CAddItem_Click
    ElseIf (KeyCode = vbKeyR And ((Shift And 7) = 2)) Then
        CRemoveItem_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClear_Click
    End If
End Sub

Private Sub Form_Load()
    DTPDate.Value = Date
    getSupplier
    getItem
    MGridInitialise
    clearControls
    TTransactionNo.Text = getNewTransactionNo
End Sub

Private Sub MGrid_Click()
Dim r As Long, i As Long

    If MGrid.Rows > 0 Then
        r = MGrid.Row
        LSlNo.Caption = Val(MGrid.TextMatrix(r, gSerialNo))
        CoItem.Text = Trim(MGrid.TextMatrix(r, gItem))
        TQuantity.Text = Val(MGrid.TextMatrix(r, gQuantity))
        LUnit.Caption = Trim(MGrid.TextMatrix(r, gUnit))
        LTax.Caption = Trim(MGrid.TextMatrix(r, gTax))
        TRate.Text = Val(MGrid.TextMatrix(r, gPurchaseRate))
        TMRP.Text = Val(MGrid.TextMatrix(r, gMRP))
        LTotalAmount.Caption = Val(MGrid.TextMatrix(r, gToTalAmount))
    Else
    End If
End Sub

Private Sub MGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub getTotal()
    LTotalAmount.Caption = Val(TRate.Text) * Val(TQuantity.Text)
End Sub

Private Sub TAddress_GotFocus()
    TAddress.SelStart = 0
    TAddress.SelLength = Len(TAddress.Text)
End Sub



Private Sub TextBox1_Change()

End Sub



Private Sub TMRP_GotFocus()
    TMRP.SelStart = 0
    TMRP.SelLength = Len(TMRP.Text)
End Sub





Private Sub TPayment_Change()
    LBalance.Caption = Format(Trim(LGrandAmount) - Val(TPayment.Text), "0.00")
End Sub

Private Sub TPayment_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        CSave.SetFocus
    End If
End Sub

Private Sub TQuantity_Change()
    getTotal
End Sub

Private Sub TQuantity_GotFocus()
    TQuantity.SelStart = 0
    TQuantity.SelLength = Len(TQuantity.Text)
End Sub

Private Sub TRate_Change()
    getTotal
End Sub

Private Sub TRate_GotFocus()
    TRate.SelStart = 0
    TRate.SelLength = Len(TRate.Text)
End Sub

Private Sub TTransactionNo_GotFocus()
    TTransactionNo.SelStart = 0
    TTransactionNo.SelLength = Len(TTransactionNo.Text)
End Sub

Private Sub TTransactionNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        clearControls
        getTransactionDetails
    End If
End Sub

Public Sub getTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.BillingName,ItemMaster.Code,Transaction.Tax as ItemTax,Units.UnitName,Transaction.* From ItemMaster,Transaction,Units Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'P' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!TransactionDate
        CoSupplier.Text = "" & rs!SupplierName
        TAddress.Text = "" & rs!SupplierAddress
        TNarration.Text = "" & rs!Narration
        TPayment.Text = "" & rs!PurchasePayment
        TDiscount.Text = Format("" & rs!Discount, "0.00")
        
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gQuantity) = "" & rs!PurchaseQuantity
            MGrid.TextMatrix(r, gBatch) = "" & rs!Batch
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gTax) = "" & rs!ItemTax
            MGrid.TextMatrix(r, gPurchaseRate) = Format("" & rs!PurchaseRate, "0.00")
            MGrid.TextMatrix(r, gGrossValue) = Format(Val("" & rs!PurchaseQuantity) * Val("" & rs!PurchaseRate), "0.00")
            MGrid.TextMatrix(r, gTaxAmount) = Format(Val("" & rs!PurchaseQuantity) * Val("" & rs!PurchaseRate) * rs!ItemTax / 100, "0.00")
            MGrid.TextMatrix(r, gToTalAmount) = Format(Val(MGrid.TextMatrix(r, gGrossValue) + Val(MGrid.TextMatrix(r, gTaxAmount))), "0.00")
            MGrid.TextMatrix(r, gMRP) = Format("" & rs!UnitMRP, "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gUnitQuantity) = "" & rs!UnitQuantity
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
    
    
    
    LSlNo.Caption = MGrid.Rows + 1
    LGrandAmount.Caption = Format(getGrandTotal - Val(TDiscount.Text), "0.00")
    LBalance.Caption = Format((Val(LGrandAmount.Caption) - Val(TPayment.Text)), "0.00")
End Sub

Private Sub addToAccountRegister()
Dim rs As Recordset, sTransactionNo As String, LSerialNo As Long
        
    sTransactionNo = Val(TTransactionNo.Text)
    
    Set rs = db.OpenRecordset("Select * From AccountTransaction Where (AccountTransaction.Type = 'PU') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') And (AccountTransaction.InventoryBillNo='" & Trim(TTransactionNo.Text) & "') And (AccountTransaction.InventoryType='P') ")
    
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    LSerialNo = 1
    'Supplier To Purchase Account
    If (Val(LGrandAmount.Caption)) > 0 Then
         rs.AddNew
         rs!BillNo = Trim(TTransactionNo.Text)
         rs!Type = "PU"
         rs!AccountCode = IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID)
         rs!AddedDate = DTPDate.Value
         rs!EditedDate = DTPDate.Value
         rs!Credit = Val(LGrandAmount.Caption)
         rs!Debit = 0
         rs!Narration = Trim(TNarration.Text)
         rs!AddedBy = sCurrentUserCode
         rs!EditedBy = sCurrentUserCode
         rs!SerialNo = LSerialNo
         rs!FinancialCode = getFinancialCode(DTPDate.Value)
         rs!InventoryBillNo = Trim(TTransactionNo.Text)
         rs!InventoryType = "P"
         rs!GCode = getGCodeOfAccount(IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID))
         rs!CreditedDebitedTo = "Purchase"
         rs!Mode = "Credit"
         rs.Update
        
         rs.AddNew
         rs!BillNo = Trim(TTransactionNo.Text)
         rs!Type = "PU"
         rs!AccountCode = sPurchaseAccount
         rs!AddedDate = DTPDate.Value
         rs!EditedDate = DTPDate.Value
         rs!Debit = Val(LGrandAmount.Caption)
         rs!Credit = 0
         rs!Narration = Trim(TNarration.Text)
         rs!AddedBy = sCurrentUserCode
         rs!EditedBy = sCurrentUserCode
         rs!SerialNo = LSerialNo + 1
         rs!FinancialCode = getFinancialCode(DTPDate.Value)
         rs!InventoryBillNo = Trim(TTransactionNo.Text)
         rs!InventoryType = "P"
         rs!GCode = getGCodeOfAccount(sPurchaseAccount)
         rs!CreditedDebitedTo = CoSupplier.Text
         rs!Mode = "Credit"
         rs.Update
         LSerialNo = LSerialNo + 2
    End If
        
    'Cash To Supplier (Advance)
    
    If Val(TPayment.Text) > 0 Then
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!Type = "PU"
        rs!AccountCode = IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TPayment.Text)
        rs!Credit = 0
        rs!Narration = Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "P"
        rs!GCode = getGCodeOfAccount(IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID))
        rs!CreditedDebitedTo = "Cash"
        rs!Mode = "Cash"
        rs.Update
        
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!Type = "PU"
        rs!AccountCode = sCashAccount
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = 0
        rs!Credit = Val(TPayment.Text)
        rs!Narration = Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo + 1
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "P"
        rs!GCode = getGCodeOfAccount(sCashAccount)
        rs!CreditedDebitedTo = CoSupplier.Text
        rs!Mode = "Cash"
        rs.Update
        LSerialNo = LSerialNo + 2
    End If
    
    'Purchase Discount From Supplier
    If Val(TDiscount.Text) > 0 Then
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!Type = "PU"
        rs!AccountCode = IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TDiscount.Text)
        rs!Credit = 0
        rs!Narration = Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "P"
        rs!GCode = getGCodeOfAccount(IIf(CoSupplier.ListIndex <> -1, sAccountCode(CoSupplier.ListIndex + 1), sGeneralSupplierAccountID))
        rs!CreditedDebitedTo = "Purchase Discount"
        rs!Mode = "Credit"
        rs.Update
        
        rs.AddNew
        rs!BillNo = Trim(TTransactionNo.Text)
        rs!Type = "PU"
        rs!AccountCode = sPurchaseDiscounts
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = 0
        rs!Credit = Val(TDiscount.Text)
        rs!Narration = Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo + 1
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "P"
        rs!GCode = getGCodeOfAccount(sPurchaseDiscounts)
        rs!CreditedDebitedTo = CoSupplier.Text
        rs!Mode = "Credit"
        rs.Update
    End If

End Sub
