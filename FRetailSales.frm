VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRetailSales 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales [Form 8B]"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FRetailSales.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   12510
   StartUpPosition =   1  'CenterOwner
   Tag             =   "a"
   Begin VB.CommandButton CAddItem 
      Height          =   500
      Left            =   255
      Picture         =   "FRetailSales.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1365
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   500
      Left            =   1695
      Picture         =   "FRetailSales.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7200
      Width           =   1365
   End
   Begin VB.CommandButton CClear 
      Height          =   500
      Left            =   3135
      Picture         =   "FRetailSales.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7200
      Width           =   1365
   End
   Begin VB.CommandButton CNew 
      Height          =   500
      Left            =   270
      Picture         =   "FRetailSales.frx":205968
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7875
      Width           =   1365
   End
   Begin VB.CommandButton CPrint 
      Height          =   500
      Left            =   1725
      Picture         =   "FRetailSales.frx":207DCA
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7875
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   500
      Left            =   8325
      Picture         =   "FRetailSales.frx":20A22C
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7950
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   500
      Left            =   9765
      Picture         =   "FRetailSales.frx":20C68E
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7950
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   500
      Left            =   4695
      Picture         =   "FRetailSales.frx":20EAF0
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   45
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3255
      Left            =   60
      TabIndex        =   13
      Top             =   1470
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   5741
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
      Left            =   3015
      TabIndex        =   14
      Top             =   45
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
   Begin MSFlexGridLib.MSFlexGrid MGridItemDetails 
      Height          =   795
      Left            =   135
      TabIndex        =   5
      Top             =   6375
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   1402
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSForms.TextBox TDiscount 
      Height          =   405
      Left            =   10560
      TabIndex        =   65
      Top             =   6180
      Width           =   1605
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "2831;714"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TPayment 
      Height          =   405
      Left            =   10560
      TabIndex        =   66
      Top             =   6570
      Width           =   1605
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "2831;714"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TServiceCharge 
      Height          =   405
      Left            =   10560
      TabIndex        =   64
      Top             =   5805
      Width           =   1605
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "2831;714"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TMRPChange 
      Height          =   435
      Left            =   5565
      TabIndex        =   6
      Top             =   5280
      Width           =   1050
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "1852;767"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LMFRShortName 
      Height          =   345
      Left            =   5865
      TabIndex        =   63
      Top             =   6975
      Width           =   990
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Size            =   "1746;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LCurrentStock 
      Height          =   420
      Left            =   4740
      TabIndex        =   62
      Top             =   7995
      Visible         =   0   'False
      Width           =   1680
      VariousPropertyBits=   8388627
      Caption         =   "CurrentStock"
      Size            =   "2963;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LMRP 
      Height          =   345
      Left            =   4620
      TabIndex        =   61
      Top             =   7380
      Visible         =   0   'False
      Width           =   1680
      VariousPropertyBits=   8388627
      Caption         =   "MRP"
      Size            =   "2963;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBatch 
      Height          =   345
      Left            =   4590
      TabIndex        =   60
      Top             =   7695
      Visible         =   0   'False
      Width           =   1680
      VariousPropertyBits=   8388627
      Caption         =   "Batch"
      Size            =   "2963;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LPurchaseRate 
      Height          =   345
      Left            =   4650
      TabIndex        =   59
      Top             =   7005
      Visible         =   0   'False
      Width           =   1680
      VariousPropertyBits=   8388627
      Caption         =   "P Rate"
      Size            =   "2963;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label25 
      Height          =   345
      Left            =   1305
      TabIndex        =   58
      Top             =   6090
      Width           =   1680
      VariousPropertyBits=   8388627
      Caption         =   "MRP"
      Size            =   "2963;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Index           =   1
      Left            =   435
      TabIndex        =   57
      Top             =   6090
      Width           =   555
      VariousPropertyBits=   8388627
      Caption         =   "Qty"
      Size            =   "979;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   345
      Left            =   2835
      TabIndex        =   56
      Top             =   6090
      Width           =   1680
      VariousPropertyBits=   8388627
      Caption         =   "P Rate"
      Size            =   "2963;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label24 
      Height          =   375
      Left            =   9060
      TabIndex        =   55
      Top             =   5790
      Width           =   1605
      ForeColor       =   -2147483642
      VariousPropertyBits=   8388627
      Caption         =   "Service Charge"
      Size            =   "2831;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TItemDiscount 
      Height          =   435
      Left            =   8280
      TabIndex        =   9
      Top             =   4740
      Width           =   1035
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "1826;767"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTaxAmount 
      Height          =   360
      Left            =   10020
      TabIndex        =   54
      Top             =   4830
      Width           =   1095
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "iiiiiiiiiiiiiiiii"
      Size            =   "1931;635"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LNetValue 
      Height          =   360
      Left            =   7080
      TabIndex        =   53
      Top             =   8010
      Visible         =   0   'False
      Width           =   1035
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "iiiiiiiiiiiiiiiii"
      Size            =   "1826;635"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LGrossValue 
      Height          =   360
      Left            =   6990
      TabIndex        =   52
      Top             =   7650
      Visible         =   0   'False
      Width           =   1095
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "iiiiiiiiiiiiiiiii"
      Size            =   "1931;635"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label23 
      Height          =   330
      Left            =   8070
      TabIndex        =   51
      Top             =   1110
      Width           =   840
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Discount"
      Size            =   "1482;582"
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label22 
      Height          =   330
      Left            =   6045
      TabIndex        =   50
      Top             =   8025
      Visible         =   0   'False
      Width           =   1020
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Net Value"
      Size            =   "1799;582"
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   4245
      Left            =   45
      Top             =   1005
      Width           =   12375
   End
   Begin MSForms.TextBox TTax 
      Height          =   435
      Left            =   9360
      TabIndex        =   10
      Top             =   4740
      Width           =   585
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "1032;767"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   435
      Left            =   9015
      TabIndex        =   47
      Top             =   1095
      Width           =   585
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Tax%"
      Size            =   "1032;767"
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label12 
      Height          =   420
      Left            =   4860
      TabIndex        =   46
      Top             =   6585
      Visible         =   0   'False
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
   Begin MSForms.TextBox TWarranty 
      Height          =   435
      Left            =   7635
      TabIndex        =   8
      Top             =   4740
      Width           =   600
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "1058;767"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label11 
      Height          =   435
      Left            =   7140
      TabIndex        =   45
      Top             =   1095
      Width           =   810
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Wrty"
      Size            =   "1429;767"
      BorderColor     =   -2147483641
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TRate 
      Height          =   435
      Left            =   5550
      TabIndex        =   12
      Top             =   4740
      Width           =   1065
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "1879;767"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalAmount 
      Height          =   435
      Left            =   10920
      TabIndex        =   44
      Top             =   4785
      Width           =   1395
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2461;767"
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TQuantity 
      Height          =   435
      Left            =   6660
      TabIndex        =   7
      Top             =   4740
      Width           =   930
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "1640;767"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoItem 
      Height          =   435
      Left            =   945
      TabIndex        =   4
      Top             =   4740
      Width           =   4605
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "8123;767"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LSlNo 
      Height          =   435
      Left            =   150
      TabIndex        =   43
      Top             =   4815
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
   Begin MSForms.Label LGrandTotalAmount 
      Height          =   375
      Left            =   10980
      TabIndex        =   42
      Top             =   5355
      Width           =   1455
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Total Amt"
      Size            =   "2566;661"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LTotalGrossValue 
      Height          =   375
      Left            =   8190
      TabIndex        =   41
      Top             =   5520
      Visible         =   0   'False
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
   Begin MSForms.Label Label15 
      Height          =   375
      Left            =   9660
      TabIndex        =   40
      Top             =   6210
      Width           =   855
      ForeColor       =   -2147483642
      VariousPropertyBits=   8388627
      Caption         =   "Discount"
      Size            =   "1508;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalTaxAmount 
      Height          =   375
      Left            =   10200
      TabIndex        =   39
      Top             =   5355
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
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   5940
      TabIndex        =   38
      Top             =   7605
      Visible         =   0   'False
      Width           =   1095
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "G.Value"
      Size            =   "1931;582"
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   9495
      TabIndex        =   37
      Top             =   1095
      Width           =   1185
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Tax Amt"
      Size            =   "2090;582"
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   45
      X2              =   12405
      Y1              =   4695
      Y2              =   4725
   End
   Begin MSForms.Label LBalance 
      Height          =   405
      Left            =   10560
      TabIndex        =   36
      Top             =   6960
      Width           =   1605
      BackColor       =   16777215
      Size            =   "2831;714"
      BorderColor     =   12632256
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LBalancelb 
      Height          =   375
      Left            =   9765
      TabIndex        =   35
      Top             =   7050
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
   Begin MSForms.Label LPayment 
      Height          =   375
      Left            =   9645
      TabIndex        =   34
      Top             =   6615
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
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   75
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
   Begin MSForms.TextBox TTransactionNo 
      Height          =   420
      Left            =   1380
      TabIndex        =   0
      Top             =   45
      Width           =   1590
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "2805;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   8565
      TabIndex        =   25
      Top             =   120
      Width           =   1335
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "2355;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoCustomer 
      Height          =   420
      Left            =   9225
      TabIndex        =   2
      Top             =   60
      Width           =   3210
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "5662;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress 
      Height          =   420
      Left            =   9225
      TabIndex        =   3
      Top             =   495
      Width           =   3210
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "5662;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LGrandAmount 
      Height          =   570
      Left            =   8625
      TabIndex        =   24
      Top             =   7335
      Width           =   3300
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Grand Amount"
      Size            =   "5821;1005"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Index           =   0
      Left            =   90
      TabIndex        =   23
      Top             =   1095
      Width           =   555
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Sl No"
      Size            =   "979;582"
      FontName        =   "Sylfaen"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   285
      Left            =   1875
      TabIndex        =   22
      Top             =   1095
      Width           =   435
      ForeColor       =   -2147483634
      VariousPropertyBits=   276824083
      Caption         =   "Item"
      Size            =   "767;503"
      FontName        =   "Sylfaen"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label16 
      Height          =   435
      Left            =   6420
      TabIndex        =   21
      Top             =   1080
      Width           =   855
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Qty"
      Size            =   "1508;767"
      FontName        =   "Sylfaen"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label19 
      Height          =   330
      Left            =   10575
      TabIndex        =   20
      Top             =   1080
      Width           =   1305
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2302;582"
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   225
      TabIndex        =   19
      Top             =   555
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
   Begin MSForms.TextBox TNarration 
      Height          =   420
      Left            =   1395
      TabIndex        =   1
      Top             =   480
      Width           =   3180
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      Size            =   "5609;741"
      BorderColor     =   8421504
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LUnit 
      Height          =   330
      Left            =   6045
      TabIndex        =   18
      Top             =   6630
      Visible         =   0   'False
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
   Begin MSForms.Label Label10 
      Height          =   435
      Left            =   5520
      TabIndex        =   17
      Top             =   1095
      Width           =   1035
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Rate"
      Size            =   "1826;767"
      FontName        =   "Sylfaen"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LManufacturer 
      Height          =   345
      Left            =   5970
      TabIndex        =   16
      Top             =   6195
      Visible         =   0   'False
      Width           =   1545
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Size            =   "2725;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   4485
      TabIndex        =   15
      Top             =   6180
      Visible         =   0   'False
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
   Begin MSForms.Label Label21 
      Height          =   495
      Left            =   30
      TabIndex        =   49
      Top             =   1005
      Width           =   12405
      BackColor       =   15724527
      VariousPropertyBits=   8388627
      Size            =   "21881;873"
      Picture         =   "FRetailSales.frx":210F52
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label20 
      Height          =   495
      Left            =   30
      TabIndex        =   48
      Top             =   1005
      Width           =   11685
      BackColor       =   15724527
      VariousPropertyBits=   8388627
      Size            =   "20611;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FRetailSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim dMRP As Double, dQuantity As Double, dBatchMRP() As Double, dBatchQuantity() As Double
Dim sCustomerCode() As String, sCustomerAddress() As String, sAccountCode() As String
Dim sItemCode() As String, sBillingName() As String
Dim gIQuantity As Single, gIMRP As Single, gIPurchaseRate As Single, gIPurchaseCode As Single, gIBatch As Single
Dim gSerialNo As Single, gItem As Single, gPurchaseRate As Single, gQuantity As Single, gTaxAmount As Single, gTax As Single, gNetValue As Single, gItemDiscount As Single, gBatch As Single, gWarranty As Single, gGrossValue As Single, gUnit As Single, gSaleRate As Single, gMRP As Single, gToTalAmount As Single, gBillingName As Single, gItemCode As Single, gMFRShortName As Single

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
       
    If Val(TQuantity.Text) > Val(LCurrentStock.Caption) Then
        lYN = MsgBox("There is no enough Stock, Do you want to Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TQuantity.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(TRate.Text) = 0 Then
        lYN = MsgBox("Rate given is Zero, Do you want to Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TRate.SetFocus
            Exit Sub
        End If
    End If

'    If Val(TRate.Text) < Val(LMRP.Caption) Then
'        lYN = MsgBox("Rate given is less than MRP, Do you Want to continue ?", vbDefaultButton2 Or vbYesNo)
'        If lYN = vbYes Then
'        Else
'            TRate.SetFocus
'            Exit Sub
'        End If
'    End If
        
    If Val(LSlNo.Caption) > MGrid.Rows Then 'Add
        MGrid.AddItem ""
        MGrid.TextMatrix(MGrid.Rows - 1, gSerialNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(MGrid.Rows - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gTax) = TTax.Text
        MGrid.TextMatrix(MGrid.Rows - 1, gBatch) = getUniversaloFor((LPurchaseRate.Caption) & (LMRP.Caption))
        MGrid.TextMatrix(MGrid.Rows - 1, gWarranty) = Val(TWarranty.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gItemDiscount) = Format(Val(TItemDiscount.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gNetValue) = Format(Val(MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue)) - (MGrid.TextMatrix(MGrid.Rows - 1, gItemDiscount)), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount) = Format(Val(MGrid.TextMatrix(MGrid.Rows - 1, gNetValue)) * Val(MGrid.TextMatrix(MGrid.Rows - 1, gTax)) / 100, "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gToTalAmount) = Format((Val(MGrid.TextMatrix(MGrid.Rows - 1, gNetValue)) + Val(MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount))), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gMFRShortName) = LMFRShortName.Caption & ""
        MGrid.TextMatrix(MGrid.Rows - 1, gMRP) = Val(LMRP.Caption)
        MGrid.TextMatrix(MGrid.Rows - 1, gPurchaseRate) = Val(LPurchaseRate.Caption)
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(r - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(r - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(r - 1, gBatch) = getUniversaloFor((LPurchaseRate.Caption) & (LMRP.Caption))
        MGrid.TextMatrix(r - 1, gTax) = TTax.Text
        MGrid.TextMatrix(r - 1, gWarranty) = Val(TWarranty.Text)
        MGrid.TextMatrix(r - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(r - 1, gGrossValue) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(r - 1, gItemDiscount) = Format(Val(TItemDiscount.Text), "0.00")
        MGrid.TextMatrix(r - 1, gNetValue) = Format(Val(MGrid.TextMatrix(r - 1, gGrossValue)) - (MGrid.TextMatrix(r - 1, gItemDiscount)), "0.00")
        MGrid.TextMatrix(r - 1, gTaxAmount) = Format(Val(MGrid.TextMatrix(r - 1, gNetValue)) * Val(MGrid.TextMatrix(r - 1, gTax)) / 100, "0.00")
        MGrid.TextMatrix(r - 1, gToTalAmount) = Format((Val(MGrid.TextMatrix(r - 1, gNetValue)) + Val(MGrid.TextMatrix(r - 1, gTaxAmount))), "0.00")
        MGrid.TextMatrix(r - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gMFRShortName) = LMFRShortName.Caption & ""
        MGrid.TextMatrix(r - 1, gMRP) = Val(LMRP.Caption)
        MGrid.TextMatrix(r - 1, gPurchaseRate) = Val(LPurchaseRate.Caption)
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
Private Sub MGridItemDetailsInitialise()
'INITIALISES MGridItemDetails
        'SETTING CONSTANTS
    gIQuantity = 0
    gIMRP = 1
    gIPurchaseRate = 2
    gIBatch = 3
    
    MGridItemDetails.Clear
    MGridItemDetails.Rows = 1 'FOR SKIPING ERROR
    MGridItemDetails.Cols = 1 'FOR SKIPING ERROR
    MGridItemDetails.FixedCols = 0
    MGridItemDetails.FixedRows = 0
    MGridItemDetails.Cols = 4
    MGridItemDetails.Rows = 0
    MGridItemDetails.ColWidth(gIQuantity) = 1270
    MGridItemDetails.ColWidth(gIMRP) = 1400
    MGridItemDetails.ColWidth(gIPurchaseRate) = 1400
    MGridItemDetails.ColWidth(gIBatch) = 0
    MGridItemDetails.RowHeightMin = 350
End Sub
Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gItem = 1
    gBatch = 2
    gUnit = 3
    gSaleRate = 4
    gQuantity = 5
    gWarranty = 6
    gGrossValue = 7
    gItemDiscount = 8
    gNetValue = 9
    gTax = 10
    gTaxAmount = 11
    gToTalAmount = 12
    gBillingName = 13
    gItemCode = 14
    gMFRShortName = 15
    gMRP = 16
    gPurchaseRate = 17
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 18
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 890
    MGrid.ColWidth(gItem) = 4500
    MGrid.ColWidth(gQuantity) = 800
    MGrid.ColWidth(gTax) = 600
    MGrid.ColWidth(gBatch) = 0
    MGrid.ColWidth(gWarranty) = 700
    MGrid.ColWidth(gUnit) = 0
    MGrid.ColWidth(gSaleRate) = 1160
    MGrid.ColWidth(gItemDiscount) = 800
    MGrid.ColWidth(gNetValue) = 0
    MGrid.ColWidth(gGrossValue) = 0
    MGrid.ColWidth(gToTalAmount) = 1160
    MGrid.ColWidth(gTaxAmount) = 1160
    MGrid.ColWidth(gBillingName) = 0
    MGrid.ColWidth(gItemCode) = 0
    MGrid.ColWidth(gMFRShortName) = 0
    MGrid.ColWidth(gMRP) = 0
    MGrid.ColWidth(gPurchaseRate) = 0
    
    MGrid.ColAlignment(gItem) = vbLeftJustify
    MGrid.ColAlignment(gUnit) = vbLeftJustify
        
    MGrid.RowHeightMin = 350
End Sub

Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String
    
    Set rs = db.OpenRecordset("Select Max(Val( Transaction.TransactionNo)) As TNo From Transaction Where ( Transaction.TransactionType = 'S' )")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function
Private Sub getCustomer()
Dim rs As Recordset
    
    CoCustomer.Clear
    Set rs = db.OpenRecordset("Select CustomerMaster.CustomerCode,CustomerMaster.AccountCode,CustomerMaster.CustomerName,CustomerMaster.Address1,CustomerMaster.Address2,CustomerMaster.Address3 From CustomerMaster Where (CustomerMaster.Status = True) Order By CustomerMaster.CustomerName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If

    ReDim sCustomerCode(rs.RecordCount) As String
    ReDim sCustomerAddress(rs.RecordCount) As String
    ReDim sAccountCode(rs.RecordCount) As String
    While rs.EOF = False
        CoCustomer.AddItem "" & rs!CustomerName
        sCustomerCode(CoCustomer.ListCount) = "" & rs!CustomerCode
        sCustomerAddress(CoCustomer.ListCount) = "" & rs!Address1 & " " & rs!Address2 & " " & rs!Address3
        sAccountCode(CoCustomer.ListCount) = "" & rs!AccountCode
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
        Exit Sub
    End If

    If (CoItem.ListIndex = -1) Then
        LManufacturer.Caption = ""
        LUnit.Caption = ""
        TTax.Text = ""
        LMFRShortName.Caption = ""
        Exit Sub
    End If
    Set rs = db.OpenRecordset("Select ItemMaster.Tax,Manufacturer.ShortName,Manufacturer.ManufacturerName,Units.UnitName,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('O','P','SR','SA') ) And (Transaction.ItemCode = ItemMaster.Code )) As InStock,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('S','PR') ) And (Transaction.ItemCode = ItemMaster.Code )) As OutStock From ItemMaster,Units,Manufacturer Where (ItemMaster.Code = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code = ItemMaster.ManufacturerCode )")
    If rs.RecordCount > 0 Then
        LManufacturer.Caption = "" & rs!ManufacturerName
        LUnit.Caption = "" & rs!UnitName
        TTax.Text = "" & rs!Tax
        LMFRShortName.Caption = "" & rs!ManufacturerName
    Else
        LManufacturer.Caption = ""
        LUnit.Caption = ""
        TTax.Text = ""
        LMFRShortName.Caption = ""
    End If
    rs.Close
                    
    MGridItemDetails.Rows = 0
    TQuantity.Text = ""
    TRate.Text = ""
    TMRPChange.Text = ""

'    r = 0
'    While r < MGrid.Rows
'        If sItemCode(CoItem.ListIndex + 1) = MGrid.TextMatrix(r, gItemCode) Then
'            dQuantity = dQuantity + MGrid.TextMatrix(r, gQuantity)
'        End If
'        r = r + 1
'    Wend
'    Set rs = db.OpenRecordset("Select (Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('S') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.TransactionNo = '" & Trim(TTransactionNo.Text) & "')) As Quantity From Transaction ")
'    If rs.RecordCount > 0 Then
'        dQuantity = dQuantity - Val("" & rs!Quantity)
'    End If

    Set rs = db.OpenRecordset("Select Transaction.Batch,Transaction.MRP,Transaction.Purchaserate,(Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('O','P','SR','SA') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.MRP = Transaction.MRP)And (T.PurchaseRate = Transaction.PurchaseRate) ) As InStock,(Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('S','W','PR') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.MRP = Transaction.MRP)And (T.PurchaseRate = Transaction.PurchaseRate)) As OutStock From Transaction Where (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) Group By Transaction.Batch,Transaction.MRP,Transaction.PurchaseRate")
    While rs.EOF = False
    If (Val("" & rs!InStock) - Val("" & rs!OutStock) - dQuantity) > 0 Then
        MGridItemDetails.AddItem Val("" & rs!InStock) - Val("" & rs!OutStock) - dQuantity & vbTab & rs!MRP & vbTab & rs!PurchaseRate & vbTab & rs!Batch
'        TQuantity.Text = MGridItemDetails.TextMatrix(0, gIQuantity)
        LCurrentStock.Caption = MGridItemDetails.TextMatrix(0, gIQuantity)
        TRate.Text = Format((Val(MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIMRP)) * 100) / (Val(TTax.Text) + 100), "0.00")
        TMRPChange.Text = Format(Val(MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIMRP)), "0.00")
        LMRP.Caption = MGridItemDetails.TextMatrix(0, gIMRP)
        LBatch.Caption = MGridItemDetails.TextMatrix(0, gIBatch)
        LPurchaseRate.Caption = MGridItemDetails.TextMatrix(0, gIPurchaseRate)
     End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Function getMRP()
Dim rs As Recordset, r As Long
        
    dMRP = 0
    If (CoItem.ListIndex = -1) Then
        Exit Function
    End If
    
    Set rs = db.OpenRecordset("Select Transaction.MRP From Transaction Where (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) Group By Transaction.MRP")
    If rs.RecordCount > 0 Then
        dMRP = Val("" & rs!MRP)
    End If
    rs.Close
    
    getMRP = dMRP
End Function

Private Sub clearControls()
    
    'TTransactionNo.Text = getNewTransactionNo
    'DTPDate.Value = Date
    TNarration.Text = ""
    CoCustomer.Text = ""
    TAddress.Text = ""
    TWarranty.Text = ""
    TMRPChange.Text = ""
    MGrid.Rows = 0
    MGridItemDetails.Rows = 0
    TPayment.Text = ""
    LBalance.Caption = ""
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TQuantity.Text = ""
    TTax.Text = ""
    LUnit.Caption = ""
    TRate.Text = ""
    LTotalAmount.Caption = ""
    TServiceCharge.Text = ""
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    getTotal
    TDiscount.Text = 0#
    TServiceCharge.Text = ""
    TItemDiscount.Text = ""
End Sub

Private Sub clearEditControls()
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    LMRP.Caption = ""
    TTax.Text = ""
    TRate.Text = ""
    TMRPChange.Text = ""
    LTotalAmount.Caption = ""
    TDiscount.Text = 0#
    LBalance.Caption = 0
    TWarranty.Text = 0
    TServiceCharge.Text = ""
    TItemDiscount.Text = ""
    MGridItemDetails.Rows = 0
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
    getGrandTotal = dGrandTotal + Val(TServiceCharge.Text)
    LTotalGrossValue.Caption = Format(dGrossValue, "0.00")
    LGrandTotalAmount.Caption = Format(dGrandTotal, "0.00")
    LTotalTaxAmount.Caption = Format(dTax, "0.00")
End Function

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean
    bFound = False
    If (MsgBox("Do you want to Delete the Bill ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'S' )")
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
    Set rs = db.OpenRecordset("Select * From AccountTransaction Where (AccountTransaction.Type = 'SB') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') And (AccountTransaction.InventoryBillNo='" & TTransactionNo.Text & "') And (AccountTransaction.InventoryType='SB') ")
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

Private Sub CoCustomer_Change()
    If CoCustomer.ListIndex <> -1 Then
        TAddress.Text = sCustomerAddress(CoCustomer.ListIndex + 1)
    Else
        TAddress.Text = ""
    End If
End Sub

Private Sub CoCustomer_GotFocus()
    CoCustomer.SelStart = 0
    CoCustomer.SelLength = Len(CoCustomer.Text)
End Sub

Private Sub CoCustomer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FCustomerRegister.Show vbModal
        getCustomer
    End If
End Sub

Private Sub CPrint_Click()
     printSales
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
Dim r As Long, lYN As Long, sStatus As String, sCustomer As String

    If Val(TTransactionNo.Text) = 0 Then
        MsgBox "Please Enter Valid Transaction No !", vbInformation
        TTransactionNo.SetFocus
        Exit Sub
    End If
    
    If Trim(CoCustomer.Text) = "" Then
        lYN = MsgBox("Do You Want To Consider General Customer !", vbDefaultButton2 Or vbYesNo)
        If lYN <> vbYes Then
            CoCustomer.SetFocus
            Exit Sub
        End If
    End If
    
    If MGrid.Rows = 0 Then
        MsgBox "No Items Entered !", vbInformation
        CoItem.SetFocus
        Exit Sub
    End If
    
    
    'SAVES DATA TO Transaction TABLE
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'S' )")
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
        rs!TransactionType = "S"
        rs!TransactionDate = DTPDate.Value
        rs!TransactionTime = Format(Time, "HH:MM AMPM")
        rs!Narration = Trim(TNarration.Text)
        rs!SupplierCode = ""
        rs!SupplierName = ""
        rs!SupplierAddress = ""
        rs!CustomerCode = IIf(CoCustomer.ListIndex = -1, "", sCustomerCode(CoCustomer.ListIndex + 1))
        rs!CustomerName = Trim(CoCustomer.Text)
        rs!CustomerAddress = Trim(TAddress.Text)
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!ItemCode = Trim(MGrid.TextMatrix(r, gItemCode))
        rs!Quantity = Val(MGrid.TextMatrix(r, gQuantity))
        rs!PurchaseRate = 0
        rs!SaleRate = Val(MGrid.TextMatrix(r, gSaleRate))
        rs!SalePayment = IIf(Val(TPayment.Text) = 0, "0", Val(TPayment.Text))
        rs!MRP = 0
        rs!ReferenceNo = ""
        rs!Batch = Trim(MGrid.TextMatrix(r, gBatch))
        rs!Warranty = Trim(MGrid.TextMatrix(r, gWarranty))
        rs!MRP = Trim(MGrid.TextMatrix(r, gMRP))
        rs!PurchaseRate = Trim(MGrid.TextMatrix(r, gPurchaseRate))
        rs!ReferenceDate = Date
        rs!Tax = Val(MGrid.TextMatrix(r, gTax))
        rs!ServiceCharge = IIf(Val(TServiceCharge.Text) = 0, "0", Val(TServiceCharge.Text))
        rs!ItemDiscount = Val(MGrid.TextMatrix(r, gItemDiscount))
        rs!Discount = IIf(Val(TDiscount.Text) = 0, "0", Val(TDiscount.Text))
        rs.Update
        r = r + 1
    Wend
    rs.Close
    
    addToAccountRegister
    
    MsgBox "Successfully Saved !", vbInformation
    
    lYN = MsgBox("Do you want to take Print ?", vbDefaultButton2 Or vbYesNo)
    If lYN = vbYes Then
        printSales
'        printSaleBill Trim(TTransactionNo.Text), db
    Else
        
    End If
    clearControls
    TTransactionNo.Text = getNewTransactionNo
    TTransactionNo.SetFocus
End Sub

Private Sub addToAccountRegister()
Dim rs As Recordset, sTransactionNo As String, LSerialNo As Long
    
    sTransactionNo = Val(TTransactionNo.Text)
    
    Set rs = db.OpenRecordset("Select * From AccountTransaction Where (AccountTransaction.Type = 'SB') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "') And (AccountTransaction.InventoryBillNo='" & Trim(TTransactionNo.Text) & "') And (AccountTransaction.InventoryType='SB') ")
    
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    'Customer To Sale Account
    LSerialNo = 1
    If (Val(LGrandAmount.Caption) + Val(TServiceCharge.Text)) > 0 Then
         rs.AddNew
         rs!BillNo = sTransactionNo
         rs!Type = "SB"
         rs!AccountCode = IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID)
         rs!AddedDate = DTPDate.Value
         rs!EditedDate = DTPDate.Value
         rs!Credit = 0
         rs!Debit = Val(LGrandAmount.Caption) + Val(TServiceCharge.Text)
         rs!Narration = "Sale Amount Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
         rs!AddedBy = sCurrentUserCode
         rs!EditedBy = sCurrentUserCode
         rs!SerialNo = LSerialNo
         rs!FinancialCode = getFinancialCode(DTPDate.Value)
         rs!InventoryBillNo = Trim(TTransactionNo.Text)
         rs!InventoryType = "SB"
         rs!GCode = getGCodeOfAccount(IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID))
         rs!CreditedDebitedTo = "Sales Form 8B"
         rs!Mode = "Credit"
         rs.Update
        
         rs.AddNew
         rs!BillNo = sTransactionNo
         rs!Type = "SB"
         rs!AccountCode = sSalesForm8B
         rs!AddedDate = DTPDate.Value
         rs!EditedDate = DTPDate.Value
         rs!Debit = 0
         rs!Credit = Val(LGrandAmount.Caption) + Val(TServiceCharge.Text)
         rs!Narration = "Sales Form 8B Amount Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
         rs!AddedBy = sCurrentUserCode
         rs!EditedBy = sCurrentUserCode
         rs!SerialNo = LSerialNo + 1
         rs!FinancialCode = getFinancialCode(DTPDate.Value)
         rs!InventoryBillNo = Trim(TTransactionNo.Text)
         rs!InventoryType = "SB"
         rs!GCode = getGCodeOfAccount(sSaleAccount)
         rs!CreditedDebitedTo = CoCustomer.Text
         rs!Mode = "Credit"
         rs.Update
         LSerialNo = LSerialNo + 2
    End If
    
    'Customer To Cash (Advance)
    If Val(TPayment.Text) > 0 Then
        rs.AddNew
        rs!BillNo = sTransactionNo
        rs!Type = "SB"
        rs!AccountCode = IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = 0
        rs!Credit = Val(TPayment.Text)
        rs!Narration = "Sales Form 8B Advance Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "SB"
        rs!GCode = getGCodeOfAccount(IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID))
        rs!CreditedDebitedTo = "Cash"
        rs!Mode = "Cash"
        rs.Update
        
        rs.AddNew
        rs!BillNo = sTransactionNo
        rs!Type = "SB"
        rs!AccountCode = sCashAccount
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TPayment.Text)
        rs!Credit = 0
        rs!Narration = "Sales Form 8B Advance Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo + 1
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "SB"
        rs!GCode = getGCodeOfAccount(sCashAccount)
        rs!CreditedDebitedTo = CoCustomer.Text
        rs!Mode = "Cash"
        rs.Update
        LSerialNo = LSerialNo + 2
    End If
    
    'Sale Discount From Customer
    If Val(TDiscount.Text) > 0 Then
        rs.AddNew
        rs!BillNo = sTransactionNo
        rs!Type = "SB"
        rs!AccountCode = IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = 0
        rs!Credit = Val(TDiscount.Text)
        rs!Narration = "Sales Form 8B Disount Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "SB"
        rs!GCode = getGCodeOfAccount(IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sGeneralCustomerAccountID))
        rs!CreditedDebitedTo = "Sales Form 8B Discount"
        rs!Mode = "Credit"
        rs.Update
        
        rs.AddNew
        rs!BillNo = sTransactionNo
        rs!Type = "SB"
        rs!AccountCode = sSaleDiscounts
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TDiscount.Text)
        rs!Credit = 0
        rs!Narration = "Sales Form 8B Discount Bill No" & TTransactionNo.Text & " " & Trim(TNarration.Text)
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = LSerialNo + 1
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!InventoryBillNo = Trim(TTransactionNo.Text)
        rs!InventoryType = "SB"
        rs!GCode = getGCodeOfAccount(sSaleDiscounts)
        rs!CreditedDebitedTo = CoCustomer.Text
        rs!Mode = "Credit"
        rs.Update
    End If
    
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
    getCustomer
    getItem
    MGridItemDetailsInitialise
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
        TTax.Text = Val(MGrid.TextMatrix(r, gTax))
        LUnit.Caption = Trim(MGrid.TextMatrix(r, gUnit))
        TWarranty.Text = Val(MGrid.TextMatrix(r, gWarranty))
        TItemDiscount.Text = Val(MGrid.TextMatrix(r, gItemDiscount))
        TRate.Text = Val(MGrid.TextMatrix(r, gSaleRate))
        TMRPChange.Text = Format(Val(MGrid.TextMatrix(r, gSaleRate)) + (Val(MGrid.TextMatrix(r, gSaleRate)) * Val(MGrid.TextMatrix(r, gTax)) / 100), "0.00")
        LTotalAmount.Caption = Val(MGrid.TextMatrix(r, gToTalAmount))
        LPurchaseRate.Caption = Val(MGrid.TextMatrix(r, gPurchaseRate))
        LMRP.Caption = Val(MGrid.TextMatrix(r, gMRP))
    Else
    End If
End Sub

Private Sub MGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub getTotal()
    LGrossValue.Caption = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
    LNetValue = Format((Val(TRate.Text) * Val(TQuantity.Text)) - Val(TItemDiscount.Text), "0.00")
    LTaxAmount = Format(Val(LNetValue.Caption) * Val(TTax.Text) / 100, "0.00")
    LTotalAmount = Format(Val(LNetValue.Caption) + Val(LTaxAmount.Caption), "0.00")
End Sub

Private Sub MGridItemDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If (MGridItemDetails.Rows > 0) Then
'            TQuantity.Text = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIQuantity)
            LCurrentStock.Caption = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIQuantity)
            TRate.Text = Format((Val(MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIMRP)) * 100) / (Val(TTax.Text) + 100), "0.00")
            TMRPChange.Text = Format(Val(MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIMRP)), "0.00")
            LMRP.Caption = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIMRP)
            LBatch.Caption = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIBatch)
            LPurchaseRate.Caption = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIPurchaseRate)
        End If
        SendKeys "{TAB}"
    End If
End Sub


Private Sub TMRPChange_Change()
    TRate.Text = Format((Val(TMRPChange.Text) * 100) / (Val(TTax.Text) + 100), "0.00")
End Sub

Private Sub TPayment_Change()
    LBalance.Caption = Val(LGrandAmount.Caption) - Val(TPayment.Text)
End Sub

Private Sub TPayment_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    If Val(LGrandAmount.Caption) < Val(TPayment.Text) Then
        TPayment.SetFocus
        Exit Sub
    ElseIf Val(TPayment.Text) = 0 Then
        LBalance.Caption = Format(LGrandAmount, "0.00")
    Else
        LBalance.Caption = Format(Trim(LGrandAmount) - Val(TPayment.Text), "0.00")
    End If
    CSave.SetFocus
  End If
End Sub

Private Sub TServiceCharge_Change()
   getGrandTotal
End Sub

Private Sub TServiceCharge_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then
    If Val(TServiceCharge.Text) = 0 Then
        TPayment.Text = Format(LGrandAmount, "0.00")
        LBalance.Caption = 0#
    End If
    TDiscount.SetFocus
  End If
End Sub

Private Sub TServiceCharge_LostFocus()
    LGrandAmount.Caption = Format(getGrandTotal - Val(TDiscount.Text), "0.00")
    TPayment.Text = Format(getGrandTotal - Val(TDiscount.Text), "0.00")
End Sub

Private Sub TAddress_GotFocus()
    TAddress.SelStart = 0
    TAddress.SelLength = Len(TAddress.Text)
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
LGrandAmount.Caption = Format(getGrandTotal - Val(TDiscount.Text), "0.00")
TPayment.Text = Format(getGrandTotal - Val(TDiscount.Text), "0.00")
End Sub

Private Sub TItemDiscount_Change()
    getTotal
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

Private Sub TTax_Change()
    getTotal
End Sub

Private Sub TTax_GotFocus()
    TTax.SelStart = 0
    TTax.SelLength = Len(TTax.Text)
End Sub

Private Sub TTransactionNo_GotFocus()
    TTransactionNo.SelStart = 0
    TTransactionNo.SelLength = Len(TTransactionNo.Text)
End Sub

Private Sub TTransactionNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        clearControls
        getTransactionDetails
'        LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    End If
End Sub

Public Sub getTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.BillingName,ItemMaster.Code,Transaction.Tax As ItemTax,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'S' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode) Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!TransactionDate
        CoCustomer.Text = "" & rs!CustomerName
        TAddress.Text = "" & rs!CustomerAddress
        TNarration.Text = "" & rs!Narration
        TPayment.Text = Format("" & rs!SalePayment, "0.00")
        TDiscount.Text = Format("" & rs!Discount, "0.00")
        TServiceCharge.Text = Format("" & rs!ServiceCharge, "0.00")
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gQuantity) = "" & rs!Quantity
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gBatch) = "" & rs!Batch
            MGrid.TextMatrix(r, gTax) = "" & rs!ItemTax
            MGrid.TextMatrix(r, gWarranty) = "" & rs!Warranty
            MGrid.TextMatrix(r, gSaleRate) = Format("" & rs!SaleRate, "0.00")
            MGrid.TextMatrix(r, gGrossValue) = Format(Val("" & rs!Quantity) * Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gItemDiscount) = Format("" & rs!ItemDiscount, "0.00")
            MGrid.TextMatrix(r, gNetValue) = Format(Val(MGrid.TextMatrix(r, gGrossValue)) - Val("" & rs!ItemDiscount), "0.00")
            MGrid.TextMatrix(r, gTaxAmount) = Format(Val(MGrid.TextMatrix(r, gNetValue)) * rs!ItemTax / 100, "0.00")
            MGrid.TextMatrix(r, gToTalAmount) = Format(Val(MGrid.TextMatrix(r, gNetValue) + Val(MGrid.TextMatrix(r, gTaxAmount))), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gMFRShortName) = "" & rs!ShortName
            MGrid.TextMatrix(r, gMRP) = Format("" & rs!MRP, "0.00")
            MGrid.TextMatrix(r, gPurchaseRate) = "" & rs!PurchaseRate
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
    
    LSlNo.Caption = MGrid.Rows + 1
    LGrandAmount.Caption = Format(getGrandTotal - Val(TDiscount.Text), "0.00")
    LBalance.Caption = Format(Trim(LGrandAmount.Caption) - Val(TPayment.Text), "0.00")
    
End Sub

Private Function getCentralAlignmentStartingPos(lPrintWidth As Long, sWord As String) As Long
Dim dPos As Long, lWordLen As Long
    lWordLen = Len(sWord)
    dPos = (lPrintWidth / 2) - (lWordLen / 2)
    getCentralAlignmentStartingPos = dPos
End Function

Private Sub TWarranty_GotFocus()
    TWarranty.SelStart = 0
    TWarranty.SelLength = Len(TWarranty.Text)
End Sub

Private Sub printSales()

On Error GoTo GoOut
    Dim i, x, y As Double
    Dim ITaxAmount, IGrossvalue, INetValue, IDiscount, IQty, ITotalValue, IRate As Double

    ITaxAmount = 0
    IGrossvalue = 0
    INetValue = 0
    IDiscount = 0
    IQty = 0
    ITotalValue = 0
    IRate = 0

    i = 0
    x = 500
    
    y = NewPage + 400
    
    While (i < MGrid.Rows)
    
        Printer.FontSize = 9
        Printer.FontBold = False
                
        x = 550
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Trim(MGrid.TextMatrix(i, gSerialNo))

        x = 1100
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Trim(MGrid.TextMatrix(i, gBillingName))
        
        x = 5250
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gTax), "0.00")
        
        x = 5950
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gSaleRate), "0.00")
        
        x = 7250
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Trim(MGrid.TextMatrix(i, gQuantity))
        
'        x = 7250
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(MGrid.TextMatrix(i, gGrossValue), "0.00")
        
        x = 7950
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gItemDiscount), "0.00")

'        x = 9100
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(MGrid.TextMatrix(i, gNetValue), "0.00")
        
         x = 8950
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gTaxAmount), "0.00")
        
         x = 9950
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gToTalAmount), "0.00")
        
        ITaxAmount = ITaxAmount + Val(MGrid.TextMatrix(i, gTaxAmount))
        IDiscount = IDiscount + Val(MGrid.TextMatrix(i, gItemDiscount))
        IGrossvalue = IGrossvalue + Val(MGrid.TextMatrix(i, gGrossValue))
        INetValue = INetValue + Val(MGrid.TextMatrix(i, gNetValue))
        IQty = IQty + Val(MGrid.TextMatrix(i, gQuantity))
        IRate = IRate + Val(MGrid.TextMatrix(i, gSaleRate))
        
'        If Val(MGrid.TextMatrix(i, gTax)) = 4 Then
'            Check = True
'            TGrossvalue = Tgross + Val(MGrid.TextMatrix(i, gGrossValue))
'            Taxamt = Tex + Val(MGrid.TextMatrix(i, gTaxAmount))
'            TNetamt = Tnet + Val(MGrid.TextMatrix(i, gTotalAmount))
'        Else
'            Check1 = True
'            TGrossvalue1 = Tgross1 + Val(MGrid.TextMatrix(i, gGrossValue))
'            Taxamt1 = Tex1 + Val(MGrid.TextMatrix(i, gTaxAmount))
'            TNetamt1 = Tnet1 + Val(MGrid.TextMatrix(i, gTotalAmount))
'        End If
        
        i = i + 1
        y = y + 300
        If (y > 13000) Then
            y = NewPage + 400
        End If
    Wend
        
        y = 9800
        x = 1700
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "TOTAL"
        
'        x = 5550
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(IRate, "0.00")
        
        x = 7250
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print IQty
        
'        x = 7250
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(IGrossvalue, "0.00")
        
        x = 7950
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(IDiscount, "0.00")

'        x = 9100
'        Printer.CurrentX = x
'        Printer.CurrentY = y
'        Printer.Print Format(INetValue, "0.00")
        
         x = 8950
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(ITaxAmount, "0.00")
        
         x = 9950
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(INetValue + ITaxAmount, "0.00")
    
    y = 10100
    Printer.FontBold = True
    Printer.CurrentX = 8750
    Printer.CurrentY = y
    Printer.Print "  Servicing Charge:"
    
    Printer.CurrentX = 10600
    Printer.CurrentY = y
    Printer.Print Format(Val(TServiceCharge.Text), "0.00")
    
    y = y + 500
    Printer.FontBold = True
    Printer.CurrentX = 8750
    Printer.CurrentY = y
    Printer.Print "        Discount Amt:"
    
    Printer.CurrentX = 10600
    Printer.CurrentY = y
    Printer.Print Format(Val(TDiscount.Text), "0.00")
        
    y = y + 500
    Printer.FontSize = 16
    Printer.FontBold = True
    Printer.CurrentX = 9200
    Printer.Font = "Rupee"
    Printer.CurrentY = y
    Printer.Print "`"
    
    Printer.CurrentX = 9900
    Printer.CurrentY = y
    Printer.Font = "Arial"
    Printer.FontSize = 12
    Printer.Print Format(Val(LGrandAmount.Caption), "0.00")
    
'    Printer.Print Tab(5); String(110, "-")
'    Printer.Print Tab(10); "Tax"; Tab(20); "Gross Value"; Tab(35); "Tax Amt"; Tab(50); "Cess Amt"; Tab(65); "Net Amount";
'    Printer.Print Tab(5); String(110, "-")
'
'    If Check = True Then
'        Printer.Print Tab(10); "4.00"; Tab(20); Format(TGrossvalue, "0.00"); Tab(35); Format(Taxamt, "0.00"); Tab(50); Format(Taxamt * 0.01, "0.00"); Tab(67); Format(TNetamt, "0.00");
'    End If
'    If Check1 = True Then
'        Printer.Print Tab(10); "12.50"; Tab(20); Format(TGrossvalue1, "0.00"); Tab(35); Format(Taxamt1, "0.00"); Tab(50); Format(Taxamt1 * 0.01, "0.00"); Tab(67); Format(TNetamt1, "0.00");
'    End If
'    Printer.Print Tab(5); String(110, "-")
    
    x = 500
    y = 13900
    Printer.FontSize = 10
    Printer.FontUnderline = False
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Customer Signature"
    
    x = 9300
    y = 13900
    Printer.FontSize = 10
    Printer.FontUnderline = False
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "For REDLINES"
    
    x = 500
    y = 14200
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.FontUnderline = False
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Received With Good Condition"
    
    Printer.EndDoc
    
    x = MsgBox("Successfully Printed !", vbInformation)
    
GoOut:
End Sub

Private Function NewPage() As Long

    Dim i, j, x, y, D, M, YR, DT1, TOPH As Double
    Dim Declaration(10) As String
    
    Printer.ScaleMode = 1
    Printer.FontName = "Arial"
    Printer.FontBold = False
    y = 400
    x = 450
    
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "CST NO :"
  
    x = x + 8500
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "TIN NO :"
    
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 14
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("RED LINES")) / 2)
    Printer.CurrentY = 400
    Printer.Print "RED LINES"
    Printer.FontUnderline = False
    Printer.FontBold = False
    x = 400
    y = 800
'
'    Printer.FontUnderline = True
'    Printer.FontSize = 16
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.FontUnderline = True
'    Printer.Print "Ink - Opening Stock"
'
'    Printer.FontBold = False

    Printer.FontSize = 10
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Peruvazhiyambalam , Tirur ")) / 2)
    Printer.CurrentY = 800
    Printer.Print "Peruvazhiyambalam , Tirur "
    
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("THE KERALA VALUE ADDED TAX RULES 2005 FORM NO.8B")) / 2)
    Printer.CurrentY = 1000
    Printer.Print "THE KERALA VALUE ADDED TAX RULES 2005 FORM NO.8B"

    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("(For Customers When input tax credit is not required)[See Rule 58(10)]")) / 2)
    Printer.CurrentY = 1200
    Printer.Print "(For Customers When input tax credit is not required)[See Rule 58(10)]"
    
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("TAX INVOICE")) / 2)
    Printer.CurrentY = 1400
    Printer.Print "TAX INVOICE"

    Printer.FontSize = 10
    Printer.FontUnderline = False
    Printer.CurrentX = x
    y = y + 1100
    Printer.CurrentY = y
    Printer.Print "Invoice No"
    
    x = x + 1100
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print ": "
    
    x = x + 200
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print Trim(TTransactionNo.Text)
    
    x = x + 6500
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.FontUnderline = False
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Customer"

    x = x + 1000
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print ": "

    x = x + 200
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print Trim(CoCustomer.Text)
    
    D = Trim(Day(DTPDate))
    M = Trim(Month(DTPDate))
    YR = Trim(Year(DTPDate))
    If Len(D) = 1 Then D = "0" & D
    If Len(M) = 1 Then M = "0" & M
    DT1 = D & "-" & M & "-" & YR
    
    x = 600
    y = y + 200
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Date"
    
    x = x + 900
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print ": "
    
    x = x + 200
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print Trim(DT1)
    
    x = x + 6500
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Address"
    
    x = x + 1000
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print ": "
    
    x = x + 100
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print Trim(TAddress.Text)
   
    x = 500
    y = y + 1600
 
    
    Printer.FontBold = True
    Printer.FontSize = 9
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "SNo"

    x = 100 + 1000
    Printer.FontSize = 9
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Particulars"

    x = 100 + 5200
    Printer.FontSize = 9
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Tax % "

    x = 100 + 5900
    Printer.FontSize = 9
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Rate"

    x = 100 + 7200
    Printer.FontSize = 9
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Qty"

'    x = 100 + 7100
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "GR Value"

    x = 100 + 7900
    Printer.FontSize = 9
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Disc"

'    x = 100 + 9000
'    Printer.FontSize = 9
'    Printer.CurrentX = x
'    Printer.CurrentY = y
'    Printer.Print "Net Amt"
    
    x = 100 + 8900
    Printer.FontSize = 9
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Tax Amt"
    
    x = 100 + 9900
    Printer.FontSize = 9
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Total"
    
    'HORIZONTAL LINES
    Printer.Line (400, 3600)-(11000, 3600)
    Printer.Line (400, 4000)-(11000, 4000)
    Printer.Line (400, 10000)-(11000, 10000)
    Printer.Line (400, 9700)-(11000, 9700)

    
    'FIRST AND LAST VERTICAL LINE
    Printer.Line (400, 3600)-(400, 10000)
    Printer.Line (11000, 3600)-(11000, 10000)
    
    'INNER LINES
    Printer.Line (1000, 3600)-(1000, 10000)
    Printer.Line (5200, 3600)-(5200, 10000)
    Printer.Line (5900, 3600)-(5900, 10000)
    Printer.Line (7200, 3600)-(7200, 10000)
    Printer.Line (7900, 3600)-(7900, 10000)
    Printer.Line (8900, 3600)-(8900, 10000)
    Printer.Line (9900, 3600)-(9900, 10000)
'    Printer.Line (10000, 3600)-(10000, 10000)
'    Printer.Line (10900, 3600)-(10900, 10000)
    
    

    
    Printer.FontSize = 10
    Printer.FontItalic = True
    Printer.FontBold = False
    Printer.CurrentY = 11600
    Printer.CurrentX = 1000
    Printer.Print (NumberToWords(Val(LGrandAmount.Caption & "")))
    Printer.FontItalic = False
'    Print #1, Chr(27) & "!" & Chr(4) & "|Amount in Words:" & Left(NumberToWords(Val(LGrandAmount.Caption & "")) & Space(66), 66) & " Balance                                |" & Chr(0) & Chr(27) & "!" & Chr(29) & Right(Space(13) & Format("0" & LBalance.Caption, "0.00"), 13) & "|" & Chr(0)
    
    Printer.FontSize = 10
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Declaration")) / 2)
    Printer.CurrentY = 12600
    Printer.Print "Declaration"
    Printer.FontBold = False
    Printer.FontUnderline = False

    Declaration(0) = "DECLARATION : Certified that all the particulars shown in the above Tax Invoice are true and correct and that my/our registration under"
    Declaration(1) = "KVAT ACT is valid as on the date of this bill"
  
    
TOPH = 200
    
    For i = 0 To 2
        Printer.FontSize = 9
        Printer.CurrentX = 550
        Printer.CurrentY = Printer.CurrentY + TOPH
        If i = 2 Then
            Printer.Print Declaration(i)
        Else
            Printer.Print Declaration(i);
        End If
    Next
NewPage = y
End Function

