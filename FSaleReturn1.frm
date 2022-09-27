VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FSaleReturn 
   Caption         =   "FServiceReturn"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   Picture         =   "FSaleReturn1.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   14580
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3285
      Left            =   135
      TabIndex        =   19
      Top             =   1560
      Width           =   14275
      _ExtentX        =   25188
      _ExtentY        =   5794
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
   Begin VB.CommandButton CDelete 
      Height          =   500
      Left            =   4770
      Picture         =   "FSaleReturn1.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   500
      Left            =   11925
      Picture         =   "FSaleReturn1.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7335
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   500
      Left            =   10485
      Picture         =   "FSaleReturn1.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7335
      Width           =   1365
   End
   Begin VB.CommandButton CPrint 
      Height          =   500
      Left            =   1800
      Picture         =   "FSaleReturn1.frx":205968
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7410
      Width           =   1365
   End
   Begin VB.CommandButton CNew 
      Height          =   500
      Left            =   345
      Picture         =   "FSaleReturn1.frx":207DCA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7410
      Width           =   1365
   End
   Begin VB.CommandButton CClear 
      Height          =   500
      Left            =   3210
      Picture         =   "FSaleReturn1.frx":20A22C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6735
      Width           =   1365
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   500
      Left            =   1770
      Picture         =   "FSaleReturn1.frx":20C68E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6735
      Width           =   1365
   End
   Begin VB.CommandButton CAddItem 
      Height          =   500
      Left            =   330
      Picture         =   "FSaleReturn1.frx":20EAF0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6735
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   420
      Left            =   3090
      TabIndex        =   20
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
      Format          =   53608451
      CurrentDate     =   40544
   End
   Begin MSFlexGridLib.MSFlexGrid MGridItemDetails 
      Height          =   795
      Left            =   210
      TabIndex        =   5
      Top             =   5910
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
      Appearance      =   0
   End
   Begin MSForms.TextBox TRefferenceNo 
      Height          =   375
      Left            =   5850
      TabIndex        =   60
      Top             =   720
      Width           =   1590
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2805;661"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   375
      Left            =   4920
      TabIndex        =   59
      Top             =   720
      Width           =   795
      VariousPropertyBits=   8388627
      Caption         =   "Reff .No"
      Size            =   "1402;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   4560
      TabIndex        =   55
      Top             =   5910
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
      Left            =   6045
      TabIndex        =   54
      Top             =   5925
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
      Height          =   435
      Left            =   6240
      TabIndex        =   53
      Top             =   1245
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
   Begin MSForms.Label LUnit 
      Height          =   330
      Left            =   6120
      TabIndex        =   52
      Top             =   6360
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
      Left            =   1470
      TabIndex        =   1
      Top             =   630
      Width           =   3180
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "5609;741"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   300
      TabIndex        =   51
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
   Begin MSForms.Label Label16 
      Height          =   435
      Left            =   5400
      TabIndex        =   49
      Top             =   1245
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
   Begin MSForms.Label Label14 
      Height          =   285
      Left            =   1950
      TabIndex        =   48
      Top             =   1245
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
   Begin MSForms.Label Label13 
      Height          =   330
      Index           =   0
      Left            =   165
      TabIndex        =   47
      Top             =   1245
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
   Begin MSForms.Label LGrandAmount 
      Height          =   570
      Left            =   10785
      TabIndex        =   46
      Top             =   6615
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
   Begin MSForms.TextBox TAddress 
      Height          =   420
      Left            =   11145
      TabIndex        =   3
      Top             =   645
      Width           =   3210
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "5662;741"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoCustomer 
      Height          =   420
      Left            =   11145
      TabIndex        =   2
      Top             =   135
      Width           =   3210
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5662;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   10485
      TabIndex        =   45
      Top             =   195
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
   Begin MSForms.TextBox TTransactionNo 
      Height          =   420
      Left            =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1590
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "2805;741"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   315
      TabIndex        =   44
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
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   120
      X2              =   14445
      Y1              =   4845
      Y2              =   4845
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   11610
      TabIndex        =   43
      Top             =   1245
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
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   8880
      TabIndex        =   42
      Top             =   1245
      Width           =   1095
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "G.Value"
      Size            =   "1931;582"
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalTaxAmount 
      Height          =   375
      Left            =   11835
      TabIndex        =   41
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
   Begin MSForms.Label LTotalGrossValue 
      Height          =   375
      Left            =   9060
      TabIndex        =   40
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
   Begin MSForms.Label LGrandTotalAmount 
      Height          =   375
      Left            =   12540
      TabIndex        =   39
      Top             =   5655
      Width           =   1455
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Total Tax"
      Size            =   "2566;661"
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
      Left            =   225
      TabIndex        =   38
      Top             =   5085
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
      Left            =   1020
      TabIndex        =   4
      Top             =   5010
      Width           =   4560
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "8043;767"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TQuantity 
      Height          =   435
      Left            =   5595
      TabIndex        =   6
      Top             =   5010
      Width           =   885
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "1561;767"
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalAmount 
      Height          =   435
      Left            =   12570
      TabIndex        =   37
      Top             =   5055
      Width           =   1560
      ForeColor       =   0
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2752;767"
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
      Left            =   6480
      TabIndex        =   7
      Top             =   5010
      Width           =   975
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "1720;767"
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label11 
      Height          =   435
      Left            =   7335
      TabIndex        =   36
      Top             =   1245
      Width           =   810
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Warranty"
      Size            =   "1429;767"
      BorderColor     =   -2147483641
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TWarranty 
      Height          =   435
      Left            =   7590
      TabIndex        =   8
      Top             =   5010
      Width           =   795
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "1402;767"
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label12 
      Height          =   420
      Left            =   4935
      TabIndex        =   35
      Top             =   6315
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
   Begin MSForms.Label Label17 
      Height          =   435
      Left            =   8265
      TabIndex        =   34
      Top             =   1245
      Width           =   585
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Tax%"
      Size            =   "1032;767"
      FontName        =   "Sylfaen"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TTax 
      Height          =   435
      Left            =   8400
      TabIndex        =   9
      Top             =   5010
      Width           =   585
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "1032;767"
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label22 
      Height          =   330
      Left            =   10815
      TabIndex        =   33
      Top             =   1245
      Width           =   1020
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Net Value"
      Size            =   "1799;582"
      FontName        =   "Sylfaen"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label23 
      Height          =   330
      Left            =   9900
      TabIndex        =   32
      Top             =   1245
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
   Begin MSForms.Label LGrossValue 
      Height          =   360
      Left            =   9015
      TabIndex        =   31
      Top             =   5055
      Width           =   1095
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "iiiiiiiiiiiiiiiii"
      Size            =   "1931;635"
      FontName        =   "Arial Narrow"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LNetValue 
      Height          =   360
      Left            =   10980
      TabIndex        =   30
      Top             =   5070
      Width           =   1035
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "iiiiiiiiiiiiiiiii"
      Size            =   "1826;635"
      FontName        =   "Arial Narrow"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTaxAmount 
      Height          =   360
      Left            =   11685
      TabIndex        =   29
      Top             =   5055
      Width           =   1095
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "iiiiiiiiiiiiiiiii"
      Size            =   "1931;635"
      FontName        =   "Arial Narrow"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TItemDiscount 
      Height          =   435
      Left            =   9945
      TabIndex        =   10
      Top             =   5040
      Width           =   1035
      VariousPropertyBits=   746604571
      BackColor       =   16777215
      BorderStyle     =   1
      Size            =   "1826;767"
      BorderColor     =   4210752
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   345
      Left            =   2910
      TabIndex        =   28
      Top             =   5625
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
   Begin MSForms.Label Label13 
      Height          =   330
      Index           =   1
      Left            =   510
      TabIndex        =   27
      Top             =   5625
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
   Begin MSForms.Label Label25 
      Height          =   345
      Left            =   1380
      TabIndex        =   26
      Top             =   5625
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
   Begin MSForms.Label LPurchaseRate 
      Height          =   345
      Left            =   4725
      TabIndex        =   25
      Top             =   6735
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
   Begin MSForms.Label LBatch 
      Height          =   345
      Left            =   4665
      TabIndex        =   24
      Top             =   7425
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
   Begin MSForms.Label LMRP 
      Height          =   345
      Left            =   4695
      TabIndex        =   23
      Top             =   7110
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
   Begin MSForms.Label LCurrentStock 
      Height          =   420
      Left            =   4815
      TabIndex        =   22
      Top             =   7725
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
   Begin MSForms.Label LMFRShortName 
      Height          =   345
      Left            =   7560
      TabIndex        =   21
      Top             =   5925
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
   Begin MSForms.Label Label20 
      Height          =   495
      Left            =   120
      TabIndex        =   56
      Top             =   1155
      Width           =   13035
      BackColor       =   15724527
      VariousPropertyBits=   8388627
      Size            =   "22992;873"
      Picture         =   "FSaleReturn1.frx":210F52
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label19 
      Height          =   330
      Left            =   12645
      TabIndex        =   50
      Top             =   1245
      Width           =   1305
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2302;582"
      FontName        =   "Sylfaen"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   4440
      Left            =   120
      Top             =   1170
      Width           =   14325
   End
   Begin MSForms.Label Label2 
      Height          =   495
      Left            =   1390
      TabIndex        =   61
      Top             =   1155
      Width           =   13035
      BackColor       =   15724527
      VariousPropertyBits=   8388627
      Size            =   "22992;873"
      Picture         =   "FSaleReturn1.frx":227C14
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label21 
      Height          =   495
      Left            =   1125
      TabIndex        =   57
      Top             =   1155
      Width           =   13350
      BackColor       =   15724527
      Size            =   "23548;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label18 
      Height          =   495
      Left            =   150
      TabIndex        =   58
      Top             =   1140
      Width           =   13410
      BackColor       =   15724527
      Size            =   "23654;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FSaleReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim dMRP As Double, dQuantity As Double, dBatchMRP() As Double, dBatchQuantity() As Double
Dim sCustomerCode() As String, sCustomerAddress() As String, sAccountCode() As String
Dim sItemCode() As String, sBillingName() As String
Dim gIQuantity As Single, gIMRP As Single, gIPurchaseRate As Single, gIPurchaseCode As Single, gIBatch As Single
Dim gSerialNo As Single, gItem As Single, gPurchaseRate As Single, gQuantity As Single, gTaxAmount As Single, gTax As Single, gNetValue As Single, gItemDiscount As Single, gBatch As Single, gWarranty As Single, gGrossValue As Single, gUnit As Single, gSaleRate As Single, gMRP As Single, gTotalAmount As Single, gBillingName As Single, gItemCode As Single, gMFRShortName As Single

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
       
'    If Val(TQuantity.Text) > Val(LCurrentStock.Caption) Then
'        lYN = MsgBox("There is no enough Stock, Do you want to Continue ?", vbDefaultButton2 Or vbYesNo)
'        If lYN = vbYes Then
'        Else
'            TQuantity.SetFocus
'            Exit Sub
'        End If
'    End If
    
    If Val(TRate.Text) = 0 Then
        lYN = MsgBox("Rate given is Zero, Do you want to Continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TRate.SetFocus
            Exit Sub
        End If
    End If

    If Val(TRate.Text) < Val(LMRP.Caption) Then
        lYN = MsgBox("Rate given is less than MRP, Do you Want to continue ?", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
            TRate.SetFocus
            Exit Sub
        End If
    End If
        
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
        MGrid.TextMatrix(MGrid.Rows - 1, gTotalAmount) = Format((Val(MGrid.TextMatrix(MGrid.Rows - 1, gNetValue)) + Val(MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount))), "0.00")
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
        MGrid.TextMatrix(r - 1, gTotalAmount) = Format((Val(MGrid.TextMatrix(r - 1, gNetValue)) + Val(MGrid.TextMatrix(r - 1, gTaxAmount))), "0.00")
        MGrid.TextMatrix(r - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gMFRShortName) = LMFRShortName.Caption & ""
        MGrid.TextMatrix(r - 1, gMRP) = Val(LMRP.Caption)
        MGrid.TextMatrix(r - 1, gPurchaseRate) = Val(LPurchaseRate.Caption)
    End If
    
    clearEditControls
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    CoItem.SetFocus
End Sub

Private Sub CClear_Click()
    MGrid.Rows = 0
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
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
    gQuantity = 3
    gUnit = 4
    gSaleRate = 5
    gWarranty = 6
    gTax = 7
    gGrossValue = 8
    gItemDiscount = 9
    gNetValue = 10
    gTaxAmount = 11
    gTotalAmount = 12
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
    MGrid.ColWidth(gNetValue) = 1160
    MGrid.ColWidth(gGrossValue) = 1160
    MGrid.ColWidth(gTotalAmount) = 1160
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
    
    Set rs = db.OpenRecordset("Select Max(Val( Transaction.TransactionNo)) As TNo From Transaction Where ( Transaction.TransactionType = 'SR' )")
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
'    If (Val("" & rs!InStock) - Val("" & rs!OutStock) - dQuantity) > 0 Then
        MGridItemDetails.AddItem Val("" & rs!InStock) - Val("" & rs!OutStock) - dQuantity & vbTab & rs!MRP & vbTab & rs!PurchaseRate & vbTab & rs!Batch
        TQuantity.Text = MGridItemDetails.TextMatrix(0, gIQuantity)
        LCurrentStock.Caption = MGridItemDetails.TextMatrix(0, gIQuantity)
        TRate.Text = MGridItemDetails.TextMatrix(0, gIMRP)
        LMRP.Caption = MGridItemDetails.TextMatrix(0, gIMRP)
        LBatch.Caption = MGridItemDetails.TextMatrix(0, gIBatch)
        LPurchaseRate.Caption = MGridItemDetails.TextMatrix(0, gIPurchaseRate)
'     End If
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
    DTPDate.Value = Date
    TNarration.Text = ""
    CoCustomer.Text = ""
    TAddress.Text = ""
    TWarranty.Text = ""
    MGrid.Rows = 0
    MGridItemDetails.Rows = 0
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TQuantity.Text = ""
    TTax.Text = ""
    LUnit.Caption = ""
    TRate.Text = ""
    LTotalAmount.Caption = ""
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    getTotal
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
    LTotalAmount.Caption = ""
    TWarranty.Text = 0
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
        dGrandTotal = dGrandTotal + Val(MGrid.TextMatrix(r, gTotalAmount))
        dTax = dTax + Val(MGrid.TextMatrix(r, gTaxAmount))
        dGrossValue = dGrossValue + Val(MGrid.TextMatrix(r, gGrossValue))
        r = r + 1
    Wend
    getGrandTotal = dGrandTotal
    LTotalGrossValue.Caption = Format(dGrossValue, "0.00")
    LGrandTotalAmount.Caption = Format(dGrandTotal, "0.00")
    LTotalTaxAmount.Caption = Format(dTax, "0.00")
End Function

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean
    bFound = False
    If (MsgBox("Do you want to Delete the Bill ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'SR' )")
        While rs.EOF = False
            bFound = True
            rs.Delete
            rs.MoveNext
        Wend
        rs.Close
        
        If bFound Then
            MsgBox "Successfully Deleted !", vbInformation
            clearControls
            TTransactionNo.Text = getNewTransactionNo
        Else
            MsgBox "Bill Not Found !", vbInformation
        End If
    End If
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
        FCustomerMaster.Show vbModal
        getCustomer
    End If
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
    
    If CoCustomer.ListIndex = -1 Then
        lYN = MsgBox("Do you want to consider General Customer !", vbDefaultButton2 Or vbYesNo)
        If lYN = vbYes Then
        Else
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
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'SR' )")
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
        rs!TransactionType = "SR"
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
        rs!MRP = 0
        rs!ReferenceNo = ""
        rs!Batch = Trim(MGrid.TextMatrix(r, gBatch))
        rs!Warranty = Trim(MGrid.TextMatrix(r, gWarranty))
        rs!MRP = Trim(MGrid.TextMatrix(r, gMRP))
        rs!PurchaseRate = Trim(MGrid.TextMatrix(r, gPurchaseRate))
        rs!ReferenceDate = Date
        rs!Tax = Val(MGrid.TextMatrix(r, gTax))
        rs!ItemDiscount = Val(MGrid.TextMatrix(r, gItemDiscount))
        rs.Update
        r = r + 1
    Wend
    rs.Close
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
    ElseIf (KeyCode = vbKeyA And ((Shift And 7) = 2)) Then
        CAddItem_Click
    ElseIf (KeyCode = vbKeyR And ((Shift And 7) = 2)) Then
        CRemoveItem_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClear_Click
    End If
End Sub

Private Sub Form_Load()

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
        LTotalAmount.Caption = Val(MGrid.TextMatrix(r, gTotalAmount))
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
    LGrossValue.Caption = Val(TRate.Text) * Val(TQuantity.Text)
    LNetValue = (Val(TRate.Text) * Val(TQuantity.Text)) - Val(TItemDiscount.Text)
    LTaxAmount = Val(LNetValue.Caption) * Val(TTax.Text) / 100
    LTotalAmount = Val(LNetValue.Caption) + Val(LTaxAmount.Caption)
End Sub

Private Sub MGridItemDetails_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        If (MGridItemDetails.Rows > 0) Then
            TQuantity.Text = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIQuantity)
            LCurrentStock.Caption = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIQuantity)
            TRate.Text = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIMRP)
            LMRP.Caption = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIMRP)
            LBatch.Caption = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIBatch)
            LPurchaseRate.Caption = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIPurchaseRate)
        End If
        SendKeys "{TAB}"
    End If
End Sub





Private Sub TAddress_GotFocus()
    TAddress.SelStart = 0
    TAddress.SelLength = Len(TAddress.Text)
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



Private Sub TRefferenceNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        clearControls
        getTransactionDetailsByRefference
        TTransactionNo.Text = getNewTransactionNo
'        LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    End If
End Sub
Private Sub getTransactionDetailsByRefference()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.BillingName,ItemMaster.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & Trim(TRefferenceNo.Text) & "' ) And (Transaction.TransactionType = 'S' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode) Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        CoCustomer.Text = "" & rs!CustomerName
        TAddress.Text = "" & rs!CustomerAddress
        TNarration.Text = "" & rs!Narration
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gQuantity) = "" & rs!Quantity
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gBatch) = "" & rs!Batch
            MGrid.TextMatrix(r, gTax) = "" & rs!Tax
            MGrid.TextMatrix(r, gWarranty) = "" & rs!Warranty
            MGrid.TextMatrix(r, gSaleRate) = Format("" & rs!SaleRate, "0.00")
            MGrid.TextMatrix(r, gGrossValue) = Format(Val("" & rs!Quantity) * Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gItemDiscount) = Format("" & rs!ItemDiscount, "0.00")
            MGrid.TextMatrix(r, gNetValue) = Format(Val(MGrid.TextMatrix(r, gGrossValue)) - Val("" & rs!ItemDiscount), "0.00")
            MGrid.TextMatrix(r, gTaxAmount) = Format(Val(MGrid.TextMatrix(r, gNetValue)) * rs!Tax / 100, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format(Val(MGrid.TextMatrix(r, gNetValue) + Val(MGrid.TextMatrix(r, gTaxAmount))), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gMFRShortName) = "" & rs!ShortName
            MGrid.TextMatrix(r, gMRP) = "" & rs!MRP
            MGrid.TextMatrix(r, gPurchaseRate) = "" & rs!PurchaseRate
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
    
    LSlNo.Caption = MGrid.Rows + 1
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")

    
End Sub
Private Sub TTax_Change()
    TTax.SelStart = 0
    TTax.SelLength = Len(TTax.Text)
    getTotal
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

Private Sub getTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.BillingName,ItemMaster.Code,Transaction.Tax As ItemTax,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'SR' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode) Order By Transaction.SerialNo")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!TransactionDate
        CoCustomer.Text = "" & rs!CustomerName
        TAddress.Text = "" & rs!CustomerAddress
        TNarration.Text = "" & rs!Narration
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
            MGrid.TextMatrix(r, gTotalAmount) = Format(Val(MGrid.TextMatrix(r, gNetValue) + Val(MGrid.TextMatrix(r, gTaxAmount))), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gMFRShortName) = "" & rs!ShortName
            MGrid.TextMatrix(r, gMRP) = "" & rs!MRP
            MGrid.TextMatrix(r, gPurchaseRate) = "" & rs!PurchaseRate
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
    
    LSlNo.Caption = MGrid.Rows + 1
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    
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

