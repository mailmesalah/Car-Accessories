VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FWholeSales 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local Sales  (Whole Sales)"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FWholeSales.frx":0000
   ScaleHeight     =   8670
   ScaleWidth      =   11745
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CFirst 
      Height          =   435
      Left            =   5940
      Picture         =   "FWholeSales.frx":19D696
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8940
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton CNext 
      Height          =   435
      Left            =   7320
      Picture         =   "FWholeSales.frx":19F690
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8940
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton CLast 
      Height          =   435
      Left            =   10065
      Picture         =   "FWholeSales.frx":1A168A
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   8940
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton CPrevious 
      Height          =   435
      Left            =   8685
      Picture         =   "FWholeSales.frx":1A3684
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8940
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox TPayment 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8475
      TabIndex        =   44
      Top             =   6315
      Width           =   2850
   End
   Begin VB.TextBox TQuantity 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7680
      TabIndex        =   5
      Text            =   " "
      Top             =   4800
      Width           =   1320
   End
   Begin VB.TextBox TRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9045
      TabIndex        =   6
      Text            =   " "
      Top             =   4800
      Width           =   1320
   End
   Begin VB.TextBox TPartNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5640
      TabIndex        =   42
      Text            =   " "
      Top             =   4800
      Width           =   1320
   End
   Begin VB.CommandButton CNew 
      Height          =   450
      Left            =   1095
      Picture         =   "FWholeSales.frx":1A567E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8070
      Width           =   1485
   End
   Begin VB.CommandButton CPrint 
      Height          =   450
      Left            =   2550
      Picture         =   "FWholeSales.frx":1AAC7C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8070
      Width           =   1485
   End
   Begin VB.CommandButton CSave 
      Height          =   450
      Left            =   7875
      Picture         =   "FWholeSales.frx":1B027A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8070
      Width           =   1485
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   450
      Left            =   9330
      Picture         =   "FWholeSales.frx":1B5878
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8070
      Width           =   1485
   End
   Begin VB.CommandButton CAddItem 
      Height          =   450
      Left            =   270
      Picture         =   "FWholeSales.frx":1BAE76
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7515
      Width           =   1485
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   450
      Left            =   1710
      Picture         =   "FWholeSales.frx":1C0474
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7515
      Width           =   1485
   End
   Begin VB.CommandButton CClear 
      Height          =   450
      Left            =   3150
      Picture         =   "FWholeSales.frx":1C5A72
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7515
      Width           =   1485
   End
   Begin VB.CommandButton CDelete 
      Height          =   450
      Left            =   4995
      Picture         =   "FWholeSales.frx":1CB070
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   75
      Width           =   1485
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   330
      Left            =   2820
      TabIndex        =   15
      Top             =   180
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
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
      Format          =   54394883
      CurrentDate     =   40544
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3285
      Left            =   150
      TabIndex        =   27
      Top             =   1380
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   5794
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MGridItemDetails 
      Height          =   795
      Left            =   240
      TabIndex        =   58
      Top             =   5685
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
   Begin MSForms.Label Label8 
      Height          =   270
      Left            =   0
      TabIndex        =   60
      Top             =   5400
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "QTY"
      Size            =   "2593;476"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   270
      Left            =   1560
      TabIndex        =   61
      Top             =   5400
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "MRP"
      Size            =   "2593;476"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label15 
      Height          =   270
      Left            =   2970
      TabIndex        =   59
      Top             =   5400
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "P   R"
      Size            =   "2593;476"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   3135
      TabIndex        =   50
      Top             =   7500
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   1695
      TabIndex        =   57
      Top             =   7500
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   255
      TabIndex        =   56
      Top             =   7500
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   1080
      TabIndex        =   55
      Top             =   8055
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2535
      TabIndex        =   54
      Top             =   8055
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   7860
      TabIndex        =   53
      Top             =   8055
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   9315
      TabIndex        =   52
      Top             =   8055
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   4980
      TabIndex        =   51
      Top             =   60
      Width           =   1530
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   4365
      Left            =   120
      Top             =   990
      Width           =   11520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   135
      X2              =   11595
      Y1              =   4665
      Y2              =   4680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   120
      X2              =   11610
      Y1              =   1365
      Y2              =   1365
   End
   Begin MSForms.Label LBalance 
      Height          =   330
      Left            =   8475
      TabIndex        =   43
      Top             =   6825
      Width           =   2850
      BackColor       =   -2147483643
      Size            =   "5027;573"
      BorderStyle     =   1
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label LMFRShortName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3345
      TabIndex        =   41
      Top             =   6465
      Width           =   1650
   End
   Begin MSForms.Label Label18 
      Height          =   330
      Left            =   210
      TabIndex        =   40
      Top             =   7125
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Rack"
      Size            =   "2593;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LRack 
      Height          =   330
      Left            =   1695
      TabIndex        =   39
      Top             =   7125
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   210
      TabIndex        =   38
      Top             =   6465
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Manufacturer"
      Size            =   "2593;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LManufacturer 
      Height          =   330
      Left            =   1665
      TabIndex        =   37
      Top             =   6465
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LCurrentStock 
      Height          =   330
      Left            =   1695
      TabIndex        =   36
      Top             =   6795
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   330
      Left            =   210
      TabIndex        =   35
      Top             =   6795
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Current Stock"
      Size            =   "2593;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoItem 
      CausesValidation=   0   'False
      Height          =   420
      Left            =   885
      TabIndex        =   4
      Top             =   4800
      Width           =   4545
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "8017;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LTotalAmount 
      Height          =   390
      Left            =   10425
      TabIndex        =   34
      Top             =   4800
      Width           =   1140
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2011;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Index           =   0
      Left            =   270
      TabIndex        =   33
      Top             =   1035
      Width           =   555
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Sl No"
      Size            =   "979;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1395
      TabIndex        =   32
      Top             =   1005
      Width           =   3480
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Item"
      Size            =   "6138;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   5820
      TabIndex        =   31
      Top             =   1035
      Width           =   1320
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Part No"
      Size            =   "2328;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   330
      Left            =   7095
      TabIndex        =   30
      Top             =   1035
      Width           =   1170
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LUnit 
      Height          =   330
      Left            =   7020
      TabIndex        =   29
      Top             =   4800
      Width           =   600
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "1058;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LSlNo 
      Height          =   420
      Left            =   285
      TabIndex        =   28
      Top             =   4800
      Width           =   555
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "SLNo"
      Size            =   "979;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBalancelb 
      Height          =   330
      Left            =   7635
      TabIndex        =   26
      Top             =   6840
      Width           =   720
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "1270;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LPayment 
      Height          =   330
      Left            =   7545
      TabIndex        =   25
      Top             =   6360
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Payment"
      Size            =   "1508;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Left            =   4500
      TabIndex        =   24
      Top             =   585
      Width           =   495
      VariousPropertyBits=   8388627
      Caption         =   "Title"
      Size            =   "873;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoTitle 
      Height          =   330
      Left            =   4980
      TabIndex        =   2
      Top             =   585
      Width           =   2025
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3572;573"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   405
      TabIndex        =   23
      Top             =   180
      Width           =   465
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "820;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   180
      Width           =   1590
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2805;573"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   7800
      TabIndex        =   22
      Top             =   180
      Width           =   375
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "661;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoCustomer 
      Height          =   330
      Left            =   8340
      TabIndex        =   3
      Top             =   180
      Width           =   3210
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5662;573"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress 
      Height          =   330
      Left            =   8340
      TabIndex        =   8
      Top             =   585
      Width           =   3210
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5662;573"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LGrandAmount 
      Height          =   570
      Left            =   7935
      TabIndex        =   21
      Top             =   5595
      Width           =   3795
      VariousPropertyBits=   8388627
      Caption         =   "Grand Amount"
      Size            =   "6694;1005"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   180
      TabIndex        =   20
      Top             =   585
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "1508;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TNarration 
      Height          =   330
      Left            =   1200
      TabIndex        =   1
      Top             =   585
      Width           =   3180
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5609;573"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label10 
      Height          =   330
      Left            =   8160
      TabIndex        =   19
      Top             =   1035
      Width           =   840
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Qty"
      Size            =   "1482;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label19 
      Height          =   330
      Left            =   9810
      TabIndex        =   18
      Top             =   1020
      Width           =   1305
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2302;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label16 
      Height          =   330
      Left            =   8880
      TabIndex        =   17
      Top             =   1035
      Width           =   1170
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Rate"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label12 
      BackColor       =   &H00404040&
      Height          =   375
      Index           =   9
      Left            =   135
      TabIndex        =   49
      Top             =   1005
      Width           =   11475
   End
End
Attribute VB_Name = "FWholeSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCustomerCode() As String, sCustomerAddress() As String
Dim sItemCode() As String, sBillingName() As String, sAccountCode() As String, sPartNo() As String
Dim FstFlag As Boolean, NxtFlag As Boolean, PrvsFlag As Boolean, LstFlag As Boolean
Dim gSerialNo As Single, gItem As Single, gPartNo As Single, gIQuantity As Single, gPurchaseRate As Single, gARate As Single, gBRate As Single, gCRate As Single, gIMRP As Single, gQuantity As Single, gUnit As Single, gSaleRate As Single, gMRP As Single, gTotalAmount As Single, gBillingName As Single, gItemCode As Single, gMFRShortName As Single
Private Sub MGridItemDetailsInitialise()
'INITIALISES MGridItemDetails
        'SETTING CONSTANTS
    gIQuantity = 0
    gIMRP = 1
    gPurchaseRate = 2
    
    MGridItemDetails.Clear
    MGridItemDetails.Rows = 1 'FOR SKIPING ERROR
    MGridItemDetails.Cols = 1 'FOR SKIPING ERROR
    MGridItemDetails.FixedCols = 0
    MGridItemDetails.FixedRows = 0
    MGridItemDetails.Cols = 3
    MGridItemDetails.Rows = 0
    MGridItemDetails.ColWidth(gIQuantity) = 1270
    MGridItemDetails.ColWidth(gIMRP) = 1400
    MGridItemDetails.ColWidth(gPurchaseRate) = 1400
    MGridItemDetails.RowHeightMin = 350
End Sub


Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

    If Trim(CoItem.Text) = "" Then
        MsgBox "Please Select a Item !", vbInformation
        CoItem.SetFocus
        Exit Sub
    End If
    If Val(TQuantity.Text) = 0 Then
        MsgBox "Please Enter Quantity greater than Zero !", vbInformation
        TQuantity.SetFocus
        Exit Sub
    End If
    
'    If (CoBatch.ListIndex = -1) Then
'        MsgBox "Please select a valid Batch !", vbInformation
'    ElseIf Val(TQuantity.Text) > dBatchQuantity(CoBatch.ListIndex + 1) Then
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
    
'    If Val(TRate.Text) < dBatchWholeSaleRate(CoBatch.ListIndex + 1) Then
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
        MGrid.TextMatrix(MGrid.Rows - 1, gPartNo) = Trim(TPartNo.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gTotalAmount) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gBillingName) = IIf(CoItem.ListIndex = -1, CoItem.Text, sBillingName(CoItem.ListIndex + 1))
        MGrid.TextMatrix(MGrid.Rows - 1, gItemCode) = IIf(CoItem.ListIndex = -1, "0", sItemCode(CoItem.ListIndex + 1))
        MGrid.TextMatrix(MGrid.Rows - 1, gMFRShortName) = LMFRShortName.Caption & ""
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(r - 1, gPartNo) = Trim(TPartNo.Text)
        MGrid.TextMatrix(r - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(r - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(r - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(r - 1, gTotalAmount) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(r - 1, gBillingName) = IIf(CoItem.ListIndex = -1, CoItem.Text, sBillingName(CoItem.ListIndex + 1))
        MGrid.TextMatrix(r - 1, gItemCode) = IIf(CoItem.ListIndex = -1, "0", sItemCode(CoItem.ListIndex + 1))
        MGrid.TextMatrix(r - 1, gMFRShortName) = LMFRShortName.Caption & ""
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
    gPartNo = 2
    gUnit = 3
    gQuantity = 4
    gSaleRate = 5
    gTotalAmount = 6
    gBillingName = 7
    gItemCode = 8
    gMFRShortName = 9
    
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 10
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 890
    MGrid.ColWidth(gItem) = 4940
    MGrid.ColWidth(gPartNo) = 1150
    MGrid.ColWidth(gQuantity) = 1000
    MGrid.ColWidth(gUnit) = 730
    MGrid.ColWidth(gSaleRate) = 1160
    MGrid.ColWidth(gTotalAmount) = 1160
    MGrid.ColWidth(gBillingName) = 0
    MGrid.ColWidth(gItemCode) = 0
    MGrid.ColWidth(gMFRShortName) = 0
    MGrid.RowHeightMin = 350
End Sub

Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String
    Set rs = db.OpenRecordset("Select Max(Val( Transaction.TransactionNo)) As TNo From Transaction Where ( Transaction.TransactionType = 'SW' )")
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
        
    Set rs = db.OpenRecordset("Select ItemMaster.Code,ItemMaster.PartNo,ItemMaster.ItemName,ItemMaster.BillingName From ItemMaster Where (ItemMaster.Type = 'BItem' ) Order By ItemMaster.ItemName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sItemCode(rs.RecordCount + 1) As String
    ReDim sBillingName(rs.RecordCount + 1) As String
    ReDim sPartNo(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoItem.AddItem "" & rs!ItemName
        sItemCode(CoItem.ListCount) = "" & rs!Code
        sBillingName(CoItem.ListCount) = "" & rs!BillingName
        sPartNo(CoItem.ListCount) = "" & rs!PartNo
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub getItemDetails()
Dim rs As Recordset, r As Long, dQuantity As Double
    MGridItemDetails.Rows = 0
    TQuantity.Text = ""
    TRate.Text = ""
    If (CoItem.ListIndex = -1) Then
        Exit Sub
    Else
        TPartNo.Text = Trim("" & sPartNo(CoItem.ListIndex + 1))
    End If

    Set rs = db.OpenRecordset("Select Manufacturer.ShortName,Manufacturer.ManufacturerName,Units.UnitName,ItemMaster.ItemName,ItemMaster.Rack,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('O','P','SR','SA') ) And (Transaction.ItemCode = ItemMaster.Code )) As InStock,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('S','SW','PR','F8','FB') ) And (Transaction.ItemCode = ItemMaster.Code )) As OutStock From ItemMaster,Units,Manufacturer Where (ItemMaster.Code = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code = ItemMaster.ManufacturerCode )")
    If rs.RecordCount > 0 Then
        LManufacturer.Caption = "" & rs!ManufacturerName
        LUnit.Caption = "" & rs!UnitName
        LCurrentStock.Caption = Val("" & rs!InStock) - Val("" & rs!OutStock)
        LMFRShortName.Caption = "" & rs!ShortName
        LRack.Caption = "" & rs!Rack
    Else
        LManufacturer.Caption = ""
        LUnit.Caption = ""
        LCurrentStock.Caption = ""
        LMFRShortName.Caption = ""
        LRack.Caption = ""
    End If
    rs.Close
    


    dQuantity = 0
    r = 0
    While r < MGrid.Rows
        If sItemCode(CoItem.ListIndex + 1) = MGrid.TextMatrix(r, gItemCode) Then
            LCurrentStock.Caption = Val(LCurrentStock.Caption) - MGrid.TextMatrix(r, gQuantity)
        End If
        r = r + 1
    Wend
    
    Set rs = db.OpenRecordset("Select (Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('SW') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.TransactionNo = '" & Trim(TTransactionNo.Text) & "')) As Quantity From Transaction ")
    If rs.RecordCount > 0 Then
        LCurrentStock.Caption = Val(LCurrentStock.Caption) + Val("" & rs!Quantity)
    End If
    
    Set rs = db.OpenRecordset("Select Transaction.MRP,Transaction.PurchaseRate,(Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('O','P','SR','SA') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.MRP = Transaction.MRP) And (T.PurchaseRate = Transaction.PurchaseRate) ) As InStock,(Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('S','SW','PR','F8','FB') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.MRP = Transaction.MRP)And (T.PurchaseRate = Transaction.PurchaseRate)) As OutStock From Transaction Where (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) Group By Transaction.MRP,Transaction.PurchaseRate")
    While rs.EOF = False
        If (Val("" & rs!InStock) - Val("" & rs!OutStock) - dQuantity) > 0 Then
            MGridItemDetails.AddItem Val("" & rs!InStock) - Val("" & rs!OutStock) - dQuantity & vbTab & rs!MRP & vbTab & rs!PurchaseRate
        Else
        
        End If
        If MGridItemDetails.Rows <> 0 Then
            TRate.Text = Val(MGridItemDetails.TextMatrix(0, gIMRP))
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub clearControls()
    DTPDate.Value = Date
    TNarration.Text = ""
    CoCustomer.Text = ""
    TAddress.Text = ""
    MGrid.Rows = 0
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TPartNo.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    TRate.Text = ""
    TPayment.Text = ""
    LBalance.Caption = ""
    LTotalAmount.Caption = ""
    
    CoTitle.Clear
    CoTitle.AddItem "ESTIMATE"
    CoTitle.AddItem "KM AUTO SPARES"
    CoTitle.ListIndex = 0

   

 
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
End Sub

Private Sub clearEditControls()
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TPartNo.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    TRate.Text = ""
    LTotalAmount.Caption = ""
    LBalance.Caption = ""
End Sub

Private Function getGrandTotal() As Double
Dim dGrandTotal As Double, r As Long
    
    r = 0
    dGrandTotal = 0
    While r < MGrid.Rows
        dGrandTotal = dGrandTotal + Val(MGrid.TextMatrix(r, gTotalAmount))
        r = r + 1
    Wend
    getGrandTotal = dGrandTotal
End Function

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean

    bFound = False
    If (MsgBox("Do you want to Delete the Bill ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'SW' )")
        While rs.EOF = False
            bFound = True
            rs.Delete
            rs.MoveNext
        Wend
        rs.Close
        'Delete From Account Register
        deleteFromAccountRegister
        
        If bFound Then
            MsgBox "Successfully Deleted !", vbInformation
            clearControls
            TTransactionNo.Text = getNewTransactionNo
        Else
            MsgBox "Bill Not Found !", vbInformation
        End If
    End If
End Sub
Private Sub deleteFromAccountRegister()
Dim rs As Recordset
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (AccountRegister.SpecialAccount In( 'WholeSales','SWBillVoucher' ))")
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub CNew_Click()
    clearControls
    TTransactionNo.Text = getNewTransactionNo
    TTransactionNo.SetFocus
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
    If KeyCode = 13 And CoItem.ListIndex <> -1 Then
        If MGridItemDetails.Rows > 0 Then
            MGridItemDetails.SetFocus
        Else
            TQuantity.SetFocus
        End If
    ElseIf KeyCode = 13 And CoItem.ListIndex = -1 Then
        getPartNoDetails
    End If
    If KeyCode = 27 Then
        TPayment.Text = Format(LGrandAmount.Caption, "#0.00")
        LBalance.Caption = Format(LGrandAmount.Caption, "#0.00")
        TPayment.SetFocus
    End If
End Sub
Private Sub CoItem_LostFocus()
CoItem.Text = Trim(UCase(CoItem.Text))
If CoItem.ListIndex = -1 Then
    TPartNo.Text = "-"
End If
End Sub
Private Sub CoCustomer_Change()
    If CoCustomer.ListIndex <> -1 Then
        TAddress.Text = sCustomerAddress(CoCustomer.ListIndex + 1)
    Else
        TAddress.Text = ""
    End If
End Sub
Private Sub getPartNoDetails()
Dim rs As Recordset, r As Long
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName From ItemMaster Where(ItemMaster.PartNo = '" & Trim(CoItem.Text) & "')")
    If rs.RecordCount > 0 Then
        TPartNo.Text = Trim(CoItem.Text)
        CoItem.Text = "" & rs!ItemName
    End If
    rs.Close
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

Private Sub CPrint_Click()
'    printSale
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
    
    If MGrid.Rows = 0 Then
        MsgBox "No Items Entered !", vbInformation
        CoItem.SetFocus
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
  'SAVES DATA TO Transaction TABLE
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'SW' )")
    If rs.RecordCount > 0 Then  'Edit
         
        'SAVES DATA TO TransactionRegister ReadyMade
        While rs.EOF = False
            rs.Delete
            rs.MoveNext
        Wend
    Else
        TTransactionNo.Text = getNewTransactionNo
    End If
    
    r = 0
    While r < MGrid.Rows
        rs.AddNew
        rs!TransactionNo = Trim(TTransactionNo.Text)
        rs!TransactionType = "SW"
        rs!TransactionDate = DTPDate.Value
        rs!TransactionTime = Format(Time, "HH:MM AMPM")
        rs!Narration = Trim(TNarration.Text)
        rs!CustomerCode = IIf(CoCustomer.ListIndex = -1, "", sCustomerCode(CoCustomer.ListIndex + 1))
        rs!CustomerName = Trim(CoCustomer.Text)
        rs!CustomerAddress = Trim(TAddress.Text)
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!TempItemName = Trim(MGrid.TextMatrix(r, gItem))
        rs!ItemCode = Trim(MGrid.TextMatrix(r, gItemCode))
        rs!Quantity = Val(MGrid.TextMatrix(r, gQuantity))
        rs!SaleRate = Val(MGrid.TextMatrix(r, gSaleRate))
        rs!SalePayment = IIf(TPayment.Text = "", "0", Val(TPayment.Text))
        rs!PurchaseRate = 0
        rs!WholeSaleRate = 0
        rs!PurchasePayment = 0
        rs!ARate = 0
        rs!BRate = 0
        rs!CRate = 0
        rs!MRP = 0
        rs!SupplierCode = ""
        rs!SupplierName = ""
        rs!SupplierAddress = ""
        rs!ReferenceNo = ""
        rs!ReferenceDate = Date
        rs.Update
        r = r + 1
    Wend
    rs.Close
    
    'Add to Accounts Details
    addToAccountRegister
    
    MsgBox "Successfully Saved !", vbInformation
    lYN = MsgBox("Do you want to take Print ?", vbDefaultButton2 Or vbYesNo)
    If lYN = vbYes Then
        printSale
    Else
        
    End If
    
    clearControls
    TTransactionNo.Text = getNewTransactionNo
    TTransactionNo.SetFocus
End Sub
Private Sub addToAccountRegister()
Dim rs As Recordset, sTransactionNo As String
    
    Set rs = db.OpenRecordset("Select Max(Val(AccountRegister.TransactionNo)) As TransactionNo From AccountRegister Where AccountRegister.Type In ('R','P')")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TransactionNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') ) And (AccountRegister.SpecialAccount = 'WholeSales' )And (AccountRegister.Type In ('P','R') )And (AccountRegister.BillNo = '" & TTransactionNo.Text & "' )  Order By AccountRegister.Type ")
    If (rs.RecordCount > 0) Then
        sTransactionNo = "" & rs!TransactionNo
        rs.MoveFirst
    End If
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    rs.AddNew
    rs!TransactionNo = sTransactionNo
    rs!Type = "P"
    rs!TransactionDate = DTPDate.Value
    rs!TransactionTime = Format(Time, "HH:MM AMPM")
    rs!AccountCode = IIf(CoCustomer.ListIndex = -1, Trim(generalCustomerAccountID), sAccountCode(CoCustomer.ListIndex + 1))
    rs!Narration = "- Bill Amount"
    rs!CashOrCredit = "Credit"
    rs!Income = 0
    rs!Expense = Val(LGrandAmount.Caption)
    rs!BillNo = "" & TTransactionNo.Text
    rs!SpecialAccount = "WholeSales"
    rs.Update

    
    rs.AddNew
    rs!TransactionNo = Val(sTransactionNo) + 1
    rs!Type = "R"
    rs!TransactionDate = DTPDate.Value
    rs!TransactionTime = Format(Time, "HH:MM AMPM")
    rs!AccountCode = IIf(CoCustomer.ListIndex = -1, Trim(generalCustomerAccountID), sAccountCode(CoCustomer.ListIndex + 1))
    rs!Narration = CoCustomer.Text & "- Advance"
    rs!CashOrCredit = "Cash"
    rs!Income = Val(TPayment.Text)
    rs!Expense = 0
    rs!BillNo = "" & TTransactionNo.Text
    rs!SpecialAccount = "WholeSales"
    rs.Update
    rs.Close
End Sub
Private Function getNewAccountTNo(sMode As String) As String

    Dim rs As Recordset, sTCode As String
    
    Set rs = db.OpenRecordset("Select Max(val(AccountRegister.TransactionNo))As ACode From AccountRegister Where (AccountRegister.TransactionType = '" & sMode & "' )")
    If rs.RecordCount > 0 Then
        sTCode = Val("" & rs!ACode) + 1
    Else
        sTCode = "1"
    
    End If
    rs.Close
    
    getNewAccountTNo = sTCode

End Function
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
    getCustomer
    getItem
    MGridInitialise
    MGridItemDetailsInitialise
    clearControls
    TTransactionNo.Text = getNewTransactionNo
End Sub

Private Sub MGrid_Click()
Dim r As Long, i As Long

    If MGrid.Rows > 0 Then
        r = MGrid.Row
        LSlNo.Caption = Val(MGrid.TextMatrix(r, gSerialNo))
        CoItem.Text = Trim(MGrid.TextMatrix(r, gItem))
        TPartNo.Text = Trim(MGrid.TextMatrix(r, gPartNo))
        TQuantity.Text = Val(MGrid.TextMatrix(r, gQuantity))
        LUnit.Caption = Trim(MGrid.TextMatrix(r, gUnit))
        TRate.Text = Val(MGrid.TextMatrix(r, gSaleRate))
        LTotalAmount.Caption = Val(MGrid.TextMatrix(r, gTotalAmount))
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
Private Sub TAddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    CoItem.SetFocus
End If
End Sub

Private Sub TPartNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TQuantity.SetFocus
End If
End Sub

Private Sub TPayment_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
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
Private Sub TQuantity_Change()
    getTotal
End Sub

Private Sub TQuantity_GotFocus()
    TQuantity.SelStart = 0
    TQuantity.SelLength = Len(TQuantity.Text)
End Sub

Private Sub TQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TQuantity.Text <> "" Then
    TRate.SetFocus
End If
End Sub

Private Sub TRate_Change()
    getTotal
End Sub

Private Sub TRate_GotFocus()
    TRate.SelStart = 0
    TRate.SelLength = Len(TRate.Text)
End Sub

Private Sub TRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TRate <> "" Then
    CAddItem.SetFocus
End If
End Sub
Private Sub TTransactionNo_GotFocus()
    TTransactionNo.SelStart = 0
    TTransactionNo.SelLength = Len(TTransactionNo.Text)
End Sub

Private Sub TTransactionNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        clearControls
        getTransactionDetails
        LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    End If
End Sub
Private Sub CFirst_Click()
    FstFlag = True
    NxtFlag = False
    PrvsFlag = False
    LstFlag = False
    TTransactionNo.Text = 1
    TransactionDetails
End Sub
Private Sub CNext_Click()
Dim Tbno As String
    FstFlag = False
    NxtFlag = True
    PrvsFlag = False
    LstFlag = False
    TTransactionNo.Text = Val(TTransactionNo.Text) + 1
    Tbno = getNewTransactionNo
    If Val(TTransactionNo.Text) >= Tbno Then
        TTransactionNo.Text = TTransactionNo.Text - 1
    ElseIf Trim(TTransactionNo.Text) = "" Then
        TTransactionNo.Text = 1
    End If
    TransactionDetails
End Sub
Private Sub CPrevious_Click()
    FstFlag = False
    NxtFlag = False
    PrvsFlag = True
    LstFlag = False
    TTransactionNo.Text = Val(TTransactionNo.Text) - 1
    If Val(TTransactionNo.Text) = 0 And Trim(TTransactionNo.Text) = "" Then
        TTransactionNo.Text = 1
    End If
    TransactionDetails
End Sub
Private Sub CLast_Click()
    FstFlag = False
    NxtFlag = False
    PrvsFlag = False
    LstFlag = True
    TTransactionNo.Text = getNewTransactionNo
    If Val(TTransactionNo.Text) <> 1 Then
        TTransactionNo.Text = Val(TTransactionNo.Text) - 1
    End If
    TransactionDetails
End Sub
Private Sub TransactionDetails()
Dim rs As Recordset, r As Long, slNo As Long
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.PartNo,ItemMaster.BillingName,ItemMaster.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'SW' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode) Union Select Transaction.TempItemName As ItemName,'' As PartNo,Transaction.TempItemName As BillingName,0 As Code,'' As UnitName,Transaction.*,'' As ShortName From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'SW' ) Order By Transaction.SerialNo Asc,Code Desc")
    MGrid.Rows = 0
    
    If FstFlag = True Then
        If rs.EOF = True Then
            MsgBox "No Bills Are Recorded !", vbInformation
            Exit Sub
        End If
    ElseIf NxtFlag = True Then
        If rs.EOF = True Then
            MsgBox "This is Last Bill", vbInformation
        End If
    ElseIf PrvsFlag = True Then
        If rs.EOF = True Then
            MsgBox "This is First Bill", vbInformation
        End If
    ElseIf LstFlag = True Then
        If rs.EOF = True Then
            MsgBox "No Bills Are Recorded !", vbInformation
            Exit Sub
        End If
    End If
    
    If rs.RecordCount > 0 Then
    
        DTPDate.Value = rs!TransactionDate
        CoCustomer.Text = "" & rs!CustomerName
        TAddress.Text = "" & rs!CustomerAddress
        TNarration.Text = "" & rs!Narration
        TPayment.Text = Format("" & rs!SalePayment, "0.00")
      
        r = 0
        rs.MoveFirst
        While rs.EOF = False
        If slNo = Val("" & rs!SerialNo) Then
        
        Else
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gPartNo) = "" & rs!PartNo
            MGrid.TextMatrix(r, gQuantity) = "" & rs!Quantity
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gSaleRate) = Format("" & rs!SaleRate, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format(Val("" & rs!Quantity) * Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gMFRShortName) = "" & rs!ShortName
            slNo = Val("" & rs!SerialNo)
            r = r + 1
        End If
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
    
    LSlNo.Caption = MGrid.Rows + 1
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    LBalance.Caption = Format(Trim(LGrandAmount) - Val(TPayment.Text), "0.00")
End Sub
Private Sub getTransactionDetails()
Dim rs As Recordset, r As Long, slNo As Long
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.PartNo,ItemMaster.BillingName,ItemMaster.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'SW' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode) Union Select Transaction.TempItemName As ItemName,'' As PartNo,Transaction.TempItemName As BillingName,0 As Code,'' As UnitName,Transaction.*,'' As ShortName From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'SW' ) Order By Transaction.SerialNo Asc,Code Desc")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
    
        DTPDate.Value = rs!TransactionDate
        CoCustomer.Text = "" & rs!CustomerName
        TAddress.Text = "" & rs!CustomerAddress
        TNarration.Text = "" & rs!Narration
        TPayment.Text = Format("" & rs!SalePayment, "0.00")
      
        r = 0
        rs.MoveFirst
        While rs.EOF = False
        If slNo = Val("" & rs!SerialNo) Then
        
        Else
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gPartNo) = "" & rs!PartNo
            MGrid.TextMatrix(r, gQuantity) = "" & rs!Quantity
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gSaleRate) = Format("" & rs!SaleRate, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format(Val("" & rs!Quantity) * Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gMFRShortName) = "" & rs!ShortName
            slNo = Val("" & rs!SerialNo)
            r = r + 1
        End If
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
    
    LSlNo.Caption = MGrid.Rows + 1
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    LBalance.Caption = Format(Trim(LGrandAmount) - Val(TPayment.Text), "0.00")
End Sub
Private Sub printSale()
On Error GoTo GoOut
Open "LPT1:" For Output As #1
Dim i As Integer, Tamt As Double
    
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1,
    Print #1, Chr(27) & "!" & Chr(20) & Space(20) & Chr(0) & Chr(27) & "!" & Chr(50) & Trim(CoTitle.Text) & Chr(27) & "!" & Chr(0)
    Print #1,
    Print #1, Chr(27) & "!" & Chr(20) & "Customer : " & Left(Trim(CoCustomer.Text) & Space(22), 22) & Space(10) & " BNo: " & Left(Trim(TTransactionNo.Text) & Space(22), 22)
    Print #1, Chr(27) & "!" & Chr(20) & Space(44) & "Date: " & Left(Format(DTPDate.Value, "dd-MMM-yyyy") & Space(12), 12) & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "---------------------------------------------------------------------" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "| SNo |  Item Name        |Part.No |   Qty   |   Rate   |  Amount   |" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "|-----|-------------------|--------|---------|----------|-----------|" & Chr(0)
    
    Tamt = 0
    i = 0
    While i < MGrid.Rows
        Print #1, Chr(27) & "!" & Chr(4) & "|" & Left(MGrid.TextMatrix(i, gSerialNo) & Space(5), 5) & "|" & Left(" " & MGrid.TextMatrix(i, gItem) & Space(19), 19) & "|" & Left(MGrid.TextMatrix(i, gPartNo) & Space(8), 8) & "|" & Right(Space(9) & MGrid.TextMatrix(i, gQuantity) & " ", 9) & "|" & Right(Space(10) & Format("0" & MGrid.TextMatrix(i, gSaleRate), "0.00"), 10) & "|" & Right(Space(11) & Format("0" & MGrid.TextMatrix(i, gTotalAmount), "0.00"), 11) & "|" & Chr(0)
        Tamt = Tamt + Val(MGrid.TextMatrix(i, gTotalAmount))
        i = i + 1
    Wend
    
    Print #1, Chr(27) & "!" & Chr(4) & "|-----|-------------------|--------|---------|----------|-----------|" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "|  " & Left(NumberToWords(Tamt) & Space(54), 54) & Chr(0) & Chr(27) & "!" & Chr(29) & Right(Space(12) & Format("" & Val(Tamt), "0.00"), 12) & Chr(0) & "|"
    Print #1, Chr(27) & "!" & Chr(4) & "---------------------------------------------------------------------" & Chr(0)
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""

    Close #1
    MsgBox "Successfully Printed !", vbInformation
GoOut:
End Sub
Private Function getCentralAlignmentStartingPos(lPrintWidth As Long, sWord As String) As Long
Dim dPos As Long, lWordLen As Long
    lWordLen = Len(sWord)
    dPos = (lPrintWidth / 2) - (lWordLen / 2)
    getCentralAlignmentStartingPos = dPos
End Function

