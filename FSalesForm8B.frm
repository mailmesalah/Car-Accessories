VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FSalesForm8B 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales (Form 8B)"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FSalesForm8B.frx":0000
   ScaleHeight     =   8205
   ScaleWidth      =   14655
   StartUpPosition =   1  'CenterOwner
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
      Left            =   10890
      TabIndex        =   63
      Top             =   6480
      Width           =   2850
   End
   Begin VB.CommandButton CFirst 
      Height          =   435
      Left            =   8475
      Picture         =   "FSalesForm8B.frx":21F396
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   10965
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton CNext 
      Height          =   435
      Left            =   9855
      Picture         =   "FSalesForm8B.frx":221390
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   10965
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton CLast 
      Height          =   435
      Left            =   12600
      Picture         =   "FSalesForm8B.frx":22338A
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   10965
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton CPrevious 
      Height          =   435
      Left            =   11220
      Picture         =   "FSalesForm8B.frx":225384
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   10965
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton CDelete 
      Height          =   450
      Left            =   4410
      Picture         =   "FSalesForm8B.frx":22737E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   165
      Width           =   1485
   End
   Begin VB.CommandButton CClear 
      Height          =   450
      Left            =   3300
      Picture         =   "FSalesForm8B.frx":22C97C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6870
      Width           =   1485
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   450
      Left            =   1845
      Picture         =   "FSalesForm8B.frx":231F7A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6870
      Width           =   1485
   End
   Begin VB.CommandButton CAddItem 
      Height          =   450
      Left            =   405
      Picture         =   "FSalesForm8B.frx":237578
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6870
      Width           =   1485
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   450
      Left            =   12060
      Picture         =   "FSalesForm8B.frx":23CB76
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7470
      Width           =   1485
   End
   Begin VB.CommandButton CSave 
      Height          =   450
      Left            =   10605
      Picture         =   "FSalesForm8B.frx":242174
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7470
      Width           =   1485
   End
   Begin VB.CommandButton CPrint 
      Height          =   450
      Left            =   2745
      Picture         =   "FSalesForm8B.frx":247772
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7470
      Width           =   1485
   End
   Begin VB.CommandButton CNew 
      Height          =   450
      Left            =   1245
      Picture         =   "FSalesForm8B.frx":24CD70
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7470
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   2805
      Left            =   120
      TabIndex        =   15
      Top             =   2070
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   4948
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
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   325
      Left            =   2760
      TabIndex        =   16
      Top             =   180
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20643843
      CurrentDate     =   40544
   End
   Begin MSFlexGridLib.MSFlexGrid MGridItemDetails 
      Height          =   795
      Left            =   5280
      TabIndex        =   67
      Top             =   6525
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
   Begin MSForms.Label Label24 
      Height          =   270
      Left            =   8010
      TabIndex        =   70
      Top             =   6240
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
   Begin MSForms.Label Label20 
      Height          =   270
      Left            =   6600
      TabIndex        =   69
      Top             =   6240
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
   Begin MSForms.Label Label9 
      Height          =   270
      Left            =   5040
      TabIndex        =   68
      Top             =   6240
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
   Begin MSForms.Label LBalance 
      Height          =   325
      Left            =   10890
      TabIndex        =   66
      Top             =   6990
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
   Begin MSForms.Label LBalancelb 
      Height          =   325
      Left            =   10050
      TabIndex        =   65
      Top             =   7005
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
      Height          =   325
      Left            =   9960
      TabIndex        =   64
      Top             =   6525
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
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   390
      TabIndex        =   62
      Top             =   6855
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   1830
      TabIndex        =   61
      Top             =   6855
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   3285
      TabIndex        =   60
      Top             =   6855
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   1230
      TabIndex        =   59
      Top             =   7455
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2730
      TabIndex        =   58
      Top             =   7455
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   10590
      TabIndex        =   57
      Top             =   7455
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   12045
      TabIndex        =   56
      Top             =   7455
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   4395
      TabIndex        =   55
      Top             =   150
      Width           =   1530
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   60
      X2              =   14580
      Y1              =   2010
      Y2              =   2025
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   3960
      Left            =   75
      Top             =   1590
      Width           =   14520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   90
      X2              =   14595
      Y1              =   4890
      Y2              =   4890
   End
   Begin MSForms.TextBox TNarration 
      Height          =   330
      Left            =   1140
      TabIndex        =   1
      Top             =   525
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
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   210
      TabIndex        =   49
      Top             =   510
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "1508;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   325
      Left            =   1140
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
   Begin MSForms.Label Label1 
      Height          =   325
      Left            =   345
      TabIndex        =   48
      Top             =   165
      Width           =   465
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "820;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress 
      Height          =   330
      Left            =   11130
      TabIndex        =   47
      Top             =   405
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5292;573"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoCustomer 
      Height          =   325
      Left            =   11130
      TabIndex        =   3
      Top             =   45
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;573"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TTinNo 
      Height          =   330
      Left            =   11130
      TabIndex        =   46
      Top             =   1170
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5292;573"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TPhone 
      Height          =   330
      Left            =   11130
      TabIndex        =   45
      Top             =   795
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5292;573"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Left            =   10260
      TabIndex        =   44
      Top             =   1200
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Tin No"
      Size            =   "1508;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   10260
      TabIndex        =   43
      Top             =   795
      Width           =   705
      VariousPropertyBits=   8388627
      Caption         =   "Cont No"
      Size            =   "1244;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   10260
      TabIndex        =   42
      Top             =   420
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Place"
      Size            =   "1508;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   325
      Left            =   10140
      TabIndex        =   41
      Top             =   120
      Width           =   945
      VariousPropertyBits=   8388627
      Caption         =   "Customer"
      Size            =   "1667;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoAccount 
      Height          =   330
      Left            =   1140
      TabIndex        =   2
      Top             =   870
      Width           =   1410
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2487;573"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label8 
      Height          =   330
      Left            =   210
      TabIndex        =   40
      Top             =   870
      Width           =   945
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "1667;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label10 
      Height          =   390
      Left            =   225
      TabIndex        =   39
      Top             =   1650
      Width           =   495
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Sl.No"
      Size            =   "873;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label12 
      Height          =   390
      Left            =   6105
      TabIndex        =   38
      Top             =   1650
      Width           =   855
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Tax %"
      Size            =   "1508;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label13 
      Height          =   375
      Index           =   0
      Left            =   9225
      TabIndex        =   37
      Top             =   1650
      Width           =   705
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Rate"
      Size            =   "1244;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label14 
      Height          =   390
      Left            =   7065
      TabIndex        =   36
      Top             =   1650
      Width           =   855
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Qty"
      Size            =   "1508;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label15 
      Height          =   375
      Left            =   1785
      TabIndex        =   35
      Top             =   1650
      Width           =   945
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Item"
      Size            =   "1667;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label16 
      Height          =   375
      Left            =   4305
      TabIndex        =   34
      Top             =   1650
      Width           =   825
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Part No"
      Size            =   "1455;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label17 
      Height          =   390
      Left            =   13305
      TabIndex        =   33
      Top             =   1650
      Width           =   855
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Amount"
      Size            =   "1508;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label18 
      Height          =   375
      Left            =   11745
      TabIndex        =   32
      Top             =   1650
      Width           =   705
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Tax Amt"
      Size            =   "1244;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label19 
      Height          =   390
      Left            =   10305
      TabIndex        =   31
      Top             =   1650
      Width           =   1215
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Gross Value"
      Size            =   "2143;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoItem 
      Height          =   420
      Left            =   870
      TabIndex        =   4
      Top             =   5010
      Width           =   2970
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5239;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TQuantity 
      Height          =   420
      Left            =   6825
      TabIndex        =   5
      Top             =   5010
      Width           =   1080
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "1905;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TTax 
      Height          =   420
      Left            =   5955
      TabIndex        =   30
      Top             =   5010
      Width           =   840
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "1482;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TRate 
      Height          =   420
      Left            =   8865
      TabIndex        =   6
      Top             =   5010
      Width           =   1440
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2540;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LSlNo 
      Height          =   420
      Left            =   225
      TabIndex        =   29
      Top             =   5010
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
   Begin MSForms.Label LTotalAmount 
      Height          =   390
      Left            =   12825
      TabIndex        =   28
      Top             =   5010
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
   Begin MSForms.Label LUnit 
      Height          =   330
      Left            =   8040
      TabIndex        =   27
      Top             =   5025
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
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   7785
      TabIndex        =   26
      Top             =   1650
      Width           =   810
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "1429;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LGrandAmount 
      Height          =   570
      Left            =   10785
      TabIndex        =   25
      Top             =   5760
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
   Begin MSForms.TextBox TPartNo 
      Height          =   420
      Left            =   3870
      TabIndex        =   24
      Top             =   5010
      Width           =   2040
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3598;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label21 
      Height          =   330
      Left            =   135
      TabIndex        =   23
      Top             =   6210
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
   Begin MSForms.Label LCurrentStock 
      Height          =   330
      Left            =   1620
      TabIndex        =   22
      Top             =   6210
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LManufacturer 
      Height          =   330
      Left            =   1590
      TabIndex        =   21
      Top             =   5880
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;573"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label22 
      Height          =   330
      Left            =   135
      TabIndex        =   20
      Top             =   5880
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
   Begin MSForms.Label LRack 
      Height          =   345
      Left            =   1620
      TabIndex        =   19
      Top             =   6540
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label23 
      Height          =   330
      Left            =   135
      TabIndex        =   18
      Top             =   6540
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
      Left            =   3270
      TabIndex        =   17
      Top             =   5880
      Width           =   1650
   End
   Begin MSForms.Label Label11 
      Height          =   405
      Left            =   90
      TabIndex        =   54
      Top             =   1605
      Width           =   14475
      BackColor       =   4210752
      Size            =   "25532;714"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FSalesForm8B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim sCustomerCode() As String, sCustomerAddress() As String, sCustomerPhone() As String, sCustomerTinNo() As String
Dim sItemCode() As String, sBillingName() As String, sPartNo() As String
Dim sAccountCode() As String
Dim FstFlag As Boolean, NxtFlag As Boolean, PrvsFlag As Boolean, LstFlag As Boolean
Dim gSerialNo As Single, gItem As Single, gPartNo As Single, gCRate As Single, gPurchaseRate As Single, gBRate As Single, gARate As Single, gIQuantity As Single, gQuantity As Single, gUnit As Single, gTax As Single, gSaleRate As Single, gGrossValue As Single, gTaxAmount As Single, gBillingName As Single, gItemCode As Single, gMFRShortName As Single, gTotalAmount As Single
Private Sub MGridInitialise()
  
    gSerialNo = 0
    gItem = 1
    gPartNo = 2
    gTax = 3
    gQuantity = 4
    gUnit = 5
    gSaleRate = 6
    gGrossValue = 7
    gTaxAmount = 8
    gTotalAmount = 9
    gBillingName = 10
    gItemCode = 11
    gMFRShortName = 12
   
    
    MGrid.Clear
    MGrid.Rows = 1
    MGrid.Cols = 1
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 13
    MGrid.Rows = 0
    
    MGrid.ColWidth(gSerialNo) = 800
    MGrid.ColWidth(gItem) = 2800
    MGrid.ColWidth(gPartNo) = 2000
    MGrid.ColWidth(gTax) = 1000
    MGrid.ColWidth(gQuantity) = 1200
    MGrid.ColWidth(gUnit) = 900
    MGrid.ColWidth(gSaleRate) = 1300
    MGrid.ColWidth(gGrossValue) = 1300
    MGrid.ColWidth(gTaxAmount) = 1300
    MGrid.ColWidth(gTotalAmount) = 1400
    MGrid.ColWidth(gBillingName) = 0
    MGrid.ColWidth(gItemCode) = 0
    MGrid.ColWidth(gMFRShortName) = 0
    MGrid.RowHeightMin = 350
    
End Sub
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
'    MsgBox dBatchQuantity(CoBatch.ListIndex + 1)
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
    
'    If Val(TRate.Text) < dBatchMRP(CoBatch.ListIndex + 1) Then
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
        MGrid.TextMatrix(MGrid.Rows - 1, gPartNo) = IIf(Trim(TPartNo.Text) = "", "-", Trim(TPartNo.Text))
        MGrid.TextMatrix(MGrid.Rows - 1, gTax) = Format(Val(TTax.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount) = Format(Val(MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue)) * Val(MGrid.TextMatrix(MGrid.Rows - 1, gTax)) / 100, "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTotalAmount) = Format(Val(MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue)) + Val(MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount)), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gMFRShortName) = LMFRShortName.Caption & ""
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(MGrid.Rows - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gPartNo) = IIf(Trim(TPartNo.Text) = "", "-", Trim(TPartNo.Text))
        MGrid.TextMatrix(MGrid.Rows - 1, gTax) = Format(Val(TTax.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount) = Format(Val(MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue)) * Val(MGrid.TextMatrix(MGrid.Rows - 1, gTax)) / 100, "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTotalAmount) = Format(Val(MGrid.TextMatrix(MGrid.Rows - 1, gGrossValue)) + Val(MGrid.TextMatrix(MGrid.Rows - 1, gTaxAmount)), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gMFRShortName) = LMFRShortName.Caption & ""
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

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean
    bFound = False
    If (MsgBox("Do you want to Delete the Bill ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'FB' )")
        While rs.EOF = False
            bFound = True
            rs.Delete
            rs.MoveNext
        Wend
        rs.Close
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
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.BillNo = '" & Trim(TTransactionNo.Text) & "' ) And (AccountRegister.SpecialAccount In( 'Form8B','FBBillVoucher' ))")
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
Private Sub CoCustomer_Change()
    If CoCustomer.ListIndex <> -1 Then
        TAddress.Text = sCustomerAddress(CoCustomer.ListIndex + 1)
        TPhone.Text = sCustomerPhone(CoCustomer.ListIndex + 1)
        TTinNo.Text = sCustomerTinNo(CoCustomer.ListIndex + 1)
    Else
        TAddress.Text = ""
        TPhone.Text = ""
        TTinNo.Text = ""
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
    If KeyCode = 13 And CoItem.ListIndex = -1 Then
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
Private Sub getPartNoDetails()
Dim rs As Recordset, r As Long
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName From ItemMaster Where(ItemMaster.PartNo = '" & Trim(CoItem.Text) & "')")
    If rs.RecordCount > 0 Then
        TPartNo.Text = Trim(CoItem.Text)
        CoItem.Text = "" & rs!ItemName
    End If
    rs.Close
End Sub
Private Sub CPrint_Click()
    printSales
End Sub
Private Sub printSales()

On Error GoTo GoOut
    Dim i, x, y As Double
    Dim TaxAmount, TGrossvalue, Taxamt, TNetamt, TGrossvalue1, Taxamt1, TNetamt1 As Double

    TaxAmount = 0
    TGrossvalue = 0
    Taxamt = 0
    TNetamt = 0
    TGrossvalue1 = 0
    Taxamt1 = 0
    TNetamt1 = 0

    i = 0
    x = 500
    
    y = NewPage + 400
    
    While (i < MGrid.Rows)
    
        Printer.FontSize = 10
        Printer.FontBold = False
                
        x = 550
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Trim(MGrid.TextMatrix(i, gSerialNo))

        x = 1100
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Trim(MGrid.TextMatrix(i, gBillingName))
        
        x = 3700
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Trim(MGrid.TextMatrix(i, gPartNo))

        x = 4900
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gTax), "0.00")
        
        x = 5950
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gSaleRate), "0.00")
        
        x = 6950
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Trim(MGrid.TextMatrix(i, gQuantity))
        
        x = 7750
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gGrossValue), "0.00")
        
        x = 9300
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gTaxAmount), "0.00")

        x = 10600
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gTotalAmount), "0.00")
        
        TaxAmount = TaxAmount + Val(MGrid.TextMatrix(i, gTaxAmount))
        
        
        If Val(MGrid.TextMatrix(i, gTax)) = 4 Then
            Check = True
            TGrossvalue = Tgross + Val(MGrid.TextMatrix(i, gGrossValue))
            Taxamt = Tex + Val(MGrid.TextMatrix(i, gTaxAmount))
            TNetamt = Tnet + Val(MGrid.TextMatrix(i, gTotalAmount))
        Else
            Check1 = True
            TGrossvalue1 = Tgross1 + Val(MGrid.TextMatrix(i, gGrossValue))
            Taxamt1 = Tex1 + Val(MGrid.TextMatrix(i, gTaxAmount))
            TNetamt1 = Tnet1 + Val(MGrid.TextMatrix(i, gTotalAmount))
        End If
        
        i = i + 1
        y = y + 300
        If (y > 13000) Then
            y = NewPage + 400
        End If
    Wend
    
    y = 11100
    Printer.FontBold = True
    Printer.CurrentX = 3000
    Printer.CurrentY = y
    Printer.Print "Grand Totalt"
    
    Printer.CurrentX = 10600
    Printer.CurrentY = y
    Printer.Print Format(LGrandAmount.Caption, "0.00")
    
      
    Printer.Print Tab(5); String(90, "-")
    Printer.Print Tab(10); "Tax"; Tab(20); "Gross Value"; Tab(35); "Tax Amt"; Tab(50); "Net Amount";
    Printer.Print Tab(5); String(90, "-")
    
    If Check = True Then
        Printer.Print Tab(10); "4.00"; Tab(20); Format(TGrossvalue, "0.00"); Tab(35); Format(Taxamt, "0.00"); Tab(50); Format(TNetamt, "0.00");
    End If
    If Check1 = True Then
        Printer.Print Tab(10); "12.50"; Tab(20); Format(TGrossvalue1, "0.00"); Tab(35); Format(Taxamt1, "0.00"); Tab(50); Format(TNetamt1, "0.00");
    End If
    Printer.Print Tab(5); String(90, "-")
    
    x = 9300
    y = 14600
    Printer.FontSize = 10
    Printer.FontUnderline = False
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "For K.M AUTO SPARES"
    
    x = 9900
    y = 15200
    Printer.FontSize = 10
    Printer.FontUnderline = False
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Manager"
    
    Printer.EndDoc
    
    x = MsgBox("Successfully Printed !", vbInformation)
    
GoOut:
End Sub

Private Function NewPage() As Long

    Dim i, j, x, y As Double
    Dim Declaration(10) As String
    
    Printer.ScaleMode = 1
    Printer.FontName = "Arial"
    Printer.FontBold = False
    y = 400
    x = 450
    
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "TIN NO : 32100522309"
  
    x = x + 9500
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Phone : 6451285"
    
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.FontSize = 14
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("K.M AUTO SPARES")) / 2)
    Printer.CurrentY = 400
    Printer.Print "K.M AUTO SPARES"
    Printer.FontUnderline = False
    Printer.FontBold = False
    x = 400
    y = 800

    Printer.FontSize = 10
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("EZHOOR ROAD TIRUR - 1")) / 2)
    Printer.CurrentY = 800
    Printer.Print "EZHOOR ROAD TIRUR - 1"
    
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
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "SNo"

    x = 100 + 1000
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Particulars"


    x = 100 + 3650
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "PartNo"

    x = 100 + 4850
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Tax % "

    x = 100 + 5900
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Rate"

    x = 100 + 6950
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Qty"

    x = 100 + 7700
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Gross Value"

    x = 100 + 9100
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Tax Amount"

    x = 100 + 10500
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Total Amount"
    
    'HORIZONTAL LINES
    Printer.Line (400, 3600)-(12000, 3600)
    Printer.Line (400, 4000)-(12000, 4000)
    Printer.Line (400, 11000)-(12000, 11000)
    
    'FIRST AND LAST VERTICAL LINE
    Printer.Line (400, 3600)-(400, 11000)
    Printer.Line (12000, 3600)-(12000, 11000)
    
    'INNER LINES
    Printer.Line (1000, 3600)-(1000, 11000)
    Printer.Line (3500, 3600)-(3500, 11000)
    Printer.Line (4800, 3600)-(4800, 11000)
    Printer.Line (5600, 3600)-(5600, 11000)
    Printer.Line (6800, 3600)-(6800, 11000)
    Printer.Line (7600, 3600)-(7600, 11000)
    Printer.Line (9100, 3600)-(9100, 11000)
    Printer.Line (10500, 3600)-(10500, 11000)
    
    Printer.FontSize = 10
    Printer.FontUnderline = True
    Printer.FontBold = True
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Declaration")) / 2)
    Printer.CurrentY = 13600
    Printer.Print "Declaration"
    Printer.FontBold = False
    Printer.FontUnderline = False

    Declaration(0) = "Certified  that all the particulars  shown in the above Tax Invoice are true and correct in all respects and the goods which the tax charged"
    Declaration(1) = "and collected are in accordance with the provitions as the KVAT ACT 2003 and the rules made there under. It is also certified that my/our"
    Declaration(2) = "under KVAT 2003 is not subject to any suspension/cancellation and it is valid as on the date of this Bill."
    
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
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'FB' )")
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
        rs!TransactionType = "FB"
        rs!TransactionDate = DTPDate.Value
        rs!TransactionTime = Format(Time, "HH:MM AMPM")
        rs!Narration = Trim(TNarration.Text)
        rs!SupplierCode = ""
        rs!SupplierName = ""
        rs!SupplierAddress = ""
        rs!CustomerCode = sCustomerCode(CoCustomer.ListIndex + 1)
        rs!CustomerName = Trim(CoCustomer.Text)
        rs!CustomerAddress = Trim(TAddress.Text)
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!ItemCode = Trim(MGrid.TextMatrix(r, gItemCode))
        rs!Quantity = Val(MGrid.TextMatrix(r, gQuantity))
        rs!PurchaseRate = 0
        rs!SaleRate = Val(MGrid.TextMatrix(r, gSaleRate))
        rs!SalePayment = IIf(Val(TPayment.Text) = 0, "0", Val(TPayment.Text))
        rs!MRP = 0
        rs!Tax = Val(MGrid.TextMatrix(r, gTax))
        rs!ReferenceNo = ""
        rs!ReferenceDate = Date
        rs.Update
        r = r + 1
    Wend
    rs.Close
    
    MsgBox "Successfully Saved !", vbInformation
    addToAccountRegister
    
'    lYN = MsgBox("Do you want to take Print ?", vbDefaultButton2 Or vbYesNo)
'    If lYN = vbYes Then
'        printSale
'    Else
'
'    End If
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
    
    
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') ) And (AccountRegister.SpecialAccount = 'Form8B' )And (AccountRegister.Type In ('P','R') )And (AccountRegister.BillNo = '" & TTransactionNo.Text & "' )  Order By AccountRegister.Type ")
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
    rs!SpecialAccount = "Form8B"
    rs.Update

    
    rs.AddNew
    rs!TransactionNo = Val(sTransactionNo) + 1
    rs!Type = "R"
    rs!TransactionDate = DTPDate.Value
    rs!TransactionTime = Format(Time, "HH:MM AMPM")
    rs!AccountCode = IIf(CoCustomer.ListIndex = -1, Trim(generalCustomerAccountID), sAccountCode(CoCustomer.ListIndex + 1))
    rs!Narration = "- Advance"
    rs!CashOrCredit = "Cash"
    rs!Income = Val(TPayment.Text)
    rs!Expense = 0
    rs!BillNo = "" & TTransactionNo.Text
    rs!SpecialAccount = "Form8B"
    rs.Update
    rs.Close
End Sub
Private Sub Form_Load()
getCustomer
getItem
MGridInitialise
MGridItemDetailsInitialise
clearControls
TTransactionNo.Text = getNewTransactionNo
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
Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String
    
    Set rs = db.OpenRecordset("Select Max(Val( Transaction.TransactionNo)) As TNo From Transaction Where ( Transaction.TransactionType = 'FB' )")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function

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
Private Sub getCustomer()
Dim rs As Recordset
    
    CoCustomer.Clear
    
    Set rs = db.OpenRecordset("Select CustomerMaster.CustomerCode,CustomerMaster.AccountCode,CustomerMaster.CustomerName,CustomerMaster.Address1,CustomerMaster.Address2,CustomerMaster.Address3,CustomerMaster.Phone,CustomerMaster.TinNo From CustomerMaster Where (CustomerMaster.Status = True) Order By CustomerMaster.CustomerName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If

    ReDim sCustomerCode(rs.RecordCount) As String
    ReDim sCustomerAddress(rs.RecordCount) As String
    ReDim sCustomerPhone(rs.RecordCount) As String
    ReDim sCustomerTinNo(rs.RecordCount) As String
    ReDim sAccountCode(rs.RecordCount) As String
    While rs.EOF = False
        CoCustomer.AddItem "" & rs!CustomerName
        sCustomerCode(CoCustomer.ListCount) = "" & rs!CustomerCode
        sCustomerAddress(CoCustomer.ListCount) = "" & rs!Address1 & " " & rs!Address2 & " " & rs!Address3
        sCustomerPhone(CoCustomer.ListCount) = "" & rs!Phone
        sCustomerTinNo(CoCustomer.ListCount) = "" & rs!TinNo
        sAccountCode(CoCustomer.ListCount) = "" & rs!AccountCode
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getItem()
Dim rs As Recordset
    
    CoItem.Clear
    
    Set rs = db.OpenRecordset("Select ItemMaster.Code,ItemMaster.ItemName,ItemMaster.PartNo,ItemMaster.BillingName From ItemMaster Where (ItemMaster.Type = 'BItem' ) Order By ItemMaster.ItemName")
    
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
        sPartNo(rs.RecordCount) = "" & rs!PartNo
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

    Set rs = db.OpenRecordset("Select Manufacturer.ShortName,Manufacturer.ManufacturerName,Units.UnitName,ItemMaster.ItemName,ItemMaster.PartNo,ItemMaster.Tax,ItemMaster.Rack,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('O','P','SR','SA') ) And (Transaction.ItemCode = ItemMaster.Code )) As InStock,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('S','PR','F8','FB') ) And (Transaction.ItemCode = ItemMaster.Code )) As OutStock From ItemMaster,Units,Manufacturer Where (ItemMaster.Code = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code = ItemMaster.ManufacturerCode )")
    If rs.RecordCount > 0 Then
        LManufacturer.Caption = "" & rs!ManufacturerName
        LUnit.Caption = "" & rs!UnitName
        LCurrentStock.Caption = Val("" & rs!InStock) - Val("" & rs!OutStock)
        LMFRShortName.Caption = "" & rs!ShortName
        LRack.Caption = "" & rs!Rack
        TPartNo.Text = "" & rs!PartNo
        TTax.Text = "" & rs!Tax
    Else
        LManufacturer.Caption = ""
        LUnit.Caption = ""
        LCurrentStock.Caption = ""
        LMFRShortName.Caption = ""
        TPartNo.Text = ""
        LRack.Caption = ""
        TTax.Text = ""
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
    
    Set rs = db.OpenRecordset("Select (Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('FB') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.TransactionNo = '" & Trim(TTransactionNo.Text) & "')) As Quantity From Transaction ")
    If rs.RecordCount > 0 Then
        LCurrentStock.Caption = Val(LCurrentStock.Caption) + Val("" & rs!Quantity)
    End If
    
    Set rs = db.OpenRecordset("Select Transaction.MRP,Transaction.PurchaseRate,(Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('O','P','SR','SA') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.MRP = Transaction.MRP) And (T.PurchaseRate = Transaction.PurchaseRate) ) As InStock,(Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('S','PR','F8','FB') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.MRP = Transaction.MRP)And (T.PurchaseRate = Transaction.PurchaseRate)) As OutStock From Transaction Where (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) Group By Transaction.MRP,Transaction.PurchaseRate")
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
Private Sub getTotal()
    LTotalAmount.Caption = Val(TRate.Text) * Val(TQuantity.Text)
End Sub
Private Sub clearControls()
    
    'TTransactionNo.Text = getNewTransactionNo
    DTPDate.Value = Date
    TNarration.Text = ""
    CoCustomer.Text = ""
    TAddress.Text = ""
    TPartNo.Text = ""
    TPhone.Text = ""
    TTinNo.Text = ""
    MGrid.Rows = 0
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TPartNo.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    TRate.Text = ""
    TTax.Text = ""
    TPayment.Text = ""
    LTotalAmount.Caption = ""
    LBalance.Caption = ""
    
    FstFlag = False
    NxtFlag = False
    PrvsFlag = False
    LstFlag = False
    
    CoAccount.Clear
    CoAccount.AddItem "Debit"
    CoAccount.AddItem "Credit"
    

    
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
End Sub

Private Sub clearEditControls()
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    TPartNo.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    TRate.Text = ""
    TTax = ""
    LTotalAmount.Caption = ""
    LBalance.Caption = ""
    FstFlag = False
    NxtFlag = False
    PrvsFlag = False
    LstFlag = False
    
End Sub


Private Sub MGrid_Click()
Dim r As Long, i As Long

    If MGrid.Rows > 0 Then
        r = MGrid.Row
        LSlNo.Caption = Val(MGrid.TextMatrix(r, gSerialNo))
        CoItem.Text = Trim(MGrid.TextMatrix(r, gItem))
        TPartNo.Text = Trim(MGrid.TextMatrix(r, gPartNo))
        TTax.Text = Trim(MGrid.TextMatrix(r, gTax))
        TQuantity.Text = Val(MGrid.TextMatrix(r, gQuantity))
        LUnit.Caption = Trim(MGrid.TextMatrix(r, gUnit))
        TRate.Text = Val(MGrid.TextMatrix(r, gSaleRate))
        LTotalAmount.Caption = Val(MGrid.TextMatrix(r, gTotalAmount))
    End If
End Sub


Private Sub TQuantity_Change()
    getTotal
End Sub

Private Sub TRate_Change()
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
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.PartNo,ItemMaster.Tax,ItemMaster.BillingName,ItemMaster.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'FB' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode) Order By Transaction.SerialNo")
    
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
    
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!TransactionDate
        CoCustomer.Text = "" & rs!CustomerName
        TAddress.Text = "" & rs!CustomerAddress
        TNarration.Text = "" & rs!Narration
        TPayment.Text = Format("" & rs!SalePayment, "0.00")
        CoAccount.Text = IIf(rs!WholeSaleType = Null, "Debit", "" & rs!WholeSaleType)
        
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gPartNo) = "" & rs!PartNo
            MGrid.TextMatrix(r, gTax) = "" & rs!Tax
            MGrid.TextMatrix(r, gQuantity) = "" & rs!Quantity
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gSaleRate) = Format("" & rs!SaleRate, "0.00")
            MGrid.TextMatrix(r, gGrossValue) = Format(Val("" & rs!Quantity) * Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gTaxAmount) = Format(Val(MGrid.TextMatrix(r, gGrossValue)) * Val(MGrid.TextMatrix(r, gTax)) / 100, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format(Val(MGrid.TextMatrix(r, gTaxAmount)) + Val(MGrid.TextMatrix(r, gGrossValue)), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gMFRShortName) = "" & rs!ShortName
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
Private Sub getTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.PartNo,ItemMaster.BillingName,ItemMaster.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'FB' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode) Order By Transaction.SerialNo")
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
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gSerialNo) = "" & rs!SerialNo
            MGrid.TextMatrix(r, gItem) = "" & rs!ItemName
            MGrid.TextMatrix(r, gPartNo) = "" & rs!PartNo
            MGrid.TextMatrix(r, gTax) = "" & rs!Tax
            MGrid.TextMatrix(r, gQuantity) = "" & rs!Quantity
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gSaleRate) = Format("" & rs!SaleRate, "0.00")
            MGrid.TextMatrix(r, gGrossValue) = Format(Val("" & rs!Quantity) * Val("" & rs!SaleRate), "0.00")
            MGrid.TextMatrix(r, gTaxAmount) = Format(Val(MGrid.TextMatrix(r, gGrossValue)) * Val(MGrid.TextMatrix(r, gTax)) / 100, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format(Val(MGrid.TextMatrix(r, gTaxAmount)) + Val(MGrid.TextMatrix(r, gGrossValue)), "0.00")
            MGrid.TextMatrix(r, gBillingName) = "" & rs!BillingName
            MGrid.TextMatrix(r, gItemCode) = "" & rs!Code
            MGrid.TextMatrix(r, gMFRShortName) = "" & rs!ShortName
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

