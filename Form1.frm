VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRetailSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   11835
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CNew 
      Caption         =   "New"
      Height          =   570
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8565
      Width           =   2175
   End
   Begin VB.CommandButton CPrint 
      Caption         =   "Print"
      Height          =   570
      Left            =   3165
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8565
      Width           =   2175
   End
   Begin VB.CommandButton CSave 
      Caption         =   "Save"
      Height          =   570
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8565
      Width           =   2175
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   570
      Left            =   8670
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8565
      Width           =   2175
   End
   Begin VB.CommandButton CAddItem 
      Caption         =   "Add"
      Height          =   435
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7725
      Width           =   1545
   End
   Begin VB.CommandButton CRemoveItem 
      Caption         =   "Remove"
      Height          =   435
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7725
      Width           =   1545
   End
   Begin VB.CommandButton CClear 
      Caption         =   "Clear"
      Height          =   435
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7725
      Width           =   1545
   End
   Begin VB.CommandButton CDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   135
      Width           =   1545
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   420
      Left            =   2760
      TabIndex        =   16
      Top             =   195
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   50987011
      CurrentDate     =   40544
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3285
      Left            =   210
      TabIndex        =   17
      Top             =   1605
      Width           =   11355
      _ExtentX        =   20029
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
   End
   Begin MSFlexGridLib.MSFlexGrid MGridItemDetails 
      Height          =   795
      Left            =   255
      TabIndex        =   5
      Top             =   5745
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
   Begin MSForms.TextBox TPartNo 
      Height          =   420
      Left            =   6090
      TabIndex        =   44
      Top             =   4980
      Width           =   1125
      VariousPropertyBits=   746604571
      Size            =   "1984;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label15 
      Height          =   270
      Left            =   2985
      TabIndex        =   43
      Top             =   5460
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Purchase Rate"
      Size            =   "2593;476"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label12 
      Height          =   345
      Left            =   1425
      TabIndex        =   42
      Top             =   5460
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
      Left            =   555
      TabIndex        =   41
      Top             =   5460
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
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   345
      TabIndex        =   40
      Top             =   225
      Width           =   345
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "609;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   420
      Left            =   1140
      TabIndex        =   0
      Top             =   195
      Width           =   1590
      VariousPropertyBits=   746604571
      Size            =   "2805;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   7425
      TabIndex        =   39
      Top             =   255
      Width           =   375
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "661;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoCustomer 
      Height          =   420
      Left            =   7935
      TabIndex        =   3
      Top             =   210
      Width           =   3210
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "5662;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAddress 
      Height          =   420
      Left            =   7935
      TabIndex        =   38
      Top             =   720
      Width           =   3210
      VariousPropertyBits=   746604571
      Size            =   "5662;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LGrandAmount 
      Height          =   570
      Left            =   7875
      TabIndex        =   37
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
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   285
      TabIndex        =   36
      Top             =   7020
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Current Stock"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LCurrentStock 
      Height          =   405
      Left            =   1770
      TabIndex        =   35
      Top             =   7020
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   330
      TabIndex        =   34
      Top             =   780
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "1508;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TNarration 
      Height          =   420
      Left            =   1140
      TabIndex        =   1
      Top             =   705
      Width           =   3180
      VariousPropertyBits=   746604571
      Size            =   "5609;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LManufacturer 
      Height          =   405
      Left            =   1740
      TabIndex        =   33
      Top             =   6690
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   300
      Left            =   285
      TabIndex        =   32
      Top             =   6690
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Manufacturer"
      Size            =   "2593;529"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LSlNo 
      Height          =   420
      Left            =   345
      TabIndex        =   31
      Top             =   5040
      Width           =   555
      VariousPropertyBits=   8388627
      Caption         =   "SLNo"
      Size            =   "979;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label10 
      Height          =   330
      Left            =   8730
      TabIndex        =   30
      Top             =   1230
      Width           =   1560
      VariousPropertyBits=   8388627
      Caption         =   "Rate"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TRate 
      Height          =   420
      Left            =   9015
      TabIndex        =   7
      Top             =   4980
      Width           =   1125
      VariousPropertyBits=   746604571
      Size            =   "1984;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LUnit 
      Height          =   330
      Left            =   8295
      TabIndex        =   29
      Top             =   5025
      Width           =   600
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "1058;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   330
      Left            =   7965
      TabIndex        =   28
      Top             =   1230
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   5445
      TabIndex        =   27
      Top             =   1230
      Width           =   2400
      VariousPropertyBits=   8388627
      Caption         =   "Part No"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label19 
      Height          =   330
      Left            =   9990
      TabIndex        =   26
      Top             =   1230
      Width           =   1380
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
      Left            =   7095
      TabIndex        =   25
      Top             =   1230
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Qty"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1815
      TabIndex        =   24
      Top             =   1230
      Width           =   3480
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
      Index           =   0
      Left            =   300
      TabIndex        =   23
      Top             =   1230
      Width           =   555
      VariousPropertyBits=   8388627
      Caption         =   "Sl No"
      Size            =   "979;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalAmount 
      Height          =   390
      Left            =   10185
      TabIndex        =   22
      Top             =   4995
      Width           =   1140
      VariousPropertyBits=   8388627
      Caption         =   "Total Amount"
      Size            =   "2011;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TQuantity 
      Height          =   420
      Left            =   7275
      TabIndex        =   6
      Top             =   4980
      Width           =   960
      VariousPropertyBits=   746604571
      Size            =   "1693;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoItem 
      Height          =   420
      Left            =   1170
      TabIndex        =   4
      Top             =   4980
      Width           =   4860
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "8572;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LRack 
      Height          =   345
      Left            =   1770
      TabIndex        =   21
      Top             =   7350
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   420
      Left            =   285
      TabIndex        =   20
      Top             =   7350
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Rack"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Left            =   4425
      TabIndex        =   19
      Top             =   795
      Width           =   495
      VariousPropertyBits=   8388627
      Caption         =   "Title"
      Size            =   "873;503"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoTitle 
      Height          =   420
      Left            =   4905
      TabIndex        =   2
      Top             =   735
      Width           =   2025
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3572;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Height          =   405
      Left            =   3420
      TabIndex        =   18
      Top             =   6690
      Width           =   1650
   End
End
Attribute VB_Name = "FRetailSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dBatchMRP() As Double, dBatchQuantity() As Double, dBatchPurchaseRate() As Double
Dim sCustomerCode() As String, sCustomerAddress() As String, sAccountCode() As String
Dim sItemCode() As String, sBillingName() As String, sPartNo() As String
Dim gSerialNo As Single, gItem As Single, gPartNo As Single, gQuantity As Single, gPurchaseRate As Single, gIMRP As Single, gIQuantity As Single, gUnit As Single, gSaleRate As Single, gMRP As Single, gTotalAmount As Single, gBillingName As Single, gItemCode As Single, gMFRShortName As Single

Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

    If Trim(CoItem.Text) = 0 Then
        MsgBox "Please Select a Item !", vbInformation
        CoItem.SetFocus
        Exit Sub
    End If
    If Val(TQuantity.Text) = 0 Then
        MsgBox "Please Enter Quantity greater than Zero !", vbInformation
        TQuantity.SetFocus
        Exit Sub
    End If
    
'    If Val(TQuantity.Text) > IIf(MGridItemDetails.Rows > 0, MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIQuantity), MGridItemDetails.Rows = 0) Then
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
        
    If Val(LSlNo.Caption) > MGrid.Rows Then 'Add
        MGrid.AddItem ""
        MGrid.TextMatrix(MGrid.Rows - 1, gSerialNo) = Val(LSlNo.Caption)
        MGrid.TextMatrix(MGrid.Rows - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gPartNo) = IIf(Trim(TPartNo.Text) = "", "-", Trim(TPartNo.Text))
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTotalAmount) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gBillingName) = IIf(CoItem.ListIndex = -1, CoItem.Text, sBillingName(CoItem.ListIndex + 1))
        MGrid.TextMatrix(MGrid.Rows - 1, gItemCode) = IIf(CoItem.ListIndex = -1, "0", sItemCode(CoItem.ListIndex + 1))
        MGrid.TextMatrix(MGrid.Rows - 1, gMFRShortName) = LMFRShortName.Caption & ""
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(r - 1, gPartNo) = IIf(Trim(TPartNo.Text) = "", "-", Trim(TPartNo.Text))
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
    CoItem.SetFocus
End Sub

Private Sub CClear_Click()
    MGrid.Rows = 0
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
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
    gQuantity = 3
    gUnit = 4
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
On Error GoTo GoOut
Dim rs As Recordset, sTransactionNo As String

    Set rs = db.OpenRecordset("Select Max(Val( Transaction.TransactionNo)) As TNo From Transaction Where ( Transaction.TransactionType = 'S' )")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo

    Set rs = Localdb.OpenRecordset("Select Max(Val( Transaction.TransactionNo)) As TNo From Transaction Where ( Transaction.TransactionType = 'S' )")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
    
GoOut:
End Function

Private Sub getCustomer()
On Error GoTo GoOut
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
    

    
    Set rs = Localdb.OpenRecordset("Select CustomerMaster.CustomerCode,CustomerMaster.AccountCode,CustomerMaster.CustomerName,CustomerMaster.Address1,CustomerMaster.Address2,CustomerMaster.Address3 From CustomerMaster Where (CustomerMaster.Status = True) Order By CustomerMaster.CustomerName")
    CoCustomer.Clear
    
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
    
GoOut:
        

End Sub

Private Sub getItem()
On Error GoTo GoOut
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
    
  
    Set rs = Localdb.OpenRecordset("Select ItemMaster.Code,ItemMaster.PartNo,ItemMaster.ItemName,ItemMaster.BillingName From ItemMaster Where (ItemMaster.Type = 'BItem' ) Order By ItemMaster.ItemName")
    CoItem.Clear
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
    
GoOut:

End Sub
Private Sub getPartNoDetails()
Dim rs As Recordset, r As Long
On Error GoTo GoOut

    Set rs = db.OpenRecordset("Select ItemMaster.ItemName From ItemMaster Where(ItemMaster.PartNo = '" & Trim(CoItem.Text) & "')")
    If rs.RecordCount > 0 Then
        TPartNo.Text = Trim(CoItem.Text)
        CoItem.Text = "" & rs!ItemName
    End If
    rs.Close
    
    Set rs = Localdb.OpenRecordset("Select ItemMaster.ItemName From ItemMaster Where(ItemMaster.PartNo = '" & Trim(CoItem.Text) & "')")
    If rs.RecordCount > 0 Then
        TPartNo.Text = Trim(CoItem.Text)
        CoItem.Text = "" & rs!ItemName
    End If
    rs.Close
    
GoOut:
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
Private Sub getItemDetails()
Dim rs As Recordset, r As Long, dQuantity As Double
    MGridItemDetails.Rows = 0
    TQuantity.Text = ""
    TRate.Text = ""
    If (CoItem.ListIndex = -1) Then
        Exit Sub
    End If
    If CoItem.ListIndex > -1 Then
        TPartNo.Text = Trim("" & sPartNo(CoItem.ListIndex + 1))
    End If
    dQuantity = 0
    r = 0
    While r < MGrid.Rows
        If sItemCode(CoItem.ListIndex + 1) = MGrid.TextMatrix(r, gItemCode) Then
            dQuantity = dQuantity + MGrid.TextMatrix(r, gQuantity)
        End If
        r = r + 1
    Wend
    
    Set rs = db.OpenRecordset("Select Manufacturer.ShortName,Manufacturer.ManufacturerName,Units.UnitName,ItemMaster.ItemName,ItemMaster.Rack,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('O','P','SR','SA') ) And (Transaction.ItemCode = ItemMaster.Code )) As InStock,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('S','SW','PR') ) And (Transaction.ItemCode = ItemMaster.Code )) As OutStock From ItemMaster,Units,Manufacturer Where (ItemMaster.Code = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code = ItemMaster.ManufacturerCode )")
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
    
    Set rs = db.OpenRecordset("Select (Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('S') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.TransactionNo = '" & Trim(TTransactionNo.Text) & "')) As Quantity From Transaction ")
    If rs.RecordCount > 0 Then
        dQuantity = dQuantity - Val("" & rs!Quantity)
    End If
    
    Set rs = db.OpenRecordset("Select Transaction.MRP,Transaction.PurchaseRate,(Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('O','P','SR','SA') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.MRP = Transaction.MRP) And (T.PurchaseRate = Transaction.PurchaseRate) ) As InStock,(Select Sum(T.Quantity) From Transaction As T Where (T.TransactionType In ('S','SW','PR') ) And (T.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (T.MRP = Transaction.MRP)And (T.PurchaseRate = Transaction.PurchaseRate)) As OutStock From Transaction Where (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) Group By Transaction.MRP,Transaction.PurchaseRate")
    While rs.EOF = False
        If (Val("" & rs!InStock) - Val("" & rs!OutStock) - dQuantity) > 0 Then
            MGridItemDetails.AddItem Val("" & rs!InStock) - Val("" & rs!OutStock) - dQuantity & vbTab & rs!MRP & vbTab & rs!PurchaseRate
        Else
        
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub clearControls()
    
    TTransactionNo.Text = getNewTransactionNo
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
On Error GoTo GoOut
Dim rs As Recordset, lYN As Long, bFound As Boolean

    bFound = False
    If (MsgBox("Do you want to Delete the Bill ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = Localdb.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'S' )")
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
GoOut:
End Sub

Private Sub CNew_Click()
    clearControls
    TTransactionNo.Text = getNewTransactionNo
End Sub
Private Sub CoItem_Change()
    getItemDetails
'    getBatchDetailsOfItem
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
'        getBatchDetailsOfItem
    End If
    If KeyCode = 27 Then
        CSave.SetFocus
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

Private Sub CPrint_Click()
    printSale
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
On Error GoTo GoOut
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
    Set rs = Localdb.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'S' )")
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
        rs!CustomerCode = sCustomerCode(CoCustomer.ListIndex + 1)
        rs!CustomerName = Trim(CoCustomer.Text)
        rs!CustomerAddress = Trim(TAddress.Text)
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!ItemCode = Trim(MGrid.TextMatrix(r, gItemCode))
        rs!TempItemName = Trim(MGrid.TextMatrix(r, gItem))
        rs!Quantity = Val(MGrid.TextMatrix(r, gQuantity))
        rs!PurchaseRate = 0
        rs!SaleRate = Val(MGrid.TextMatrix(r, gSaleRate))
        rs!MRP = 0
        rs!ItemDiscount = 0
        rs!DiscountSaleRate = 0
        rs!ReferenceNo = ""
        rs!ReferenceDate = Date
        rs.Update
        r = r + 1
    Wend
    rs.Close
    
    'Add to Accounts Details
    AddToAccount
    
    MsgBox "Successfully Saved !", vbInformation
    lYN = MsgBox("Do you want to take Print ?", vbDefaultButton2 Or vbYesNo)
    If lYN = vbYes Then
        printSale
    Else
        
    End If
    
    clearControls
    TTransactionNo.Text = getNewTransactionNo
    TTransactionNo.SetFocus
GoOut:
End Sub
Private Sub AddToAccount()
On Error GoTo GoOut
Dim rs As Recordset

    'SAVES DATA TO ACCOUNTREGISTER TABLE
    Set rs = Localdb.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.BillNo = '" & "SF" & Trim(TTransactionNo.Text) & "' )")
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend

    'BILLAMOUNT
    rs.AddNew
    rs!TransactionNo = getNewAccountTNo("P")
    rs!TransactionType = "P"
    rs!AccountCode = IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sCustomerGroupCode)
    rs!TransactionDate = DTPDate.Value
    rs!TransactionTime = Format(Time, "HH:MM AMPM")
    rs!Expense = Val(LGrandAmount.Caption)
    rs!Income = 0
    rs!Narration = "LocalSales Bill Amount On Retail Bill No" & Trim(TTransactionNo.Text)
    rs!BillNo = "SL" & Trim(TTransactionNo.Text)
    rs!CashOrCredit = "CR"
    rs.Update

    'ADVANCE
    rs.AddNew
    rs!TransactionNo = getNewAccountTNo("R")
    rs!TransactionType = "R"
    rs!AccountCode = IIf(CoCustomer.ListIndex <> -1, sAccountCode(CoCustomer.ListIndex + 1), sCustomerGroupCode)
    rs!TransactionDate = DTPDate.Value
    rs!TransactionTime = Format(Time, "HH:MM AMPM")
    rs!Expense = 0
    rs!Income = Val(LGrandAmount.Caption)
    rs!Narration = "LocalSales Payment On Retail Bill No" & Trim(TTransactionNo.Text)
    rs!BillNo = "SL" & Trim(TTransactionNo.Text)
    rs!CashOrCredit = "CA"
    rs.Update
    rs.Close
GoOut:
End Sub
Private Function getNewAccountTNo(sMode As String) As String
On Error GoTo GoOut
Dim rs As Recordset, sTCode As String
    
    Set rs = Localdb.OpenRecordset("Select Max(val(AccountRegister.TransactionNo))As ACode From AccountRegister Where (AccountRegister.TransactionType = '" & sMode & "' )")
    If rs.RecordCount > 0 Then
        sTCode = Val("" & rs!ACode) + 1
    Else
        sTCode = "1"
    
    End If
    rs.Close
    
    getNewAccountTNo = sTCode
    
    Set rs = db.OpenRecordset("Select Max(val(AccountRegister.TransactionNo))As ACode From AccountRegister Where (AccountRegister.TransactionType = '" & sMode & "' )")
    If rs.RecordCount > 0 Then
        sTCode = Val("" & rs!ACode) + 1
    Else
        sTCode = "1"
    
    End If
    rs.Close
    
    getNewAccountTNo = sTCode
    
GoOut:

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

Private Sub MGridItemDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If (MGridItemDetails.Rows > 0) Then
            TQuantity.Text = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIQuantity)
            TRate.Text = MGridItemDetails.TextMatrix(MGridItemDetails.Row, gIMRP)
        End If
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TAddress_GotFocus()
    TAddress.SelStart = 0
    TAddress.SelLength = Len(TAddress.Text)
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
        getTransactionDetails
        LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    End If
End Sub

Private Sub getTransactionDetails()
'On Error GoTo GoOut
Dim rs As Recordset, r As Long, slNo As Long
    Set rs = Localdb.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.PartNo,ItemMaster.BillingName,ItemMaster.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & TTransactionNo.Text & "' ) And (Transaction.TransactionType = 'S' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode)  Union Select Transaction.TempItemName As ItemName,'' As PartNo,Transaction.TempItemName As BillingName,0 As Code,'' As UnitName,Transaction.*,'' As ShortName From Transaction Where (Transaction.TransactionNo = '" & TTransactionNo.Text & "' ) And (Transaction.TransactionType = 'S' ) Order By Transaction.SerialNo Asc,Code Desc")
    'Set rs = Localdb.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.PartNo,ItemMaster.BillingName,ItemMaster.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'S' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode) Order By Transaction.SerialNo")
    MGrid.Rows = 0
    slNo = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = rs!TransactionDate
        CoCustomer.Text = "" & rs!CustomerName
        TAddress.Text = "" & rs!CustomerAddress
        TNarration.Text = "" & rs!Narration
        
        r = 0
        rs.MoveFirst
        While rs.EOF = False
            If slNo = Val("" & rs!SerialNo) Then
            
            Else
                MGrid.AddItem ""
                MGrid.TextMatrix(r, gItem) = "" & rs!TempItemName
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
            End If
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
    
    LSlNo.Caption = MGrid.Rows + 1
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
GoOut:
End Sub

Private Sub printSale()
On Error GoTo GoOut
Open "LPT1:" For Output As #1
Dim i As Integer, Tamt As Double
    
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1,
    Print #1, Chr(27) & "!" & Chr(20) & Space(20) & Chr(0) & Chr(27) & "!" & Chr(50) & "ESTIMATE" & Chr(27) & "!" & Chr(0)
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

