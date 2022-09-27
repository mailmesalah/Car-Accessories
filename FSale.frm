VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FSRetailSales 
   Caption         =   "Local WholeSale"
   ClientHeight    =   11220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
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
   ScaleHeight     =   11220
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   285
      Width           =   1545
   End
   Begin VB.CommandButton CClear 
      Caption         =   "Clear"
      Height          =   435
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8115
      Width           =   1545
   End
   Begin VB.CommandButton CRemoveItem 
      Caption         =   "Remove"
      Height          =   435
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8115
      Width           =   1545
   End
   Begin VB.CommandButton CAddItem 
      Caption         =   "Add"
      Height          =   435
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8115
      Width           =   1545
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   570
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton CSave 
      Caption         =   "Save"
      Height          =   570
      Left            =   7365
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton CPrint 
      Caption         =   "Print"
      Height          =   570
      Left            =   2355
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton CNew 
      Caption         =   "New"
      Height          =   570
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9360
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   420
      Left            =   3375
      TabIndex        =   1
      Top             =   300
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   59179011
      CurrentDate     =   40544
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3285
      Left            =   285
      TabIndex        =   5
      Top             =   2340
      Width           =   11235
      _ExtentX        =   19817
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
   Begin MSForms.Label Label2 
      Height          =   285
      Left            =   5040
      TabIndex        =   41
      Top             =   900
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
      Left            =   5640
      TabIndex        =   40
      Top             =   840
      Width           =   1515
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2672;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   420
      Left            =   480
      TabIndex        =   39
      Top             =   7200
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
   Begin MSForms.Label LRack 
      Height          =   345
      Left            =   1950
      TabIndex        =   38
      Top             =   7215
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoItem 
      Height          =   420
      Left            =   1185
      TabIndex        =   6
      Top             =   5775
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
   Begin MSForms.TextBox TQuantity 
      Height          =   420
      Left            =   7305
      TabIndex        =   8
      Top             =   5775
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
   Begin MSForms.Label LTotalAmount 
      Height          =   390
      Left            =   10215
      TabIndex        =   37
      Top             =   5790
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
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   330
      TabIndex        =   36
      Top             =   2025
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
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1845
      TabIndex        =   35
      Top             =   2025
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
   Begin MSForms.Label Label16 
      Height          =   330
      Left            =   7125
      TabIndex        =   34
      Top             =   2025
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
   Begin MSForms.Label Label19 
      Height          =   330
      Left            =   10020
      TabIndex        =   33
      Top             =   2025
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
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   5475
      TabIndex        =   32
      Top             =   2025
      Width           =   2400
      VariousPropertyBits=   8388627
      Caption         =   "Batch"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   330
      Left            =   7995
      TabIndex        =   31
      Top             =   2025
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
   Begin MSForms.Label LUnit 
      Height          =   330
      Left            =   8325
      TabIndex        =   30
      Top             =   5820
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
   Begin MSForms.TextBox TRate 
      Height          =   420
      Left            =   9030
      TabIndex        =   9
      Top             =   5775
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
   Begin MSForms.Label Label10 
      Height          =   330
      Left            =   8760
      TabIndex        =   29
      Top             =   2025
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
   Begin MSForms.ComboBox CoBatch 
      Height          =   420
      Left            =   6105
      TabIndex        =   7
      Top             =   5775
      Width           =   1140
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2011;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LSlNo 
      Height          =   420
      Left            =   375
      TabIndex        =   28
      Top             =   5835
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
   Begin VB.Label LMFRShortName 
      BackStyle       =   0  'Transparent
      Height          =   405
      Left            =   3960
      TabIndex        =   27
      Top             =   6420
      Width           =   1650
   End
   Begin MSForms.Label LBatchStock 
      Height          =   345
      Left            =   3720
      TabIndex        =   26
      Top             =   6795
      Width           =   1455
      VariousPropertyBits=   8388627
      Size            =   "2566;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   510
      TabIndex        =   25
      Top             =   6420
      Width           =   1470
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
      Height          =   405
      Left            =   1995
      TabIndex        =   24
      Top             =   6420
      Width           =   1890
      VariousPropertyBits=   8388627
      Size            =   "3334;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TNarration 
      Height          =   420
      Left            =   1755
      TabIndex        =   2
      Top             =   810
      Width           =   3180
      VariousPropertyBits=   746604571
      Size            =   "5609;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   345
      TabIndex        =   23
      Top             =   885
      Width           =   1335
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
      Left            =   1980
      TabIndex        =   22
      Top             =   6795
      Width           =   1710
      VariousPropertyBits=   8388627
      Size            =   "3016;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   510
      TabIndex        =   21
      Top             =   6780
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
   Begin MSForms.Label LGrandAmount 
      Height          =   570
      Left            =   7680
      TabIndex        =   20
      Top             =   6450
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
   Begin MSForms.TextBox TAddress 
      Height          =   420
      Left            =   8310
      TabIndex        =   4
      Top             =   825
      Width           =   3210
      VariousPropertyBits=   746604571
      Size            =   "5662;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoCustomer 
      Height          =   420
      Left            =   8310
      TabIndex        =   3
      Top             =   315
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
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   7800
      TabIndex        =   19
      Top             =   360
      Width           =   375
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "661;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TTransactionNo 
      Height          =   420
      Left            =   1755
      TabIndex        =   0
      Top             =   300
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
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   330
      Width           =   465
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "820;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FSRetailSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dBatchMRP() As Double, dBatchQuantity() As Double
Dim sCustomerCode() As String, sCustomerAddress() As String
Dim sItemCode() As String, sBillingName() As String
Dim gSerialNo As Single, gItem As Single, gBatch As Single, gQuantity As Single, gUnit As Single, gSaleRate As Single, gMRP As Single, gTotalAmount As Single, gBillingName As Single, gItemCode As Single, gMFRShortName As Single

Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

    If CoItem.ListIndex = -1 Then
        MsgBox "Please Select a Item !", vbInformation
        CoItem.SetFocus
        Exit Sub
    End If
    
    'If CoBatch.ListIndex = -1 Then
    '    MsgBox "Please Select a Batch !", vbInformation
    '    CoBatch.SetFocus
    '    Exit Sub
    'End If
    
    If Val(TQuantity.Text) = 0 Then
        MsgBox "Please Enter Quantity greater than Zero !", vbInformation
        TQuantity.SetFocus
        Exit Sub
    End If
    
    If (CoBatch.ListIndex = -1) Then
        MsgBox "Please select a valid Batch !", vbInformation
    ElseIf Val(TQuantity.Text) > dBatchQuantity(CoBatch.ListIndex + 1) Then
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
    
    If Val(TRate.Text) < dBatchMRP(CoBatch.ListIndex + 1) Then
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
        MGrid.TextMatrix(MGrid.Rows - 1, gBatch) = Trim(CoBatch.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gTotalAmount) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gMFRShortName) = LMFRShortName.Caption & ""
    Else
        r = Val(LSlNo.Caption)
        MGrid.TextMatrix(r - 1, gItem) = Trim(CoItem.Text)
        MGrid.TextMatrix(r - 1, gBatch) = Trim(CoBatch.Text)
        MGrid.TextMatrix(r - 1, gQuantity) = Val(TQuantity.Text)
        MGrid.TextMatrix(r - 1, gUnit) = LUnit.Caption
        MGrid.TextMatrix(r - 1, gSaleRate) = Format(Val(TRate.Text), "0.00")
        MGrid.TextMatrix(r - 1, gTotalAmount) = Format(Val(TRate.Text) * Val(TQuantity.Text), "0.00")
        MGrid.TextMatrix(r - 1, gBillingName) = sBillingName(CoItem.ListIndex + 1)
        MGrid.TextMatrix(r - 1, gItemCode) = sItemCode(CoItem.ListIndex + 1)
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
    gBatch = 2
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
    MGrid.ColWidth(gBatch) = 1150
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
    
    Set rs = db.OpenRecordset("Select CustomerMaster.CustomerCode,CustomerMaster.CustomerName,CustomerMaster.Address1,CustomerMaster.Address2,CustomerMaster.Address3 From CustomerMaster Where (CustomerMaster.Status = True) Order By CustomerMaster.CustomerName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    ReDim sCustomerCode(rs.RecordCount) As String
    ReDim sCustomerAddress(rs.RecordCount) As String
    While rs.EOF = False
        CoCustomer.AddItem "" & rs!CustomerName
        sCustomerCode(CoCustomer.ListCount) = "" & rs!CustomerCode
        sCustomerAddress(CoCustomer.ListCount) = "" & rs!Address1 & " " & rs!Address2 & " " & rs!Address3
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

Private Sub getBatchDetailsOfItem()
Dim rs As Recordset
    CoBatch.Clear
    If CoItem.ListIndex = -1 Then
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select Transaction.Batch,(Select Sum(TN.Quantity)  From Transaction As TN Where (TN.ItemCode = Transaction.ItemCode ) And (TN.Batch = Transaction.Batch ) And (TN.TransactionType In ('O','P','SR','SA') ) ) As InStock,(Select Sum(TN.Quantity)  From Transaction As TN Where (TN.ItemCode = Transaction.ItemCode ) And (TN.Batch = Transaction.Batch ) And (TN.TransactionType In ('S','PR') ) ) As OutStock,(Select TN.MRP From Transaction As TN Where (TN.ItemCode = Transaction.ItemCode ) And (TN.Batch = Transaction.Batch ) And (TN.TransactionType In ('P','O') ) Group By TN.MRP) As MRP From Transaction Where (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "' ) Group By Transaction.Batch,Transaction.ItemCode")
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
        ReDim dBatchMRP(rs.RecordCount + 1) As Double
        ReDim dBatchQuantity(rs.RecordCount + 1) As Double
        While rs.EOF = False
            CoBatch.AddItem "" & rs!Batch
            dBatchMRP(CoBatch.ListCount) = Val("" & rs!MRP)
            dBatchQuantity(CoBatch.ListCount) = Val("" & rs!Instock) - Val("" & rs!Outstock)
            rs.MoveNext
        Wend
    End If
    rs.Close
    CoBatch.ListIndex = IIf(CoBatch.ListCount > 0, 0, -1)
End Sub

Private Sub getItemDetails()
Dim rs As Recordset, r As Long
    If (CoItem.ListIndex = -1) Then
        LManufacturer.Caption = ""
        LUnit.Caption = ""
        LCurrentStock.Caption = ""
        LMFRShortName.Caption = ""
        Exit Sub
    End If
    Set rs = db.OpenRecordset("Select Manufacturer.ShortName,Manufacturer.ManufacturerName,Units.UnitName,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('O','P','SR','SA') ) And (Transaction.ItemCode = ItemMaster.Code )) As InStock,(Select Sum(Transaction.Quantity) From Transaction Where (Transaction.TransactionType In ('S','PR') ) And (Transaction.ItemCode = ItemMaster.Code )) As OutStock From ItemMaster,Units,Manufacturer Where (ItemMaster.Code = '" & sItemCode(CoItem.ListIndex + 1) & "' ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code = ItemMaster.ManufacturerCode )")
    If rs.RecordCount > 0 Then
        LManufacturer.Caption = "" & rs!ManufacturerName
        LUnit.Caption = "" & rs!UnitName
        LCurrentStock.Caption = Val("" & rs!Instock) - Val("" & rs!Outstock)
        LMFRShortName.Caption = rs!ShortName & ""
    Else
        LManufacturer.Caption = ""
        LUnit.Caption = ""
        LCurrentStock.Caption = ""
        LMFRShortName.Caption = ""
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("Select Sum(Transaction.Quantity) As Quantity From Transaction Where (Transaction.TransactionNo='" & Trim(TTransactionNo.Text) & "') And (Transaction.TransactionType='S') And (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "')")
    If rs.RecordCount > 0 Then
        LCurrentStock.Caption = Val("" & LCurrentStock.Caption) + Val("" & rs!Quantity)
    End If
    rs.Close
    
    r = 0
    While r < MGrid.Rows
        If (Trim(MGrid.TextMatrix(r, gItemCode)) = sItemCode(CoItem.ListIndex + 1)) Then
            LCurrentStock.Caption = Val("" & LCurrentStock.Caption) - Val(MGrid.TextMatrix(r, gQuantity))
        End If
        r = r + 1
    Wend
End Sub

Private Sub clearControls()
    
    'TTransactionNo.Text = getNewTransactionNo
    DTPDate.Value = Date
    TNarration.Text = ""
    CoCustomer.Text = ""
    TAddress.Text = ""
    MGrid.Rows = 0
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    CoBatch.Text = ""
    TQuantity.Text = ""
    LUnit.Caption = ""
    TRate.Text = ""
    LTotalAmount.Caption = ""
    LGrandAmount.Caption = Format(getGrandTotal, "0.00")
End Sub

Private Sub clearEditControls()
    LSlNo.Caption = MGrid.Rows + 1
    CoItem.Text = ""
    CoBatch.Text = ""
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

Private Sub CoBatch_Change()
Dim r As Long, gridBatchStock As Long, rs As Recordset
    If CoBatch.ListIndex > -1 Then
        r = 0
        gridBatchStock = 0
        While r < MGrid.Rows
            If (Trim(MGrid.TextMatrix(r, gItemCode)) = sItemCode(CoItem.ListIndex + 1) And MGrid.TextMatrix(r, gBatch) = Trim(CoBatch.Text)) Then
                gridBatchStock = gridBatchStock + Val(MGrid.TextMatrix(r, gQuantity))
            End If
            r = r + 1
        Wend
        LBatchStock.Caption = 0
        
        Set rs = db.OpenRecordset("Select Sum(Transaction.Quantity) As Quantity From Transaction Where (Transaction.TransactionNo='" & Trim(TTransactionNo.Text) & "') And (Transaction.TransactionType='S') And (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "') And (Transaction.Batch = '" & Trim(CoBatch.Text) & "')")
        If rs.RecordCount > 0 Then
            LBatchStock.Caption = Val("" & LBatchStock.Caption) + Val("" & rs!Quantity)
        End If
        rs.Close
        
        LBatchStock.Caption = (Val("" & LBatchStock.Caption) + Val("" & dBatchQuantity(CoBatch.ListIndex + 1))) - gridBatchStock
        TRate.Text = Val("" & dBatchMRP(CoBatch.ListIndex + 1))
    Else
        LBatchStock.Caption = ""
        TRate.Text = ""
    End If
End Sub

Private Sub CoItem_Change()
    getItemDetails
    getBatchDetailsOfItem
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

Private Sub CPrint_Click()
    'printSale
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
        rs!CustomerCode = sCustomerCode(CoCustomer.ListIndex + 1)
        rs!CustomerName = Trim(CoCustomer.Text)
        rs!CustomerAddress = Trim(TAddress.Text)
        rs!SerialNo = Val(MGrid.TextMatrix(r, gSerialNo))
        rs!ItemCode = Trim(MGrid.TextMatrix(r, gItemCode))
        rs!Batch = Trim(MGrid.TextMatrix(r, gBatch))
        rs!Quantity = Val(MGrid.TextMatrix(r, gQuantity))
        rs!PurchaseRate = 0
        rs!SaleRate = Val(MGrid.TextMatrix(r, gSaleRate))
        rs!MRP = 0
        rs!ReferenceNo = ""
        rs!ReferenceDate = Date
        rs.Update
        r = r + 1
    Wend
    rs.Close
    
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
    clearControls
    TTransactionNo.Text = getNewTransactionNo
End Sub

Private Sub MGrid_Click()
Dim r As Long, i As Long

    If MGrid.Rows > 0 Then
        r = MGrid.Row
        LSlNo.Caption = Val(MGrid.TextMatrix(r, gSerialNo))
        CoItem.Text = Trim(MGrid.TextMatrix(r, gItem))
        CoBatch.Text = Trim(MGrid.TextMatrix(r, gBatch))
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
        LGrandAmount.Caption = Format(getGrandTotal, "0.00")
    End If
End Sub

Private Sub getTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName,ItemMaster.BillingName,ItemMaster.Code,Units.UnitName,Transaction.*,Manufacturer.ShortName From ItemMaster,Transaction,Units,Manufacturer Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'S' ) And (ItemMaster.Code = Transaction.ItemCode ) And (Units.Code = ItemMaster.UnitCode ) And (Manufacturer.Code=ItemMaster.ManufacturerCode) Order By Transaction.SerialNo")
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
            MGrid.TextMatrix(r, gBatch) = "" & rs!Batch
            MGrid.TextMatrix(r, gQuantity) = "" & rs!Quantity
            MGrid.TextMatrix(r, gUnit) = "" & rs!UnitName
            MGrid.TextMatrix(r, gSaleRate) = Format("" & rs!SaleRate, "0.00")
            MGrid.TextMatrix(r, gTotalAmount) = Format(Val("" & rs!Quantity) * Val("" & rs!SaleRate), "0.00")
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

Private Sub printSale()
DoEvents
    
    If (MGrid.Rows = 0) Then
        GoTo GoOut
    End If
    
    Dim i As Long, lines As Long
    
    'Checking if the data is already entered
    Dim rs As Recordset
    
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.TransactionNo = '" & Trim(TTransactionNo.Text) & "' ) And (Transaction.TransactionType = 'S' )")
    If (rs.RecordCount = 0) Then
        GoTo GoOut
        rs.Close
        db.Close
    End If
    rs.Close
    Set rs = db.OpenRecordset("Select CompanyMaster.* From CompanyMaster")
    Open "C:\SaleBill.txt" For Output As #1
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1, ""
    Print #1, ""
    Print #1, Chr(27) & "!" & Chr(20) & Space(getCentralAlignmentStartingPos(40, rs!Company)) & Chr(0) & Chr(27) & "!" & Chr(50) & "  " & rs!Company & Chr(27) & "!" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(20) & Space(getCentralAlignmentStartingPos(90, rs!Address1 & "," & rs!Address2 & "   " & rs!Phone)) & rs!Address1 & "," & rs!Address2 & "   " & rs!Phone & Chr(0)
    Print #1, Chr(27) & "!" & Chr(20) & "  DL No: " & rs!DLNo1 & Chr(0)
    Print #1, Chr(27) & "!" & Chr(20) & "  DL No: " & rs!DLNo2 & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
    rs.Close
    Print #1, Chr(27) & "!" & Chr(4) & "  Invoice No: " & Left(Trim(TTransactionNo.Text) & Space(27), 27) & Space(20) & "Date: " & Left(Format(DTPDate.Value, "dd/mm/yy") & Space(30), 30) & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "  Patient: " & Left(Trim(Trim(CoCustomer.Text)) & Space(30), 30) & Space(20) & "Dr  : " & Left(Trim(CoDoctor.Text) & Space(30), 30) & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & " No Item                                           MFR  Batch    ExpDate Qty   Total      " & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
        
    i = 0
    lines = 0
    While i < MGrid.Rows
        
        If (lines >= 10) Then
    
            Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
            Print #1, ""
            Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
            Print #1, Chr(27) & "!" & Chr(4) & " Wishing you a speedy recovery.                                                Pharmacist "
            Print #1,
            Print #1,
            Print #1,
            'PAGE BREAK SHOULD BE GIVEN HERE
              
            Print #1,
            Print #1,
            Set rs = db.OpenRecordset("Select CompanyMaster.* From CompanyMaster")
            Print #1, Chr(27) & "!" & Chr(20) & Space(getCentralAlignmentStartingPos(40, rs!Company)) & Chr(0) & Chr(27) & "!" & Chr(50) & "  " & rs!Company & Chr(27) & "!" & Chr(0)
            Print #1, Chr(27) & "!" & Chr(20) & Space(getCentralAlignmentStartingPos(90, rs!Address1 & "," & rs!Address2 & "   " & rs!Phone)) & rs!Address1 & "," & rs!Address2 & "   " & rs!Phone & Chr(0)
            Print #1, Chr(27) & "!" & Chr(20) & "  DL No: " & rs!DLNo1 & Chr(0)
            Print #1, Chr(27) & "!" & Chr(20) & "  DL No: " & rs!DLNo2 & Chr(0)
            Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
            rs.Close
            Print #1, Chr(27) & "!" & Chr(4) & "  Invoice No: " & Left(Trim(TTransactionNo.Text) & Space(27), 27) & Space(20) & "Date: " & Left(Format(DTPDate.Value, "dd/mm/yy") & Space(30), 30) & Chr(0)
            Print #1, Chr(27) & "!" & Chr(4) & "  Patient: " & Left(Trim(Trim(CoCustomer.Text)) & Space(30), 30) & Space(20) & "Dr  : " & Left(Trim(CoDoctor.Text) & Space(30), 30) & Chr(0)
            Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
            Print #1, Chr(27) & "!" & Chr(4) & " No Item                                           MFR  Batch    ExpDate Qty   Total      " & Chr(0)
            Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
            'Resetting line count
            lines = 0
        End If
        
        Print #1, Chr(27) & "!" & Chr(4) & " " & Right(Space(2) & MGrid.TextMatrix(i, gSerialNo), 2) & " " & Left(MGrid.TextMatrix(i, gBillingName) & Space(46), 46) & " " & Left(MGrid.TextMatrix(i, gMFRShortName) & Space(4), 4) & " " & Left(Trim(Right(Space(8) & MGrid.TextMatrix(i, gBatch), 8)) & Space(8), 8) & " " & Left(Format(MGrid.TextMatrix(i, gExpiry), "mm/yyyy") & Space(7), 7) & " " & Right(Space(3) & MGrid.TextMatrix(i, gQuantity), 3) & " " & Right(Space(9) & MGrid.TextMatrix(i, gTotalAmount), 9) & Chr(0)
        lines = lines + 1
        i = i + 1
    Wend
        
    While lines <= 10
        lines = lines + 1
        Print #1, ""
    Wend
    
    Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "                                                   Grand Total: " & Chr(0) & Chr(27) & "!" & Chr(44) & Right(Space(10) & Format(LGrandAmount.Caption, "0.00"), 10) & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & "------------------------------------------------------------------------------------------" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(4) & " Wishing you a speedy recovery.                                                Pharmacist "
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
    Print #1, ""
    Close #1
 
    Shell "C:\Print.bat C:\SaleBill.txt", vbHide
    MsgBox "Successfully Printed !", vbInformation
GoOut:
End Sub

Private Function getCentralAlignmentStartingPos(lPrintWidth As Long, sWord As String) As Long
Dim dPos As Long, lWordLen As Long
    lWordLen = Len(sWord)
    dPos = (lPrintWidth / 2) - (lWordLen / 2)
    getCentralAlignmentStartingPos = dPos
End Function
