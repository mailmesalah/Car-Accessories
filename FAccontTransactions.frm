VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FAccontTransactions 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accont Transactions"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   11895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CPrint 
      Caption         =   "Print"
      Height          =   570
      Left            =   1035
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8025
      Width           =   2175
   End
   Begin VB.CommandButton CSave 
      Caption         =   "Save"
      Height          =   570
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8025
      Width           =   2175
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   570
      Left            =   8850
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8025
      Width           =   2175
   End
   Begin VB.CommandButton CAddItem 
      Caption         =   "Add Item"
      Height          =   435
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6735
      Width           =   1545
   End
   Begin VB.CommandButton CRemoveItem 
      Caption         =   "Remove Item"
      Height          =   435
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6735
      Width           =   1545
   End
   Begin VB.CommandButton CClear 
      Caption         =   "Clear"
      Height          =   435
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6735
      Width           =   1545
   End
   Begin VB.CommandButton CDelete 
      Caption         =   "Delete"
      Height          =   435
      Left            =   2010
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   165
      Width           =   1545
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3660
      Left            =   315
      TabIndex        =   13
      Top             =   1170
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   6456
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
      Height          =   420
      Left            =   345
      TabIndex        =   14
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   53739523
      CurrentDate     =   40458
   End
   Begin MSForms.Label Label6 
      Height          =   420
      Left            =   8040
      TabIndex        =   30
      Top             =   7080
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Current Balance"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LCurrentBalance 
      Height          =   345
      Left            =   9720
      TabIndex        =   29
      Top             =   7080
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2725;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LVoucherNo 
      Height          =   390
      Left            =   315
      TabIndex        =   28
      Top             =   4995
      Width           =   1155
      VariousPropertyBits=   8388627
      Caption         =   "Voucher No"
      Size            =   "2037;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoAccounts 
      Height          =   420
      Left            =   3675
      TabIndex        =   2
      Top             =   4980
      Width           =   4035
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7117;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TReceipt 
      Height          =   420
      Left            =   7755
      TabIndex        =   3
      Top             =   4980
      Width           =   1710
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3016;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TPayment 
      Height          =   420
      Left            =   9510
      TabIndex        =   4
      Top             =   4980
      Width           =   1710
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3016;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   360
      TabIndex        =   27
      Top             =   840
      Width           =   1050
      VariousPropertyBits=   8388627
      Caption         =   "Voucher No"
      Size            =   "1852;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1545
      TabIndex        =   26
      Top             =   840
      Width           =   2400
      VariousPropertyBits=   8388627
      Caption         =   "Bill No"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label16 
      Height          =   330
      Left            =   8010
      TabIndex        =   25
      Top             =   840
      Width           =   1170
      VariousPropertyBits=   8388627
      Caption         =   "Receipt"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   9585
      TabIndex        =   24
      Top             =   840
      Width           =   1560
      VariousPropertyBits=   8388627
      Caption         =   "Payment"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalReceipt 
      Height          =   420
      Left            =   9705
      TabIndex        =   23
      Top             =   5775
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "Total Receipt0"
      Size            =   "2725;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   8040
      TabIndex        =   22
      Top             =   6150
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Total Payment"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalPayment 
      Height          =   345
      Left            =   9705
      TabIndex        =   21
      Top             =   6150
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "Total Payment0"
      Size            =   "2725;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label11 
      Height          =   420
      Left            =   8040
      TabIndex        =   20
      Top             =   5745
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Total Receipt"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBalance 
      Height          =   345
      Left            =   9705
      TabIndex        =   19
      Top             =   6600
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2725;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   420
      Left            =   8040
      TabIndex        =   18
      Top             =   6600
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TNarration 
      Height          =   420
      Left            =   3660
      TabIndex        =   5
      Top             =   5625
      Width           =   4065
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "7170;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   4590
      TabIndex        =   17
      Top             =   840
      Width           =   2400
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   1425
      TabIndex        =   16
      Top             =   5655
      Width           =   2400
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TBillNo 
      Height          =   420
      Left            =   1665
      TabIndex        =   1
      Top             =   4980
      Width           =   1965
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3466;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoSaleType 
      Height          =   420
      Left            =   8670
      TabIndex        =   0
      Top             =   180
      Width           =   2880
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5080;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   6975
      TabIndex        =   15
      Top             =   225
      Width           =   1335
      VariousPropertyBits=   8388627
      Caption         =   "Type"
      Size            =   "2355;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FAccontTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Database
Dim sAccountCode() As String
Dim gVoucherNo As Single, gAccount As Single, gNarration As Single, gReceipt As Single, gPayment As Single, gAccountCode As Single, gSpecialAccount As Single, gBillNo As Single
Dim lSelectedVoucherNo As Long, lFirstVoucherNo As Long

Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

    If (CoSaleType.ListIndex = -1) Then
        MsgBox "Please Select a Sale type !", vbInformation
        CoSaleType.SetFocus
        Exit Sub
    End If
    
    If ifNotBillNoExist(Trim(TBillNo.Text), Trim(CoSaleType.Text)) Then
        MsgBox "Please select a valid Bill No !", vbInformation
        TBillNo.SetFocus
        Exit Sub
    End If
    
    If CoAccounts.ListIndex = -1 Then
        MsgBox "Please Select an Account !", vbInformation
        CoAccounts.SetFocus
        Exit Sub
    End If
    
    If Val(TReceipt.Text) = 0 And Val(TPayment.Text) = 0 Then
        MsgBox "Please Enter Receipt or Payment !", vbInformation
        TReceipt.SetFocus
        Exit Sub
    End If
    
    If Val(TReceipt.Text) > 0 And Val(TPayment.Text) > 0 Then
        MsgBox "Please Enter Receipt or Payment only one at a time !", vbInformation
        TReceipt.SetFocus
        Exit Sub
    End If
        
    If Val(lSelectedVoucherNo) = 0 Then 'Add
        MGrid.AddItem ""
        MGrid.TextMatrix(MGrid.Rows - 1, gVoucherNo) = LVoucherNo.Caption
        MGrid.TextMatrix(MGrid.Rows - 1, gBillNo) = IIf(CoSaleType.ListIndex = 0, "R-", "O-") & Trim(TBillNo.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gAccount) = Trim(CoAccounts.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gNarration) = Trim(TNarration.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(Val(TReceipt.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(Val(TPayment.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gAccountCode) = sAccountCode(CoAccounts.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gSpecialAccount) = IIf(CoSaleType.ListIndex = 0, "LocalWholeSales", "Purchase")
    Else
        r = MGrid.Row
        MGrid.TextMatrix(r, gBillNo) = IIf(CoSaleType.ListIndex = 0, "R-", "O-") & Trim(TBillNo.Text)
        MGrid.TextMatrix(r, gAccount) = Trim(CoAccounts.Text)
        MGrid.TextMatrix(r, gNarration) = Trim(TNarration.Text)
        MGrid.TextMatrix(r, gReceipt) = Format(Val(TReceipt.Text), "0.00")
        MGrid.TextMatrix(r, gPayment) = Format(Val(TPayment.Text), "0.00")
        MGrid.TextMatrix(r, gAccountCode) = sAccountCode(CoAccounts.ListIndex + 1)
        MGrid.TextMatrix(r, gSpecialAccount) = IIf(CoSaleType.ListIndex = 0, "LocalWholeSales", "Purchase")
    End If
    lSelectedVoucherNo = 0
    LVoucherNo.Caption = getNewTransactionNo
    clearEditControls
    setBalance
    CoAccounts.SetFocus
End Sub
Private Function getBalance() As Double
Dim dBalance, TPayment, tBalance As Double
Dim rs As Recordset
    dBalance = 0
    TPayment = 0
    tBalance = 0
    Set rs = Localdb.OpenRecordset("Select Transaction.TransactionNo,Transaction.SaleRate, Transaction.Quantity ,Transaction.WholeSalePayment From Transaction Where (Transaction.TransactionNo = '" & Trim(TBillNo.Text) & "' ) And (Transaction.TransactionType = 'SW' ) ")
    While rs.EOF = False
        tBalance = Val("" & rs!SaleRate) * Val("" & rs!Quantity)
        TPayment = Val("" & rs!WholeSalePayment)
        dBalance = dBalance + tBalance
        rs.MoveNext
    Wend
    dBalance = dBalance - TPayment
    rs.Close
    getBalance = dBalance
End Function
Private Function ifNotBillNoExist(sBillNo As String, sType As String) As Boolean
Dim rs As Recordset, bNotExist As Boolean
    bNotExist = True
    If sType = "Local Sales(WholeSales)" Then
        Set rs = Localdb.OpenRecordset("Select Transaction.TransactionNo From Transaction Where (Transaction.TransactionNo = '" & sBillNo & "' ) And (Transaction.TransactionType = 'SW' ) ")
        If rs.RecordCount > 0 Then
            bNotExist = False
        End If
        rs.Close
    ElseIf sType = "Purchase" Then
        Set rs = Localdb.OpenRecordset("Select Transaction.TransactionNo From Transaction Where (Transaction.TransactionNo = '" & sBillNo & "' ) And (Transaction.TransactionType = 'P' ) ")
        If rs.RecordCount > 0 Then
            bNotExist = False
        End If
        rs.Close
    End If
    
    ifNotBillNoExist = bNotExist
End Function
Private Sub getBillNoAccountDetails(sBillNo As String, sType As String)
    Dim rs As Recordset
    If sType = "Local Sales(WholeSales)" Then
        Set rs = Localdb.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName From AccountMaster,Transaction,CustomerMaster Where (Transaction.TransactionNo = '" & sBillNo & "' ) And (Transaction.TransactionType ='SW') And (AccountMaster.Code = CustomerMaster.AccountCode And CustomerMaster.CustomerCode = Transaction.CustomerCode )")
        If rs.RecordCount > 0 Then
            CoAccounts.Text = "" & rs!AccountName
        End If
    ElseIf sType = "Purchase" Then
        Set rs = db.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName From AccountMaster,Transaction,SupplierMaster Where (Transaction.TransactionNo = '" & sBillNo & "' ) And (Transaction.TransactionType ='P') And (AccountMaster.Code = SupplierMaster.AccountCode And SupplierMaster.SupplierCode = Transaction.SupplierCode )")
        If rs.RecordCount > 0 Then
            CoAccounts.Text = "" & rs!AccountName
        End If
        rs.Close
    End If
End Sub
Private Sub CClear_Click()
    MGrid.Rows = 0
    setBalance
    lSelectedVoucherNo = 0
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gVoucherNo = 0
    gBillNo = 1
    gAccount = 2
    gNarration = 3
    gReceipt = 4
    gPayment = 5
    gAccountCode = 6
    gSpecialAccount = 7
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 8
    MGrid.Rows = 0
    MGrid.ColWidth(gVoucherNo) = 1200
    MGrid.ColWidth(gBillNo) = 2965
    MGrid.ColWidth(gAccount) = 3300
    MGrid.ColWidth(gNarration) = 0
    MGrid.ColWidth(gReceipt) = 1700
    MGrid.ColWidth(gPayment) = 1700
    MGrid.ColWidth(gAccountCode) = 0
    MGrid.ColWidth(gSpecialAccount) = 0
    MGrid.RowHeightMin = 350
End Sub

Private Function getNewTransactionNo() As String
Dim rs As Recordset, lTransactionNo As String, lRecordCount As Long, r As Long, lBigestNo As Long
     
    lBigestNo = 0
    
    Set rs = db.OpenRecordset("Select Max(Val( AccountRegister.TransactionNo)) As TNo From AccountRegister ")
    If rs.RecordCount > 0 Then
        lTransactionNo = Val("" & rs!TNo)
    Else
        lTransactionNo = 0
    End If
    rs.Close
    
    r = 0
    While r < MGrid.Rows
        If lBigestNo < Val(MGrid.TextMatrix(r, gVoucherNo)) Then
            lBigestNo = Val(MGrid.TextMatrix(r, gVoucherNo))
        End If
        r = r + 1
    Wend
    
    If lBigestNo < lTransactionNo Then
        lBigestNo = lTransactionNo
    End If
    
    getNewTransactionNo = lBigestNo + 1
End Function

Private Sub getAccounts()
Dim rs As Recordset
    
    CoAccounts.Clear
    
    Set rs = db.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName From AccountMaster Where (AccountMaster.GroupCode = '" & customerAccountParentID & "'  ) And (AccountMaster.Type='BAccount')")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sAccountCode(rs.RecordCount + 1) As String
    ReDim sItemBillingName(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoAccounts.AddItem "" & rs!AccountName
        sAccountCode(CoAccounts.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub clearControls()
    
    DTPDate.Value = Date
    MGrid.Rows = 0
    TBillNo.Text = ""
    CoAccounts.ListIndex = -1
    TNarration.Text = ""
    TReceipt.Text = ""
    TPayment.Text = ""
    LVoucherNo.Caption = getNewTransactionNo
    setBalance
End Sub

Private Sub clearEditControls()
    TBillNo.Text = ""
    CoAccounts.ListIndex = -1
    TNarration.Text = ""
    TReceipt.Text = ""
    TPayment.Text = ""
    LVoucherNo.Caption = getNewTransactionNo
End Sub

Private Sub CDelete_Click()
Dim rs As Recordset, lYN As Long, bFound As Boolean
    bFound = False
    If (MsgBox("Do you want to Delete this day's Transaction ?", vbDefaultButton2 Or vbYesNo) = vbYes) Then
        Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') And (AccountRegister.SpecialAccount='SaleAndPurchaseOrder' Or AccountRegister.SpecialAccount='SaleAndPurchaseReadyMade') )")
        While rs.EOF = False
            bFound = True
            rs.Delete
            rs.MoveNext
        Wend
        rs.Close
        
        If bFound Then
            MsgBox "Successfully Deleted !", vbInformation
            clearControls
            getTransactionDetails
            LVoucherNo.Caption = getNewTransactionNo
        Else
            MsgBox "Bill Not Found !", vbInformation
        End If
    End If
End Sub

Private Sub CoAccounts_GotFocus()
    CoAccounts.SelStart = 0
    CoAccounts.SelLength = Len(CoAccounts.Text)
End Sub

Private Sub CoAccounts_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim r As Long
    If KeyCode = 113 Then
        FAccountMaster.Show vbModal
        getAccounts
    End If
End Sub

Private Sub CPrint_Click()
    'printOpeningStock
End Sub

Private Sub CRemoveItem_Click()
Dim r As Long
    If MGrid.Rows <= 0 Then
        Exit Sub
    End If
    
    If MGrid.TextMatrix(MGrid.Row, gSpecialAccount) = "LocalWholeSales" Or MGrid.TextMatrix(MGrid.Row, gSpecialAccount) = "Purchase" Then
        If MGrid.Rows = 1 Then
            MGrid.Rows = 0
            clearEditControls
        Else
            MGrid.RemoveItem (MGrid.Row)
            clearEditControls
        End If
        setBalance
    Else
    
    End If
    lSelectedVoucherNo = 0
End Sub

Private Sub CSave_Click()
Dim rs As Recordset
Dim r As Long, lYN As Long, sStatus As String

    If MGrid.Rows = 0 Then
        MsgBox "No Items Entered !", vbInformation
        CoAccounts.SetFocus
        Exit Sub
    End If
    
    
    'SAVES DATA TO AccountRegister TABLE
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') And (AccountRegister.SpecialAccount='LocalWholeSales' Or AccountRegister.SpecialAccount='Purchase') )")
        
    'SAVES DATA TO TransactionRegister ReadyMade
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    r = 0
    While r < MGrid.Rows
        If MGrid.TextMatrix(r, gSpecialAccount) = "LocalWholeSales" Or MGrid.TextMatrix(r, gSpecialAccount) = "Purchase" Then
            rs.AddNew
            rs!TransactionNo = Val(MGrid.TextMatrix(r, gVoucherNo))
            rs!TransactionType = IIf(Val(MGrid.TextMatrix(r, gReceipt)) > 0, "R", "P")
            rs!TransactionDate = DTPDate.Value
            rs!TransactionTime = Format(Time, "HH:MM AMPM")
            rs!AccountCode = Trim(MGrid.TextMatrix(r, gAccountCode))
            rs!Narration = Trim(MGrid.TextMatrix(r, gNarration))
            rs!CashOrCredit = "Cash"
            rs!Income = Val(MGrid.TextMatrix(r, gReceipt))
            rs!Expense = Val(MGrid.TextMatrix(r, gPayment))
            rs!BillNo = Right(Trim(MGrid.TextMatrix(r, gBillNo)), Len(Trim(MGrid.TextMatrix(r, gBillNo))) - 2)
            rs!SpecialAccount = Trim(MGrid.TextMatrix(r, gSpecialAccount))
            rs.Update
        End If
        
        r = r + 1
    Wend
    rs.Close
    
    MsgBox "Successfully Saved !", vbInformation
    clearControls
    getTransactionDetails
End Sub
Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub DTPDate_LostFocus()
    getTransactionDetails
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
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
    Set db = OpenDatabase("Storage.mdb", False, False, "MS Access;PWD=12345abcde")
    CoSaleType.AddItem "Local Sales(WholeSales)"
    CoSaleType.AddItem "Purchase"
    getAccounts
    MGridInitialise
    clearControls
    LVoucherNo.Caption = getNewTransactionNo
End Sub

Private Sub MGrid_Click()
Dim r As Long, i As Long

    If MGrid.Rows <= 0 Then
        Exit Sub
    End If
    
    If MGrid.TextMatrix(MGrid.Row, gSpecialAccount) = "LocalWholeSales" Or MGrid.TextMatrix(MGrid.Row, gSpecialAccount) = "Purchase" Then
        r = MGrid.Row
        TBillNo.Text = Right(Trim(MGrid.TextMatrix(r, gBillNo)), Len(Trim(MGrid.TextMatrix(r, gBillNo))) - 2)
        lSelectedVoucherNo = Val(MGrid.TextMatrix(r, gVoucherNo))
        LVoucherNo.Caption = lSelectedVoucherNo
        TNarration.Text = Trim(MGrid.TextMatrix(r, gNarration))
        TReceipt.Text = Val(MGrid.TextMatrix(r, gReceipt))
        TPayment.Text = Val(MGrid.TextMatrix(r, gPayment))
        CoAccounts.Text = Trim(MGrid.TextMatrix(r, gAccount))
    Else
    End If
End Sub

Private Sub MGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub TBillNo_LostFocus()
    getBillNoAccountDetails Trim(TBillNo.Text), CoSaleType.Text
    LCurrentBalance.Caption = Format(getBalance, "0.00")
End Sub

Private Sub TReceipt_GotFocus()
    TReceipt.SelStart = 0
    TReceipt.SelLength = Len(TReceipt.Text)
End Sub

Private Sub TPayment_GotFocus()
    TPayment.SelStart = 0
    TPayment.SelLength = Len(TPayment.Text)
End Sub

Private Sub getTransactionDetails()
Dim rs As Recordset, r As Long
        
    Set rs = db.OpenRecordset("Select AccountRegister.*,AccountMaster.AccountName From AccountRegister,AccountMaster Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') And AccountMaster.Code=AccountRegister.AccountCode And (AccountRegister.SpecialAccount='SaleAndPurchaseOrder' Or AccountRegister.SpecialAccount='SaleAndPurchaseReadyMade') ) Order By Val(AccountRegister.TransactionNo)")
    MGrid.Rows = 0
    If rs.RecordCount > 0 Then
        DTPDate.Value = DateValue("" & rs!TransactionDate)

        r = 0
        rs.MoveFirst
        While rs.EOF = False
            MGrid.AddItem ""
            MGrid.TextMatrix(r, gVoucherNo) = "" & rs!TransactionNo
            MGrid.TextMatrix(r, gAccountCode) = "" & rs!AccountCode
            MGrid.TextMatrix(r, gAccount) = "" & rs!AccountName
            MGrid.TextMatrix(r, gNarration) = "" & rs!Narration
            MGrid.TextMatrix(r, gReceipt) = "" & rs!Income
            MGrid.TextMatrix(r, gPayment) = "" & rs!Expense
            MGrid.TextMatrix(r, gBillNo) = "" & IIf("" & rs!SpecialAccount = "SaleAndPurchaseReadyMade", "R-", "O-") & rs!BillNo
            MGrid.TextMatrix(r, gSpecialAccount) = "" & rs!SpecialAccount
            r = r + 1
            rs.MoveNext
        Wend
        rs.Close
    Else
        rs.Close
    End If
        
    setBalance
End Sub

Private Sub setBalance()
    getTotalReceiptPayment
    LBalance.Caption = Format(Val(LTotalReceipt.Caption) - Val(LTotalPayment.Caption), "0.00")
    LCurrentBalance.Caption = Format(Val(LCurrentBalance.Caption) - Val(LTotalReceipt.Caption), "0.00")
End Sub
Private Sub getTotalReceiptPayment()
Dim r As Long, dReceipt As Double, dPayment As Double
    r = 0
    dReceipt = 0
    dPayment = 0
    While r < MGrid.Rows
        dReceipt = dReceipt + Val(MGrid.TextMatrix(r, gReceipt))
        dPayment = dPayment + Val(MGrid.TextMatrix(r, gPayment))
        r = r + 1
    Wend
    LTotalReceipt.Caption = Format(dReceipt, "0.00")
    LTotalPayment.Caption = Format(dPayment, "0.00")
End Sub

Private Sub printOpeningStock()

On Error GoTo GoOut
    Dim i, x, y As Double
    
    i = 0
    x = 500
    y = NewPage + 400
    
    While (i < MGrid.Rows)
    
        Printer.FontSize = 10
        Printer.FontBold = False
                
        x = 500
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Trim(MGrid.TextMatrix(i, gVoucherNo))
                       
        x = 1300
        Printer.CurrentX = x
        Printer.CurrentY = y
        'Printer.Print trimStringForPrinting(Trim(MGrid.TextMatrix(i, gBillingName) & "/" & MGrid.TextMatrix(i, gColour)), 6000)
        
        x = 7600
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Trim(MGrid.TextMatrix(i, gReceipt))
        
        x = 9100
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print Format(MGrid.TextMatrix(i, gPayment), "0.00")
        
        x = 10600
        Printer.CurrentX = x
        Printer.CurrentY = y
        'Printer.Print Format(MGrid.TextMatrix(i, gTotalAmount), "0.00")
        
        i = i + 1
        y = y + 400
        If (y > 13000) Then
            y = NewPage + 400
        End If
    Wend
    
    y = 13300
    Printer.FontBold = True
    Printer.CurrentX = 3000
    Printer.CurrentY = y
    Printer.Print "Bill Amount"
    
    Printer.CurrentX = 10600
    Printer.CurrentY = y
    'Printer.Print Format(LGrandAmount.Caption, "0.00")
    
    y = y + 400
    Printer.FontBold = True
    Printer.CurrentX = 3000
    Printer.CurrentY = y
    Printer.Print "Advance"
    
    Printer.CurrentX = 10600
    Printer.CurrentY = y
    'Printer.Print Format(TAdvance.Text, "0.00")
    
    y = y + 400
    Printer.FontBold = True
    Printer.CurrentX = 3000
    Printer.CurrentY = y
    Printer.Print "Balance"
    
    Printer.CurrentX = 10600
    Printer.CurrentY = y
    'Printer.Print Format(LBalance.Caption, "0.00")
    
    Printer.EndDoc
    
    x = MsgBox("Successfully Printed !", vbInformation)
    
GoOut:
End Sub

Private Function NewPage() As Long

    Dim i, j, x, y As Double
    
    Printer.ScaleMode = 1
    Printer.FontName = "Arial"
    
    Printer.FontBold = True
    Printer.FontSize = 20
    Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("DYNAMIC DIGITAL SPOT")) / 2)
    Printer.CurrentY = 400
    Printer.Print "DYNAMIC DIGITAL SPOT"
    
    x = 400
    y = 800
    
    Printer.FontUnderline = True
    Printer.FontSize = 16
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.FontUnderline = True
    Printer.Print "Ink - Opening Stock"
    
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.FontUnderline = False
    Printer.CurrentX = x
    y = y + 500
    Printer.CurrentY = y
    Printer.Print "No"
    
    x = x + 800
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print ": "
    
    x = x + 100
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print Trim(lFirstVoucherNo)
    
    x = x + 6000
    Printer.FontBold = False
    Printer.FontSize = 10
    Printer.FontUnderline = False
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Customer"
    
    x = x + 800
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print ": "
    
    x = x + 100
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    'Printer.Print trimStringForPrinting(Trim(CoCustomer.Text), 3500)
    
    x = 400
    y = y + 400
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Ref No"
    
    x = x + 800
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print ": "
    
    x = x + 100
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    'Printer.Print Trim(TRefNo.Text)
    
    x = x + 6000
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Address"
    
    x = x + 800
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print ": "
    
    x = x + 100
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    'Printer.Print trimStringForPrinting(Trim(TAddress.Text), 3500)
    
    
    x = 500
    y = y + 1200
    
    Printer.FontBold = True
    
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "SlNo"
    
    x = 100 + 1200
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Description"
    
    x = 100 + 7500
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Qty"
    
    x = 100 + 9000
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Rate"
    
    x = 100 + 10500
    Printer.FontSize = 10
    Printer.CurrentX = x
    Printer.CurrentY = y
    Printer.Print "Amount"
    
    'HORIZONTAL LINES
    Printer.Line (400, 2800)-(12000, 2800)
    Printer.Line (400, 3200)-(12000, 3200)
    Printer.Line (400, 13200)-(12000, 13200)
    
    'FIRST AND LAST VERTICAL LINE
    Printer.Line (400, 2800)-(400, 13200)
    Printer.Line (12000, 2800)-(12000, 13200)
    
    'INNER LINES
    Printer.Line (1200, 2800)-(1200, 13200)
    'Printer.Line (6500, 2800)-(6500, 13200)
    Printer.Line (7500, 2800)-(7500, 13200)
    Printer.Line (9000, 2800)-(9000, 13200)
    Printer.Line (10500, 2800)-(10500, 13200)
    NewPage = y
End Function

