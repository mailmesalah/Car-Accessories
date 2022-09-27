VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FBillVoucher1 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Voucher"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FBillVoucher1.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   11895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CDelete 
      Height          =   450
      Left            =   2070
      Picture         =   "FBillVoucher1.frx":225B92
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   405
      Width           =   1505
   End
   Begin VB.CommandButton CClear 
      Height          =   450
      Left            =   3720
      Picture         =   "FBillVoucher1.frx":228434
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   1505
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   450
      Left            =   2190
      Picture         =   "FBillVoucher1.frx":22ACD6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1505
   End
   Begin VB.CommandButton CAddItem 
      Height          =   450
      Left            =   660
      Picture         =   "FBillVoucher1.frx":22D578
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7080
      Width           =   1505
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   450
      Left            =   9720
      Picture         =   "FBillVoucher1.frx":22FE1A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8160
      Width           =   1505
   End
   Begin VB.CommandButton CSave 
      Height          =   450
      Left            =   8040
      Picture         =   "FBillVoucher1.frx":2326BC
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8160
      Width           =   1505
   End
   Begin VB.CommandButton CExport 
      Height          =   450
      Left            =   2280
      Picture         =   "FBillVoucher1.frx":234F5E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
      Width           =   1505
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3660
      Left            =   375
      TabIndex        =   8
      Top             =   1980
      Width           =   11190
      _ExtentX        =   19738
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
      Left            =   405
      TabIndex        =   0
      Top             =   420
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
      Format          =   53673987
      CurrentDate     =   40458
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   5
      Left            =   2040
      TabIndex        =   36
      Top             =   375
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   4
      Left            =   630
      TabIndex        =   35
      Top             =   7050
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   3
      Left            =   2160
      TabIndex        =   34
      Top             =   7050
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   2
      Left            =   3690
      TabIndex        =   33
      Top             =   7050
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   1
      Left            =   2250
      TabIndex        =   32
      Top             =   8130
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   0
      Left            =   8010
      TabIndex        =   31
      Top             =   8130
      Width           =   1545
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000012&
      Height          =   510
      Index           =   6
      Left            =   9690
      TabIndex        =   30
      Top             =   8130
      Width           =   1545
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   6405
      TabIndex        =   28
      Top             =   465
      Width           =   2400
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Type"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoType 
      Height          =   420
      Left            =   8730
      TabIndex        =   1
      Top             =   420
      Width           =   2625
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4630;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBillNo 
      Height          =   420
      Left            =   1755
      TabIndex        =   2
      Top             =   5790
      Width           =   1965
      VariousPropertyBits=   746604571
      Size            =   "3466;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   1515
      TabIndex        =   27
      Top             =   6480
      Width           =   2400
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   4680
      TabIndex        =   26
      Top             =   1620
      Width           =   2400
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TNarration 
      Height          =   420
      Left            =   3750
      TabIndex        =   4
      Top             =   6450
      Width           =   4065
      VariousPropertyBits=   746604571
      Size            =   "7170;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   420
      Left            =   8205
      TabIndex        =   25
      Top             =   7425
      Width           =   1470
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBalance 
      Height          =   345
      Left            =   9690
      TabIndex        =   24
      Top             =   7425
      Width           =   1545
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2725;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label11 
      Height          =   420
      Left            =   8160
      TabIndex        =   23
      Top             =   6570
      Width           =   1470
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Total Receipt"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalPayment 
      Height          =   345
      Left            =   9690
      TabIndex        =   22
      Top             =   6975
      Width           =   1545
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Total Payment0"
      Size            =   "2725;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   8205
      TabIndex        =   21
      Top             =   6975
      Width           =   1470
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Total Payment"
      Size            =   "2593;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LTotalReceipt 
      Height          =   420
      Left            =   9690
      TabIndex        =   20
      Top             =   6600
      Width           =   1545
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Total Receipt0"
      Size            =   "2725;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   9675
      TabIndex        =   19
      Top             =   1620
      Width           =   1560
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Payment"
      Size            =   "2752;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label16 
      Height          =   330
      Left            =   8100
      TabIndex        =   18
      Top             =   1620
      Width           =   1170
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Receipt"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1635
      TabIndex        =   17
      Top             =   1620
      Width           =   2400
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Bill No"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   450
      TabIndex        =   16
      Top             =   1620
      Width           =   1050
      ForeColor       =   -2147483644
      VariousPropertyBits=   8388627
      Caption         =   "Voucher No"
      Size            =   "1852;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TPayment 
      Height          =   420
      Left            =   9600
      TabIndex        =   6
      Top             =   5790
      Width           =   1710
      VariousPropertyBits=   746604571
      Size            =   "3016;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TReceipt 
      Height          =   420
      Left            =   7845
      TabIndex        =   5
      Top             =   5790
      Width           =   1710
      VariousPropertyBits=   746604571
      Size            =   "3016;741"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoAccounts 
      Height          =   420
      Left            =   3765
      TabIndex        =   3
      Top             =   5790
      Width           =   4035
      VariousPropertyBits=   746604571
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
   Begin MSForms.Label LVoucherNo 
      Height          =   390
      Left            =   405
      TabIndex        =   15
      Top             =   5805
      Width           =   1155
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Voucher No"
      Size            =   "2037;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   4800
      Left            =   360
      Top             =   1560
      Width           =   11235
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   360
      X2              =   11550
      Y1              =   5670
      Y2              =   5670
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   360
      X2              =   11580
      Y1              =   1965
      Y2              =   1965
   End
   Begin MSForms.Label Label5 
      Height          =   390
      Left            =   375
      TabIndex        =   37
      Top             =   1575
      Width           =   11190
      BackColor       =   4210752
      Size            =   "19738;688"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FBillVoucher1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim sAccountCode() As String
Dim gVoucherNo As Single, gAccount As Single, gNarration As Single, gReceipt As Single, gPayment As Single, gAccountCode As Single, gSpecialAccount As Single, gBillNo As Single
Dim lSelectedVoucherNo As Long, lFirstVoucherNo As Long
Dim sSpecialAccount(6) As String

Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

    If (CoType.ListIndex = -1) Then
        MsgBox "Please Select a Type !", vbInformation
        CoType.SetFocus
        Exit Sub
    End If
    
    If ifNotBillNoExist(Trim(TBillNo.Text), sSpecialAccount(CoType.ListIndex)) Then
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
        MGrid.TextMatrix(MGrid.Rows - 1, gBillNo) = IIf(Val(TReceipt.Text) > 0, "R-", "P-") & Trim(TBillNo.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gAccount) = Trim(CoAccounts.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gNarration) = Trim(TNarration.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(Val(TReceipt.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(Val(TPayment.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gAccountCode) = sAccountCode(CoAccounts.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gSpecialAccount) = sSpecialAccount(CoType.ListIndex)
    Else
        r = MGrid.Row
        MGrid.TextMatrix(r, gBillNo) = IIf(TReceipt.Text < 0, "R-", "P-") & Trim(TBillNo.Text)
        MGrid.TextMatrix(r, gAccount) = Trim(CoAccounts.Text)
        MGrid.TextMatrix(r, gNarration) = Trim(TNarration.Text)
        MGrid.TextMatrix(r, gReceipt) = Format(Val(TReceipt.Text), "0.00")
        MGrid.TextMatrix(r, gPayment) = Format(Val(TPayment.Text), "0.00")
        MGrid.TextMatrix(r, gAccountCode) = sAccountCode(CoAccounts.ListIndex + 1)
        MGrid.TextMatrix(r, gSpecialAccount) = sSpecialAccount(CoType.ListIndex)
    End If
    lSelectedVoucherNo = 0
    LVoucherNo.Caption = getNewTransactionNo
    
    clearEditControls
    setBalance
    TBillNo.SetFocus
    'CoAccounts.SetFocus
End Sub

Private Function ifNotBillNoExist(sBillNo As String, sType As String) As Boolean
Dim rs As Recordset, bNotExist As Boolean
    bNotExist = True
    
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.BillNo = '" & sBillNo & "' ) And (AccountRegister.SpecialAccount='" & sType & "') ")
    If rs.RecordCount > 0 Then
        bNotExist = False
    End If
    rs.Close
    
    ifNotBillNoExist = bNotExist
End Function

Private Sub getBillNoAccountDetails(sBillNo As String, sType As String)
    Dim rs As Recordset
    If CoType.ListIndex = 0 Then
        Set rs = db.OpenRecordset("Select AccountMaster.AccountName From AccountMaster,AccountRegister,CustomerMaster Where (AccountRegister.BillNo = '" & sBillNo & "' ) And (AccountRegister.SpecialAccount = '" & sType & "' ) And (CustomerMaster.AccountCode = AccountMaster.Code)")
        If rs.RecordCount > 0 Then
            CoAccounts.Text = "" & rs!AccountName
        End If
        rs.Close
    ElseIf CoType.ListIndex = 1 Then
        Set rs = db.OpenRecordset("Select AccountMaster.AccountName From AccountMaster,AccountRegister,SupplierMaster Where (AccountRegister.BillNo = '" & sBillNo & "' ) And (AccountRegister.SpecialAccount = '" & sType & "' ) And (SupplierMaster.AccountCode = AccountMaster.Code)")
        If rs.RecordCount > 0 Then
            CoAccounts.Text = "" & rs!AccountName
        End If
        rs.Close
    Else
        Exit Sub
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
    lTransactionNo = getNewAccountCode() - 1

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
    
    Set rs = db.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName From AccountMaster Where (AccountMaster.Status = True ) And (AccountMaster.Type = 'BAccount' )")
    
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
        Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') And (AccountRegister.SpecialAccount='" & sSpecialAccount(CoType.ListIndex) & "') )")
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

Private Sub CExport_Click()
On Error GoTo ErrHandler
Dim oExcel As Object, oExcelSheet As Object
Dim lReturnValue As Long
Dim lRowCount As Long, lColCount As Long

    If MGrid.Rows = 0 Then
        MsgBox "Empty Data!", vbInformation
        Exit Sub
    End If
  
    OLEExcel.CreateEmbed vbNullString, "Excel.Sheet"
    
    lRowCount = MGrid.Rows
    lColCount = MGrid.Cols
    ReDim xData(1 To lRowCount + 3, 1 To lColCount) As Variant
    Dim i As Long, j As Long

    Set oExcel = OLEExcel.object
    Set oExcelSheet = oExcel.Sheets(1)

    xData(1, 1) = "Voucher No"
    xData(1, 2) = "Bill No"
    xData(1, 3) = "Account"
    xData(1, 4) = "Narration"
    xData(1, 5) = "Receipt"
    xData(1, 6) = "Payment"
    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    xData(i + 1, 5) = Format(LTotalReceipt.Caption, "0.00")
    xData(i + 1, 6) = Format(LTotalPayment.Caption, "0.00")
    xData(i + 2, 5) = "Balance"
    xData(i + 2, 6) = Format(LBalance.Caption, "0.00")
    
    oExcelSheet.Range("A3:F" & lRowCount + 5).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Account Transaction of " & Format(DTPDate.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:F" & lRowCount + 5).Select
    oExcel.Application.Selection.AutoFormat

On Error Resume Next
    ' Delete the existing test file (if any)...
    Kill App.Path & "\Reports\Bill Voucher " & Format(DTPDate.Value, "dd-MMM-yyyy") & ".xlsx"

  ' Save the file as a native XLS file...
    oExcel.SaveAs App.Path & "\Reports\Bill Voucher " & Format(DTPDate.Value, "dd-MMM-yyyy") & ".xlsx"
    
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
  ' Close the OLE object and remove it...
    OLEExcel.Close
    OLEExcel.Delete
    
    'lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\DayBook " & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\Bill Voucher " & Format(DTPDate.Value, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical

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
    'print
End Sub


Private Sub CRemoveItem_Click()
Dim r As Long
    If MGrid.Rows <= 0 Then
        Exit Sub
    End If
    
    If MGrid.TextMatrix(MGrid.Row, gSpecialAccount) = "BillVoucher" Then
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
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') And (AccountRegister.SpecialAccount='" & sSpecialAccount(CoType.ListIndex) & "') )")
        
    'SAVES DATA TO TransactionRegister ReadyMade
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    r = 0
    While r < MGrid.Rows
        rs.AddNew
        rs!TransactionNo = Val(MGrid.TextMatrix(r, gVoucherNo))
        rs!Type = IIf(Val(MGrid.TextMatrix(r, gReceipt)) > 0, "R", "P")
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

    sSpecialAccount(0) = "LocalWholeSales"
    sSpecialAccount(1) = "Purchase"

    
    CoType.AddItem "LocalWholeSales"
    CoType.AddItem "Purchase"

    getAccounts
    MGridInitialise
    clearControls
    LVoucherNo.Caption = getNewTransactionNo
End Sub

Private Sub MGrid_Click()
Dim r As Long, i As Long

    If MGrid.Rows > 0 Then
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
    getBillNoAccountDetails Trim(TBillNo.Text), CoType.Text
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
    If (CoType.ListIndex = -1) Then
        Exit Sub
    End If
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') ) And (AccountRegister.SpecialAccount = 'LocalWholeSales' )And (AccountRegister.Type In ('P','R') )And (AccountRegister.BillNo = '" & TBillNo.Text & "' )  Order By AccountRegister.Type ")
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
            MGrid.TextMatrix(r, gBillNo) = "" & IIf("" & rs!TransactionType = "R", "R-", "P-") & rs!TransactionNo
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


