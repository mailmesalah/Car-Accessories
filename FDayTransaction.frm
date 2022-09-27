VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FDayTransaction 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Account Transaction"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
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
   Icon            =   "FDayTransaction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FDayTransaction.frx":000C
   ScaleHeight     =   7665
   ScaleWidth      =   11565
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CAddItem 
      Height          =   505
      Left            =   495
      Picture         =   "FDayTransaction.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6345
      Width           =   1365
   End
   Begin VB.CommandButton CRemoveItem 
      Height          =   505
      Left            =   1920
      Picture         =   "FDayTransaction.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6345
      Width           =   1365
   End
   Begin VB.CommandButton CClear 
      Height          =   505
      Left            =   3345
      Picture         =   "FDayTransaction.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6345
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   505
      Left            =   2045
      Picture         =   "FDayTransaction.frx":205974
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   135
      Width           =   1365
   End
   Begin VB.CommandButton CPrint 
      Height          =   505
      Left            =   495
      Picture         =   "FDayTransaction.frx":207DD6
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7050
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   8235
      Picture         =   "FDayTransaction.frx":20A238
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7050
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   9720
      Picture         =   "FDayTransaction.frx":20C69A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7050
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3660
      Left            =   165
      TabIndex        =   6
      Top             =   1275
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
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   420
      Left            =   285
      TabIndex        =   0
      Top             =   180
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   741
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   45678595
      CurrentDate     =   40458
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   4790
      Left            =   150
      Top             =   780
      Width           =   11295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   150
      X2              =   11430
      Y1              =   4965
      Y2              =   4950
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   4635
      TabIndex        =   20
      Top             =   825
      Width           =   2400
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TNarration 
      Height          =   420
      Left            =   4275
      TabIndex        =   2
      Top             =   5040
      Width           =   3300
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5821;741"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   450
      Left            =   7965
      TabIndex        =   19
      Top             =   6585
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "2593;794"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBalance 
      Height          =   450
      Left            =   9450
      TabIndex        =   18
      Top             =   6585
      Width           =   1545
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2725;794"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label11 
      Height          =   420
      Left            =   7920
      TabIndex        =   17
      Top             =   5730
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
   Begin MSForms.Label LTotalPayment 
      Height          =   345
      Left            =   9450
      TabIndex        =   16
      Top             =   6135
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
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   7965
      TabIndex        =   15
      Top             =   6135
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
   Begin MSForms.Label LTotalReceipt 
      Height          =   420
      Left            =   9450
      TabIndex        =   14
      Top             =   5760
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
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   9450
      TabIndex        =   13
      Top             =   825
      Width           =   1560
      ForeColor       =   -2147483643
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
      Left            =   7875
      TabIndex        =   12
      Top             =   825
      Width           =   1170
      ForeColor       =   -2147483643
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
      Left            =   1650
      TabIndex        =   11
      Top             =   825
      Width           =   2400
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "4233;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   225
      TabIndex        =   10
      Top             =   825
      Width           =   1050
      ForeColor       =   -2147483643
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
      Left            =   9375
      TabIndex        =   4
      Top             =   5040
      Width           =   1710
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3016;741"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TReceipt 
      Height          =   420
      Left            =   7620
      TabIndex        =   3
      Top             =   5040
      Width           =   1710
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3016;741"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoAccounts 
      Height          =   420
      Left            =   1455
      TabIndex        =   1
      Top             =   5040
      Width           =   2775
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4895;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LVoucherNo 
      Height          =   390
      Left            =   180
      TabIndex        =   9
      Top             =   5040
      Width           =   1155
      ForeColor       =   -2147483642
      VariousPropertyBits=   8388627
      Caption         =   "Voucher No"
      Size            =   "2037;688"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   780
      Width           =   11325
      BackColor       =   15724527
      Size            =   "19976;873"
      Picture         =   "FDayTransaction.frx":20EAFC
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FDayTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sAccountCode() As String
Dim gVoucherNo As Single, gAccount As Single, gNarration As Single, gReceipt As Single, gPayment As Single, gAccountCode As Single, gSpecialAccount As Single
Dim lSelectedVoucherNo As Long, lFirstVoucherNo As Long

Private Sub CAddItem_Click()
Dim lYN As Long, r As Long

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
        MGrid.TextMatrix(MGrid.Rows - 1, gAccount) = Trim(CoAccounts.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gNarration) = Trim(TNarration.Text)
        MGrid.TextMatrix(MGrid.Rows - 1, gReceipt) = Format(Val(TReceipt.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gPayment) = Format(Val(TPayment.Text), "0.00")
        MGrid.TextMatrix(MGrid.Rows - 1, gAccountCode) = sAccountCode(CoAccounts.ListIndex + 1)
        MGrid.TextMatrix(MGrid.Rows - 1, gSpecialAccount) = "Payroll"
    Else
        r = MGrid.Row
        MGrid.TextMatrix(r, gAccount) = Trim(CoAccounts.Text)
        MGrid.TextMatrix(r, gNarration) = Trim(TNarration.Text)
        MGrid.TextMatrix(r, gReceipt) = Format(Val(TReceipt.Text), "0.00")
        MGrid.TextMatrix(r, gPayment) = Format(Val(TPayment.Text), "0.00")
        MGrid.TextMatrix(r, gAccountCode) = sAccountCode(CoAccounts.ListIndex + 1)
        MGrid.TextMatrix(r, gSpecialAccount) = "Payroll"
    End If
    lSelectedVoucherNo = 0
    LVoucherNo.Caption = getNewTransactionNo
    clearEditControls
    setBalance
    CoAccounts.SetFocus
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
    gAccount = 1
    gNarration = 2
    gReceipt = 3
    gPayment = 4
    gAccountCode = 5
    gSpecialAccount = 6
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 7
    MGrid.Rows = 0
    MGrid.ColWidth(gVoucherNo) = 1200
    MGrid.ColWidth(gAccount) = 2965
    MGrid.ColWidth(gNarration) = 3300
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
    
    Set rs = db.OpenRecordset("Select AccountMaster.Code,AccountMaster.AccountName From AccountMaster Where (AccountMaster.Type ='BAccount')")
    
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
    CoAccounts.ListIndex = -1
    TNarration.Text = ""
    TReceipt.Text = ""
    TPayment.Text = ""
    LVoucherNo.Caption = getNewTransactionNo
    setBalance
End Sub

Private Sub clearEditControls()
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
        Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') And AccountRegister.SpecialAccount='Payroll' )")
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
        FAccountRegister.Show vbModal
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
    
    If MGrid.TextMatrix(MGrid.Row, gSpecialAccount) = "Payroll" Then
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
    Set rs = db.OpenRecordset("Select AccountRegister.* From AccountRegister Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') And AccountRegister.SpecialAccount='Payroll' )")
        
    'SAVES DATA TO TransactionRegister ReadyMade
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    r = 0
    While r < MGrid.Rows
        If (MGrid.TextMatrix(r, gSpecialAccount) = "Payroll") Then
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
            rs!BillNo = ""
            rs!SpecialAccount = "Payroll"
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
        CoAccounts.SetFocus
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
    getAccounts
    MGridInitialise
    clearControls
    getTransactionDetails
    LVoucherNo.Caption = getNewTransactionNo
End Sub

Private Sub MGrid_Click()
Dim r As Long, i As Long
    
    If MGrid.Rows <= 0 Then
        Exit Sub
    End If
    
    If MGrid.TextMatrix(MGrid.Row, gSpecialAccount) = "Payroll" Then
        r = MGrid.Row
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
        CoAccounts.SetFocus
    End If
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
            
        Set rs = db.OpenRecordset("Select AccountRegister.*,AccountMaster.AccountName From AccountRegister,AccountMaster Where (AccountRegister.TransactionDate = cDate('" & DTPDate.Value & "') And AccountMaster.Code=AccountRegister.AccountCode And AccountRegister.SpecialAccount = 'Payroll' ) Order By Val(AccountRegister.TransactionNo)")
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
