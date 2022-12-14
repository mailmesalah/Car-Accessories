VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FOpeningBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts -Opening Balance"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FOpeningBalance.frx":0000
   ScaleHeight     =   4080
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CNew 
      Height          =   525
      Left            =   75
      Picture         =   "FOpeningBalance.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3390
      Width           =   1365
   End
   Begin VB.CommandButton CPrint 
      Height          =   525
      Left            =   1605
      Picture         =   "FOpeningBalance.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3390
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   525
      Left            =   3135
      Picture         =   "FOpeningBalance.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3390
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   525
      Left            =   4665
      Picture         =   "FOpeningBalance.frx":205968
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3390
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   525
      Left            =   3915
      Picture         =   "FOpeningBalance.frx":207DCA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   390
      Left            =   2310
      TabIndex        =   1
      Top             =   90
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   688
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
      Format          =   20643843
      CurrentDate     =   40458
   End
   Begin MSForms.Label Label2 
      Height          =   375
      Left            =   3930
      TabIndex        =   16
      Top             =   1185
      Width           =   690
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Type"
      Size            =   "1217;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoType 
      Height          =   375
      Left            =   4515
      TabIndex        =   4
      Top             =   1200
      Width           =   1470
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2593;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LCurrentBalance 
      Height          =   375
      Left            =   540
      TabIndex        =   11
      Top             =   2595
      Width           =   4770
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Current Balance"
      Size            =   "8414;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   375
      Left            =   555
      TabIndex        =   10
      Top             =   1620
      Width           =   1785
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "3149;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TNarration 
      Height          =   375
      Left            =   2010
      TabIndex        =   5
      Top             =   1620
      Width           =   3975
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "7011;661"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   375
      Left            =   555
      TabIndex        =   9
      Top             =   1185
      Width           =   1785
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Amount"
      Size            =   "3149;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TAmount 
      Height          =   375
      Left            =   2010
      TabIndex        =   3
      Top             =   1185
      Width           =   1740
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "3069;661"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   375
      Left            =   420
      TabIndex        =   8
      Top             =   90
      Width           =   465
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "No"
      Size            =   "820;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TVoucherNo 
      Height          =   375
      Left            =   885
      TabIndex        =   0
      Top             =   90
      Width           =   1395
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "2461;661"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   375
      Left            =   555
      TabIndex        =   7
      Top             =   735
      Width           =   1785
      ForeColor       =   -2147483641
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "3149;661"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoAccount 
      Height          =   375
      Left            =   2010
      TabIndex        =   2
      Top             =   750
      Width           =   3975
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7011;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FOpeningBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sAccountCode() As String

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDelete_Click()
Dim rs As Recordset

    If Trim(TVoucherNo.Text) = "" Then
        MsgBox "Please Enter a Transaction No !", vbInformation
        TVoucherNo.SetFocus
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select AccountTransaction.* From AccountTransaction Where (AccountTransaction.BillNo = '" & Trim(TVoucherNo.Text) & "' ) And (AccountTransaction.Type = 'O' ) And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    MsgBox "Successfully Deleted !", vbInformation
    clearControls
    TVoucherNo.Text = getNewTransactionNo
    TVoucherNo.SetFocus
End Sub

Private Sub CNew_Click()
    clearControls
    TVoucherNo.Text = getNewTransactionNo
    TVoucherNo.SetFocus
End Sub

Private Sub CoAccount_Change()
Dim dBalance As Double
    dBalance = getCurrentBalanceOf(sAccountCode(CoAccount.ListIndex + 1))
    LCurrentBalance.Caption = "Current Balance is " & IIf(dBalance >= 0, Format(dBalance, "0.00") & " Dr", Format(Abs(dBalance), "0.00") & " Cr")
End Sub

Private Sub CoAccount_GotFocus()
    CoAccount.SelStart = 0
    CoAccount.SelLength = Len(CoAccount.Text)
End Sub

Private Sub CSave_Click()
Dim rs As Recordset
Dim sStatus As String

    If Trim(TVoucherNo.Text) = "" Then
        MsgBox "Please give a Transaction No to Edit or Click New to Add new !", vbInformation
        CNew.SetFocus
        Exit Sub
    End If
    
    If CoAccount.ListIndex = -1 Then
        MsgBox "Please Select an Account !", vbInformation
        CoAccount.SetFocus
        Exit Sub
    End If
    
    If Val("" & TAmount.Text) <= 0 Then
        MsgBox "Please Enter valid Amount !", vbInformation
        TAmount.SetFocus
        Exit Sub
    End If
    
    If CoType.ListIndex = -1 Then
        MsgBox "Please Select Debit or Credit !", vbInformation
        CoType.SetFocus
        Exit Sub
    End If
    
    If Trim(TNarration.Text) = "" Then
        MsgBox "Please Enter valid Narration !", vbInformation
        TNarration.SetFocus
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select AccountTransaction.* From AccountTransaction Where (AccountTransaction.BillNo = '" & Trim(TVoucherNo.Text) & "' ) And (AccountTransaction.Type = 'O' ) And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
    
    If CoType.Text = "Dr" Then
        rs.AddNew
        rs!BillNo = "" & TVoucherNo.Text
        rs!Type = "O"
        rs!AccountCode = sAccountCode(CoAccount.ListIndex + 1)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Credit = 0
        rs!Debit = Val(TAmount.Text)
        rs!Narration = "" & TNarration.Text
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = "1"
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!GCode = getGCodeOfAccount(sAccountCode(CoAccount.ListIndex + 1))
        rs!CreditedDebitedTo = "Opening Stock"
        rs!Mode = "Credit"
        rs.Update
        
        rs.AddNew
        rs!BillNo = "" & TVoucherNo.Text
        rs!Type = "O"
        rs!AccountCode = sReceivableAccount
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TAmount.Text)
        rs!Credit = 0
        rs!Narration = "" & TNarration.Text
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = "2"
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!GCode = getGCodeOfAccount(sReceivableAccount)
        rs!CreditedDebitedTo = CoAccount.Text
        rs!Mode = "Credit"
        rs.Update
    Else
        rs.AddNew
        rs!BillNo = "" & TVoucherNo.Text
        rs!Type = "O"
        rs!AccountCode = sAccountCode(CoAccount.ListIndex + 1)
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Credit = Val(TAmount.Text)
        rs!Debit = 0
        rs!Narration = "" & TNarration.Text
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = "1"
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!GCode = getGCodeOfAccount(sAccountCode(CoAccount.ListIndex + 1))
        rs!CreditedDebitedTo = "Opening Stock"
        rs!Mode = "Credit"
        rs.Update
        
        rs.AddNew
        rs!BillNo = "" & TVoucherNo.Text
        rs!Type = "O"
        rs!AccountCode = sPayableAccount
        rs!AddedDate = DTPDate.Value
        rs!EditedDate = DTPDate.Value
        rs!Debit = Val(TAmount.Text)
        rs!Credit = 0
        rs!Narration = "" & TNarration.Text
        rs!AddedBy = sCurrentUserCode
        rs!EditedBy = sCurrentUserCode
        rs!SerialNo = "2"
        rs!FinancialCode = getFinancialCode(DTPDate.Value)
        rs!GCode = getGCodeOfAccount(sPayableAccount)
        rs!CreditedDebitedTo = CoAccount.Text
        rs!Mode = "Credit"
        rs.Update
    End If
    MsgBox "Successfully " & sStatus & " !", vbInformation
    clearControls
    TVoucherNo.Text = getNewTransactionNo
    TVoucherNo.SetFocus
End Sub

Private Sub DTPDate_Change()
    TVoucherNo.Text = getNewTransactionNo
End Sub

Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       CoAccount.SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    DTPDate.Value = Date
    TVoucherNo = getNewTransactionNo
    getAccountsToCombo
    
    CoType.AddItem "Dr"
    CoType.AddItem "Cr"
End Sub

Private Function getNewTransactionNo() As String
Dim rs As Recordset, sTransactionNo As String

    Set rs = db.OpenRecordset("Select Max(Val(AccountTransaction.BillNo)) As TNo From AccountTransaction Where (AccountTransaction.Type = 'O') And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    If rs.RecordCount > 0 Then
        sTransactionNo = Val("" & rs!TNo) + 1
    Else
        sTransactionNo = 1
    End If
    rs.Close
    
    getNewTransactionNo = sTransactionNo
End Function

Private Sub getAccountsToCombo()
Dim rs As Recordset
    
    CoAccount.Clear
    
    Set rs = db.OpenRecordset("Select AccountRegister.Code,AccountRegister.AccountName From AccountRegister Where (AccountRegister.Type = 'BAccount' And AccountRegister.IsEnabled = True ) Order By AccountRegister.AccountName")
    ReDim sAccountCode(rs.RecordCount) As String
    While rs.EOF = False
        CoAccount.AddItem UCase("" & rs!AccountName)
        sAccountCode(CoAccount.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Public Sub getTransactionDetails(sTransactionNo As String)
Dim rs As Recordset

    clearControls
    Set rs = db.OpenRecordset("Select AccountTransaction.*,AccountRegister.AccountName From AccountTransaction,AccountRegister Where (AccountTransaction.BillNo = '" & Trim(sTransactionNo) & "' )  And (AccountTransaction.Type = 'O' ) And (AccountRegister.Code=AccountTransaction.AccountCode) And (AccountTransaction.FinancialCode='" & getFinancialCode(DTPDate.Value) & "')")
    If rs.RecordCount > 0 Then
        DTPDate.Value = DateValue("" & rs!AddedDate)
        TAmount.Text = Val("" & rs!Credit) + Val("" & rs!Debit)
        TNarration.Text = "" & rs!Narration
        CoAccount.Text = "" & rs!AccountName
        CoType = IIf(Val("" & rs!Debit) > 0, "Dr", "Cr")
    Else
        TVoucherNo.Text = getNewTransactionNo
        TVoucherNo.SetFocus
    End If
    rs.Close
End Sub

Private Sub clearControls()
    
    'DTPDate.Value = Date
    CoAccount.Text = ""
    TAmount.Text = ""
    TNarration.Text = ""
    LCurrentBalance.Caption = ""
    CoType.Text = ""
End Sub

Private Sub TAmount_GotFocus()
    TAmount.SelStart = 0
    TAmount.SelLength = Len(TAmount.Text)
End Sub

Private Sub TNarration_GotFocus()
    TNarration.SelStart = 0
    TNarration.SelLength = Len(TNarration.Text)
End Sub

Private Sub TVoucherNo_GotFocus()
    TVoucherNo.SelStart = 0
    TVoucherNo.SelLength = Len(TVoucherNo.Text)
End Sub

Private Sub TVoucherNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        getTransactionDetails (TVoucherNo.Text)
    End If
End Sub

Private Sub printPayment()

    'Dim i, j, x, y As Double
   
    'Printer.ScaleMode = 1
    'Printer.FontName = "Arial"
    
    'Printer.FontBold = True
    'Printer.FontSize = 20
    'Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("DYNAMIC DIGITAL SPOT")) / 2)
    'Printer.CurrentY = 400
    'Printer.Print "DYNAMIC DIGITAL SPOT"
    
    'x = 400
    'y = 900
    
    'Printer.FontSize = 12
    'Printer.CurrentX = (Val(Printer.Width) / 2) - (Val(Printer.TextWidth("Receipt")) / 2)
    'Printer.CurrentY = y
    'Printer.Print "Receipt"
    
    'Printer.FontBold = False
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'y = y + 500
    'Printer.CurrentY = y
    'Printer.Print "No"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Trim(TVoucherNo.Text)
    
    'Printer.FontBold = False
    'Printer.FontSize = 10
    'Printer.FontUnderline = False
    'Printer.CurrentX = 4000
    'Printer.CurrentY = y
    'Printer.Print Format(DTPDate.Value, "dd-MMM-yyyy")
    
    'x = 400
    'y = y + 400
    'Printer.FontBold = False
    'Printer.FontSize = 10
    'Printer.FontUnderline = False
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print "Account"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Trim(CoAccount.Text)
    
    'x = 400
    'y = y + 400
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print "Amount"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Format(TAmount.Text, "0.00")
    
    'x = 400
    'y = y + 400
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print "Narration"
    
    'x = x + 1000
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print ": "
    
    'x = x + 100
    'Printer.FontSize = 10
    'Printer.CurrentX = x
    'Printer.CurrentY = y
    'Printer.Print Trim(TNarration.Text & " " & CoSubAccount.Text)
    
    'Printer.EndDoc
    
    'MsgBox "Successfully send to Printer !", vbInformation
    
    DoEvents    'will not wait to complete the printing,lets to do other things while printing
    
    Dim i As Long, lines As Long, lReturnValue As Long
    
    'checking if the data is already entered
    On Error GoTo GoOut
    
    Open "LPT1:" For Output As #1
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1, Chr(27) & "j" & Chr(216)
    Print #1,
    Print #1, Chr(27) & "!" & Chr(20) & "    " & Chr(0) & Chr(27) & "!" & Chr(50) & "DeXtop" & Chr(27) & "!" & Chr(0) & Chr(27) & "!" & Chr(20) & " Software Innovations" & Chr(0)
    Print #1, Chr(27) & "!" & Chr(20) & "    " & "Payment" & Chr(0)
    Print #1,
    Print #1, Chr(27) & "!" & Chr(20) & "No        : R/" & Left(Trim(TVoucherNo.Text & "") & Space(22), 22) & Space(90) & " Date: " & Left(Format(DTPDate.Value, "dd-MMM-yyyy") & Space(12), 12) & Chr(0)
    Print #1, Chr(27) & "!" & Chr(20) & "Account   : " & Left(CoAccount.Text & Space(40), 40)
    Print #1, Chr(27) & "!" & Chr(20) & "Narration : " & Left(TNarration.Text & Space(40), 40)
    Print #1, Chr(27) & "!" & Chr(20) & "Amount    : " & Right(Space(12) & Format(TAmount.Text, "0.00"), 12)
    'Print #1, Chr(27) & "!" & Chr(20) & "Narration : " & Left(Trim(TNarration.Text & " " & CoSubAccount.Text) & Space(40), 40)
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
    
    lReturnValue = MsgBox("Successfully Send to Printed !", vbInformation)
    Exit Sub
GoOut:
    MsgBox "Check If Printer is available, " & Err.Description, vbInformation
End Sub


