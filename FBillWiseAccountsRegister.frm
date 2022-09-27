VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FBillWiseAccountsRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill - Wise Accounts Register"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FBillWiseAccountsRegister.frx":0000
   ScaleHeight     =   7470
   ScaleWidth      =   12225
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   500
      Left            =   10305
      Picture         =   "FBillWiseAccountsRegister.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6930
      Width           =   1365
   End
   Begin VB.CommandButton CToExcel 
      Height          =   500
      Left            =   2140
      Picture         =   "FBillWiseAccountsRegister.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6930
      Width           =   1365
   End
   Begin VB.CommandButton CShow 
      Height          =   500
      Left            =   480
      Picture         =   "FBillWiseAccountsRegister.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6930
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   105
      TabIndex        =   4
      Top             =   1755
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   7329
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColorFixed  =   8421504
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   4575
      Left            =   90
      Top             =   1365
      Width           =   12045
   End
   Begin MSForms.ComboBox CoAccountType 
      Height          =   345
      Left            =   9585
      TabIndex        =   1
      Top             =   510
      Width           =   2490
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4392;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TBillNo 
      Height          =   345
      Left            =   9600
      TabIndex        =   2
      Top             =   915
      Width           =   2490
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "4392;609"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoSpecialAccount 
      Height          =   345
      Left            =   9585
      TabIndex        =   0
      Top             =   105
      Width           =   2490
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4392;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   345
      Left            =   8505
      TabIndex        =   19
      Top             =   510
      Width           =   705
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Type"
      Size            =   "1244;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LBalance 
      Height          =   405
      Left            =   10710
      TabIndex        =   18
      Top             =   6570
      Width           =   1770
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "3122;714"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LPayment 
      Height          =   405
      Left            =   10710
      TabIndex        =   17
      Top             =   6090
      Width           =   1755
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "3096;714"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   10440
      TabIndex        =   16
      Top             =   1455
      Width           =   1140
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Payment"
      Size            =   "2011;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   2625
      TabIndex        =   15
      Top             =   1470
      Width           =   2610
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "4604;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   6450
      TabIndex        =   14
      Top             =   1455
      Width           =   1095
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Description"
      Size            =   "1931;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   45
      TabIndex        =   13
      Top             =   1455
      Width           =   1410
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Date"
      Size            =   "2487;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label15 
      Height          =   330
      Left            =   1230
      TabIndex        =   12
      Top             =   1470
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Bill No"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   8925
      TabIndex        =   11
      Top             =   1455
      Width           =   1140
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Receipt"
      Size            =   "2011;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LReceipt 
      Height          =   405
      Left            =   9015
      TabIndex        =   10
      Top             =   6105
      Width           =   1665
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2937;714"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5115
      TabIndex        =   9
      Top             =   -540
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label7 
      Height          =   345
      Left            =   7755
      TabIndex        =   8
      Top             =   105
      Width           =   1365
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Special Account"
      Size            =   "2408;617"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   345
      Left            =   8565
      TabIndex        =   7
      Top             =   915
      Width           =   570
      ForeColor       =   -2147483640
      VariousPropertyBits=   8388627
      Caption         =   "Bill No"
      Size            =   "1005;609"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   495
      Left            =   90
      TabIndex        =   20
      Top             =   1365
      Width           =   12045
      BackColor       =   15724527
      Size            =   "21255;873"
      Picture         =   "FBillWiseAccountsRegister.frx":205968
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FBillWiseAccountsRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gDate As Single, gBillNo As Single, gAccount As Single, gDescription As Single, gReceipt As Single, gPayment As Single
Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gDate = 0
    gBillNo = 1
    gAccount = 2
    gDescription = 3
    gReceipt = 4
    gPayment = 5
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 6
    MGrid.Rows = 0
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gBillNo) = 1200
    MGrid.ColWidth(gAccount) = 2075
    MGrid.ColWidth(gDescription) = 4000
    MGrid.ColWidth(gReceipt) = 1500
    MGrid.ColWidth(gPayment) = 1500
    
    MGrid.RowHeightMin = 350
End Sub

Private Sub CoAccountType_Change()
    MGrid.Rows = 0
End Sub
Private Sub CoSpecialAccount_Change()
    MGrid.Rows = 0
End Sub

Private Sub CShow_Click()
Dim rs As Recordset
    MGrid.Rows = 0

    If (CoSpecialAccount.ListIndex = 0) Then
        If (Val("" & TBillNo.Text) = 0) And ((CoAccountType.ListIndex) = -1 Or (CoAccountType.ListIndex) = 0) Then
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount In( 'Sales','SRBillVoucher') ) And (AccountRegister.CashOrCredit In ('Cash','Credit') ) Order By Val(AccountRegister.TransactionNo) Asc,AccountRegister.TransactionDate")
        ElseIf (Val("" & TBillNo.Text) = 0) And (CoAccountType.ListIndex) = 1 Then
            TBillNo.Text = Trim(TBillNo.Text)
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount In( 'Sales','SRBillVoucher' )) And (AccountRegister.CashOrCredit In ('Cash') ) Order By Val(AccountRegister.TransactionNo) Asc,AccountRegister.TransactionDate")
        ElseIf (Val("" & TBillNo.Text) <> 0) And (CoAccountType.ListIndex) = 1 Then
            TBillNo.Text = Trim(TBillNo.Text)
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount In ( 'Sales','SRBillVoucher' ))And (AccountRegister.BillNo ='" & TBillNo.Text & "') And (AccountRegister.CashOrCredit ='Cash' ) Order By Val(AccountRegister.TransactionNo) Asc,AccountRegister.TransactionDate")
        Else
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount In( 'Sales','SRBillVoucher' )) And (AccountRegister.BillNo ='" & TBillNo.Text & "') Order By Val(AccountRegister.TransactionNo) Asc,AccountRegister.TransactionDate")
        End If
        
    ElseIf (CoSpecialAccount.ListIndex = 1) Then
        If (Val("" & TBillNo.Text) = 0) And ((CoAccountType.ListIndex) = -1 Or (CoAccountType.ListIndex) = 0) Then
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount In( 'Purchase','PBillVoucher') ) And (AccountRegister.CashOrCredit In ('Cash','Credit') ) Order By Val(AccountRegister.TransactionNo) Asc,AccountRegister.TransactionDate")
        ElseIf (Val("" & TBillNo.Text) = 0) And (CoAccountType.ListIndex) = 1 Then
            TBillNo.Text = Trim(TBillNo.Text)
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount In ( 'Purchase','PBillVoucher' )) And (AccountRegister.CashOrCredit In ('Cash') ) Order By Val(AccountRegister.TransactionNo) Asc,AccountRegister.TransactionDate")
        ElseIf (Val("" & TBillNo.Text) <> 0) And (CoAccountType.ListIndex) = 1 Then
            TBillNo.Text = Trim(TBillNo.Text)
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount In ( 'Purchase','PBillVoucher' ))And (AccountRegister.BillNo ='" & TBillNo.Text & "') And (AccountRegister.CashOrCredit ='Cash' ) Order By Val(AccountRegister.TransactionNo) Asc,AccountRegister.TransactionDate")
        Else
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount In ('Purchase','PBillVoucher' )) And (AccountRegister.BillNo ='" & TBillNo.Text & "') Order By Val(AccountRegister.TransactionNo) Asc,AccountRegister.TransactionDate")
        End If
        

        
    Else
        MsgBox "Please Select a SpecialAccount !", vbInformation
        CoSpecialAccount.SetFocus
        Exit Sub
    End If
    If rs.RecordCount = 0 Then Exit Sub
    While rs.EOF = False
        MGrid.AddItem Format("" & rs!TransactionDate, "dd-MM-yyyy") & vbTab & "" & rs!BillNo & vbTab & "" & rs!AccountName & vbTab & "" & rs!Narration & vbTab & Format(Val("" & rs!Income), "0.00") & vbTab & Format(Val("" & rs!Expense), "0.00")
        rs.MoveNext
    Wend
    rs.Close
    
    getTotals
End Sub

Private Sub CToExcel_Click()
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
    ReDim xData(1 To lRowCount + 2, 1 To lColCount) As Variant
    Dim i As Long, j As Long

    Set oExcel = OLEExcel.object
    Set oExcelSheet = oExcel.Sheets(1)

    xData(1, 1) = "Date"
    xData(1, 2) = "Bill No"
    xData(1, 3) = "Account"
    xData(1, 4) = "Description"
    xData(1, 5) = "Reciept"
    xData(1, 6) = "Payment"

    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    'xData(i + 1, 4) = LBillAmount.Caption
    'xData(i + 1, 5) = LAdvance.Caption
    'xData(i + 1, 6) = LBalance.Caption
    
    oExcelSheet.Range("A3:F" & lRowCount + 4).Value = xData

    'oExcelSheet.Cells(1, 1).Value = "Laser Sale Bill Wise Summary From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:F" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    'lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Private Sub DTPFrom_Change()
    MGrid.Rows = 0
    getTotals
End Sub

Private Sub DTPTo_Change()
    MGrid.Rows = 0
    getTotals
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CShow_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CToExcel_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    MGridInitialise
    
    CoSpecialAccount.AddItem "Sale"
'    CoSpecialAccount.AddItem "Purchase"

 
    CoAccountType.AddItem "Credit"
    CoAccountType.AddItem "Cash"
End Sub

Private Sub getTotals()
Dim r As Long
Dim dReceipt As Double, dPayment As Double
    r = 0
    dReceipt = 0
    dPayment = 0
    While r < MGrid.Rows
        dReceipt = dReceipt + Val(MGrid.TextMatrix(r, gReceipt))
        dPayment = dPayment + Val(MGrid.TextMatrix(r, gPayment))
        r = r + 1
    Wend
    LReceipt.Caption = Format("" & dReceipt, "0.00")
    LPayment.Caption = Format("" & dPayment, "0.00")
    LBalance.Caption = Format("" & (Val(dPayment) - Val(dReceipt)), "0.00")
End Sub
Private Sub TBillNo_Change()
    MGrid.Rows = 0
End Sub
