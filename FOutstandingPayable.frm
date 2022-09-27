VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FOutstandingPayable 
   Caption         =   "Outstanding Payable"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FOutstandingPayable.frx":0000
   ScaleHeight     =   6705
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CShow 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   240
      Picture         =   "FOutstandingPayable.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6030
      Width           =   1365
   End
   Begin VB.CommandButton CToExcel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7635
      Picture         =   "FOutstandingPayable.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6030
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9315
      Picture         =   "FOutstandingPayable.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6015
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   3315
      Left            =   75
      TabIndex        =   3
      Top             =   1605
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   5847
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
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   345
      Left            =   1650
      TabIndex        =   18
      Top             =   120
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
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
      Format          =   45678595
      CurrentDate     =   40909
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   1650
      TabIndex        =   19
      Top             =   555
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
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
      Format          =   45678595
      CurrentDate     =   40909
   End
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   105
      TabIndex        =   22
      Top             =   1215
      Width           =   1245
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Date"
      Size            =   "2196;582"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   90
      TabIndex        =   21
      Top             =   540
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   90
      TabIndex        =   20
      Top             =   150
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "1905;582"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   9300
      TabIndex        =   16
      Top             =   1215
      Width           =   1140
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Balance"
      Size            =   "2011;582"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   3675
      TabIndex        =   15
      Top             =   1230
      Width           =   1830
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "3228;741"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   6405
      TabIndex        =   14
      Top             =   1215
      Width           =   1170
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Debit"
      Size            =   "2064;741"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1635
      TabIndex        =   13
      Top             =   1215
      Width           =   1770
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Account Name"
      Size            =   "3122;582"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label17 
      Height          =   330
      Left            =   7920
      TabIndex        =   12
      Top             =   1215
      Width           =   1140
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Credit"
      Size            =   "2011;582"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LReceipt 
      Height          =   405
      Left            =   6255
      TabIndex        =   11
      Top             =   5085
      Width           =   1500
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2646;714"
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
      Left            =   7785
      TabIndex        =   10
      Top             =   5085
      Width           =   1395
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2461;714"
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
      Left            =   3645
      TabIndex        =   9
      Top             =   15
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label7 
      Height          =   420
      Left            =   6030
      TabIndex        =   8
      Top             =   210
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Type"
      Size            =   "2566;741"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LBalance 
      Height          =   405
      Left            =   9240
      TabIndex        =   7
      Top             =   5085
      Width           =   1470
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2593;714"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   420
      Left            =   6030
      TabIndex        =   6
      Top             =   675
      Width           =   1455
      VariousPropertyBits=   8388627
      Caption         =   "Account"
      Size            =   "2566;741"
      FontName        =   "Arial"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoAccount 
      Height          =   405
      Left            =   7680
      TabIndex        =   5
      Top             =   675
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;706"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoType 
      Height          =   405
      Left            =   7680
      TabIndex        =   4
      Top             =   225
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;706"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   495
      Left            =   60
      TabIndex        =   17
      Top             =   1140
      Width           =   10710
      BackColor       =   15724527
      Size            =   "18891;873"
      Picture         =   "FOutstandingPayable.frx":205968
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FOutstandingPayable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim gDate As Single, gVoucherNo As Single, gAccount As Single, gDescription As Single, gDebit As Single, gCredit As Single, gBalance As Single, gType As Single, gInvoiceType As Single, gInvoiceBillNo As Single, gNarration As Single
Dim sAccountCode() As String
Dim sAddress() As String

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gDate = 0
    gAccount = 1
    gDescription = 2
    gVoucherNo = 3
    gDebit = 4
    gCredit = 5
    gBalance = 6
    gType = 7
    gInvoiceType = 8
    gInvoiceBillNo = 9
    gNarration = 10
    
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 11
    MGrid.Rows = 0
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gAccount) = 4000
    MGrid.ColWidth(gDescription) = 1500
    MGrid.ColWidth(gVoucherNo) = 800
    MGrid.ColWidth(gDebit) = 1450
    MGrid.ColWidth(gCredit) = 1450
    MGrid.ColWidth(gBalance) = 1450
    MGrid.ColWidth(gType) = 0
    MGrid.ColWidth(gInvoiceType) = 0
    MGrid.ColWidth(gInvoiceBillNo) = 0
    MGrid.ColWidth(gNarration) = 0
    
    MGrid.RowHeightMin = 350
End Sub


Private Sub getAccounts()
Dim rs As Recordset
    
    CoAccount.Clear
    
    If (CoType.ListIndex = 0) Then
        Exit Sub
    ElseIf (CoType.ListIndex > 0) Then
    
        Set rs = db.OpenRecordset("Select AccountRegister.Code,AccountRegister.AccountName From AccountRegister Where (AccountRegister.GroupCode='" & sSupplierAccountParentID & "' )")
    Else
        Exit Sub
    End If
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sAccountCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoAccount.AddItem "" & rs!AccountName
        sAccountCode(CoAccount.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CoType_Change()
    getAccounts
End Sub

Private Sub CShow_Click()
Dim rs As Recordset, dOpeningBalance As Double, dBalance As Double, sCode As String
    
    MGrid.Rows = 0
    If (CoType.ListIndex = 0) Then
            Set rs = db.OpenRecordset("Select Sum(AR.Debit-AR.Credit) As OpeningBalance From AccountTransaction As AR Where (AR.GCode = '" & sSupplierAccountParentID & "') And (AR.EditedDate < cDate('" & DTPFrom.Value & "')) ")
            If rs.RecordCount > 0 Then
                dOpeningBalance = Val("" & rs!OpeningBalance)
            End If
            
            Set rs = db.OpenRecordset("Select AccountRegister.AccountName,AccountTransaction.* From AccountRegister,AccountTransaction Where (AccountRegister.Code = AccountTransaction.AccountCode )  And (AccountTransaction.GCode = '" & sSupplierAccountParentID & "' ) And (AccountTransaction.EditedDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Order By AccountTransaction.EditedDate,Val(AccountTransaction.BillNo)")
                        
    
    ElseIf (CoType.ListIndex = 1) Then
    
        If (CoAccount.ListIndex >= 0) Then
        
            Set rs = db.OpenRecordset("Select Sum(AR.Debit-AR.Credit) As OpeningBalance From AccountTransaction As AR Where ((AR.GCode='" & sAccountCode(CoAccount.ListIndex + 1) & "')) And (AR.EditedDate < cDate('" & DTPFrom.Value & "'))")
            If rs.RecordCount > 0 Then
                dOpeningBalance = Val("" & rs!OpeningBalance)
            End If
            
            Set rs = db.OpenRecordset("Select AccountRegister.AccountName,AccountTransaction.* From AccountRegister,AccountTransaction Where ((AccountTransaction.AccountCode='" & sAccountCode(CoAccount.ListIndex + 1) & "') ) And (AccountRegister.Code = AccountTransaction.AccountCode )  And (AccountTransaction.EditedDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Order By AccountTransaction.EditedDate,Val(AccountTransaction.BillNo)")
                            
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    dBalance = dOpeningBalance
    If dOpeningBalance <> 0 Then
        MGrid.AddItem Format(DTPFrom.Value, "dd.mm.yyyy") & vbTab & "Opening Balance" & vbTab & vbTab & "" & vbTab & Format(IIf(dOpeningBalance > 0, dOpeningBalance, 0), "0.00") & vbTab & Format(IIf(dOpeningBalance < 0, Abs(dOpeningBalance), 0), "0.00") & vbTab & Format(Abs(dBalance), "0.00") & IIf(dBalance >= 0, " Dr", " Cr")
    End If
    
    If (CoType.ListIndex = 0) Then
        While rs.EOF = False
            dBalance = dBalance + Val("" & rs!Debit) - Val(rs!Credit)
            MGrid.AddItem ""
            MGrid.TextMatrix((MGrid.Rows - 1), gDate) = Format("" & rs!EditedDate, "dd.mm.yyyy")
            MGrid.TextMatrix((MGrid.Rows - 1), gAccount) = UCase("" & rs!AccountName)
            MGrid.TextMatrix((MGrid.Rows - 1), gDescription) = UCase(IIf("" & rs!Type = "R" Or "" & rs!Type = "PY", "Receipt", IIf("" & rs!Type = "P" Or "" & rs!Type = "RV", "Payment", IIf("" & rs!Type = "PU", "Purchase", IIf("" & rs!Type = "PR", "Purchase Return", IIf("" & rs!Type = "S8", "Sales Form 8", IIf("" & rs!Type = "SB", "Sales Form 8B", IIf("" & rs!Type = "S8R", "Sales Return Form 8", IIf("" & rs!Type = "SBR", "Sales Return Form 8B", IIf("" & rs!Type = "O", "Opening Balance", "Others"))))))))))
            MGrid.TextMatrix((MGrid.Rows - 1), gVoucherNo) = "" & rs!BillNo
            MGrid.TextMatrix((MGrid.Rows - 1), gDebit) = Format(Val("" & rs!Debit), "0.00")
            MGrid.TextMatrix((MGrid.Rows - 1), gCredit) = Format(Val("" & rs!Credit), "0.00")
            MGrid.TextMatrix((MGrid.Rows - 1), gBalance) = Format(Abs(Val("" & dBalance)), "0.00") & IIf(dBalance >= 0, " Dr", " Cr")
            MGrid.TextMatrix((MGrid.Rows - 1), gType) = "" & rs!Type
            MGrid.TextMatrix((MGrid.Rows - 1), gInvoiceType) = "" & rs!InventoryType
            MGrid.TextMatrix((MGrid.Rows - 1), gInvoiceBillNo) = "" & rs!InventoryBillNo
            MGrid.TextMatrix((MGrid.Rows - 1), gNarration) = "" & rs!Narration
            
            rs.MoveNext
        Wend
    Else
        While rs.EOF = False
            dBalance = dBalance + Val("" & rs!Debit) - Val("" & rs!Credit)
            MGrid.AddItem ""
            MGrid.TextMatrix((MGrid.Rows - 1), gDate) = Format("" & rs!EditedDate, "dd.mm.yyyy")
            MGrid.TextMatrix((MGrid.Rows - 1), gAccount) = UCase("" & rs!AccountName)
            MGrid.TextMatrix((MGrid.Rows - 1), gDescription) = UCase(IIf("" & rs!Type = "R" Or "" & rs!Type = "PY", "Receipt", IIf("" & rs!Type = "P" Or "" & rs!Type = "RV", "Payment", IIf("" & rs!Type = "PU", "Purchase", IIf("" & rs!Type = "PR", "Purchase Return", IIf("" & rs!Type = "S8", "Sales Form 8", IIf("" & rs!Type = "SB", "Sales Form 8B", IIf("" & rs!Type = "S8R", "Sales Return Form 8", IIf("" & rs!Type = "SBR", "Sales Return Form 8B", IIf("" & rs!Type = "O", "Opening Balance", "Others"))))))))))
            MGrid.TextMatrix((MGrid.Rows - 1), gVoucherNo) = "" & rs!BillNo
            MGrid.TextMatrix((MGrid.Rows - 1), gDebit) = Format(Val("" & rs!Debit), "0.00")
            MGrid.TextMatrix((MGrid.Rows - 1), gCredit) = Format(Val("" & rs!Credit), "0.00")
            MGrid.TextMatrix((MGrid.Rows - 1), gBalance) = Format(Abs(Val("" & dBalance)), "0.00") & IIf(dBalance >= 0, " Dr", " Cr")
            MGrid.TextMatrix((MGrid.Rows - 1), gType) = "" & rs!Type
            MGrid.TextMatrix((MGrid.Rows - 1), gInvoiceType) = "" & rs!InventoryType
            MGrid.TextMatrix((MGrid.Rows - 1), gInvoiceBillNo) = "" & rs!InventoryBillNo
            MGrid.TextMatrix((MGrid.Rows - 1), gNarration) = "" & rs!Narration
            
            rs.MoveNext
        Wend
    End If
        
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
    xData(1, 2) = "Account"
    xData(1, 3) = "Description"
    xData(1, 4) = "Debit"
    xData(1, 5) = "Credit"
    xData(1, 6) = "Balance"

    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    'xData(i + 1, 4) = LBillAmount.Caption
    'xData(i + 1, 5) = LAdvance.Caption
    'xData(i + 1, 6) = LBalance.Caption
    
    oExcelSheet.Range("A3:F" & lRowCount + 4).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Outstanding Payables From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:F" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\Outstanding Payables " & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\Outstanding Payables " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    'lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\LaserSaleBillWiseSummary " & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\Outstanding Payables " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CShow_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CToExcel_Click
    'ElseIf (KeyCode = vbKeyP And ((Shift And 7) = 2)) Then
        'CPrint_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    DTPFrom.Value = Date
    DTPTo.Value = Date
    MGridInitialise
    CoType.AddItem "All"
    CoType.AddItem "Select Supplier"
End Sub

Private Sub getTotals()
Dim r As Long
Dim dReceipt As Double, dPayment As Double
    r = 0
    dReceipt = 0
    dPayment = 0
    While r < MGrid.Rows
        dReceipt = dReceipt + Val(MGrid.TextMatrix(r, gDebit))
        dPayment = dPayment + Val(MGrid.TextMatrix(r, gCredit))
        r = r + 1
    Wend
    LReceipt.Caption = Format("" & dReceipt, "0.00")
    LPayment.Caption = Format("" & dPayment, "0.00")
    LBalance.Caption = Format("" & Abs((Val(dReceipt) - Val(dPayment))), "0.00") & IIf(Val(dReceipt) - Val(dPayment) > -1, " Dr", " Cr")
End Sub

