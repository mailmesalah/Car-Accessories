VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FBillWiseAccountsRegister1 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill - Wise Accounts Register"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FBillWiseAccountsRegister1.frx":0000
   ScaleHeight     =   8445
   ScaleWidth      =   12510
   StartUpPosition =   1  'CenterOwner
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
      Height          =   435
      Left            =   375
      Picture         =   "FBillWiseAccountsRegister1.frx":225B92
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7590
      Width           =   1485
   End
   Begin VB.CommandButton CPrint 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2070
      Picture         =   "FBillWiseAccountsRegister1.frx":22B190
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7590
      Width           =   1485
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
      Height          =   435
      Left            =   8415
      Picture         =   "FBillWiseAccountsRegister1.frx":23078E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7590
      Width           =   1485
   End
   Begin VB.CommandButton CClose 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10110
      Picture         =   "FBillWiseAccountsRegister1.frx":235D8C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7590
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   225
      TabIndex        =   6
      Top             =   1830
      Width           =   12030
      _ExtentX        =   21220
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
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   465
      Index           =   2
      Left            =   10080
      TabIndex        =   23
      Top             =   7575
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   465
      Index           =   1
      Left            =   8400
      TabIndex        =   22
      Top             =   7575
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   465
      Index           =   0
      Left            =   360
      TabIndex        =   21
      Top             =   7575
      Width           =   1530
   End
   Begin VB.Label LBack 
      BackColor       =   &H80000008&
      Height          =   465
      Index           =   5
      Left            =   2040
      TabIndex        =   20
      Top             =   7575
      Width           =   1530
   End
   Begin MSForms.Label LBalance 
      Height          =   405
      Left            =   10830
      TabIndex        =   18
      Top             =   7125
      Width           =   1770
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
      Left            =   10830
      TabIndex        =   17
      Top             =   6165
      Width           =   1755
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
      Left            =   10560
      TabIndex        =   16
      Top             =   1485
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
      Left            =   2955
      TabIndex        =   15
      Top             =   1500
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
      Left            =   6570
      TabIndex        =   14
      Top             =   1485
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
      Left            =   165
      TabIndex        =   13
      Top             =   1485
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
      Left            =   1350
      TabIndex        =   12
      Top             =   1500
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
      Left            =   9045
      TabIndex        =   11
      Top             =   1485
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
      Left            =   9135
      TabIndex        =   10
      Top             =   6180
      Width           =   1665
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
      Top             =   30
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.ComboBox CoSaleType 
      Height          =   405
      Left            =   8715
      TabIndex        =   0
      Top             =   165
      Width           =   3240
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5715;706"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   420
      Left            =   7800
      TabIndex        =   8
      Top             =   240
      Width           =   705
      VariousPropertyBits=   8388627
      Caption         =   "Type"
      Size            =   "1244;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBillNo 
      Height          =   405
      Left            =   8715
      TabIndex        =   1
      Top             =   645
      Width           =   3240
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5715;706"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   405
      Left            =   7800
      TabIndex        =   7
      Top             =   645
      Width           =   855
      VariousPropertyBits=   8388627
      Caption         =   "Bill No"
      Size            =   "1508;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label12 
      BackColor       =   &H00404040&
      Height          =   465
      Index           =   9
      Left            =   225
      TabIndex        =   19
      Top             =   1395
      Width           =   12030
   End
End
Attribute VB_Name = "FBillWiseAccountsRegister1"
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

Private Sub CoSaleType_Change()

End Sub

Private Sub CPrint_Click()
    If MGrid.Rows = 0 Then
        MsgBox "Empty Grid !", vbInformation
        Exit Sub
    End If
    'printReport
End Sub

Private Sub CShow_Click()
Dim rs As Recordset
    MGrid.Rows = 0
    If (CoSaleType.ListIndex = 0) Then
        If (Val("" & TBillNo.Text) = 0) Then
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount = 'LocalWholeSales' ) Order By Val(AccountRegister.BillNo) Desc,AccountRegister.TransactionDate")
        Else
            TBillNo.Text = Trim(TBillNo.Text)
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount = 'LocalWholeSales' ) And (AccountRegister.BillNo='" & TBillNo.Text & "') Order By Val(AccountRegister.BillNo) Desc,AccountRegister.TransactionDate")
        End If
    ElseIf (CoSaleType.ListIndex = 1) Then
        If (Val("" & TBillNo.Text) = 0) Then
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount = 'Purchase') Order By Val(AccountRegister.BillNo) Desc,AccountRegister.TransactionDate")
        Else
            TBillNo.Text = Trim(TBillNo.Text)
            Set rs = db.OpenRecordset("Select AccountMaster.AccountName,AccountRegister.Type,AccountRegister.TransactionDate,AccountRegister.Narration,AccountRegister.Income,AccountRegister.Expense,AccountRegister.BillNo From AccountMaster,AccountRegister Where (AccountMaster.Code = AccountRegister.AccountCode ) And (AccountRegister.SpecialAccount = 'Purchase') And (AccountRegister.BillNo='" & TBillNo.Text & "') Order By Val(AccountRegister.BillNo) Desc,AccountRegister.TransactionDate")
        End If
    Else
        Exit Sub
    End If
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
    xData(1, 3) = "Description"
    xData(1, 4) = "Bill Amount"
    xData(1, 5) = "Advance"
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
    ElseIf (KeyCode = vbKeyP And ((Shift And 7) = 2)) Then
        CPrint_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()

    MGridInitialise
    CoSaleType.AddItem "Local Sales(WholeSales)"
    CoSaleType.AddItem "Purchase"
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

