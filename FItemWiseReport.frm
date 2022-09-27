VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FItemWiseReport 
   Caption         =   "ItemWiseReport"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FItemWiseReport.frx":0000
   ScaleHeight     =   7020
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CShow 
      Height          =   505
      Left            =   360
      Picture         =   "FItemWiseReport.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6375
      Width           =   1365
   End
   Begin VB.CommandButton CToExcel 
      Height          =   505
      Left            =   1815
      Picture         =   "FItemWiseReport.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6375
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   8790
      Picture         =   "FItemWiseReport.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6375
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   90
      TabIndex        =   0
      Top             =   1650
      Width           =   10500
      _ExtentX        =   18521
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
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   345
      Left            =   885
      TabIndex        =   4
      Top             =   210
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
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
      Format          =   45678595
      CurrentDate     =   40458
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   885
      TabIndex        =   5
      Top             =   645
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
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
      Format          =   45678595
      CurrentDate     =   40458
   End
   Begin VB.Shape Shape1 
      Height          =   4725
      Left            =   75
      Top             =   1140
      Width           =   10545
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   4980
      TabIndex        =   17
      Top             =   180
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   120
      TabIndex        =   16
      Top             =   630
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "1905;582"
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   30
      TabIndex        =   14
      Top             =   1245
      Width           =   1410
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Date"
      Size            =   "2487;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   330
      Left            =   1665
      TabIndex        =   13
      Top             =   1245
      Width           =   3285
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Item Name"
      Size            =   "5794;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   7455
      TabIndex        =   12
      Top             =   1245
      Width           =   1170
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Qty"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   8775
      TabIndex        =   11
      Top             =   1245
      Width           =   1455
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Bill Amount"
      Size            =   "2566;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LBillAmount 
      Height          =   435
      Left            =   9075
      TabIndex        =   10
      Top             =   5865
      Width           =   1620
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2857;767"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LQuantity 
      Height          =   435
      Left            =   7560
      TabIndex        =   9
      Top             =   5865
      Width           =   1500
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2646;767"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   6000
      TabIndex        =   8
      Top             =   1245
      Width           =   1410
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Sale Rate"
      Size            =   "2487;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoItem 
      Height          =   420
      Left            =   7800
      TabIndex        =   7
      Top             =   210
      Width           =   2775
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "4895;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Index           =   0
      Left            =   6840
      TabIndex        =   6
      Top             =   210
      Width           =   840
      VariousPropertyBits=   8388627
      Caption         =   "Item"
      Size            =   "1482;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Line Line2 
      X1              =   75
      X2              =   10590
      Y1              =   1605
      Y2              =   1605
   End
   Begin MSForms.Label Label2 
      Height          =   495
      Index           =   120
      Left            =   75
      TabIndex        =   18
      Top             =   1140
      Width           =   10545
      BackColor       =   15724527
      Size            =   "18600;873"
      Picture         =   "FItemWiseReport.frx":205968
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FItemWiseReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sItemCode() As String
Dim gDate As Single, gItemName As Single, gSaleRate As Single, gQuantity As Single, gBillAmount As Single
Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gDate = 0
    gItemName = 1
    gSaleRate = 2
    gQuantity = 3
    gBillAmount = 4

    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 5
    MGrid.Rows = 0
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gItemName) = 4500
    MGrid.ColWidth(gSaleRate) = 1400
    MGrid.ColWidth(gQuantity) = 1400
    MGrid.ColWidth(gBillAmount) = 1500
    MGrid.RowHeightMin = 350
End Sub

Private Sub CoItem_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        CShow.SetFocus
    End If
End Sub
Private Sub CShow_Click()
Dim rs As Recordset

    If CoItem.ListIndex = -1 Then
        MsgBox "Please Select an Item !", vbInformation
        CoItem.SetFocus
    Exit Sub
    End If
    
    MGrid.Rows = 0
    Set rs = db.OpenRecordset(" Select ItemMaster.ItemName,Transaction.TransactionDate,Sum(Transaction.Quantity) As Qty,Transaction.SaleRate,Sum(Transaction.Quantity*Transaction.SaleRate) As BillAmount From ItemMaster,Transaction Where (ItemMaster.Code = Transaction.ItemCode) And (Transaction.ItemCode = '" & sItemCode(CoItem.ListIndex + 1) & "') And (Transaction.TransactionType = 'S' ) And (Transaction.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Group By ItemMaster.ItemName,Transaction.TransactionDate,Transaction.SaleRate Order By Transaction.TransactionDate")
    While rs.EOF = False
        MGrid.AddItem Format("" & rs!TransactionDate, "dd-MM-yyyy") & vbTab & "" & rs!ItemName & vbTab & Format(Val("" & rs!SaleRate), "0.00") & vbTab & Format(Val("" & rs!Qty), "0") & vbTab & Format(Val("" & rs!BillAmount), "0.00")
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
    xData(1, 2) = "Item Name"
    xData(1, 3) = "Sale Rate"
    xData(1, 4) = "Quantity"
    xData(1, 5) = "Bill Amount"
    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    xData(i + 1, 4) = LQuantity.Caption
    xData(i + 1, 5) = LBillAmount.Caption
    
    oExcelSheet.Range("A3:E" & lRowCount + 4).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Item - Wise Sale Register From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:E" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\" & "Item-Wise Sale Register" & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\" & "Item-Wise Sale Register" & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\" & "Item-Wise Sale Register" & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\" & "Item-Wise Sale Register" & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
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
Private Sub DTPFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DTPTo.SetFocus
End If
End Sub
Private Sub DTPTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CoItem.SetFocus
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CShow_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CToExcel_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub
Private Sub Form_Load()
    MGridInitialise
    DTPFrom.Value = Date
    DTPTo.Value = Date
    getItem
End Sub
Private Sub getItem()
Dim rs As Recordset
    
    CoItem.Clear
    
    Set rs = db.OpenRecordset("Select ItemMaster.Code,ItemMaster.ItemName From ItemMaster Where ItemMaster.Type='BItem' Order By ItemMaster.ItemName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sItemCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoItem.AddItem "" & rs!ItemName
        sItemCode(CoItem.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub getTotals()
Dim r As Long
Dim dQuantity As Double, dBillAmount As Double
    r = 0
    dBillAmount = 0
    While r < MGrid.Rows
        dQuantity = dQuantity + Val(MGrid.TextMatrix(r, gQuantity))
        dBillAmount = dBillAmount + Val(MGrid.TextMatrix(r, gBillAmount))
        r = r + 1
    Wend
    LQuantity.Caption = Format("" & dQuantity, "0")
    LBillAmount.Caption = Format("" & dBillAmount, "0.00")
End Sub


