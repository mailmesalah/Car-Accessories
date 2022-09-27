VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FTaxReport 
   Caption         =   "Tax Report"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FTaxReport.frx":0000
   ScaleHeight     =   7020
   ScaleWidth      =   14505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CShow 
      Height          =   505
      Left            =   360
      Picture         =   "FTaxReport.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6375
      Width           =   1365
   End
   Begin VB.CommandButton CToExcel 
      Height          =   505
      Left            =   1815
      Picture         =   "FTaxReport.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6375
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   12960
      Picture         =   "FTaxReport.frx":203506
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
      Width           =   14280
      _ExtentX        =   25188
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
      Format          =   82575363
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
      Format          =   82575363
      CurrentDate     =   40458
   End
   Begin MSForms.Label Label13 
      Height          =   330
      Left            =   9165
      TabIndex        =   22
      Top             =   1260
      Width           =   1245
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Gross Amount"
      Size            =   "2196;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label12 
      Height          =   330
      Left            =   12690
      TabIndex        =   21
      Top             =   1260
      Width           =   1245
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "ToTal Amount"
      Size            =   "2196;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label11 
      Height          =   330
      Left            =   10260
      TabIndex        =   20
      Top             =   1260
      Width           =   1170
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Tax"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label10 
      Height          =   330
      Left            =   5190
      TabIndex        =   19
      Top             =   1260
      Width           =   1170
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Item"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   330
      Left            =   3765
      TabIndex        =   18
      Top             =   1260
      Width           =   1305
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Serial No"
      Size            =   "2302;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   330
      Left            =   1065
      TabIndex        =   17
      Top             =   1260
      Width           =   1185
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Bill No"
      Size            =   "2090;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape Shape1 
      Height          =   4725
      Left            =   75
      Top             =   1140
      Width           =   14310
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   4500
      TabIndex        =   15
      Top             =   75
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   120
      TabIndex        =   14
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
      TabIndex        =   13
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
      Left            =   150
      TabIndex        =   12
      Top             =   1260
      Width           =   870
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Date"
      Size            =   "1535;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   330
      Left            =   2355
      TabIndex        =   11
      Top             =   1260
      Width           =   1290
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Supplier"
      Size            =   "2275;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   8115
      TabIndex        =   10
      Top             =   1260
      Width           =   1020
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Quantity"
      Size            =   "1799;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   11265
      TabIndex        =   9
      Top             =   1260
      Width           =   1245
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Tax Amount"
      Size            =   "2196;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label LToTalAmount 
      Height          =   435
      Left            =   12930
      TabIndex        =   8
      Top             =   5865
      Width           =   1350
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2381;767"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   255
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label LTaxAmount 
      Height          =   435
      Left            =   11640
      TabIndex        =   7
      Top             =   5865
      Width           =   1215
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "2143;767"
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
      Left            =   6780
      TabIndex        =   6
      Top             =   1260
      Width           =   1170
      ForeColor       =   -2147483643
      VariousPropertyBits=   8388627
      Caption         =   "Rate"
      Size            =   "2064;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   495
      Index           =   120
      Left            =   -390
      TabIndex        =   16
      Top             =   1140
      Width           =   15240
      BackColor       =   15724527
      Size            =   "26882;873"
      Picture         =   "FTaxReport.frx":205968
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FTaxReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gDate As Single, gBillNo As Single, gSupplier As Single, gSerialNo As Single, gItemName As Single, gRate As Single, gQuantity As Single, gGrossAmount As Single, gTax As Single, gTaxAmount As Single, gToTalAmount As Single

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gDate = 0
    gBillNo = 1
    gSupplier = 2
    gSerialNo = 3
    gItemName = 4
    gRate = 5
    gQuantity = 6
    gGrossAmount = 7
    gTax = 8
    gTaxAmount = 9
    gToTalAmount = 10
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 11
    MGrid.Rows = 20
    MGrid.ColWidth(gDate) = 1200
    MGrid.ColWidth(gBillNo) = 800
    MGrid.ColWidth(gSupplier) = 1900
    MGrid.ColWidth(gSerialNo) = 800
    MGrid.ColWidth(gItemName) = 1900
    MGrid.ColWidth(gRate) = 1400
    MGrid.ColWidth(gQuantity) = 1000
    MGrid.ColWidth(gGrossAmount) = 1400
    MGrid.ColWidth(gTax) = 800
    MGrid.ColWidth(gTaxAmount) = 1400
    MGrid.ColWidth(gToTalAmount) = 1400
    MGrid.RowHeightMin = 350
End Sub

Private Sub CShow_Click()
Dim rs As Recordset
Dim dTax As Double, dTotalAmount As Double, dTaxAmount As Double, dRate As Double, dQuantity As Double

    MGrid.Rows = 0
    Set rs = db.OpenRecordset("Select Transaction.*,ItemMaster.ItemName,SupplierMaster.SupplierName As Supplier From Transaction,Itemmaster,SupplierMaster Where (Transaction.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "')) And (Transaction.TransactionType='P') And (SupplierMaster.SupplierCode=Transaction.SupplierCode) And (ItemMaster.Code=Transaction.SupplierCode) And (Transaction.Tax > 0) Order By Transaction.TransactionDate, Val(Transaction.TransactionNo)")
    While rs.EOF = False
        dTax = Val("" & rs!Tax)
        dQuantity = Val("" & rs!Quantity)
        dRate = Val("" & rs!PurchaseRate)
        dTaxAmount = Val((dRate * dQuantity) * (dTax / 100))
        dTotalAmount = Val((dRate * dQuantity)) + dTaxAmount
        MGrid.AddItem Format("" & rs!TransactionDate, "dd-MM-yyyy") & vbTab & "" & rs!TransactionNo & vbTab & "" & rs!Supplier & vbTab & Val("" & rs!SerialNo) & vbTab & "" & rs!ItemName & vbTab & Format(dRate, "0.00") & vbTab & Format(dQuantity, "0") & vbTab & Format(dQuantity * dRate, "0.00") & vbTab & Format(dTax, "0.00") & vbTab & Format(dTaxAmount, "0.00") & vbTab & Format(dTotalAmount, "0.00")
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
    xData(1, 3) = "Supplier"
    xData(1, 4) = "Serial No"
    xData(1, 5) = "Item Name"
    xData(1, 6) = "Rate"
    xData(1, 7) = "Quantity"
    xData(1, 8) = "Gross Amount"
    xData(1, 9) = "Tax"
    xData(1, 10) = "Tax Amount"
    xData(1, 11) = "Total Amount"
    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    xData(i + 1, 10) = Format(Val(LTaxAmount.Caption), "0.00")
    xData(i + 1, 11) = Format(Val(LToTalAmount.Caption), "0.00")
    
    oExcelSheet.Range("A3:K" & lRowCount + 4).Value = xData
    oExcelSheet.Cells(1, 1).Value = "Tax Report " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")
    oExcelSheet.Range("A1:K" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\" & "Tax Report" & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\" & "Tax Report" & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\" & "Tax Report" & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\" & "Tax Report" & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CShow_Click
    ElseIf (KeyCode = vbKeyE And ((Shift And 7) = 2)) Then
        CToExcel_Click
    ElseIf (KeyCode = vbKeyL And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    MGridInitialise
    DTPFrom.Value = Date
    DTPTo.Value = Date
End Sub

Private Sub getTotals()
Dim r As Long
Dim dTaxAmount As Double, dTotalAmount As Double
    r = 0
    dTaxAmount = 0
    dTotalAmount = 0
    While r < MGrid.Rows
        dTaxAmount = dTaxAmount + Val(MGrid.TextMatrix(r, gTaxAmount))
        dTotalAmount = dTotalAmount + Val(MGrid.TextMatrix(r, gToTalAmount))
        r = r + 1
    Wend
    LTaxAmount.Caption = Format(dTaxAmount, "0.00")
    LToTalAmount.Caption = Format(dTotalAmount, "0.00")
End Sub


