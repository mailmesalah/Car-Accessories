VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FPurchaseRegister 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Purchase Summary"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13275
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
   Icon            =   "FPurchaseRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FPurchaseRegister.frx":000C
   ScaleHeight     =   7005
   ScaleWidth      =   13275
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
      Height          =   505
      Left            =   375
      Picture         =   "FPurchaseRegister.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6375
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
      Height          =   505
      Left            =   2055
      Picture         =   "FPurchaseRegister.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6375
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
      Height          =   505
      Left            =   11540
      Picture         =   "FPurchaseRegister.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6375
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   13005
      _ExtentX        =   22939
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
      Left            =   1200
      TabIndex        =   0
      Top             =   150
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
      Format          =   20119555
      CurrentDate     =   40909
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   1200
      TabIndex        =   1
      Top             =   585
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
      Format          =   20119555
      CurrentDate     =   40909
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   4695
      Left            =   105
      Top             =   1080
      Width           =   13035
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   405
      Left            =   10050
      Top             =   225
      Width           =   3000
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   105
      TabIndex        =   15
      Top             =   1155
      Width           =   1245
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Serial No"
      Size            =   "2196;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoPurchase 
      Height          =   405
      Left            =   10065
      TabIndex        =   2
      Top             =   240
      Width           =   3000
      VariousPropertyBits=   746604571
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
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   4455
      TabIndex        =   14
      Top             =   1140
      Width           =   4245
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Supplier - Description"
      Size            =   "7488;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   11340
      TabIndex        =   13
      Top             =   1140
      Width           =   1455
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Bill Amount"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1170
      TabIndex        =   12
      Top             =   1140
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
      Left            =   2355
      TabIndex        =   11
      Top             =   1140
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
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   165
      TabIndex        =   10
      Top             =   210
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   165
      TabIndex        =   9
      Top             =   600
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LBillAmount 
      Height          =   495
      Left            =   11085
      TabIndex        =   8
      Top             =   5820
      Width           =   2055
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "3625;873"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5025
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label2 
      Height          =   495
      Index           =   120
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   13005
      BackColor       =   15724527
      Size            =   "22939;873"
      Picture         =   "FPurchaseRegister.frx":205974
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FPurchaseRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim gSerialNo As Single, gDate As Single, gBillNo As Single, gDescription As Single, gBillAmount As Single

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gDate = 1
    gBillNo = 2
    gDescription = 3
    gBillAmount = 4
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 5
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 1000
    MGrid.ColWidth(gDate) = 1400
    MGrid.ColWidth(gBillNo) = 1200
    MGrid.ColWidth(gDescription) = 7575
    MGrid.ColWidth(gBillAmount) = 1500
    
    MGrid.RowHeightMin = 350
End Sub

Private Sub CShow_Click()
Dim rs As Recordset
    MGrid.Rows = 0
    If CoPurchase.ListIndex = 0 Then
         Set rs = db.OpenRecordset("Select Transaction.TransactionNo,Transaction.TransactionDate,Transaction.Narration,Transaction.SupplierName,Sum((Transaction.PurchaseQuantity*Transaction.PurchaseRate)+((Transaction.PurchaseQuantity*Transaction.PurchaseRate)*Transaction.Tax/100)) As BillAmount From ItemMaster,Transaction Where (Transaction.TransactionType = 'P' ) And (Transaction.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') )And (ItemMaster.Code = Transaction.ItemCode ) Group By Transaction.TransactionNo,Transaction.TransactionDate,Transaction.Narration,Transaction.SupplierName Order By Transaction.TransactionDate,Transaction.TransactionNo")
       ' Set rs = db.OpenRecordset("Select Transaction.TransactionNo,Transaction.TransactionDate,Transaction.Narration,Transaction.SupplierName,Sum(Transaction.Quantity*Transaction.PurchaseRate) As BillAmount From Transaction Where (Transaction.TransactionType = 'P' ) And (Transaction.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Group By Transaction.TransactionNo,Transaction.TransactionDate,Transaction.Narration,Transaction.SupplierName Order By Transaction.TransactionDate,Transaction.TransactionNo")
    ElseIf CoPurchase.ListIndex = 1 Then
        Set rs = db.OpenRecordset("Select Transaction.TransactionNo,Transaction.TransactionDate,Transaction.Narration,Transaction.SupplierName,Sum(Transaction.Quantity*Transaction.PurchaseRate) As BillAmount From Transaction Where (Transaction.TransactionType = 'PR' ) And (Transaction.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Group By Transaction.TransactionNo,Transaction.TransactionDate,Transaction.Narration,Transaction.SupplierName Order By Transaction.TransactionDate,Transaction.TransactionNo")
    Else
        MsgBox "Select a Purchase Type!", vbInformation
        CoPurchase.SetFocus
        Exit Sub
    End If
        
    While rs.EOF = False
        MGrid.AddItem MGrid.Rows + 1 & vbTab & Format("" & rs!TransactionDate, "dd-MM-yyyy") & vbTab & "" & rs!TransactionNo & vbTab & "" & rs!SupplierName & " -" & rs!Narration & vbTab & Format(Val("" & rs!BillAmount), "0.00")
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
     
    xData(1, 1) = "Sl No"
    xData(1, 2) = "Date"
    xData(1, 3) = "Bill No"
    xData(1, 4) = "Supplier-Description"
    xData(1, 5) = "Bill Amount"
    
    For i = 2 To lRowCount + 1
        xData(i, 1) = MGrid.TextMatrix(i - 2, gSerialNo)
        xData(i, 2) = Format(MGrid.TextMatrix(i - 2, gDate), "dd/MMM/yyyy")
        xData(i, 3) = MGrid.TextMatrix(i - 2, gBillNo)
        xData(i, 4) = MGrid.TextMatrix(i - 2, gDescription)
        xData(i, 5) = MGrid.TextMatrix(i - 2, gBillAmount)
    Next i
    
    xData(i, 5) = LBillAmount.Caption
    
    oExcelSheet.Range("A3:E" & lRowCount + 4).Value = xData

    oExcelSheet.Cells(1, 1).Value = CoPurchase.Text & " From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:E" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\" & CoPurchase.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\" & CoPurchase.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\" & CoPurchase.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\" & CoPurchase.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Private Sub DTPFrom_Change()
    MGrid.Rows = 0
    getTotals
End Sub
Private Sub DTPFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DTPTo.SetFocus
End If
End Sub
Private Sub DTPTo_Change()
    MGrid.Rows = 0
    getTotals
End Sub
Private Sub DTPTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CoPurchase.SetFocus
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CShow_Click
    ElseIf (KeyCode = vbKeyE And ((Shift And 7) = 2)) Then
        CToExcel_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()

    MGridInitialise
    DTPFrom.Value = Date
    DTPTo.Value = Date
    CoPurchase.AddItem "Purchase"
    CoPurchase.AddItem "Purchase Return"
End Sub

Private Sub getTotals()
Dim r As Long
Dim dBillAmount As Double
    r = 0
    dBillAmount = 0
    While r < MGrid.Rows
        dBillAmount = dBillAmount + Val(MGrid.TextMatrix(r, gBillAmount))
        r = r + 1
    Wend
    LBillAmount.Caption = Format("" & dBillAmount, "0.00")
End Sub

