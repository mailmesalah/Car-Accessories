VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRateComparison 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Rate Comparison"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13110
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
   Icon            =   "FRateComparison.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FRateComparison.frx":000C
   ScaleHeight     =   6270
   ScaleWidth      =   13110
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
      Height          =   525
      Left            =   135
      Picture         =   "FRateComparison.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5565
      Width           =   1395
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
      Height          =   525
      Left            =   1815
      Picture         =   "FRateComparison.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5565
      Width           =   1395
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
      Height          =   525
      Left            =   11295
      Picture         =   "FRateComparison.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5565
      Width           =   1395
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   120
      TabIndex        =   2
      Top             =   1185
      Width           =   12870
      _ExtentX        =   22701
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
      Height          =   4650
      Left            =   105
      Top             =   720
      Width           =   12900
   End
   Begin MSForms.Label Label8 
      Height          =   330
      Left            =   105
      TabIndex        =   13
      Top             =   765
      Width           =   1155
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Serial No"
      Size            =   "2037;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   4560
      TabIndex        =   12
      Top             =   765
      Width           =   1155
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Bill No"
      Size            =   "2037;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   6435
      TabIndex        =   11
      Top             =   765
      Width           =   1155
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Narration"
      Size            =   "2037;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Index           =   0
      Left            =   8535
      TabIndex        =   10
      Top             =   765
      Width           =   1155
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "P.Rate"
      Size            =   "2037;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   330
      Left            =   11310
      TabIndex        =   9
      Top             =   765
      Width           =   1635
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Quantity"
      Size            =   "2884;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   9840
      TabIndex        =   8
      Top             =   765
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "MRP"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   1500
      TabIndex        =   7
      Top             =   765
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Supplier"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   8130
      TabIndex        =   6
      Top             =   120
      Width           =   1620
      VariousPropertyBits=   8388627
      Caption         =   "Item"
      Size            =   "2857;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoProduct 
      Height          =   405
      Left            =   9435
      TabIndex        =   0
      Top             =   90
      Width           =   3525
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "6218;714"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   4905
      TabIndex        =   5
      Top             =   -120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label9 
      Height          =   495
      Left            =   90
      TabIndex        =   14
      Top             =   715
      Width           =   12945
      BackColor       =   15724527
      Size            =   "22834;873"
      Picture         =   "FRateComparison.frx":205974
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FRateComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim sProductCode() As String
Dim gSerialNo As Single, gSupplier As Single, gBillNo As Single, gNarration As Single, gPurchaseRate As Single, gMRP As Single, gQuantity As Single

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gSupplier = 1
    gBillNo = 2
    gNarration = 3
    gPurchaseRate = 4
    gMRP = 5
    gQuantity = 6
        
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 7
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 1000
    MGrid.ColWidth(gSupplier) = 3000
    MGrid.ColWidth(gBillNo) = 1500
    MGrid.ColWidth(gNarration) = 2500
    MGrid.ColWidth(gPurchaseRate) = 1500
    MGrid.ColWidth(gMRP) = 1500
    MGrid.ColWidth(gQuantity) = 1500
    MGrid.RowHeightMin = 350
End Sub

Private Sub getCategoryData()
Dim rs As Recordset
    
    CoProduct.Clear
    Set rs = db.OpenRecordset("Select ItemMaster.Code As ProductCode,ItemMaster.ItemName As ProductName From ItemMaster Where (ItemMaster.Type = 'BItem' ) Order By ItemMaster.ItemName")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sProductCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoProduct.AddItem "" & rs!ProductName
        sProductCode(CoProduct.ListCount) = "" & rs!ProductCode
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CoProduct_Change()
    MGrid.Rows = 0
End Sub

Private Sub CShow_Click()
Dim rs As Recordset

    MGrid.Rows = 0
    If CoProduct.ListIndex >= 0 Then
        Set rs = db.OpenRecordset("Select Transaction.TransactionType,Transaction.TransactionNo,Transaction.SupplierName,Transaction.Narration,Transaction.PurchaseRate,Transaction.MRP,Transaction.UnitQuantity,Sum(Transaction.Quantity) As Quantity From Transaction Where (Transaction.ItemCode='" & sProductCode(CoProduct.ListIndex + 1) & "' And Transaction.TransactionType In ('P','O')) Group By Transaction.TransactionType,Transaction.TransactionNo,Transaction.SupplierName,Transaction.Narration,Transaction.PurchaseRate,Transaction.MRP,Transaction.UnitQuantity Order By Transaction.TransactionNo")
    Else
        MsgBox "Select a Product !", vbInformation
        CoProduct.SetFocus
        Exit Sub
    End If
        
    While rs.EOF = False
        MGrid.AddItem MGrid.Rows + 1 & vbTab & rs!SupplierName & vbTab & rs!TransactionType & "-" & rs!TransactionNo & vbTab & rs!Narration & vbTab & Format(rs!PurchaseRate / rs!UnitQuantity, "0.00") & vbTab & Format("" & rs!MRP, "0.00") & vbTab & rs!Quantity
        rs.MoveNext
    Wend
    rs.Close
        
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

    xData(1, 1) = "Sl.No"
    xData(1, 2) = "Supplier"
    xData(1, 3) = "Bill No"
    xData(1, 4) = "Narration"
    xData(1, 5) = "P.Rate"
    xData(1, 6) = "MRP"
    xData(1, 7) = "Quantity"

    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    oExcelSheet.Range("A3:G" & lRowCount + 4).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Rate Comparison Of : " & CoProduct.Text

    oExcelSheet.Range("A1:G" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\Rate Comparison Of " & CoProduct.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\Rate Comparison Of " & CoProduct.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\Rate Comparison Of " & CoProduct.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\Rate Comparison Of " & CoProduct.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
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
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()

    MGridInitialise
    getCategoryData
End Sub

