VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FMinimumStock 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Minimum Stock Comparison"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
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
   Icon            =   "FMinimumStock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FMinimumStock.frx":000C
   ScaleHeight     =   6390
   ScaleWidth      =   14010
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
      Left            =   255
      Picture         =   "FMinimumStock.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5775
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
      Left            =   1935
      Picture         =   "FMinimumStock.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5775
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
      Left            =   12135
      Picture         =   "FMinimumStock.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5775
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4125
      Left            =   135
      TabIndex        =   4
      Top             =   1140
      Width           =   13710
      _ExtentX        =   24183
      _ExtentY        =   7276
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   4635
      Left            =   120
      Top             =   640
      Width           =   13745
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   435
      Left            =   10755
      Top             =   135
      Width           =   3015
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Left            =   210
      TabIndex        =   12
      Top             =   735
      Width           =   975
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Serial No"
      Size            =   "1720;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label7 
      Height          =   330
      Left            =   10785
      TabIndex        =   11
      Top             =   735
      Width           =   1635
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Status"
      Size            =   "2884;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   8235
      TabIndex        =   10
      Top             =   750
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Balance Stock"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   6750
      TabIndex        =   9
      Top             =   750
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Current Stock"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   8880
      TabIndex        =   8
      Top             =   195
      Width           =   1620
      VariousPropertyBits=   8388627
      Caption         =   "Category"
      Size            =   "2857;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoCategory 
      Height          =   390
      Left            =   10770
      TabIndex        =   0
      Top             =   165
      Width           =   3015
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "5318;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1725
      TabIndex        =   7
      Top             =   735
      Width           =   2130
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Product"
      Size            =   "3757;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label15 
      Height          =   330
      Left            =   5250
      TabIndex        =   6
      Top             =   750
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Minimum Stock"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5025
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label Label3 
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   645
      Width           =   13740
      BackColor       =   15724527
      Size            =   "24236;873"
      Picture         =   "FMinimumStock.frx":205974
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FMinimumStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCategoryCode() As String
Dim gSerialNo As Single, gProduct As Single, gMinimumStock As Single, gCurrentStock As Single, gBalanceStock As Single, gStatus As Single

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gProduct = 1
    gMinimumStock = 2
    gCurrentStock = 3
    gBalanceStock = 4
    gStatus = 5
        
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 6
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 1000
    MGrid.ColWidth(gProduct) = 4200
    MGrid.ColWidth(gMinimumStock) = 1500
    MGrid.ColWidth(gCurrentStock) = 1500
    MGrid.ColWidth(gBalanceStock) = 1500
    MGrid.ColWidth(gStatus) = 3650
    MGrid.RowHeightMin = 350
End Sub

Private Sub getCatogaryData()
Dim rs As Recordset
    
    CoCategory.Clear
    Set rs = db.OpenRecordset("Select ItemMaster.Code,ItemMaster.ItemName From ItemMaster Where (ItemMaster.Type = 'AGroup' )")
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sCategoryCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoCategory.AddItem "" & rs!ItemName
        sCategoryCode(CoCategory.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub


Private Sub CShow_Click()
Dim rs As Recordset, dMinimumStock As Double, dCurrentStock As Double, dBalanceStock As Double, sStatus As String

    MGrid.Rows = 0
    If CoCategory.ListIndex >= 0 Then
        Set rs = db.OpenRecordset("Select ItemMaster.ItemName,(Select Sum(T.Quantity) From Transaction As T Where T.TransactionType In('O','P','SR','SA') And T.ItemCode=ItemMaster.Code) As StockIn, (Select Sum(T.Quantity) From Transaction As T Where T.TransactionType In('S','PR') And T.ItemCode=ItemMaster.Code) As StockOut, (Select MM.MinimumStock From ItemMaster As MM Where MM.Code=ItemMaster.Code)As MinimumStock From ItemMaster,Transaction Where (Transaction.ItemCode=ItemMaster.Code And ItemMaster.GroupCode='" & sCategoryCode(CoCategory.ListIndex + 1) & "') Group By ItemMaster.ItemName,ItemMaster.Code Order By ItemMaster.ItemName")
    Else
        MsgBox "Select a Category !", vbInformation
        CoCategory.SetFocus
        Exit Sub
    End If
        
    While rs.EOF = False
        dMinimumStock = Val("" & rs!MinimumStock)
        dCurrentStock = Val("" & rs!StockIn) - Val("" & rs!StockOut)
        dBalanceStock = dCurrentStock - dMinimumStock
        If dBalanceStock = 0 Then
            sStatus = "Now at Minimum Stock"
        ElseIf dBalanceStock < 0 Then
            sStatus = "Below Minimum Stock"
        Else
            sStatus = "Current Stock is above Minimum Stock"
        End If
        MGrid.AddItem MGrid.Rows + 1 & vbTab & "" & rs!ItemName & vbTab & dMinimumStock & vbTab & dCurrentStock & vbTab & dBalanceStock & vbTab & sStatus
        MGrid.Col = gStatus
        MGrid.Row = MGrid.Rows - 1
        MGrid.CellFontBold = True
        If dBalanceStock = 0 Then
            MGrid.CellForeColor = vbGreen
        ElseIf dBalanceStock < 0 Then
            MGrid.CellForeColor = vbRed
        Else
            MGrid.CellForeColor = vbBlue
        End If
        
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

    xData(1, 1) = "Product"
    xData(1, 2) = "Minimum Stock"
    xData(1, 3) = "Current Stock"
    xData(1, 4) = "Balance Stock"
    xData(1, 5) = "Status"
    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
        
    oExcelSheet.Range("A3:E" & lRowCount + 4).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Minimum Stock Comparison"

    oExcelSheet.Range("A1:E" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\Minimum Stock Comparison " & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\Minimum Stock Comparison " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\Minimum Stock Comparison " & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\Minimum Stock Comparison " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
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
    getCatogaryData
End Sub

