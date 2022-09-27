VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FWarrantyDetails 
   Caption         =   "Warranty Details"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   Picture         =   "FWarrantyDetails.frx":0000
   ScaleHeight     =   6885
   ScaleWidth      =   13185
   StartUpPosition =   2  'CenterScreen
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
      Left            =   11355
      Picture         =   "FWarrantyDetails.frx":1FEC42
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6255
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
      Left            =   1875
      Picture         =   "FWarrantyDetails.frx":2010A4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6255
      Width           =   1365
   End
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
      Left            =   195
      Picture         =   "FWarrantyDetails.frx":203506
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6255
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   75
      TabIndex        =   6
      Top             =   1560
      Width           =   12975
      _ExtentX        =   22886
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
      Left            =   1380
      TabIndex        =   0
      Top             =   135
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
      Format          =   20643843
      CurrentDate     =   40909
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   1380
      TabIndex        =   1
      Top             =   570
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
      Format          =   20643843
      CurrentDate     =   40909
   End
   Begin VB.OLE OLEExcel 
      Height          =   975
      Left            =   5085
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label LBillAmount 
      Height          =   495
      Left            =   10905
      TabIndex        =   14
      Top             =   5850
      Width           =   2310
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "4075;873"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   330
      Left            =   360
      TabIndex        =   13
      Top             =   585
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   360
      TabIndex        =   12
      Top             =   195
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "From"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label15 
      Height          =   330
      Left            =   2310
      TabIndex        =   11
      Top             =   1200
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
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1125
      TabIndex        =   10
      Top             =   1200
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
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   11355
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Warranty"
      Size            =   "2566;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   420
      Left            =   3990
      TabIndex        =   8
      Top             =   1200
      Width           =   4245
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Customer - Description"
      Size            =   "7488;741"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox CoItem 
      Height          =   435
      Left            =   9915
      TabIndex        =   2
      Top             =   195
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;767"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   180
      TabIndex        =   7
      Top             =   1200
      Width           =   1065
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Serial No"
      Size            =   "1879;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   4665
      Left            =   60
      Top             =   1080
      Width           =   13005
   End
   Begin MSForms.Label gfdgfd 
      Height          =   495
      Index           =   120
      Left            =   75
      TabIndex        =   16
      Top             =   1080
      Width           =   12975
      BackColor       =   0
      Size            =   "22886;873"
      Picture         =   "FWarrantyDetails.frx":205968
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FWarrantyDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sItemCode() As String
Dim gSerialNo As Single, gDate As Single, gBillNo As Single, gDescription As Single, gStatus As Single

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
    gStatus = 4
    
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
    MGrid.ColWidth(gDescription) = 6575
    MGrid.ColWidth(gStatus) = 2500
    
    MGrid.RowHeightMin = 350
End Sub

Private Sub CShow_Click()
Dim rs As Recordset, dDate As Date, sStatus As String
    MGrid.Rows = 0
    If CoItem.ListIndex = -1 Then
        MsgBox "Select a Sale Type!", vbInformation
        CoItem.SetFocus
        Exit Sub
    Else
        Set rs = db.OpenRecordset("Select Transaction.TransactionNo,Transaction.TransactionDate,Transaction.Narration,Transaction.CustomerName,Transaction.Warranty From Transaction Where (Transaction.TransactionType = 'S' ) And (Transaction.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "') ) Group By Transaction.TransactionNo,Transaction.Warranty,Transaction.TransactionDate,Transaction.Narration,Transaction.CustomerName Order By Transaction.TransactionDate,Transaction.TransactionNo")
    End If
        
    While rs.EOF = False
     
        dDate = DateValue("" & rs!TransactionDate)
        If DateDiff("M", Date, (dDate + (("" & rs!Warranty) * 30))) = 0 Then
            sStatus = "Expires This Month"
        ElseIf DateDiff("M", Date, (dDate + (("" & rs!Warranty) * 30))) < 0 Then
            sStatus = "Already Expired"
        Else
            sStatus = "Expires in " & DateDiff("M", (dDate + (("" & rs!Warranty) * 30)), Date) & " Months"
        End If
        
        MGrid.AddItem MGrid.Rows + 1 & vbTab & Format("" & rs!TransactionDate, "dd-MM-yyyy") & vbTab & "" & rs!TransactionNo & vbTab & "" & rs!CustomerName & " -" & rs!Narration & vbTab & sStatus
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

    xData(1, 1) = "Sl No"
    xData(1, 2) = "Date"
    xData(1, 3) = "Bill No"
    xData(1, 4) = "Customer-Description"
    xData(1, 5) = "Warranty"
    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    xData(i + 1, 5) = LBillAmount.Caption
    
    oExcelSheet.Range("A3:E" & lRowCount + 4).Value = xData
    
    oExcelSheet.Cells(1, 1).Value = CoItem.Text & " From " & Format(DTPFrom.Value, "dd-MM-yyyy") & " To " & Format(DTPTo.Value, "dd-MM-yyyy")

    oExcelSheet.Range("A1:E" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\" & CoItem.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\" & CoItem.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\" & CoItem.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\" & CoItem.Text & Format(Date, "dd-MMM-yyyy") & ".xlsx"
    
    MsgBox "Successfully Exported !", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Private Sub DTPFrom_Change()
    MGrid.Rows = 0

End Sub

Private Sub DTPFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    DTPTo.SetFocus
End If
End Sub

Private Sub DTPTo_Change()
    MGrid.Rows = 0
End Sub
Private Sub DTPTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CoItem.SetFocus
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
    getItem
End Sub

Private Sub getItem()
Dim rs As Recordset
    
    CoItem.Clear
    
    Set rs = db.OpenRecordset("Select ItemMaster.Code,ItemMaster.ItemName From ItemMaster Where (ItemMaster.Type = 'BItem' ) Order By ItemMaster.ItemName")
    
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


