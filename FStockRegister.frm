VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FStockRegister 
   BackColor       =   &H00EFEFEF&
   Caption         =   "Stock Register"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14700
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
   Icon            =   "FStockRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FStockRegister.frx":000C
   ScaleHeight     =   7065
   ScaleWidth      =   14700
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
      Left            =   735
      Picture         =   "FStockRegister.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   4
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
      Left            =   2415
      Picture         =   "FStockRegister.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Left            =   12495
      Picture         =   "FStockRegister.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6375
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid MGrid 
      Height          =   4155
      Left            =   135
      TabIndex        =   5
      Top             =   1560
      Width           =   14460
      _ExtentX        =   25506
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
      Left            =   1290
      TabIndex        =   0
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
      Format          =   20643843
      CurrentDate     =   40909
   End
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   510
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   4605
      Left            =   120
      Top             =   1125
      Width           =   14490
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   400
      Left            =   11490
      Top             =   585
      Width           =   3000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   400
      Left            =   11490
      Top             =   150
      Width           =   3000
   End
   Begin MSForms.Label Label10 
      Height          =   330
      Left            =   60
      TabIndex        =   21
      Top             =   1185
      Width           =   1290
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Serial No"
      Size            =   "2275;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label9 
      Height          =   330
      Left            =   12660
      TabIndex        =   20
      Top             =   1170
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Stock Value"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label8 
      Height          =   330
      Left            =   11190
      TabIndex        =   19
      Top             =   1185
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
   Begin MSForms.Label Label7 
      Height          =   330
      Left            =   9765
      TabIndex        =   18
      Top             =   1185
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
   Begin MSForms.Label Label6 
      Height          =   330
      Left            =   8160
      TabIndex        =   17
      Top             =   1185
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Stock Out"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   6675
      TabIndex        =   16
      Top             =   1185
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Stock In"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label2 
      Height          =   330
      Index           =   0
      Left            =   9570
      TabIndex        =   15
      Top             =   600
      Width           =   1620
      VariousPropertyBits=   8388627
      Caption         =   "Filter By"
      Size            =   "2857;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   9570
      TabIndex        =   14
      Top             =   210
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
   Begin MSForms.ComboBox CoFilterBy 
      Height          =   400
      Left            =   11490
      TabIndex        =   3
      Top             =   585
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
   Begin MSForms.ComboBox CoCategory 
      Height          =   400
      Left            =   11490
      TabIndex        =   2
      Top             =   150
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
   Begin MSForms.Label Label14 
      Height          =   330
      Left            =   1440
      TabIndex        =   13
      Top             =   1185
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
      TabIndex        =   12
      Top             =   1185
      Width           =   1605
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "Opening Stock"
      Size            =   "2831;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   390
      TabIndex        =   11
      Top             =   165
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
      Left            =   390
      TabIndex        =   10
      Top             =   540
      Width           =   1080
      VariousPropertyBits=   8388627
      Caption         =   "To"
      Size            =   "1905;582"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label LAmount 
      Height          =   495
      Left            =   12555
      TabIndex        =   9
      Top             =   5895
      Width           =   1965
      VariousPropertyBits=   8388627
      Caption         =   "0.00"
      Size            =   "3466;873"
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
      Left            =   5145
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.Label gfdgfd 
      Height          =   495
      Index           =   120
      Left            =   120
      TabIndex        =   23
      Top             =   1155
      Width           =   12975
      BackColor       =   0
      VariousPropertyBits=   8388627
      Size            =   "22886;873"
      Picture         =   "FStockRegister.frx":205974
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label gfdgfd 
      Height          =   495
      Index           =   0
      Left            =   1650
      TabIndex        =   24
      Top             =   1125
      Width           =   12975
      BackColor       =   0
      Size            =   "22886;873"
      Picture         =   "FStockRegister.frx":21F936
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   465
      Index           =   120
      Left            =   105
      TabIndex        =   22
      Top             =   1140
      Width           =   14520
      BackColor       =   16711680
      Size            =   "25612;820"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "FStockRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim sFilterCode() As String
Dim gSerialNo As Single, gProduct As Single, gOpeningStock As Single, gStockIn As Single, gStockOut As Single, gClosingStock As Single, gMRP As Single, gStockValue As Single

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub MGridInitialise()
'INITIALISES MGRID
        'SETTING CONSTANTS
    gSerialNo = 0
    gProduct = 1
    gOpeningStock = 2
    gStockIn = 3
    gStockOut = 4
    gClosingStock = 5
    gMRP = 6
    gStockValue = 7
    
    MGrid.Clear
    MGrid.Rows = 1 'FOR SKIPING ERROR
    MGrid.Cols = 1 'FOR SKIPING ERROR
    MGrid.FixedCols = 0
    MGrid.FixedRows = 0
    MGrid.Cols = 8
    MGrid.Rows = 0
    MGrid.ColWidth(gSerialNo) = 1000
    MGrid.ColWidth(gProduct) = 4200
    MGrid.ColWidth(gOpeningStock) = 1500
    MGrid.ColWidth(gStockIn) = 1500
    MGrid.ColWidth(gStockOut) = 1500
    MGrid.ColWidth(gClosingStock) = 1500
    MGrid.ColWidth(gMRP) = 1500
    MGrid.ColWidth(gStockValue) = 1500
    MGrid.RowHeightMin = 350
End Sub

Private Sub getFilterData()
Dim rs As Recordset
    
    CoFilterBy.Clear
    If CoCategory.ListIndex = 0 Then 'Group
        Set rs = db.OpenRecordset("Select ItemMaster.Code As FilterCode,ItemMaster.ItemName As FilterName From ItemMaster Where (ItemMaster.Type = 'AGroup' ) Order By ItemMaster.ItemName")
    ElseIf CoCategory.ListIndex = 1 Then 'Manufacturer
        Set rs = db.OpenRecordset("Select Manufacturer.Code As FilterCode,Manufacturer.ManufacturerName As FilterName From Manufacturer Order By Manufacturer.ManufacturerName")
    'ElseIf CoCategory.ListIndex = 2 Then 'Supplier
    '    Set rs = db.OpenRecordset("Select SupplierMaster.SupplierCode As FilterCode,SupplierMaster.SupplierName As FilterName From SupplierMaster Order By SupplierMaster.SupplierCode")
    ElseIf CoCategory.ListIndex = 2 Then 'All Product
        Exit Sub
    ElseIf CoCategory.ListIndex = 3 Then 'Specific Product
        Set rs = db.OpenRecordset("Select ItemMaster.Code As FilterCode,ItemMaster.ItemName As FilterName From ItemMaster Where (ItemMaster.Type = 'BItem' ) Order By ItemMaster.ItemName")
    Else
        Exit Sub
    End If
    
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    ReDim sFilterCode(rs.RecordCount + 1) As String
    While rs.EOF = False
        CoFilterBy.AddItem "" & rs!FilterName
        sFilterCode(CoFilterBy.ListCount) = "" & rs!FilterCode
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub CoCategory_Change()
    getFilterData
End Sub

Private Sub CShow_Click()
Dim rs As Recordset
Dim dOpeningStock As Double, dStockIn As Double, dStockOut As Double, dClosingStock As Double, dMRP As Double
    MGrid.Rows = 0
    If CoCategory.ListIndex = 0 Then
        If CoFilterBy.ListIndex < 0 Then
            MsgBox "Please Select a Filter Value !", vbInformation
            Exit Sub
        End If
        Set rs = db.OpenRecordset("Select IM.ItemName," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningInStock," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningOutStock," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockIn," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockOut," _
        & " (Select Sum(T.MRP)/Sum(1) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('P','O')))As MeanMRP" _
        & " From ItemMaster As IM" _
        & " Where (IM.Code In (Select ItemMaster.Code From ItemMaster Where(ItemMaster.GroupCode='" & sFilterCode(CoFilterBy.ListIndex + 1) & "')) ) Order By IM.ItemName")
    ElseIf CoCategory.ListIndex = 1 Then
        If CoFilterBy.ListIndex < 0 Then
            MsgBox "Please Select a Filter Value !", vbInformation
            Exit Sub
        End If
        Set rs = db.OpenRecordset("Select IM.ItemName," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningInStock," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningOutStock," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockIn," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockOut," _
        & " (Select Sum(T.MRP)/Sum(1) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('P','O')))As MeanMRP" _
        & " From ItemMaster As IM" _
        & " Where (IM.Code In (Select ItemMaster.Code From ItemMaster Where(ItemMaster.ManufacturerCode='" & sFilterCode(CoFilterBy.ListIndex + 1) & "')) ) Order By IM.ItemName")
    'ElseIf CoCategory.ListIndex = 2 Then
    '    If CoFilterBy.ListIndex < 0 Then
    '        MsgBox "Please Select a Filter Value !", vbInformation
    '        Exit Sub
    '    End If
    '    Set rs = db.OpenRecordset("Select IM.ItemName," _
    '    & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningInStock," _
    '    & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningOutStock," _
    '    & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockIn," _
    '    & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockOut," _
    '    & " (Select Sum(T.MRP)/Sum(1) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('P','O')))As MeanMRP" _
    '    & " From ItemMaster As IM" _
    '    & " Where (IM.Code In (Select T.ItemCode From Transaction As T Where(T.SupplierCode='" & sFilterCode(CoFilterBy.ListIndex + 1) & "')Group By T.ItemCode)) Order By IM.ItemName")
    ElseIf CoCategory.ListIndex = 2 Then
        Set rs = db.OpenRecordset("Select IM.ItemName," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningInStock," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningOutStock," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockIn," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockOut," _
        & " (Select Sum(T.MRP)/Sum(1) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('P','O')))As MeanMRP" _
        & " From ItemMaster As IM" _
        & " Where (IM.Code In (Select ItemMaster.Code From ItemMaster Where(ItemMaster.Type='BItem')) ) Order By IM.ItemName")
    ElseIf CoCategory.ListIndex = 3 Then
        If CoFilterBy.ListIndex < 0 Then
            MsgBox "Please Select a Filter Value !", vbInformation
            Exit Sub
        End If
        Set rs = db.OpenRecordset("Select IM.ItemName," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningInStock," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate < cDate('" & DTPFrom.Value & "'))As OpeningOutStock," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('O','P','SR','SA')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockIn," _
        & " (Select Sum(T.Quantity) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('S','PR','F8','FB')) And T.TransactionDate Between cDate('" & DTPFrom.Value & "') And cDate('" & DTPTo.Value & "'))As StockOut," _
        & " (Select Sum(T.MRP)/Sum(1) From Transaction As T Where(T.ItemCode=IM.Code) And (T.TransactionType In ('P','O')))As MeanMRP" _
        & " From ItemMaster As IM" _
        & " Where (IM.Code='" & sFilterCode(CoFilterBy.ListIndex + 1) & "' ) Order By IM.ItemName")
    Else
        MsgBox "Select a Category!", vbInformation
        CoCategory.SetFocus
        Exit Sub
    End If
        
    While rs.EOF = False
        dOpeningStock = Val("" & rs!OpeningInStock) - Val("" & rs!OpeningOutStock)
        dStockIn = Val("" & rs!StockIn)
        dStockOut = Val("" & rs!StockOut)
        dClosingStock = (dOpeningStock + dStockIn) - dStockOut
        dMRP = Val("" & rs!MeanMRP)
        MGrid.AddItem MGrid.Rows + 1 & vbTab & "" & rs!ItemName & vbTab & dOpeningStock & vbTab & dStockIn & vbTab & dStockOut & vbTab & dClosingStock & vbTab & Format(dMRP, "0.00") & vbTab & Format(dClosingStock * dMRP, "0.00")
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
    
'    xData(1, 1) = "Sl.No"
'    xData(1, 2) = "Product"
'    xData(1, 3) = "Opening Stock"
'    xData(1, 4) = "Stock In"
'    xData(1, 5) = "Stock Out"
'    xData(1, 6) = "Closing Stock"
'    xData(1, 7) = "MRP"
'    xData(1, 8) = "Stock Value"
    
    xData(1, 1) = "Sl.No"
    xData(1, 2) = "Product"
    xData(1, 3) = "Opening Stock"
    xData(1, 4) = "Stock In"
    xData(1, 5) = "Stock Out"
    xData(1, 6) = "Closing Stock"
    xData(1, 7) = "MRP"
    xData(1, 8) = "Stock Value"

    
    For i = 1 To lRowCount
       For j = 1 To lColCount
          xData(i + 1, j) = MGrid.TextMatrix(i - 1, j - 1)
       Next j
    Next i
    
    xData(i + 1, 8) = LAmount.Caption
    
    oExcelSheet.Range("A3:H" & lRowCount + 4).Value = xData

    oExcelSheet.Cells(1, 1).Value = "Stock Register From " & Format(DTPFrom.Value, "dd-IM-yyyy") & " To " & Format(DTPTo.Value, "dd-IM-yyyy")

    oExcelSheet.Range("A1:H" & lRowCount + 4).Select
    oExcel.Application.Selection.AutoFormat
On Error Resume Next

    Kill App.Path & "\Reports\StockRegister " & Format(Date, "dd-IMM-yyyy") & ".xlsx"

    oExcel.SaveAs App.Path & "\Reports\StockRegister " & Format(Date, "dd-IMM-yyyy") & ".xlsx"
    Set oExcel = Nothing
    Set oExcelSheet = Nothing
    
    lReturnValue = Shell(App.Path & "\EXCEL.exe - """ & App.Path & "\Reports\StockRegister " & Format(Date, "dd-IMM-yyyy") & ".xlsx""", vbNormalFocus)

    OLEExcel.Close
    OLEExcel.Delete
    
    Dim xlTmp As Excel.Application
    Set xlTmp = New Excel.Application
    xlTmp.DisplayFullScreen = True
    xlTmp.Visible = True
    xlTmp.Workbooks.Open App.Path & "\Reports\StockRegister " & Format(Date, "dd-IMM-yyyy") & ".xlsx"
    
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
    CoCategory.SetFocus
End If
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
    DTPFrom.Value = Date
    DTPTo.Value = Date
    CoCategory.AddItem "Group"
    CoCategory.AddItem "Manufacturer"
    'CoCategory.AddItem "Supplier"
    CoCategory.AddItem "All Items"
    CoCategory.AddItem "Specific Item"
End Sub

Private Sub getTotals()
Dim r As Long
Dim dBillAmount As Double
    r = 0
    dBillAmount = 0
    While r < MGrid.Rows
        dBillAmount = dBillAmount + Val(MGrid.TextMatrix(r, gStockValue))
        r = r + 1
    Wend
    LAmount.Caption = Format("" & dBillAmount, "0.00")
End Sub

