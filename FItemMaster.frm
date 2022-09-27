VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FItemMaster 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Master"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9840
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FItemMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FItemMaster.frx":000C
   ScaleHeight     =   6765
   ScaleWidth      =   9840
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CAddGroup 
      Height          =   505
      Left            =   285
      Picture         =   "FItemMaster.frx":1FEC4E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5565
      Width           =   1365
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Height          =   505
      Left            =   8355
      Picture         =   "FItemMaster.frx":2010B0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6165
      Width           =   1365
   End
   Begin VB.CommandButton CSave 
      Height          =   505
      Left            =   6840
      Picture         =   "FItemMaster.frx":203512
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6165
      Width           =   1365
   End
   Begin VB.CommandButton CDelete 
      Height          =   505
      Left            =   1755
      Picture         =   "FItemMaster.frx":205974
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6165
      Width           =   1365
   End
   Begin VB.CommandButton CAddNew 
      Height          =   505
      Left            =   285
      Picture         =   "FItemMaster.frx":207DD6
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6165
      Width           =   1365
   End
   Begin VB.CommandButton CFindNext 
      CausesValidation=   0   'False
      Height          =   505
      Left            =   2960
      Picture         =   "FItemMaster.frx":20A238
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5085
      Width           =   1365
   End
   Begin MSComctlLib.TreeView TrItems 
      Height          =   4785
      Left            =   255
      TabIndex        =   9
      Top             =   195
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   8440
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.TextBox TTax 
      Height          =   405
      Left            =   6615
      TabIndex        =   3
      Top             =   2422
      Width           =   3000
      VariousPropertyBits=   746604571
      MaxLength       =   40
      BorderStyle     =   1
      Size            =   "5292;714"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label9 
      Height          =   405
      Left            =   4800
      TabIndex        =   25
      Top             =   2430
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Tax"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoUnit 
      Height          =   405
      Left            =   6615
      TabIndex        =   5
      Top             =   3378
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;714"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Shape Shape2 
      Height          =   4815
      Left            =   225
      Top             =   180
      Width           =   4065
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   375
      Index           =   10
      Left            =   225
      Shape           =   4  'Rounded Rectangle
      Top             =   5085
      Width           =   2595
   End
   Begin MSForms.TextBox TRack 
      Height          =   405
      Left            =   6615
      TabIndex        =   4
      Top             =   2900
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5292;714"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   4800
      TabIndex        =   24
      Top             =   2925
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Rack"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoManufacturer 
      Height          =   405
      Left            =   6615
      TabIndex        =   2
      Top             =   1944
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;714"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   4800
      TabIndex        =   23
      Top             =   3405
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoItem 
      Height          =   405
      Left            =   6615
      TabIndex        =   0
      Top             =   988
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;714"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoStatus 
      Height          =   405
      Left            =   6615
      TabIndex        =   7
      Top             =   4335
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5292;714"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label8 
      Height          =   405
      Left            =   4800
      TabIndex        =   21
      Top             =   4380
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Status"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   405
      Left            =   4800
      TabIndex        =   20
      Top             =   3885
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Minimum Stock"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TMinimumStock 
      Height          =   405
      Left            =   6615
      TabIndex        =   6
      Top             =   3856
      Width           =   3000
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5292;714"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label6 
      Height          =   405
      Left            =   4800
      TabIndex        =   19
      Top             =   1950
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Manufacturer"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   405
      Left            =   4800
      TabIndex        =   18
      Top             =   1470
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Billing Name"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBillingName 
      Height          =   405
      Left            =   6615
      TabIndex        =   1
      Top             =   1466
      Width           =   3000
      VariousPropertyBits=   746604571
      MaxLength       =   40
      BorderStyle     =   1
      Size            =   "5292;714"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   405
      Left            =   4800
      TabIndex        =   17
      Top             =   975
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Name"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   405
      Left            =   4800
      TabIndex        =   16
      Top             =   495
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TItemCode 
      Height          =   405
      Left            =   6615
      TabIndex        =   10
      Top             =   510
      Width           =   3000
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "5292;706"
      BorderColor     =   8421504
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TFind 
      Height          =   315
      Left            =   270
      TabIndex        =   12
      Top             =   5100
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;556"
      SpecialEffect   =   0
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FItemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCreateNewGroup As Boolean
Dim sUnitCode() As String, sManufacturerCode() As String

Private Sub getItem()
Dim rs As Recordset
    
    CoItem.Clear
    
    Set rs = db.OpenRecordset("Select ItemMaster.ItemName From ItemMaster  Order By ItemMaster.ItemName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    While rs.EOF = False
        CoItem.AddItem "" & rs!ItemName
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getManufacturer()
Dim rs As Recordset
    
    CoManufacturer.Clear
    
    Set rs = db.OpenRecordset("Select Manufacturer.Code,Manufacturer.ManufacturerName From Manufacturer Order By Manufacturer.ManufacturerName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    ReDim sManufacturerCode(rs.RecordCount) As String
    While rs.EOF = False
        CoManufacturer.AddItem "" & rs!ManufacturerName
        sManufacturerCode(CoManufacturer.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub getUnit()
Dim rs As Recordset
    
    CoUnit.Clear
    
    Set rs = db.OpenRecordset("Select Units.Code,Units.UnitName From Units Order By Units.UnitName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    ReDim sUnitCode(rs.RecordCount) As String
    While rs.EOF = False
        CoUnit.AddItem "" & rs!UnitName
        sUnitCode(CoUnit.ListCount) = "" & rs!Code
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub CAddNew_Click()
    On Error Resume Next
    If (TrItems.Nodes.Count = 0) Then
        MsgBox "Please Create a a Group First !", vbInformation
        Exit Sub
    End If
    
    If Left(Trim(TrItems.SelectedItem.Key), 1) = "B" Then
        MsgBox "Please Select any Group to create Item !", vbInformation
        Exit Sub
    End If

    clearEditCodtrols
    enableDisableControlsOnAdd
    TItemCode = getNewItemcode
    bCreateNewGroup = False
    CoItem.SetFocus
End Sub

Private Sub CAddGroup_Click()
    clearEditCodtrols
    enableDisableControlsOnGroup
    TItemCode = getNewItemcode
    CoItem.SetFocus
    bCreateNewGroup = True
End Sub

Private Sub enableDisableControlsOnGroup()
    TItemCode.Enabled = False
    CoItem.Enabled = True
    TBillingName.Enabled = False
    TTax.Enabled = False
    TRack.Enabled = False
    CoManufacturer.Enabled = False
    CoUnit.Enabled = False
    TMinimumStock.Enabled = False
    CoStatus.Enabled = True
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDelete_Click()
Dim rs As Recordset
    
    If Trim(TItemCode.Text) = "" Then
        MsgBox "Please Select Any Item to Delete !", vbInformation
        Exit Sub
    End If
        
    If checkAlreadyUsed(Trim(TItemCode.Text)) Then
        MsgBox "The Item is Already Used !", vbInformation
        Exit Sub
    End If
    
    If checkForChildItems(Trim(TItemCode.Text)) Then
        MsgBox "The Group has Items, Please Delete them First !", vbInformation
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select ItemMaster.* From ItemMaster Where (ItemMaster.Code = '" & Trim(TItemCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        rs.Delete
        rs.Close
    Else
        rs.Close
        MsgBox "The Item doesnt Exist !", vbInformation
        Exit Sub
    End If
    
    MsgBox "Successfully Deleted the Item !", vbInformation
   
    refreshTree
    clearEditCodtrols
End Sub

Private Function checkForChildItems(sItemCode As String) As Boolean
Dim rs As Recordset, bFound As Boolean

    Set rs = db.OpenRecordset("Select ItemMaster.* From ItemMaster Where (ItemMaster.GroupCode = '" & Trim(sItemCode) & "' )")
    If rs.RecordCount > 0 Then
        bFound = True
    Else
        bFound = False
    End If
    rs.Close
    
    checkForChildItems = bFound
End Function

Private Sub CFindNext_Click()
Static lFindIndex As Long
Static sFindWord As String
    
    If Trim(TFind.Text) <> sFindWord Then
        lFindIndex = 1
    Else
        lFindIndex = lFindIndex + 1
    End If
    
    sFindWord = Trim(TFind.Text)
    
    Do While lFindIndex <= TrItems.Nodes.Count
        
        If InStr(1, LCase(TrItems.Nodes.Item(lFindIndex)), LCase(sFindWord), vbTextCompare) > 0 Then
            TrItems.Nodes.Item(lFindIndex).Selected = True
            getDetailsOfItem
            TrItems.SetFocus
            Exit Do
        End If
        lFindIndex = lFindIndex + 1
    Loop
    
    If lFindIndex > TrItems.Nodes.Count Then
        MsgBox "No more Items !", vbInformation
        lFindIndex = 1
        Exit Sub
    End If
End Sub

Private Sub CoManufacturer_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FManufacturer.Show vbModal
        getManufacturer
    End If
End Sub
Private Sub CoStatus_GotFocus()
    CoStatus.SelStart = 0
    CoStatus.SelLength = Len(CoStatus.Text)
End Sub

Private Sub CoUnit_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 113 Then
        FUnits.Show vbModal
        getUnit
    End If
End Sub

Private Sub CSave_Click()
Dim rs As Recordset, sStatus As String, sItemCode As String, sParenttype As String
Dim sParentCode As String

    If Trim(TItemCode.Text) = "" Then
        MsgBox "Please Select an Item to Edit or click Add New button To add new Item", vbInformation
        Exit Sub
    ElseIf Trim(CoItem.Text) = "" Then
        MsgBox "Please Enter needed Informations !", vbInformation
        CoItem.SetFocus
        Exit Sub
'    ElseIf Trim(TPartNo.Text) = "" And Not bCreateNewGroup Then
'        MsgBox "Please Enter needed Informations !", vbInformation
'        TPartNo.SetFocus
'        Exit Sub
    ElseIf Trim(TBillingName.Text) = "" And Not bCreateNewGroup Then
        MsgBox "Please Enter needed Informations !", vbInformation
        TBillingName.SetFocus
        Exit Sub
'    ElseIf Val(TTax.Text) = 0 And Not bCreateNewGroup Then
'        MsgBox "Please Enter needed Informations !", vbInformation
'        TTax.SetFocus
'        Exit Sub
    ElseIf CoManufacturer.ListIndex = -1 And Not bCreateNewGroup Then
        MsgBox "Please Enter needed Informations !", vbInformation
        CoManufacturer.SetFocus
        Exit Sub
    ElseIf CoUnit.ListIndex = -1 And Not bCreateNewGroup Then
        MsgBox "Please Enter needed Informations !", vbInformation
        CoUnit.SetFocus
        Exit Sub
    End If
    
    'Determines GroupCode
    If TrItems.Nodes.Count > 0 Then
        If (bCreateNewGroup) Then
            sParenttype = ""
            sParentCode = ""
        Else
            sParenttype = Trim(Left(TrItems.SelectedItem.Key, 1))
            sParentCode = Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1))
        End If
    Else
        sParenttype = ""
        sParentCode = ""
    End If

    Set rs = db.OpenRecordset("Select ItemMaster.* From ItemMaster Where (ItemMaster.Code = '" & Trim(TItemCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        sStatus = "Edited"
        rs.Edit
    Else
        sStatus = "Added"
        TItemCode.Text = getNewItemcode()
        rs.AddNew
        rs!Code = Trim(TItemCode.Text)
        rs!Type = IIf(Trim(sParenttype) = "", "AGroup", "BItem")
        rs!GroupCode = sParentCode
    End If
    rs!ItemName = Trim(CoItem.Text)
    rs!BillingName = Trim(TBillingName.Text)
    rs!ManufacturerCode = sManufacturerCode(CoManufacturer.ListIndex + 1)
    rs!Rack = Trim(TRack.Text)
    rs!Tax = Trim(TTax.Text)
    rs!UnitCode = sUnitCode(CoUnit.ListIndex + 1)
    rs!MinimumStock = Val(TMinimumStock.Text)
    rs!Status = IIf((CoStatus.ListIndex = 0), True, False)
    rs.Update
    rs.Close
    
    MsgBox "Successfully " & sStatus & " !", vbInformation
    getItem
    refreshTree
    clearEditCodtrols
    bCreateNewGroup = False
    getLastItem
End Sub

Private Sub getLastItem()
Dim rs As Recordset, sCode As String, i As Long
    Set rs = db.OpenRecordset("Select Max(Val(ItemMaster.Code)) As Code From ItemMaster")
    If rs.RecordCount > 0 Then
        sCode = "" & rs!Code
    Else
        sCode = "Nill"
    End If
    rs.Close
    
    i = 1
    Do While i <= TrItems.Nodes.Count
        If Right(TrItems.Nodes.Item(i).Key, Len(TrItems.Nodes.Item(i).Key) - 1) = sCode Then
            TrItems.Nodes.Item(i).Selected = True
            TrItems.SetFocus
            Exit Do
        End If
        i = i + 1
    Loop
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF And ((Shift And 7) = 2)) Then
        CFindNext_Click
    ElseIf (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CAddNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDelete_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyX And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    
    CoStatus.AddItem "Active"
    CoStatus.AddItem "Inactive"
    
    refreshTree
    enableDisableControls
    getItem
    getManufacturer
    getUnit
    bCreateNewGroup = False
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    
    TrItems.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select ItemMaster.Code,ItemMaster.ItemName,ItemMaster.Type,ItemMaster.GroupCode From ItemMaster Order By ItemMaster.Type,ItemMaster.ItemName")
    While rs.EOF = False
        If Trim(rs!Type) = "AGroup" Then
            TrItems.Nodes.Add , , "A" & rs!Code, rs!ItemName
            TrItems.Nodes(TrItems.Nodes.Count).Bold = True
            TrItems.Nodes(TrItems.Nodes.Count).ForeColor = &H808080
        ElseIf Trim(rs!Type) = "BItem" Then
            TrItems.Nodes.Add "A" & rs!GroupCode, tvwChild, "B" & rs!Code, rs!ItemName
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub TBillingName_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
If KeyCode = 13 Then
    If TBillingName = "" Then
        TBillingName = CoItem.Text
    End If
End If
End Sub
Private Sub TItemCode_GotFocus()
    TItemCode.SelStart = 0
    TItemCode.SelLength = Len(TItemCode.Text)
End Sub


Private Sub CoItem_GotFocus()
    CoItem.SelStart = 0
    CoItem.SelLength = Len(CoItem.Text)
End Sub

Private Sub TFind_GotFocus()
    TFind.SelStart = 0
    TFind.SelLength = Len(TFind.Text)
End Sub
Private Sub TTax_GotFocus()
    TTax.SelStart = 0
    TTax.SelLength = Len(TTax.Text)
End Sub

Private Sub TrItems_Click()
    enableDisableControls
    'getDetailsOfItem
End Sub

Private Sub TrItems_NodeClick(ByVal Node As MSComctlLib.Node)
    enableDisableControls
    If TrItems.Nodes.Count > 0 Then
        getDetailsOfItem
    End If
End Sub

Private Sub getDetailsOfItem()
Dim rs As Recordset
    
    If (Left(TrItems.SelectedItem.Key, 1) = "A") Then
        Set rs = db.OpenRecordset("Select '' As ManufacturerName,ItemMaster.Code,ItemMaster.ItemName,ItemMaster.BillingName,ItemMaster.PartNo,ItemMaster.ManufacturerCode,ItemMaster.Tax,ItemMaster.Rack,ItemMaster.MinimumStock,ItemMaster.Status,'' As UnitName From ItemMaster Where (ItemMaster.Code = '" & Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1)) & "' )")
    ElseIf (Left(TrItems.SelectedItem.Key, 1) = "B") Then
        Set rs = db.OpenRecordset("Select Manufacturer.ManufacturerName,ItemMaster.Code,ItemMaster.ItemName,ItemMaster.BillingName,ItemMaster.PartNo,ItemMaster.ManufacturerCode,ItemMaster.Tax,ItemMaster.Rack,ItemMaster.MinimumStock,ItemMaster.Status,Units.UnitName From Manufacturer,ItemMaster,Units Where (Manufacturer.Code = ItemMaster.ManufacturerCode ) And (Units.Code = ItemMaster.UnitCode ) And (ItemMaster.Code = '" & Trim(Right(TrItems.SelectedItem.Key, Len(TrItems.SelectedItem.Key) - 1)) & "' )")
    Else
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        
        TItemCode.Text = "" & rs!Code
        CoItem.Text = "" & rs!ItemName
        TBillingName.Text = "" & rs!BillingName
        CoManufacturer.Text = "" & rs!ManufacturerName
        TRack.Text = "" & rs!Rack
        TTax.Text = "" & rs!Tax
        CoUnit.Text = "" & rs!UnitName
        TMinimumStock.Text = Val("" & rs!MinimumStock)
        CoStatus.ListIndex = IIf((rs!Status = True), 0, 1)
    Else
        clearEditCodtrols
    End If
    rs.Close
End Sub

Private Sub enableDisableControlsOnAdd()
    
    If Left(TrItems.SelectedItem.Key, 1) = "A" Then
        
        TItemCode.Enabled = False
        CoItem.Enabled = True
        TBillingName.Enabled = True
        TTax.Enabled = True
        TRack.Enabled = True
        CoManufacturer.Enabled = True
        CoUnit.Enabled = True
        TMinimumStock.Enabled = True
        CoStatus.Enabled = True
    ElseIf Left(TrItems.SelectedItem.Key, 1) = "B" Then
        
        TItemCode.Enabled = False
        CoItem.Enabled = True
        TBillingName.Enabled = True
        TTax.Enabled = True
        TRack.Enabled = True
        CoManufacturer.Enabled = True
        CoUnit.Enabled = True
        TMinimumStock.Enabled = True
        CoStatus.Enabled = True
    End If
End Sub

Private Sub enableDisableControls()
    
    If TrItems.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Left(TrItems.SelectedItem.Key, 1) = "A" Then
        
        TItemCode.Enabled = False
        CoItem.Enabled = True
        TBillingName.Enabled = False
        TTax.Enabled = False
        CoManufacturer.Enabled = False
        CoUnit.Enabled = False
        TRack.Enabled = False
        TMinimumStock.Enabled = False
        CoStatus.Enabled = True
        bCreateNewGroup = True
    ElseIf Left(TrItems.SelectedItem.Key, 1) = "B" Then
    
        TItemCode.Enabled = False
        CoItem.Enabled = True
        TTax.Enabled = True
        TBillingName.Enabled = True
        CoManufacturer.Enabled = True
        TRack.Enabled = True
        CoUnit.Enabled = True
        TMinimumStock.Enabled = True
        CoStatus.Enabled = True
        bCreateNewGroup = False
    End If
End Sub

Private Sub clearEditCodtrols()
    TItemCode.Text = ""
    CoItem.Text = ""
    TBillingName.Text = ""
    CoManufacturer.Text = ""
    CoUnit.Text = ""
    TRack = ""
    TTax = ""
    TMinimumStock.Text = ""
    CoStatus.ListIndex = 0
End Sub

Private Function getParentItem(sItemCode As String) As String
Dim rs As Recordset, sParentCode As String
    
    Set rs = db.OpenRecordset("Select ItemMaster.GroupCode From ItemMaster Where (ItemMaster.Code = '" & Trim(sItemCode) & "' )")
    If rs.RecordCount > 0 Then
        sParentCode = "" & rs!GroupCode
    Else
        sParentCode = ""
    End If
    rs.Close
    
    getParentItem = sParentCode
End Function

Private Function getNewItemcode() As String
Dim rs As Recordset, sItemCode As String
    
    Set rs = db.OpenRecordset("Select Max(val(ItemMaster.Code))As ACode From ItemMaster")
    If rs.RecordCount > 0 Then
        sItemCode = Val("" & rs!ACode) + 1
    Else
        sItemCode = "1"
    
    End If
    rs.Close
    
    getNewItemcode = sItemCode
End Function

Private Function checkAlreadyUsed(sMCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.ItemCode = '" & sMCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    
    checkAlreadyUsed = bExist
End Function

