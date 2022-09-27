VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FMedicineMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medicine Master"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
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
   Icon            =   "FMedicineMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CAddGroup 
      Caption         =   "Add Group"
      Height          =   540
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5880
      Width           =   1890
   End
   Begin VB.CommandButton CClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   540
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6510
      Width           =   1890
   End
   Begin VB.CommandButton CSave 
      Caption         =   "Save"
      Height          =   540
      Left            =   6015
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6510
      Width           =   1890
   End
   Begin VB.CommandButton CDeleteMedicine 
      Caption         =   "Delete Medicine"
      Height          =   540
      Left            =   2025
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6510
      Width           =   1890
   End
   Begin VB.CommandButton CAddNew 
      Caption         =   "Add New"
      Height          =   540
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6510
      Width           =   1890
   End
   Begin VB.CommandButton CFindNext 
      Caption         =   "Find Next"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5085
      Width           =   1470
   End
   Begin MSComctlLib.TreeView TrMedicines 
      Height          =   4785
      Left            =   285
      TabIndex        =   0
      Top             =   210
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
   Begin MSForms.TextBox TRack 
      Height          =   345
      Left            =   7140
      TabIndex        =   6
      Top             =   2475
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   405
      Left            =   4920
      TabIndex        =   23
      Top             =   2520
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
      Height          =   345
      Left            =   7140
      TabIndex        =   4
      Top             =   1620
      Width           =   2520
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4445;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   405
      Left            =   4920
      TabIndex        =   22
      Top             =   2055
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Unit"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox CoUnit 
      Height          =   345
      Left            =   7140
      TabIndex        =   5
      Top             =   2055
      Width           =   2520
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4445;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoMedicine 
      Height          =   345
      Left            =   7140
      TabIndex        =   2
      Top             =   750
      Width           =   2520
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4445;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox CoStatus 
      Height          =   330
      Left            =   7140
      TabIndex        =   8
      Top             =   3345
      Width           =   2520
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "4445;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label8 
      Height          =   405
      Left            =   4920
      TabIndex        =   20
      Top             =   3360
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
      Left            =   4920
      TabIndex        =   19
      Top             =   2955
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
      Height          =   345
      Left            =   7140
      TabIndex        =   7
      Top             =   2910
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;609"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label6 
      Height          =   405
      Left            =   4920
      TabIndex        =   18
      Top             =   1620
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
      Left            =   4920
      TabIndex        =   17
      Top             =   1215
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
      Height          =   345
      Left            =   7140
      TabIndex        =   3
      Top             =   1185
      Width           =   2520
      VariousPropertyBits=   746604571
      Size            =   "4445;609"
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
      Left            =   4920
      TabIndex        =   16
      Top             =   750
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
      Left            =   4920
      TabIndex        =   15
      Top             =   300
      Width           =   1530
      VariousPropertyBits=   8388627
      Caption         =   "Code"
      Size            =   "2699;714"
      FontName        =   "Arial Narrow"
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TMedicineCode 
      Height          =   330
      Left            =   7140
      TabIndex        =   1
      Top             =   330
      Width           =   2520
      VariousPropertyBits=   746604575
      Size            =   "4445;582"
      SpecialEffect   =   0
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox TFind 
      Height          =   315
      Left            =   270
      TabIndex        =   11
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
Attribute VB_Name = "FMedicineMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Database
Dim bCreateNewGroup As Boolean
Dim sUnitCode() As String, sManufacturerCode() As String

Private Sub getMedicine()
Dim rs As Recordset
    
    CoMedicine.Clear
    
    Set rs = db.OpenRecordset("Select MedicineMaster.MedicineName From MedicineMaster  Order By MedicineMaster.MedicineName")
    Do While rs.EOF = False
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    End If
    
    While rs.EOF = False
        CoMedicine.AddItem "" & rs!MedicineName
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
    
    If (TrMedicines.Nodes.Count = 0) Then
        MsgBox "Please Create a a Group First !", vbInformation
        Exit Sub
    End If
    
    If Left(Trim(TrMedicines.SelectedItem.Key), 1) = "B" Then
        MsgBox "Please Select any Group to create Medicine !", vbInformation
        Exit Sub
    End If

    clearEditCodtrols
    enableDisableControlsOnAdd
    TMedicineCode = getNewMedicinecode
    CoMedicine.SetFocus
End Sub

Private Sub CAddGroup_Click()
    clearEditCodtrols
    enableDisableControlsOnGroup
    TMedicineCode = getNewMedicinecode
    CoMedicine.SetFocus
    bCreateNewGroup = True
End Sub

Private Sub enableDisableControlsOnGroup()
    TMedicineCode.Enabled = False
    CoMedicine.Enabled = True
    TBillingName.Enabled = False
    CoManufacturer.Enabled = False
    TRack.Enabled = False
    CoUnit.Enabled = False
    TMinimumStock.Enabled = False
    CoStatus.Enabled = True
End Sub

Private Sub CClose_Click()
    Unload Me
End Sub

Private Sub CDeleteMedicine_Click()
Dim rs As Recordset
    
    If Trim(TMedicineCode.Text) = "" Then
        MsgBox "Please Select Any Medicine to Delete !", vbInformation
        Exit Sub
    End If
        
    If checkAlreadyUsed(Trim(TMedicineCode.Text)) Then
        MsgBox "The Medicine is Already Used !", vbInformation
        Exit Sub
    End If
    
    If checkForChildMedicines(Trim(TMedicineCode.Text)) Then
        MsgBox "The Group has Medicine Items, Please Delete them First !", vbInformation
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("Select MedicineMaster.* From MedicineMaster Where (MedicineMaster.Code = '" & Trim(TMedicineCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        rs.Delete
        rs.Close
    Else
        rs.Close
        MsgBox "The Medicine doesnt Exist !", vbInformation
        Exit Sub
    End If
    
    MsgBox "Successfully Deleted the Medicine !", vbInformation
    
    refreshTree
    clearEditCodtrols
End Sub

Private Function checkForChildMedicines(sMedicineCode As String) As Boolean
Dim rs As Recordset, bFound As Boolean

    Set rs = db.OpenRecordset("Select MedicineMaster.* From MedicineMaster Where (MedicineMaster.GroupCode = '" & Trim(sMedicineCode) & "' )")
    If rs.RecordCount > 0 Then
        bFound = True
    Else
        bFound = False
    End If
    rs.Close
    
    checkForChildMedicines = bFound
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
    
    Do While lFindIndex <= TrMedicines.Nodes.Count
        
        If InStr(1, LCase(TrMedicines.Nodes.Item(lFindIndex)), LCase(sFindWord), vbTextCompare) > 0 Then
            TrMedicines.Nodes.Item(lFindIndex).Selected = True
            getDetailsOfMedicine
            TrMedicines.SetFocus
            Exit Do
        End If
        lFindIndex = lFindIndex + 1
    Loop
    
    If lFindIndex > TrMedicines.Nodes.Count Then
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
Dim rs As Recordset, sStatus As String, sMedicineCode As String, sParenttype As String
Dim sParentCode As String

    If Trim(TMedicineCode.Text) = "" Then
        MsgBox "Please Select a Medicine to Edit or click Add New button To add new Medicine", vbInformation
        Exit Sub
    ElseIf Trim(CoMedicine.Text) = "" Then
        MsgBox "Please Enter needed Informations !", vbInformation
        CoMedicine.SetFocus
        Exit Sub
    ElseIf Trim(TBillingName.Text) = "" And Not bCreateNewGroup Then
        MsgBox "Please Enter needed Informations !", vbInformation
        TBillingName.SetFocus
        Exit Sub
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
    If TrMedicines.Nodes.Count > 0 Then
        If (bCreateNewGroup) Then
            sParenttype = ""
            sParentCode = ""
        Else
            sParenttype = Trim(Left(TrMedicines.SelectedItem.Key, 1))
            sParentCode = Trim(Right(TrMedicines.SelectedItem.Key, Len(TrMedicines.SelectedItem.Key) - 1))
        End If
    Else
        sParenttype = ""
        sParentCode = ""
    End If

    Set rs = db.OpenRecordset("Select MedicineMaster.* From MedicineMaster Where (MedicineMaster.Code = '" & Trim(TMedicineCode.Text) & "' )")
    If rs.RecordCount > 0 Then
        sStatus = "Edited"
        rs.Edit
    Else
        sStatus = "Added"
        TMedicineCode.Text = getNewMedicinecode()
        rs.AddNew
        rs!Code = Trim(TMedicineCode.Text)
        rs!Type = IIf(Trim(sParenttype) = "", "AGroup", "BMedicine")
        rs!GroupCode = sParentCode
    End If
    rs!MedicineName = Trim(CoMedicine.Text)
    rs!BillingName = Trim(TBillingName.Text)
    rs!ManufacturerCode = sManufacturerCode(CoManufacturer.ListIndex + 1)
    rs!UnitCode = sUnitCode(CoUnit.ListIndex + 1)
    rs!MinimumStock = Val(TMinimumStock.Text)
    rs!Status = IIf((CoStatus.ListIndex = 0), True, False)
    rs!Rack = Trim(TRack.Text)
    rs.Update
    rs.Close
    
    MsgBox "Successfully " & sStatus & " !", vbInformation
    
    refreshTree
    clearEditCodtrols
    bCreateNewGroup = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyF And ((Shift And 7) = 2)) Then
        CFindNext_Click
    ElseIf (KeyCode = vbKeyN And ((Shift And 7) = 2)) Then
        CAddNew_Click
    ElseIf (KeyCode = vbKeyD And ((Shift And 7) = 2)) Then
        CDeleteMedicine_Click
    ElseIf (KeyCode = vbKeyS And ((Shift And 7) = 2)) Then
        CSave_Click
    ElseIf (KeyCode = vbKeyC And ((Shift And 7) = 2)) Then
        CClose_Click
    End If
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase("Storage.mdb", False, False, "MS Access;PWD=12345abcde")
    
    CoStatus.AddItem "Enabled"
    CoStatus.AddItem "Disabled"
    
    refreshTree
    enableDisableControls
    getMedicine
    getManufacturer
    getUnit
    bCreateNewGroup = False
End Sub

Private Sub refreshTree()
Dim rs As Recordset
    
    TrMedicines.Nodes.Clear
    
    Set rs = db.OpenRecordset("Select MedicineMaster.Code,MedicineMaster.MedicineName,MedicineMaster.Type,MedicineMaster.GroupCode From MedicineMaster Order By MedicineMaster.Type,MedicineMaster.MedicineName")
    While rs.EOF = False
        If Trim(rs!Type) = "AGroup" Then
            TrMedicines.Nodes.Add , , "A" & rs!Code, rs!MedicineName
            'TrMedicines.Nodes(TrMedicines.Nodes.Count).Bold = True
            'TrMedicines.Nodes(TrMedicines.Nodes.Count).ForeColor = &H808080
        ElseIf Trim(rs!Type) = "BMedicine" Then
            TrMedicines.Nodes.Add "A" & rs!GroupCode, tvwChild, "B" & rs!Code, rs!MedicineName
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub TMedicineCode_GotFocus()
    TMedicineCode.SelStart = 0
    TMedicineCode.SelLength = Len(TMedicineCode.Text)
End Sub


Private Sub CoMedicine_GotFocus()
    CoMedicine.SelStart = 0
    CoMedicine.SelLength = Len(CoMedicine.Text)
End Sub

Private Sub TFind_GotFocus()
    TFind.SelStart = 0
    TFind.SelLength = Len(TFind.Text)
End Sub

Private Sub TrMedicines_Click()
    enableDisableControls
    'getDetailsOfMedicine
End Sub

Private Sub TrMedicines_NodeClick(ByVal Node As MSComctlLib.Node)
    enableDisableControls
    If TrMedicines.Nodes.Count > 0 Then
        getDetailsOfMedicine
    End If
End Sub

Private Sub getDetailsOfMedicine()
Dim rs As Recordset
    
    If (Left(TrMedicines.SelectedItem.Key, 1) = "A") Then
        Set rs = db.OpenRecordset("Select '' As ManufacturerName,MedicineMaster.Code,MedicineMaster.MedicineName,MedicineMaster.BillingName,MedicineMaster.ManufacturerCode,MedicineMaster.MinimumStock,MedicineMaster.Status,MedicineMaster.Rack,'' As UnitName From MedicineMaster Where (MedicineMaster.Code = '" & Trim(Right(TrMedicines.SelectedItem.Key, Len(TrMedicines.SelectedItem.Key) - 1)) & "' )")
    ElseIf (Left(TrMedicines.SelectedItem.Key, 1) = "B") Then
        Set rs = db.OpenRecordset("Select Manufacturer.ManufacturerName,MedicineMaster.Code,MedicineMaster.MedicineName,MedicineMaster.BillingName,MedicineMaster.ManufacturerCode,MedicineMaster.MinimumStock,MedicineMaster.Status,MedicineMaster.Rack,Units.UnitName From Manufacturer,MedicineMaster,Units Where (Manufacturer.Code = MedicineMaster.ManufacturerCode ) And (Units.Code = MedicineMaster.UnitCode ) And (MedicineMaster.Code = '" & Trim(Right(TrMedicines.SelectedItem.Key, Len(TrMedicines.SelectedItem.Key) - 1)) & "' )")
    Else
        Exit Sub
    End If
    
    If rs.RecordCount > 0 Then
        
        TMedicineCode.Text = "" & rs!Code
        CoMedicine.Text = "" & rs!MedicineName
        TBillingName.Text = "" & rs!BillingName
        CoManufacturer.Text = "" & rs!ManufacturerName
        TRack.Text = "" & rs!Rack
        CoUnit.Text = "" & rs!UnitName
        TMinimumStock.Text = Val("" & rs!MinimumStock)
        CoStatus.ListIndex = IIf((rs!Status = True), 0, 1)
    Else
        clearEditCodtrols
    End If
    rs.Close
End Sub

Private Sub enableDisableControlsOnAdd()
    
    If Left(TrMedicines.SelectedItem.Key, 1) = "A" Then
        
        TMedicineCode.Enabled = False
        CoMedicine.Enabled = True
        TBillingName.Enabled = True
        CoManufacturer.Enabled = True
        TRack.Enabled = True
        CoUnit.Enabled = True
        TMinimumStock.Enabled = True
        CoStatus.Enabled = True
    ElseIf Left(TrMedicines.SelectedItem.Key, 1) = "B" Then
        
        TMedicineCode.Enabled = False
        CoMedicine.Enabled = True
        TBillingName.Enabled = True
        CoManufacturer.Enabled = True
        TRack.Enabled = True
        CoUnit.Enabled = True
        TMinimumStock.Enabled = True
        CoStatus.Enabled = True
    End If
End Sub

Private Sub enableDisableControls()
    
    If TrMedicines.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    If Left(TrMedicines.SelectedItem.Key, 1) = "A" Then
        
        TMedicineCode.Enabled = False
        CoMedicine.Enabled = True
        TBillingName.Enabled = False
        CoManufacturer.Enabled = False
        TRack.Enabled = False
        CoUnit.Enabled = False
        TMinimumStock.Enabled = False
        CoStatus.Enabled = True
    ElseIf Left(TrMedicines.SelectedItem.Key, 1) = "B" Then
        
        TMedicineCode.Enabled = False
        CoMedicine.Enabled = True
        TBillingName.Enabled = True
        CoManufacturer.Enabled = True
        TRack.Enabled = True
        CoUnit.Enabled = True
        TMinimumStock.Enabled = True
        CoStatus.Enabled = True
    End If
End Sub

Private Sub clearEditCodtrols()
    TMedicineCode.Text = ""
    CoMedicine.Text = ""
    TBillingName.Text = ""
    CoManufacturer.Text = ""
    TRack.Text = ""
    CoUnit.Text = ""
    TMinimumStock.Text = ""
    CoStatus.Text = ""
End Sub

Private Function getParentMedicine(sMedicineCode As String) As String
Dim rs As Recordset, sParentCode As String
    
    Set rs = db.OpenRecordset("Select MedicineMaster.GroupCode From MedicineMaster Where (MedicineMaster.Code = '" & Trim(sMedicineCode) & "' )")
    If rs.RecordCount > 0 Then
        sParentCode = "" & rs!GroupCode
    Else
        sParentCode = ""
    End If
    rs.Close
    
    getParentMedicine = sParentCode
End Function

Private Function getNewMedicinecode() As String
Dim rs As Recordset, sMedicineCode As String
    
    Set rs = db.OpenRecordset("Select Max(val(MedicineMaster.Code))As ACode From MedicineMaster")
    If rs.RecordCount > 0 Then
        sMedicineCode = Val("" & rs!ACode) + 1
    Else
        sMedicineCode = "1"
    
    End If
    rs.Close
    
    getNewMedicinecode = sMedicineCode
End Function

Private Function checkAlreadyUsed(sMCode As String) As Boolean
Dim rs As Recordset
Dim bExist As Boolean
    bExist = False
    Set rs = db.OpenRecordset("Select Transaction.* From Transaction Where (Transaction.MedicineCode = '" & sMCode & "' )")
    If rs.RecordCount > 0 Then
        bExist = True
    End If
    rs.Close
    
    checkAlreadyUsed = bExist
End Function

