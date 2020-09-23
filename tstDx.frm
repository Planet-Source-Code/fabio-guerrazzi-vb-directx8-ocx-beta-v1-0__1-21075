VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{73DA769C-6D87-4330-898E-0FB2771367FB}#1.0#0"; "VBDirectx8.ocx"
Begin VB.Form tstDx 
   Caption         =   "Dx8 debug project"
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   10575
   Icon            =   "tstDx.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VBDirectX8.D3D8Tree D3D8Tree1 
      Height          =   4455
      Left            =   7500
      TabIndex        =   10
      Top             =   480
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   7858
   End
   Begin VBDirectX8.D3D8 D3D81 
      Height          =   5715
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   10081
   End
   Begin VB.Frame Frame1 
      Caption         =   "    Walk"
      Height          =   1155
      Left            =   7440
      TabIndex        =   2
      Top             =   5040
      Width           =   3075
      Begin VB.CommandButton Command1 
         Caption         =   "?"
         Height          =   195
         Left            =   2580
         TabIndex        =   8
         Top             =   0
         Width           =   435
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   0
         Width           =   195
      End
      Begin MSComctlLib.Slider ISpeed 
         Height          =   195
         Left            =   1620
         TabIndex        =   5
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         _Version        =   393216
      End
      Begin MSComctlLib.Slider LSpeed 
         Height          =   195
         Left            =   1620
         TabIndex        =   3
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   344
         _Version        =   393216
         Min             =   1
         Max             =   50
         SelStart        =   1
         TickFrequency   =   5
         Value           =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Inertia Level"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Linear Speed"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1275
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3900
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5775
      Left            =   7140
      Max             =   0
      Min             =   -1000
      TabIndex        =   0
      Top             =   420
      Value           =   -20
      Width           =   255
   End
   Begin MSComctlLib.ImageList ILTree 
      Left            =   3120
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   46
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":0CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":0DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":0EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1008
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":111A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":122C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":133E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1450
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1562
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1674
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1786
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1898
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":19AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1CE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":1F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2016
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2128
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":223A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":234C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":245E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2570
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2682
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2794
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":28A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":29B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2CEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":2F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":3024
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":3136
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":3248
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":335A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tstDx.frx":346C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ILTree"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "saveas"
            ImageIndex      =   38
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "properties"
            ImageIndex      =   39
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "select"
            Object.ToolTipText     =   "Select"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "zoom"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pan"
            ImageIndex      =   10
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "displaymode"
            ImageIndex      =   3
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grid"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "texture"
            ImageIndex      =   28
            Style           =   1
            Value           =   1
         EndProperty
      EndProperty
   End
   Begin VB.Menu File_m 
      Caption         =   "File"
      Begin VB.Menu Open_i 
         Caption         =   "Open"
      End
      Begin VB.Menu Save_i 
         Caption         =   "Save"
      End
      Begin VB.Menu ImportDXF_i 
         Caption         =   "Insert 3D DXF File"
      End
      Begin VB.Menu inAscii 
         Caption         =   "Insert Ascii File"
      End
      Begin VB.Menu InsLand 
         Caption         =   "Insert Land"
      End
      Begin VB.Menu exit_i 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu scenes_m 
      Caption         =   "Scenes"
      Begin VB.Menu NewScene 
         Caption         =   "New"
      End
      Begin VB.Menu Test1 
         Caption         =   "Simple scene (No x files) Test 1"
      End
      Begin VB.Menu Test2 
         Caption         =   "Simple scene (No x files) Test 2"
      End
      Begin VB.Menu Colors 
         Caption         =   "Gamma Internal colors"
      End
      Begin VB.Menu Test3 
         Caption         =   "Environment Game Test 1"
      End
   End
   Begin VB.Menu opts 
      Caption         =   "Options"
      Begin VB.Menu DeviceChange 
         Caption         =   "Change Device"
      End
   End
   Begin VB.Menu mView 
      Caption         =   "View"
      Begin VB.Menu viewas 
         Caption         =   "WireFrame"
         Index           =   0
      End
      Begin VB.Menu viewas 
         Caption         =   "Flat"
         Index           =   1
      End
      Begin VB.Menu viewas 
         Caption         =   "Gauraud"
         Index           =   2
      End
   End
   Begin VB.Menu Helpm 
      Caption         =   "Help"
      Begin VB.Menu Walk_i 
         Caption         =   "Walk Keys"
      End
      Begin VB.Menu About_i 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "tstDx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub About_i_Click()
  Dim St$
  St = "Fabio Guerrazzi, http://digilander.iol.it/WarZi/default.htm - e-mail fabiog2@libero.it"
  MsgBox St
End Sub

Private Sub Check2_Click()
 If Check2 = 1 Then
    D3D81.EditState = 3
 Else
    D3D81.EditState = 0
 End If
 
End Sub

Private Sub Colors_Click()
  Dim i As Integer
  Dim X As Integer
  Dim Y As Integer
  Dim s As Integer
  Dim Hnd$
  
' Load 80 boxes, with all internal preset colors
  
  Screen.MousePointer = 11
  Y = -2
  
  With D3D81
       .ClearScene
       .BackColor = QBColor(8) 'Pic.BackColor
       .SetCameraView 0, 0, -50
  
       .newFrame "Gamma"
       
       For i = 1 To 80
         
           Hnd = .GenerateHandle
          .AddBox "Gamma", Hnd, 10, 10, 10
          .Mesh_SetColor Hnd, i, 0
          
          
          X = X + 1
          If X = 11 Then
             Y = Y + 1
             X = 1
          End If
          
          
          .Mesh_Translate Hnd, 0, (X * 13), (Y * 13)
    Next
  
    .RefreshTree
    .DrawScene
  
  End With

 
 Screen.MousePointer = 0

End Sub

Private Sub Command1_Click()
  Walk_i_Click
End Sub

Private Sub DeviceChange_Click()
 
 ' try commenting the exit sub.. but it doesn't works to me :/
 ' can't switch to FullScreen.. i don't know why yet
  
 Exit Sub
 
 
 ' .InitMode:  0-ask for device windowed or fullscreen, 1-Windowed, 2-FullScreen
   D3D81.InitMode = 0
   
   If Not D3D81.InitDevice Then
       MsgBox "Unable to CreateDevice." ' no video card available
       End
   End If

   Select Case D3D81.DeviceMode ' here i get which InitDevice found as best mode
     Case 2:    Caption = "Accelleration Mode: Medium (D3DDEVTYPE_REF) "
     Case 3:    Caption = "Accelleration Mode: Low! (D3DDEVTYPE_SW) "
     Case 1:    Caption = "Accelleration Mode: Best (D3DDEVTYPE_HAL)"
   End Select

End Sub

Private Sub exit_i_Click()
   End
End Sub

Private Sub Form_Load()
   
   
 ' .InitMode:  0-ask for device/window or fullscreen, 1-Windowed, 2-FullScreen
   D3D81.InitMode = 1
   
   If Not D3D81.InitDevice Then
       MsgBox "Unable to CreateDevice." ' no video card available
       End
   End If

   Select Case D3D81.DeviceMode ' here i get which InitDevice found as best mode
     Case 2:    Caption = "Accelleration Mode: Medium (D3DDEVTYPE_REF) "
     Case 3:    Caption = "Accelleration Mode: Low! (D3DDEVTYPE_SW) "
     Case 1:    Caption = "Accelleration Mode: Best (D3DDEVTYPE_HAL)"
   End Select


   D3D81.MediaPath = App.Path & "\Media\" ' sets the media folder where to get all x files and textures
   
End Sub

Private Sub Form_Resize()
Exit Sub
  VScroll1.Left = ScaleWidth - VScroll1.Width
  VScroll1.Height = ScaleHeight - Toolbar1.Height
  D3D81.Width = VScroll1.Left - D3D81.Left
  D3D81.Height = ScaleHeight - Toolbar1.Height
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
  D3D81.ClearScene
End Sub


Private Sub ImportDXF_i_Click()
 Dim File$
 With CommonDialog1
   .Filter = "Autocad DXF (*.DXF)|*.DXF"
   .ShowOpen
   File = .FileName
 End With
 
 With D3D81
        .BackColor = QBColor(8) 'Pic.BackColor
       .SetCameraView 0, 5, -50

        .newFrame "dxf1"
        .xMesh.Init
        .xMesh.LoadDXF File
        .AddUserMesh "dxf1", "myMesh2", .xMesh.ResolveMesh
        .Mesh_AddMaterial "myMesh2", 0, QBColor(9), &H6FC0C0C0, QBColor(15), 100
       ' .Mesh_AddTexture "myMesh2", 0, "c:\banana.bmp" ' "E:\Dx8SDK\samples\Multimedia\VBSamples\Media\ground2.bmp"
        
        .xMesh.Init
       ' .Frame_Translate "user1", -8, 1, 0
       ' .Mesh_AddMaterial "myMesh1", 0, QBColor(12), &H6FC0C0C0, QBColor(1), 150
       ' .Mesh_AddTexture "myMesh2", 0, "c:\banana.bmp" ' "E:\Dx8SDK\samples\Multimedia\VBSamples\Media\ground2.bmp"
    
       .DrawScene
 
 
 End With
 
End Sub

Private Sub inAscii_Click()
 Dim File$
 With CommonDialog1
   .Filter = "Ascii 3D user file (*.DAT)|*.DAT"
   .ShowOpen
   File = .FileName
 End With
 If Len(File) = 0 Then Exit Sub
 
 With D3D81
        .BackColor = QBColor(8)
        .SetCameraView 0, 5, -50
        .Pause = True        ' turn on the pause else drawing the scene it will crash
        Set DxCTL = D3D81
          LoadPCMFile File   ' add my own file structure.. any other is ok
        Set DxCTL = Nothing
        .Pause = False
        .RefreshTree
        .DrawScene
 End With

End Sub


Private Sub InsLand_Click()
 Dim File$
 With CommonDialog1
   .Filter = "Delaunay File (*.txt)|*.TXT"
   .ShowOpen
   File = .FileName
 End With
 If Len(File) = 0 Then Exit Sub
 
 With D3D81
        .BackColor = QBColor(8) 'Pic.BackColor
        .SetCameraView 0, 5, -50
        .xMesh.Init
        .newFrame "Delaunay"
        .LoadLand File
        .AddUserMesh "Delaunay", "MyLand", .xMesh.ResolveMesh
        .xMesh.Init
        .Mesh_AddTexture "myLand", 0, App.Path & "\media\seafloor.bmp" ' "c:\texr1.bmp" ' "E:\Dx8SDK\samples\Multimedia\Media\env3.bmp" 'App.Path & "\media\seafloor.bmp" '
        
        .DrawScene
 End With

End Sub

Private Sub ISpeed_Click()
 On Error Resume Next
 D3D81.InertiaSpeed = ISpeed / 10 ' ' 0 to 1
End Sub

Private Sub LSpeed_Click()
        
        D3D81.LinearSpeed = LSpeed   ' 0 to 20

End Sub

Private Sub NewScene_Click()
  D3D81.ClearScene
'  D3D81.SetWorld 0, 0, 0, 100, 100, 100
 ' D3D81.SetCameraView 0, 5, -200
  D3D81.DrawScene
End Sub


Private Sub Open_i_Click()
 Dim File$
 With CommonDialog1
   .Filter = "DirectX File (*.x)|*.x"
   .ShowOpen
   File = .FileName
 End With
 
 With D3D81
        
        .BackColor = QBColor(8) 'Pic.BackColor
        .SetCameraView 0, 5, -50
        .LoadXFile File, "xFrame1"
        .RefreshTree
    ' .Mesh_AddTexture "myMesh1", 0, "E:\Dx8SDK\samples\Multimedia\Media\spheremap.bmp"

       .DrawScene
 
 End With
 

End Sub





Private Sub Test1_Click()

Dim s As Integer
   
With D3D81 ' <--- here starts the control features
   
     
       .BackColor = QBColor(8) 'Pic.BackColor
       .SetCameraView 0, 0, -50
       
     
     ' Objects/Scene Construction..
     ' few methods to build almost all :/
     ' look the structure in this way:
     '  Frames are the object you can move and rotate in the scene
     '  Meshes are objects added to single frames
     '   such
     '    I add a frame called "Frame1" where next i add two mesh objects "m1" and "m2"
     '    so far, the structure could be view like an object tree where i define all object
     '    by knot moving them
     
     ' AddNewFrame:                                    ' Add a new frame to the scene
     ' Addxxx <framename>,<meshname>, <parameters>     ' add a mesh to the frame object
     ' Mesh_AddMaterial <meshName>, <material_count>, _
     '                  <Ambient color>,<Diffuse color>,<specular color> _
     '                  <power 0-255>,<optional 1=sets as transparent object>
     '                  ' apply a material/color to the mesh
     ' frame_Translate  <posxyz) ' translate frame
     
       
       .newFrame "Plane1"
       .AddBox "Plane1", "Box1", 10, 10, 0.1
      ' .AddShapeSphere "Plane1", "sf1", 50, 16, 16
       .Mesh_AddMaterial "sf1", 0, QBColor(12), &H6FC0C0C0, QBColor(15), 100, 1
       .Mesh_AddMaterial "Box1", 0, QBColor(12), &H6FC0C0C0, QBColor(7), 0
       
       .newFrame "Cube1" ' Add new frame as cube
       .AddBox "Cube1", "2", 2, 2, 4
       .Frame_Translate "Cube1", 0, 1, 0
       .Mesh_AddMaterial "2", 0, QBColor(1), &H6FC0C0C0, QBColor(7), 0
      
               
       .newFrame "Sphere1" ' Add new frame as sphere
       .AddShapeSphere "Sphere1", "sfera1", 3, 16, 16
       .Frame_Translate "Sphere1", 0, 5, 0
       .Mesh_AddMaterial "sfera1", 0, QBColor(6), &H6FC0C0C0, QBColor(15), 100, 1
     '  .Frame_Select("Sphere1") = True
       
      ' .Mesh_AddTexture "sfera1", 0, "E:\Dx8SDK\samples\Multimedia\VBSamples\Media\texr1.bmp"
     '  .Mesh_AddTexture "Box1", 0, "E:\Dx8SDK\samples\Multimedia\VBSamples\Media\dx_logo.bmp"
    
    
    
    
' Here an extruded object using the xMesh modeler class
' just add the 2D contour and calls the Extrude method with the height value
' Next, it is treat as a normal mesh object and we can apply material or textures
' Note: at the moment the textures are applyed only by faces
'
        .newFrame "user1"
        .xMesh.Init
        .xMesh.AddProfileVertex 0, 0
        .xMesh.AddProfileVertex 0, 5
        .xMesh.AddProfileVertex 3, 2
        .xMesh.AddProfileVertex 5, 5
        .xMesh.AddProfileVertex 5, 0
        .xMesh.Extrude 3
        .AddUserMesh "user1", "myMesh1", .xMesh.ResolveMesh
        .xMesh.Init
        .Frame_Translate "user1", -8, 1, 0
        .Mesh_AddMaterial "myMesh1", 0, QBColor(12), &H6FC0C0C0, QBColor(1), 150
        
  ' uncomment these lines typing any valid bitmap
  
    '    .Mesh_AddTexture "myMesh1", 0, "c:\texr1.bmp" '"E:\Dx8SDK\samples\Multimedia\Media\env3.bmp"
    '    .Mesh_AddTexture "Box1", 0, "E:\Dx8SDK\samples\Multimedia\VBSamples\Media\dx_logo.bmp", 1
    
        .newFrame "Poly1"
        .AddPolygon "Poly1", "MyPoly", 30, 4
        .Mesh_AddMaterial "MyPoly", 0, QBColor(14), &H6FC0C0C0, QBColor(12), 0, 1
      
       .RefreshTree
       .DrawScene
   End With
   

End Sub

Private Sub Test2_Click()
    With D3D81
  
      '  .ClearScene
        
        .newFrame "F1"
        .xMesh.Init
        .xMesh.AddProfileVertex 0, 0
        .xMesh.AddProfileVertex 100, 0
        .xMesh.AddProfileVertex 100, 100
        .xMesh.AddProfileVertex 0, 100
        .xMesh.Extrude 0  ' this will cause a generation of a single polygon
        .AddUserMesh "F1", "M1", .xMesh.ResolveMesh
        .xMesh.Init
        .Mesh_SetColor "M1", 11, 0
          
        .newFrame "Fondale"
        .xMesh.Init
        .xMesh.AddProfileVertex 0, 95
        .xMesh.AddProfileVertex 0, 100
        .xMesh.AddProfileVertex 100, 100
        .xMesh.AddProfileVertex 100, 95
        .xMesh.Extrude 50
        .AddUserMesh "Fondale", "M2", .xMesh.ResolveMesh
        .xMesh.Init
        .Mesh_SetColor "M2", 14, 0
          
          
        .newFrame "Laterale"
        .xMesh.Init
        .xMesh.AddProfileVertex 0, 0
        .xMesh.AddProfileVertex 5, 0
        .xMesh.AddProfileVertex 5, 95
        .xMesh.AddProfileVertex 0, 95
        .xMesh.Extrude 50
        .AddUserMesh "Laterale", "M3", .xMesh.ResolveMesh
        .xMesh.Init
        .Mesh_SetColor "M3", 56, 0
      
      
        .newFrame "Front"
        .xMesh.Init
        .xMesh.AddProfileVertex -10, -10
        .xMesh.AddProfileVertex -8, -10
        .xMesh.AddProfileVertex -8, 60
        .xMesh.AddProfileVertex -10, 60
        .xMesh.Extrude 25
        .AddUserMesh "Front", "M4", .xMesh.ResolveMesh
        .xMesh.Init
        .Mesh_SetColor "M3", 37, 0
          
         .Frame_Translate "Front", 0, -10, 0
                    
        ' .mesh_addTexture "M1", 0, "c:\0\012.jpg"
         .SetCameraView 0, 0, -150
         .RefreshTree
         .DrawScene
                    
              
  End With


' RefreshList1
 

End Sub

Private Sub Test3_Click()
    
    Dim Hnd$
    
    With D3D81
  
        .ClearScene
        .SetCameraView 0, 0, -50
      
      ' sets movement parameters
        
        .InertiaSpeed = 0.8 ' 0 to 1
        .LinearSpeed = 5    ' 0 to 20
        .EditState = 3 ' process keyboard input, field camera movement
        
      ' set up video controls
        ISpeed = 8
        LSpeed = 5
        Check2 = 1
        
      ' add skybox
        
        Hnd = .GenerateHandle        ' build a random frame name for future calls
        .AddSkyMap "skybox2.x", Hnd  ' set the skymap (it MUST be a cube x file with its own texure)
        
      
      ' add terrain
      
        Hnd = .GenerateHandle
        .LoadXFile "SeaFloor.x", Hnd
        .Frame_Scale Hnd, 3
        .Frame_Translate Hnd, 0, 25, 0
'       .Frame_Scale Hnd, 30
'       .Frame_Translate Hnd, 0, 270, 0
        
  
   ' insert an user ascii file as building
   ' you can insert here whatever you want.. a dxf an x file or any other 3D objects
        Set DxCTL = D3D81
          .Pause = True        ' turn on the pause else drawing the scene it will crash
          LoadPCMFile App.Path & "\media\REND7.DAT"
          .Pause = False
        Set DxCTL = Nothing
        
        .RefreshTree
        .DrawScene
        .SetFocus
              
  End With


End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "open": Open_i_Click
      '  Case "save": Save_Click
        Case "select": D3D81.EditState = 1
        Case "pan": D3D81.EditState = 0
      '  Case "grid": vgrid_Click
    End Select
End Sub

Private Sub viewas_Click(Index As Integer)
 D3D81.RenderMode = Index
End Sub

Private Sub VScroll1_Change()
 D3D81.SetCameraView 0, 5, VScroll1
End Sub


Private Sub VScroll1_Scroll()
  D3D81.SetCameraView 0, 5, VScroll1
End Sub


Private Sub Walk_i_Click()
 Dim St As String
 
    St = "Key 5 :  Slide Right"
    St = St & vbCrLf & "Key 4: Slide Left"
    St = St & vbCrLf & "Key PageUp:  Move up"
    St = St & vbCrLf & "Key PageDown:  Move down"
    
    St = St & vbCrLf & "Key Up: Move Forward"
    St = St & vbCrLf & "Key Down:  Move Backward"
    
    St = St & vbCrLf & "Key Right: Yaw right"
    St = St & vbCrLf & "Key Left:  Yaw left"
    
    St = St & vbCrLf & "KeyZ: turn down"
    St = St & vbCrLf & "KeyA:  turn up"
    MsgBox St
End Sub


