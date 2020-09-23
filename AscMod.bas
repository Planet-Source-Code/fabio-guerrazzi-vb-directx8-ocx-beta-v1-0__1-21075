Attribute VB_Name = "AsciiModule"
Type rCoord
   X As Double
   Y As Double
   z1 As Double
   z2 As Double
End Type

Type RkPar
    Mode As Integer
    Tip As Integer
    nv As Integer
    Crd() As rCoord
    Col As Long
    dis As Boolean
End Type

Type ASLayer
   nome As String
   nPar As Integer ' Numero Pareti
   Par() As RkPar
End Type




Type Coord
  X As Double
  Y As Double
  z As Double
End Type

Type Box3D
   xmin As Double
   xmax As Double
   ymin As Double
   ymax As Double
   zmin As Double
   zmax As Double
End Type

Const BIG = 1E+30

Public DxCTL As Control
Sub LoadPCMFile(File As String)
  
 ' this is my own file structure to generate buildings.
 ' you can use whatever programmable structure you want
 
  
  Dim St As String, X#, Y#, z1#, z2#, Dummy$
  Dim Piani%, Fondazioni%
  Dim PN() As ASLayer
  Dim i%, j%, k%, s%
  Dim Mode%, Tip%, nv%
  Dim Col&, dis%
  Dim Dex$
  
  Dim v() As Coord
  Dim Bx As Box3D
  
  
  
If Len(File) = 0 Then Exit Sub
  
  ReDim Preserve PN(1)
  
  n = FreeFile
  Open File For Input As #n
  Line Input #n, Dummy ' DIMENSIONAMENTI
  Input #n, Piani
  Input #n, i, PN(0).nPar
  
  ReDim Preserve PN(Piani)
  For i = 1 To Piani
        Input #n, j, PN(i).nPar
  Next
  
  
  For i = 0 To Piani
         Line Input #n, nome$  ' PARETInn  (Nome del Layer)
         If i = 0 Then nome = "0"
         
         ReDim PN(i).Par(PN(i).nPar)
         
         For k = 1 To PN(i).nPar
                Input #n, j, Mode, Tip, nv
                ReDim PN(i).Par(k).Crd(nv)
                PN(i).Par(k).nv = nv
                PN(i).Par(k).Mode = Mode
                PN(i).Par(k).Tip = Tip
                
                For s = 1 To nv
                     Input #n, X, Y, z1, z2
                     PN(i).Par(k).Crd(s).X = X
                     PN(i).Par(k).Crd(s).Y = Y
                     PN(i).Par(k).Crd(s).z1 = z1
                     PN(i).Par(k).Crd(s).z2 = z2
                Next
                Input #n, Col, dis
                PN(i).Par(k).Col = Col
                PN(i).Par(k).dis = dis
                
                ReDim v(nv)
                Bx = InitBox3D
                
                z# = 0
                For s = 1 To nv
                    v(s).X = PN(i).Par(k).Crd(s).X
                    v(s).Y = PN(i).Par(k).Crd(s).Y
                    v(s).z = PN(i).Par(k).Crd(s).z1
                    z# = Max(PN(i).Par(k).Crd(s).z2, z#)
                    AssignBox3D Bx, v(s)
                Next
                
                cl% = GetQBColor(Col)
              ' here i have all 2D contour, now perform the Dx Extrusion
                AddExtrusion v, nv, z, 0, 0, 0, cl
   
         Next
  
  Next
  
  

End Sub

Function Max(v1 As Double, v2 As Double) As Double
  If v1 > v2 Then Max = v1 Else Max = v2
End Function

Function Min(v1 As Double, v2 As Double) As Double
  If v1 < v2 Then Min = v1 Else Min = v2
End Function

Sub AssignBox3D(b As Box3D, P As Coord)
   
    With P
       If .X < b.xmin Then b.xmin = .X
       If .Y < b.ymin Then b.ymin = .Y
       If .z < b.zmin Then b.zmin = .z
       
       If .X > b.xmax Then b.xmax = .X
       If .Y > b.ymax Then b.ymax = .Y
       If .z > b.zmax Then b.zmax = .z
   End With
   
End Sub

Sub AddExtrusion(v() As Coord, nv As Integer, Height#, X#, Y#, z#, Color%)
    
'  might be better doing a single frame then add all meshes to it
'  or a root frame with its tree with frames/meshes
'  now.. for clarity i use a frame by mesh
With DxCTL
         FName$ = "F-" & .GenerateHandle
         mName$ = "M-" & .GenerateHandle
        
        .newFrame FName
        .xMesh.Init
         For i% = 1 To nv
            .xMesh.AddProfileVertex CSng(v(i).X), CSng(v(i).Y)
         Next
         .xMesh.Extrude CSng(Height)
        
        .AddUserMesh FName, mName, .xMesh.ResolveMesh
        .xMesh.Init
        If v(1).z <> 0 Then .Frame_Translate FName, 0, CSng(v(1).z), 0
        .Mesh_AddMaterial mName, 0, QBColor(Color), QBColor(Color), QBColor(0), 1
'&H6FC0C0C0
      '  .Mesh_AddTexture mName, 0, "c:\banana.bmp"
End With
    
End Sub


Function GetQBColor(Rg As Long) As Integer
Dim c As Integer

' Riottiene il codice QBColor dal Long RGB corrispondente

c = -1

Select Case Rg
     Case 0: c = 0
     Case 8388608: c = 1
     Case 32768: c = 2
     Case 8421376: c = 3
     Case 128: c = 4
     Case 8388736: c = 5
     Case 32896: c = 6
     Case 12632256: c = 7
     Case 8421504: c = 8
     Case 16711680: c = 9
     Case 65280: c = 10
     Case 16776960: c = 11
     Case 255: c = 12
     Case 16711935: c = 13
     Case 65535: c = 14
     Case 16777215: c = 15
     Case Else
        c = 0    ' nel caso non sia un Long Generato da QBColor restituisce "Nero"
End Select

 GetQBColor = c
    

End Function


Function QBColor2Plastica(Color As Integer) As Integer

' VB QbColor            Plastica
' =============         ===========
'0   Nero                  8
'1   Blu                   4
'2   Verde                 3
'3   Azzurro               9
'4   Rosso                 2
'5   Fucsia                7
'6   Giallo                5
'7   Bianco                33
'8   Grigio                21
'9   Blu chiaro            32
'10  Verde limone          28
'11  Azzurro chiaro        70
'12  Rosso chiaro          30
'13  Fucsia chiaro         31
'14  Giallo chiaro         69
'15  Bianco brillante      1

Dim Rt%
Select Case Color
    Case 0: Rt = 8
    Case 1: Rt = 4
    Case 2: Rt = 3
    Case 3: Rt = 9
    Case 4: Rt = 2
    Case 5: Rt = 7
    Case 6: Rt = 5
    Case 7: Rt = 33
    Case 8: Rt = 21
    Case 9: Rt = 32
    Case 10: Rt = 28
    Case 11: Rt = 70
    Case 12: Rt = 30
    Case 13: Rt = 31
    Case 14: Rt = 69
    Case 15: Rt = 1
    Case Else
      Rt = 1
End Select


QBColor2Plastica = Rt

End Function


Function InitBox3D() As Box3D
 
 With InitBox3D
   
   .xmin = BIG
   .ymin = BIG
   .zmin = BIG
   
   .xmax = -BIG
   .ymax = -BIG
   .zmax = -BIG

 End With
 
 
End Function

