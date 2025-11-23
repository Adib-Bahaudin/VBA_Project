VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub copi_code()

    Dim Var1 As Shape
    Dim Var2 As Shape
    Dim Var3 As Shape
    Dim Var4 As Shape
    
    ' Pastikan ada objek yang dipilih
    If ActiveShape Is Nothing Then
        MsgBox "Tidak ada objek yang dipilih.", vbExclamation
        Exit Sub
    End If
    
    ' copi objek
    Set Var1 = ActiveShape
    Set Var2 = Var1.Duplicate
    Set Var3 = Var1.Duplicate
    Set Var4 = Var1.Duplicate
    
    ActiveDocument.Unit = cdrMillimeter         ' merubah satuan menjadi mm
    ActiveDocument.ReferencePoint = cdrCenter   ' fokus objek berada di center
    
    ' menyusun objek
    Offset = Var1.SizeWidth + 2
    
    Var2.Move Offset, 0
    Var3.Move (Offset * 2), 0
    Var4.Move (Offset * 3), 0
    
    ' memilih objek
    ActiveDocument.ClearSelection   ' kosongkan selection dulu
    ActiveDocument.AddToSelection Var1, Var2, Var3, Var4
    
    ' menyatukan objek
    Dim Varall As Shape
    Set Varall = ActiveSelection.Group
    
    ' set objek ketengah halaman
    Varall.AlignAndDistribute 3, 3, 2, 0, False, 2

End Sub

Sub Cetak_Pas_Foto()

    If ActiveShape Is Nothing Then
        MsgBox "Tidak ada objek yang dipilih.", vbExclamation
        Exit Sub
    End If
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.MasterPage.SetSize 148#, 210#
    
    If ActiveShape.SizeWidth = 40 And ActiveShape.SizeHeight = 60 Then
    
        Dim En As form_4x6
        Set En = New form_4x6
        En.Show
        
        If En.IsCancelled Then
            Exit Sub
        End If
        
        Call Empat_Enam(En.Jumlah)
        
    ElseIf ActiveShape.SizeWidth = 30 And ActiveShape.SizeHeight = 40 Then
    
        Dim Em As form_3x4
        Set Em = New form_3x4
        Em.Show
        
        If Em.IsCancelled Then
            Exit Sub
        End If
        
        Call Tiga_Empat(Em.Jumlah)
        
    Else
        MsgBox "Ukuran Pas Foto Tidak Valid.", vbExclamation
    End If
End Sub

Private Sub Tiga_Empat(Jumlah As Integer)

    Dim Var() As Shape
    Dim i As Integer
    
    ReDim Var(1 To 4)
    
    Set Var(1) = ActiveShape
    
    For i = 2 To 4
        Set Var(i) = Var(1).Duplicate
    Next i
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    
    Dim Offset As Integer
    Dim H_Offset As Integer
    Offset = Var(1).SizeWidth + 2
    H_Offset = Var(1).SizeHeight + 2
    
    If Jumlah >= 4 Then
        Var(2).Move Offset, 0
        Var(3).Move (Offset * 2), 0
        Var(4).Move (Offset * 3), 0
    End If
    
    ActiveDocument.ClearSelection
    ActiveDocument.AddToSelection Var()
    
    Dim Varall As Shape
    Set Varall = ActiveSelection.Group
    
    Varall.AlignAndDistribute 3, 3, 2, 0, False, 2
End Sub

Private Sub Empat_Enam(Jumlah As Integer)

    Dim Var() As Shape
    Dim i As Integer
    
    ReDim Var(1 To 3)
    
    Set Var(1) = ActiveShape
    
    For i = 2 To 3
        Set Var(i) = Var(1).Duplicate
    Next i
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    
    Dim Offset As Integer
    Dim H_Offset As Integer
    Offset = Var(1).SizeWidth + 2
    H_Offset = Var(1).SizeHeight + 2
    
    If Jumlah >= 3 Then
        Var(2).Move Offset, 0
        Var(3).Move (Offset * 2), 0
    End If
    
    ActiveDocument.ClearSelection
    ActiveDocument.AddToSelection Var()
    
    Dim Varall As Shape
    Set Varall = ActiveSelection.Group
    
    Varall.AlignAndDistribute 3, 3, 2, 0, False, 2
    
    If Jumlah >= 6 Then
    
        Dim Varall6 As Shape
        Set Varall6 = Varall.Duplicate
        
        Varall6.Move 0, H_Offset
        
        If Jumlah = 9 Then
            
            Dim Varall9 As Shape
            Set Varall9 = Varall.Duplicate
            
            Varall9.Move 0, (H_Offset * 2)
            
            ActiveDocument.ClearSelection
            ActiveDocument.AddToSelection Varall, Varall6, Varall9
            
            Dim Var9all As Shape
            Set Var9all = ActiveSelection.Group
        
            Var9all.AlignAndDistribute 3, 3, 2, 0, False, 2
            
            Exit Sub
        
        End If
        
        ActiveDocument.ClearSelection
        ActiveDocument.AddToSelection Varall, Varall6
        
        Dim Var6all As Shape
        Set Var6all = ActiveSelection.Group
        
        Var6all.AlignAndDistribute 3, 3, 2, 0, False, 2
        
    End If
End Sub
