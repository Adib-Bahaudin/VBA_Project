VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_2x3 
   Caption         =   "Cetak Pas Foto 2x3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3585
   OleObjectBlob   =   "form_2x3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_2x3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsCancelled As Boolean
Public Jumlah As Integer

Private Sub UserForm_Initialize()

    ComboBox1.Clear
    ComboBox1.AddItem 6
    ComboBox1.AddItem 12
    ComboBox1.AddItem 18
    ComboBox1.AddItem 24
    ComboBox1.AddItem 30
    ComboBox1.AddItem 36
    
    ComboBox1.ListIndex = 0

    CommandButton1.Default = True
    
    IsCancelled = True
    
End Sub

Private Sub CommandButton1_Click()
    
    If CheckBox1.Value = True Then
        Jumlah = 36
    Else
        Dim Inp As Integer
        Inp = CInt(ComboBox1.Text)
        If Inp <= 6 Then
            Jumlah = 6
        ElseIf Inp > 6 And Inp <= 12 Then
            Jumlah = 12
        ElseIf Inp > 12 And Inp <= 18 Then
            Jumlah = 18
        ElseIf Inp > 18 And Inp <= 24 Then
            Jumlah = 24
        ElseIf Inp > 24 And Inp <= 30 Then
            Jumlah = 30
        ElseIf Inp > 30 And Inp <= 36 Then
            Jumlah = 36
        Else
            MsgBox "Error : Jumlah Tidak Valid.", vbExclamation
        End If
    End If
        
    IsCancelled = False
    Unload Me
    
End Sub

