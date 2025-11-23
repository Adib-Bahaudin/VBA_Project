VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_4x6 
   Caption         =   "Cetak Pas Foto 4x6"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4155
   OleObjectBlob   =   "form_4x6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_4x6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsCancelled As Boolean
Public Jumlah As Integer

Private Sub UserForm_Initialize()

    OptionButton1 = True
    
    CommandButton2.Cancel = True
    CommandButton1.Default = True
    
    IsCancelled = True
    
End Sub

Private Sub CommandButton1_Click()
    
    If OptionButton1.Value = True Then
        Jumlah = 3
    ElseIf OptionButton2.Value = True Then
        Jumlah = 6
    ElseIf OptionButton3.Value = True Then
        Jumlah = 9
    End If
        
    IsCancelled = False
    Unload Me
    
End Sub

Private Sub CommandButton2_Click()

    IsCancelled = True
    Unload Me
    
End Sub
