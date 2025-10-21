VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm12 
   Caption         =   "Extractor By Ignacio Solar"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   OleObjectBlob   =   "UserForm12.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub boton_Click()
Call er
End Sub

Private Sub UserForm_Initialize()
ComboBox1.AddItem ("Ezentis Energia") 'JULIO YAÑEZ
ComboBox1.AddItem ("Ezentis Chile") 'CARLOS MORENO
End Sub
