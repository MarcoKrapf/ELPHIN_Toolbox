VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFortschritt 
   Caption         =   "[Fortschrittsbalken]"
   ClientHeight    =   615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4230
   OleObjectBlob   =   "frmFortschritt.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFortschritt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Call modAuxiliary.PlaceUserFormInCenter(Me)
End Sub
