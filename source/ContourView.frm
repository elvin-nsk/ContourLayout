VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContourView 
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8625.001
   OleObjectBlob   =   "ContourView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContourView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'===============================================================================
' # State

Private Const MIN_OFFSET As Double = 0.025

Public IsOk As Boolean
Public IsCancel As Boolean

Public OffsetHandler As TextBoxHandler

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = "Построение контура - " & APP_DISPLAYNAME & " (v" & APP_VERSION & ")"
    btnOk.Default = True
    
    Set OffsetHandler = _
        TextBoxHandler.New_(Offset, TextBoxTypeDouble)
End Sub

'===============================================================================
' # Handlers

Private Sub Offset_AfterUpdate()
    If Offset.Value < MIN_OFFSET And Offset.Value > -MIN_OFFSET Then
        If Offset.Value < 0 Then
            Offset.Value = VBA.CStr(-MIN_OFFSET)
        Else
            Offset.Value = VBA.CStr(MIN_OFFSET)
        End If
    End If
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub btnOk_Click()
    FormОК
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================
' # Logic

Private Sub FormОК()
    Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================
' # Helpers


'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Сancel = True
        FormCancel
    End If
End Sub
