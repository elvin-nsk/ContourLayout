VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LayoutView 
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5745
   OleObjectBlob   =   "LayoutView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LayoutView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # State

Private Type This
    NumWidth As TextBoxHandler
    NumHeight As TextBoxHandler
    Space As TextBoxHandler
    MarksInnerOffset As TextBoxHandler
    MarksPath As TextBoxHandler
End Type
Private This As This

Public IsOk As Boolean
Public IsCancel As Boolean
Public Calculated As FitCalc

Public PageSize As Size
Public PlaceSize As Size

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = "Расклад на лист - " & APP_DISPLAYNAME & " v" & APP_VERSION
    btnOk.Default = True
    
    With This
        Set .NumWidth = _
            TextBoxHandler.New_(NumWidth, TextBoxTypeLong, 1)
        Set .NumHeight = _
            TextBoxHandler.New_(NumHeight, TextBoxTypeLong, 1)
        Set .Space = _
            TextBoxHandler.New_(Space, TextBoxTypeDouble)
        Set .MarksInnerOffset = _
            TextBoxHandler.New_(MarksInnerOffset, TextBoxTypeLong)
        Set .MarksPath = _
            TextBoxHandler.New_(MarksPath, TextBoxTypeString)
            
        .MarksInnerOffset = 10
    End With
End Sub

'===============================================================================
' # Handlers

Private Sub UserForm_Activate()
    CalcLayoutAndValidate
End Sub

Private Sub Recalculate_Click()
    CalcLayoutAndValidate
End Sub

Private Sub NumHeight_Change()
    CalcInfo
End Sub

Private Sub NumWidth_Change()
    CalcInfo
End Sub

Private Sub Rotate_Change()
    CalcInfo
End Sub

Private Sub Space_Change()
    CalcInfo
End Sub

Private Sub BrowseMarksPath_Click()
    Dim File As String
    If AskForMarksFile(FileSpec.New_(This.MarksPath).Path, File) = Ok Then
        This.MarksPath = File
    End If
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

Private Sub CalcLayoutAndValidate()
    DoEvents
    Dim Sheet As Size: Set Sheet = CalcSheetSize
    Set Calculated = FitCalc.New_(Sheet, PlaceSize)
    Dim HSpace As Double, VSpace As Double
    With This
        .NumWidth = Calculated.NumWidth
        .NumHeight = Calculated.NumHeight
        Rotate = Calculated.Rotate
        If .NumWidth <= 1 Or .NumHeight <= 1 Then
            .Space = 0
            Exit Sub
        End If
        HSpace = _
            (Sheet.Width - Calculated.PlaceSize.Width * Calculated.NumWidth) _
          / (Calculated.NumWidth - 1)
        VSpace = _
            (Sheet.Height - Calculated.PlaceSize.Height * Calculated.NumHeight) _
          / (Calculated.NumHeight - 1)
        .Space = Fix(MinOfTwo(HSpace, VSpace))
    End With
End Sub

Private Sub CalcInfo()
    DoEvents
    Dim Sheet As Size: Set Sheet = CalcSheetSize
    SheetSize = Fix(Sheet.Width) & " x " & Fix(Sheet.Height) & " мм"
    Dim Lay As Size: Set Lay = CalcLayoutSize(Sheet)
    LayoutSize = Fix(Lay.Width) + 1 & " x " & Fix(Lay.Height) + 1 & " мм"
    LayoutCount = This.NumWidth * This.NumHeight
    If This.NumWidth = 0 Or This.NumHeight = 0 Then
        btnOk.Enabled = False
        Exit Sub
    Else
        btnOk.Enabled = True
    End If
End Sub

Private Property Get CalcSheetSize() As Size
    Dim MarksOffset As Double
    If OptionMarks Then MarksOffset = This.MarksInnerOffset * 2
    Set CalcSheetSize = _
        Size.New_( _
            Width:= _
                PageSize.Width _
              - PAGE_PADDING_LEFT - PAGE_PADDING_RIGHT - MarksOffset, _
            Height:= _
                PageSize.Height _
              - PAGE_PADDING_TOP - PAGE_PADDING_BOTTOM - MarksOffset _
        )
End Property

Private Property Get CalcLayoutSize(ByVal SheetSize As Size) As Size
    Dim Place As Size
    If Rotate Then Set Place = PlaceSize.Swap Else Set Place = PlaceSize
    Set CalcLayoutSize = _
        Size.New_( _
            Width:=((Place.Width + This.Space) * This.NumWidth) - This.Space, _
            Height:=((Place.Height + This.Space) * This.NumHeight) - This.Space _
        )
End Property

Public Function AskForMarksFile( _
                    ByVal InitialDir As String, _
                    ByRef File As String _
                ) As BooleanResult
  Dim Files As Collection
  With New FileBrowser
    .Filter = "CorelDraw (*.cdr)" & VBA.Chr(0) & "*.cdr"
    .InitialDir = InitialDir
    .MultiSelect = False
    .Title = "Выберите файл меток"
    Set Files = .ShowFileOpenDialog
  End With
  If Files.Count = 0 Then Exit Function
  File = FileSpec.New_(Files(1))
  AskForMarksFile = Ok
End Function

'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Сancel = True
        FormCancel
    End If
End Sub
