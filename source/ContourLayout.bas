Attribute VB_Name = "ContourLayout"
'===============================================================================
'   Макрос          : ContourLayout
'   Версия          : 2025.01.15
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "ContourLayout"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_FILEBASENAME As String = "elvin_" & APP_NAME
Public Const APP_VERSION As String = "2025.01.15"

'===============================================================================
' # Globals

Public Const CONTOUR_COLOR As String = "CMYK,USER,0,0,0,100"
Private Const SOME_CONST As String = ""

'===============================================================================
' # Entry points

Sub Contour()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
        
    If InputData.ExpectPage.Fail Then GoTo Finally
    
    Dim Cfg As Dictionary
    If ShowContourView(Cfg) = Fail Then GoTo Finally
    
    BoostStart "Построение контура"
    
    
    '??? PROFIT!
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

Sub Layout()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
        
    If InputData.ExpectPage.Fail Then GoTo Finally
    
    BoostStart "Layout"
    
    '??? PROFIT!
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub


'===============================================================================
' # Helpers



'-------------------------------------------------------------------------------

Private Function ShowContourView(ByRef Cfg As Dictionary) As BooleanResult
    Dim FileBinder As JsonFileBinder: Set FileBinder = BindConfig
    Set Cfg = FileBinder.GetOrMakeSubDictionary("Contour")
    Dim View As New ContourView
    Dim ViewBinder As ViewToDictionaryBinder: Set ViewBinder = _
        ViewToDictionaryBinder.New_( _
            Dictionary:=Cfg, _
            View:=View, _
            ControlNames:=Pack("Offset") _
        )
    View.Show vbModal
    ViewBinder.RefreshDictionary
    ShowContourView = View.IsOk
End Function

Private Function BindConfig() As JsonFileBinder
    Set BindConfig = JsonFileBinder.New_(APP_FILEBASENAME)
End Function

'===============================================================================
' # Tests

Private Sub TestSomething()
'
End Sub
