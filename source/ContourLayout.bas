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
Public Const CONTOUR_FILLET_MULT As Double = 1
Public Const CONTOUR_ZERO_FILLET_MULT As Double = 0.005
Private Const SOME_CONST As String = ""

'===============================================================================
' # Entry points

Sub Contour()

    Const ENTRY_NAME = "Построение контура"

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
        
    Dim Shape As Shape
    With InputData.ExpectPage
        If .Fail Then Exit Sub
        If .Shapes.Count > 1 Then
            Warn "Нужен один растровый объект на странице.", ENTRY_NAME
            Exit Sub
        End If
        Set Shape = .Shape
    End With
    
    If Not ValidForTrace(Shape) Then
        Warn "Объект должен быть растровым изображением.", ENTRY_NAME
        Exit Sub
    End If
    
    Dim Cfg As Dictionary
    If ShowContourView(Cfg) = Fail Then Exit Sub
    
    BoostStart ENTRY_NAME
    
    ContourMain Shape, Cfg
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

Sub Layout()

    Const ENTRY_NAME = "Расклад на лист"

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
        
    If InputData.ExpectPage.Fail Then GoTo Finally
    
    BoostStart ENTRY_NAME
    
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





'===============================================================================
' # Tests

Private Sub TestSomething()
'
End Sub
