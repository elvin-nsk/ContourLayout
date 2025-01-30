Attribute VB_Name = "ContourLayout"
'===============================================================================
'   Макрос          : ContourLayout
'   Версия          : 2025.01.30
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
Public Const APP_VERSION As String = "2025.01.30"

'===============================================================================
' # Globals

Public Const CONTOUR_THICKNESS As Double = 0.076 '0.076 = hairline
Public Const CONTOUR_COLOR As String = "CMYK,USER,0,0,0,100"
Public Const CONTOUR_FILLET_MULT As Double = 1
Public Const CONTOUR_ZERO_FILLET_MULT As Double = 0.005
Public Const PAGE_PADDING_TOP As Double = 15
Public Const PAGE_PADDING_LEFT As Double = 10
Public Const PAGE_PADDING_RIGHT As Double = PAGE_PADDING_LEFT
Public Const PAGE_PADDING_BOTTOM As Double = 55

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
    If ShiftKeyPressed Then
        ReadContourCfg Cfg
    Else
        If ShowContourView(Cfg) = Fail Then Exit Sub
    End If
    
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
        
    Dim Shapes As ShapeRange
    With InputData.ExpectPage
        If .Fail Then Exit Sub
        Set Shapes = .Shapes
    End With
        
    ActiveDocument.Unit = cdrMillimeter
    Dim PageSize As Size: Set PageSize = Size.NewFromRect(ActivePage.BoundingBox)
    Dim PlaceSize As Size: Set PlaceSize = Size.NewFromRect(Shapes.BoundingBox)
        
    Dim Cfg As Dictionary, LayoutInfo As LayoutInfo
    If ShowLayoutView(PageSize, PlaceSize, Cfg, LayoutInfo) _
        = Fail Then Exit Sub
    
    If LayoutInfo.NumWidth * LayoutInfo.NumHeight = 0 Then Exit Sub
    
    BoostStart ENTRY_NAME
    
    LayoutMain Shapes, Cfg, LayoutInfo
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Tests

Private Sub TestSomething()
End Sub
