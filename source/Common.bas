Attribute VB_Name = "Common"
Option Explicit

'===============================================================================

Private Const TRACEABLE_BITMAP_SIZE As Long = 1000
Private Const MM_PER_INCH = 25.4

'===============================================================================
' # Contour helpers

Public Sub ContourMain(ByVal BitmapShape As Shape, ByVal Cfg As Dictionary)
    Dim Offset As Double: Offset = Cfg!Offset
    Dim FilletAmount As Double: FilletAmount = _
        Abs(Offset) * CONTOUR_FILLET_MULT
    If FilletAmount = 0 Then _
        FilletAmount = AverageDim(BitmapShape) * CONTOUR_ZERO_FILLET_MULT
     
    Dim Traced As Shape
    Set Traced = MakeTrace(BitmapShape)
    
    'Debug
    'Traced.Fill.ApplyUniformFill CreateCMYKColor(100, 0, 0, 0)
    
    Dim Contour As Shape
    Set Contour = MakeContour(Traced, Offset)
    Contour.Outline.Color = CreateColor(CONTOUR_COLOR)
    Contour.Fill.ApplyNoFill

    Smoothen Contour.Curve, FilletAmount
End Sub

Public Function MakeTrace(ByVal BitmapShape As Shape) As Shape
    
    Dim ShapeToProcess As Shape: Set ShapeToProcess = BitmapShape.Duplicate
                
    If ShapeToProcess.Bitmap.SizeWidth > TRACEABLE_BITMAP_SIZE _
    Or ShapeToProcess.Bitmap.SizeHeight > TRACEABLE_BITMAP_SIZE Then
        Set ShapeToProcess = ShapeToProcess.ConvertToBitmapEx( _
            ShapeToProcess.Bitmap.Mode, , _
            ShapeToProcess.Bitmap.Transparent, _
            TRACEABLE_BITMAP_SIZE / (GreaterDim(ShapeToProcess) / MM_PER_INCH) _
        )
    End If
    
    'тут мы слегка раздуваем битмап, чтобы нормально трассировался
    'если будет впритык к краю - будут куски фона по углам
    With ShapeToProcess.Bitmap.CropEnvelope
        .CopyAssign .Contour(AverageDim(BitmapShape) / 10, cdrContourOutside)
    End With
    
    Set MakeTrace = TraceBitmap(ShapeToProcess)

    ShapeToProcess.Delete
    
End Function

Public Function MakeContour(ByVal Shape As Shape, ByVal Offset As Double) As Shape
    Dim Direction As cdrContourDirection
    If Offset > 0 Then
        Direction = cdrContourOutside
    ElseIf Offset < 0 Then
        Direction = cdrContourInside
    ElseIf Offset = 0 Then
        Set MakeContour = CreateBoundary(Shape)
        Shape.Delete
        Exit Function
    End If
    
    With Shape.Curve
        .CopyAssign .Contour( _
            Direction:=Direction, _
            Offset:=Abs(Offset), _
            EndCapType:=cdrContourRoundCap, _
            CornerType:=cdrContourCornerRound _
        )
    End With
    Set MakeContour = Shape
End Function

Public Property Get ValidForTrace(ByVal RawShape As Shape) As Boolean
    If RawShape.Type = cdrBitmapShape Then ValidForTrace = True
End Property

Private Function TraceBitmap(ByVal BitmapShape As Shape) As Shape
    Dim TraceResult As ShapeRange
    With BitmapShape.Bitmap.Trace(cdrTraceLineArt)
        .BackgroundRemovalMode = cdrTraceBackgroundAutomatic
        '.CornerSmoothness = 50
        '.DetailLevelPercent = 25
        '.MergeAdjacentObjects = True
        '.SetColorCount 8
        '.MergeAdjacentObjects = True
        .RemoveBackground = True
        '.RemoveEntireBackColor = False
        Set TraceResult = .Finish
    End With
    Set TraceBitmap = CreateBoundary(TraceResult)
    TraceResult.Delete
End Function

Private Sub Smoothen( _
                ByVal Curve As Curve, _
                Optional ByVal FilletAmount As Double = 1 _
            )
    With Curve.Nodes.All
        .Smoothen 1
        .AutoReduce 1
        .Fillet FilletAmount
        .AutoReduce 1
        .Fillet FilletAmount / 2
        .AutoReduce 1
    End With
End Sub

'===============================================================================
' # View and config

Public Function ShowContourView(ByRef Cfg As Dictionary) As BooleanResult
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
' # Common helpers
