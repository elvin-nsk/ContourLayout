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
    If Contour.Type = cdrGroupShape Then _
            Set Contour = Contour.UngroupAllEx.Combine
    Contour.Outline.SetProperties CONTOUR_THICKNESS, , CreateColor(CONTOUR_COLOR)
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
        
    Set MakeContour = _
        Shape.CreateContour( _
            Direction:=Direction, _
            Offset:=Abs(Offset), _
            Steps:=1, _
            EndCapType:=cdrContourRoundCap, _
            CornerType:=cdrContourCornerRound _
        ).Separate.FirstShape
    
    Shape.Delete
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
' # Layout helpers

Public Sub LayoutMain( _
               ByVal Shapes As ShapeRange, _
               ByVal Cfg As Dictionary, _
               LayoutInfo As LayoutInfo _
           )
    If LayoutInfo.Rotate Then Shapes.Rotate 90
    Dim Sorted As SortedMotifs: Set Sorted = _
        DuplicateMotifs(Shapes, LayoutInfo.NumWidth * LayoutInfo.NumHeight)
    
    Dim MarksOffset As Double
    If Cfg!OptionMarks Then MarksOffset = Cfg!MarksInnerOffset
    
    With Composer.NewCompose( _
        Elements:=Sorted.Elements, _
        StartingPoint:= _
            Point.New_( _
                PAGE_PADDING_LEFT + MarksOffset, _
                ActivePage.TopY - PAGE_PADDING_TOP - MarksOffset _
            ), _
        MaxPlacesInWidth:=LayoutInfo.NumWidth, _
        MaxPlacesInHeight:=LayoutInfo.NumHeight, _
        HorizontalSpace:=LayoutInfo.Space, _
        VerticalSpace:=LayoutInfo.Space _
    )
    End With
    
    If Not Cfg!OptionMarks Then Exit Sub
    
    Dim Composed As ShapeRange: Set Composed = ActivePage.Shapes.All
    Dim Mark As Shape: Set Mark = Import(Cfg!MarksPath)
    Dim Marks As ShapeRange: Set Marks = SetMarks(Mark, Composed, MarksOffset)
    Mark.Delete
    
    With ActiveDocument
        .MasterPage.SetSize ActivePage.SizeWidth, ActivePage.SizeHeight
    End With
    Dim NewPage As Page
    With Sorted
        If .Contours.Count > 0 And .Images.Count > 0 Then
            Set NewPage = ActiveDocument.AddPages(1)
            .Contours.MoveToLayer NewPage.ActiveLayer
            Marks.CopyToLayer NewPage.ActiveLayer
            With NewPage.Shapes.All.CreateDocumentFrom.ActivePage
                .SetSize .Shapes.All.SizeWidth, .Shapes.All.SizeHeight
                .Shapes.All.SetPositionEx cdrBottomLeft, .LeftX, .BottomY
                
            End With
        End If
    End With
    
End Sub

Private Function SetMarks( _
                     ByVal Mark As Shape, _
                     ByVal Area As ShapeRange, _
                     ByVal MarksInnerOffset As Double _
                 ) As ShapeRange
    Set SetMarks = CreateShapeRange
    SetMarks.Add MarkDuplicate(Mark, Area.LeftX - MarksInnerOffset, Area.TopY + MarksInnerOffset, 0)
    SetMarks.Add MarkDuplicate(Mark, Area.RightX + MarksInnerOffset, Area.TopY + MarksInnerOffset, -90)
    SetMarks.Add MarkDuplicate(Mark, Area.RightX + MarksInnerOffset, Area.BottomY - MarksInnerOffset, -180)
    SetMarks.Add MarkDuplicate(Mark, Area.LeftX - MarksInnerOffset, Area.BottomY - MarksInnerOffset, -270)
End Function

Private Function MarkDuplicate( _
                ByVal Mark As Shape, _
                ByVal x As Double, y As Double, _
                ByVal RotateByAngle As Double _
            ) As Shape
    Set MarkDuplicate = Mark.Duplicate
    SetPositionByRotationCenter MarkDuplicate, x, y
    MarkDuplicate.Rotate RotateByAngle
End Function

Public Property Get DuplicateMotifs( _
                        ByVal Motif As ShapeRange, _
                        ByVal Count As Long _
                    ) As SortedMotifs
    Set DuplicateMotifs = New SortedMotifs
    Dim Image As Shape, Contour As Shape
    GetSortedMotif Motif, Image, Contour
    Dim TempMotif As ShapeRange
    Dim TempShape As Shape
    Dim i As Long
    For i = 1 To Count
        Set TempMotif = CreateShapeRange
        If IsSome(Image) Then
            Set TempShape = Image.Duplicate
            TempMotif.Add TempShape
            DuplicateMotifs.Images.Add TempShape
        End If
        If IsSome(Contour) Then
            Set TempShape = Contour.Duplicate
            TempMotif.Add TempShape
            DuplicateMotifs.Contours.Add TempShape
        End If
        DuplicateMotifs.Elements.Add ComposerElement.New_(TempMotif)
    Next i
    Motif.Delete
End Property

Private Sub GetSortedMotif( _
                ByVal Motif As ShapeRange, _
                ByRef Image As Shape, _
                ByRef Contour As Shape _
            )
    Dim Shapes As New ShapeRange: Shapes.AddRange Motif
    Set Image = Motif.Shapes.FindShape(Type:=cdrBitmapShape)
    If IsSome(Image) Then Shapes.RemoveRange PackShapes(Image)
    If Shapes.Count = 1 Then
        Set Contour = Shapes.FirstShape
    ElseIf Shapes.Count > 1 Then
        Set Contour = Shapes.Group
    End If
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

Public Sub ReadContourCfg(ByRef Cfg As Dictionary)
    Set Cfg = BindConfig.GetOrMakeSubDictionary("Contour")
End Sub

Public Function ShowLayoutView( _
                    ByVal PageSize As Size, _
                    ByVal PlaceSize As Size, _
                    ByRef Cfg As Dictionary, _
                    ByRef LayoutInfo As LayoutInfo _
                ) As BooleanResult
    Dim FileBinder As JsonFileBinder: Set FileBinder = BindConfig
    Set Cfg = FileBinder.GetOrMakeSubDictionary("Layout")
    Dim View As New LayoutView
    Dim ViewBinder As ViewToDictionaryBinder: Set ViewBinder = _
        ViewToDictionaryBinder.New_( _
            Dictionary:=Cfg, _
            View:=View, _
            ControlNames:= _
                Pack("OptionMarks", "MarksInnerOffset", "MarksPath") _
        )
    Set View.PlaceSize = PlaceSize
    Set View.PageSize = PageSize
    View.Show vbModal
    ShowLayoutView = View.IsOk
    If Not View.IsOk Then Exit Function
    
    ViewBinder.RefreshDictionary
    Set LayoutInfo = New LayoutInfo
    LayoutInfo.NumWidth = View.NumWidth
    LayoutInfo.NumHeight = View.NumHeight
    LayoutInfo.Rotate = View.Rotate
    LayoutInfo.Space = View.Space
End Function

Private Function BindConfig() As JsonFileBinder
    Set BindConfig = JsonFileBinder.New_(APP_FILEBASENAME)
End Function

'===============================================================================
' # Common helpers
