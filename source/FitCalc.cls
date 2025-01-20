VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FitCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
'   Модуль          : FitCalc
'   Версия          : 2025.01.17
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Сайт            : https://github.com/elvin-nsk/LowCoupledFromCore
'   Назначение      : считает раскладку, как лучше: с поворотом или без
'   Зависимости     : Size, LibCore
'===============================================================================

'@Interface
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Type This
    NumHeight As Long
    NumWidth As Long
    SheetSize As Size
    PlaceSize As Size
    HSpace As Double
    VSpace As Double
    Rotate As Boolean
    Total As Long
End Type
Private This As This

'===============================================================================
' # Constructors

Friend Function New_( _
                    ByVal SheetSize As Size, _
                    ByVal PlaceSize As Size, _
                    Optional ByVal HSpace As Double, _
                    Optional ByVal VSpace As Double _
                ) As FitCalc
    Set New_ = New FitCalc
    New_.Inject SheetSize, PlaceSize, HSpace, VSpace
End Function

Friend Sub Inject( _
               ByVal SheetSize As Size, _
               ByVal PlaceSize As Size, _
               ByVal HSpace As Double, _
               ByVal VSpace As Double _
           )
    With This
        Set .SheetSize = SheetSize
        Set .PlaceSize = PlaceSize
        .HSpace = HSpace
        .VSpace = VSpace
    End With
    If PlaceSize.Width <> 0 And PlaceSize.Height <> 0 Then
        Calc
    End If
End Sub

'===============================================================================
' # Logic

Private Sub Calc()
    Dim AsIs As Long
    Dim Rotated As Long
    Dim AsIsNumWidth As Long, AsIsNumHeight As Long
    Dim RotatedNumWidth As Long, RotatedNumHeight As Long
    
    With This
        AsIsNumWidth = _
            Fix((.SheetSize.Width + .HSpace) / (.PlaceSize.Width + .HSpace))
        AsIsNumHeight = _
            Fix((.SheetSize.Height + .VSpace) / (.PlaceSize.Height + .VSpace))
        AsIs = AsIsNumWidth * AsIsNumHeight
        
        RotatedNumWidth = _
            Fix((.SheetSize.Width + .HSpace) / (.PlaceSize.Height + .HSpace))
        RotatedNumHeight = _
            Fix((.SheetSize.Height + .VSpace) / (.PlaceSize.Width + .VSpace))
        Rotated = RotatedNumWidth * RotatedNumHeight
    
        If AsIs < Rotated Then
            .Rotate = True
            .NumWidth = RotatedNumWidth
            .NumHeight = RotatedNumHeight
            .Total = Rotated
        Else
            .Rotate = False
            .NumWidth = AsIsNumWidth
            .NumHeight = AsIsNumHeight
            .Total = AsIs
        End If
    End With
End Sub

'===============================================================================
' # Public

Public Property Get NumHeight() As Long
    NumHeight = This.NumHeight
End Property

Public Property Get NumWidth() As Long
    NumWidth = This.NumWidth
End Property

Public Property Get PlaceSize() As Size
    If This.Rotate Then
        Set PlaceSize = This.PlaceSize.Swap
    Else
        Set PlaceSize = This.PlaceSize
    End If
End Property

Public Property Get Rotate() As Boolean
    Rotate = This.Rotate
End Property

Public Property Get Total() As Long
    Total = This.Total
End Property

Public Property Get Self() As FitCalc
    Set Self = Me
End Property

'===============================================================================
' # Helpers
