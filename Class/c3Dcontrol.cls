VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls3Dcontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub ThreeDControl(DrawObj As Object, Ctrl As Control, Optional BevelWidth% = 1, Optional BoarderWidth% = 0, Optional InsertInner As Boolean = False, Optional InsertOuter As Boolean = False, Optional Outline As Boolean = False)
' BevelWidth% = bevel width (pixels)
' BoarderWidth% = surround distance from control (pixels)
' InsertInner% = True is 3D inset border
' InsertOuter% = True is 3D inset border
     
    On Local Error Resume Next
    PixX% = Screen.TwipsPerPixelX
    PixY% = Screen.TwipsPerPixelY
    CTop% = Ctrl.top - PixX%
    CLft% = Ctrl.left - PixY%
    CRgt% = Ctrl.left + Ctrl.Width
    CBtm% = Ctrl.top + Ctrl.Height
    '     ' Color used below:
    Const Shadow = &H80000010
    Const Highlight = &H80000014
    
    
    If InsertInner Then 'inset border
        For i% = 0 To (BevelWidth% - 1)
           AddX% = i% * PixX%
           AddY% = i% * PixY%
    '            Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CRgt% + AddX%, CTop% - AddY%), Shadow
    '            Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CLft% - AddX%, CBtm% + AddY%), Shadow
    '            Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CRgt% + AddX% + PixX%, CBtm% + AddY%), Highlight
    '            Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CRgt% + AddX%, CBtm% + AddY%), Highlight
           DrawObj.Line (CLft% - AddX%, CTop% - AddY%)-(CRgt% + AddX%, CTop% - AddY%), vb3DShadow
           DrawObj.Line (CLft% - AddX%, CTop% - AddY%)-(CLft% - AddX%, CBtm% + AddY%), vb3DShadow
           DrawObj.Line (CLft% - AddX%, CBtm% + AddY%)-(CRgt% + AddX% + PixX%, CBtm% + AddY%), vb3DHighlight
           DrawObj.Line (CRgt% + AddX%, CTop% - AddY%)-(CRgt% + AddX%, CBtm% + AddY%), vb3DHighlight
        Next
    Else 'outset border
       For i% = 0 To (BevelWidth% - 1)
          AddX% = i% * PixX%
          AddY% = i% * PixY%
    '           Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + AddX%, CTop% - AddY%), Shadow
    '           Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - AddX%, CBtm% + AddY%), Shadow
    '           Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - AddX% - PixX%, CTop% - AddY%), Highlight
    '           Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - AddX%, CTop% - AddY%), Highlight
          DrawObj.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + AddX%, CTop% - AddY%), vb3DShadow
          DrawObj.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - AddX%, CBtm% + AddY%), vb3DShadow
          DrawObj.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - AddX% - PixX%, CTop% - AddY%), vb3DHighlight
          DrawObj.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - AddX%, CTop% - AddY%), vb3DHighlight
       Next
    End If
    
    If BoarderWidth > 0 Then
        If InsertOuter Then 'inset border
            For i% = BoarderWidth% To (BevelWidth% + BoarderWidth% - 1)
               AddX% = i% * PixX%
               AddY% = i% * PixY%
        '            Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CRgt% + AddX%, CTop% - AddY%), Shadow
        '            Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CLft% - AddX%, CBtm% + AddY%), Shadow
        '            Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CRgt% + AddX% + PixX%, CBtm% + AddY%), Highlight
        '            Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CRgt% + AddX%, CBtm% + AddY%), Highlight
               DrawObj.Line (CLft% - AddX%, CTop% - AddY%)-(CRgt% + AddX%, CTop% - AddY%), vb3DShadow
               DrawObj.Line (CLft% - AddX%, CTop% - AddY%)-(CLft% - AddX%, CBtm% + AddY%), vb3DShadow
               DrawObj.Line (CLft% - AddX%, CBtm% + AddY%)-(CRgt% + AddX% + PixX%, CBtm% + AddY%), vb3DHighlight
               DrawObj.Line (CRgt% + AddX%, CTop% - AddY%)-(CRgt% + AddX%, CBtm% + AddY%), vb3DHighlight
            Next
        Else 'outset border
           For i% = BoarderWidth% To (BevelWidth% + BoarderWidth% - 1)
              AddX% = i% * PixX%
              AddY% = i% * PixY%
        '           Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + AddX%, CTop% - AddY%), Shadow
        '           Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - AddX%, CBtm% + AddY%), Shadow
        '           Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - AddX% - PixX%, CTop% - AddY%), Highlight
        '           Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - AddX%, CTop% - AddY%), Highlight
              DrawObj.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + AddX%, CTop% - AddY%), vb3DShadow
              DrawObj.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - AddX%, CBtm% + AddY%), vb3DShadow
              DrawObj.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - AddX% - PixX%, CTop% - AddY%), vb3DHighlight
              DrawObj.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - AddX%, CTop% - AddY%), vb3DHighlight
           Next
        End If
    End If
    
    If Outline Then
        If BoarderWidth > 0 Then
            i% = BevelWidth% + BoarderWidth%
            AddX% = i% * PixX%
            AddY% = i% * PixY%
            DrawObj.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + AddX%, CTop% - AddY%), vb3DDKShadow
            DrawObj.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - AddX%, CBtm% + AddY%), vb3DDKShadow
            DrawObj.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - AddX% - PixX%, CTop% - AddY%), vb3DDKShadow
            DrawObj.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - AddX%, CTop% - AddY%), vb3DDKShadow
        Else
            AddX% = BevelWidth% * PixX%
            AddY% = BevelWidth% * PixY%
            DrawObj.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + AddX%, CTop% - AddY%), vb3DDKShadow
            DrawObj.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - AddX%, CBtm% + AddY%), vb3DDKShadow
            DrawObj.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - AddX% - PixX%, CTop% - AddY%), vb3DDKShadow
            DrawObj.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - AddX%, CTop% - AddY%), vb3DDKShadow
        End If
    End If
            
    On Local Error GoTo 0

End Sub


