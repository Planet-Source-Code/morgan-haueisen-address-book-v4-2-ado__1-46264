VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/* Copyright (c) 1997-2001
'/*************************************/
Option Explicit

'/* Windows API/Global Declarations for : Determine the Available Screen Real Estate
Private Type Rect
     left As Long
     top As Long
     right As Long
     bottom As Long
End Type

Private Const SPI_GETWORKAREA& = 48
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

'/* Windows API/Global Declarations for IsScreenFontSmall
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'/* Windows API/Global Declarations for Resizing
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Private Const SSTAB_DIALOG_OFFSET As Long = 75000
Private Const SSTAB_TYPE_NAME As String = "SSTAB"
Private Const ResizeRepositionCommand As String = "@"
Private Const ResizeRepositionLeft As String = "L"
Private Const ResizeRepositionTop As String = "T"
Private Const ResizeRepositionWidth As String = "W"
Private Const ResizeRepositionHeight As String = "H"

Private Type ControlPosition
  ControlInstance As Control                         'Reference to the control instance.
  OriginalLeft As Long                               'Original Left position of the Control.
  OriginalTop As Long                                'Original Top position of the Control.
  OriginalWidth As Long                              'Original Width of the Control.
  OriginalHeight As Long                             'Original Height of the Control.
End Type

Private m_SourceForm As Form                         'The form to be resized.
Private m_FormWidth As Long                          'Original form width.
Private m_FormHeight As Long                         'Original form height.
Private m_Controls() As ControlPosition              'Array for storing control information.
Private m_IsFirstResize As Boolean                   'Flag indicating first resize.

Public Enum OnTopPar
    ON_TOP_YES = False
    ON_TOP_NO = True
End Enum


Public Sub ResizeFonts(xForm As Form)
  Const DESIGNFONTSIZE As Integer = 96 '/* 96 = Small & 120 = Large
  Dim i As Integer
  Dim FontRatio As Single
    
    '/* check to see what size of fonts are being used
    FontRatio = DESIGNFONTSIZE / ScreenFontSize()
    If FontRatio = 1 Then Exit Sub

    On Local Error Resume Next
    
    For i = 0 To xForm.Controls.Count - 1 '/* Number of controls on the form
    
        If xForm.Controls(i).Font.Size <= 8 And FontRatio < 1 Then
            xForm.Controls(i).Font.Name = "Small Fonts"
            xForm.Controls(i).Font.Size = xForm.Controls(i).Font.Size - 0.5
        Else
            xForm.Controls(i).Font.Size = xForm.Controls(i).Font.Size * FontRatio
        End If
    
        If TypeOf xForm.Controls(i) Is Line Then
            xForm.Controls(i).X1 = xForm.Controls(i).X1 * FontRatio
            xForm.Controls(i).Y1 = xForm.Controls(i).Y1 * FontRatio
            xForm.Controls(i).X2 = xForm.Controls(i).X2 * FontRatio
            xForm.Controls(i).Y2 = xForm.Controls(i).Y2 * FontRatio
        ElseIf TypeOf xForm.Controls(i) Is PictureBox Then
            xForm.Controls(i).Width = xForm.Controls(i).Width * FontRatio
            xForm.Controls(i).Height = xForm.Controls(i).Height * FontRatio
            xForm.Controls(i).top = xForm.Controls(i).top * FontRatio
            xForm.Controls(i).left = xForm.Controls(i).left * FontRatio
            xForm.Controls(i).ScaleHeight = xForm.Controls(i).ScaleHeight * FontRatio
            xForm.Controls(i).ScaleWidth = xForm.Controls(i).ScaleWidth * FontRatio
        Else
            xForm.Controls(i).Width = xForm.Controls(i).Width * FontRatio
            xForm.Controls(i).Height = xForm.Controls(i).Height * FontRatio
            xForm.Controls(i).top = xForm.Controls(i).top * FontRatio
            xForm.Controls(i).left = xForm.Controls(i).left * FontRatio
            '/* If control is a grid
            xForm.Controls(i).ColWidth = xForm.Controls(i).ColWidth * FontRatio
            xForm.Controls(i).RowHeight = xForm.Controls(i).RowHeight * FontRatio
            '/* If Control is a Toolbar
            xForm.Controls(i).ButtonHeight = xForm.Controls(i).ButtonHeight * FontRatio
            xForm.Controls(i).ButtonWidth = xForm.Controls(i).ButtonWidth * FontRatio
        End If
    Next i
    On Local Error GoTo 0

End Sub

Private Function ScreenFontSize() As Integer
  Dim hWndDesk As Long
  Dim hDCDesk As Long
  Dim logPix As Long
  Dim r As Long
  Const LOGPIXELSX = 88
  'Const LOGPIXELSY = 90
    
    hWndDesk = GetDesktopWindow()
    hDCDesk = GetWindowDC(hWndDesk)
    logPix = GetDeviceCaps(hDCDesk, LOGPIXELSX)
    r = ReleaseDC(hWndDesk, hDCDesk)
    '/* 96 = Small & 120 = Large
    ScreenFontSize = logPix
End Function

Public Sub ResizeSaveInitialSizes(xForm As Form)
    Set m_SourceForm = xForm
    m_IsFirstResize = True
    Call ResizeSaveInitialStates
End Sub
Public Sub ResizeControls(xForm As Form)
  Dim HeightChange As Double, WidthChange As Double
  Dim ControlInstance As Control
  Dim ResizeRepositionFlags As String
  Dim i As Integer

    Set m_SourceForm = xForm
    
    '/* Check if the form has any controls on it.
    If m_SourceForm.Controls.Count = 0 Then Exit Sub
    
    '/* Check if the form is being minimized.
    If m_SourceForm.WindowState = vbMinimized Then Exit Sub
    
    '/* Check if it's the first time the form is to be resized.
    If m_IsFirstResize Then
        Call ResizeSaveInitialStates
        Exit Sub
    End If

    '/* Prevent the Window from redrawing until the user releases the mouse.
    Call LockWindowUpdate(m_SourceForm.hwnd)
    
    '/* Calculate the change in the form's size.
    If (TypeOf m_SourceForm Is MDIForm) Then
        HeightChange = m_SourceForm.Height / m_FormHeight
        WidthChange = m_SourceForm.Width / m_FormWidth
    Else
        HeightChange = m_SourceForm.ScaleHeight / m_FormHeight
        WidthChange = m_SourceForm.ScaleWidth / m_FormWidth
    End If

    '/* Iterate through the array containing the form's controls.
    For i = 0 To UBound(m_Controls)
        Set ControlInstance = m_Controls(i).ControlInstance
    
        '/* Use the Control that is currently being referenced.
        With ControlInstance
            On Local Error Resume Next
            '/* Check if the Tag Property contains valid resize or reposition instructions.
            If (left(UCase(.Tag), 1) = ResizeRepositionCommand) Then
                'Retrieve any specific resize and reposition instructions.
                ResizeRepositionFlags = UCase(.Tag)
            Else
                '/* Default all resize and resposition instructions.
                ResizeRepositionFlags = ResizeRepositionCommand & ResizeRepositionLeft & ResizeRepositionTop & ResizeRepositionWidth & ResizeRepositionHeight
            End If
            
            '/* Resize and Reposition the control.
            If (TypeOf ControlInstance Is Line) Then
                '/* Line control.
                If (InStr(ResizeRepositionFlags, ResizeRepositionLeft) <> 0) Then .X1 = (m_Controls(i).OriginalLeft * WidthChange)
                If (InStr(ResizeRepositionFlags, ResizeRepositionTop) <> 0) Then .Y1 = (m_Controls(i).OriginalTop * HeightChange)
                If (InStr(ResizeRepositionFlags, ResizeRepositionWidth) <> 0) Then .X2 = (m_Controls(i).OriginalWidth * WidthChange)
                If (InStr(ResizeRepositionFlags, ResizeRepositionHeight) <> 0) Then .Y2 = (m_Controls(i).OriginalHeight * HeightChange)
            ElseIf (UCase(TypeName(ControlInstance.Container)) = SSTAB_TYPE_NAME) Then
                '/* Control is placed on a SSTab Dialog Control.
                If (.left < 0) Then
                    If (m_Controls(i).OriginalLeft > 0) Then
                        m_Controls(i).OriginalLeft = m_Controls(i).OriginalLeft - SSTAB_DIALOG_OFFSET
                    End If
            
                    If (InStr(ResizeRepositionFlags, ResizeRepositionLeft) <> 0) Then .left = ((m_Controls(i).OriginalLeft + SSTAB_DIALOG_OFFSET) * WidthChange) - SSTAB_DIALOG_OFFSET
                Else
                    If (m_Controls(i).OriginalLeft < 0) Then
                        m_Controls(i).OriginalLeft = m_Controls(i).OriginalLeft + SSTAB_DIALOG_OFFSET
                    End If
            
                    If (InStr(ResizeRepositionFlags, ResizeRepositionLeft) <> 0) Then .left = m_Controls(i).OriginalLeft * WidthChange
                End If
        
                If (InStr(ResizeRepositionFlags, ResizeRepositionTop) <> 0) Then .top = (m_Controls(i).OriginalTop * HeightChange)
                If (InStr(ResizeRepositionFlags, ResizeRepositionWidth) <> 0) Then .Width = (m_Controls(i).OriginalWidth * WidthChange)
                If (InStr(ResizeRepositionFlags, ResizeRepositionHeight) <> 0) Then .Height = (m_Controls(i).OriginalHeight * HeightChange)
            Else
                '/* All other controls.
                If (InStr(ResizeRepositionFlags, ResizeRepositionLeft) <> 0) Then .left = (m_Controls(i).OriginalLeft * WidthChange)
                If (InStr(ResizeRepositionFlags, ResizeRepositionTop) <> 0) Then .top = (m_Controls(i).OriginalTop * HeightChange)
                If (InStr(ResizeRepositionFlags, ResizeRepositionWidth) <> 0) Then .Width = (m_Controls(i).OriginalWidth * WidthChange)
                If (InStr(ResizeRepositionFlags, ResizeRepositionHeight) <> 0) Then .Height = (m_Controls(i).OriginalHeight * HeightChange)
            End If
            On Local Error GoTo 0
        End With
    Next i
    
    '/* Release the memory being used by the ControlInstance variable.
    Set ControlInstance = Nothing
    
    '/* Allow Windows to refresh the form's display.
    Call LockWindowUpdate(0)
End Sub
Private Sub ResizeSaveInitialStates()
  '/* Description : Store the dimensions and positions of all controls on the Form.            '
  Dim i As Integer

    '/* Check if the form has any controls on it.
    If m_SourceForm.Controls.Count = 0 Then Exit Sub

    With m_SourceForm
        '/* Check if the form being resized is an MDI form.
        If (TypeOf m_SourceForm Is MDIForm) Then
            '/* Set the FormWidth and FormHeight variables to the Form's Width and Height.
            m_FormWidth = .Width
            m_FormHeight = .Height
        Else
            '/* Set the FormWidth and FormHeight variables to the Form's Scale Width and Height.
            m_FormWidth = .ScaleWidth
            m_FormHeight = .ScaleHeight
        End If
    End With

    '/* Frees the memory used by the dynamic array storing all control information.
    Erase m_Controls
    
    '/* Redimension the array storing all control information to the number of controls on the
    '/* form that is being resized.
    ReDim m_Controls(0 To (m_SourceForm.Controls.Count - 1))

    '/* Iterate through each of the Controls on the form to be resized.
    For i = 0 To (m_SourceForm.Controls.Count - 1)
        '/* Use the control that is currently being referenced.
        With m_SourceForm.Controls(i)
            '/* Add the controls reference to the array containing all control information.
            Set m_Controls(i).ControlInstance = m_SourceForm.Controls(i)
            '/* Check if the control currently being referenced is a line control.
            If (TypeOf m_SourceForm.Controls(i) Is Line) Then
                '/* Add a reference to the control and it's position into the Controls array.
                m_Controls(i).OriginalLeft = .X1
                m_Controls(i).OriginalTop = .Y1
                m_Controls(i).OriginalWidth = .X2
                m_Controls(i).OriginalHeight = .Y2
            Else
                On Local Error Resume Next
                '/* Add a reference to the control and it's position into the Controls array.
                m_Controls(i).OriginalLeft = .left
                m_Controls(i).OriginalTop = .top
                m_Controls(i).OriginalWidth = .Width
                m_Controls(i).OriginalHeight = .Height
                On Local Error GoTo 0
            End If
        End With
    Next i

    '/* Notify the class that the form has been resized.
    m_IsFirstResize = False
    
End Sub


Public Sub CenterObject(ParentObject As Object, ChildObject As Object)
    ChildObject.top = (ParentObject.Height - ChildObject.Height) / 2
    ChildObject.left = (ParentObject.Width - ChildObject.Width) / 2
End Sub

Public Function scrnTop() As Long
     Dim Rc As Rect
     Dim r As Long
     Dim msg As String
     
     r = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
     scrnTop = Rc.top * Screen.TwipsPerPixelY
End Function

Public Function scrnRight() As Long
     Dim Rc As Rect
     Dim r As Long
     Dim msg As String
     
     r = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
     scrnRight = Rc.right * Screen.TwipsPerPixelX
End Function
Public Function scrnBottom() As Long
     Dim Rc As Rect
     Dim r As Long
     Dim msg As String
     
     r = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
     scrnBottom = Rc.bottom * Screen.TwipsPerPixelY
End Function

Public Function scrnLeft() As Long
     Dim Rc As Rect
     Dim r As Long
     Dim msg As String
     
     r = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
     scrnLeft = Rc.left * Screen.TwipsPerPixelX
End Function


Public Sub CenterForm(xForm As Form, Optional ByVal MoveUpDn As Single = 1)
  Dim Rc As Rect
  Dim rVal As Long
  Dim T As Long, B As Long, L As Long, r As Long
  Dim mT As Long, mL As Long
     
    rVal = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
    
    T = Rc.top * Screen.TwipsPerPixelY
    B = Rc.bottom * Screen.TwipsPerPixelY
    L = Rc.left * Screen.TwipsPerPixelX
    r = Rc.right * Screen.TwipsPerPixelX
    
    mT = Abs((B / 2 * MoveUpDn) - (xForm.Height / 2))
    mL = Abs((r / 2) - (xForm.Width / 2))
    
    If mT < T Then mT = T
    If mT > B - xForm.Height Then mT = B - xForm.Height
    If mL < L Then mL = L
    
    xForm.Move mL, mT

End Sub
Public Sub HCenterObject(Obj As Object)
     Dim Rc As Rect
     Dim r As Long, L As Long
     
     r = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
     
     L = Rc.right * Screen.TwipsPerPixelX
     L = L / 2 - Obj.Width / 2
     Obj.left = L

End Sub

Public Sub FitScreen(xForm As Form, Optional ByVal Percent As Single = 100)
 Dim Rc As Rect
 Dim r As Long
    
    On Local Error Resume Next
    r = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
    'left - " & rc.Left
    'top - " & rc.Top
    'rght - " & rc.Right
    'bttm - " & rc.Bottom
    
    If Percent > 100 Then Percent = 100
    Percent = Percent / 100
    
    xForm.Move Rc.left * Screen.TwipsPerPixelX * Percent, Rc.top * Screen.TwipsPerPixelY * Percent, _
            Rc.right * Screen.TwipsPerPixelX * Percent, Rc.bottom * Screen.TwipsPerPixelY * Percent
            
    If Percent < 1 Then Call CenterForm(xForm)
    
    On Local Error GoTo 0
End Sub

Public Sub OnTop(xForm As Form, Optional TurnOff As OnTopPar = ON_TOP_YES)
    If TurnOff Then
        Call SetWindowPos(xForm.hwnd, -2, 0, 0, 0, 0, 3)
    Else
        Call SetWindowPos(xForm.hwnd, -1, 0, 0, 0, 0, 3)
    End If
End Sub

Public Sub vFitScreen(xForm As Form, Optional Center As Boolean = True, Optional UseFullScreen As Boolean = False)
     Dim Rc As Rect
     Dim r As Long, scTop As Long, scBottom As Long
     
     r = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
     
     scTop = Rc.top * Screen.TwipsPerPixelY
     scBottom = Rc.bottom * Screen.TwipsPerPixelY
     
     If scTop = 0 And Not UseFullScreen Then scTop = 250
     xForm.Height = scBottom - scTop
     If Center Then CenterForm xForm
     xForm.top = scTop
    
End Sub
Public Sub hFitScreen(xForm As Form, Optional Center As Boolean = True)
     Dim Rc As Rect
     Dim r As Long
     
     r = SystemParametersInfo(SPI_GETWORKAREA, 0&, Rc, 0&)
     xForm.left = Rc.left * Screen.TwipsPerPixelX
     xForm.Width = (Rc.right - Rc.left) * Screen.TwipsPerPixelX
     If Center Then CenterForm xForm

End Sub

Private Sub Class_Initialize()
    Set m_SourceForm = Nothing
    m_IsFirstResize = True
End Sub

Public Property Get Width() As Variant
    Width = GetSystemMetrics(0&)
End Property

Public Property Get Height() As Variant
    Height = GetSystemMetrics(1&)
End Property

