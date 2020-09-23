VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/* Copyright (c) 1997-2002
'/*************************************/
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Sub FcaseTextBox(MyCTRL As Object)
    '/**************************************/
    '/* Change to formal case while typing */
    '/* Place in MyCTRL_KeyUp event        */
    '/**************************************/
    Dim x As Integer
    Dim tString As String
    
    If MyCTRL.Tag = vbNullString Then
        MyCTRL.Tag = MyCTRL.SelStart
        tString = FcaseStringF(MyCTRL.Text, True)
        MyCTRL.Text = LTrim$(tString)
        MyCTRL.SelStart = MyCTRL.Tag
        MyCTRL.Tag = vbNullString
    End If
End Sub
Public Function FcaseStringF(ByVal tString As String, Optional NoRTrim As Boolean = False, Optional CaseType As Integer = vbProperCase) As String
    Dim x As Integer
    '/*************************************/
    '/* Convert any string to Formal Case */
    '/*************************************/
    On Local Error Resume Next
    tString = tString & vbNullString
    If NoRTrim Then
        tString = LTrim(tString)
    Else
        tString = Trim(tString)
    End If
    If tString = vbNullString Then
        FcaseStringF = vbNullString
        Exit Function
    End If
    
    Select Case CaseType
    Case vbUpperCase
        FcaseStringF = UCase(tString)
    Case vbLowerCase
        FcaseStringF = LCase(tString)
    Case vbProperCase
    
        tString = UCase$(Left$(tString, 1)) & LCase$(Mid$(tString, 2))
        '/* Look for space */
        x = InStr(tString, " ")
        If x Then
           Do
               Mid$(tString, x + 1, 1) = UCase$(Mid$(tString, x + 1, 1))
               x = x + 1
               x = InStr(x, tString, " ")
               If x = 0 Or x + 1 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for . */
        x = InStr(tString, ".")
        If x Then
           Do
               Mid$(tString, x + 1, 1) = UCase$(Mid$(tString, x + 1, 1))
               x = x + 1
               x = InStr(x, tString, ".")
               If x = 0 Or x + 1 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for / */
        x = InStr(tString, "/")
        If x Then
           Do
               Mid$(tString, x + 1, 1) = UCase$(Mid$(tString, x + 1, 1))
               x = x + 1
               x = InStr(x, tString, "/")
               If x = 0 Or x + 1 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for & */
        x = InStr(tString, "&")
        If x Then
           Do
               Mid$(tString, x + 1, 1) = UCase$(Mid$(tString, x + 1, 1))
               x = x + 1
               x = InStr(x, tString, "&")
               If x = 0 Or x + 1 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for Mc */
        x = InStr(tString, "Mc")
        If x Then
           Do
               Mid$(tString, x + 2, 1) = UCase$(Mid$(tString, x + 2, 1))
               x = x + 2
               x = InStr(x, tString, "Mc")
               If x = 0 Or x + 2 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for O' */
        x = InStr(tString, "O'")
        If x Then
           Do
               Mid$(tString, x + 2, 1) = UCase$(Mid$(tString, x + 2, 1))
               x = x + 2
               x = InStr(x, tString, "O'")
               If x = 0 Or x + 2 > Len(tString) Then Exit Do
           Loop
        End If
        '/* Look for - */
        x = InStr(tString, "-")
        If x Then
           Do
               Mid$(tString, x + 1, 1) = UCase$(Mid$(tString, x + 1, 1))
               x = x + 1
               x = InStr(x, tString, "-")
               If x = 0 Or x + 1 > Len(tString) Then Exit Do
           Loop
        End If
    End Select
    FcaseStringF = tString
    
End Function

Public Function MaskedText(MyCTRL As TextBox, ByVal MyProfile As String) As Boolean
'/* Place in MyCTRL_KeyUp for imediate check */
'/* Place in MyCTRL_KeyPress for 1 char delay check */

 Const GoodMatchColor = &H80000008
 Const BadMatchColor = vbRed

 On Local Error GoTo BadCheckProfile

 Dim MyCursorPlace As Long
 Dim MyLen As Long
 Dim MyPlace As Long
 Dim MyBuffer As String
 Dim MyText As String
 Dim MyChar As String * 1
 Dim MyProfilePlace As Long

    MyPlace = 1
    MyProfilePlace = 1
    
    '/* If there are more characters than allowed then remove them */
    If Len(MyCTRL.Text) > Len(MyProfile) Then
        MyCTRL.Text = Left(MyCTRL.Text, Len(MyProfile))
        MyCTRL.SelStart = Len(MyCTRL.Text)
        Beep
    End If
    
    MyText = MyCTRL.Text
    MyLen = Len(MyText)
    MyCursorPlace = MyCTRL.SelStart
    
    '/* The parser takes the pattern as the transition map. Starting
    '/* at the beginning of the map, it compares the current character
    '/* with the state of the parser
    Do While MyPlace <= MyLen
        MyChar = Mid(MyText, MyPlace, 1)
        Select Case Mid(MyProfile, MyProfilePlace, 1)
            '/* Check if MyProfile calls for a numeric input */
         Case "#"
            If IsNumeric(MyChar) Then
                '/* Add the character to the buffer
                MyBuffer = MyBuffer & MyChar
                '/* Move to the next character
                MyPlace = MyPlace + 1
                '/* Move to the next valid parser state
                MyProfilePlace = MyProfilePlace + 1
                '/* Make sure we are indicating a valid transition state
                MyCTRL.ForeColor = GoodMatchColor
            Else
                '/* The character does not match the parser's state so
                '/* indicate an invalid state and exit the parser
                MyCTRL.ForeColor = BadMatchColor
                GoTo ExitCheckProfile
            End If
         Case "@"
            If Not IsNumeric(MyChar) Then
                '/* Add the character to the buffer
                MyBuffer = MyBuffer & MyChar
                '/* Move to the next character
                MyPlace = MyPlace + 1
                '/* Move to the next valid parser state
                MyProfilePlace = MyProfilePlace + 1
                '/* Make sure we are indicating a valid transition state
                MyCTRL.ForeColor = GoodMatchColor
            Else
                '/* The character does not match the parser's state so
                '/* indicate an invalid state and exit the parser
                MyCTRL.ForeColor = BadMatchColor
                GoTo ExitCheckProfile
            End If
         Case Else
            If MyChar = Mid(MyProfile, MyProfilePlace, 1) Then
                '/* If it Is here Then add the character to the buffer */
                MyBuffer = MyBuffer & MyChar
                '/* Move to next character position */
                MyPlace = MyPlace + 1
                '/* Move to next parser state */
                MyProfilePlace = MyProfilePlace + 1
                '/* Indicate a valid transition state to the user */
                MyCTRL.ForeColor = GoodMatchColor
            Else
                '/* The required character is not present and in this case we
                '/* insert it meeting the requirements of the parser state
                MyBuffer = MyBuffer & Mid(MyProfile, MyProfilePlace, 1)
                '/* Shift the parser to the next state but stay with the current
                '/* character to see if it matches the next state transition
                MyProfilePlace = MyProfilePlace + 1
                '/* Move the input cursor
                '/* over one to account for the added character
                MyCursorPlace = MyCursorPlace + 1
            End If
        End Select
    Loop
    
    '/* If this is a function */
    If Len(MyBuffer) = Len(MyProfile) Then
        MaskedText = True
    Else
        MaskedText = False
    End If
    
    MyCTRL.Text = MyBuffer
    MyCTRL.SelStart = MyCursorPlace
    
    '=======================================
ExitCheckProfile:
    '/* exit code here */
Exit Function


BadCheckProfile:
    MsgBox Err.Description, vbCritical, "CheckProfile"
    '/* If this is a function */
    MaskedText = False
    Resume ExitCheckProfile
End Function
Function IsNullOrEmpty(ByVal InVal As Variant) As Boolean
'/* This function returns True if the value passed as InVal is
'/* Null or a zero-length string.

    If IsNull(InVal) Then
        IsNullOrEmpty = True
    Else
        IsNullOrEmpty = (Len(InVal) = 0)
    End If
End Function

Public Function IsNumber(ByVal CheckString As String, Optional KeyAscii As Integer = 0, Optional AllowDecPoint As Boolean = True, Optional AllowNegative As Boolean = False) As Boolean
    If KeyAscii > 0 And KeyAscii <> 8 Then
        If Not AllowNegative And KeyAscii = 45 Then KeyAscii = 0
        If Not AllowDecPoint And KeyAscii = 46 Then KeyAscii = 0
        If CheckString = vbNullString And KeyAscii = 46 Then CheckString = "0"
        If Not IsNumeric(CheckString & Chr(KeyAscii)) Then
            KeyAscii = False
            IsNumber = False
        Else
            IsNumber = True
        End If
    Else
        IsNumber = IsNumeric(CheckString)
    End If
End Function
