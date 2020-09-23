Attribute VB_Name = "modZXBasic"
Option Explicit

'ZX-BASIC
' Copyright © 2002 by gonchuki
' e-mail: gonchuki@yahoo.es

'This module is not yet complete, as some of the original
'functions are not implemented here. Some of them because
'i don't need them, and the others because they are only
'applicable to the ZX Spectrum

Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private frmViewport  As Form
Private picViewport As PictureBox
Private bCOLORS(0 To 7, 0 To 1) As Long, CI As Long, CP As Long, CB As Long
Private useBright As Boolean
Private Keys As String, KeysL As String

'the initialization routines...
Public Property Set viewForm(ByVal theForm As Form)
    Set frmViewport = theForm
End Property

Public Property Set viewPic(ByVal thePic As PictureBox)
    Set picViewport = thePic
End Property

Public Sub InitBasic()
'we initialize this array with the original colors used in the ZX Spectrum, the second bound is to represent the color when BRIGHT is set to 1
    bCOLORS(0, 0) = &H0
    bCOLORS(0, 1) = &H202020
    bCOLORS(1, 0) = &HC00000
    bCOLORS(1, 1) = &HFF0000
    bCOLORS(2, 0) = &HDF
    bCOLORS(2, 1) = &HFF
    bCOLORS(3, 0) = &H6000DF
    bCOLORS(3, 1) = &H8000FF
    bCOLORS(4, 0) = &HC060&
    bCOLORS(4, 1) = &HFF80&
    bCOLORS(5, 0) = &HDF60
    bCOLORS(5, 1) = &HFF80
    bCOLORS(6, 0) = &HDFDF&
    bCOLORS(6, 1) = &HFFFF&
    bCOLORS(7, 0) = &HDFDFDF
    bCOLORS(7, 1) = &HFFFFFF
    
'and these are the default colors used
    CP = 7: CI = 0: CB = 7
End Sub

Public Sub AddKey(ByVal KeyAscii As Long)
    If KeyAscii = 8 Then 'backspace key pressed...
        If Len(Keys) Then Keys = Left$(Keys, Len(Keys) - 1)
    ElseIf KeyAscii > 12 Then 'we simply skip the first control codes that are useless here
        Keys = Keys & Chr$(KeyAscii)
    End If
End Sub

'*********************************
'*  the replacement routines...  *
'*********************************

Sub BORDER(ByVal COLOR As Long)
    CB = COLOR
    frmViewport.BackColor = bCOLORS(COLOR, Abs(useBright))
End Sub

Sub PAPER(ByVal COLOR As Long)
    CP = COLOR
    picViewport.BackColor = bCOLORS(COLOR, Abs(useBright))
End Sub

Sub INK(ByVal COLOR As Long)
    CI = COLOR
    picViewport.ForeColor = bCOLORS(CI, Abs(useBright))
End Sub

Sub BRIGHT(ByVal STATUS As Byte)
    useBright = STATUS
End Sub

Sub PRINT_AT(ByVal Y As Long, ByVal X As Long, Optional ByVal PAPERCOLOR As Long = -1, Optional ByVal INKCOLOR As Long = -1, Optional ByVal HASBRIGHT As Long = -1, Optional ByVal TEXT As String)
Dim UB As Boolean: If HASBRIGHT > -1 Then UB = HASBRIGHT Else UB = useBright
If PAPERCOLOR > -1 Then SetBkColor picViewport.hDC, bCOLORS(PAPERCOLOR, Abs(UB)) Else SetBkColor picViewport.hDC, bCOLORS(CP, Abs(useBright))
If INKCOLOR > -1 Then SetTextColor picViewport.hDC, bCOLORS(INKCOLOR, Abs(UB)) Else SetTextColor picViewport.hDC, bCOLORS(CI, Abs(useBright))

    TextOut picViewport.hDC, X * 7, Y * 9, TEXT, Len(TEXT)
    picViewport.Refresh
End Sub

Sub CLS()
    frmViewport.CLS
    picViewport.CLS
End Sub

Sub INPUTs(ByRef VAR As Variant)
    'clean-up the variables
    VAR = "": Keys = "": PRINT_AT 23, 0, , , , String$(32, " ")
    Do 'we loop until the ENTER key is pressed
        If Len(Keys) Then
            If Asc(Right$(Keys, 1)) = 13 Then
                If Len(Keys) > 1 Then
                    VAR = Left$(Keys, Len(Keys) - 1)
                    Keys = Left$(Keys, Len(Keys) - 1)
                Else
                    VAR = " "
                End If
            End If
        End If
            If KeysL <> Keys Then
                If Len(Keys) Then
                    PRINT_AT 23, 0, , , , Format$(Keys, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
                Else
                    PRINT_AT 23, 0, , , , String$(32, " ")
                End If
                KeysL = Keys
            End If
        DoEvents
    Loop While VAR = ""
    PRINT_AT 23, 0, , , , String$(32, " ")
End Sub

Sub PAUSE(ByVal TIME As Long)
    Sleep TIME
End Sub

Function INKEY$()
    Dim VAR As String
    Keys = ""
    Do
        If Len(Keys) Then VAR = UCase$(Keys)
        DoEvents
    Loop While VAR = ""
    INKEY$ = VAR
End Function

Sub NEWs()
Dim I As Long, J As Long, K As Long
'add some effects so it appears that we are restarting the spectrum
    PAPER 0: BORDER 0: CLS: PAUSE 500
    InitBasic
    PAPER CP: BORDER CB: INK CI: BRIGHT 0: CLS
    PRINT_AT 8, 0, , 2, , "        GONCHUKI SYSTEMS        "
    PRINT_AT 10, 0, , 1, , " ZX SPECTRUM BASIC CODE PORTER  "
    PRINT_AT 14, 0, , 4, , " COPYRIGHT © 2002 BY GONCHUKI   "
    PRINT_AT 16, 0, , 3, , "   E-MAIL: GONCHUKI@YAHOO.ES    "
    
    For K = 0 To 1
    For J = 0 To 7
        For I = 1 To 32
            PRINT_AT 23, I, , J, K, "*"
            PAUSE 5
        Next
    Next
    Next
    CLS
End Sub

