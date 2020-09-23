Attribute VB_Name = "modNimSpectrum"
Sub Main()
'the non-avoidable VB initialization...

Set modZXBasic.viewForm = frmNim
Set modZXBasic.viewPic = frmNim.picView
Call modZXBasic.InitBasic
frmNim.Show

'and we start our pseudo-Basic routine...
NEWs
RUN
End Sub

Sub RUN()

10 Rem ******************************************************************
12 Rem * Gonchuki Systems Productions - VB Port from the "Run" magazine *
15 Rem ******************************************************************
20 BORDER 1
21 PAPER 1
22 INK 7
    'Not used in a PC (this function handles directly the memory of the Spectrum changing in this case the system variables)
'25 POKE 23609, 100 'well, this one may be used, the effect this instruction does is that you hear a sound each time you press a key...
'26 POKE 23658, 8   'this changes the cursor to Caps so we write all in upper case... a slight work-aroud was used in the ZX-BASIC module so i don't have to take care of it...
30 CLS
35 Randomize 'ADDED BY ME. There's a VB fault (or PC fault?) that makes the game to have always the same random nubers generated, the so-intelligent ZX Spectrum did not that fault.
40 PRINT_AT 21, 0, , , , "NUMBER OF COLUMNS? (3-8)"
50 INPUTs L
60 If L < 3 Or L > 8 Then GoTo 50
70 PRINT_AT 21, 0, , , , "DIFFICULTY LEVEL?  (0-5)"
80 INPUTs D
90 If D < 0 Or D > 5 Then GoTo 80
100 Dim C1() As Long: ReDim C1(L) 'A SMALL ADAPTATION SINCE VB FAILS TO COMPARE CORRECTLY TWO VARIANTS...
110 Dim B1() As Long: ReDim B1(L, 4)
120 Dim S1() As Long: ReDim S1(4)
130 Let B$ = "                                "
140 For M = 1 To L
150 Let X = Int(Rnd * 8) + 1
160 Let C1(M) = X
170 Next M
180 CLS
190 PRINT_AT 0, 0, 7, 2, , "             N I M              "
'200 PRINT_AT
210 PRINT_AT 2, 0, , , , "Column:  "
220 For C = 1 To L
230 PRINT_AT 2, C * 3 + 7, , , , C
240 Next C
250 PRINT_AT 13, 0, , , , "Quantity:"
260 For M = 1 To L
270 PRINT_AT 13, M * 3 + 7, , 5, , C1(M)
280 Next M
290 For C = 1 To L
300 For M = 1 To C1(C)
310 PRINT_AT 12 - M, C * 3 + 7, , 6, 1, "*"
320 Next M
330 Next C
340 Let J = 0
350 PRINT_AT 15, 0, 7, , , B$
360 PRINT_AT 17, 0, , , , "MOVE:"
370 PRINT_AT 19, 0, , , , "PIECES:"
380 PRINT_AT 19, 24, , , , "LEVEL: " & D
390 GoSub 1170
400 If Rnd > 0.5 Then GoTo 570
410 GoSub 1140
420 PRINT_AT 17, 23, 1, 6, 1, "YOUR TURN"
430 PRINT_AT 21, 0, , , , B$
440 PRINT_AT 21, 0, , 6, , "COLUMN? "
450 INPUTs M
460 If Val(M) < 1 Or Val(M) > L Then GoTo 430
470 PRINT_AT 21, 10, , 6, , M
480 PRINT_AT 21, 15, , 6, , "QUANTITY? "
490 INPUTs N
500 PRINT_AT 21, 27, , 6, , N
505 PAUSE 400 'ADDED BY ME
510 If Val(N) > C1(M) Or Val(N) < 1 Then GoTo 430
520 GoSub 1230
530 GoSub 1170
540 If P > 0 Then GoTo 570
550 PRINT_AT 21, 0, 5, 0, , "   CONGRATULATIONS, YOU WIN!!!  "
560 GoTo 1090
570 GoSub 1140
580 PRINT_AT 17, 23, 1, 6, 1, " MY TURN "
590 PRINT_AT 21, 0, , , , B$
595 PAUSE 400 'ADDED BY ME
600 If Rnd * 10 > D * 2 Then GoTo 810
610 For M = 1 To L
620 Let X = C1(M)
630 For C = 4 To 1 Step -1
640 Let Z = Int(X / 2)
650 Let B1(M, C) = X - 2 * Z
660 Let X = Z
670 Next C
680 Next M
690 For C = 1 To 4
700 Let X = 0
710 For I = 1 To L
720 Let X = X + B1(I, C)
730 Next I
740 Let S1(C) = X - 2 * Int(X / 2)
750 Next C
760 Let X = 0
770 For I = 1 To 4
780 Let X = X + S1(I)
790 Next I
800 If X Then GoTo 860
810 For M = 1 To L
820 If C1(M) = 0 Then GoTo 850
830 Let N = Int(C1(M) * Rnd + 1)
840 GoTo 1010
850 Next M
860 For C = 1 To 4
870 If S1(C) > 0 Then GoTo 890
880 Next C
890 For M = 1 To L
900 If B1(M, C) = 1 Then GoTo 920
910 Next M
920 Let N = 0
930 For C = C To 4
940 If S1(C) = 0 Then GoTo 1000
950 Let X = 2 ^ (4 - C)
960 If B1(M, C) = 0 Then GoTo 990
970 Let N = N + X
980 GoTo 1000
990 Let N = N - X
1000 Next C
1010 PRINT_AT 21, 0, , 6, , "COLUMN?   " & M & "    QUANTITY?   " & N
1020 PAUSE 1000
'1030 BEEP 0.1, 40
1040 GoSub 1230
1050 GoSub 1170
1060 If P Then GoTo 410
1070 PRINT_AT 21, 0, 5, 0, , "       S O R R Y, I  W I N       "
1080 For I = 30 To 0 Step -1
'1081 BEEP 0.025, I
1082 Next I
1083 GoTo 1100
1090 For I = 0 To 30
'1091 BEEP 0.025, I
1092 Next I
1100 PRINT_AT 7, 0, 2, 6, , "DO YOU WANT TO PLAY AGAIN? (Y/N)"
1110 Let R = INKEY$
1120 If R = "Y" Then GoTo 10
1125 If R <> "N" Then GoTo 1110
1130 End 'GOTO 10000 in Basic
1140 Let J = J + 1
1150 PRINT_AT 17, 8, , , , J
1160 Return
1170 Let P = 0
1180 For M = 1 To L
1190 Let P = P + C1(M)
1200 Next M
1210 PRINT_AT 19, 8, , , , P & " "
1220 Return
1230 Let X = M * 3 + 7
1240 For I = 1 To N
1250 PRINT_AT 12 - C1(M), X, , , , " "
1260 Let C1(M) = C1(M) - 1
1265 PAUSE 30 'TINY ANIMATOIN ADDED BY ME, SO THE USER FEELS THAT THE PIECES ARE BEING TAKEN AWAY
1270 Next I
1280 PRINT_AT 13, X, , 5, , C1(M)
1290 Return
End Sub
