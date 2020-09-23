Attribute VB_Name = "ModScoreArrays"
'ModScoreArrays.bas  ~RRChess~

Option Explicit

' Static piece position scores & some openings

' 16 score tables 8x8 integers = 2048 B
' Score range -10 to 20 ?
Public WPawns_Open() As Integer          ' < 40 half-moves?
Public BPawns_Open() As Integer          ' < 40 half-moves?
Public WPawns() As Integer               ' >= 40 half-moves?
Public BPawns() As Integer               ' >= 40 half-moves?
Public WPawns_End() As Integer           ' < 8 Black pieces
Public BPawns_End() As Integer           ' < 8 White pieces
Public WRook() As Integer
Public BRook() As Integer
Public WKnight() As Integer
Public BKnight() As Integer
Public WWBishop_BBBishop_Open() As Integer   ' < 20 half-moves White's white bishop & Black's black bishop
Public WBBishop_BWBishop_Open() As Integer   ' < 20 half-moves White's black bishop & Black's white bishop
Public WWBishop_BBBishop() As Integer   ' White's white bishop & Black's black bishop
Public WBBishop_BWBishop() As Integer   ' White's black bishop & Black's white bishop
Public WQueen_Open() As Integer         ' < 20  half-moves?
Public BQueen_Open() As Integer         ' < 20  half-moves?
Public WBQueen() As Integer             ' >= 20 half-moves White & Black queens
Public WKing() As Integer               ' < 60 half-moves?
Public BKing() As Integer               ' < 60 half-moves?
Public WBKing_End() As Integer          ' >= 60 half-moves White & Black king's later
Public Opener() As Byte       ' Full board
Public SOpenString$()

Private RCPN() As Byte  ' To check if piece is there.  Blocked??


Public Sub FillScoreArrays()
Dim a$
Dim REV As Long  ' Reverse rows for some blacks
REV = 1

' 20 score tables 8x8 integers = 2560 B
' Score range -99 to 20 ?
ReDim WPawns_Open(1 To 8, 1 To 8) As Integer
ReDim BPawns_Open(1 To 8, 1 To 8) As Integer
ReDim WPawns(1 To 8, 1 To 8) As Integer
ReDim BPawns(1 To 8, 1 To 8) As Integer
ReDim WPawns_End(1 To 8, 1 To 8) As Integer
ReDim BPawns_End(1 To 8, 1 To 8) As Integer
ReDim WRook(1 To 8, 1 To 8) As Integer
ReDim BRook(1 To 8, 1 To 8) As Integer
ReDim WKnight(1 To 8, 1 To 8) As Integer
ReDim BKnight(1 To 8, 1 To 8) As Integer
ReDim WWBishop_BBBishop_Open(1 To 8, 1 To 8) As Integer
ReDim WBBishop_BWBishop_Open(1 To 8, 1 To 8) As Integer
ReDim WWBishop_BBBishop(1 To 8, 1 To 8) As Integer
ReDim WBBishop_BWBishop(1 To 8, 1 To 8) As Integer
ReDim WQueen_Open(1 To 8, 1 To 8) As Integer
ReDim BQueen_Open(1 To 8, 1 To 8) As Integer
ReDim WBQueen(1 To 8, 1 To 8) As Integer
ReDim WKing(1 To 8, 1 To 8) As Integer
ReDim BKing(1 To 8, 1 To 8) As Integer
ReDim WBKing_End(1 To 8, 1 To 8) As Integer


'#  PAWNS ####################################################

' " WPawns_Open"
a$ = ""
a$ = a$ + "000 000 000 000 000 000 000 000 "  ' R=8, C= 1 to 8
a$ = a$ + "000 000 000 000 000 000 000 000 "
a$ = a$ + "000 000 008 008 008 000 000 000 "
a$ = a$ + "000 000 008 008 008 000 000 000 "
a$ = a$ + "-05 -05 016 025 025 -25 -05 -05 "
a$ = a$ + "000 005 010 020 020 -25 005 000 "
a$ = a$ + "000 000 -05 -05 -05 000 000 000 "
a$ = a$ + "000 000 000 000 000 000 000 000 "  ' R=1, C= 1 to 8
FillScoreArray WPawns_Open(), a$
' " BPawns_Open"
FillScoreArray BPawns_Open(), a$, REV

' " WPawns"
a$ = ""
a$ = a$ + "012 012 012 012 012 012 012 012 "
a$ = a$ + "004 004 004 004 004 004 004 004 "
a$ = a$ + "004 004 004 004 004 004 004 004 "
a$ = a$ + "002 002 006 006 006 004 002 002 "
a$ = a$ + "001 002 004 004 004 004 002 001 "
a$ = a$ + "001 002 002 002 002 002 002 001 "
a$ = a$ + "000 000 000 -02 -02 000 000 000 "
a$ = a$ + "000 000 000 000 000 000 000 000 "
FillScoreArray WPawns(), a$
' " BPawns"
FillScoreArray BPawns(), a$, REV

' " WPawns_End"
a$ = ""
a$ = a$ + "020 020 020 020 020 020 020 020 "  ' R=8, C= 1 to 8
a$ = a$ + "010 010 010 010 010 010 010 010 "
a$ = a$ + "008 008 008 008 008 008 008 008 "
a$ = a$ + "006 006 006 006 006 006 006 006 "
a$ = a$ + "005 005 005 005 005 005 005 005 "
a$ = a$ + "003 003 003 003 003 003 003 003 "
a$ = a$ + "003 003 003 003 003 003 003 003 "
a$ = a$ + "000 000 000 000 000 000 000 000 "  ' R=1, C= 1 to 8
FillScoreArray WPawns_End(), a$
' " BPawns_End"
FillScoreArray BPawns_End(), a$, REV

'#  ROOKS ###################################################
' " WRook"
a$ = ""
a$ = a$ + "006 006 006 006 006 006 006 006 "
a$ = a$ + "008 008 008 008 008 008 008 008 "
a$ = a$ + "006 006 006 006 006 006 006 006 "
a$ = a$ + "000 004 004 006 006 004 004 000 "
a$ = a$ + "000 002 004 006 006 004 002 000 "
a$ = a$ + "002 002 002 004 004 002 002 002 "
a$ = a$ + "000 000 000 004 004 000 000 000 "
a$ = a$ + "000 -06 -05 016 016 005 -06 000 "
FillScoreArray WRook(), a$
' " BRook"
FillScoreArray BRook(), a$, REV

'#  KNIGHTS #################################################
' " WKnight"
a$ = ""
a$ = a$ + "002 002 002 002 002 002 002 002 "
a$ = a$ + "004 004 004 004 004 004 004 004 "
a$ = a$ + "006 006 006 006 006 006 006 006 "
a$ = a$ + "004 005 005 008 008 005 005 004 "
a$ = a$ + "003 004 004 008 008 004 004 003 "
a$ = a$ + "-02 003 013 003 003 013 003 -02 "
a$ = a$ + "002 002 002 002 002 002 002 002 "
a$ = a$ + "002 -02 002 002 002 002 -02 002 "
FillScoreArray WKnight(), a$
' " BKnight"
FillScoreArray BKnight(), a$, REV

'#  BISHOPS #################################################
' " WWBishop_BBBishop_Open"  White white bishop, Black black bishop
a$ = ""
a$ = a$ + "-05 -05 -05 -05 -05 -99 -05 -05 "
a$ = a$ + "001 008 005 005 005 005 008 001 "
a$ = a$ + "001 005 008 -10 005 008 005 001 "
a$ = a$ + "005 005 016 008 008 005 005 001 "
a$ = a$ + "005 005 016 008 008 005 005 001 "
a$ = a$ + "001 006 008 -10 005 008 006 001 "
a$ = a$ + "001 008 005 005 005 005 008 001 "
a$ = a$ + "-05 -05 000 -05 000 -99 002 -05 "
FillScoreArray WWBishop_BBBishop_Open(), a$
' " WBBishop_BWBishop_Open"  White black bishop, Black white bishop
a$ = ""
a$ = a$ + "-05 -05 -99 -05 -05 -05 -05 -05 "
a$ = a$ + "001 008 005 010 005 005 008 001 "
a$ = a$ + "001 005 008 006 -10 008 005 001 "
a$ = a$ + "001 005 006 008 008 006 005 001 "
a$ = a$ + "001 005 006 008 008 006 005 001 "
a$ = a$ + "001 005 008 006 -10 008 005 001 "
a$ = a$ + "001 008 005 010 005 005 008 001 "
a$ = a$ + "-05 -05 -99 -05 -05 -05 -05 -05 "
FillScoreArray WBBishop_BWBishop_Open(), a$

' " WWBishop_BBBishop"  White white bishop, Black black bishop
a$ = ""
a$ = a$ + "-05 -05 -05 -05 -05 -45 -05 -05 "
a$ = a$ + "001 008 005 005 006 005 008 001 "
a$ = a$ + "001 006 008 007 004 008 005 001 "
a$ = a$ + "001 005 006 008 008 006 005 001 "
a$ = a$ + "001 005 006 008 008 006 005 001 "
a$ = a$ + "001 005 008 007 004 008 005 001 "
a$ = a$ + "001 008 005 005 006 005 008 001 "
a$ = a$ + "-05 -05 -05 -05 -05 -45 -05 -05 "
FillScoreArray WWBishop_BBBishop(), a$
' " WBBishop_BWBishop"  White black bishop, Black white bishop
a$ = ""
a$ = a$ + "-05 -05 -45 -05 -05 -05 -05 -05 "
a$ = a$ + "001 008 005 005 005 005 008 001 "
a$ = a$ + "001 005 008 004 004 008 005 001 "
a$ = a$ + "001 005 008 008 008 008 006 001 "
a$ = a$ + "001 005 008 008 008 008 006 001 "
a$ = a$ + "001 005 008 004 004 008 005 001 "
a$ = a$ + "001 008 005 005 005 005 008 001 "
a$ = a$ + "-05 -05 -45 -05 -05 -05 -05 -05 "
FillScoreArray WBBishop_BWBishop(), a$


'#  QUEENS ##################################################
' " WQueen_Open"
a$ = ""
a$ = a$ + "-05 -05 -10 -10 -10 -10 -05 -05 "
a$ = a$ + "-05 -05 -10 -10 -10 -10 -05 -05 "
a$ = a$ + "-05 -05 -10 -10 -10 -10 -05 -05 "
a$ = a$ + "-05 -05 -10 -10 -10 -10 -05 -05 "
a$ = a$ + "-05 -05 -10 -10 -10 -10 -05 -05 "
a$ = a$ + "-05 -15 -15 -10 -10 -15 -15 -05 "
a$ = a$ + "000 000 002 005 005 002 000 000 "
a$ = a$ + "000 000 000 035 000 000 000 000 "
FillScoreArray WQueen_Open(), a$
' " BQueen_Open"
FillScoreArray BQueen_Open(), a$, REV

' " WBQueen"  ie both
a$ = ""
a$ = a$ + "000 000 000 000 000 000 000 000 "
a$ = a$ + "000 005 005 005 005 005 005 000 "
a$ = a$ + "000 005 010 010 010 010 005 000 "
a$ = a$ + "000 005 010 012 012 010 005 000 "
a$ = a$ + "000 005 010 012 012 010 005 000 "
a$ = a$ + "000 005 010 010 010 010 005 000 "
a$ = a$ + "000 005 010 010 010 010 005 000 "
a$ = a$ + "000 000 000 000 000 000 000 000 "

FillScoreArray WBQueen(), a$

'#  KINGS ###################################################
' " WKing"
a$ = ""
a$ = a$ + "-20 -20 -20 -20 -20 -20 -20 -20 "
a$ = a$ + "-40 -40 -40 -40 -40 -40 -40 -40 "
a$ = a$ + "-38 -38 -38 -38 -38 -38 -38 -38 "
a$ = a$ + "-28 -28 -28 -28 -28 -28 -28 -28 "
a$ = a$ + "-16 -16 -16 -16 -16 -16 -16 -16 "
a$ = a$ + "-09 -09 -09 -09 -09 -09 -09 -09 "
a$ = a$ + "-05 -05 -05 -05 -05 -05 -05 -05 "
a$ = a$ + "000 010 010 -10 005 -10 010 000 "
FillScoreArray WKing(), a$
' " BKing"
FillScoreArray BKing(), a$, REV

' " WBKing_End"
a$ = ""
a$ = a$ + "000 000 000 000 000 000 000 000 "
a$ = a$ + "000 005 005 005 005 005 005 000 "
a$ = a$ + "000 005 008 008 008 008 005 000 "
a$ = a$ + "000 005 008 008 008 008 005 000 "
a$ = a$ + "000 005 008 008 008 008 005 000 "
a$ = a$ + "000 005 008 008 008 008 005 000 "
a$ = a$ + "000 005 005 005 005 005 005 000 "
a$ = a$ + "000 000 000 000 000 000 000 000 "
FillScoreArray WBKing_End(), a$
'############################################################
'############################################################


ReDim Opener(1 To 8, 1 To 8)   ' As Byte
' 64*20 = 1280 B
' FULL BOARD
a$ = ""
a$ = a$ + "007 008 009 010 011 009 008 007 "
a$ = a$ + "012 012 012 012 012 012 012 012 "
a$ = a$ + "000 000 000 000 000 000 000 000 "
a$ = a$ + "000 000 000 000 000 000 000 000 "
a$ = a$ + "000 000 000 000 000 000 000 000 "
a$ = a$ + "000 000 000 000 000 000 000 000 "
a$ = a$ + "006 006 006 006 006 006 006 006 "
a$ = a$ + "001 002 003 004 005 003 002 001 "
FillOpenArray Opener(), a$
a$ = ""


' SOME SIMPLE OPENINGS
' These could be placed in a file
ReDim SOpenString$(1 To 80)
' WHITE'S FIRST BLACK'S RESPONSE
' HalfMove = 0
'3
SOpenString$(1) = "e2-e4 e7-e5 "  ' P-K4 P-K4    >.5
SOpenString$(2) = "e2-e4 e7-e6 "  ' P-K4 P-K3    >.3
SOpenString$(3) = "e2-e4 c7-c5 "  ' P-K4 P-QB4   <=.3
'3
SOpenString$(4) = "d2-d4 d7-d5 "  ' P-Q4  P-Q4   >.5
SOpenString$(5) = "d2-d4 d7-d6 "  ' P-Q4  P-Q3   >.3
SOpenString$(6) = "d2-d4 g8-f6 "  ' P-Q4  N-KB3  <=.3
'2
SOpenString$(7) = "g1-f3 g8-f6 "  ' N-KB3 N-KB3  >=.5
SOpenString$(8) = "g1-f3 d7-d5 "  ' N-KB3 P-Q4   <.5
'2
SOpenString$(9) = "b1-c3 e7-e5 "  ' N-QB3 P-K4   >=.5
SOpenString$(10) = "b1-c3 b8-c6 " ' N-QB3 N-QB3  <.5
'1s
SOpenString$(11) = "e2-e3 e7-e5 " ' P-K3 P-K4
SOpenString$(12) = "d2-d3 d7-d5 " ' P-Q3 P-Q4
SOpenString$(13) = "g2-g3 e7-e5 " ' P-KN3 P-K4
SOpenString$(14) = "b2-b3 d7-d5 " ' P-QN3 P-Q4
SOpenString$(15) = "c2-c4 e7-e5 " ' P-QB4 P-K4
SOpenString$(16) = "c2-c3 e7-e5 " ' P-QB3 P-K4

' WHITE'S SECOND BLACK'S SECOND RESPONSE
' HalfMove=2
'2
SOpenString$(21) = "e2-e4 e7-e5 g1-f3 b8-c6"  ' 1. P-K4 P-K4 2. N-KB3 N-QB3   >.5
SOpenString$(22) = "e2-e4 e7-e5 g1-f3 d7-d6"  ' 1. P-K4 P-K4 2. N-KB3 P-Q3    <=.5
'2
SOpenString$(23) = "d2-d4 d7-d5 g1-f3 g8-f6"  ' 1. P-Q4 P-Q4 2. N-KB3 N-KB3   >.5
SOpenString$(24) = "d2-d4 d7-d5 g1-f3 e7-e6"  ' 1. P-Q4 P-Q4 2. N-KB3 P-K3    <=.5

'-----------------------------------------------
' WHITE'S FIRST MOVE if comp = "W"
' HalfMove = -1
SOpenString$(40) = "e2-e4 "  ' P-K4   >=.5
SOpenString$(41) = "d2-d4 "  ' P-Q4

' BLACK'S SECOND WHITE'S SECOND RESPONSE
' HalfMove = 1
SOpenString$(42) = "e2-e4 e7-e5 g1-f3 "  ' 1. P-K4 P-K4 2. N-KB3    >.5
SOpenString$(43) = "e2-e4 e7-e5 b1-c3 "  ' 1. P-K4 P-K4 2. N-QB3
SOpenString$(44) = "d2-d4 d7-d5 g1-f3 "  ' 1. P-Q4 P-Q4 2. N-KB3    >.5
SOpenString$(45) = "d2-d4 d7-d5 b1-c3 "  ' 1. P-Q4 P-Q4 2. N-QB3

End Sub

Public Function BlackOMoves(Index As Integer) As Boolean
' Public MoveString$
' Public HalfMove, OpenString$,SOpenString$()
Dim Rand As Single
Dim P$
Dim k As Long
   MoveString$ = ""
   BlackOMoves = True
   Randomize
   Rand = Rnd
   
   Select Case HalfMove
   Case 0   ' Black's 1st response
      Select Case OpenString$
      Case "e2-e4 " ' P-K4
'3
'SOpenString$(1) = "e2-e4 e7-e5 "  ' P-K4 P-K4    >.5
'SOpenString$(2) = "e2-e4 e7-e6 "  ' P-K4 P-K3    >.3
'SOpenString$(3) = "e2-e4 c7-c5 "  ' P-K4 P-QB4   <=.3
            Select Case Rand
            Case Is > 0.5  ' 1. P-K4 P-K4
               P$ = Mid$(SOpenString$(1), 7, 5): MakeOMove P$, Index ': Exit Function
            Case Is > 0.3  ' 1. P-K4 P-K3
               P$ = Mid$(SOpenString$(2), 7, 5): MakeOMove P$, Index ': Exit Function
            Case Else  ' 1. P-K4 P-QB4
               P$ = Mid$(SOpenString$(3), 7, 5): MakeOMove P$, Index ': Exit Function
            End Select
      Case "d2-d4 "   ' P-Q4
'3
'SOpenString$(4) = "d2-d4 d7-d5 "  ' P-Q4  P-Q4   >.5
'SOpenString$(5) = "d2-d4 d7-d6 "  ' P-Q4  P-Q3   >.3
'SOpenString$(6) = "d2-d4 g8-f6 "  ' P-Q4  N-KB3  <=.3
            Select Case Rand
            Case Is > 0.5
               P$ = Mid$(SOpenString$(4), 7, 5): MakeOMove P$, Index ': Exit Function
            Case Is > 0.3
               P$ = Mid$(SOpenString$(5), 7, 5): MakeOMove P$, Index ': Exit Function
            Case Else
               P$ = Mid$(SOpenString$(6), 7, 5): MakeOMove P$, Index ': Exit Function
            End Select
      Case "g1-f3 "
'2
'SOpenString$(7) = "g1-f3 g8-f6 "  ' N-KB3 N-KB3  >=.5
'SOpenString$(8) = "g1-f3 d7-d5 "  ' N-KB3 P-Q4   <.5
            Select Case Rand
            Case Is > 0.5
               P$ = Mid$(SOpenString$(7), 7, 5): MakeOMove P$, Index ': Exit Function
            Case Else
               P$ = Mid$(SOpenString$(8), 7, 5): MakeOMove P$, Index ': Exit Function
            End Select
      Case "b1-c3 "
'2
'SOpenString$(9) = "b1-c3 e7-e5 "  ' N-QB3 P-K4   >=.5
'SOpenString$(10) = "b1-c3 b8-c6 " ' N-QB3 N-QB3  <.5
            Select Case Rand
            Case Is > 0.5
               P$ = Mid$(SOpenString$(9), 7, 5): MakeOMove P$, Index ': Exit Function
            Case Else
               P$ = Mid$(SOpenString$(10), 7, 5): MakeOMove P$, Index ': Exit Function
            End Select
      Case Else
'1s  fixed responses
'SOpenString$(11) = "e2-e3 e7-e5 " ' P-K3 P-K4
'SOpenString$(12) = "d2-d3 d7-d5 " ' P-Q3 P-Q4
'SOpenString$(13) = "g2-g3 e7-e5 " ' P-KN3 P-K4
'SOpenString$(14) = "b2-b3 d7-d5 " ' P-QN3 P-Q4
'SOpenString$(15) = "c2-c4 e7-e5 " ' P-QB4 P-K4
'SOpenString$(16) = "c2-c3 e7-e5 " ' P-QB3 P-K4
            For k = 11 To 16
               If Left$(OpenString$, 6) = Left$(SOpenString$(k), 6) Then
                  P$ = Mid$(SOpenString$(k), 7, 5): MakeOMove P$, Index ': Exit Function
               End If
            Next k
      End Select
      
   Case 2   ' Black's 2nd
      Select Case OpenString$
      Case "e2-e4 e7-e5 g1-f3 "
'2
'SOpenString$(21) = "e2-e4 e7-e5 g1-f3 b8-c6"  ' 1. P-K4 P-K4 2. N-KB3 N-QB3   >.5
'SOpenString$(22) = "e2-e4 e7-e5 g1-f3 d7-d6"  ' 1. P-K4 P-K4 2. N-KB3 P-Q3    <=.5
            Select Case Rand
            Case Is > 0.5
               P$ = Mid$(SOpenString$(21), 19, 5):
               MakeOMove P$, Index ': Exit Function
            Case Else
               P$ = Mid$(SOpenString$(22), 19, 5):
               MakeOMove P$, Index ': Exit Function
            End Select
      Case "d2-d4 d7-d5 g1-f3 "
'2
'SOpenString$(23) = "d2-d4 d7-d5 g1-f3 g8-f6"  ' 1. P-Q4 P-Q4 2. N-KB3 N-KB3   >.5
'SOpenString$(24) = "d2-d4 d7-d5 g1-f3 e7-e6"  ' 1. P-Q4 P-Q4 2. N-KB3 P-K3    <=.5
            Select Case Rand
            Case Is > 0.5
               P$ = Mid$(SOpenString$(23), 19, 5):
               MakeOMove P$, Index ': Exit Function
            Case Else
               P$ = Mid$(SOpenString$(24), 19, 5):
               MakeOMove P$, Index ': Exit Function
            End Select
      
      End Select
   Case Else
      MoveString$ = ""
   End Select
   If MoveString$ = "" Then BlackOMoves = False
   
End Function

Public Function WhiteOMoves(Index As Integer) As Boolean
' Public MoveString$
' Public HalfMove, OpenString$,SOpenString$()
Dim Rand As Single
Dim P$
   MoveString$ = ""
   WhiteOMoves = True
   Randomize
   Rand = Rnd
   Select Case HalfMove
   Case -1
'SOpenString$(40) = "e2-e4 "  ' P-K4   >=.5
'SOpenString$(41) = "d2-d4 "  ' P-Q4
         Select Case Rand
         Case Is >= 0.5
            P$ = SOpenString$(40): MakeOMove P$, Index ': Exit Function
         Case Else
            P$ = SOpenString$(41): MakeOMove P$, Index ': Exit Function
         End Select
   Case 1
      Select Case OpenString$
      Case "e2-e4 e7-e5 "
'SOpenString$(42) = "e2-e4 e7-e5 g1-f3 "  ' 1. P-K4 P-K4 2. N-KB3    >.5
'SOpenString$(43) = "e2-e4 e7-e5 b1-c3 "  ' 1. P-K4 P-K4 2. N-QB3
         Select Case Rand
         Case Is >= 0.5
            P$ = Mid$(SOpenString$(42), 13, 5): MakeOMove P$, Index ': Exit Function
         Case Else
            P$ = Mid$(SOpenString$(43), 13, 5): MakeOMove P$, Index ': Exit Function
         End Select
      Case "d2-d4 d7-d5 "
'SOpenString$(44) = "d2-d4 d7-d5 g1-f3 "  ' 1. P-Q4 P-Q4 2. N-KB3    >.5
'SOpenString$(45) = "d2-d4 d7-d5 b1-c3 "  ' 1. P-Q4 P-Q4 2. N-QB3
         Select Case Rand
         Case Is >= 0.5
            P$ = Mid$(SOpenString$(44), 13, 5): MakeOMove P$, Index ': Exit Function
         Case Else
            P$ = Mid$(SOpenString$(45), 13, 5): MakeOMove P$, Index ': Exit Function
         End Select
      End Select
   Case Else
      MoveString$ = ""
   End Select
   
   If MoveString$ = "" Then WhiteOMoves = False

End Function

Public Sub MakeOMove(M$, Index As Integer)
' For MoveBarred set:-
' Public RowS As Long, ColS As Long
' Public RowE As Long, ColE As Long
' ie MoveBarred(PieceNum As Long, Index As Integer) As Boolean
' ie NB Public Valid RowS,ColS & RowE,RowE (Start,End) MUST BE SET BEFORE ENTRY!

' Public MoveString$

' Enter with MoveString$ = ""
' EG
' M$ e7-e5
' bRCBoard(2, 3, Index) = 0
' bRCBoard(4, 3, Index) = WPn
' Asc("a") = 97
Dim P$
Dim PN As Long
Dim DPN As Long
   ColS = Asc(Mid$(M$, 1, 1)) - 96
   RowS = Val(Mid$(M$, 2, 1))
   ColE = Asc(Mid$(M$, 4, 1)) - 96
   RowE = Val(Mid$(M$, 5, 1))
   PN = bRCBoard(RowS, ColS, Index)
   DPN = bRCBoard(RowE, ColE, Index)
   If PN <> 0 And DPN = 0 Then
      If Not MoveBarred(PN, Index) Then
         If PN <= 6 Then ' White piece
            If Not aWKingInCheck Then
               bRCBoard(RowS, ColS, Index) = 0
               bRCBoard(RowE, ColE, Index) = PN
               ConvPNtoPNDescrip PN, P$
               MoveString$ = P$ & " " & Trim$(M$)
            End If
         Else  ' BlackPiece
            If Not aBKingInCheck Then
               bRCBoard(RowS, ColS, Index) = 0
               bRCBoard(RowE, ColE, Index) = PN
               ConvPNtoPNDescrip PN, P$
               MoveString$ = P$ & " " & Trim$(M$)
            End If
         End If
       End If
   End If
End Sub

Public Sub FillScoreArray(bARR() As Integer, a$, Optional REV As Long = 0)
Dim N As Long
Dim R As Long, C As Long
Dim RR As Long
   N = 1
      For R = 8 To 1 Step -1
      RR = R
      If REV > 0 Then RR = 9 - R
      For C = 1 To 8
         bARR(RR, C) = Val(Mid$(a$, N, 3))
         N = N + 4
      Next C
      Next R
End Sub

Public Sub FillOpenArray(bARR() As Byte, a$)
Dim N As Long
Dim R As Long, C As Long
   N = 1
      For R = 8 To 1 Step -1
      For C = 1 To 8
         bARR(R, C) = Val(Mid$(a$, N, 3))
         N = N + 4
      Next C
      Next R
End Sub


