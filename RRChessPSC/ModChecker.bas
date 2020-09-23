Attribute VB_Name = "ModChecker"
' ModChecker.bas  ~RRChess~

Option Explicit

Private aGLOBEP As Boolean ' En passant
Private BTBoard() As Byte   ' BlockTestBoard

Public Function LegalMove(M$) As Boolean
' Only used for Human moves on Board 0
' IN: eg M$ = "WQ a1-h8"
Dim PN As Long ' PieceNum
   Message$ = ""
   LegalMove = False
   If Len(M$) < 8 Then Exit Function
   ' Find RowSE/ColSE from M$
   RowS = Val(Mid$(M$, 5, 1))
   ColS = Asc(Mid$(M$, 4)) - 96 ' a = 97
   RowE = Val(Mid$(M$, 8, 1))
   ColE = Asc(Mid$(M$, 7)) - 96 ' a = 97
   'IndexS = 8 * (8 - RowS) + ColS - 1
   'IndexE = 8 * (8 - RowE) + ColE - 1
   PN = bRCBoard(RowS, ColS, 0)  ' NB Board 0
   
   If Not LegalDirection(M$) Then Exit Function
   ' Legal direction - now test if move barred
   If MoveBarred(PN, 0) Then Exit Function ' NB Board 0
   LegalMove = True
End Function

Private Function LegalDirection(M$) As Boolean
' Only used for Human moves
' eg M$ = "WQ a1-h8"
' NB Public Valid RowS,ColS & RowE,RowE (Start,End) MUST BE SET BEFORE ENTRY!
   LegalDirection = True
   Select Case Left$(M$, 2)
   Case "WR", "BR"
      If RowS = RowE Then Exit Function
      If ColS = ColE Then Exit Function
   Case "WB", "BB"
      If Abs(RowE - RowS) = Abs(ColE - ColS) Then Exit Function
   Case "WN", "BN"
      If Abs(ColE - ColS) = 2 Then
         If Abs(RowE - RowS) = 1 Then Exit Function
      ElseIf Abs(RowE - RowS) = 2 Then
         If Abs(ColE - ColS) = 1 Then Exit Function
      End If
   Case "WQ", "BQ"
      If RowS = RowE Then Exit Function
      If ColS = ColE Then Exit Function
      If Abs(RowE - RowS) = Abs(ColE - ColS) Then Exit Function
   Case "WK", "BK"
      If Abs(RowE - RowS) = 0 Then
         If Abs(ColE - ColS) = 1 Then Exit Function
         If Abs(ColE - ColS) = 2 Then Exit Function
      ElseIf Abs(ColE - ColS) = 0 Then
         If Abs(RowE - RowS) = 1 Then Exit Function
      End If
      If Abs(RowE - RowS) = Abs(ColE - ColS) Then
         If Abs(ColE - ColS) = 1 Then Exit Function
      End If
   Case "WP"
      If ColS = ColE Then
         If (RowE - RowS) = 1 Then Exit Function
         If RowS = 2 And RowE = 4 Then Exit Function
      Else   ' Diagonal take or en passant take
         If (RowE - RowS) = 1 Then
            If Abs(ColE - ColS) = 1 Then Exit Function
         End If
      End If
   Case "BP"
      If ColS = ColE Then
         If (RowS - RowE) = 1 Then Exit Function
         If RowS = 7 And RowE = 5 Then Exit Function
      Else  ' Diagonal take or en passant take
         If (RowS - RowE) = 1 Then
            If Abs(ColE - ColS) = 1 Then Exit Function
         End If
      End If
   End Select
   LegalDirection = False
End Function

Public Function MoveBarred(PieceNum As Long, Index As Integer) As Boolean
' NB Public Valid RowS,ColS & RowE,RowE (Start,End) MUST BE SET BEFORE ENTRY!

' PieceNum  @  RowS,ColS on Board Index looking for approval to move to RowE,ColE.
' Is it blocked before End by any color or by same color @ End.
' Also test/flag castling, en passant & promotion
' Messages for Human player(s).
Dim k As Long
Dim PN As Long    ' Source piece number
Dim DPN As Long   ' Dest piece number
' W/B King positions
Dim rkw As Long, ckw As Long
Dim rkb As Long, ckb As Long
Dim aEnPassant As Boolean
Dim aCast As Boolean
   'On Error GoTo ExitBarred
   aWKingInCheck = False
   aBKingInCheck = False
   MoveBarred = True
   Message$ = ""
   
   If PieceNum = 0 Then Exit Function
   
   If RowS < 1 Or RowS > 8 Then Exit Function
   If RowE < 1 Or RowE > 8 Then Exit Function
   If ColS < 1 Or ColS > 8 Then Exit Function
   If ColE < 1 Or ColE > 8 Then Exit Function
   
   ReDim BTBoard(1 To 8, 1 To 8) As Byte   ' BlockTestBoard
   
   CopyMemory BTBoard(1, 1), bRCBoard(1, 1, Index), 64
   
   ' Legal moves might still put own K in check!
   ' So before checking this need to save settings
   ' WPawn(1 To 8), BPawn(1 To 8) &
   ' aWPawnPromotion, aBPawnPromotion etc
   ' and restore them if own king put in check
   ' because board will be restored.

   SaveALLBOOLS Index, MaxIndex + 1
   
   aEnPassant = False
   aCast = False
   
   If PieceNum <= WPn Then
      PlayColor$ = "W": OppColor$ = "B"
   Else
      PlayColor$ = "B": OppColor$ = "W"
   End If
   
   Select Case PieceNum
   Case WRn, BRn   ' Rooks horz or vert
      MoveBarred = Not IsHorzVertOK(Index)
      ' Not blocked
      ' Block castling?
      If Not MoveBarred Then
         If PlayColor$ = "W" Then  'WR
            If RowS = 1 Then
               If ColS = 1 Then
                  WKRR(WQRMoved, Index) = 1 ' WQR Moved
               ElseIf ColS = 8 Then
                  WKRR(WKRMoved, Index) = 1 ' WKR Moved
               End If
            End If
         Else   ' BR
            If RowS = 8 Then
               If ColS = 1 Then
                  BKRR(BQRMoved, Index) = 1 ' BQR Moved
               ElseIf ColS = 8 Then
                  BKRR(BKRMoved, Index) = 1 ' BKR Moved
               End If
            End If
         End If
      End If
      
   Case WNn, BNn   ' Knights
      ' Get landing square
      PN = bRCBoard(RowE, ColE, Index)
      If PN <> 0 Then
         If PlayColor$ = "W" Then
            If PN >= BRn Then MoveBarred = False
         Else  ' PColor$= "B"
            If PN <= WPn Then MoveBarred = False
         End If
      Else
         MoveBarred = False  ' Empty landing square
      End If
   
   Case WBn, BBn   ' Bishops  diagonals
      MoveBarred = Not AreDiaginalsOK(Index)
      
   Case WQn, BQn   ' Queens
      If RowE = RowS Or ColE = ColS Then
         MoveBarred = Not IsHorzVertOK(Index)
      Else  ' Must be diagonal
         MoveBarred = Not AreDiaginalsOK(Index)
      End If
   
   Case WPn   ' White pawns
      If ColE = ColS Then
         MoveBarred = Not IsWP_UpOK(Index)  ' Destination empty. Maybe row 8 promotion
         ' WP move up
         If Not MoveBarred Then
            'WEPP set all false
            For k = 0 To 3
               WEPP(k, Index) = 0  ' cancel WEPPs
            Next k
            If RowS = 2 And RowE = 4 Then
               WEPP(WEPSet, Index) = ColS
               'aEnPassant = True
            End If
            If RowS = 7 And RowE = 8 Then WEPP(WEPProm, Index) = 1  ' WPawn Promotion = True
         End If
      ElseIf Abs(ColE - ColS) = 1 Then ' Must be up diagonal 1 \/
         MoveBarred = Not IsWP_DiagonalOK(Index)  ' Empty en passant or opp color
         If Not MoveBarred Then
            'WEPP set all false
            For k = 0 To 3
               WEPP(k, Index) = 0  ' cancel WEPPs
            Next k
            If bRCBoard(RowS + 1, ColE, Index) = 0 Then  ' Only en passant allowed
               If RowE = 6 Then
                  If BEPP(BEPSet, Index) = ColE Then  ' B EnPassantSET=True
                     aEnPassant = True
                     WEPP(WEPOK, Index) = 1
                     BEPP(BEPSet, Index) = 0
                  Else
                     MoveBarred = True
                  End If
               Else
                  MoveBarred = True
               End If
            Else    ' Must be taking opp color & could promote
               If RowS = 7 And RowE = 8 Then WEPP(WEPProm, Index) = 1 ' W PawnPromotion = True
            End If
         End If
      End If
   
   Case BPn   ' Black pawns
      If ColE = ColS Then
         MoveBarred = Not IsBP_DnOK(Index)  ' Destination empty. Maybe row 8 promotion
         ' BP move down
         If Not MoveBarred Then
            'BEPP set all false
            For k = 0 To 3
               BEPP(k, Index) = 0  ' cancel BEPPs
            Next k
            If RowS = 7 And RowE = 5 Then
               BEPP(BEPSet, Index) = ColS
               'aEnPassant = True
               BEPP(BEPOK, Index) = 0
            End If
            If RowS = 2 And RowE = 1 Then BEPP(BEPProm, Index) = 1 ' B PawnPromotion = True
         End If
      ElseIf Abs(ColE - ColS) = 1 Then ' Must be up diagonal 1 \/
         MoveBarred = Not IsBP_DiagonalOK(Index)  ' Empty en passant or opp color
         If Not MoveBarred Then
            'BEPP set all false
            For k = 0 To 3
               BEPP(k, Index) = 0  ' cancel BEPPs
            Next k
            If bRCBoard(RowS - 1, ColE, Index) = 0 Then  ' Only en passant allowed
               If RowE = 3 Then
                  If WEPP(WEPSet, Index) = ColE Then   ' W EnPassantSET=True
                     aEnPassant = True
                     BEPP(BEPOK, Index) = 1
                     WEPP(WEPSet, Index) = 0
                  Else
                     MoveBarred = True
                  End If
               Else
                  MoveBarred = True
               End If
            Else  ' Must be taking opp color & could promote
               If RowS = 2 And RowE = 1 Then BEPP(BEPProm, Index) = 1 ' B PawnPromotion = True
            End If
         End If
      End If
   
   Case WKn, BKn   ' Kings
      If RowE = RowS Or ColE = ColS Then
         If Abs(ColE - ColS) = 1 Or Abs(RowE - RowS) = 1 Then  ' Horz or vert 1
            MoveBarred = Not IsHorzVertOK(Index)
            If Not MoveBarred Then
               If PlayColor$ = "W" Then WKRR(WKMoved, Index) = 1 Else BKRR(BKMoved, Index) = 1
               ' WK Moved = OR BK Moved = True
            End If
         End If
      ElseIf Abs(RowE - RowS) = Abs(ColE - ColS) Then
         If Abs(ColE - ColS) = 1 Then   ' Diagonal 1
            MoveBarred = Not AreDiaginalsOK(Index)
            If Not MoveBarred Then
               If PlayColor$ = "W" Then WKRR(WKMoved, Index) = 1 Else BKRR(BKMoved, Index) = 1
               ' WK Moved = True OR BK Moved = True
            End If
         End If
      End If
      
      ' Will these K moves put own K in check?
      Message$ = ""
      If Not MoveBarred Then
            If PlayColor$ = "W" Then   ' Check if WK in check
               If RC_Targetted(RowE, ColE, "B", Index) Then
                  Message$ = "White King put in check" ': Stop
                  aWKingInCheck = True
               End If
            Else  ' Check if BK in check
               If RC_Targetted(RowE, ColE, "W", Index) Then
                  Message$ = "Black King put in check" ': Stop
                  aBKingInCheck = True
               End If
            End If
      End If
      
      If Message$ <> "" Then GoTo ExitBarred ' King moved & into check so need to cancel
                                             ' WKRR(WKMoved, Index) = 1 Or BKRR(BKMoved, Index) = 1
      
      ' Block castling?
      If Abs(ColE - ColS) = 2 And (RowE = RowS) And (RowE = 1 Or RowE = 8) Then ' Castling attempt
               
         Cast(0, Index) = 0 ' Castling(Index) = False
         Cast(1, Index) = 0 ' KSideCastling(Index) = False
         Cast(2, Index) = 0 ' QSideCastling(Index) = False
               
         ' Check K positions & R presence
         If ColS <> 5 Then GoTo ExitBarred
         If (ColE <> 7 And ColE <> 3) Then GoTo ExitBarred
         If PlayColor$ = "W" Then
            If RowS <> 1 Then GoTo ExitBarred
            If ColE = 7 Then  ' ie K to col 7
               If bRCBoard(1, 8, Index) <> WRn Then GoTo ExitBarred
            End If
            If ColE = 3 Then   ' ie K to col 3
               If bRCBoard(1, 1, Index) <> WRn Then GoTo ExitBarred
               If bRCBoard(1, 2, Index) <> 0 Then GoTo ExitBarred
            End If
         Else   ' PlayColor$ = "B"
            If RowS <> 8 Then GoTo ExitBarred
            If ColE = 7 Then
               If bRCBoard(8, 8, Index) <> BRn Then GoTo ExitBarred
            End If
            If ColE = 3 Then
               If bRCBoard(8, 1, Index) <> BRn Then GoTo ExitBarred
               If bRCBoard(8, 2, Index) <> 0 Then GoTo ExitBarred
            End If
         End If
      
         MoveBarred = Not IsHorzVertOK(Index)
         
         If MoveBarred Then GoTo ExitBarred
      
         MoveBarred = True ' Again
         Message$ = ""
         ' Set Message$ if illegal
         If PlayColor$ = "W" Then
            If WKRR(WKMoved, Index) = 1 Then Message$ = " White king has been moved"
            If WKRR(WKMoved, Index) > 1 Then Message$ = " Castling done"
            If ColE > ColS Then  ' W KingSide
               If WKRR(WKRMoved, Index) = 1 Then Message$ = " White king rook has been moved"
            Else    ' W QueenSide
               If WKRR(WQRMoved, Index) = 1 Then Message$ = " White queen rook has been moved"
            End If
         Else  ' PlayColor$ = "B"
            If BKRR(BKMoved, Index) = 1 Then Message$ = " Black king has been moved"
            If BKRR(WKMoved, Index) > 1 Then Message$ = " Castling done"
            If ColE > ColS Then  ' B KingSide
               If BKRR(BKRMoved, Index) = 1 Then Message$ = " Black king rook has been moved"
            Else    ' B QueenSide
               If BKRR(BQRMoved, Index) = 1 Then Message$ = " Black queen rook has been moved"
            End If
         End If   ' If PlayColor$ = "W" Then, Else  ' PlayColor$ = "B"
   
         If Message$ <> "" Then GoTo ExitBarred
      
         If Message$ = "" Then
            ' OK so far. Now to check if king out of, cross & into check
            If PlayColor$ = "W" Then
               ' Check if WK in check
               If RC_Targetted(RowS, ColS, "B", Index) Then Message$ = "White King in check, castling out of check not allowed!"
               ' Check if landing square in check
               If RC_Targetted(RowE, ColE, "B", Index) Then Message$ = "White King castling into check!"
               'OK so far. Check if crossing check
               If ColE > ColS Then  ' W KingSide Row 1 Col 5 to 7 Check Col 6
                  If RC_Targetted(1, 6, "B", Index) Then Message$ = "White King castling across check not allowed!"
               Else  ' W QueenSide Row 1 Col 5 to 3 Check Col 4
                  If RC_Targetted(1, 4, "B", Index) Then Message$ = "White King castling across check not allowed!"
               End If
            Else  ' PlayColor$ = "B"
               ' Check if BK in check
               If RC_Targetted(RowS, ColS, "W", Index) Then Message$ = "Black King in check, castling out of check not allowed!"
               ' Check if landing square in check
               If RC_Targetted(RowE, ColE, "W", Index) Then Message$ = "Black King castling into check!"
               'OK so far. Check if crossing check
               If ColE > ColS Then  ' B KingSide  Row 8 Col 5 to 7 Check Col 6
                  If RC_Targetted(8, 6, "W", Index) Then Message$ = "Black King castling across check not allowed!"
               Else  ' B QueenSide  Row 8 Col 5 to 3 Check Col 4
                  If RC_Targetted(8, 4, "W", Index) Then Message$ = "Black King castling across check not allowed!"
               End If
            End If   ' If PlayColor$ = "W" Then, Else  ' PlayColor$ = "B"
      
         End If   ' If Message$ = "" Then
   
         If Message$ <> "" Then GoTo ExitBarred
         
         ' Castling OK now
         MoveBarred = False
         aCast = True
         If ColE > ColS Then  ' KSideCastling
            If PlayColor$ = "W" Then
               Cast(WKSCastOK, Index) = 1   ' W KSideCastling = True
            Else
               Cast(BKSCastOK, Index) = 1   ' B KSideCastling = True
            End If
         Else   ' QSideCastling
            If PlayColor$ = "W" Then
               Cast(WQSCastOK, Index) = 1   ' W QSideCastling = True
            Else
               Cast(BQSCastOK, Index) = 1   ' B QSideCastling = True
            End If
         End If
   
      End If   ' If Abs(ColE - ColS) = 2 And (RowE = RowS) Then  ' Castling attempt
      
   End Select
   
   ' Check if King put in discovered check
   If Not MoveBarred Then
      If Message$ = "" Then
         If Not aCast Then  ' Castling dealt with. NB For this Prom Pawn to 8th or 1st row is sufficient
            ' Make the move   'PieceNum, RowS,ColS -> RowE,ColE
            DPN = bRCBoard(RowE, ColE, Index)
            If DPN = 5 Or DPN = 11 Then GoTo ExitBarred
            bRCBoard(RowE, ColE, Index) = PieceNum
            bRCBoard(RowS, ColS, Index) = 0
            If aEnPassant Then   ' also ??
               bRCBoard(RowS, ColE, Index) = 0
            End If
            
            '''' ERROR CHECK '''''
            If Not FindKingRC("W", rkw, ckw, Index) Then
                  Message$ = "No White King in MoveBarred ??"
                  MsgBox Message$, vbCritical, "ERROR"
                  GoTo ExitBarred
            End If
            If Not FindKingRC("B", rkb, ckb, Index) Then
                  Message$ = "No Black King in MoveBarred ??"
                  MsgBox Message$, vbCritical, "ERROR"
                  GoTo ExitBarred
            End If
            
            If PlayColor$ = "W" Then
               
               If WEPP(WEPProm, Index) = 1 Then
                  bRCBoard(RowE, ColE, Index) = WQn
                  PromPiece$ = "Q"
               End If
               If RC_Targetted(rkw, ckw, "B", Index) Then
                  Message$ = "White King in check"
                  aWKingInCheck = True
                  'MoveBarred = True  ' No, MATER may need to run through this
                  ' TESTED again in MoveCounter for comp play
               End If
               
               If RC_Targetted(rkb, ckb, "W", Index) Then
                  aBKingInCheck = True
               End If
            
            Else  ' PlayColor$ = "B"
               
               If BEPP(BEPProm, Index) = 1 Then
                  bRCBoard(RowE, ColE, Index) = BQn
                  PromPiece$ = "Q"
               End If
               If RC_Targetted(rkb, ckb, "W", Index) Then
                  Message$ = "Black King in check"
                  aBKingInCheck = True
                  'MoveBarred = True  ' No, MATER may need to run through this
                  ' TESTED again in MoveCounter for comp play
               End If
   
               If RC_Targetted(rkw, ckw, "B", Index) Then
                  aWKingInCheck = True
               End If
            
            End If
         End If
      End If
   End If
   
   If MoveBarred Then GoTo ExitBarred  'RestoreALLBOOLS MaxIndex + 1, Index
   
   If Not aEnPassant Then
      If PlayColor$ = "B" Then
         WEPP(WEPSet, Index) = 0
         BEPP(BEPOK, Index) = 0
      Else
         BEPP(BEPSet, Index) = 0
         WEPP(WEPOK, Index) = 0
      End If
   End If
   ' Restore bRCBoard(,,Index)
   CopyMemory bRCBoard(1, 1, Index), BTBoard(1, 1), 64
   Exit Function
'===============
ExitBarred:
   MoveBarred = True
   RestoreALLBOOLS MaxIndex + 1, Index
   CopyMemory bRCBoard(1, 1, Index), BTBoard(1, 1), 64
End Function

Private Function IsHorzVertOK(Index As Integer) As Boolean
' NB. Public PlayColor$, OppColor$  ' Piece color, Opposite color set
' NB. Public Valid RowS,ColS & RowE,RowE (Start,End) MUST BE SET BEFORE ENTRY!
' WR or BR  also  WQ or BQ if ColE = ColS (Up/Dn) or RowE = RowS (Left/Right)
' Piece number @ R,C Incrementing row, Incrementing col
' Track from horz/Vert squares next to RowS,ColS to RowE,ColE
Dim PN As Long, R As Long, C As Long
   IsHorzVertOK = False ' Not OK
   
   If RowE > RowS Then  ' Move UP
      C = ColS
      For R = RowS + 1 To RowE
         PN = bRCBoard(R, C, Index) ' bRCBoard number @ R,C
         If PN <> 0 Then  ' a piece @ R,C
            If PlayColor$ = "W" Then
               If PN <= WPn Then Exit Function  ' White piece hits another white piece
               If R < RowE Then Exit Function   ' Black piece @ R,C before destination RowE
            Else     ' Black piece @ R,C
               If PN >= BRn Then Exit Function  ' Black piece hits another black piece
               If R < RowE Then Exit Function   ' White piece @ R,C before destination RowE
            End If
         End If
      Next R
      IsHorzVertOK = True  ' Destination square OK
      Exit Function
   End If
   
   If RowE < RowS Then  ' Move DOWN
      C = ColS
      For R = RowS - 1 To RowE Step -1
         PN = bRCBoard(R, C, Index) ' bRCBoard number @ R,C
         If PN <> 0 Then
            If PlayColor$ = "W" Then
               If PN <= WPn Then Exit Function  ' White piece hits another white piece
               If R > RowE Then Exit Function   ' Black piece @ R,C before destination RowE
            Else     ' Black piece @ R,C
               If PN >= BRn Then Exit Function  ' Black piece hits another black piece
               If R > RowE Then Exit Function   ' White piece @ R,C before destination RowE
            End If
         End If
      Next R
      IsHorzVertOK = True  ' Destination square OK
      Exit Function
   End If

   If ColE < ColS Then  ' Move LEFT
      R = RowS
      For C = ColS - 1 To ColE Step -1
         PN = bRCBoard(R, C, Index) ' bRCBoard number @ R,C
         If PN <> 0 Then
            If PlayColor$ = "W" Then
               If PN <= WPn Then Exit Function  ' White piece hits another white piece
               If C > ColE Then Exit Function   ' Black piece @ R,C before destination RowE
            Else     ' Black piece @ R,C
               If PN >= BRn Then Exit Function  ' Black piece hits another black piece
               If C > ColE Then Exit Function   ' White piece @ R,C before destination RowE
            End If
         End If
      Next C
      IsHorzVertOK = True  ' Destination square OK
      Exit Function
   End If
   
   If ColE > ColS Then  ' Move RIGHT
      R = RowS
      For C = ColS + 1 To ColE
         PN = bRCBoard(R, C, Index) ' bRCBoard number @ R,C
         If PN <> 0 Then
            If PlayColor$ = "W" Then
               If PN <= WPn Then Exit Function  ' White piece hits another white piece
               If C < ColE Then Exit Function   ' Black piece @ R,C before destination RowE
            Else     ' Black piece @ R,C
               If PN >= BRn Then Exit Function  ' Black piece hits another black piece
               If C < ColE Then Exit Function   ' White piece @ R,C before destination RowE
            End If
         End If
      Next C
      IsHorzVertOK = True  ' Destination square OK
      Exit Function
   End If
End Function

Private Function AreDiaginalsOK(Index As Integer) As Boolean
' Public PlayColor$, OppColor$  ' Piece color, Opposite color
' WB or BB  also  WQ or BQ if NOT (ColE = ColS (Up/Dn) And RowE = RowS (Left/Right))
' Piece number @ R,C Incrementing row, Incrementing col
' Track from diagonal squares next to RowS,ColS to RowE,ColE
Dim PN As Long, R As Long, C As Long
   AreDiaginalsOK = False ' Not OK
   R = RowS: C = ColS
   If ColE > ColS And RowE > RowS Then ' Track along Top RIGHT diagonal
      Do ' TR
         C = C + 1
         R = R + 1
         If C > ColE Or R > RowE Then Exit Function  'Beyond destination ?
         PN = bRCBoard(R, C, Index) ' bRCBoard number @ R,C
         If PN <> 0 Then
            If PN <= WPn Then   ' White piece @ R,C
               If PlayColor$ = "W" Then
                  Exit Function ' Hit white piece
               Else ' PColor$= "B"
                  If R < RowE Then Exit Function ' Black piece @ R,C before destination RowE,ColE
               End If
            Else     ' Black piece @ R,C
               If PlayColor$ = "B" Then
                  Exit Function ' Hit black piece
               Else ' Pcolor$= "W"
                  If R < RowE Then Exit Function ' White piece @ R,C before destination RowE,ColE
               End If
            End If
         End If
         If C = ColE And R = RowE Then Exit Do
      Loop
      AreDiaginalsOK = True ' Destination square OK
      Exit Function
   End If

   If ColE < ColS And RowE > RowS Then ' Top LEFT
      Do ' TL
         C = C - 1
         R = R + 1
         If C < ColE Or R > RowE Then Exit Function     'Beyond destination ?
         PN = bRCBoard(R, C, Index) ' bRCBoard number @ R,C
         If PN <> 0 Then
            If PN <= WPn Then   ' White piece @ R,C
               If PlayColor$ = "W" Then
                  Exit Function ' Hit white piece
               Else ' PColor$= "B"
                  If R < RowE Then Exit Function ' Black piece @ R,C before destination RowE,ColE
               End If
            Else     ' Black piece @ R,C
               If PlayColor$ = "B" Then
                  Exit Function ' Hit black piece
               Else ' Pcolor$= "W"
                  If R < RowE Then Exit Function ' White piece @ R,C before destination RowE,ColE
               End If
            End If
         End If
         If C = ColE And R = RowE Then Exit Do
      Loop
      AreDiaginalsOK = True ' Destination square OK
      Exit Function
   End If
   
   If ColE > ColS And RowE < RowS Then ' Bottom RIGHT
      Do ' BR
         C = C + 1
         R = R - 1
         If C > ColE Or R < RowE Then Exit Function     'Beyond destination ?
         PN = bRCBoard(R, C, Index) ' bRCBoard number @ R,C
         If PN <> 0 Then
            If PN <= WPn Then   ' White piece @ R,C
               If PlayColor$ = "W" Then
                  Exit Function ' Hit white piece
               Else ' PColor$= "B"
                  If C < ColE Then Exit Function ' Black piece @ R,C before destination RowE,ColE
               End If
            Else     ' Black piece @ R,C
               If PlayColor$ = "B" Then
                  Exit Function ' Hit black piece
               Else ' Pcolor$= "W"
                  If C < ColE Then Exit Function ' White piece @ R,C before destination RowE,ColE
               End If
            End If
         End If
         If C = ColE And R = RowE Then Exit Do
      Loop
      AreDiaginalsOK = True ' Destination square OK
      Exit Function
   End If
   
   If ColE < ColS And RowE < RowS Then ' Bottom LEFT
      Do ' BL
         C = C - 1
         R = R - 1
         If C < ColE Or R < RowE Then Exit Function     'Beyond destination ?
         PN = bRCBoard(R, C, Index) ' bRCBoard number @ R,C
         If PN <> 0 Then
            If PN <= WPn Then   ' White piece @ R,C
               If PlayColor$ = "W" Then
                  Exit Function ' Hit white piece
               Else ' PColor$= "B"
                  If C > ColE Then Exit Function ' Black piece @ R,C before destination RowE,ColE
               End If
            Else     ' Black piece @ R,C
               If PlayColor$ = "B" Then
                  Exit Function ' Hit black piece
               Else ' Pcolor$= "W"
                  If C > ColE Then Exit Function ' White piece @ R,C before destination RowE,ColE
               End If
            End If
         End If
         If C = ColE And R = RowE Then Exit Do
      Loop
      AreDiaginalsOK = True ' Destination square OK
      Exit Function
   End If
End Function

Private Function IsWP_UpOK(Index As Integer) As Boolean
' WPawn UP
   IsWP_UpOK = False
   If bRCBoard(RowE, ColS, Index) <> 0 Then Exit Function         ' RowE up blocked
   ' End square clear
   If (RowE - RowS) = 2 And RowS <> 2 Then Exit Function          ' 2 move wrong
   If RowS = 2 And RowE = 4 Then    ' 2 up from row 2
      If bRCBoard(RowS + 1, ColS, Index) <> 0 Then Exit Function  ' Row 3 occupied
   End If
   IsWP_UpOK = True  ' WP clear to move up
End Function

Private Function IsWP_DiagonalOK(Index As Integer) As Boolean
' WPawn up diagonal
Dim PN As Long ' Piece number
   IsWP_DiagonalOK = False
   PN = bRCBoard(RowS + 1, ColE, Index)
   If PN <> 0 Then
      If PN <= WPn Then Exit Function  ' Diag. blocked by same color.
      ' Else OppColor
      IsWP_DiagonalOK = True
   Else  ' PN=0 possible en passant ?
      IsWP_DiagonalOK = True
   End If
End Function

Private Function IsBP_DnOK(Index As Integer) As Boolean
' BPawn down
   IsBP_DnOK = False
   If bRCBoard(RowE, ColS, Index) <> 0 Then Exit Function         ' RowE up blocked
   ' End square clear
   If (RowS - RowE) = 2 And RowS <> 7 Then Exit Function          ' 2 move wrong
   If RowS = 7 And RowE = 5 Then    ' 2 down from row 7
      If bRCBoard(RowS - 1, ColS, Index) <> 0 Then Exit Function  ' Row 6 occupied
   End If
   IsBP_DnOK = True   ' WP clear to move up
End Function

Private Function IsBP_DiagonalOK(Index As Integer) As Boolean
' BPawn
Dim PN As Long ' Piece number
   IsBP_DiagonalOK = False
   PN = bRCBoard(RowS - 1, ColE, Index)
   If PN <> 0 Then
      If PN >= BRn Then Exit Function  ' Diag. blocked by same color.
      ' Else OppColor
      IsBP_DiagonalOK = True
   Else  ' PN=0 possible en passant ?
      IsBP_DiagonalOK = True
   End If
End Function
'#### End of MoveBarred routines ###############################################################

Public Function RC_Targetted(R As Long, C As Long, WB$, Index As Integer, _
               Optional ExcludeKing As Boolean = False, Optional ExcludeDiagPawnMove As Boolean = False) As Boolean

' Returns:-
' Public AttackingPiece @
' Public rat As Long, cat As Long     ' r,c of AttackingPiece

' Enter with bRCBoard(1-8 columns, 8 - 1 rows) set
'            WB$ = "W" look for attack by White on Square(r,c) row & column to check
'         or WB$ = "B" look for attack by Black on Square(r,c) row & column to check

' ExcludeKing         = False DEFAULT Include attack by opposing king
' ExcludeKing         = True  Exclude attack by opposing king
' ExcludeDiagPawnMove = False DEFAULT Include diagonal Exclude forward pawn move
' ExcludeDiagPawnMove = True  Exclude diagonal Include forward pawn move
Dim k As Long
Dim rs As Long, cs As Long  ' test r & c
Dim s As Long ' Sign
'   RC_Targetted = False
'   If r < 1 Then Exit Function
'   If r > 8 Then Exit Function
'   If c < 1 Then Exit Function
'   If c > 8 Then Exit Function
   RC_Targetted = True
'   Attackng King
'   ReDim AttackType(1 To 6)     ' Unlikely to be more than 3
'   ReDim AttackTypeRow(1 To 6)  ' Unlikely to be more than 3
'   ReDim AttackTypeCol(1 To 6)  ' Unlikely to be more than 3
'   NumAttacking = 0
   
' NumAttackingSquare
' AttackPNOnSquare()
' rat(),cat()

   ' Check BN or WN [8] [2]
   
   If WB$ = "B" Then ' Look for attack by Black on Square(r,c)
      If NumBN > 0 Then AttackingPiece = BNn Else GoTo TestBRBQ
   Else  ' Look for attack by White on Square(r,c)
      If NumWN > 0 Then AttackingPiece = WNn Else GoTo TestBRBQ
   End If
   
   For rs = R - 2 To R + 2 Step 4   ' U/D 2
      If 1 <= rs And rs <= 8 Then
         For cs = C - 1 To C + 1 Step 2   ' L/R 1
            If 1 <= cs And cs <= 8 Then
               If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BN or WN
                  rat = rs: cat = cs
                  Exit Function
               End If
            
            End If
         Next cs
      End If
   Next rs
   For cs = C - 2 To C + 2 Step 4   ' L/R 2
      If 1 <= cs And cs <= 8 Then
         For rs = R - 1 To R + 1 Step 2   ' U/D 1
            If 1 <= rs And rs <= 8 Then
               If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BN or WN
                  rat = rs: cat = cs
                  Exit Function
               End If
            End If
         Next rs
      End If
   Next cs
   
   '---------------------------------------------------
TestBRBQ:
   ' Check BR & BQ [7 & 10] or WR & WQ [1] [4]
   For k = 0 To 1
      If k = 0 Then
         If WB$ = "B" Then ' Look for attack by Black on Square(r,c)
            If NumBR > 0 Then AttackingPiece = BRn Else GoTo Nextk
         Else  ' Look for attack by White on Square(r,c)
            If NumWR > 0 Then AttackingPiece = WRn Else GoTo Nextk
         End If
      Else  ' k=1
         If WB$ = "B" Then ' Look for attack by B on Square(r,c)
            If NumBQ > 0 Then AttackingPiece = BQn Else GoTo TestBBBQ
         Else  ' Look for attack by W on Square(r,c)
            If NumWQ > 0 Then AttackingPiece = WQn Else GoTo TestBBBQ
         End If
      End If
      ' Search Up
      cs = C
      If R < 8 Then
         For rs = R + 1 To 8
            If bRCBoard(rs, cs, Index) <> 0 Then
               If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BR or WR, BQ or WQ
                  rat = rs: cat = cs
                  Exit Function
               End If
               Exit For
            End If
         Next rs
      End If
      ' Dn
      If R > 1 Then
         For rs = R - 1 To 1 Step -1
            If bRCBoard(rs, cs, Index) <> 0 Then
               If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BR or WR, BQ or WQ
                  rat = rs: cat = cs
                  Exit Function
               End If
               Exit For
            End If
         Next rs
      End If
      ' Right
      rs = R
      If C < 8 Then
         For cs = C + 1 To 8
            If bRCBoard(rs, cs, Index) <> 0 Then
               If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BR or WR, BQ or WQ
                  rat = rs: cat = cs
                  Exit Function
               End If
               Exit For
            End If
         Next cs
      End If
      ' Left
      If C > 1 Then
         For cs = C - 1 To 1 Step -1
            If bRCBoard(rs, cs, Index) <> 0 Then
               If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BR or WR, BQ or WQ
                  rat = rs: cat = cs
                  Exit Function
               End If
               Exit For
            End If
         Next cs
      End If
Nextk:
   Next k
   '---------------------------------------------------
TestBBBQ:
   ' Check BB & BQ  [9 & 10]  WB & WQ [3 & 4]
   For k = 0 To 1
      If k = 0 Then
         If WB$ = "B" Then ' Look for attack by Black on Square(r,c)
            If NumBB > 0 Then AttackingPiece = BBn Else GoTo Nextkk
         Else  ' Look for attack by White on Square(r,c)
            If NumWB > 0 Then AttackingPiece = WBn Else GoTo Nextkk
         End If
      Else  ' k=1
         If WB$ = "B" Then ' Look for attack by B on Square(r,c)
            If NumBQ > 0 Then AttackingPiece = BQn Else GoTo TestBPWP
         Else  ' Look for attack by W on Square(r,c)
            If NumWQ > 0 Then AttackingPiece = WQn Else GoTo TestBPWP
         End If
      End If
      'Search TLeft
      rs = R + 1
      cs = C - 1
      Do
         If cs < 1 Then Exit Do
         If rs > 8 Then Exit Do
         If bRCBoard(rs, cs, Index) <> 0 Then
            If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BB or WB, BQ or WQ
               rat = rs: cat = cs
               Exit Function
            End If
            Exit Do
         End If
         rs = rs + 1
         cs = cs - 1
      Loop
      ' TRight
      rs = R + 1
      cs = C + 1
      Do
         If cs > 8 Then Exit Do
         If rs > 8 Then Exit Do
         If bRCBoard(rs, cs, Index) <> 0 Then
            If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BB or WB, BQ or WQ
               rat = rs: cat = cs
               Exit Function
            End If
            Exit Do
         End If
         rs = rs + 1
         cs = cs + 1
      Loop
      ' BRight
      rs = R - 1
      cs = C + 1
      Do
         If cs > 8 Then Exit Do
         If rs < 1 Then Exit Do
         If bRCBoard(rs, cs, Index) <> 0 Then
            If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BB or WB, BQ or WQ
               rat = rs: cat = cs
               Exit Function
            End If
            Exit Do
         End If
         rs = rs - 1
         cs = cs + 1
      Loop
      ' BLeft
      rs = R - 1
      cs = C - 1
      Do
         If cs < 1 Then Exit Do
         If rs < 1 Then Exit Do
         If bRCBoard(rs, cs, Index) <> 0 Then
            If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BB or WB, BQ or WQ
               rat = rs: cat = cs
               Exit Function
            End If
            Exit Do
         End If
         rs = rs - 1
         cs = cs - 1
      Loop
Nextkk:
   Next k
   '---------------------------------------------------
   
TestBPWP:
'            WB$ = "W" look for attack by White on Square(r,c) row & column to check
'         or WB$ = "B" look for attack by Black on Square(r,c) row & column to check
   
   ' Check BP or WP [12] [6]
   If WB$ = "B" Then ' Look for attack by Black on Square(r,c)
      ' Only an attack if square occupied by opp color
      If R >= 7 Then ' BP cannot attack square @ r
         GoTo TestKing
      End If
      If NumBP > 0 Then AttackingPiece = BPn Else GoTo TestKing
      s = 1       ' Attack Down, so pawn has to be @ r+1 down to attacked square
   Else  ' Look for attack by White on Square(r,c)
      ' Only an attack if square occupied by opp color
      If R <= 2 Then ' WP cannot attack square @ r
         GoTo TestKing
      End If
      If NumWP > 0 Then AttackingPiece = WPn Else GoTo TestKing
      s = -1    ' Attack Up, so pawn has to be @ r-1 down to attacked square
   End If
   
' ExcludeDiagPawnMove = False Include diagonal Exclude forward pawn move, default
' ExcludeDiagPawnMove = True  Exclude diagonal Include forward pawn move

   If ExcludeDiagPawnMove Then GoTo ForwardPawn
   
' DiagonalPawn

   aGLOBEP = False
   rs = R + s
   If rs >= 1 And rs <= 8 Then
      
      cs = C - 1  ' Check if pawn @ Left diagonal
      If cs >= 1 Then
'            WB$ = "W" look for attack by White on Square(r,c) row & column to check
'         or WB$ = "B" look for attack by Black on Square(r,c) row & column to check
         'If (WB$ = "W" And bRCBoard(rs, cs, Index) <= WPn) Or _
         '   (WB$ = "B" And bRCBoard(rs, cs, Index) >= BRn) Then
         If (WB$ = "W" And bRCBoard(rs, cs, Index) = WPn) Or _
            (WB$ = "B" And bRCBoard(rs, cs, Index) = BPn) Then
            If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BP or WP
               rat = rs: cat = cs
               Exit Function
            End If
         End If
      End If
      cs = C + 1  ' Check if pawn @ Right diagonal
      If cs <= 8 Then
'            WB$ = "W" look for attack by White on Square(r,c) row & column to check
'         or WB$ = "B" look for attack by Black on Square(r,c) row & column to check
         'If (WB$ = "W" And bRCBoard(rs, cs, Index) <= WPn) Or _
         '   (WB$ = "B" And bRCBoard(rs, cs, Index) >= BRn) Then
         If (WB$ = "W" And bRCBoard(rs, cs, Index) = WPn) Or _
            (WB$ = "B" And bRCBoard(rs, cs, Index) = BPn) Then
            If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BP or WP
               rat = rs: cat = cs
               Exit Function
            End If
         End If
      End If
      ' En Passant take ?
      If WB$ = "W" Then
         If R = 5 Then ' Check if WP left or right & BEPP(BPSET,Index)=C
            cs = C - 1  ' left
            If cs > 1 Then ' Is there a WP to the left
               If bRCBoard(R, cs, Index) = WPn And BEPP(BEPSet, Index) = C Then
                  rat = R: cat = cs
                  AttackingPiece = WPn
                  aGLOBEP = True
                  Exit Function
               End If
            End If
            cs = C + 1 ' right
            If cs <= 8 Then   ' Is there a WP to the right
               If bRCBoard(R, cs, Index) = WPn And BEPP(BEPSet, Index) = C Then
                  rat = R: cat = cs
                  AttackingPiece = WPn
                  aGLOBEP = True
                  Exit Function
               End If
            End If
         End If
      Else  ' WB$="B"
         If R = 4 Then ' Check if BP left or right & WEPP(WPSET,Index)=C
            cs = C - 1  ' left
            If cs > 1 Then ' Is there a BP to the left
               If bRCBoard(R, cs, Index) = BPn And WEPP(WEPSet, Index) = C Then
                  rat = R: cat = cs
                  AttackingPiece = BPn
                  aGLOBEP = True
                  Exit Function
               End If
            End If
            cs = C + 1 ' right
            If cs <= 8 Then  ' Is there a BP to the right
               If bRCBoard(R, cs, Index) = BPn And WEPP(WEPSet, Index) = C Then
                  rat = R: cat = cs    ' Might have been zeroed
                  AttackingPiece = BPn
                  aGLOBEP = True
                  Exit Function
               End If
            End If
         End If
      End If
      
   End If
   GoTo TestKing  ' ie exclude ForwardPawn move
'---------------------------------------------------
ForwardPawn:
'   If WB$ = "B" Then ' Look for attack by Black on Square(r,c)
'   S = 1       ' Attack Down, so pawn has to be @ r+1 down to attacked square
'   If WB$ = "W" Then   ' Look for attack by White on Square(r,c)
'   S = -1    ' Attack Up, so pawn has to be @ r-1 down to attacked square
   If bRCBoard(R, C, Index) = 0 Then ' Square available for pawn to move to
      ' Move forward 1 square
      rs = R + s   '''
      If rs >= 1 And rs <= 8 Then
         cs = C
         If (WB$ = "W" And bRCBoard(rs, cs, Index) = WPn) Or _
            (WB$ = "B" And bRCBoard(rs, cs, Index) = BPn) Then
            rat = rs: cat = cs
            Exit Function
         Else  ' No pawn 1 back to move into square test if pawn a 2 squares back
         End If
      Else
         GoTo TestKing
      End If
      
   '         If WB$ = "W" look for move by White pawn to Square(r,c) row & column to check
   '         or WB$ = "B" look for move by Black pawn to Square(r,c) row & column to check
      rs = R + 2 * s ' Possible 2 move from 7th(B) or 2nd(W)
      cs = C
      
      If bRCBoard(R + s, C, Index) = 0 Then ' 1 square back clear
         If (WB$ = "W" And bRCBoard(2, cs, Index) = WPn) Or _
            (WB$ = "B" And bRCBoard(7, cs, Index) = BPn) Then
               rat = rs: cat = cs
               Exit Function
         Else  ' No pawn 2 squares back to move into destination square
         End If
      End If
   End If
   
'---------------------------------------------------
TestKing:
   If ExcludeKing = False Then
      ' Check BK [11] WK [5]
      If WB$ = "B" Then ' Look for attack by Black King on Square(R,C)
         AttackingPiece = BKn
      Else  ' Look for attack by White King on Square(R,C)
         AttackingPiece = WKn
      End If
      For rs = R - 1 To R + 1
         If 1 <= rs And rs <= 8 Then
            For cs = C - 1 To C + 1
               If 1 <= cs And cs <= 8 Then
                  
                  'If Not (RS = R And CS = C) Then
                  If rs <> R Or cs <> C Then
                     If bRCBoard(rs, cs, Index) = AttackingPiece Then ' BK or WK
                        rat = rs: cat = cs
                        Exit Function
                     End If
                  End If
               
               End If
            Next cs
         End If
      Next rs
   End If

   AttackingPiece = 0
   RC_Targetted = False
End Function

Public Function TestForCheckMate(WB$, Index As Integer) As Boolean
' Come in with Board Index & King in check!

' Called from WHITE_PIECE_SECOND_MOVE (Mod2MvMate) &
'   CheckIfWMatesB, CheckIfBMatesW & SetPiece (Setup on Form1)

' WB$ = "W" to test if WK in checkmate
' WB$ = "B" to test if BK in checkmate
' Index is bRCBoard(r,c Index)
' Public rbk,cbk & rwk,cwk returned from SavePosition(frm, Index)
' AttackingPiece returned from RC_Targetted

' Public PlayColor$, OppColor$             ' Player color, Opponent color


Dim rsk As Long, csk As Long  ' takes rwk,cwk or rbk,cbk
Dim rs As Long, cs As Long    ' test row,col
Dim N As Long
Dim nKing As Long
Dim NA As Long    ' Takes attacking piece number
Dim D As Long, NP As Long   ' For multiple attacks
   
   TestForCheckMate = False
   
   If WB$ = "W" Then    ' White to move.
      FindKingRC "W", rwk, cwk, Index ' Test if WK in checkmate by black
      rsk = rwk
      csk = cwk
      OppColor$ = "B"   ' Look for attack by B
      PlayColor$ = "W"
   Else                 ' Black to move.
      FindKingRC "B", rbk, cbk, Index ' Test if BK in checkmate by white
      rsk = rbk         ' Test if BK in checkmate by white
      csk = cbk
      OppColor$ = "W"   ' Look for attack by W
      PlayColor$ = "B"
   End If
   
   
   CopyMemory TestCheckMateBoard(1, 1), bRCBoard(1, 1, Index), 64
   
   ' Check squares around King to see if attacked
   
   ' Temp clear K  ?
   nKing = bRCBoard(rsk, csk, Index)
   bRCBoard(rsk, csk, Index) = 0
   '111
   '101
   '111
   
      For rs = rsk - 1 To rsk + 1
         If 1 <= rs And rs <= 8 Then
            For cs = csk - 1 To csk + 1
               If 1 <= cs And cs <= 8 Then
                  If rs <> rsk Or cs <> csk Then
                     If Not Attacked(rs, cs, Index) Then
                        N = AttackingPiece
                        ' Restore King
                        bRCBoard(rsk, csk, Index) = nKing
                        Exit Function  ' Not checkmate
                     End If
                  
                  End If
               End If
            Next cs
         End If
      Next rs
   
   ' So no King move possible, either moves or takes into check or blocked in.
   ' Put King back
   bRCBoard(rsk, csk, Index) = nKing
   
   ReDim AttackType(1 To 6)     ' Unlikely to be more than 3
   ReDim AttackEP(1 To 6)
   ReDim AttackTypeRow(1 To 6)  ' Unlikely to be more than 3
   ReDim AttackTypeCol(1 To 6)  ' Unlikely to be more than 3
   
   ' Now need to check whether piece(s) attacking King can be taken or blocked.
   ' If WK find B piece attacking WK
   ' If BK find W piece attacking BK
   NumAttacking = 0
   
   Do
      If RC_Targetted(rsk, csk, OppColor$, Index) Then
         ' at least one piece AttackingPiece( @ rat,cat) attacking King
         NumAttacking = NumAttacking + 1
         If NumAttacking > 6 Then   ' Unlikely
            ReDim Preserve AttackType(1 To NumAttacking)
            ReDim Preserve AttackTypeRow(1 To NumAttacking)
            ReDim Preserve AttackTypeCol(1 To NumAttacking)
         End If
         AttackType(NumAttacking) = AttackingPiece
         AttackEP(NumAttacking) = aGLOBEP
         AttackTypeRow(NumAttacking) = rat
         AttackTypeCol(NumAttacking) = cat
         ' Zero AttackingPiece @ rat,cat
         bRCBoard(rat, cat, Index) = 0
         ' and check for any other attacking pieces (could be in line Q & B , Q & R etc)
      Else
         Exit Do
      End If
   Loop
   ' Restore bRCBoard(,,Index) ie restore zeroed attacking pieces
   CopyMemory bRCBoard(1, 1, Index), TestCheckMateBoard(1, 1), 64
   
   If NumAttacking = 0 Then Exit Function
   If NumAttacking = 1 Then
      TestForCheckMate = MATE(NumAttacking, rsk, csk, Index) ', AttackEP(NumAttacking))
      ' Board Index will contains solved board
      Exit Function
   Else
      ' NumAttacking King > 1
      ' Find attacking piece closest to King
      ' Need better way eg King checked by a Knight could
      ' be closest
      
      NA = 100
      For N = 1 To NumAttacking
         'AT = AttackType(n)   ' TEST
         D = Sqr((rsk - AttackTypeRow(N)) * (rsk - AttackTypeRow(N)) + _
                 (csk - AttackTypeCol(N)) * (csk - AttackTypeCol(N)))
         If D < NA Then
            NA = D
            NP = N
         End If
      Next N
      NA = NP
      N = AttackType(NP)   ' TEST
      TestForCheckMate = MATE(NA, rsk, csk, Index) ', AttackEP(NA))
   End If
End Function

Private Function MATE(NA As Long, rsk As Long, csk As Long, Index As Integer) As Boolean  ', aEP As Boolean) As Boolean
' NA as attacking piece number ie NumAttacking
' Index is bRCBoard(r,c Index)
' Public rbk,cbk & rwk,cwk  black/white king positions
' AttackingPiece returned from RC_Targetted

' Private PlayColor$, OppColor$  ' Piece color, Opposite color

   ' Checkmate logic
   ' Come here when King cannot move out of check
   ' Take attacking piece @ AttackTypeRow(NA), AttackTypeCol(NA) and see if can be taken
   '  if so take and check that not discovered check
   '    if not discovered check then is King still in check
   '      if not then not Checkmate if still in check then CHECKMATE
   ' If can't be taken or is discovered check then
   '   if R,B or Q can it be blocked?
   '     if so block and check that not discovered check
   '       if not discovered check then is King still in check
   '         if not then not Checkmate if still in check then CHECKMATE
   '         BUT if discovered check see if another peice can block attacking piece
   ' If can't be taken or blocked and not discovered check
   ' then CHECKMATE

Dim N As Long
Dim NMates As Long
Dim rs As Long, cs As Long
' Saved piece types & locations
Dim svrat As Long, svcat As Long, svAttackingPiece
Dim svrated As Long, svcated As Long, svAttackedPiece As Long
Dim NumAttacks As Long
      
   MATE = False
   'if WK find attack by white on attacking B piece
   'if BK find attack by black on attacking W piece

   ' SquareAttacks(R, C, WB$, Index, _
   ' Optional ExcludeKing As Boolean = False, Optional ExcludeDiagPawnMove As Boolean = False) As Boolean
   
   '  Returns all pieces attacking square r,c on Board Index
   '  NumAttacking
   '   ReDim AttackType(1 To 6)     ' Unlikely to be more than 3
   '   ReDim AttackTypeRow(1 To 6)  ' Unlikely to be more than 3
   '   ReDim AttackTypeCol(1 To 6)  ' Unlikely to be more than 3
   
   ' Returns:-
   ' Public AttackingPiece @
   ' Public rat As Long, cat As Long     ' r,c of AttackingPiece
   
   ' Enter with bRCBoard(1-8 columns, 8 - 1 rows) set
   '            WB$ = "W" look for attack by White on Square(r,c) row & column to check
   '         or WB$ = "B" look for attack by Black on Square(r,c) row & column to check
   
   ' ExcludeKing         = False <default> Include attack by opposing king
   ' ExcludeKing         = True  Exclude attack by opposing king
   ' ExcludeDiagPawnMove = False <default> Include diagonal Exclude forward pawn move, default
   ' ExcludeDiagPawnMove = True  Exclude diagonal Include forward pawn move
      
      
      If Not RC_Targetted(AttackTypeRow(NA), AttackTypeCol(NA), PlayColor$, Index) Then
         ' Attacking piece cannot be taken CAN IT BE BLOCKED
         ' Knight can't be blocked
         If AttackType(NA) = BNn Or AttackType(NA) = WNn Then
            MATE = True
            Exit Function
         ' Pawn can't be blocked
         ElseIf AttackType(NA) = WPn Or AttackType(NA) = BPn Then
            
            ' Test if pawn can be taken En Passant
            rs = AttackTypeRow(NA)
            If AttackType(NA) = WPn Then
               If rs = 4 Then  ' Check if BP can be taken by WP en passant
                  cs = AttackTypeCol(NA) - 1 ' look left for WP
                  If cs > 1 Then
                     If bRCBoard(rs, cs, Index) = WPn And WEPP(WEPSet, Index) = cs Then
                        Exit Function  ' No Mate
                     End If
                  End If
                  cs = AttackTypeCol(NA) + 1 ' look right for WP
                  If cs < 8 Then
                     If bRCBoard(rs, cs, Index) = WPn And WEPP(WEPSet, Index) = cs Then
                        Exit Function  ' No Mate
                     End If
                  End If
                  MATE = True
                  Exit Function
               End If
            ElseIf AttackType(NA) = BPn Then
               If rs = 5 Then  ' Check if WP can be taken by BP en passant
                  cs = AttackTypeCol(NA) - 1 ' look left for BP
                  If cs > 1 Then
                     If bRCBoard(rs, cs, Index) = BPn And BEPP(BEPSet, Index) = cs Then
                        Exit Function  ' No Mate
                     End If
                  End If
                  cs = AttackTypeCol(NA) + 1 ' look right for BP
                  If cs < 8 Then
                     If bRCBoard(rs, cs, Index) = BPn And BEPP(BEPSet, Index) = cs Then
                        Exit Function  ' No Mate
                     End If
                  End If
                  MATE = True
                  Exit Function
               End If
            End If
           MATE = True
           Exit Function
         
         Else  ' R,Q or B  Can it be blocked?
            ' R & Q
            If (AttackType(NA) = BRn Or AttackType(NA) = WRn) Or _
               (AttackType(NA) = BQn Or AttackType(NA) = WQn) Then
               
               ' ROW BLOCKING?
               If AttackTypeRow(NA) = rsk Then ' Attack piece & King on SAME ROW
                  If Abs(AttackTypeCol(NA) - csk) = 1 Then ' Can't be blocked
                     MATE = True
                     Exit Function
                  Else  ' Look at all blank squares between Piece & King
                     ' Unrolled For
                     If csk > AttackTypeCol(NA) Then  ' King to right of attacking piece
                        For N = AttackTypeCol(NA) + 1 To csk - 1  ' look from left of K
                           ' if WK find intercept by white, if BK find intercept by black
                           If CheckForIntercept(rsk, N, Index, rsk, csk) Then
                              Exit Function  ' Can be blocked
                           End If
                        Next N
                        MATE = True
                        Exit Function  ' Cannot be blocked
                     Else ' csk < AttackTypeCol(NA)
                        For N = csk + 1 To AttackTypeCol(NA) - 1    ' look from K to right
                           If CheckForIntercept(rsk, N, Index, rsk, csk) Then
                              Exit Function  ' Cannot be blocked
                           End If
                        Next N
                        MATE = True
                        Exit Function  ' Cannot be blocked
                     End If
                  End If
               
               ' COLUMN BLOCKING?
               ElseIf AttackTypeCol(NA) = csk Then ' Attack piece & King on SAME COLUMN
                  If Abs(AttackTypeRow(NA) - rsk) = 1 Then ' Can't be blocked
                     MATE = True
                     Exit Function
                  Else
                     If rsk > AttackTypeRow(NA) Then     ' ??
                        For N = AttackTypeRow(NA) + 1 To rsk - 1  ' look below to K
                           If CheckForIntercept(N, csk, Index, rsk, csk) Then
                              Exit Function  ' Can be blocked
                           End If
                        Next N
                        MATE = True
                        Exit Function  ' Cannot be blocked
                     Else  ' rsk < AttackTypeRow(NA)
                        For N = rsk + 1 To AttackTypeRow(NA) - 1  ' look from K to above
                           If CheckForIntercept(N, csk, Index, rsk, csk) Then
                              Exit Function  ' Can be blocked
                           End If
                        Next N
                        MATE = True
                        Exit Function  ' Can't be blocked
                     End If
                  End If
               End If
            End If ' NOT ROW OR COLUMN BLOCKING
            
            ' MUST BE BISHOP OR QUEEN DIAGONAL ATTACK

            If (Abs(AttackTypeRow(NA) - rsk) = 1) And _
               (Abs(AttackTypeCol(NA) - csk) = 1) Then
               ' On diagonal next to king - can't be blocked NB attacking piece can't be taken here
                MATE = True
                Exit Function
            Else  ' Test if can be blocked   X 4 possible diagonals
               ' King @ rsk,csk    Attacking piece at AttackTypeRow(NA),AttackTypeCol(NA)
               If AttackTypeRow(NA) < rsk Then 'BL or BR
                  If AttackTypeCol(NA) < csk Then  ' BL, AP SW of King
                     rs = rsk - 1: cs = csk - 1
                     Do
                        If rs <= AttackTypeRow(NA) Then Exit Do
                        If cs <= AttackTypeCol(NA) Then Exit Do
                        If CheckForIntercept(rs, cs, Index, rsk, csk) Then
                           Exit Function  ' Can be blocked
                        End If
                        rs = rs - 1
                        cs = cs - 1
                     Loop
                  Else  ' AttackTypeCol(NA) > csk ' BR, AP SE of King
                     rs = rsk - 1: cs = csk + 1
                     Do
                        If rs <= AttackTypeRow(NA) Then Exit Do
                        If cs >= AttackTypeCol(NA) Then Exit Do
                        If CheckForIntercept(rs, cs, Index, rsk, csk) Then
                           Exit Function  ' Can be blocked
                        End If
                        rs = rs - 1
                        cs = cs + 1
                     Loop
                  End If
               
               Else  ' AttackTypeRow(NA) > rsk  TL or TR
                  If AttackTypeCol(NA) < csk Then  ' TL, AP NW of King
                     rs = rsk + 1: cs = csk - 1
                     Do
                        If rs >= AttackTypeRow(NA) Then Exit Do
                        If cs <= AttackTypeCol(NA) Then Exit Do
                        If CheckForIntercept(rs, cs, Index, rsk, csk) Then
                           Exit Function  ' Can be blocked
                        End If
                        rs = rs + 1
                        cs = cs - 1
                     Loop
                  Else  ' AttackTypeCol(NA) > csk  TR, AP NE of King
                     rs = rsk + 1: cs = csk + 1
                     Do
                        If rs >= AttackTypeRow(NA) Then Exit Do
                        If cs >= AttackTypeCol(NA) Then Exit Do
                        If CheckForIntercept(rs, cs, Index, rsk, csk) Then
                           Exit Function  ' Can be blocked
                        End If
                        rs = rs + 1
                        cs = cs + 1
                     Loop
                  End If
                  
               End If  ' End of diagonal tests
               ' No diagonal blocking therefore:-
               MATE = True
               Exit Function
            End If   ' Check if next to KIng
         End If   ' R,Q or B  Can it be blocked?

      Else
         ' AttackingPiece can be taken
         ' SquareAttacked by PlayColor$ piece AttackingPiece( @ rat,cat) & can be a king
         ' Need to find how many pieces can take the checking piece
         ' Save the checking piece's location
         ' 1st found Checking piece location and type
         svrated = AttackTypeRow(NA)
         svcated = AttackTypeCol(NA)
         svAttackingPiece = bRCBoard(svrated, svcated, Index)
                  
         ' Find if more pieces attacking in-line
         NumAttacking = 0
         Do
            If RC_Targetted(svrated, svcated, PlayColor$, Index) Then
               ' at least one piece AttackingPiece( @ rat,cat) attacking checking piece
               NumAttacking = NumAttacking + 1
               If NumAttacking > 6 Then   ' Unlikely
                  ReDim Preserve AttackType(1 To NumAttacking)
                  ReDim Preserve AttackTypeRow(1 To NumAttacking)
                  ReDim Preserve AttackTypeCol(1 To NumAttacking)
               End If
               AttackType(NumAttacking) = AttackingPiece
               AttackTypeRow(NumAttacking) = rat
               AttackTypeCol(NumAttacking) = cat
               ' Zero AttackingPiece @ rat,cat
               bRCBoard(rat, cat, Index) = 0
               ' and check for any other attacking pieces
               ' (could be in line R & R, Q & B , Q & R rarely B & B)
            Else
               Exit Do
            End If
         Loop
         ' Restore bRCBoard(,,Index) ie restore zeroed attacking pieces
         CopyMemory bRCBoard(1, 1, Index), TestCheckMateBoard(1, 1), 64
         
         If NumAttacking > 1 Then
            FindKingRC OppColor$, rs, cs, Index
            For N = 1 To NumAttacking
               Select Case AttackType(N)
               Case BRn, WRn, BBn, WBn, BQn, WQn ', WKn, BKn
                  bRCBoard(svrated, svcated, Index) = 0 ' Temp zero checking piece
                  ' Does that piece PN @ RP,CP attack the blank square
                  ' @ svrated, svcated on board Index and not self-check?
                  If Not IsPieceAttackingRC(AttackType(N), AttackTypeRow(N), AttackTypeCol(N), _
                     svrated, svcated, Index) Then ', RS, CS) Then
                     ' Mark in list as not in line with AttakingPiece
                     AttackType(N) = 0
                  End If
               Case Else  ' Knights, Pawns, King OK
               End Select
            Next N
         End If
         ' Restore bRCBoard(,,Index) ie restore zeroed attacking pieces
         CopyMemory bRCBoard(1, 1, Index), TestCheckMateBoard(1, 1), 64
         
         NMates = 0
         
         ' Let each piece take the checking piece & see if King still in check
         NumAttacks = NumAttacking
         For N = 1 To NumAttacking
            ' Take each piece attacking the checking piece
            ' NB if in line (RR,QR,QB,BB,KR,KB,KQ) then cannot take with the furthest away
            ' So need to check squares between
            '   (AttackTypeRow(N), AttackTypeCol(N) & (svrated, svcated)
            '   before taking attacking the checking piece
            If AttackType(N) <> 0 Then
               ' Zero the piece attacking the checking piece
               bRCBoard(AttackTypeRow(N), AttackTypeCol(N), Index) = 0
               If aGLOBEP Then
                  bRCBoard(svrated, svcated, Index) = 0
                  If PlayColor$ = "W" And AttackType(N) = WPn Then
                     bRCBoard(svrated + 1, svcated, Index) = AttackType(N)
                  ElseIf PlayColor$ = "B" And AttackType(N) = BPn Then
                     bRCBoard(svrated - 1, svcated, Index) = AttackType(N)
                  Else
                     ' Take the checking piece
                     bRCBoard(svrated, svcated, Index) = AttackType(N)
                  End If
               Else
                  ' Take the checking piece
                  bRCBoard(svrated, svcated, Index) = AttackType(N)  ' Unless En passant
               End If
               ' Is King still in check
               If FindKingRC(PlayColor$, rs, cs, Index) Then
                  If RC_Targetted(rs, cs, OppColor$, Index) Then
                     NMates = NMates + 1
                  Else
                     NMates = 0
                  End If
               Else
                  MsgBox "ERROR" & vbCrLf & "NO " & PlayColor$ & " KING in Function MATE (ModChecker.bas) ???"
               End If
               ' Restore bRCBoard(r, c, Index)
               CopyMemory bRCBoard(1, 1, Index), TestCheckMateBoard(1, 1), 64
            Else
               NumAttacks = NumAttacks - 1
            End If
         Next N
         
         If NMates > 0 And (NMates = NumAttacks) Then
            MATE = True
         Else
            MATE = False
         End If
      End If
End Function

Private Function Attacked(rs As Long, cs As Long, Index As Integer) As Boolean
' Private PlayColor$, OppColor$  ' Piece color, Opposite color
   Attacked = False
   If bRCBoard(rs, cs, Index) = 0 Then
      If RC_Targetted(rs, cs, OppColor$, Index) Then
         Attacked = True
      'Else 'Not attacked
      End If
      Exit Function
      'End If
   Else  ' Square occupied, is it opposite color?
      ' If so, when king takes, will it be in check?
      If (OppColor$ = "W" And bRCBoard(rs, cs, Index) <= WPn) Or _
         (OppColor$ = "B" And bRCBoard(rs, cs, Index) >= BRn) Then
         If RC_Targetted(rs, cs, OppColor$, Index) Then
            Attacked = True
         'Else 'Not attacked
         End If
         Exit Function
      End If
   End If
   Attacked = True
End Function

Private Function CheckForIntercept(rs As Long, cs As Long, Index As Integer, rsk As Long, csk As Long) As Boolean
' Check to see if a piece can attack sq RS,CS
' rsk,csk King position NOT USED SO FAR
Dim R As Long, C As Long
Dim PN As Long
   CheckForIntercept = False
   For R = 1 To 8
   For C = 1 To 8
      PN = bRCBoard(R, C, Index)
      If PN <> 0 Then
         ' If PlayColor$ = "W" looking for checkmate ON White by Black
         '    hence here an intercept by White
         ' If PlayColor$ = "B" looking for checkmate ON Black by White
         '    hence here an intercept by Black
         If PlayColor$ = "B" And PN >= BRn Or _
            PlayColor$ = "W" And PN <= WPn Then
            ' Does that piece @ R,C attack the square @ rsk, n
'            If PN = 6 And R = 4 And C = 7 Then Stop
            
            
            If IsPieceAttackingRC(PN, R, C, rs, cs, Index) Then ', rsk, csk) Then
               CheckForIntercept = True
               Exit Function
            End If
         End If
      End If
   Next C
   Next R
End Function

Private Function IsPieceAttackingRC(PN As Long, RP As Long, CP As Long, RSQ As Long, CSQ As Long, Index As Integer) As Boolean ', _
'         rsk As Long, csk As Long) As Boolean
' Does that piece PN @ RP,CP attack the blank square @ RSQ,CSQ on board Index
' rsk,csk King position
Dim KColor$, OKColor$
Dim rk As Long, ck As Long
   IsPieceAttackingRC = False
   Select Case PN
   Case WNn, BNn
      If Abs(CSQ - CP) = 2 And Abs(RSQ - RP) = 1 Then GoTo TestDiscoveredCheck
      If Abs(CSQ - CP) = 1 And Abs(RSQ - RP) = 2 Then GoTo TestDiscoveredCheck
      Exit Function
   Case WPn
      If CP <> CSQ Then Exit Function  ' WP cannot intercept
      If RSQ <= 2 Then Exit Function ' WP cannot intercept
      If RP + 1 = RSQ Then GoTo TestDiscoveredCheck
      If RP = 2 And RP + 2 = RSQ Then   ' 2 move pawn
         If bRCBoard(RP + 1, CSQ, Index) = 0 Then
            GoTo TestDiscoveredCheck
         End If
         Exit Function
      End If
      Exit Function
   Case BPn
      If CP <> CSQ Then Exit Function  ' BP cannot intercept
      If RSQ >= 7 Then Exit Function ' BP cannot intercept
      If RP - 1 = RSQ Then GoTo TestDiscoveredCheck
      If RP = 7 And RP - 2 = RSQ Then  ' 2 move pawn
         If bRCBoard(RP - 1, CSQ, Index) = 0 Then
            GoTo TestDiscoveredCheck
         End If
         Exit Function
       End If
       Exit Function
   Case WRn, BRn
      If Not CheckHorzVerts(PN, RP, CP, RSQ, CSQ, Index) Then Exit Function
    Case WBn, BBn
      If Not CheckDiags(PN, RP, CP, RSQ, CSQ, Index) Then Exit Function
    Case WQn, BQn
      If Not CheckHorzVerts(PN, RP, CP, RSQ, CSQ, Index) And _
         Not CheckDiags(PN, RP, CP, RSQ, CSQ, Index) Then
         Exit Function
      End If
    Case Else
      Exit Function  ' ie exclude Kings
    End Select
TestDiscoveredCheck:
' Public PlayColor$, OppColor$             ' Player color, Opponent color
   If PN <= WPn Then
      KColor$ = "W"
      OKColor$ = "B"
   Else
      KColor$ = "B"
      OKColor$ = "W"
   End If
   ' Temp move PN
   FindKingRC KColor$, rk, ck, Index
   
   bRCBoard(RSQ, CSQ, Index) = PN
   bRCBoard(RP, CP, Index) = 0
   ' Check if OppColor$ King has discovered check
   If Not RC_Targetted(rk, ck, OKColor$, Index, True) Then ' King not in check
      bRCBoard(RSQ, CSQ, Index) = 0
      bRCBoard(RP, CP, Index) = PN
      IsPieceAttackingRC = True
      Exit Function ' Can be blocked
   End If
   bRCBoard(RSQ, CSQ, Index) = 0
   bRCBoard(RP, CP, Index) = PN
End Function

Private Function CheckHorzVerts(PN As Long, RP As Long, CP As Long, RSQ As Long, CSQ As Long, Index As Integer) As Boolean
' PN NOT USED SO FAR
Dim Row As Long, Col As Long
   CheckHorzVerts = False
   If RP = RSQ Then ' Same row move along columns
      If CP < CSQ Then
         For Col = CP + 1 To CSQ  'P+1 -> SQ
            If bRCBoard(RP, Col, Index) <> 0 Then Exit Function
         Next Col
      Else  ' CP > CSQ
         For Col = CSQ To CP - 1  ' SQ -> P-1
            If bRCBoard(RP, Col, Index) <> 0 Then Exit Function
         Next Col
      End If
   ElseIf CP = CSQ Then   ' Same column move along rows
      If RP < RSQ Then
         For Row = RP + 1 To RSQ  ' P+1 -> SQ
            If bRCBoard(Row, CP, Index) <> 0 Then Exit Function
         Next Row
      Else  ' RP > RSQ
         For Row = RSQ To RP - 1 ' SQ -> P-1
            If bRCBoard(Row, CP, Index) <> 0 Then Exit Function
         Next Row
      End If
   Else  ' Not same row or column
      Exit Function
   End If
   CheckHorzVerts = True
End Function

Private Function CheckDiags(PN As Long, RP As Long, CP As Long, RSQ As Long, CSQ As Long, Index As Integer) As Boolean
' PN NOT USED SO FAR
Dim Row As Long, Col As Long
   CheckDiags = False
   If Abs(RSQ - RP) <> Abs(CSQ - CP) Then Exit Function
   If CSQ > CP Then  ' ->
      If RSQ > RP Then ' TR
         Row = RP + 1: Col = CP + 1
         Do
            If bRCBoard(Row, Col, Index) <> 0 Then Exit Function
            Row = Row + 1: Col = Col + 1
            If Col > CSQ Then Exit Do
         Loop
      ElseIf RSQ < RP Then ' BR
         Row = RP - 1: Col = CP + 1
         Do
            If bRCBoard(Row, Col, Index) <> 0 Then Exit Function
            Row = Row - 1: Col = Col + 1
            If Col > CSQ Then Exit Do
         Loop
      End If
   ElseIf CSQ < CP Then ' <-
      If RSQ > RP Then ' TL
         Row = RP + 1: Col = CP - 1
         Do
            If bRCBoard(Row, Col, Index) <> 0 Then Exit Function
            Row = Row + 1: Col = Col - 1
            If Col < CSQ Then Exit Do
         Loop
      ElseIf CSQ < CP And RSQ < RP Then ' BL
         Row = RP - 1: Col = CP - 1
         Do
            If bRCBoard(Row, Col, Index) <> 0 Then Exit Function
            Row = Row - 1: Col = Col - 1
            If Col < CSQ Then Exit Do
         Loop
      End If
   End If
   CheckDiags = True
End Function

Public Function FindKingRC(a$, rk As Long, ck As Long, Index As Integer) As Boolean
' a$ ="B" find Black king
' a$ ="W" find White king
' on bRCBoard(r,c,Index)
' Returns rk,ck
Dim KingNum As Long
Dim R As Long, C As Long
   FindKingRC = True
   If a$ = "B" Then
      KingNum = BKn
      For R = 8 To 1 Step -1
      For C = 1 To 8
         If bRCBoard(R, C, Index) = KingNum Then
            rk = R: ck = C
            Exit Function
         End If
      Next C
      Next R
   Else
      KingNum = WKn
      For R = 1 To 8
      For C = 1 To 8
         If bRCBoard(R, C, Index) = KingNum Then
            rk = R: ck = C
            Exit Function
         End If
      Next C
      Next R
   End If
   ' ERROR NO KING
   FindKingRC = False
End Function


