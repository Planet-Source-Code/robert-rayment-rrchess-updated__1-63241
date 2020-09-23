Attribute VB_Name = "ModMate"
' ModMate.bas  ~RRChess~

Option Explicit

Public NFirstMoves As Long

Private rk As Long, ck As Long      ' King position
Private rOff As Long, cOff As Long  ' Offsets from piece start square
Private PColor$                     ' Player's color
Private OColor$                     ' Opposite color
' FAM Depth MateNum 1,2,3,,
Private MOVEDONE As Boolean, ALLDONE As Boolean, FOUNDMATE As Boolean
Private StartIndex As Integer
Private AMAT As Long
Private NMates As Long
Private a$

Public Function MATER(WB$, Index As Integer, MateNum As Long) As Boolean
' WB$ "W" or "B"
' MateNum = 1 look for mate in 1
' MateNum = 3 look for mate in 2
' MateNum = 5 look for mate in 3
' MateNum = 7 look for mate in 4
' MateNum = 9 look for mate in 5
Dim k As Long

   MATER = False
   
   SavePosition Form1, 0   ' Save main dispay to Board 0
   SavePosition Form1, 1   ' Save main dispay to Board 1
   
   NFirstMoves = MoveCounter(WB$, Index + 1, 1)
   If NFirstMoves = 0 Then Exit Function
   
   If WB$ = "W" Then    ' White to move.
      PColor$ = "W"
      OColor$ = "B"   ' Look for attack by B
   Else                 ' Black to move.
      PColor$ = "B"
      OColor$ = "W"   ' Look for attack by W
   End If
   ' Save for resetting after Halt ?

   Message$ = ""
   aWKingInCheck = False
   aBKingInCheck = False
   ThePromPiece$ = ""

   For k = Index + 1 To MaxIndex
       If UDSearch = 0 Then
         cs(k) = 0
         rs(k) = 1
       Else ' UDSearch = 1 Opp diag search
         cs(k) = 9
         rs(k) = 8
       End If
       NumMoves(k) = 0
       NumMates(k) = 0
   Next k
   StartIndex = Index
   AMAT = 0
   NMates = 0
   
   ' Progress numbers
   Form1.LabPB(0) = 0
   Form1.LabPB(1) = NFirstMoves
   Form1.LabPB(2) = (MateNum + 1) \ 2
   Form1.Refresh
   
   MATER = MATERPLUS(PColor$, OColor$, Index, MateNum)
   
   
   aWKingInCheck = False
   aBKingInCheck = False
   If WB$ = "W" Then
      If FindKingRC("B", rbk, cbk, 1) Then
         If RC_Targetted(rbk, cbk, "W", 1) Then aBKingInCheck = True
      End If
   Else
      If FindKingRC("W", rwk, cwk, 1) Then
         If RC_Targetted(rwk, cwk, "B", 1) Then aWKingInCheck = True
      End If
   End If
   
End Function


Public Function MATERPLUS(PColor$, OColor$, Index As Integer, MateNum As Long) As Boolean
' MateNum=1,3,5,7,9,,  ie Mate in 1,2,3,4,5,, moves
' Progress
Dim npb As Long
   npb = 0
   Do
      If TestKeys() And Message$ = "HALTED" Then
         Message$ = "HALTED"
         Exit Function
      End If

      DoEvents ' To deter White out & Not Responding message
      
      Index = Index + 1
      Call MOVE_MATE(PColor$, OColor$, Index, FOUNDMATE, ALLDONE, MOVEDONE, MateNum)

      Select Case Index 'Depth
      
      Case 1 ' ODD   W1/B1  Depth<=MateNum
         If ALLDONE Then
            MATERPLUS = False
            Message$ = "NO MATE"
            Exit Do
         End If
            
         If MateNum = 1 Then  ' ie mate in 1 move
            If FOUNDMATE Then
               MATERPLUS = True
               Exit Do
            End If
            Index = Index - 1 ' Stay on same 'board'
            If MOVEDONE = True Then
               npb = npb + 1
               Form1.LabPB(0) = npb
            End If
            GoTo ContinueLooping
         Else
            If MOVEDONE = True Then
               npb = npb + 1
               Form1.LabPB(0) = npb
            End If
            CheckMOVEDONE Index
         End If
      Case 2 ' EVEN  B2/W2
         
            If ALLDONE Then
               If MateNum = 3 Then  ' ie mate in 2 moves
                  If NumMates(MateNum) > 0 And NumMoves(2) = NumMates(MateNum) Then
                     MATERPLUS = True
                     Exit Do   ' Key found
                  Else  ' EG maybe for black to make any move depends on a particular 1st white move
                     Index = 0    ' 0 -> board 1 W1/B1
                     Reset_COUNTERS Index + 2, MateNum
                     a$ = PColor$: PColor$ = OColor$: OColor$ = a$
                     GoTo ContinueLooping
                  End If
               ElseIf MateNum > 3 Then ' ie mate in > 2 moves
                  If AMAT = 1 And NumMoves(2) > 0 Then
                     MATERPLUS = True
                     Exit Do   ' Key found
                  Else  ' EG maybe for black to make any move depends on a particular 1st white move
                     Index = 0    ' 0 -> board 1 W1/B1
                     Reset_COUNTERS Index + 2, MateNum
                     a$ = PColor$: PColor$ = OColor$: OColor$ = a$
                     GoTo ContinueLooping
                  End If
               End If
            End If
            CheckMOVEDONE Index
'''''''''''''''''''''''''''''''''''''
      Case 3 To MateNum - 2
         If (Index And 1) = 1 Then  ' ODD
            If ALLDONE Then ' NOMATE Eg Try another first W move
               Index = Index - 3 ' Eg 0 -> board 1  W1/B1
               Reset_COUNTERS Index + 2, MateNum
               GoTo ContinueLooping
            End If
            If FOUNDMATE Then ' Early mate Eg Try another  previous black move
               Index = Index - 2 ' Eg 1 -> board 2  B2/W2
               Reset_COUNTERS Index + 2, MateNum
               a$ = PColor$: PColor$ = OColor$: OColor$ = a$
               GoTo ContinueLooping
            End If
            CheckMOVEDONE Index
         Else   'EVEN
            If ALLDONE Then
               If AMAT = 1 And NumMoves(Index) > 0 Then ' Eg Check if mate with all black 1st moves
                  Index = Index - 3 ' Eg 1 -> board 2   B2/W2
                  Reset_COUNTERS Index + 2, MateNum
               Else     ' NOMATE Eg Try another previous white move
                  Index = Index - 2 '  Eg 2 -> board 3  W3/B3
                  Reset_COUNTERS Index + 2, MateNum
                  a$ = PColor$: PColor$ = OColor$: OColor$ = a$
               End If
               GoTo ContinueLooping
            End If
            CheckMOVEDONE Index
         End If
'''''''''''''''''''''''''''''''''''''
      Case MateNum - 1 'Eg 6 ' EVEN B6/W6
            If ALLDONE Then
               If NumMates(MateNum) > 0 And NumMoves(MateNum - 1) = NumMates(MateNum) Then
                  AMAT = 1 ' Eg Check if mate with all black 2nd moves
                  Index = MateNum - 4 ' Eg 3 -> board 4   B4/W4
                  Reset_COUNTERS Index + 2, MateNum
               Else  ' Eg No mate for every 3rd black move
                  Index = MateNum - 3 ' Eg 4 -> board 5  W5/B5
                  Reset_COUNTERS Index + 2, MateNum
                  a$ = PColor$: PColor$ = OColor$: OColor$ = a$
               End If
               GoTo ContinueLooping
            End If
            CheckMOVEDONE Index
      Case MateNum 'Eg 7 ' ODD W7/B7
            If ALLDONE Then ' NOMATE Eg Try another  previous white move
              Index = MateNum - 3 ' Eg 4 -> board 5  W5/B5
              Reset_COUNTERS Index + 2, MateNum
              GoTo ContinueLooping
            End If
            If FOUNDMATE Then ' Test the rest of moves on previous 'board' MateNum-1
              Index = MateNum - 2 ' Eg 5 -> board 6  B6/W6
              cs(MateNum) = 0
              rs(MateNum) = 1    ' Eg Keep NumMates(7)
              If UDSearch = 1 Then
                 cs(MateNum) = 9
                 rs(MateNum) = 8     ' Eg Keep NumMates(7)
              End If
              NumMoves(MateNum) = 0
              a$ = PColor$: PColor$ = OColor$: OColor$ = a$
              GoTo ContinueLooping
            End If
            Index = Index - 1 ' Stay on same 'board', to find a mate or until ALLDONE
      End Select
ContinueLooping:
   Loop
End Function

Public Sub MOVE_MATE(PCul$, OCul$, ByVal Index As Integer, _
      FOUNDMATE As Boolean, ALLDONE As Boolean, MOVEDONE As Boolean, MateNum As Long)
Dim NP As Long
Dim csup As Long
 
 ' FAM
   FOUNDMATE = False
   ALLDONE = False
   MOVEDONE = False
   CopyMemory bRCBoard(1, 1, Index), bRCBoard(1, 1, Index - 1), 64 ' <-
   CopyALLBOOLS Index - 1, Index  '->
   
   If DOFLAG(Index) = 1 Then GoTo INDO ' Jump into Do Loop with PieceNum(Index)
   
   If UDSearch = 0 Then
      cs(Index) = cs(Index) + 1
      If cs(Index) > 8 Then
         cs(Index) = 1
         rs(Index) = rs(Index) + 1
      End If
      csup = 0
      If rs(Index) > 8 Then
         cs(Index) = csup
         rs(Index) = 1
      End If
   Else  ' UDSearch = 1 ' Opp diag scan, 2nd mate search
      cs(Index) = cs(Index) - 1
      If cs(Index) < 1 Then
         cs(Index) = 8
         rs(Index) = rs(Index) - 1
      End If
      csup = 9
      If rs(Index) < 1 Then
         cs(Index) = csup
         rs(Index) = 8
      End If
   End If
   
   If cs(Index) = csup Then
      ALLDONE = True
      Exit Sub 'EXIT SUB ALLDONE   ' >>>>>
   Else
       ' Get a Piece
      NP = bRCBoard(rs(Index), cs(Index), Index) ' <> 0 Then
      If NP <> 0 Then
         If (PCul$ = "W" And NP <= WPn) Or _
            (PCul$ = "B" And NP >= BRn) Then
            
            PieceNum(Index) = NP
            M(Index) = 1: Q(Index) = 1
            Do
               If TheMoveOK(Index) Then
                  NumMoves(Index) = NumMoves(Index) + 1
                  If PCul$ = PColor$ Then
                     If (PCul$ = "W" And aBKingInCheck) Or _
                         PCul$ = "B" And aWKingInCheck Then
                        If TestForCheckMate(OCul$, Index) Then
                           NumMates(Index) = NumMates(Index) + 1
                           FOUNDMATE = True
                           Exit Sub
                        Else  ' No Mate found here but King in check
                           DOFLAG(Index) = 1
                           MOVEDONE = True ' Move on unless Index = MateNum when MOVEDONE ignored
                           Exit Sub
                        End If   ' BK/WK in checkmate ?
                     End If   ' BK/WK in check ?
                  End If
                  ' Come here for opp color or King not in check
                  If Index < MateNum Then
                     DOFLAG(Index) = 1    'Allow Do Loop Re-entry for further piece movement
                     MOVEDONE = True 'EXIT SUB MOVEDONE    ' >>>>>>  No End Mate
                     Exit Sub
                  End If
               End If  ' MoveOK ?
INDO:
               If Not RestorePieces(Index) Then Exit Do  ' Restore & Check validity of new move for piece
            Loop
            ' No more directions/increments for piece so find another piece
            DOFLAG(Index) = 0
         End If  ' Piece  = 0
      End If  ' or wrong color
   End If   ' ALLDONE ?
End Sub

Public Function TestCheckMate(PCul$, OCul$, Index As Integer) As Boolean
   ' Find OCul$ King and see if checkmated by PCul$ on board Index
   TestCheckMate = False
   FindKingRC OCul$, rk, ck, Index                 ' Find OCul$ King,
   If RC_Targetted(rk, ck, PCul$, Index) Then      ' see if attacked by PCul$
      If TestForCheckMate(OCul$, Index) Then       ' if so, is OCul$ mated?
         TestCheckMate = True
      End If
   End If
End Function

Public Function RestorePieces(ByVal Index As Integer) As Boolean
' All Public or Private
   RestorePieces = True
   CopyMemory bRCBoard(1, 1, Index), bRCBoard(1, 1, Index - 1), 64
   CountPieces Index
   CheckEPPPCA Index
   CopyALLBOOLS Index - 1, Index
   If Not PieceDirec(M(Index), Q(Index), PieceNum(Index), aMovOK(Index)) Then
     RestorePieces = False 'Will Exit Do
   End If
End Function


Public Function TheMoveOK(Index As Integer) As Boolean
' Could put into Function MOVER   () As Boolean

' Public:-
' aWKingInCheck, aBKingInCheck
' aMovOK(Index), M(Index), Q(Index), _
' PieceNum(Index), DestPieceNum, PieceIndex$(Index), _
' rs(Index), cs(Index), rd(Index), cd(Index)
   
   '  Do
   svM(Index) = M(Index): svQ(Index) = Q(Index) ' Save in case need to reverse EP,PP,CA
   svDOFLAG(Index) = DOFLAG(Index)
   ' Test proposed move and make it if OK
   MOVER Index, aMovOK(Index), M(Index), Q(Index), _
      PieceNum(Index), DestPieceNum, PieceIndex$(Index), _
      rs(Index), cs(Index), rd(Index), cd(Index)
   If aMovOK(Index) Then
      If (PlayColor$ = "W" And Not aWKingInCheck) Or _
         (PlayColor$ = "B" And Not aBKingInCheck) Then
         TheMoveOK = True
      Else
         TheMoveOK = False
      End If
   Else
      TheMoveOK = False
   End If
End Function


Public Sub MOVER(Index As Integer, aMoveOK As Boolean, M As Long, Q As Long, _
                 PN As Long, DPN As Long, P$, _
                 SR As Long, SC As Long, DR As Long, DC As Long)
' MOVER Index,aMoveOK,M,Q,PN,DPN,P$,SR,SC,DR,DC
' Public aWKingInCheck, aBKingInCheck
' PN   Piece number
' DPN  Destination piece number
' P$   Piece description
' IN:  SR,SC source, M,Q,PN
' OUT: DPN,P$,DR,DC destination square
Dim PieceColor$
Dim LandingColor$
   aMoveOK = False
   aWKingInCheck = False
   aBKingInCheck = False
   GetPieceOffsets M, Q, PN, P$, rOff, cOff  ' Returns P$, rOff, cOff
   PieceColor$ = Left$(P$, 1)
   DR = SR + rOff: DC = SC + cOff  ' destination
   If (1 <= DR And DR <= 8) And (1 <= DC And DC <= 8) Then  ' On board ?
      DPN = bRCBoard(DR, DC, Index)   ' Destination piece number (0,1-12)
      If DPN <> 0 Then
         If DPN >= BRn Then LandingColor$ = "B" Else LandingColor$ = "W"
      End If
      If DPN = 0 Or PieceColor$ <> LandingColor$ Then ' ie not same color @ destination   ' W 1-6 B 7-12
         RowS = SR: ColS = SC
         RowE = DR: ColE = DC
         If Not MoveBarred(PN, Index) Then    ' Move not barred
            'If (PieceColor$ = "W" And Not aWKingInCheck) Or _
            '   (PieceColor$ = "B" And Not aBKingInCheck) Then
               aMoveOK = True
               bRCBoard(DR, DC, Index) = PN     ' move piece to DR,DC
               bRCBoard(SR, SC, Index) = 0      ' clear piece from start position
               
               TestPawnPromotion PN, DPN, SR, SC, DR, DC, Index ' P->Q  PN =Q, DPN =Dest
               TestEnPassant PN, DPN, SR, SC, DR, DC, Index     ' Opp P to 0
               TestCastling PN, SR, SC, DR, DC, Index

               ' EVEN IF KING IN CHECK HERE
               ' NEED TO ALLOW PN TO MOVE ON
               ' IE Q+1 or M+1 ON THE SAME BOARD
               ' HENCE A NEW MOVE FOR PN MAY
               ' TAKE THE KING OUT OF CHECK
               ' See if further pieces need action
            
            'End If
         End If   ' White/Black move barred ?
      End If   ' DPN = 0 or opp color
   End If   ' White/Black piece off the board ?
End Sub

Public Sub CheckEPPPCA(Index As Integer)
' Reverses En passant, Pawn promotion or Castling
   
   If PieceNum(Index) = WKn And aCastling(Index) = 1 Then
      M(Index) = svM(Index) + 1 ' Reverse CA and move on
      aCastling(Index) = 0
   ElseIf PieceNum(Index) = BKn And aCastling(Index) = 2 Then
      M(Index) = svM(Index) + 1 ' Reverse CA and move on
      aCastling(Index) = 0
   End If
   
   If PieceNum(Index) = WPn And aEnPassant(Index) Then
      M(Index) = svM(Index) + 1 ' Reverse EP and move on
   ElseIf PieceNum(Index) = BPn And aEnPassant(Index) Then
      M(Index) = svM(Index) + 1 ' Reverse EP and move on
   End If
   
   If PieceNum(Index) = WQn And aPawnProm(Index) Then
      M(Index) = svM(Index) + 1 ' Reverse PP and move on
      PieceNum(Index) = WPn
   ElseIf PieceNum(Index) = BQn And aPawnProm(Index) Then
      M(Index) = svM(Index) + 1 ' Reverse PP and move on
      PieceNum(Index) = BPn
   End If
End Sub

Public Sub TestPawnPromotion(PN As Long, DPN As Long, SR As Long, SC As Long, DR As Long, DC As Long, Index As Integer)
' Called only when move OK
' Only promotes to a Q!
' PN = PieceNum
' DPN = destination piece number ( 0 or black/white piece)
' SR,SC - DR,DC  source - destination of PN

   aPawnProm(Index) = False
   If PN = WPn Then
      If WEPP(WEPProm, Index) = 1 Then ' W Pawn Promotion
         PN = WQn
         NumWQ = NumWQ + 1
         NumWP = NumWP - 1
         If DPN <> 0 Then  ' Will need to decrease number of B Piece
            'BR,BN,BB or BQ only
            Select Case DPN
            Case 7: NumBR = NumBR - 1
            Case 8: NumBN = NumBN - 1
            Case 9: NumBB = NumBB - 1
            Case 10: NumBQ = NumBQ - 1
            End Select
         End If
         bRCBoard(DR, DC, Index) = PN  ' move white piece (Q) to DR,DC
         bRCBoard(SR, SC, Index) = 0   ' clear piece from start position
         WEPP(WEPProm, Index) = 0
         aPawnProm(Index) = True
      
         If TestForCheck("W", Index) Then aBKingInCheck = True
      
      End If
   ElseIf PN = BPn Then ' PColor$ = "B"
      If BEPP(BEPProm, Index) = 1 Then ' B Pawn Promotion
         PN = BQn
         NumBQ = NumBQ + 1
         NumBP = NumBP - 1
         If DPN <> 0 Then  ' Will need to decrease number of W Piece
            'WR,WN,WB or WQ only
            Select Case DPN
            Case 1: NumWR = NumWR - 1
            Case 2: NumWN = NumWN - 1
            Case 3: NumWB = NumWB - 1
            Case 4: NumWQ = NumWQ - 1
            End Select
         End If
         bRCBoard(DR, DC, Index) = PN  ' move black piece to DR,DC
         bRCBoard(SR, SC, Index) = 0   ' clear piece from start position
         BEPP(BEPProm, Index) = 0
         aPawnProm(Index) = True
            
         If TestForCheck("B", Index) Then aWKingInCheck = True
      
      End If
   End If
End Sub

Public Sub TestEnPassant(PN As Long, DPN As Long, SR As Long, SC As Long, DR As Long, DC As Long, Index As Integer)
' NOT ALL ARGUMENTS USED SO FAR IE  DPN,SC & DR
' Called only when move OK
' PN = PieceNum
' DPN = Destination piece
' SR,SC - DR,DC  source - destination of PN
   aEnPassant(Index) = False
   If PN = WPn Then
         If WEPP(WEPOK, Index) = 1 Then   ' MoveBarred set this as a valid take
            bRCBoard(SR, DC, Index) = 0   ' Clear BP
            NumBP = NumBP - 1
            aEnPassant(Index) = True
         End If
         If TestForCheck("W", Index) Then aBKingInCheck = True
   ElseIf PN = BPn Then
         If BEPP(BEPOK, Index) = 1 Then   ' MoveBarred set this as a valid take
            bRCBoard(SR, DC, Index) = 0   ' Clear WP
            NumWP = NumWP - 1
            aEnPassant(Index) = True
         End If
         If TestForCheck("B", Index) Then aWKingInCheck = True
   End If
End Sub

Public Sub TestCastling(PN As Long, SR As Long, SC As Long, DR As Long, DC As Long, Index As Integer)
' NOT ALL ARGUMENTS USED SO FAR IE  SR, SC,DR & DC
' Called only when move OK
' Public PColor$
   aCastling(Index) = 0
   'If PColor$ = "W" And PN = WKn Then
   If PN = WKn Then
      If Cast(WKSCastOK, Index) = 1 Then
            bRCBoard(1, 6, Index) = WRn
            bRCBoard(1, 8, Index) = 0
            aCastling(Index) = 1
            Cast(WKSCastOK, Index) = 0
      ElseIf Cast(WQSCastOK, Index) = 1 Then
            bRCBoard(1, 4, Index) = WRn
            bRCBoard(1, 1, Index) = 0
            aCastling(Index) = 1
            Cast(WQSCastOK, Index) = 0
      End If
      If TestForCheck("W", Index) Then aBKingInCheck = True
   ElseIf PN = BKn Then
      If Cast(BKSCastOK, Index) = 1 Then
            bRCBoard(8, 6, Index) = BRn
            bRCBoard(8, 8, Index) = 0
            aCastling(Index) = 2
            Cast(BKSCastOK, Index) = 0
      ElseIf Cast(BQSCastOK, Index) = 1 Then
            bRCBoard(8, 4, Index) = BRn
            bRCBoard(8, 1, Index) = 0
            aCastling(Index) = 2
            Cast(BQSCastOK, Index) = 0
      End If
      If TestForCheck("B", Index) Then aWKingInCheck = True
   End If
End Sub

Public Function TestForCheck(WB$, Index As Integer)
Dim rk As Long, ck As Long
   TestForCheck = False
   If WB$ = "W" Then
         If Not FindKingRC("B", rk, ck, Index) Then
         Message$ = "No Black King in TestForCheck ??"
         MsgBox Message$, vbCritical, "ERROR"
      End If
      If RC_Targetted(rk, ck, "W", Index) Then TestForCheck = True
   Else  ' WB$ = "B"
      If Not FindKingRC("W", rk, ck, Index) Then
         Message$ = "No White King in TestForCheck ??"
         MsgBox Message$, vbCritical, "ERROR"
      End If
      If RC_Targetted(rk, ck, "B", Index) Then TestForCheck = True
   End If
End Function

Public Sub GetPieceOffsets(ByVal M As Long, ByVal Q As Long, ByVal PieceNum As Long, _
                            Piece$, ROF As Long, COF As Long)
'IN:  M,Q,PieceNum the piece number
'OUT: Piece$,ROF,COF Piece description, row & column steps
' Increments step from start position
' M direction
' Q the piece step size from the start position
' (M & Q set by Function PieceDirec)

   Select Case PieceNum
   Case WRn, BRn
      Select Case M
      Case 1: COF = Q:  ROF = 0   ' ->
      Case 2: COF = 0:  ROF = Q   ' ^
      Case 3: COF = -Q: ROF = 0   ' <-
      Case 4: COF = 0:  ROF = -Q  ' v
      End Select
      If PieceNum = WRn Then Piece$ = "WR" Else Piece$ = "BR"
   Case WBn, BBn
      If UDSearch = 0 Then
         Select Case M
         Case 1: COF = Q:  ROF = Q   ' tr
         Case 2: COF = -Q: ROF = Q   ' tl
         Case 3: COF = -Q: ROF = -Q  ' bl
         Case 4: COF = Q:  ROF = -Q  ' br
         End Select
      Else  ' UDSearch = 1 ' Opp diag scan, searching 2nd mate
         Select Case M
         Case 1: COF = -Q: ROF = -Q  ' bl
         Case 2: COF = Q:  ROF = -Q  ' br
         Case 3: COF = Q:  ROF = Q   ' tr
         Case 4: COF = -Q: ROF = Q   ' tl
         End Select
      End If
      If PieceNum = WBn Then Piece$ = "WB" Else Piece$ = "BB"
   Case WQn, BQn
      If UDSearch = 0 Then
         Select Case M
         Case 1: COF = Q:  ROF = 0   ' ->
         Case 2: COF = 0:  ROF = Q   ' ^
         Case 3: COF = -Q: ROF = 0   ' <-
         Case 4: COF = 0:  ROF = -Q  ' v
         
         Case 5: COF = Q:  ROF = Q   ' tr
         Case 6: COF = -Q: ROF = Q   ' tl
         Case 7: COF = -Q: ROF = -Q  ' bl
         Case 8: COF = Q:  ROF = -Q  ' br
         End Select
      Else  ' UDSearch = 1
         Select Case M
         Case 1: COF = Q:  ROF = Q   ' tr
         Case 2: COF = -Q: ROF = Q   ' tl
         Case 3: COF = -Q: ROF = -Q  ' bl
         Case 4: COF = Q:  ROF = -Q  ' br
         
         Case 5: COF = Q:  ROF = 0   ' ->
         Case 6: COF = 0:  ROF = Q   ' ^
         Case 7: COF = -Q: ROF = 0   ' <-
         Case 8: COF = 0:  ROF = -Q  ' v
      End Select
      
      End If
      
      If PieceNum = WQn Then Piece$ = "WQ" Else Piece$ = "BQ"
   Case WNn, BNn
      If UDSearch = 0 Then
         Select Case M
         Case 1: COF = 2:  ROF = 1   ' tr1
         Case 2: COF = 1:  ROF = 2   ' tr2
         Case 3: COF = -1: ROF = 2   ' tl1
         Case 4: COF = -2: ROF = 1   ' tl2
         
         Case 5: COF = -2: ROF = -1  ' bl1
         Case 6: COF = -1: ROF = -2  ' bl2
         Case 7: COF = 1:  ROF = -2  ' br1
         Case 8: COF = 2:  ROF = -1  ' br2
         End Select
      Else  ' UDSearch = 1
         Select Case M
         Case 1: COF = -2: ROF = -1  ' bl1
         Case 2: COF = -1: ROF = -2  ' bl2
         Case 3: COF = 1:  ROF = -2  ' br1
         Case 4: COF = 2:  ROF = -1  ' br2
         Case 5: COF = 2:  ROF = 1   ' tr1
         Case 6: COF = 1:  ROF = 2   ' tr2
         Case 7: COF = -1: ROF = 2   ' tl1
         Case 8: COF = -2: ROF = 1   ' tl2
         End Select
      End If
      If PieceNum = WNn Then Piece$ = "WN" Else Piece$ = "BN"
   Case WPn
      Select Case M
      Case 1: COF = 1:  ROF = 1   ' tr
      Case 2: COF = -1: ROF = 1   ' tl
      Case 3: COF = 0:  ROF = 1   ' ^
      Case 4: COF = 0:  ROF = 2   ' ^^
      End Select
      Piece$ = "WP"
   Case BPn
      Select Case M
      Case 1: COF = 1:  ROF = -1  ' br
      Case 2: COF = -1: ROF = -1  ' bl
      Case 3: COF = 0:  ROF = -1  ' v
      Case 4: COF = 0:  ROF = -2  ' vv
      End Select
      Piece$ = "BP"
   Case WKn, BKn
      Select Case M
      Case 1: COF = 2:  ROF = 0   ' >>
      Case 2: COF = -2: ROF = 0   ' <<
      
      Case 3: COF = 1:  ROF = 0   ' ->
      Case 4: COF = 0:  ROF = 1   ' ^
      Case 5: COF = -1: ROF = 0   ' <-
      Case 6: COF = 0:  ROF = -1  ' v

      Case 7: COF = 1:  ROF = 1   ' tr
      Case 8: COF = -1: ROF = 1   ' tl
      Case 9: COF = -1: ROF = -1  ' bl
      Case 10: COF = 1: ROF = -1  ' br
      
      End Select
      If PieceNum = WKn Then Piece$ = "WK" Else Piece$ = "BK"
   End Select
End Sub

 Public Function PieceDirec(M As Long, Q As Long, _
         ByVal PieceNum As Long, MVOK As Boolean) As Boolean
         
' eg PieceDirec(MW,QW,WPieceNum,WMOK)

' Last move info:-
' M last direction
' Q piece step
' PieceBun the piece number
' MVOK says whether or not the last move was OK
'      if not will set a new direction

' Returns M,Q & True/False

   PieceDirec = False

   Select Case PieceNum
   Case WRn, WBn, BRn, BBn
      Q = Q + 1   'Increase movement
      If Not MVOK Then  ' New direction
         Q = 1
         M = M + 1
         If M > 4 Then Exit Function  ' New direction= False, will get new piece
      End If
   Case WQn, BQn
      Q = Q + 1
      If Not MVOK Then
         Q = 1
         M = M + 1
         If M > 8 Then Exit Function
      End If
   ' No Q incrementing for N,P,K
   ' just direction change
   Case WNn, BNn
      M = M + 1
      If M = 9 Then Exit Function
   Case WPn, BPn
      M = M + 1
      If M >= 5 Then
         Exit Function
      End If
   Case WKn, BKn
      M = M + 1
      If M > 10 Then Exit Function
   End Select
   PieceDirec = True
 End Function

Public Sub Reset_COUNTERS(ByVal SIndex As Integer, MateNum As Long)
Dim k As Long
   For k = SIndex To MateNum
      cs(k) = 0
      rs(k) = 1
      If UDSearch = 1 Then
         cs(k) = 9
         rs(k) = 8
      End If
      NumMoves(k) = 0: NumMates(k) = 0
      DOFLAG(k) = 0
   Next k
End Sub

Private Sub CheckMOVEDONE(ByRef Index As Integer)
   If MOVEDONE Then  ' Color Swap & Go on to next 'board'
      a$ = PColor$: PColor$ = OColor$: OColor$ = a$
   Else
      Index = Index - 1 ' Stay on same 'board' -1 because Do increments Index
   End If
End Sub

Public Function TEST_FOR_MATE(WB$, NMates As Long, MM As Long) As Boolean
'WB$ Color "W" or "B", NMates (1,3,5 or 7), MM returns 1,2,3 or 4
Dim k As Long
Dim PN As Long
Dim MNUM As Long  ', BIndex As Long

      TEST_FOR_MATE = True
      
      For k = 0 To MaxIndex
         NumMoves(k) = 0
         NumMates(k) = 0
         DOFLAG(k) = 0
      Next k
      MNUM = 0
      ' MATER(Color, StartIndex, MateNum) 1,2,3,4,5,, Checkmate on Board 0 ?
      ' Board Index = 0 Result returned for Board Index = 1
      For MNUM = 1 To NMates Step 2
         If MATER(WB$, 0, MNUM) Then Exit For
      Next MNUM
      If MNUM > NMates Then MNUM = 0
      
      
      If NFirstMoves = 0 Then
         Message$ = "STALEMATE"
         TEST_FOR_MATE = False
         Exit Function
      End If
      
      If Message$ = "HALTED" Then
         TEST_FOR_MATE = False
         Exit Function
      End If
      
      ' Source & Destination of last move
      svRowS = rs(1): svColS = cs(1)
      svRowE = rd(1): svColE = cd(1)
      Piece$ = PieceIndex$(1)
      
      If MNUM > 0 Then
          
         ' Board 0 as start
         ' Board 1 contains solution
         MM = (MNUM + 1) \ 2  ' Mate depth
         ThePromPiece$ = ""
         If a$ = "W" Then
            If Piece$ = "WP" And rd(1) = 8 Then
               ThePromPiece$ = "WQ"
            End If
         Else
            If Piece$ = "BP" And rd(1) = 1 Then
               ThePromPiece$ = "BQ"
            End If
         End If
         
         PN = bRCBoard(svRowS, svColS, 0)
         DestPieceNum = bRCBoard(svRowE, svColE, 0)
         ' Show solution
         MoveString$ = Piece$ & " "
         MoveString$ = MoveString$ & Chr$(96 + svColS) & Trim$(Str$(svRowS)) '& "-"
         If DestPieceNum = 0 Then
            MoveString$ = MoveString$ & "-"
         Else
            MoveString$ = MoveString$ & "x"
         End If
         MoveString$ = MoveString$ & Chr$(96 + svColE) & Trim$(Str$(svRowE))
         
         If ThePromPiece$ <> "" Then MoveString$ = MoveString$ & "=" & ThePromPiece$
         
         If MNUM = 1 Then  ' Immediate mate
            Message$ = MoveString$ & "   " & "CHECKMATE"
         ElseIf MNUM >= 3 Then
            Message$ = Str$(MM) & " Key move = " & MoveString$
            If aWKingInCheck Or aBKingInCheck Then Message$ = Message$ & "CHECK"
         End If
      Else
         MM = (NMates + 1) \ 2  ' Mate depth
         Message = ""
         TEST_FOR_MATE = False
      End If
End Function
      

