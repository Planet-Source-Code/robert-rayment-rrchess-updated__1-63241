Attribute VB_Name = "ModCompMove"
' ModCompMove.bas  ~RRChess~

Option Explicit

Public LBRand As Long   ' Random variation for CountnScore Bishops
Public LNRand As Long   ' Random variation for CountnScore Knights
Public mmul As Long ' Material score multiplier

Private Cnt As Integer '= 0 ,Cnt,Cnt+1,, ODD EVEN
Private OrderARR() As Long
Private OrderARR2() As Long

Private NM() As Long
Private rrs As Long, ccs As Long
Private rre As Long, cce As Long
Private svMessage$


Public Sub CompMove(Index As Integer)
' Has to return
' Message$ "CHECK","CHECKMATE", STALEMATE", etc
' ippt,
' PieceTaken
' MoveString$
Dim MM As Long
Dim MvCount As Long
Dim Comp$   ' Comp color
Dim MateNum As Long
Dim R As Long
Dim C As Long
Dim INCheck
   
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' Check for Draw for lack of material
      If WhitePieceCount = 1 And BlackPieceCount = 1 Then
         Message$ = "DRAW"
         Exit Sub
      End If
      If WhitePieceCount <= 2 And BlackPieceCount <= 2 Then
         If NumWR = 0 And NumBR = 0 Then
         If NumWQ = 0 And NumBQ = 0 Then
         If NumWP = 0 And NumBP = 0 Then
            Message$ = "DRAW"
            Exit Sub
         End If
         End If
         End If
      End If
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   ReDim OrderARR(1 To 100)
   ReDim OrderARR2(1 To 100)
   Message$ = ""            ' Public
   Select Case CorHsMove$   ' Public
   Case "CW":
      Comp$ = "W" ' Comp plays "W" on Board Index
   Case "CB"
      Comp$ = "B" ' Comp plays "B" on Board Index
   Case "CCW"
      Comp$ = "W" ' Comp plays "W" on Board Index
   Case "CCB"
      Comp$ = "B" ' Comp plays "B" on Board Index
   End Select
   ' In check flag
   INCheck = 0
   ' Get number of legal moves
   If Comp$ = "B" Then
      MvCount = MoveCounter("B", Index, 1, 0)
      If aBKingInCheck Then INCheck = 1
   Else
      MvCount = MoveCounter("W", Index, 1, 0)
      If aWKingInCheck Then INCheck = 1
   End If
   CountPieces Index    ' PieceCount
   Form1.LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
   Form1.LabPC(0).Refresh
   Form1.LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
   Form1.LabPC(1).Refresh
   
   '  Mate filtering
   MateNum = 5 ' ie Mate in 3
   
   If INCheck = 0 Then  ' No long mate if in check
      If MvCount < 16 And PieceCount < 16 Then MateNum = 7 ' ie Mate in 4
      If MvCount < 6 And PieceCount < 6 Then MateNum = 9   ' ie Mate in 5
   End If
   ' CHECKMATE?
   If TEST_FOR_MATE(Comp$, MateNum, MM) Then   ' Checkmate in 2,3,4 or 5 moves
      RestoreSavedPosition Form1, 1 ' Restore display from Board 1
      CopyMemory bRCBoard(1, 1, 0), bRCBoard(1, 1, 1), 64   ' Make Board 0 the same
      CopyALLBOOLS 1, 0 ' 1 -> 0
      SavePosition Form1, 1            ' Display -> Board 1
      Exit Sub
   End If
   If Message = "HALTED" Then Exit Sub
   
   ' Again?
   SavePosition Form1, 1   ' Display -> Board 1
   CopyALLBOOLS 0, 1 ' 0 -> 1

   ' CAPTURES ??

   ' OPENINGS
   If PieceCount = 32 Then
      If Comp$ = "B" Then
         Form1.LabPB(2) = "B"
         Form1.LabPB(1) = ""
         Form1.LabPB(1).Refresh
         Form1.LabPB(0) = ""
         Form1.LabPB(1).Refresh
         If BlackOMoves(Index) And HalfMove <= 2 Then
            Exit Sub
         End If
      Else  ' Comp$="W"
         Form1.LabPB(2) = "B"
         Form1.LabPB(1) = ""
         Form1.LabPB(1).Refresh
         Form1.LabPB(0) = ""
         Form1.LabPB(1).Refresh
         If WhiteOMoves(Index) And HalfMove <= 1 Then
            Exit Sub
         End If
      End If
   End If

   Randomize
   ' INTRODUCE SOME VARIATION IN MOVES
   ' WHEN COMPUTER PLAYS COMPUTER
   ' for use in CountnScore
   If CorHsMove$ = "CCW" Or CorHsMove$ = "CCB" Then
      LBRand = CLng(10 * (Rnd - 0.5))
      LNRand = CLng(10 * (Rnd - 0.5))
   Else
      ' No variation when Computer versus Human
      LBRand = 1
      LNRand = 1
   End If
   
      ' FIND A COMPUTER MOVE
         
      If FIND_A_MOVE(Comp$, Index) Then
         If Message = "HALTED" Then Exit Sub
         If Message = "CHECKMATE" Then Exit Sub
         
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         ' 3 Repetitions ?
         CopyMemory RepPositions(1, 1, NumRepPositions), bRCBoard(1, 1, 1), 64
         NumRepPositions = NumRepPositions + 1
         If NumRepPositions > 7 Then NumRepPositions = 1

         If NumRepPositions = 6 Then
            svMessage$ = Message$
            Message$ = "DRAW"
            For R = 1 To 8
            For C = 1 To 8
               If bRCBoard(R, C, 1) <> RepPositions(R, C, 3) Then
                 Message$ = "": Exit For
               End If
               If RepPositions(R, C, 3) <> RepPositions(R, C, 1) Then
                  Message$ = "": Exit For
               End If
            Next C
            If Message$ = "" Then
               Message$ = svMessage$
               NumRepPositions = 1
               Exit For
            End If
            Next R
         End If
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
               
      Else  ' STALEMATE ?
         If Message$ = "HALTED" Then Exit Sub
         Message$ = "STALEMATE"
         Exit Sub
      End If
End Sub

Public Function FIND_A_MOVE(Comp$, ByVal Index As Integer) As Boolean
'Comp$ computer color
Dim K1 As Long, J1 As Long
Dim K2 As Long, J2 As Long
Dim K3 As Long  ', J3 As Long
Dim Cnt As Integer
Dim PN As Long
Dim DPN As Long
Dim TScore As Long
Dim BestK1 As Long

Dim alpha1 As Long   ' Max Comp score
Dim alpha2 As Long
Dim alpha3 As Long
Dim beta1 As Long    ' Min Comp Score
Dim beta2 As Long
'Dim beta3 As Long
 
   FIND_A_MOVE = False
   
   ReDim NM(1 To 100) As Long
   BestK1 = 1     ' BestMove in K1 outer For loop
   
   Cnt = 1

' 1st COMP MOVE
alpha1 = -100000
   GETMOVES Comp$, Index, Cnt ' 1 Saves Board & BOOLS & returns no. of legal moves  ODD
                              '   Board Index to TempBoard Cnt, BOOLS to svstore Cnt
   If NM(Cnt) = 0 Then
      Message$ = "STALEMATE"
      Exit Function  ' STALEMATE
   End If
   
   Form1.LabPB(2) = "C"
   Form1.LabPB(1) = NM(Cnt)
   Form1.LabPB(1).Refresh
   
   If NM(Cnt) = 1 Then
      BestK1 = 1
      GoTo MakeMove2
   End If
   
   OrderMoves Comp$, Index, Cnt, NM(Cnt)
   For K1 = 1 To NM(Cnt)    ' 1st CMoves B/W
      Form1.LabPB(0) = K1
      DoEvents
      If TestKeys() Then UNDOMOVE Index, 1: GoTo MakeMove2
      MAKE_A_MOVE Comp$, K1, Index, Cnt, PN, DPN

' 1st HUMAN MOVE
beta1 = 100000
      GETMOVES Comp$, Index, Cnt + 1 ' 2 Returns NM(Cnt+1) & all Legal W/B HMoves  EVEN
      If NM(Cnt + 1) = 0 Then
         GoTo Comp1 ' No HMove
      End If
      OrderMoves Comp$, Index, Cnt + 1, NM(Cnt + 1)
      For J1 = 1 To NM(Cnt + 1)  ' 1st HMove
         If TestKeys() Then UNDOMOVE Index, 1: GoTo MakeMove2
         MAKE_A_MOVE Comp$, J1, Index, Cnt + 1, PN, DPN
         
         If CheckOnKing(Comp$, Index, J1, Cnt + 1, DPN, TScore) Then
            ' Mate try another 1st CMove
            GoTo Comp1
         End If
         
' 2nd COMP MOVE
alpha2 = -100000
         GETMOVES Comp$, Index, Cnt + 2 ' 3 Returns NM(Cnt+2) & all Legal B/W 2nd CMoves  ODD
         If NM(Cnt + 2) = 0 Then
            GoTo Hum1  ' No B/W move
         End If
         OrderMoves Comp$, Index, Cnt + 2, NM(Cnt + 2)
         For K2 = 1 To NM(Cnt + 2)  ' 2nd CMove
            If TestKeys() Then UNDOMOVE Index, 1: GoTo MakeMove2
            MAKE_A_MOVE Comp$, K2, Index, Cnt + 2, PN, DPN

' 2nd HUMAN MOVE
beta2 = 100000
            GETMOVES Comp$, Index, Cnt + 3 ' 4 Returns NM(Cnt+3) & all Legal W/B HMoves  EVEN
            If NM(Cnt + 3) = 0 Then
               GoTo Comp2 ' No W/B move
            End If
            OrderMoves Comp$, Index, Cnt + 3, NM(Cnt + 3)
            For J2 = 1 To NM(Cnt + 3)  ' 2nd HMove
               If TestKeys() Then UNDOMOVE Index, 1: GoTo MakeMove2
               MAKE_A_MOVE Comp$, J2, Index, Cnt + 3, PN, DPN
               
               If PLY = 4 Then
               ' 4-Ply
                  CountnScore Index
                  If Comp$ = "B" Then
                     TScore = BScore - WScore
                  Else
                     TScore = WScore - BScore
                  End If
                  CheckOnKing Comp$, Index, J2, Cnt + 3, DPN, TScore
                  If TScore < beta2 Then beta2 = TScore ' beta2 = Min CMove score
                  GoTo Hum2
               End If

               ' 5-Ply only
               If CheckOnKing(Comp$, Index, J2, Cnt + 3, DPN, TScore) Then
                  ' Mate try another 1st CMove
                   GoTo Comp1
               End If
' 3rd COMP MOVE
alpha3 = -100000
               GETMOVES Comp$, Index, Cnt + 4 ' 5 Returns NM(Cnt+4) & all Legal B/W CMoves  ODD
               If NM(Cnt + 4) = 0 Then
                  GoTo Comp1
               End If
               OrderMoves Comp$, Index, Cnt + 4, NM(Cnt + 4)
               For K3 = 1 To NM(Cnt + 4)  ' 3rd CMove
                  If TestKeys() Then UNDOMOVE Index, 1: GoTo MakeMove2
                  MAKE_A_MOVE Comp$, K3, Index, Cnt + 4, PN, DPN

                     CountnScore Index
                     If Comp$ = "B" Then
                        TScore = BScore - WScore
                     Else
                        TScore = WScore - BScore
                     End If
                     
                     CheckOnKing Comp$, Index, K3, Cnt + 4, DPN, TScore
                     If TScore > alpha3 Then alpha3 = TScore ' Max CMove score = Min HMove score

Comp3:              UNDOMOVE Index, Cnt + 4
               Next K3
               
               If alpha3 <= beta2 Then
                  beta2 = alpha3
                  If beta2 <= alpha2 Then
                     Exit For
                  End If
               End If

Hum2:           UNDOMOVE Index, Cnt + 3
            Next J2
            
            If beta2 >= alpha2 Then
               alpha2 = beta2
               If alpha2 >= beta1 Then
                  Exit For
               End If
            End If

Comp2:        UNDOMOVE Index, Cnt + 2
         Next K2
         
         If alpha2 <= beta1 Then   ' If Max Comp2 <= Min Hum1
            beta1 = alpha2         ' Min Hum1 = Max Comp2
            If beta1 <= alpha1 Then    ' If Min Hum1 <= Max Comp1
               Exit For
            End If
         End If
         
Hum1:     UNDOMOVE Index, Cnt + 1
      Next J1
      
      If beta1 > alpha1 Then
         alpha1 = beta1
         BestK1 = K1
      End If
                              ' TempBoard Cnt to Index, svStore Cnt BOOLS to Index
Comp1:  UNDOMOVE Index, Cnt   ' Restores Board & BOOLS
   Next K1

MakeMove2:
   If Not MakeBESTMove(Comp$, BestK1, Index, PN, DPN) Then
      Exit Function
   End If
   FIND_A_MOVE = True
End Function

Public Function TestKeys() As Boolean
   TestKeys = False
   If GetAsyncKeyState(vbKeyLButton) And &H8000 Then DoEvents 'Form1.Refresh
   If GetAsyncKeyState(vbKeyEscape) And &H8000 Then TestKeys = True
   If GetAsyncKeyState(vbKeyH) And &H8000 Then
      TestKeys = True
      Message = "HALTED"
   End If
   Sleep 0
End Function

Public Function CheckOnKing(WB$, Index As Integer, JK As Long, C As Integer, DPN As Long, TScore As Long) As Boolean
' Private DPN
' IN: WB$=CompCul, Index Board, JK inner move (EVEN), C=Cnt+3 or +5, DPN Dest Piece Number
'     WB$=CompCul, Index Board, JK inner move (ODD), C=Cnt+4 DPN Dest Piece Number 5-PLy
' OUT: TScore
   CheckOnKing = False
   If (C And 1) = 0 Then ' EVEN
         If WB$ = "B" Then ' Comp "B" White HMove here
            If DPN = 11 Then ' taking BK ?? shoudn't happen
               TScore = -100000
            End If
            If WMoveStore(7, C, JK) = 1 Then ' BK in check?
               TScore = TScore - 10
               If TestCheckMate("W", "B", Index) Then
                  CheckOnKing = True
                  TScore = -100000 ' Hi score for "W"
               End If
            End If
         Else  ' Comp "W"  Black HMove here
            If DPN <> 0 Then
               If DPN = 5 Then ' taking WK ?? shoudn't happen
                  TScore = -100000
               End If
            End If
            If BMoveStore(7, C, JK) = 1 Then ' WK in check?
               TScore = TScore - 10
               If TestCheckMate("B", "W", Index) Then
                  CheckOnKing = True
                  TScore = -100000 ' Hi score for "B"
                  Exit Function
               End If
            End If
         End If
   Else  ' ODD  Odd-Ply
         If WB$ = "B" Then ' Comp "B" Black Move here
            If DPN <> 0 Then
               If DPN = 5 Then ' taking WK ?? shoudn't happen
                  TScore = 100000
               End If
               If BMoveStore(7, C, JK) = 1 Then ' WK in check?
                  TScore = TScore + 10
                  ' Comp checkmate already tested to >= 6-Ply
               End If
            End If
         Else  ' Comp "W"  White Move here
            If DPN = 11 Then ' taking BK ?? shoudn't happen
               TScore = 100000
            End If
            If WMoveStore(7, C, JK) = 1 Then ' BK in check?
               TScore = TScore + 10
               ' Comp checkmate already tested to >= 6-Ply
            End If
         End If
   
   End If
End Function

Public Sub GETMOVES(WB$, Index As Integer, C As Integer)
' WB$ = Comp color "B" or "W"
' Index = bRCBoard(), BOOLS Index
' C = Cnt, Cnt+1 etc
' Public NM() ' Num of legal moves
   ' Save board
   CopyMemory TempBoard(1, 1, C), bRCBoard(1, 1, Index), 64
   ' Save BOOLS ie castling & en passant flags
   SaveALLBOOLS Index, C   ' Index to svstore C
   
   If (C And 1) = 1 Then     ' ODD count  ie Comp move 1,3,5,
      If WB$ = "B" Then
         NM(C) = MoveCounter("B", Index, C, 2)
      Else
         NM(C) = MoveCounter("W", Index, C, 2)
      End If
   Else  ' EVEN count ie Human move 2,4,6,,
      If WB$ = "B" Then
         NM(C) = MoveCounter("W", Index, C, 2)
      Else
         NM(C) = MoveCounter("B", Index, C, 2)
      End If
   End If
End Sub

Public Sub UNDOMOVE(Index As Integer, C As Integer)
' Index = bRCBoard(), BOOLS Index
' C = Cnt, Cnt+1 etc
   ' Restore board
   CopyMemory bRCBoard(1, 1, Index), TempBoard(1, 1, C), 64
   ' Restore BOOLS ie castling & en passant flags
   RestoreALLBOOLS C, Index ' svstore C to Index
'eg   CopyMemory WKRR(0, Index), svWKRR(0, C), 4

End Sub

   
Public Sub OrderMoves(WB$, Index As Integer, ByVal C As Integer, ByVal UPCount As Long)
Dim PN As Long, DPN As Long
Dim k As Long
Dim TScore As Long
' WB$ = "B" or "W"
' Index = Board Index
' C = Cnt, Cnt+1 etc
' UPCount= num legal moves stored in BMoveStore() or WMoveStore(), (1 To 8, C= 1 To 8, 1 To 100)
' Before calling OrderMoves GETMOVES will be done:-
'   including
'   CopyMemory TempBoard(1, 1, C), bRCBoard(1, 1, Index), 64
'   SaveALLBOOLS Index, C   ' Index to svstore C
'   NM()

   Select Case WB$   ' Comp color "W" or "B"
   Case "B" ' Comp "B"
   
      If (C And 1) = 1 Then ' ODD so CMove "B"
         For k = 1 To UPCount
            MakeBlackMove k, Index, C, PN, DPN
            
            'GETMOVES
            'for kk= 1 to
            '  MakeWhiteMove k, Index, C, PN, DPN
            
            CountnScore Index
            TScore = BScore - WScore
            If C = 1 And TScore < 0 Then TScore = -10000
            OrderARR(k) = TScore
            'UNDOMOVE
            ' Next kk
            
            UNDOMOVE Index, C
         Next k
         Quicksort3D OrderARR(), BMoveStore(), C, UPCount, 1
      Else  ' EVEN so HMove "W"
         For k = 1 To UPCount    ' Eg 1 B/W
            MakeWhiteMove k, Index, C, PN, DPN
            'CountnScore Index
            MatScore Index
            TScore = BScore - WScore
            CheckOnKing WB$, Index, k, C, DPN, TScore
            OrderARR(k) = TScore
            UNDOMOVE Index, C
         Next k
         Quicksort3D OrderARR(), WMoveStore(), C, UPCount, 0  ' 0 min at top, 1 max at top
      End If
   
   Case "W" ' Comp "W"
      If (C And 1) = 1 Then ' ODD so CMove "W"
         For k = 1 To UPCount
            MakeWhiteMove k, Index, C, PN, DPN
            CountnScore Index
            TScore = WScore - BScore
            OrderARR(k) = TScore
            UNDOMOVE Index, C
         Next k
         Quicksort3D OrderARR(), WMoveStore(), C, UPCount, 1
      Else  ' EVEN so HMove "B"
         For k = 1 To UPCount
            MakeBlackMove k, Index, C, PN, DPN
            'CountnScore Index
            MatScore Index
            TScore = WScore - BScore
            CheckOnKing WB$, Index, k, C, DPN, TScore
            OrderARR(k) = TScore
            UNDOMOVE Index, C
         Next k
         Quicksort3D OrderARR(), BMoveStore(), C, UPCount, 0
      End If
   End Select
End Sub

Public Sub MAKE_A_MOVE(WB$, kj As Long, Index As Integer, C As Integer, PN As Long, DPN As Long)
' kj = k1,j1,k2,j2,,,
' Index = Board index
' C = Cnt, Cnt+1,,
' PN = piece number
' DPN = destination piece number
   If (C And 1) = 1 Then ' ODD CMove
      If WB$ = "B" Then ' CMove
         MakeBlackMove kj, Index, C, PN, DPN
      Else
         MakeWhiteMove kj, Index, C, PN, DPN
      End If
   Else ' EVEN HMove
      If WB$ = "B" Then
         MakeWhiteMove kj, Index, C, PN, DPN
      Else
         MakeBlackMove kj, Index, C, PN, DPN
      End If
   End If
End Sub

Public Function MakeBESTMove(WB$, MoveNum As Long, Index As Integer, PN As Long, DPN As Long) As Boolean
'Private rrs As Long, ccs As Long
'Private rre As Long, cce As Long
' ENTER WITH
' Legal move tables:-
' NLMB = MoveCounter("B", Index, 1, 2)
' NLMW = MoveCounter("W", Index, 1, 2)

' MakeBlackMove(k As Long, Index As Integer, Lev As Integer, PN As Long, DPN As Long)
'Dim P$
   MakeBESTMove = False
   If WB$ = "B" Then
      ' Best B move
      MakeBlackMove MoveNum, Index, 1, PN, DPN
      DestPieceNum = DPN      ' Public DestPieceNum for showing piece taken
      CountPieces Index
      If Not FindKingRC("W", rwk, cwk, Index) Then
         MsgBox "ERROR: Comp move - No W King in MakeBESTMove"
         Exit Function
      Else
         If RC_Targetted(rwk, cwk, "B", Index) Then
            Message$ = "CHECK"
            GoTo MakeMoveString
         End If
      End If
   Else
      ' Best W move
      MakeWhiteMove MoveNum, Index, 1, PN, DPN
      DestPieceNum = DPN      ' Public DestPieceNum for showing piece taken
      CountPieces Index
      If Not FindKingRC("B", rbk, cbk, Index) Then
         MsgBox "ERROR: Comp move - No B King  in MakeBESTMove"
         Exit Function
      Else
         If RC_Targetted(rbk, cbk, "W", Index) Then
            Message$ = "CHECK"
         End If
      End If
   End If

MakeMoveString:
   
   AssembleMoveString PN, rrs, ccs, rre, cce, DPN
   
   MakeBESTMove = True
End Function

 Public Sub AssembleMoveString(PN As Long, rrs As Long, ccs As Long, rre As Long, cce As Long, DPN As Long)
 Dim P$
 Dim ShortString$
 
   ShortString$ = LTrim$(Str$(PN))
   If Len(ShortString$) = 1 Then ShortString$ = "0" & ShortString$
   ShortString$ = ShortString$ & LTrim$(Str$(rrs)) & LTrim$(Str$(ccs))
   ShortString$ = ShortString$ & LTrim$(Str$(rre)) & LTrim$(Str$(cce))
   
   ConvPNtoPNDescrip PN, P$
   
   
   MoveString$ = P$ & " "
   MoveString$ = MoveString$ & Chr$(96 + ccs) & LTrim$(Str$(rrs))
   If DPN = 0 Then
      MoveString$ = MoveString$ & "-"
   Else
      MoveString$ = MoveString$ & "x"
   End If
   MoveString$ = MoveString$ & Chr$(96 + cce) & LTrim$(Str$(rre))
   If PromPiece$ = "Q" Then
      MoveString$ = MoveString$ & "=Q"
   End If
 End Sub




Public Sub MakeBlackMove(k As Long, Index As Integer, C As Integer, PN As Long, DPN As Long)
'Private rrs As Long, ccs As Long
'Private rre As Long, cce As Long
   
   PN = BMoveStore(1, C, k)
   rrs = BMoveStore(2, C, k)
   ccs = BMoveStore(3, C, k)
   rre = BMoveStore(4, C, k)
   cce = BMoveStore(5, C, k)
   DPN = BMoveStore(6, C, k)
   
   bRCBoard(rrs, ccs, Index) = 0
   bRCBoard(rre, cce, Index) = BMoveStore(1, C, k)
   
   PromPiece$ = ""
   BEPP(BEPSet, Index) = 0 ' Cancel BEP
   If PN = BPn Then
      If rre = 1 Then   ' Promotion to Q
         bRCBoard(rre, cce, Index) = BQn
         PromPiece$ = "Q"
      ElseIf Abs(rrs - rre) = Abs(ccs - cce) Then
         If DPN = 0 Then  ' En passant
            If WEPP(WEPSet, Index) = cce Then
               bRCBoard(rrs, cce, Index) = 0
            End If
         End If
      ElseIf rrs = 7 And rre = 5 Then
         ' Set BEP
         BEPP(BEPSet, Index) = ccs
      End If
   End If
   If PN = BKn Then
      If rrs = 8 And rre = 8 Then
         If ccs = 5 And cce = 7 Then ' B Kingside castling
            aCastling(Index) = 0
            If BKRR(BKMoved, Index) = 0 Then
            If BKRR(BKRMoved, Index) = 0 Then
               bRCBoard(8, 8, Index) = 0
               bRCBoard(8, 6, Index) = BRn
               BKRR(BKRMoved, Index) = 2
               'aCastling(Index) = 2
            End If
            End If
         ElseIf ccs = 5 And cce = 3 Then ' B Queenside castling
            aCastling(Index) = 0
            If BKRR(BKMoved, Index) = 0 Then
            If BKRR(BQRMoved, Index) = 0 Then
               bRCBoard(8, 1, Index) = 0
               bRCBoard(8, 4, Index) = BRn
               BKRR(BKRMoved, Index) = 3
               'aCastling(Index) = 2
            End If
            End If
         End If
      End If
      BKRR(BKMoved, Index) = 1   ' BK moved
   End If
   
   If PN = BRn Then
      If rrs = 8 Then
         If ccs = 8 Then
            ' BKR moved
            BKRR(BKRMoved, Index) = 1
         ElseIf ccs = 1 Then
            ' WQR moved
            BKRR(BQRMoved, Index) = 1
         End If
      End If
   End If
   
   WEPP(WEPSet, Index) = 0
End Sub
         
Public Sub MakeWhiteMove(j As Long, Index As Integer, C As Integer, PN As Long, DPN As Long)
'Private rrs As Long, ccs As Long
'Private rre As Long, cce As Long
   
   PN = WMoveStore(1, C, j)
   rrs = WMoveStore(2, C, j)
   ccs = WMoveStore(3, C, j)
   rre = WMoveStore(4, C, j)
   cce = WMoveStore(5, C, j)
   DPN = WMoveStore(6, C, j)
   
   bRCBoard(rrs, ccs, Index) = 0
   bRCBoard(rre, cce, Index) = WMoveStore(1, C, j)
   
   PromPiece$ = ""
   WEPP(WEPSet, Index) = 0 ' Cancel WEP
   If PN = WPn Then
      If rre = 8 Then   ' Promotion to Q
         bRCBoard(rre, cce, Index) = WQn
         PromPiece$ = "Q"
      ElseIf Abs(rrs - rre) = Abs(ccs - cce) Then
         If DPN = 0 Then  ' En passant
            If BEPP(BEPSet, Index) = cce Then
               bRCBoard(rrs, cce, Index) = 0
            End If
         End If
      ElseIf rrs = 2 And rre = 4 Then
         ' Set WEP
         WEPP(WEPSet, Index) = ccs
      End If
   End If
   If PN = WKn Then
      If rrs = 1 And rre = 1 Then
         If ccs = 5 And cce = 7 Then ' W Kingside castling
            aCastling(Index) = 0
            If WKRR(WKMoved, Index) = 0 Then
            If WKRR(WKRMoved, Index) = 0 Then
               bRCBoard(1, 8, Index) = 0
               bRCBoard(1, 6, Index) = WRn
               WKRR(WKRMoved, Index) = 2
               aCastling(Index) = 1
            End If
            End If
         ElseIf ccs = 5 And cce = 3 Then ' W Queenside castling
            aCastling(Index) = 0
            If WKRR(WKMoved, Index) = 0 Then
            If WKRR(WQRMoved, Index) = 0 Then
               bRCBoard(1, 1, Index) = 0
               bRCBoard(1, 4, Index) = WRn
               WKRR(WKRMoved, Index) = 3
               aCastling(Index) = 1
            End If
            End If
         End If
      End If
      WKRR(WKMoved, Index) = 1   ' WK moved
   End If
   If PN = WRn Then
      If rrs = 1 Then
         If ccs = 8 Then
            ' WKR moved
            WKRR(WKRMoved, Index) = 1
         ElseIf ccs = 1 Then
            ' WQR moved
            WKRR(WQRMoved, Index) = 1
         End If
      End If
   End If
   
   BEPP(BEPSet, Index) = 0
End Sub


Public Function MoveCounter(WB$, ByVal Index As Integer, ByVal C As Long, Optional PNLegMoves As Long = 0) As Long

' Collects Legal moves "B" or "W", for Board Index at level C

' Counts total number of legal moves for WB$ color on board Index
' Optionally:-
' PNLegMoves = 0  Total number of legal move for W or B (default)
' PNLegMoves = 1  plus find NumLegMovesPN for all W or B pieces
' PNLegMoves = 2  plus store PN,rs,cs-rd,cd,DPN for all W or B pieces
'Dim sW As Long ' TEST

Dim PN As Long
Dim NW As Long, NB As Long
   ReDim CountBoard(1 To 8, 1 To 8)
   
   CopyMemory CountBoard(1, 1), bRCBoard(1, 1, Index), 64 ' <-
   SaveALLBOOLS Index, MaxIndex + 2
   
   MoveCounter = 0
   NW = 0: NB = 0
   If PNLegMoves > 0 Then
   If WB$ = "W" Then
      NumLegMovesWR = 0
      NumLegMovesWN = 0
      NumLegMovesWB = 0
      NumLegMovesWQ = 0
      NumLegMovesWK = 0
      NumLegMovesWP = 0
   Else
      NumLegMovesBR = 0
      NumLegMovesBN = 0
      NumLegMovesBB = 0
      NumLegMovesBQ = 0
      NumLegMovesBK = 0
      NumLegMovesBP = 0
   End If
   End If
   
   cs(Index) = 0
   rs(Index) = 1
   
   Do
      ' Restore board
      CopyMemory bRCBoard(1, 1, Index), CountBoard(1, 1), 64 ' <-
      RestoreALLBOOLS MaxIndex + 2, Index
      
      cs(Index) = cs(Index) + 1
      If cs(Index) > 8 Then
        cs(Index) = 1
        rs(Index) = rs(Index) + 1
      End If
      
      If rs(Index) > 8 Then
         cs(Index) = 0
         rs(Index) = 1
         Exit Function ' >>>>>
      Else
          ' Get a Piece
          PN = bRCBoard(rs(Index), cs(Index), Index)
          If PN <> 0 Then
          If (WB$ = "W" And PN <= WPn) Or _
             (WB$ = "B" And PN >= BRn) Then
             
             PieceNum(Index) = PN
             M(Index) = 1: Q(Index) = 1
             Do
               If TheMoveOK(Index) Then
      
                  ' Have PieceNum PN &
                  ' Returns with Public :-
                  ' DestPieceNum, PieceIndex$(Index), _
                  ' rs(Index), cs(Index), rd(Index), cd(Index)
                  ' aWKingInCheck, aBKingInCheck
                  ' all legal moves
               
                   
                  If (WB$ = "W" And Not aWKingInCheck) Or _
                     (WB$ = "B" And Not aBKingInCheck) Then
                   
                     MoveCounter = MoveCounter + 1   ' Total number of moves for W or B
                   
                     If PNLegMoves > 0 Then   ' Number of legal moves for each piece
                       If WB$ = "W" Then
                          Select Case PN
                          Case WRn: NumLegMovesWR = NumLegMovesWR + 1
                          Case WNn: NumLegMovesWN = NumLegMovesWN + 1
                          Case WBn: NumLegMovesWB = NumLegMovesWB + 1
                          Case WQn: NumLegMovesWQ = NumLegMovesWQ + 1
                          Case WKn: NumLegMovesWK = NumLegMovesWK + 1
                          Case WPn: NumLegMovesWP = NumLegMovesWP + 1
                          End Select
                       Else  ' WB$ = "B"
                          Select Case PN
                          Case BRn: NumLegMovesBR = NumLegMovesBR + 1
                          Case BNn: NumLegMovesBN = NumLegMovesBN + 1
                          Case BBn: NumLegMovesBB = NumLegMovesBB + 1
                          Case BQn: NumLegMovesBQ = NumLegMovesBQ + 1
                          Case BKn: NumLegMovesBK = NumLegMovesBK + 1
                          Case BPn: NumLegMovesBP = NumLegMovesBP + 1
                          End Select
                       End If
                     End If
                     
                     If PNLegMoves = 2 Then ' Store legal moves
                       If WB$ = "W" Then
                          NW = NW + 1
                          WMoveStore(1, C, NW) = PN
                          WMoveStore(2, C, NW) = rs(Index)
                          WMoveStore(3, C, NW) = cs(Index)
                          WMoveStore(4, C, NW) = rd(Index)
                          WMoveStore(5, C, NW) = cd(Index)
                          WMoveStore(6, C, NW) = DestPieceNum ' 0 or Opp PN
                          WMoveStore(7, C, NW) = 0
                          If aBKingInCheck Then   ' W puts BK in check @ legal move NW
                             WMoveStore(7, C, NW) = 1
                          Else
                             WMoveStore(7, C, NW) = 0
                          End If
                          If DestPieceNum <> 0 Then
                             WMoveStore(8, C, NW) = 1 ' A capture W on B
                          Else
                             WMoveStore(8, C, NW) = 0 ' No capture
                          End If
                       Else  ' WB$ = "B"
                          NB = NB + 1
                          BMoveStore(1, C, NB) = PN
                          BMoveStore(2, C, NB) = rs(Index)
                          BMoveStore(3, C, NB) = cs(Index)
                          BMoveStore(4, C, NB) = rd(Index)
                          BMoveStore(5, C, NB) = cd(Index)
                          BMoveStore(6, C, NB) = DestPieceNum ' 0 or Opp PN
                          BMoveStore(7, C, NB) = 0
                          If aWKingInCheck Then   ' B puts WK in check @ legal move NB
                             BMoveStore(7, C, NB) = 1
                          Else
                             BMoveStore(7, C, NB) = 0
                          End If
                          If DestPieceNum <> 0 Then
                             BMoveStore(8, C, NB) = 1  ' A capture B on W
                          Else
                             BMoveStore(8, C, NB) = 0  ' No capture
                          End If
                       
                       End If    ' If WB$ = "W" Then
                     End If   ' If PNLegMoves > 0 Then
                  End If   ' If (WB$ = "W" And '''
               End If   ' If TheMoveOK(Index) Then
               
               ' RestorePieces
               CopyMemory bRCBoard(1, 1, Index), CountBoard(1, 1), 64 ' <-
               'CountPieces Index      ''''''''''''''''''''''
               CheckEPPPCA (Index)
               RestoreALLBOOLS MaxIndex + 2, Index
               If Not PieceDirec(M(Index), Q(Index), PieceNum(Index), aMovOK(Index)) Then
                 Exit Do
               End If
             Loop
          End If  ' Piece  0
          End If  ' or wrong color
      End If
   Loop
Exit Function
'=================
End Function



Public Sub Quicksort3D(Arr() As Long, bARR() As Byte, C As Integer, UPCount As Long, Optional AD As Long = 0)
'  Sort as/descending according bARR(1 - NM(Cnt):-
'  WB$="B" BMoveStore(1 To 8, C, 1 To UPCount)
'  WB$="W" WMoveStore(1 To 8, C, 1 To UPCount)
'  UPCount = NM(Cnt), NM(Cnt+1),,
'  AD (default) = 0 ascending order   min at 1  top
'  AD           = 1 descending order  max at 1  top
Dim Max As Long
Dim k As Long
Dim M As Long
Dim s As Long
Dim sortL() As Long
Dim sortR() As Long
Dim LL As Long
Dim MM As Long
Dim II As Long
Dim JJ As Long
Dim PP As Long
Dim XX As Long
Dim YY As Long
Dim bYY As Byte
Dim K1 As Long
   Max = UPCount
   If Max = 1 Then Exit Sub
   k = LBound(Arr)
   If k = Max Then Exit Sub
   M = Max \ 2: ReDim sortL(M), sortR(M)
   s = 1: sortL(1) = k: sortR(1) = Max
   Do While s <> 0
      LL = sortL(s): MM = sortR(s): s = s - 1
      
      Do While LL < MM
         II = LL: JJ = MM
         PP = (LL + MM) \ 2
         XX = Arr(PP)
         Do While II <= JJ
            If AD = 0 Then ' Ascending
               Do While Arr(II) < XX: II = II + 1: Loop
               Do While XX < Arr(JJ): JJ = JJ - 1: Loop
            Else  ' Descending
               Do While Arr(II) > XX: II = II + 1: Loop
               Do While XX > Arr(JJ): JJ = JJ - 1: Loop
            End If
            
            If II <= JJ Then
               ' SWAP Arr(II), Arr(JJ)
               YY = Arr(II): Arr(II) = Arr(JJ): Arr(JJ) = YY
               ' Also SWAP bArr(1 to 8, C to 8, II), bArr(1 To 8, C To 8, JJ)
               ' ie BMoveStore(1 to 8, C, II) or WMoveStore(1 to 8, C, II)
               For K1 = 1 To 8
                  bYY = bARR(K1, C, II)
                  bARR(K1, C, II) = bARR(K1, C, JJ)
                  bARR(K1, C, JJ) = bYY
               Next K1
               II = II + 1: JJ = JJ - 1
            End If
         Loop
         If II < MM Then
            s = s + 1: sortL(s) = II: sortR(s) = MM
         End If
         MM = JJ
      Loop
   Loop
Erase sortL, sortR
End Sub

Public Sub CountnScore(ByVal Index As Integer)

' WHITE/BLACK LEGAL MOVES NEEDS CALLING FIRST IE
' MoveCounter(WB$, Index, Optional PNLegMoves As Long = 0)
' NLMW = MoveCounter("W", Index, PNLegMoves)
' NLMB = MoveCounter("B", Index, PNLegMoves)
' to give NumLegMovesW/B PN

'TEST
Dim sW As Long, sB As Long

Dim k As Long
Dim ap As Long ' adjacent pawn
Dim j As Long
Dim PN As Long
Dim CompColor$
Dim WMobility As Long, BMobility As Long

' Rest Public
   
   If Index = 0 Then ' Copy all to Board 1
      CopyMemory bRCBoard(1, 1, Index + 1), bRCBoard(1, 1, Index), 64 ' <-
      CopyALLBOOLS Index, Index + 1 ' ->
      Index = 1
   End If
   
   CompColor$ = WorBsMove ' Comps color
   WMobility = 0
   BMobility = 0
   
   MatScore Index
   ' Returns NumWR,NumBR,, etc
   ' PieceCount, WhitePieceCount, BlackPieceCount
   ' WRrow(1),WRcol(1),WRrow(2),WRcol(2),BRrow(1),BRcol(1),, etc
   
   ' GET POSITION SCORES USING PIECE'S LOCATIONS
   
   ' WHITE ROOKS
   If NumWR >= 1 Then
      WMobility = WMobility + 2 * Fix(Sqr(NumLegMovesWR))
      ' Have R,C  +/- score from score-boards
      WScore = WScore + WRook(WRrow(1), WRcol(1))
      
      If WKRR(WKMoved, Index) = 0 Then  ' WK not moved yet
         If WRrow(1) = 1 Then ' - (R->N1)  QR
            If WRcol(1) = 2 Or WRcol(1) = 7 Then WScore = WScore - 5
         End If
      Else ' K moved or castled
         If WRrow(1) > 6 Then WScore = WScore + 5  ' + Forwarding
      End If
      If WRcol(1) = BKcol Then WScore = WScore + 2 ' + R on same column as opp K
      If WRrow(1) = BKrow Then WScore = WScore + 2 ' + R on same row as opp K
      'WR-WQ
      If NumWQ > 0 Then ' WR1, rarely >1 Queen
         If WRcol(1) = WQcol(1) Then WScore = WScore + 2 ' + R on same column as Q
      End If
      DIS_Score "W", WRrow(1), WRcol(1)
      If NumWR >= 2 Then
         WScore = WScore + WRook(WRrow(2), WRcol(2))
         If WKRR(WKMoved, Index) = 0 Then ' WK not moved
            If WRrow(2) = 1 Then  ' - (R->N1)   KR
               If WRcol(2) = 2 Or WRcol(2) = 7 Then WScore = WScore - 5
            End If
         Else ' K moved or castled
            If WRrow(2) > 6 Then WScore = WScore + 5  ' + Forwarding
         End If
         If WRcol(2) = BKcol Then WScore = WScore + 2 ' R on same column as opp K
         If WRrow(2) = BKrow Then WScore = WScore + 2 ' R on same row as opp K
         If WRcol(1) = WRcol(2) Then ' + Rs on same column
            If Abs(WRrow(1) - WRrow(2)) = 1 Then
               WScore = WScore + 5  ' + Rs adjacent
            Else ' Check if clear space between Rooks
               If WRrow(1) < WRrow(2) Then
                  For k = WRrow(1) + 1 To WRrow(2) - 1
                     If bRCBoard(k, WRcol(1), Index) <> 0 Then Exit For
                  Next k
                  If k = WRrow(2) Then WScore = WScore + 5
               Else
                  For k = WRrow(2) + 1 To WRrow(1) - 1
                     If bRCBoard(k, WRcol(1), Index) <> 0 Then Exit For
                  Next k
                  If k = WRrow(1) Then WScore = WScore + 5
               End If
            End If
         End If
         'WR-WQ
         If NumWQ > 0 Then ' WR2, rarely >1 Queen
            If WRcol(2) = WQcol(1) Then WScore = WScore + 2 ' + R on same column as Q
         End If
         DIS_Score "W", WRrow(2), WRcol(2)
      End If
   End If
   
   ' BLACK ROOKS
   If NumBR >= 1 Then
      BMobility = BMobility + 2 * Fix(Sqr(NumLegMovesBR))
      BScore = BScore + BRook(BRrow(1), BRcol(1))
      
      If BKRR(BKMoved, Index) = 0 Then ' BK not moved
         If BRrow(1) = 8 Then ' - (R->N1)
            If BRcol(1) = 2 Or BRcol(1) = 7 Then BScore = BScore - 5
         End If
      Else  ' K moved or castled
         If BRrow(1) < 3 Then BScore = BScore + 5  ' + Forwarding
      End If
      If BRcol(1) = WKcol Then BScore = BScore + 2 ' + R on same column as opp K
      If BRrow(1) = WKrow Then BScore = BScore + 2 ' + R on same row as opp K
      'BR-BQ
      If NumBQ > 0 Then ' BR1, rarely >1 Queen
         If BRcol(1) = BQcol(1) Then BScore = BScore + 2 ' + R on same column as Q
      End If
      DIS_Score "B", BRrow(1), BRcol(1)
      
      If NumBR >= 2 Then
         BScore = BScore + BRook(BRrow(2), BRcol(2))
         If BKRR(BKMoved, Index) = 0 Then ' BK not moved
            If BRrow(2) = 8 Then ' - (R->N1)
               If BRcol(2) = 2 Or BRcol(2) = 7 Then BScore = BScore - 5
            End If
         Else  ' K moved or castled
            If BRrow(2) < 3 Then BScore = BScore + 5  ' + Forwarding
         End If
         If BRcol(2) = WKcol Then BScore = BScore + 2 ' R on same column as opp K
         If BRrow(2) = WKrow Then BScore = BScore + 2 ' R on same row as opp K
         If BRcol(1) = BRcol(2) Then ' + Rs on same column
            If Abs(BRrow(1) - BRrow(2)) = 1 Then
               BScore = BScore + 5  ' + Rs adjacent
            Else ' Check if clear space between Rooks
               If BRrow(1) < BRrow(2) Then
                  For k = BRrow(1) + 1 To BRrow(2) - 1
                     If bRCBoard(k, BRcol(1), Index) <> 0 Then Exit For
                  Next k
                  If k = BRrow(2) Then BScore = BScore + 5
               Else
                  For k = BRrow(2) + 1 To BRrow(1) - 1
                     If bRCBoard(k, BRcol(1), Index) <> 0 Then Exit For
                  Next k
                  If k = BRrow(1) Then BScore = BScore + 5
               End If
            End If
         End If
         'BR-BQ
         If NumBQ > 0 Then ' BR2, rarely >1 Queen
            If BRcol(2) = BQcol(1) Then BScore = BScore + 2 ' + R on same column as Q
         End If
         DIS_Score "B", BRrow(2), BRcol(2)
      End If
   End If

   ' WHITE KNIGHTS
   If NumWN >= 1 Then
      WMobility = WMobility + 2 * Fix(Sqr(NumLegMovesWN))
      WScore = WScore + WKnight(WNrow(1), WNcol(1)) * LNRand
      
      DIS_Score "W", WNrow(1), WNcol(1)
      If NumWN = 2 Then
         WScore = WScore + WKnight(WNrow(2), WNcol(2)) * LNRand
         DIS_Score "W", WNrow(2), WNcol(2)
      End If
   End If
      
   ' BLACK KNIGHTS
   If NumBN >= 1 Then
      BMobility = BMobility + 2 * Fix(Sqr(NumLegMovesBN))
      BScore = BScore + BKnight(BNrow(1), BNcol(1)) * LNRand
      
      DIS_Score "B", BNrow(1), BNcol(1)
      If NumBN = 2 Then
         BScore = BScore + BKnight(BNrow(2), BNcol(2)) * LNRand
         DIS_Score "B", BNrow(2), BNcol(2)
      End If
   End If
   
   ' WHITE BISHOPS
   If NumWB >= 1 Then
      WMobility = WMobility + 2 * Fix(Sqr(NumLegMovesWB))
      If HalfMove < 20 Then
         WScore = WScore + WWBishop_BBBishop_Open(WBrow(1), WBcol(1)) * LBRand
         WScore = WScore + WBBishop_BWBishop_Open(WBrow(1), WBcol(1)) * LBRand
      Else
         WScore = WScore + WWBishop_BBBishop(WBrow(1), WBcol(1)) * LBRand
         WScore = WScore + WBBishop_BWBishop(WBrow(1), WBcol(1)) * LBRand
      End If
      
      If WBcol(1) = WBrow(1) Then WScore = WScore + 5 ' + Bishop on main diagonal
         'WB-WQ
      If NumWQ > 0 Then ' WB1, rarely >1 Queen
         If Abs(WBrow(1) - WQrow(1)) = Abs(WBcol(1) - WQcol(1)) Then ' + B&Q on same diagonal
            WScore = WScore + 5
         End If
      End If
      DIS_Score "W", WBrow(1), WBcol(1)
      If NumWB = 2 Then
         If HalfMove < 20 Then
            WScore = WScore + WWBishop_BBBishop_Open(WBrow(2), WBcol(2)) * LBRand
            WScore = WScore + WBBishop_BWBishop_Open(WBrow(2), WBcol(2)) * LBRand
         Else
            WScore = WScore + WWBishop_BBBishop(WBrow(2), WBcol(2)) * LBRand
            WScore = WScore + WBBishop_BWBishop(WBrow(2), WBcol(2)) * LBRand
         End If
         
         If WBcol(2) = WBrow(2) Then WScore = WScore + 5 ' + Bishop on main diagonal
         'WB-WQ
         If NumWQ > 0 Then ' WB2, rarely >1 Queen
            If Abs(WBrow(2) - WQrow(1)) = Abs(WBcol(2) - WQcol(1)) Then ' + B&Q on same diagonal
               WScore = WScore + 5
            End If
         End If
         DIS_Score "W", WBrow(2), WBcol(2)
      End If
   End If
   
   ' BLACK BISHOPS
   If NumBB >= 1 Then
      BMobility = BMobility + 2 * Fix(Sqr(NumLegMovesBB))
      If HalfMove < 20 Then
         BScore = BScore + WWBishop_BBBishop_Open(BBrow(1), BBcol(1)) * LBRand
         BScore = BScore + WBBishop_BWBishop_Open(BBrow(1), BBcol(1)) * LBRand
      Else
         BScore = BScore + WWBishop_BBBishop(BBrow(1), BBcol(1)) * LBRand
         BScore = BScore + WBBishop_BWBishop(BBrow(1), BBcol(1)) * LBRand
      End If
      
      If BBcol(1) = 9 - BBrow(1) Then BScore = BScore + 5 ' + Bishop on main diagonal
      'BB-BQ
      If NumBQ > 0 Then ' BB1, rarely >1 Queen
         If Abs(BBrow(1) - BQrow(1)) = Abs(BBcol(1) - BQcol(1)) Then ' + B&Q on same diagonal
            BScore = BScore + 5
         End If
      End If
      DIS_Score "B", BBrow(1), BBcol(1)
      If NumBB = 2 Then
         If HalfMove < 20 Then
            BScore = BScore + WWBishop_BBBishop_Open(BBrow(2), BBcol(2)) * LBRand
            BScore = BScore + WBBishop_BWBishop_Open(BBrow(2), BBcol(2)) * LBRand
         Else
            BScore = BScore + WWBishop_BBBishop(BBrow(2), BBcol(2)) * LBRand
            BScore = BScore + WBBishop_BWBishop(BBrow(2), BBcol(2)) * LBRand
         End If
         
         If BBcol(2) = 9 - BBrow(2) Then BScore = BScore + 2 ' + Bishop on main diagonal
         'BB-BQ
         If NumBQ > 0 Then ' BB2, rarely >1 Queen
            If Abs(BBrow(2) - BQrow(1)) = Abs(BBcol(2) - BQcol(1)) Then ' + B&Q on same diagonal
               BScore = BScore + 5
            End If
         End If
         DIS_Score "B", BBrow(2), BBcol(2)
      End If
   End If
   
   ' WHITE QUEEN
   If NumWQ > 0 Then
      If HalfMove < 20 Then
         WMobility = WMobility - NumWQ * Fix(Sqr(NumLegMovesWQ))
         WScore = WScore + WQueen_Open(WQrow(1), WQcol(1))
      Else
         WMobility = WMobility + NumWQ * Fix(Sqr(NumLegMovesWQ))
         WScore = WScore + WBQueen(WQrow(1), WQcol(1))
         DIS_Score "W", WQrow(1), WQcol(1)
      End If
   End If
   
   ' BLACK QUEEN
   If NumBQ > 0 Then
      If HalfMove < 20 Then
         BMobility = BMobility - NumBQ * Fix(Sqr(NumLegMovesBQ))
         BScore = BScore + BQueen_Open(BQrow(1), BQcol(1))
      Else
         BMobility = BMobility + NumBQ * Fix(Sqr(NumLegMovesBQ))
         BScore = BScore + WBQueen(BQrow(1), BQcol(1))
         DIS_Score "B", BQrow(1), BQcol(1)
      End If
   End If
   
   ' WHITE KING
   If HalfMove < 60 Then
      WScore = WScore + WKing(WKrow, WKcol)
   Else
      WScore = WScore + WBKing_End(WKrow, WKcol)
      DIS_Score "W", WKrow, WKcol
   End If
   
   Select Case WKRR(WKMoved, Index)
   Case 0
      If NumWR > 0 Then
         If WRcol(1) = 1 And WKRR(WQRMoved, Index) = 0 Then
            If WKRR(WKMoved, Index) = 0 Then ' WKing not moved
               WScore = WScore + 3  ' Castling permitted when clear space
            End If
         End If
         If NumWR > 1 Then
            If WRcol(2) = 8 And WKRR(WKRMoved, Index) = 0 Then
               If WKRR(WKMoved, Index) = 0 Then ' WKing not moved
                  WScore = WScore + 3  ' Castling permitted when clear space
               End If
            End If
         End If
      End If
   Case 1   ' WK moved
   Case 2, 3   ' WK castled KSide,QSide
      WScore = WScore + 50  ' Castling done, score once
      WKRR(WKMoved, Index) = 1
   End Select
   
   ' Check pawns in front of WKing after WKing has moved
   If WKrow = 1 And (WKcol <> 5) Then
      For k = WKcol - 1 To WKcol + 1
         If k > 0 And k < 9 Then
            PN = bRCBoard(2, WKcol, Index)
            If PN <> 0 Then
            If PN <= 6 Then
               WScore = WScore + 3 ' + WK safety
            End If
            End If
         End If
      Next k
   End If
   If HalfMove < 40 Then
      WMobility = WMobility - 3 * Fix(Sqr(NumLegMovesWK))
   End If
   
   ' BLACK KING
   If HalfMove < 60 Then
      BScore = BScore + BKing(BKrow, BKcol)
   Else
      BScore = BScore + WBKing_End(BKrow, BKcol)
      DIS_Score "B", BKrow, BKcol
   End If
   
   Select Case BKRR(BKMoved, Index)
   Case 0
      If NumBR > 0 Then
         If BRcol(1) = 1 And BKRR(BQRMoved, Index) = 0 Then
            If BKRR(BKMoved, Index) = 0 Then ' BKing not moved
               BScore = BScore + 3  ' Castling permitted when clear space
            End If
         End If
         If NumBR > 1 Then
            If BRcol(2) = 8 And BKRR(BKRMoved, Index) = 0 Then
               If BKRR(BKMoved, Index) = 0 Then  ' BKing not moved
                  BScore = BScore + 3  ' Castling permitted when clear space
               End If
            End If
         End If
      End If
   Case 1   ' BK moved
   Case 2, 3   ' BK castled KSide,QSide
      BScore = BScore + 50  ' Castling done, score once
      BKRR(BKMoved, Index) = 1
   End Select
   
   ' Check pawns in front of BKing after BKing has moved
   If BKrow = 8 And (BKcol <> 5) Then
      For k = BKcol - 1 To BKcol + 1
         If k > 0 And k < 9 Then
            PN = bRCBoard(7, BKcol, Index)
            If PN >= 7 Then
               BScore = BScore + 3 ' + BK safety
            End If
         End If
      Next k
   End If
   If HalfMove < 40 Then
      BMobility = BMobility - 3 * Fix(Sqr(NumLegMovesBK))
   End If
   
   ' WHITE Pawns
   If NumWP > 0 Then
      For k = 1 To NumWP
         If HalfMove < 40 Then
            WScore = WScore + WPawns_Open(WProw(k), WPcol(k))
         ElseIf HalfMove >= 40 Then
            WScore = WScore + WPawns(WProw(k), WPcol(k))
         End If
         If BlackPieceCount < 8 Then
            WScore = WScore + WPawns_End(WProw(k), WPcol(k))
         End If
      
         If HalfMove > 12 Then
            ' Check for isolated pawn, ie no possible protecting pawns on
            ' adjacent columns. NB double pawns = 2 isolated pawns.
            If WProw(k) >= 2 Then
               ap = 0
               For j = WProw(k) To 2 Step -1
                  If WPcol(k) > 1 Then
                     If bRCBoard(j, WPcol(k) - 1, Index) = WPn Then
                        ap = 1  ' WPawn in adjacent column
                        Exit For
                     End If
                  End If
                  If WPcol(k) < 8 Then
                     If bRCBoard(j, WPcol(k) + 1, Index) = WPn Then
                        ap = 1  ' WPawn in adjacent column
                        Exit For
                     End If
                  End If
               Next j
               If ap = 0 Then BScore = BScore + 15
            End If
         End If
      Next k
   End If
   
   ' BLACK Pawns
   If NumBP > 0 Then
      For k = 1 To NumBP
         If HalfMove < 40 Then
            BScore = BScore + BPawns_Open(BProw(k), BPcol(k))
         ElseIf HalfMove >= 40 Then
            BScore = BScore + BPawns(BProw(k), BPcol(k))
         End If
         If WhitePieceCount < 8 Then
            BScore = BScore + BPawns_End(BProw(k), BPcol(k))
         End If
   
         If HalfMove > 12 Then
            ' Check for isolated pawn, ie no possible protecting pawns on
            '  adjacent columns. NB double pawns = 2 isolated pawns.
            If BProw(k) <= 7 Then
               ap = 0
               For j = BProw(k) To 7
                  If BPcol(k) > 1 Then
                     If bRCBoard(j, BPcol(k) - 1, Index) = BPn Then
                        ap = 1  ' BPawn in adjacent column
                        Exit For
                     End If
                  End If
                  If BPcol(k) < 8 Then
                     If bRCBoard(j, BPcol(k) + 1, Index) = BPn Then
                        ap = 1  ' BPawn in adjacent column
                        Exit For
                     End If
                  End If
               Next j
               If ap = 0 Then WScore = WScore + 15
            End If
         End If
      Next k
   End If
   WScore = WScore + WMobility \ 2
   BScore = BScore + BMobility \ 2
End Sub


Public Sub DIS_Score(WB$, RP, CP)
Dim DIS As Long
   If WB$ = "W" Then
      If BlackPieceCount < 10 Then
            DIS = Sqr((RP - BKrow) * (RP - BKrow) + (CP - BKcol) * (CP - BKcol))
            WScore = WScore + 50 / (DIS + 1)
      End If
   Else
      If WhitePieceCount < 10 Then
            DIS = Sqr((RP - WKrow) * (RP - WKrow) + (CP - WKcol) * (CP - WKcol))
            BScore = BScore + 50 / (DIS + 1)
      End If
   End If
End Sub


Public Sub MatScore(Index As Integer)
Dim R As Long, C As Long
Dim PN As Long
' Rest Public
   
   ' Called from CountnScore
   ' GET MATERIAL SCORE
   ' NUMBER OF PIECES
   ' PIECE'S LOCATIONS
   ' & Attack/Defend Sub  ??
   
   
   If Index = 0 Then ' Copy all to Board 1
      CopyMemory bRCBoard(1, 1, Index + 1), bRCBoard(1, 1, Index), 64 ' <-
      CopyALLBOOLS Index, Index + 1 ' ->
      Index = 1
   End If
   
   NumWR = 0
   NumWN = 0
   NumWB = 0
   NumWQ = 0
   NumWP = 0
   NumWK = 1
   WScore = 0
   
   NumBR = 0
   NumBN = 0
   NumBB = 0
   NumBQ = 0
   NumBP = 0
   NumBK = 1
   BScore = 0
   
   WScore = 0
   BScore = 0
   
   PieceCount = 0
   WhitePieceCount = 0
   BlackPieceCount = 0
   
   For R = 1 To 8
   For C = 1 To 8
      PN = bRCBoard(R, C, Index)
      If PN <> 0 Then
         Select Case PN
         Case WRn: WhitePieceCount = WhitePieceCount + 1
               NumWR = NumWR + 1                         ' Needed for CountnScore
               WRrow(NumWR) = R: WRcol(NumWR) = C        ' Needed for CountnScore
               WScore = WScore + PieceVal(WRn) * mmul    ' mmul set @ Form Initialize (6)
         Case WNn: WhitePieceCount = WhitePieceCount + 1
               NumWN = NumWN + 1
               WNrow(NumWN) = R: WNcol(NumWN) = C
               WScore = WScore + PieceVal(WNn) * mmul
         Case WBn: WhitePieceCount = WhitePieceCount + 1
               NumWB = NumWB + 1
               WBrow(NumWB) = R: WBcol(NumWB) = C
               WScore = WScore + PieceVal(WBn) * mmul
         Case WQn:  WhitePieceCount = WhitePieceCount + 1
               NumWQ = NumWQ + 1
               WQrow(NumWQ) = R: WQcol(NumWQ) = C
               WScore = WScore + PieceVal(WQn) * mmul
         Case WPn:  WhitePieceCount = WhitePieceCount + 1
               NumWP = NumWP + 1
               WProw(NumWP) = R: WPcol(NumWP) = C
               WScore = WScore + PieceVal(WPn) * mmul
         Case WKn: WKrow = R: WKcol = C
         
         Case BRn:  BlackPieceCount = BlackPieceCount + 1
               NumBR = NumBR + 1
               BRrow(NumBR) = R: BRcol(NumBR) = C
               BScore = BScore + PieceVal(BRn) * mmul
         Case BNn:  BlackPieceCount = BlackPieceCount + 1
               NumBN = NumBN + 1
               BNrow(NumBN) = R: BNcol(NumBN) = C
               BScore = BScore + PieceVal(BNn) * mmul
         Case BBn:  BlackPieceCount = BlackPieceCount + 1
               NumBB = NumBB + 1
               BBrow(NumBB) = R: BBcol(NumBB) = C
               BScore = BScore + PieceVal(BBn) * mmul
         Case BQn:  BlackPieceCount = BlackPieceCount + 1
               NumBQ = NumBQ + 1
               BQrow(NumBQ) = R: BQcol(NumBQ) = C
               BScore = BScore + PieceVal(BQn) * mmul
         Case BPn:  BlackPieceCount = BlackPieceCount + 1
               NumBP = NumBP + 1
               BProw(NumBP) = R: BPcol(NumBP) = C
               BScore = BScore + PieceVal(BPn) * mmul
         Case BKn: BKrow = R: BKcol = C
         End Select
      End If
   Next C
   Next R
   PieceCount = WhitePieceCount + BlackPieceCount
   If PieceCount > 0 Then
      PieceCount = PieceCount + 2 ' For the 2 Kings
      WhitePieceCount = WhitePieceCount + 1
      BlackPieceCount = BlackPieceCount + 1
   End If
   'AttackDefend Index
End Sub

' NOT USED SO FAR
'Public Sub AttackDefend(Index As Integer)
'' Called after MatScore so that NumWR, NumBR etc done !!
'' ReDim TestAttackBoard(1 To 8, 1 To 8)
'' Publics AttackingPiece @ rat,cat
'
'' Partial look - only looks at 1 piece attacking/defending unless a King defending
'
'Dim R As Long, C As Long
'Dim PN As Long
'Dim ATPV As Long
'Dim svrat As Long, svcat As Long
'Dim Comp$
'Dim WhiteScore As Long  ' Differential scores
'Dim BlackScore As Long
'
''TEMP
''CorHsMove$ = "B"
'   WhiteScore = 0
'   BlackScore = 0
'   Select Case CorHsMove$   ' Public
'   Case "CW":
'      Comp$ = "W" ' Comp plays "W" on Board Index
'   Case "CB"
'      Comp$ = "B" ' Comp plays "B" on Board Index
'   End Select
'
'   CopyMemory TestAttackBoard(1, 1), bRCBoard(1, 1, Index), 64 ' <-
'   SaveALLBOOLS Index, MaxIndex + 3 ' Index to svstore MaxIndex+3
'
'   For R = 1 To 8
'   For C = 1 To 8
'      PN = bRCBoard(R, C, Index)
'      If PN <> 0 Then
'            Select Case PN    ' bRCBoard(R, C, Index)
'            Case WRn To WPn   ' B -> W
'                  If RC_Targetted(R, C, "B", Index) Then ' B AttackingPiece @ rat,cat
'                     ATPV = -PieceVal(AttackingPiece)
'                     svrat = rat: svcat = cat
'                     WhiteScore = -PieceVal(PN) '* mmul
'                     If RC_Targetted(R, C, "W", Index) Then   ' W Defending
'                        If AttackingPiece = 5 Then  ' WK defending
'                           ' In Check?
'                           ' Clear ATPV and see if another B piece attacks R,C
'                           bRCBoard(svrat, svcat, Index) = 0
'                           If Not RC_Targetted(R, C, "B", Index) Then ' King would not be attacked
'                              WhiteScore = WhiteScore + ATPV '* mmul  ' therefore get some value back
'                           Else  ' K no defence, Check or CheckMate then
'                           End If
'                           CopyMemory bRCBoard(1, 1, Index), TestAttackBoard(1, 1), 64 ' <-
'                           RestoreALLBOOLS MaxIndex + 3, Index  ' svstore MaxIndex + 3 to Index
'                        Else  ' Other than King defending
'                           WhiteScore = WhiteScore + ATPV '* mmul  ' get some value back
'                        End If
'                     Else  ' No defence
'                       'Exit Sub
'                     End If
'               End If
'            End Select
'
'            Select Case PN    ' bRCBoard(R, C, Index)
'            Case BRn To BPn   ' W -> B
'                  If RC_Targetted(R, C, "W", Index) Then ' W AttackingPiece @ rat,cat
'                     ATPV = PieceVal(AttackingPiece)
'                     svrat = rat: svcat = cat
'                     BlackScore = -PieceVal(PN) '* mmul
'                     If RC_Targetted(R, C, "B", Index) Then   ' B Defending
'                        If AttackingPiece = 11 Then  ' BK defending
'                           ' In Check?
'                           ' Clear ATPV and see if another B piece attacks R,C
'                           bRCBoard(svrat, svcat, Index) = 0
'                           If Not RC_Targetted(R, C, "W", Index) Then ' King would not be attacked
'                              BlackScore = BlackScore + ATPV '* mmul
'                           Else  ' K no defence, Check or CheckMate then
'                             'Exit Sub
'                           End If
'                           CopyMemory bRCBoard(1, 1, Index), TestAttackBoard(1, 1), 64 ' <-
'                           RestoreALLBOOLS MaxIndex + 3, Index  ' svstore MaxIndex + 3 to Index
'                        Else  ' King not defending
'                           BlackScore = BlackScore + ATPV '* mmul
'                        End If
'                     Else  ' No defence
'                       'Exit Sub
'                     End If
'                  End If
'            End Select
'      End If
'   Next C
'   Next R
'   ' WhiteScore  B att W def score
'   ' BlackScore  W att B def score
'      BScore = BScore + BlackScore '+ WhiteScore)
'      WScore = WScore + WhiteScore '- BlackScore)
''   If Comp$ = "B" Then
''      BScore = BScore + BlackScore '+ WhiteScore)
''   Else
''      WScore = WScore + WhiteScore '- BlackScore)
''   End If
'
'End Sub

