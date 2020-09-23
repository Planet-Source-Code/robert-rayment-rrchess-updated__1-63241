Attribute VB_Name = "ModBoard"
' ModBoard.bas  ~RRChess~

Option Explicit


Public Sub CountPieces(Index As Integer)
Dim R As Long, C As Long
' Rest Public
   NumWR = 0
   NumWN = 0
   NumWB = 0
   NumWQ = 0
   NumWP = 0
   NumWK = 0
   
   NumBR = 0
   NumBN = 0
   NumBB = 0
   NumBQ = 0
   NumBP = 0
   NumBK = 0
   
   PieceCount = 0
   BlackPieceCount = 0
   WhitePieceCount = 0
   
   For R = 1 To 8
   For C = 1 To 8
      Select Case bRCBoard(R, C, Index)
      Case 0
      Case WRn: NumWR = NumWR + 1: WhitePieceCount = WhitePieceCount + 1
      Case WNn: NumWN = NumWN + 1: WhitePieceCount = WhitePieceCount + 1
      Case WBn: NumWB = NumWB + 1: WhitePieceCount = WhitePieceCount + 1
      Case WQn: NumWQ = NumWQ + 1: WhitePieceCount = WhitePieceCount + 1
      Case WPn: NumWP = NumWP + 1: WhitePieceCount = WhitePieceCount + 1
      Case WKn: NumWK = NumWK + 1: WhitePieceCount = WhitePieceCount + 1
      
      Case BRn: NumBR = NumBR + 1: BlackPieceCount = BlackPieceCount + 1
      Case BNn: NumBN = NumBN + 1: BlackPieceCount = BlackPieceCount + 1
      Case BBn: NumBB = NumBB + 1: BlackPieceCount = BlackPieceCount + 1
      Case BQn: NumBQ = NumBQ + 1: BlackPieceCount = BlackPieceCount + 1
      Case BPn: NumBP = NumBP + 1: BlackPieceCount = BlackPieceCount + 1
      Case BKn: NumBK = NumBK + 1: BlackPieceCount = BlackPieceCount + 1
      End Select
   Next C
   Next R
   PieceCount = BlackPieceCount + WhitePieceCount
End Sub


Public Sub ShowSquare(M As Image, MoveString$)
' From Form1   IM_MouseDown & DoSetUp
' for showing start square in LabMvString
Dim i As Long, a$
Dim R As Long, C As Long
   i = M.Index
   ' Convert Drop index to a1 -> h8 notation
   R = 8 - i \ 8
   C = i - (64 - 8 * R) + 1
   a$ = Chr$(C + 96) & Trim$(Str$(R))
   MoveString$ = M.Tag & " " & a$
End Sub

Public Sub ClearBoard(frm As Form)
   ClearBoard_KeepList frm
   ' Clear list
   frm.LabMvString = ""
   frm.ListMoves.Clear
End Sub

Public Sub ClearBoard_KeepList(frm As Form)
Dim k As Long
   For k = 0 To 63
      With frm.IM(k)
         .Picture = LoadPicture
         .DragIcon = LoadPicture
         .Tag = ""
      End With
   Next k
End Sub

Public Sub SetUp_PlacePiece(Target As Image, Source As Image)
   With Target
      If .Tag = "" Then PieceCount = PieceCount + 1
      .Picture = Source.Picture
      .DragIcon = Source.DragIcon
      .Tag = Source.Tag
      .Visible = True
   End With
   SavePosition Form1, 0      ' bRCBoard(1-8, 1-8, 0)
End Sub

Public Sub SetUp_MovePiece(Target As Image, Source As Image)
   If Target = Source Then Exit Sub
   With Target
      If .Tag <> "" Then PieceCount = PieceCount - 1
      .Picture = Source.Picture
      .DragIcon = Source.DragIcon
      .Tag = Source.Tag
      .Visible = True
   End With
   Source.Picture = LoadPicture
   Source.DragIcon = LoadPicture
   Source.Tag = ""
   SavePosition Form1, 0      ' bRCBoard(1-8, 1-8, 0)
End Sub

Public Sub FullSetUp(frm As Form)
Dim R As Long, C As Long
   For R = 8 To 1 Step -1
   For C = 1 To 8
      bRCBoard(R, C, 0) = Opener(R, C)
   Next C
   Next R
   RestoreSavedPosition frm, 0
End Sub

Public Sub CheckerPic(frm As Form, P As PictureBox, Lab1 As Label, Lab2 As Label, Lab3 As Label)
' SQ = square size   ' Chess piece icons 32 x 32
' Lab0 & Lab1 = Black & White labels
Dim i As Long, j As Long
Dim Cul As Long
Dim LX As Long, LY As Long
Dim LH As Long
   LH = 66 'Lab1.Height
   ' Set board size
   P.Width = 8 * SQ '+ 2
   P.Height = 8 * SQ '+ 2
'   frm.Line (9, 9)-(p.Width + 44, p.Height + 62), &H80C0FF, BF
   ' Light square color
   Cul = &HFFFFFF
   ' Dark square color
   P.BackColor = &HA0A0A0
   For j = 0 To SQ * 6 Step 2 * SQ
   For i = 0 To SQ * 6 Step 2 * SQ
      P.Line (i, j)-(i + SQ, j + SQ), Cul, BF
      P.Line (i, j)-(i + SQ, j + SQ), 0, B
      P.Line (i + SQ, j + SQ)-(i + 2 * SQ, j + 2 * SQ), Cul, BF
      P.Line (i + SQ, j + SQ)-(i + 2 * SQ, j + 2 * SQ), 0, B
   Next i
   Next j
   P.Line (0, 0)-(P.Width - 3, P.Height - 3), 0, B
   P.Refresh
   ' Bottom labels
   Lab1.Top = P.Top + P.Height + 16 '18 ' LabCH(0)
   Lab2.Top = P.Top + P.Height + 16 '17 ' LabCul(0)
   Lab3.Top = Lab2.Top 'P.Top + P.Height + 17 ' LabMvIndicator(0)
   frm.Show
   
   ' Row 8-1/column a-h markers
   LX = P.Left - 12
   LY = P.Top + SQ \ 2 - 2
   
   'For i = 7 To 0 Step -1  ' White at top
   'frm.Cls
   frm.Line (9 + 3, 9 + 3)-(P.Width + 44, P.Height + LH), MainColor, BF '&H80C0FF, BF
   For i = 0 To 7  ' Black at top
      If aBlackAtTop Then
         j = i
      Else
         j = 7 - i
      End If
      frm.CurrentX = LX
      frm.CurrentY = LY
      frm.Print 8 - j;
      LY = LY + SQ
   Next i
   
   LX = P.Left + SQ \ 2 - 2
   LY = P.Top + P.Height + 2
   For i = 0 To 7
      If aBlackAtTop Then
         j = i
      Else
         j = 7 - i
      End If
      frm.CurrentX = LX
      frm.CurrentY = LY
      frm.Print Chr$(97 + j);
      LX = LX + SQ
   Next i
   frm.Line (10 + 3, 10 + 3)-(P.Width + 46, P.Height + LH + 3), 0, B
   frm.Line (9 + 3, 9 + 3)-(P.Width + 44, P.Height + LH + 1), vbWhite, B
End Sub

Public Sub LoadImageBoxes(frm As Form, IM As Image)    'Form1, picBoard, IM(0)
Dim ix As Long, iy As Long
Dim N As Long
   frm.IM(0).Width = 32
   frm.IM(0).Height = 32
   ' Center IM(0) at TL
   margin = (SQ - 32) \ 2
   IM.Move margin, margin
   IM.Visible = True
   ' Make 64 image boxes
   For ix = 1 To 63
      Load frm.IM(ix)
   Next ix
   
   ' Top row, row 8
   For ix = 1 To 7
      frm.IM(ix).Move margin + ix * SQ, margin
      frm.IM(ix).Visible = True
   Next ix
   ' Move remaining images to rows 7 to 1
   N = 8
   For iy = 1 To 7
      For ix = 0 To 7
         frm.IM(N).Move margin + ix * SQ, margin + iy * SQ
         frm.IM(N).Visible = True
         N = N + 1
      Next ix
   Next iy
   
   ' Taken pieces image boxes IMWSP(0-31)
'   Temp visibility
'   For n = 0 To 31
'      frm.IMWSP(n).BorderStyle = 1
'   Next n
   iy = frm.IMWSP(0).Top
   ix = frm.IMWSP(0).Left
   frm.IMWSP(1).Top = iy
   frm.IMWSP(1).Left = ix + 22
   For N = 2 To 14 Step 2
      frm.IMWSP(N).Top = iy + 22 * N \ 2
      frm.IMWSP(N).Left = ix
      frm.IMWSP(N + 1).Top = frm.IMWSP(N).Top
      frm.IMWSP(N + 1).Left = frm.IMWSP(N).Left + 22
   Next N

   iy = frm.IMWSP(16).Top
   ix = frm.IMWSP(16).Left
   frm.IMWSP(17).Top = iy
   frm.IMWSP(17).Left = ix + 22
   For N = 18 To 30 Step 2
      frm.IMWSP(N).Top = iy + 22 * (N - 16) \ 2
      frm.IMWSP(N).Left = ix
      frm.IMWSP(N + 1).Top = frm.IMWSP(N).Top
      frm.IMWSP(N + 1).Left = frm.IMWSP(N).Left + 22
   Next N

End Sub

Public Sub SavePosition(frm As Form, Index As Integer)
' Saves postion from main display board index
' in to bRCBoard(r, c, Index)
' 1,2,3, 4, 5, 6 White R,N,B,Q,K,P
' 7,8,9,10,11,12 Black R,N,B,Q,K,P
' Sets RC board (Index) from Saved board
' Returns the WK & BK row,column ie Public rwk,cwk & rbk,cbk
Dim N As Long
Dim R As Long, C As Long
Dim P As Long
   For P = 0 To 63
      Select Case Left$(frm.IM(P).Tag, 2)
      Case "WR": N = WRn   ' 1
      Case "WN": N = WNn   ' 2
      Case "WB": N = WBn   ' 3
      Case "WQ": N = WQn   ' 4
      Case "WK": N = WKn   ' 5
      Case "WP": N = WPn   ' 6
      Case "BR": N = BRn   ' 7
      Case "BN": N = BNn   ' 8
      Case "BB":
      N = BBn   ' 9
      Case "BQ": N = BQn   ' 10
      Case "BK": N = BKn   ' 11
      Case "BP": N = BPn   ' 12
      Case Else: N = 0
      End Select
      ' Convert 0-63 to r,c
      R = 8 - P \ 8
      C = P - (64 - 8 * R) + 1
      bRCBoard(R, C, Index) = N
      ' Log king positions
      If N = 5 Then  ' WK
         rwk = R
         cwk = C
      ElseIf N = 11 Then   'BK
         rbk = R
         cbk = C
      End If
   Next P
End Sub

Public Sub ClearTakenPieces(frm As Form)
Dim k As Long
   For k = 0 To 31
      frm.IMWSP(k).Picture = LoadPicture
   Next k
   WPTaken = 0
   BPTaken = 0
End Sub

Public Sub RestoreSavedPosition(frm As Form, Index As Integer)
' Restore display from bRCBoard(r, c, Index)
Dim R As Long, C As Long
Dim pr As Long, P As Long
Dim N As Long
   ClearBoard_KeepList frm
   For R = 1 To 8  ' for c = 1  p = 56 to 0 step -8
      pr = (8 - R) * 8
      For C = 1 To 8  '
         P = pr + C - 1
         N = bRCBoard(R, C, Index)
         If N <> 0 Then
            frm.IM(P).Tag = frm.IMO(N).Tag
            frm.IM(P).Picture = frm.IMO(N).Picture
            frm.IM(P).DragIcon = frm.IMO(N).DragIcon
         Else
            frm.IM(P).Tag = ""
            frm.IM(P).Picture = LoadPicture
            frm.IM(P).DragIcon = LoadPicture
         End If
         frm.IM(P).Visible = True
      Next C
   Next R
End Sub

Public Sub RestoreBeginBoard(frm As Form)
' BeginBoard to Display
Dim R As Long, C As Long
Dim pr As Long, P As Long
Dim N As Long
   For R = 1 To 8  ' for c = 1  p = 56 to 0 step -8
      pr = (8 - R) * 8
      For C = 1 To 8  '
         P = pr + C - 1
         N = BeginBoard(R, C)
         If N <> 0 Then
            frm.IM(P).Tag = frm.IMO(N).Tag
            frm.IM(P).Picture = frm.IMO(N).Picture
            frm.IM(P).DragIcon = frm.IMO(N).DragIcon
         Else
            frm.IM(P).Tag = ""
            frm.IM(P).Picture = LoadPicture
            frm.IM(P).DragIcon = LoadPicture
         End If
      Next C
   Next R
End Sub

Public Sub MakeFENString(WB$)
' WB$ "W" or "B" next move color
' Make FENString$ from Board 0
' Public FENString$
Dim R As Long, C As Long
Dim PN As Long
Dim P$
Dim ZeroC As Long
Dim aCounting As Boolean

   FENString$ = ""
   For R = 8 To 1 Step -1
   ZeroC = 0
   For C = 1 To 8
      PN = bRCBoard(R, C, 0)
      aCounting = False
      Select Case PN
      Case 0
         ZeroC = ZeroC + 1
         aCounting = True
      Case 1: P$ = "R"
      Case 2: P$ = "N"
      Case 3: P$ = "B"
      Case 4: P$ = "Q"
      Case 5: P$ = "K"
      Case 6: P$ = "P"
      
      Case 7: P$ = "r"
      Case 8: P$ = "n"
      Case 9: P$ = "b"
      Case 10: P$ = "q"
      Case 11: P$ = "k"
      Case 12: P$ = "p"
      End Select
      If Not aCounting Then
         If ZeroC <> 0 Then
            FENString$ = FENString$ & Trim$(Str$(ZeroC))
            ZeroC = 0
         End If
         FENString$ = FENString$ & P$
      End If
   Next C
   If ZeroC <> 0 Then
   FENString$ = FENString$ & Trim$(Str$(ZeroC))
   End If
   FENString$ = FENString$ & "/"
   Next R
   If WB$ = "W" Then
   FENString$ = FENString$ & " w - -"
   Else
   FENString$ = FENString$ & " b - -"
   End If
   FENString$ = FENString$ & Str$(HalfMove + 2)
   
End Sub

Public Sub ConvPNtoPNDescrip(PN As Long, P$)
   Select Case PN
   Case 0:   P$ = ""
   Case WRn: P$ = "WR"
   Case WNn: P$ = "WN"
   Case WBn: P$ = "WB"
   Case WQn: P$ = "WQ"
   Case WKn: P$ = "WK"
   Case WPn: P$ = "WP"
   
   Case BRn: P$ = "BR"
   Case BNn: P$ = "BN"
   Case BBn: P$ = "BB"
   Case BQn: P$ = "BQ"
   Case BKn: P$ = "BK"
   Case BPn: P$ = "BP"
   End Select
End Sub

Public Sub ConvPNDescriptoPN(P$, PN As Long)
   Select Case P$
   Case "":   PN = 0
   Case "WR": PN = WRn
   Case "WN": PN = WNn
   Case "WB": PN = WBn
   Case "WQ": PN = WQn
   Case "WK": PN = WKn
   Case "WP": PN = WPn
   
   Case "BR": PN = BRn
   Case "BN": PN = BNn
   Case "BB": PN = BBn
   Case "BQ": PN = BQn
   Case "BK": PN = BKn
   Case "BP": PN = BPn
   End Select
End Sub

Public Sub SetStartUp()
   ReDim svWKRR(0 To 3, 0 To MaxIndex + 3)
   ReDim svWEPP(0 To 3, 0 To MaxIndex + 3)
   ReDim svBKRR(0 To 3, 0 To MaxIndex + 3)
   ReDim svBEPP(0 To 3, 0 To MaxIndex + 3)
   ReDim svCast(0 To 3, 0 To MaxIndex + 3)
      
   ReDim rs(0 To MaxIndex), rd(0 To MaxIndex)
   ReDim cs(0 To MaxIndex), cd(0 To MaxIndex)
   ReDim PieceNum(0 To MaxIndex)
   ReDim PieceIndex$(0 To MaxIndex)
   ReDim aMovOK(0 To MaxIndex)
   ReDim M(0 To MaxIndex), Q(0 To MaxIndex)
   ReDim svM(0 To MaxIndex), svQ(0 To MaxIndex)
   ReDim DOFLAG(0 To MaxIndex)
   ReDim svDOFLAG(0 To MaxIndex)
   
   ReDim aPawnProm(0 To MaxIndex)
   ReDim aEnPassant(0 To MaxIndex)
   ReDim aCastling(0 To MaxIndex)
'   ' Pawns' log
'   ' 0 not moved,
'   ' 2 moved 2 sqs last(possible en passant),
'   ' 3 @ far row(promotion)
   ReDim WPawn(1 To 8, 0 To MaxIndex) ' (1 to 8)
   ReDim BPawn(1 To 8, 0 To MaxIndex) ' (1 to 8)
   WPTaken = 0
   BPTaken = 0
   
   ReDim StartBoard(1 To 8, 1 To 8) As Byte
   ReDim CountBoard(1 To 8, 1 To 8, 1 To 2) As Byte
   ReDim EndBoard(1 To 8, 1 To 8) As Byte
   ReDim RepPositions(1 To 8, 1 To 8, 1 To 12)
   ReDim TestCheckMateBoard(1 To 8, 1 To 8) As Byte
   
   aCheckmate = False
   
   ReDim PieceVal(1 To 12)
   PieceVal(1) = 50   ' WRook ' Used as PieceVal(WRn)=50 or PieceVal(AttackingPiece)=50
   PieceVal(2) = 30   ' WKnight
   PieceVal(3) = 35   ' WBishop
   PieceVal(4) = 100  ' WQueen
   PieceVal(5) = 1000  ' WKing
   PieceVal(6) = 10   ' WPawn

   PieceVal(7) = 50   ' BRook ' Used as PieceVal(BRn)=50 or PieceVal(AttackingPiece)=50
   PieceVal(8) = 30   ' BKnight
   PieceVal(9) = 35   ' BBishop
   PieceVal(10) = 100 ' BQueen
   PieceVal(11) = 1000 ' BKing
   PieceVal(12) = 10  ' BPawn
End Sub

Public Sub InitArrays()
' Public MaxIndex set at Form_Initialize
   ReDim NumMoves(0 To MaxIndex)
   ReDim NumMates(0 To MaxIndex)
   ReDim DOFLAG(0 To MaxIndex)

   ' For piece position analysis
   ReDim WRrow(1 To 3), WRcol(1 To 3)
   ReDim WNrow(1 To 3), WNcol(1 To 3)
   ReDim WBrow(1 To 3), WBcol(1 To 3)
   ReDim WQrow(1 To 3), WQcol(1 To 3)
   'WKrow , WKcol
   ReDim WProw(1 To 8), WPcol(1 To 8)
   
   ReDim BRrow(1 To 3), BRcol(1 To 3)
   ReDim BNrow(1 To 3), BNcol(1 To 3)
   ReDim BBrow(1 To 3), BBcol(1 To 3)
   ReDim BQrow(1 To 3), BQcol(1 To 3)
   'BKrow , BKcol
   ReDim BProw(1 To 8), BPcol(1 To 8)
   
   ReDim WMoveStore(1 To 8, 1 To 8, 100)  ' Max moves ~84 ?
   ReDim BMoveStore(1 To 8, 1 To 8, 100)

      ' Sets ALL to 0 = False
   ReDim WKRR(0 To 3, 0 To MaxIndex)
   ReDim WEPP(0 To 3, 0 To MaxIndex)
   ReDim BKRR(0 To 3, 0 To MaxIndex)
   ReDim BEPP(0 To 3, 0 To MaxIndex)
   ReDim Cast(0 To 3, 0 To MaxIndex)

   ReDim bRCBoard(1 To 8, 1 To 8, 0 To MaxIndex)   ' Start Index = 2 but depends
   ReDim TempBoard(1 To 8, 1 To 8, 1 To MaxIndex) As Byte
   ReDim TestAttackBoard(1 To 8, 1 To 8)


End Sub

'Public WKRR() As Byte   ' ReDim WKRR(0 to 3, 0 To MaxIndex)
'Public BKRR() As Byte   ' ReDim BKRR(0 to 3, 0 To MaxIndex)
'Public WEPP() As Byte   ' ReDim WEPP(0 to 3, 0 To MaxIndex)
'Public BEPP() As Byte   ' ReDim BEPP(0 to 3, 0 To MaxIndex)
'Public Cast() As Byte   ' ReDim Cast(0 To 3, 0 To MaxIndex)

Public Sub CopyALLBOOLS(Index1 As Integer, Index2 As Integer)  '-->>
   ' Copy all Bools from Index1 to Index2
   CopyMemory WKRR(0, Index2), WKRR(0, Index1), 4  ' <<--
   CopyMemory BKRR(0, Index2), BKRR(0, Index1), 4
   CopyMemory WEPP(0, Index2), WEPP(0, Index1), 4
   CopyMemory BEPP(0, Index2), BEPP(0, Index1), 4
   CopyMemory Cast(0, Index2), Cast(0, Index1), 4
   'aCastling(Index2) = aCastling(Index1)
End Sub

Public Sub SaveALLBOOLS(BoardIndex As Integer, svIndex As Integer)   ' -->>
   '                    Src                    Dest
   '          Dest                Src
   CopyMemory svWKRR(0, svIndex), WKRR(0, BoardIndex), 4
   CopyMemory svWEPP(0, svIndex), WEPP(0, BoardIndex), 4
   CopyMemory svBKRR(0, svIndex), BKRR(0, BoardIndex), 4
   CopyMemory svBEPP(0, svIndex), BEPP(0, BoardIndex), 4
   CopyMemory svCast(0, svIndex), Cast(0, BoardIndex), 4
   'aCastling(Index2) = aCastling(Index1)
End Sub

Public Sub RestoreALLBOOLS(svIndex As Integer, BoardIndex As Integer)   ' -->>
   '                       Src                 Dest
   '          Dest                 Src
   CopyMemory WKRR(0, BoardIndex), svWKRR(0, svIndex), 4
   CopyMemory WEPP(0, BoardIndex), svWEPP(0, svIndex), 4
   CopyMemory BKRR(0, BoardIndex), svBKRR(0, svIndex), 4
   CopyMemory BEPP(0, BoardIndex), svBEPP(0, svIndex), 4
   CopyMemory Cast(0, BoardIndex), svCast(0, svIndex), 4
End Sub

