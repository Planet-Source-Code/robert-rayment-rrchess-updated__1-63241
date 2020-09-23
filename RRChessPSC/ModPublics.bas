Attribute VB_Name = "ModPublics"
' ModPublics.bas   ~ RRChess ~

Option Explicit

' Several
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
 (Destination As Any, Source As Any, ByVal Length As Long)

' For delaying show game & show working boards
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

' For detecting mouse button
Public Declare Function GetAsyncKeyState Lib "user32" _
   (ByVal vKey As KeyCodeConstants) As Long

Public Enum EPT   ' PieceType
   WRn = 1
   WNn = 2
   WBn = 3
   WQn = 4
   WKn = 5
   WPn = 6
   BRn = 7
   BNn = 8
   BBn = 9
   BQn = 10
   BKn = 11
   BPn = 12
End Enum

' White  Black
'  R 1    R 7
'  N 2    N 8
'  B 3    B 9
'  Q 4    Q 10
'  K 5    K 11
'  P 6    P 12

Public PieceVal() As Integer

Public Enum EWKRR    ' WKRR([0][1][2][3], 0 To MaxIndex)
  WKMoved = 0  ' 0   White King moved
  WKRMoved = 1 ' 1   White KRook moved
  WQRMoved = 2 ' 2   White QRook moved
  WSPARE = 3   ' 3
End Enum

Public Enum EBKRR   ' BKRR()
  BKMoved = 0  ' 0   Black King moved
  BKRMoved = 1 ' 1   Black KRook moved
  BQRMoved = 2 ' 2   Black QRook moved
  BSPARE = 3   ' 3
End Enum

Public Enum EWEPP    ' WEPP()
   WEPOK = 0    ' 0
   WEPSet = 1     ' 1 (0 or Column number)
   WEPProm = 2    ' 2
   WEPSPARE = 3   ' 3
End Enum

Public Enum EBEPP    ' BEPP()
   BEPOK = 0    ' 0
   BEPSet = 1     ' 1 (0 or Column number)
   BEPProm = 2    ' 2
   BEPSPARE = 3   ' 3
End Enum

Public Enum ECast    ' Cast()
   WKSCastOK = 0   ' 0
   WQSCastOK = 1   ' 1
   BKSCastOK = 2   ' 2
   BQSCastOK = 3   ' 3
End Enum

'() (Index) 0 = FALSE, 1 = TRUE      Byte    Board Index
Public WKRR() As Byte   ' ReDim WKRR(0 to 3, 0 To MaxIndex)
Public BKRR() As Byte   ' ReDim BKRR(0 to 3, 0 To MaxIndex)
Public WEPP() As Byte   ' ReDim WEPP(0 to 3, 0 To MaxIndex)
Public BEPP() As Byte   ' ReDim BEOO(0 to 3, 0 To MaxIndex)
Public Cast() As Byte   ' ReDim Cast(0 To 3, 0 To MaxIndex)

' WKRR(0 To 3, Index)
   ' WKMoved   ' 0   White King moved   WKRR(0, I) = 0/1 ie WKRR(WKMoved, I) = 0/1
   ' WKRMoved  ' 1   White KRook moved
   ' WQRMoved  ' 2   White QRook moved
   ' WSPARE    ' 3
' BKRR(0 To 3, Index)
   ' BKMoved   ' 0   Black King moved
   ' BKRMoved  ' 1   Black KRook moved
   ' BQRMoved  ' 2   Black QRook moved
   ' BSPARE    ' 3
' WEPP(0 To 3, Index)
   ' WEPOK     ' 0 if 1 EP can be done
   ' WEPSet    ' 1 (0 or Column number)
   ' WEPProm   ' 2 if 1 PawnProm can be done
   ' WEPSPARE  ' 3
' BEPP(0 To 3, Index)
   ' BEPOK     ' 0 if 1 EP can be done
   ' BEPSet    ' 1 (0 or Column number)
   ' BEPProm   ' 2 if 1 PawnProm can be done
   ' BEPSPARE  ' 3
' Cast(0 To 3, Index)
   ' WKSCastOK    ' 0 if 1 King-side can be done
   ' WQSCastOK    ' 1 if 1 Queen-side can be done
   ' BKSCastOK    ' 2 ditto black
   ' BQSCastOK    ' 3

' BPawnPromotion() As Boolean   ' 3   Black pawn promotion
' BPieceTaken() As Boolean      ' 4
'                               ' 5,6,7 SPARE
'' Castling check
' Castling() As Boolean         ' 0
' KSideCastling() As Boolean    ' 1
' QSideCastling() As Boolean    ' 2
' SPARE                                 ' 3

' Pawns' log
' 0 not moved,
' 2 moved 2 sqs last(possible en passant),
Public WPawn() As Byte  ' ReDim WPawn(1 to 8, 0 to 3)
Public BPawn() As Byte  ' ReDim BPawn(1 to 8, 0 to 3)

Public svWKRR() As Byte    ' ReDim svWKRR(0 to 3, 0 To MaxIndex + 3)
Public svWEPP() As Byte    ' ReDim svWKRR(0 to 3, 0 To MaxIndex + 3)
Public svBKRR() As Byte    ' ReDim svWKRR(0 to 3, 0 To MaxIndex + 3)
Public svBEPP() As Byte    ' ReDim svWKRR(0 to 3, 0 To MaxIndex + 3)
Public svCast() As Byte    ' ReDim svWKRR(0 to 3, 0 To MaxIndex + 3)
''''''''''''''''''''''''''''''''''''''''''
' For piece position analysis
' Number of W/Bpieces on board
Public NumWR As Long
Public NumWN As Long
Public NumWB As Long
Public NumWQ As Long
Public NumWP As Long
Public NumWK As Long

Public NumBR As Long
Public NumBN As Long
Public NumBB As Long
Public NumBQ As Long
Public NumBP As Long
Public NumBK As Long
   
Public WRrow() As Long, WRcol() As Long
Public WNrow() As Long, WNcol() As Long
Public WBrow() As Long, WBcol() As Long
Public WQrow() As Long, WQcol() As Long
Public WProw() As Long, WPcol() As Long
Public WKrow As Long, WKcol As Long

Public BRrow() As Long, BRcol() As Long
Public BNrow() As Long, BNcol() As Long
Public BBrow() As Long, BBcol() As Long
Public BQrow() As Long, BQcol() As Long
Public BProw() As Long, BPcol() As Long
Public BKrow As Long, BKcol As Long

Public NumLegMovesWR As Long
Public NumLegMovesWN As Long
Public NumLegMovesWB As Long
Public NumLegMovesWQ As Long
Public NumLegMovesWK As Long
Public NumLegMovesWP As Long

Public NumLegMovesBR As Long
Public NumLegMovesBN As Long
Public NumLegMovesBB As Long
Public NumLegMovesBQ As Long
Public NumLegMovesBK As Long
Public NumLegMovesBP As Long

' For storing legal moves
Public WMoveStore() As Byte
Public BMoveStore() As Byte

''''''''''''''''''''''''''''''''''''''''''
Public ippt As Long   ' piece num promoted or taken

Public MoveString$                 ' eg WQ a1-h8 always 8 characters
Public bRCBoard() As Byte          ' (1-8 row, 1-8 column, Index) Numeric boards
Public WScore As Long
Public BScore As Long

' Boards
Public BeginBoard() As Byte
Public CountBoard() As Byte
Public StartBoard() As Byte
Public TempBoard() As Byte
Public EndBoard() As Byte
Public TestAttackBoard() As Byte
Public TestCheckMateBoard() As Byte
Public RepPositions() As Byte
Public NumRepPositions As Long

' IM(IMIndex) Image holding piece icons
Public IMLeft As Long, IMTop As Long
Public IMIndex As Integer
Public margin As Long
Public PieceCount As Long
Public BlackPieceCount As Long
Public WhitePieceCount As Long
Public aLegalMove As Boolean
Public aDraggedOut As Boolean ' Indicating piece dragged off board

Public aBlackAtTop As Boolean ' Flip board flag

Public WorBsMove$       ' "W" White's, "B" Black's move
Public CorHsMove$       ' "C" Computer's, "H" Human's move
Public FirstColor$

' IMO(IMOIndex) Images holding off board piece set (Index 1 to 12)
Public IMOIndex As Integer

' IMWSP(0-31) Taken pieces image boxes
Public WPTaken As Long
Public BPTaken As Long

Public aSetUp As Boolean   ' Set up or Play
Public aHold As Boolean    ' To prevent optSelect

Public aMouseDown As Boolean

Public HalfMove As Long       ' For stepping through loaded game
Public GameOffset As Long
Public MaxTime As Long

Public STX As Long, STY As Long ' TwipsPerPixelX/Y

' Publics used in ModMate
Public rs() As Long, rd() As Long
Public cs() As Long, cd() As Long
Public UDSearch As Long ' 0 row 1-8, 1 row 8-1 search direction
Public PieceNum() As Long          ' Piece numbers
Public aMovOK() As Boolean
Public M() As Long, Q() As Long
Public svM() As Long, svQ() As Long
Public NumMoves() As Long           ' Number of opposite players moves
Public NumMates() As Long           ' Number of mates
Public DOFLAG() As Long
Public svDOFLAG() As Long
Public aPawnProm() As Boolean
Public aEnPassant()  As Boolean
Public aCastling()  As Long   ' 0 not castled, 1 W castled, 2 black castled

' Public for returning to calling Sub
Public Piece$
Public PieceIndex$()
Public svRowS As Long, svColS As Long  ' Row ,Col Start
Public svRowE As Long, svColE As Long  ' Row, Col End
Public ThePromPiece$
Public PieceColor$
Public aWKingInCheck As Boolean
Public aBKingInCheck As Boolean
Public DestPieceNum As Long

' Checker
Public Message$      ' Later make a number into Mess$()
Public Mess$() ' Message$

Public RowS As Long, ColS As Long        ' Row,Col start
Public RowE As Long, ColE As Long        ' Row,Col end
Public PlayColor$, OppColor$             ' Player color, Opponent color
Public MaxIndex As Integer

Public rwk As Long, cwk As Long  ' r,c of white king
Public rbk As Long, cbk As Long  ' r,c of black king
Public AttackingPiece As Long    ' Attacking Piece number
Public rat As Long, cat As Long  ' r,c of AttackingPiece$
Public NumAttacking As Long
Public AttackType() As Long
Public AttackEP() As Boolean
Public AttackTypeRow() As Long
Public AttackTypeCol() As Long

Public PromPiece$
Public aCheckmate As Boolean

Public OpenString$

' Files/Paths
Public PathSpec$, FileSpec$
Public LoadFENSpec$, SaveFENSpec$
Public LoadCHGSpec$, SaveCHGSpec$

Public FENString$

Public aBusy As Boolean
Public aExit As Boolean
Public PLY As Long

Public MainColor As Long, LightColor As Long


' TEST
Public MoveTot As Long
Public Solution$

Public Const SQ = 42 ' Square size   ' Chess piece icons 32 x 32

' NOT USED: POSSIBLE MESSAGE LISTS
'Public Sub SetMessages()
'   ReDim Mess$(20)
'   Mess$(1) = "White King put in Check"
'   Mess$(2) = "Black King put in Check"
'   'etc
'End Sub
