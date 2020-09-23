Attribute VB_Name = "ModFiles"
' ModFiles.bas  ~RRChess~

' For reading/saving RRChess setup files

Option Explicit

Private Title$, Filt$, InDir$
Private fnum As Long ' File number

Dim CommonDialog1 As OSDialog

Public Function Open_FEN_File(frm As Form, Entry As Integer) As Boolean
Dim k As Long
Dim a$, C$, WB$
Dim FenLine$

   On Error GoTo FENfileError
   Open_FEN_File = False
   
   If Entry <> 0 Then GoTo ReLoader
   
   If CommonDialog1 Is Nothing Then Set CommonDialog1 = New OSDialog
   
   Title$ = "Open setup FEN file"
   Filt$ = "Pics fen|*.fen"
   InDir$ = LoadFENSpec$
   FileSpec$ = ""
   
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", frm.hwnd
   
   If Len(FileSpec$) = 0 Then
      Close
      Set CommonDialog1 = Nothing
      Exit Function
   End If
   
   LoadFENSpec$ = FileSpec$
   Set CommonDialog1 = Nothing

ReLoader:
   
   If Len(LoadFENSpec$) = 0 Then Exit Function
   Solution$ = ""
   FenLine$ = ""
   GameOffset = 0
   fnum = FreeFile
   Open LoadFENSpec$ For Input As fnum
   Do
      Line Input #fnum, C$ ' eg   BR a1 Decsrip(any char)Location
      C$ = Trim$(C$)
      a$ = Left$(C$, 1)
      ' Ignore Comments apart from 'Solution
      If a$ <> "" And a$ <> "'" And a$ <> "{" And a$ <> ":" And a$ <> "[" Then
         FenLine$ = C$ ' Lines starting with ' ignored
      Else
         a$ = UCase$(Left$(C$, 9))
         If Mid$(a$, 2) = "SOLUTION" Then
            k = InStr(1, C$, " ")
            If k <> 0 Then
               Solution$ = Mid$(C$, k + 1)
            End If
         End If
      End If
   Loop Until EOF(fnum)
   Close #fnum
   DisplayFEN frm, FenLine$, WB$
   Open_FEN_File = True
   Exit Function
'==========
FENfileError:
Close fnum
MsgBox "FEN Input file error or no setup file to reload", vbCritical, "Loading"
End Function

Public Sub DisplayFEN(frm As Form, FenLine$, WB$)
Dim k As Long
Dim C$
Dim StartLineIndex  As Long
Dim PIndex As Long, PIncr As Long
Dim Dash As Boolean
   Dash = False
   ClearBoard Form1
   PieceCount = 0
   If aBlackAtTop = False Then
      Form1.Flip
   End If
   StartLineIndex = 0
   PIndex = 0
   For k = 1 To Len(FenLine$)
      C$ = Mid$(FenLine$, k, 1)
      If C$ = "/" Then
         StartLineIndex = StartLineIndex + 8
         PIndex = StartLineIndex
      ElseIf C$ = "w" And Dash Then
         WB$ = "W"
      ElseIf C$ = "b" And Dash Then
         WB$ = "B"
      ElseIf C$ = " " Then
         Dash = True
      ElseIf C$ = "-" Then
         If k + 4 <= Len(FenLine$) Then
            C$ = Mid$(FenLine$, k + 4)
            C$ = Trim$(C$)
            GameOffset = Val(C$)
            Exit For
         Else
            GameOffset = 0
         End If
      ElseIf Not Dash Then
         Piece$ = ""
         PIncr = 1
         Select Case C$
         Case Is < "9": PIncr = Val(C$)
         Case "R": IMOIndex = 1: Piece$ = "WR"
         Case "N": IMOIndex = 2: Piece$ = "WN"
         Case "B": IMOIndex = 3: Piece$ = "WB"
         Case "Q": IMOIndex = 4: Piece$ = "WQ"
         Case "K": IMOIndex = 5: Piece$ = "WK"
         Case "P": IMOIndex = 6: Piece$ = "WP"
         
         Case "r": IMOIndex = 7: Piece$ = "BR"
         Case "n": IMOIndex = 8: Piece$ = "BN"
         Case "b": IMOIndex = 9: Piece$ = "BB"
         Case "q": IMOIndex = 10: Piece$ = "BQ"
         Case "k": IMOIndex = 11: Piece$ = "BK"
         Case "p": IMOIndex = 12: Piece$ = "BP"
         End Select
         If Piece$ <> "" Then
            frm.IM(PIndex).Picture = frm.IMO(IMOIndex).Picture
            frm.IM(PIndex).DragIcon = frm.IMO(IMOIndex).DragIcon
            frm.IM(PIndex).Tag = Piece$
            frm.IM(PIndex).Visible = True
            PieceCount = PieceCount + 1
            If Left$(Piece$, 1) = "W" Then
               WhitePieceCount = WhitePieceCount + 1
            Else
               BlackPieceCount = BlackPieceCount + 1
            End If
         End If
         PIndex = PIndex + PIncr
      End If
   Next k
End Sub

Public Function Save_FEN_File(frm As Form) As Boolean
' EG  WK e1
Dim k As Long
Dim a$
Dim PN As Long ' PieceNum
Dim Row As Long, Col As Long
   Save_FEN_File = False
   
   If CommonDialog1 Is Nothing Then Set CommonDialog1 = New OSDialog
   
   Title$ = "Save board as a FEN file"
   Filt$ = "Pics FEN|*.fen"
   InDir$ = SaveFENSpec$
   FileSpec$ = ""
   
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", frm.hwnd
   
   If Len(FileSpec$) = 0 Then
      Close
      Set CommonDialog1 = Nothing
      Exit Function
   End If
   
   SaveFENSpec$ = FileSpec$
   Set CommonDialog1 = Nothing
   
   FixExtension SaveFENSpec$, "fen"
   
   fnum = FreeFile
   
   a$ = InputBox("Enter 2nd line comment", "Saving fen")
   
   Open SaveFENSpec$ For Output As fnum
   Print #fnum, "'" & FindName$(SaveFENSpec$) & "  " & Date$ & " " & Time$
   If a$ <> "" Then Print #fnum, "'" & a$
   
   SavePosition frm, 0
   ' Saves postion from main display board images
   ' in to bRCBoard(r, c, 0)
   ' 1,2,3, 4, 5, 6 White R,N,B,Q,K,P
   ' 7,8,9,10,11,12 Black r,n,b,q,k,p
   a$ = ""
   For Row = 8 To 1 Step -1
      k = 0
      For Col = 1 To 8
         PN = bRCBoard(Row, Col, 0)
         If PN = 0 Then
            k = k + 1
         Else
            If k > 0 Then
               a$ = a$ & Trim$(Str$(k))
               k = 0
            End If
            Select Case PN
            Case 1: a$ = a$ & "R"
            Case 2: a$ = a$ & "N"
            Case 3: a$ = a$ & "B"
            Case 4: a$ = a$ & "Q"
            Case 5: a$ = a$ & "K"
            Case 6: a$ = a$ & "P"
            
            Case 7: a$ = a$ & "r"
            Case 8: a$ = a$ & "n"
            Case 9: a$ = a$ & "b"
            Case 10: a$ = a$ & "q"
            Case 11: a$ = a$ & "k"
            Case 12: a$ = a$ & "p"
            End Select
         End If
      Next Col
      If k > 0 Then
         a$ = a$ & Trim$(Str$(k))
         k = 0
      End If
      If Row > 1 Then a$ = a$ & "/"
   Next Row
   Print #fnum, a$ & " w - - 0"
   Close #fnum
   Save_FEN_File = True
End Function


Public Function Save_chg_File(frm As Form, GameList As ListBox, WB$)
' WB$ NOT USED SO FAR
' Save game from ListMoves ListBox

' CHS format bulky but easy to read
' also allows 'silly' games like black moves only :)
' NB UC for piece descrip, LC for column ident
'    Move number odd numbers white even number black,
'    unless position starting with black or a 'silly game'.
'    Strict space location ie after move number & after
'    piece descrip

' eg
' 1 WN b1-a3
' 2 BN b8-h6
' 3 WP e4xd5       ' captures
' 4 WP e7-e8=WQ    ' promotion
'    En passant just diagonal pawm move to blank square
'    O-O & O-O-O not used just K moves 2 squares L/R rows 1/8

' COULD DO:_
' OR
' 1. WNb1-a3 BNb8-h6
' 2. WPe2-e4 BPe7-e5
' 3. WPe7-e8=Q BQe9xe8
' OR
' 1 Nb1-a3 Nb8-h6
' 2 e2-e4 e7-e5
' 3 e7-e8=Q Qe9xe8
' OR  PGN
' 1. Nb1-a3 b8-h6 2. e2-e4 e7-e5 3. e7-e8=Q Qe9xe8
''

Dim a$
Dim k As Long
   Save_chg_File = False
   
   If GameList.ListCount > 0 Then
   
      If CommonDialog1 Is Nothing Then Set CommonDialog1 = New OSDialog
      
      Title$ = "Save game as chg file"
      Filt$ = "Pics chg|*.chg"
      InDir$ = SaveCHGSpec$
      FileSpec$ = ""
      
      CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", frm.hwnd
      
      If Len(FileSpec$) = 0 Then
         Close
         Set CommonDialog1 = Nothing
         Exit Function
      End If
      
      SaveCHGSpec$ = FileSpec$
      Set CommonDialog1 = Nothing
      
      FixExtension SaveCHGSpec$, "chg"
      
      fnum = FreeFile
      Open SaveCHGSpec$ For Output As fnum
      Print #fnum, "'" & FindName$(SaveCHGSpec$) & "  " & Date$ & " " & Time$
      Print #fnum, FENString$
      For k = 0 To GameList.ListCount - 1
         a$ = GameList.List(k)
         Print #fnum, a$
      Next k
      Close fnum
      Save_chg_File = True
   End If   ' If GameList.ListCount > 0 Then
End Function

Public Function Open_chg_File(frm As Form, GameList As ListBox, Entry As Integer)
' Load a game chg format only!
' Reading PGN complex see PSC CodeIDs 51274 & 53306
Dim a$
Dim WB$
Dim aFEN As Boolean
Dim N As Long
Dim P As Long
Dim Num$
   On Error GoTo LoadGameError
   Open_chg_File = False
   
   If Entry <> 0 Then GoTo ReLoader
   
   If CommonDialog1 Is Nothing Then Set CommonDialog1 = New OSDialog
   
   Title$ = "Load game from chg file"
   Filt$ = "Pics chg|*.chg"
   InDir$ = LoadCHGSpec$
   FileSpec$ = ""
   
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", frm.hwnd
   
   If Len(FileSpec$) = 0 Then
      Close
      Set CommonDialog1 = Nothing
      Exit Function
   End If
   
   LoadCHGSpec$ = FileSpec$
   Set CommonDialog1 = Nothing
   
ReLoader:
   
   If Len(LoadCHGSpec$) = 0 Then Exit Function
   
   PieceCount = 0
   GameOffset = 0
   fnum = FreeFile
   Open LoadCHGSpec$ For Input As fnum
   PieceCount = 0
   
   aFEN = False
   
   GameList.Clear
   HalfMove = -1 ' Default
   GameOffset = 0
   
   ' Fill Listbox
   Do
      Line Input #fnum, a$ ' eg   BR a1 Decsrip(any char)Location
      a$ = Trim$(a$)
      If Left$(a$, 1) <> "'" And Left$(a$, 1) <> "" Then   ' Lines starting with ' ignored
         If InStr(1, a$, "/") Then ' a$=FenLine$
            DisplayFEN frm, a$, WB$
            GameOffset = 0
            aFEN = True
         Else
            P = InStr(1, a$, " ")
            If P <> 0 Then
               N = Val(Left$(a$, P - 1)) + GameOffset
               Num$ = Str$(N)
               If N > 9 Then Num$ = Trim$(Num$)
               a$ = Num$ & Mid$(a$, P)
            End If
            GameList.AddItem a$
         End If
      End If
   Loop Until EOF(fnum)
   Close fnum
   If Not aFEN Then
      FullSetUp frm
   End If
   Open_chg_File = True
   Exit Function
'===========
LoadGameError:
Close #fnum
MsgBox "Input file error or no game file to reload", vbCritical, "Loading"
End Function

Public Function GetAStartUp() As Boolean
Dim fnum As Long
   On Error GoTo RRChessInfo
   ' Defaults
   GetAStartUp = True
   
   If FileExists(PathSpec$ & "RRChessInfo.txt") Then
      fnum = FreeFile
      Open PathSpec$ & "RRChessInfo.txt" For Input As #fnum
      Input #fnum, aSound
      Close #fnum
   End If
   On Error GoTo 0
Exit Function
'========
RRChessInfo:
Close #fnum
End Function

Public Function FileExists(FSpec) As Boolean
  On Error Resume Next
  FileExists = FileLen(FSpec)
End Function

Public Sub FixExtension(FSpec$, Ext$)
' Enter Ext$ as jpg, bmp etc  no dot
Dim P As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   P = FindLastCharPos(FSpec$, ".")
   If P = 0 Then
      FSpec$ = FSpec$ & "." & Ext$
   Else
      If LCase$(Mid$(FSpec$, P + 1)) <> Ext$ Then FSpec$ = Mid$(FSpec$, 1, P) & Ext$
   End If
End Sub

Public Function FindPath$(FSpec$)
Dim P As Long
   FindPath$ = ""
   If Len(FSpec$) = 0 Then Exit Function
   P = FindLastCharPos(FSpec$, "\")
   If P = 0 Then Exit Function
   FindPath$ = Left$(FSpec$, P)
End Function

Public Function FindName$(FSpec$)
Dim P As Long
   FindName$ = ""
   If Len(FSpec$) = 0 Then Exit Function
   P = FindLastCharPos(FSpec$, "\")
   If P = 0 Then Exit Function
   FindName$ = Mid$(FSpec$, P + 1)
End Function

Public Function FindExtension$(FSpec$)
Dim P As Long
   P = FindLastCharPos(FSpec$, ".")
   If P = 0 Then
      FindExtension$ = ""
   Else
      FindExtension$ = Mid$(FSpec$, P + 1)
   End If
End Function

Public Function FindLastCharPos(InString$, SerChar$) As Long
' Also VB5
Dim P As Long
    For P = Len(InString$) To 1 Step -1
      If Mid$(InString$, P, 1) = SerChar$ Then Exit For
    Next P
    If P < 1 Then P = 0
    FindLastCharPos = P
End Function


