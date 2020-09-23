VERSION 5.00
Begin VB.Form frmPieces 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmPieces"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2760
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPieces.frx":0000
   ScaleHeight     =   4155
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3960
      Left            =   870
      TabIndex        =   1
      Top             =   60
      Width           =   1500
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":3042
         Height          =   540
         Index           =   12
         Left            =   780
         Picture         =   "frmPieces.frx":334C
         Stretch         =   -1  'True
         Tag             =   "BP"
         Top             =   3195
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":3656
         Height          =   540
         Index           =   11
         Left            =   735
         Picture         =   "frmPieces.frx":3960
         Stretch         =   -1  'True
         Tag             =   "BK"
         Top             =   2505
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":3C6A
         Height          =   540
         Index           =   10
         Left            =   750
         Picture         =   "frmPieces.frx":3F74
         Stretch         =   -1  'True
         Tag             =   "BQ"
         Top             =   1935
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":427E
         Height          =   540
         Index           =   9
         Left            =   765
         Picture         =   "frmPieces.frx":4588
         Stretch         =   -1  'True
         Tag             =   "BB"
         Top             =   1380
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":4892
         Height          =   540
         Index           =   8
         Left            =   750
         Picture         =   "frmPieces.frx":4B9C
         Stretch         =   -1  'True
         Tag             =   "BN"
         Top             =   810
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":4EA6
         Height          =   540
         Index           =   7
         Left            =   720
         Picture         =   "frmPieces.frx":51B0
         Stretch         =   -1  'True
         Tag             =   "BR"
         Top             =   240
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":54BA
         Height          =   540
         Index           =   6
         Left            =   195
         Picture         =   "frmPieces.frx":57C4
         Stretch         =   -1  'True
         Tag             =   "WP"
         Top             =   3195
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":5ACE
         Height          =   540
         Index           =   5
         Left            =   165
         Picture         =   "frmPieces.frx":5DD8
         Stretch         =   -1  'True
         Tag             =   "WK"
         Top             =   2520
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":60E2
         Height          =   540
         Index           =   4
         Left            =   150
         Picture         =   "frmPieces.frx":63EC
         Stretch         =   -1  'True
         Tag             =   "WQ"
         Top             =   1935
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":66F6
         Height          =   540
         Index           =   3
         Left            =   165
         Picture         =   "frmPieces.frx":6A00
         Stretch         =   -1  'True
         Tag             =   "WB"
         Top             =   1365
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":6D0A
         Height          =   540
         Index           =   2
         Left            =   120
         Picture         =   "frmPieces.frx":7014
         Stretch         =   -1  'True
         Tag             =   "WN"
         Top             =   810
         Width           =   540
      End
      Begin VB.Image IMO 
         DragIcon        =   "frmPieces.frx":731E
         Height          =   540
         Index           =   1
         Left            =   120
         Picture         =   "frmPieces.frx":7628
         Stretch         =   -1  'True
         Tag             =   "WR"
         Top             =   225
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   735
      Begin VB.Image IMW 
         Height          =   300
         Index           =   1
         Left            =   45
         Picture         =   "frmPieces.frx":7932
         Stretch         =   -1  'True
         Top             =   135
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   2
         Left            =   45
         Picture         =   "frmPieces.frx":7A24
         Stretch         =   -1  'True
         Top             =   480
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   3
         Left            =   15
         Picture         =   "frmPieces.frx":7B16
         Stretch         =   -1  'True
         Top             =   810
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   4
         Left            =   45
         Picture         =   "frmPieces.frx":7C08
         Stretch         =   -1  'True
         Top             =   1155
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   5
         Left            =   45
         Picture         =   "frmPieces.frx":7CFA
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   6
         Left            =   45
         Picture         =   "frmPieces.frx":7DEC
         Stretch         =   -1  'True
         Top             =   1815
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   7
         Left            =   375
         Picture         =   "frmPieces.frx":7EDE
         Stretch         =   -1  'True
         Top             =   120
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   8
         Left            =   375
         Picture         =   "frmPieces.frx":7FD0
         Stretch         =   -1  'True
         Top             =   480
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   9
         Left            =   375
         Picture         =   "frmPieces.frx":80C2
         Stretch         =   -1  'True
         Top             =   810
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   10
         Left            =   375
         Picture         =   "frmPieces.frx":828C
         Stretch         =   -1  'True
         Top             =   1155
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   11
         Left            =   375
         Picture         =   "frmPieces.frx":837E
         Stretch         =   -1  'True
         Top             =   1500
         Width           =   300
      End
      Begin VB.Image IMW 
         Height          =   300
         Index           =   12
         Left            =   390
         Picture         =   "frmPieces.frx":8470
         Stretch         =   -1  'True
         Top             =   1815
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmPieces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmPieces.frm ~ RRChess ~

' No code

Option Explicit

