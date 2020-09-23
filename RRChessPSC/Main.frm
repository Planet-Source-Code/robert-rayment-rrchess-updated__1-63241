VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  ~ RRChess ~"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9030
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   602
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdColorScheme 
      BackColor       =   &H00C0FFFF&
      Height          =   225
      Index           =   2
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6525
      Width           =   165
   End
   Begin VB.CommandButton cmdColorScheme 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   1
      Left            =   7995
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6525
      Width           =   165
   End
   Begin VB.CommandButton cmdColorScheme 
      BackColor       =   &H0080C0FF&
      Height          =   225
      Index           =   0
      Left            =   7740
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6525
      Width           =   165
   End
   Begin VB.Frame fraPLY 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   7110
      TabIndex        =   42
      Top             =   6075
      Width           =   1350
      Begin VB.OptionButton optPLY 
         BackColor       =   &H0080C0FF&
         Caption         =   "5"
         Height          =   225
         Index           =   1
         Left            =   825
         TabIndex        =   44
         Top             =   75
         Width           =   420
      End
      Begin VB.OptionButton optPLY 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   405
         TabIndex        =   43
         Top             =   75
         Width           =   375
      End
      Begin VB.Shape Shape2 
         Height          =   390
         Left            =   0
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label LabPly 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ply"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   45
         Top             =   75
         Width           =   315
      End
   End
   Begin Project1.dmFrame fraWPP 
      Height          =   600
      Left            =   2145
      TabIndex        =   40
      Top             =   2850
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1058
      BackColor       =   14737632
      Caption         =   "Choose prom piece"
      BackColor       =   14737632
      OutLineColor    =   -2147483640
      BarColor        =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      MouseIcon       =   "Main.frx":0442
      Begin VB.Image IMWPP 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   1185
         Picture         =   "Main.frx":0894
         Stretch         =   -1  'True
         Top             =   225
         Width           =   300
      End
      Begin VB.Image IMWPP 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   840
         Picture         =   "Main.frx":0986
         Stretch         =   -1  'True
         Top             =   210
         Width           =   300
      End
      Begin VB.Image IMWPP 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   480
         Picture         =   "Main.frx":0A78
         Stretch         =   -1  'True
         Top             =   225
         Width           =   300
      End
      Begin VB.Image IMWPP 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   135
         Picture         =   "Main.frx":0B6A
         Stretch         =   -1  'True
         Top             =   225
         Width           =   300
      End
   End
   Begin Project1.dmFrame fraBPP 
      Height          =   600
      Left            =   2145
      TabIndex        =   41
      Top             =   2205
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1058
      BackColor       =   14737632
      Caption         =   "Choose prom piece"
      BackColor       =   14737632
      OutLineColor    =   -2147483640
      BarColor        =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      MouseIcon       =   "Main.frx":0C5C
      Begin VB.Image IMBPP 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   105
         Picture         =   "Main.frx":10AE
         Stretch         =   -1  'True
         Top             =   225
         Width           =   300
      End
      Begin VB.Image IMBPP 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   480
         Picture         =   "Main.frx":11A0
         Stretch         =   -1  'True
         Top             =   225
         Width           =   300
      End
      Begin VB.Image IMBPP 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   855
         Picture         =   "Main.frx":1292
         Stretch         =   -1  'True
         Top             =   210
         Width           =   300
      End
      Begin VB.Image IMBPP 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   1185
         Picture         =   "Main.frx":145C
         Stretch         =   -1  'True
         Top             =   225
         Width           =   300
      End
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5070
      Left            =   450
      ScaleHeight     =   336
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   336
      TabIndex        =   0
      Top             =   510
      Width           =   5070
      Begin VB.Image IM 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   0
         Left            =   30
         Stretch         =   -1  'True
         Top             =   30
         Width           =   480
      End
   End
   Begin VB.Frame fraMates 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      TabIndex        =   23
      Top             =   6450
      Width           =   5520
      Begin VB.OptionButton optUPDN 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   5100
         TabIndex        =   32
         Top             =   45
         Width           =   375
      End
      Begin VB.OptionButton optUPDN 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   31
         Top             =   45
         Width           =   360
      End
      Begin VB.CommandButton cmdMateNum 
         BackColor       =   &H0080C0FF&
         Caption         =   ">"
         Height          =   225
         Index           =   1
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   30
         Width           =   300
      End
      Begin VB.CommandButton cmdMateNum 
         BackColor       =   &H0080C0FF&
         Caption         =   "<"
         Height          =   225
         Index           =   0
         Left            =   2595
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   30
         Width           =   285
      End
      Begin VB.CommandButton cmdFMate 
         BackColor       =   &H0080C0FF&
         Caption         =   "W mates B ?"
         Height          =   225
         Index           =   0
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   30
         Width           =   1140
      End
      Begin VB.CommandButton cmdFMate 
         BackColor       =   &H0080C0FF&
         Caption         =   "B mates W ?"
         Height          =   225
         Index           =   1
         Left            =   1185
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   30
         Width           =   1140
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Search"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4110
         TabIndex        =   33
         Top             =   45
         Width           =   525
      End
      Begin VB.Label LabMN 
         BackColor       =   &H0080C0FF&
         Caption         =   "moves."
         Height          =   210
         Index           =   1
         Left            =   3495
         TabIndex        =   30
         Top             =   45
         Width           =   510
      End
      Begin VB.Label LabMN 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "in"
         Height          =   210
         Index           =   0
         Left            =   2340
         TabIndex        =   29
         Top             =   45
         Width           =   240
      End
      Begin VB.Label LabMateNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2895
         TabIndex        =   28
         Top             =   45
         Width           =   270
      End
   End
   Begin Project1.dmFrame fraSwap 
      Height          =   3900
      Left            =   6660
      TabIndex        =   12
      Top             =   1890
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   6879
      BackColor       =   8438015
      Caption         =   "Game on"
      BackColor       =   8438015
      OutLineColor    =   -2147483640
      BarColor        =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      MouseIcon       =   "Main.frx":154E
      RRDrawWidth     =   2
      Begin VB.Frame fraStepGame 
         BackColor       =   &H0080C0FF&
         Height          =   480
         Left            =   135
         TabIndex        =   15
         Top             =   3330
         Width           =   1815
         Begin VB.CommandButton cmdStepThruGame 
            BackColor       =   &H0080C0FF&
            Caption         =   ">|"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1350
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   150
            Width           =   360
         End
         Begin VB.CommandButton cmdStepThruGame 
            BackColor       =   &H0080C0FF&
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   975
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   150
            Width           =   360
         End
         Begin VB.CommandButton cmdStepThruGame 
            BackColor       =   &H0080C0FF&
            Caption         =   "|<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   150
            Width           =   360
         End
         Begin VB.CommandButton cmdStepThruGame 
            BackColor       =   &H0080C0FF&
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   150
            Width           =   360
         End
      End
      Begin VB.ListBox ListMoves 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         ItemData        =   "Main.frx":19A0
         Left            =   105
         List            =   "Main.frx":19A2
         TabIndex        =   14
         Top             =   540
         Width           =   1830
      End
      Begin VB.Label LabMvString 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabMvString"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   255
         Width           =   1830
      End
   End
   Begin Project1.dmFrame fraOptions 
      Height          =   1590
      Left            =   6795
      TabIndex        =   5
      Top             =   195
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   2805
      BackColor       =   8438015
      Caption         =   "Options"
      BackColor       =   8438015
      OutLineColor    =   -2147483642
      BarColor        =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      MouseIcon       =   "Main.frx":19A4
      RRDrawWidth     =   2
      Begin VB.OptionButton optSetUp 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "&Set up"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   9
         Top             =   945
         Width           =   1230
      End
      Begin VB.OptionButton optSetUp 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "&Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   8
         Top             =   1245
         Width           =   1140
      End
      Begin VB.OptionButton optSetUp 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "Set full &board"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   600
         Width           =   1305
      End
      Begin VB.OptionButton optSetUp 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "&Clear board"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   315
         Width           =   1305
      End
   End
   Begin Project1.dmFrame fraPieces 
      Height          =   3555
      Left            =   9030
      TabIndex        =   3
      Top             =   135
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   6271
      BackColor       =   -2147483638
      Caption         =   "Select"
      BackColor       =   -2147483638
      OutLineColor    =   -2147483640
      BarColor        =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      MouseIcon       =   "Main.frx":1DF6
      RRDrawWidth     =   4
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":2248
         Height          =   525
         Index           =   12
         Left            =   690
         Stretch         =   -1  'True
         Tag             =   "BP"
         Top             =   2955
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":2552
         Height          =   540
         Index           =   11
         Left            =   675
         Stretch         =   -1  'True
         Tag             =   "BK"
         Top             =   2400
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":285C
         Height          =   510
         Index           =   10
         Left            =   690
         Stretch         =   -1  'True
         Tag             =   "BQ"
         Top             =   1875
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":2B66
         Height          =   525
         Index           =   9
         Left            =   675
         Stretch         =   -1  'True
         Tag             =   "BB"
         Top             =   1335
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":2E70
         Height          =   525
         Index           =   8
         Left            =   675
         Stretch         =   -1  'True
         Tag             =   "BN"
         Top             =   795
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":317A
         Height          =   540
         Index           =   7
         Left            =   675
         Stretch         =   -1  'True
         Tag             =   "BR"
         Top             =   240
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":3484
         Height          =   525
         Index           =   6
         Left            =   75
         Stretch         =   -1  'True
         Tag             =   "WP"
         Top             =   2955
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":378E
         Height          =   540
         Index           =   5
         Left            =   75
         Stretch         =   -1  'True
         Tag             =   "WK"
         Top             =   2400
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":3A98
         Height          =   510
         Index           =   4
         Left            =   75
         Stretch         =   -1  'True
         Tag             =   "WQ"
         Top             =   1875
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":3DA2
         Height          =   525
         Index           =   3
         Left            =   75
         Stretch         =   -1  'True
         Tag             =   "WB"
         Top             =   1335
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":40AC
         Height          =   525
         Index           =   2
         Left            =   75
         Stretch         =   -1  'True
         Tag             =   "WN"
         Top             =   795
         Width           =   540
      End
      Begin VB.Image IMO 
         BorderStyle     =   1  'Fixed Single
         DragIcon        =   "Main.frx":43B6
         Height          =   540
         Index           =   1
         Left            =   75
         Stretch         =   -1  'True
         Tag             =   "WR"
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.PictureBox picCB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   10380
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   22
      Top             =   165
      Width           =   150
   End
   Begin VB.Label LabColortSchemes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7185
      TabIndex        =   49
      Top             =   6540
      Width           =   450
   End
   Begin VB.Label LabPC 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabPC"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   5805
      TabIndex        =   39
      Top             =   6570
      Width           =   1095
   End
   Begin VB.Label LabTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CTime s"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   5790
      TabIndex        =   38
      Top             =   6090
      Width           =   1110
   End
   Begin VB.Label LabMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabMessage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   195
      TabIndex        =   37
      Top             =   6090
      Width           =   5535
   End
   Begin VB.Label LabPB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   6060
      TabIndex        =   36
      Top             =   195
      Width           =   285
   End
   Begin VB.Label LabPB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   6375
      TabIndex        =   35
      Top             =   195
      Width           =   345
   End
   Begin VB.Label LabPB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   2
      Left            =   5760
      TabIndex        =   34
      Top             =   195
      Width           =   270
   End
   Begin VB.Shape Shape1 
      Height          =   330
      Left            =   180
      Top             =   6420
      Width           =   5565
   End
   Begin VB.Label LabCH 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "HUMAN"
      ForeColor       =   &H00000080&
      Height          =   210
      Index           =   0
      Left            =   1350
      TabIndex        =   21
      Top             =   5670
      Width           =   1125
   End
   Begin VB.Label LabCH 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080C0FF&
      Caption         =   "COMPUTER"
      ForeColor       =   &H00000040&
      Height          =   210
      Index           =   1
      Left            =   1380
      TabIndex        =   20
      Top             =   255
      Width           =   1050
   End
   Begin VB.Label LabMvIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "to move"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3165
      TabIndex        =   11
      Top             =   255
      Width           =   840
   End
   Begin VB.Label LabMvIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "to move"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3210
      TabIndex        =   10
      Top             =   5685
      Width           =   780
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   31
      Left            =   6270
      Stretch         =   -1  'True
      Top             =   5445
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   30
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   5445
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   29
      Left            =   6270
      Stretch         =   -1  'True
      Top             =   5070
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   28
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   5085
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   27
      Left            =   6225
      Stretch         =   -1  'True
      Top             =   4740
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   26
      Left            =   5865
      Stretch         =   -1  'True
      Top             =   4755
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   25
      Left            =   6195
      Stretch         =   -1  'True
      Top             =   4395
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   24
      Left            =   5850
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   23
      Left            =   6180
      Stretch         =   -1  'True
      Top             =   4050
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   22
      Left            =   5850
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   21
      Left            =   6225
      Stretch         =   -1  'True
      Top             =   3735
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   20
      Left            =   5835
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   19
      Left            =   6255
      Stretch         =   -1  'True
      Top             =   3585
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   18
      Left            =   5820
      Stretch         =   -1  'True
      Top             =   3405
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   17
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   16
      Left            =   5805
      Stretch         =   -1  'True
      Top             =   3285
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   15
      Left            =   6195
      Stretch         =   -1  'True
      Top             =   2505
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   14
      Left            =   5820
      Stretch         =   -1  'True
      Top             =   2505
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   13
      Left            =   6195
      Stretch         =   -1  'True
      Top             =   2175
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   12
      Left            =   5835
      Stretch         =   -1  'True
      Top             =   2175
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   11
      Left            =   6195
      Stretch         =   -1  'True
      Top             =   1815
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   10
      Left            =   5850
      Stretch         =   -1  'True
      Top             =   1830
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   9
      Left            =   6180
      Stretch         =   -1  'True
      Top             =   1470
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   8
      Left            =   5835
      Stretch         =   -1  'True
      Top             =   1455
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   7
      Left            =   6195
      Stretch         =   -1  'True
      Top             =   1155
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   6
      Left            =   5835
      Stretch         =   -1  'True
      Top             =   1140
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   5
      Left            =   6195
      Stretch         =   -1  'True
      Top             =   840
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   4
      Left            =   5835
      Stretch         =   -1  'True
      Top             =   840
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   3
      Left            =   6210
      Stretch         =   -1  'True
      Top             =   690
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   2
      Left            =   5850
      Stretch         =   -1  'True
      Top             =   630
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   1
      Left            =   6210
      Stretch         =   -1  'True
      Top             =   465
      Width           =   300
   End
   Begin VB.Image IMWSP 
      Height          =   300
      Index           =   0
      Left            =   5805
      Stretch         =   -1  'True
      Top             =   570
      Width           =   300
   End
   Begin VB.Label LabPC 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabPC"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   5805
      TabIndex        =   4
      Top             =   6330
      Width           =   1095
   End
   Begin VB.Label LabCul 
      BackColor       =   &H0080C0FF&
      Caption         =   "White"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   5685
      Width           =   585
   End
   Begin VB.Label LabCul 
      BackColor       =   &H0080C0FF&
      Caption         =   "Black"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2625
      TabIndex        =   1
      Top             =   255
      Width           =   510
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Load board FEN"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save board FEN"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Reload board"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "L&oad game CHG"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "S&ave game CHG"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "R&eload game"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuPrintCB 
         Caption         =   "Print Clipboard"
         Index           =   0
      End
      Begin VB.Menu mnuPrintCB 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuStartUp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Clear board"
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Set full &board"
         Index           =   1
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Set up"
         Index           =   2
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&PLAY"
         Index           =   3
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Flip"
         Index           =   6
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "S&ound"
         Index           =   9
      End
   End
   Begin VB.Menu mnuMov 
      Caption         =   "&Moves"
      Begin VB.Menu mnuMoves 
         Caption         =   "&White to move"
         Index           =   0
      End
      Begin VB.Menu mnuMoves 
         Caption         =   "&Black to move"
         Index           =   1
      End
      Begin VB.Menu mnuMoves 
         Caption         =   "&Any moves"
         Checked         =   -1  'True
         Index           =   2
      End
   End
   Begin VB.Menu mnuComp 
      Caption         =   "&Computer"
      Begin VB.Menu mnuComputer 
         Caption         =   "&No computer"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuComputer 
         Caption         =   "Computer plays &Black"
         Index           =   1
      End
      Begin VB.Menu mnuComputer 
         Caption         =   "Computer play &White"
         Index           =   2
      End
      Begin VB.Menu mnuComputer 
         Caption         =   "Compter plays both"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form1.frm

' ~RRChess~ by Robert Rayment  26 Nov 2005

Option Explicit

' API for TileForm
Private Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
 ByVal nWidth As Long, ByVal nHeight As Long, _
 ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' API to simulate Mouse Up
'Private Const MOUSEEVENTF_LEFTUP = &H4 ' left button up
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
   ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

'API for comp move time, sec
Private Declare Function GetTickCount Lib "Kernel32" () As Long
Private Tim As Long
' TEST
Private MOV$()

Private Sub cmdColorScheme_Click(Index As Integer)
' Public MainColor As Long, LightColor As Long
Dim k As Long
   Select Case Index
   Case 0   ' Default
      MainColor = &H80C0FF
      LightColor = &HC0E0FF
      Me.BackColor = MainColor
      Me.Picture = frmPieces.Picture
      TileForm1
   Case 1   ' Alternative 1
      MainColor = &HF8E4D8 '&HD8E9EC '&HC0C0C0
      LightColor = &HFFFFFF
      Me.Picture = LoadPicture("")
      Me.BackColor = &H8F8F80 '&HF8E4D8
   Case 2   ' Alternative 2
      MainColor = &HD8E9EC '&HC0C0C0
      LightColor = &HFFFFFF
      Me.Picture = LoadPicture("")
      Me.BackColor = &HB0C0C0 '&H8F8F80
   End Select
   For k = 0 To 1
      LabCH(k).BackColor = MainColor
      LabCul(k).BackColor = MainColor
      LabMvIndicator(k).BackColor = MainColor
      LabMN(k).BackColor = MainColor
      cmdFMate(k).BackColor = MainColor
      cmdMateNum(k).BackColor = MainColor
      optUPDN(k).BackColor = MainColor
      optPLY(k).BackColor = MainColor
      LabPC(k).BackColor = vbWhite
   Next k
   For k = 0 To 3
      optSetUp(k).BackColor = MainColor
      cmdStepThruGame(k).BackColor = MainColor
   Next k
   LabPly.BackColor = MainColor
   Label3.BackColor = MainColor
   fraPLY.BackColor = MainColor
   fraOptions.BackColor = MainColor
   ListMoves.BackColor = MainColor
   fraSwap.BackColor = MainColor
   fraMates.BackColor = MainColor
   Shape1.BackColor = MainColor
   fraStepGame.BackColor = MainColor
   LabColortSchemes.BackColor = Me.BackColor
   For k = 0 To 2
      LabPB(k).BackColor = LightColor
   Next k
   LabMessage.BackColor = LightColor
   LabTime.BackColor = LightColor
   CheckerPic Form1, picBoard, LabCH(0), LabCul(0), LabMvIndicator(0)
   Me.Refresh

End Sub

' Image boxes
' IM()    on-board piece icons sitting on picBoard
' IMO()   source holding off-board piece set
' IMWSP() showing taken pieces
' IMW()   source holding taken piece images on frmPieces.frm

Private Sub Form_Initialize()
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   LoadFENSpec$ = PathSpec$
   SaveFENSpec$ = PathSpec$
   LoadCHGSpec$ = PathSpec$
   SaveCHGSpec$ = PathSpec$
   
   MaxIndex = 30
   
   ReDim BeginBoard(1 To 8, 1 To 8)
   
   InitArrays
   
   aBlackAtTop = True
   WorBsMove$ = "A"  ' Any color to move
   CorHsMove$ = "H"

   FillScoreArrays

   SetStartUp
   
   LoadWavs
   StopPlay
   aBusy = False
   aExit = False
   
   mmul = 10    ' Material score multipier
   
   MainColor = &H80C0FF
   LightColor = &HC0E0FF

End Sub

Private Sub Form_Load()
Dim k As Long

   aMouseDown = False
   PieceCount = 0
   LabPC(0) = " B pieces = 0"
   LabPC(1) = " W pieces = 0"
   
   SetStartUp
   ClearTakenPieces Form1
   HalfMove = -1
   
   ' Copy piece icons to fraPieces Select frame
   With frmPieces
      For k = 1 To 12
         IMO(k).Picture = .IMO(k).Picture
      Next k
   End With
   
   SwapPieceList True   ' Puts up Select frame fraPieces
                        ' hides Game On frame  fraSwap
   Me.Picture = frmPieces.Picture
                        
   TileForm1
   Show
   CheckerPic Form1, picBoard, LabCH(0), LabCul(0), LabMvIndicator(0)
   ' Create/Load/Locate main board image boxes
   LoadImageBoxes Form1, IM(0)
   
   ' Start with SetUp
   aHold = True
   optSetUp(0).Value = True   ' Clear board
   aSetUp = True
   aHold = False
   LabMvString = ""
   
   GetAStartUp
   
   mnuOptions(9).Checked = aSound
      
   With fraPieces
      .Left = 464
      .Top = 132
      .Visible = True
   End With
   
   LabCH(1) = "HUMAN"
   LabCH(0) = "HUMAN"
   
   mnuComputer(1).Enabled = False
   mnuComputer(2).Enabled = False
   mnuComputer(3).Enabled = False

   LabPB(0) = 0
   LabPB(1) = 0
   LabPB(2) = "H"
   
   LabMateNum = "3"
   
   ' Mate in.. nums
   optUPDN(0).Value = True  ' Search 1 & 2
   UDSearch = 0
   
   ' Start screen
   'TODO: Splash
   
   LabMessage = ""
   LabMessage.Refresh
   ' Prom frames
   fraWPP.Visible = False
   fraBPP.Visible = False
   
   PLY = 4
   optPLY(0).Value = True
   LabMessage = Str$(PLY) & " PLY"
   
   
   cmdColorScheme(0).BackColor = MainColor
   cmdColorScheme(1).BackColor = &HF8E4D8
   cmdColorScheme(2).BackColor = &HD8E9EC
   
   Caption = "~RRChess~  RRChess  " & Date & "  " & Time
End Sub

Private Sub CheckForCompMove(Index As Integer)
' Computer's Public WorBsMove$ move
Dim PieceTaken$
Dim GS As Long

' Has to return Message$ (CHECKMATE, STALEMATE?), ippt, PieceTaken
   If CorHsMove$ = "H" Then Exit Sub
   If WorBsMove$ = "A" Then Exit Sub  ' Any
   
   If aSetUp Then
      Message$ = "Press Play"
      LabMessage = Message$
      LabMessage.Refresh
      Message$ = ""
   End If
   
   LabMessage = ""
   LabMessage.Refresh
   aBusy = True
   SavePosition Form1, 1   ' Display -> Board 1
   CopyALLBOOLS 0, 1 ' 0 -> 1
   
   Me.MousePointer = vbCustom
   Me.MouseIcon = LoadResPicture(101, 1)
   
   GS = HalfMove
   
   Disabler
   If HalfMove <= 0 Then LabTime = ""
   Tim = GetTickCount

   Call CompMove(Index)
     
   Tim = (GetTickCount - Tim) \ 1000
   If Tim > MaxTime Then MaxTime = Tim
   LabTime = Str$(Tim) & " s"
    
   Me.MousePointer = vbNormal
      
   FlashPiece  ' Using MoveString
      
   RestoreSavedPosition Form1, 1 ' Restore display from Board 1
   CopyMemory bRCBoard(1, 1, 0), bRCBoard(1, 1, 1), 64   ' Make Board 0 the same
   CopyALLBOOLS 1, 0 ' 1 -> 0
   SavePosition Form1, 0            ' Display -> Board 0
   
   aBusy = False
   
   LabPB(2) = "H"
   LabPB(1) = ""
   LabPB(1).Refresh
   LabPB(0) = ""
   LabPB(1).Refresh
   
   If InStr(1, Message$, "ERROR") Then
      GameOver
      Exit Sub
   End If
   If InStr(1, Message$, "STALEMATE") Then
      Play 6
      GameOver
      Exit Sub
   End If
   
   LabMvString = MoveString$
   ListTheMove ' Does HalfMove = ListMoves.NewIndex
   If InStr(4, MoveString, "x") <> 0 Then ' BPieceTaken
      ' DestPieceNum -> PieceTaken$
      ConvPNtoPNDescrip DestPieceNum, PieceTaken$
      ShowPieceTaken PieceTaken$
   End If
   
   Play 0
   
   ' WorBsMove = Comp color
   
   If Message$ <> "" Then  ' message occasionally doesn't work ?
      If InStr(1, Message$, "CHECKMATE") Then
         If WorBsMove = "W" Then
            Message$ = "BLACK IS CHECKMATED"
            Play 2
         Else
            Message$ = "WHITE IS CHECKMATED"
            Play 4
         End If
         GameOver
         Exit Sub
      End If

      If InStr(1, Message$, "CHECK") Then
         If WorBsMove = "W" Then
            Play 1
         Else
            Play 3
         End If
         Message$ = "CHECK" ' Re-inforce ??!!
         aBusy = False
         LabMessage = Message$
         LabMessage.Refresh
         Sleep 1
      End If
      
      If Message$ = "HALTED" Then
         Play 5
         Message$ = "HALTED" ' Re-inforce ??!!
         GameOver
         Exit Sub
      ElseIf Message$ = "DRAW" Then
         Message$ = "DRAW" ' Re-inforce ??!!
         Play 7
         GameOver
         Exit Sub
      End If
   End If
   
   Select Case CorHsMove$
   
   Case "CCW"  ' Comp v Comp
      CorHsMove$ = "CCB"
      WorBsMove$ = "B"
      mnuMoves_Click (1)  ' Switch 'to move'
      CheckForCompMove 1
   Case "CCB"  ' Comp v Comp
      CorHsMove$ = "CCW"
      WorBsMove$ = "W"
      mnuMoves_Click (0)  ' Switch 'to move'
      CheckForCompMove 1
   
   Case Else
      Select Case WorBsMove$
      Case "W": mnuMoves_Click (1)  ' Now "B"
      Case "B": mnuMoves_Click (0)  ' Now "W"
      Case "A": mnuMoves_Click (2)  ' Any
      End Select
      Enabler
   End Select
End Sub

Private Sub GameOver()
' Enter with Message$ (Public)
   LabMessage = Message$
   LabMessage.Refresh
   aBusy = False
   mnuMoves_Click (2)   ' "A" Any
   CorHsMove$ = "H"
   LabTime = "Max" & Str$(MaxTime) & " s"
   Enabler
End Sub

Private Sub FlashPiece()
Dim Col As Long
Dim Row As Long
Dim IMI
Dim k As Long
   Col = Asc(Mid$(MoveString$, 4, 1)) - 96
   Row = Asc(Mid$(MoveString$, 5)) - 48
   If 1 <= Col And Col <= 8 Then
   If 1 <= Row And Row <= 8 Then
     IMI = (8 - Row) * 8 + Col - 1
     For k = 1 To 4
        IM(IMI).Visible = False
        DoEvents
        Sleep 100
        IM(IMI).Visible = True
        DoEvents
        Sleep 100
     Next k
   End If
   End If
End Sub

Private Sub Disabler()
   mnuF.Enabled = False
   mnuOpt.Enabled = False
   mnuMov.Enabled = False
   mnuComp.Enabled = False
   mnuHelp.Enabled = False
   cmdFMate(0).Enabled = False
   cmdFMate(1).Enabled = False
   fraMates.Enabled = False
   fraMates.BackColor = &HCCCCCC
   fraOptions.Enabled = False
   fraSwap.Enabled = False
   picBoard.Enabled = False
   fraPLY.Enabled = False
   cmdColorScheme(0).Enabled = False
   cmdColorScheme(1).Enabled = False
   cmdColorScheme(2).Enabled = False
End Sub

Private Sub Enabler()
   picBoard.Enabled = True
   fraMates.Enabled = True
   fraMates.BackColor = MainColor
   fraOptions.Enabled = True
   fraSwap.Enabled = True
   cmdFMate(0).Enabled = True
   cmdFMate(1).Enabled = True
   mnuHelp.Enabled = True
   mnuComp.Enabled = True
   mnuMov.Enabled = True
   mnuOpt.Enabled = True
   mnuF.Enabled = True
   fraPLY.Enabled = True
   cmdColorScheme(0).Enabled = True
   cmdColorScheme(1).Enabled = True
   cmdColorScheme(2).Enabled = True
End Sub

Private Sub cmdMateNum_Click(Index As Integer)
' Sets Mate in N moves
Dim N As Long
   N = Val(LabMateNum)
   Select Case Index
   Case 0   ' <
      If N > 1 Then N = N - 1
   Case 1   ' >
      If N < 5 Then N = N + 1
   End Select
   LabMateNum = N
End Sub

Private Sub cmdFMate_Click(CIndex As Integer)
' Find mates in up to 5 moves
Dim WB$
Dim NMates As Long
   If aBusy Then Exit Sub
   
   MakeFENString WB$
   
   LabMessage = ""
   LabMessage.Refresh
   CountPieces 0
   If PieceCount = 0 Then
      Message$ = "No pieces"
      LabMessage = Message$
      LabMessage.Refresh
      Message$ = ""
      Exit Sub
   End If
   
   aBusy = True
   picBoard.SetFocus
   If Not aSetUp Then
      Disabler
   
      Form1.LabPB(0) = 0
      Form1.LabPB(1) = NFirstMoves
      Form1.Refresh

      Select Case CIndex
      Case 0: WB$ = "W" ' Does W mate B on Board 1
      Case 1: WB$ = "B" ' Does B mate W on Board 1
      End Select
      NMates = Val(LabMateNum) * 2 - 1 ' Convert 0,1,2,3,4 to 1,3,5,7,9
      FindMate WB$, NMates
   Else
      Message$ = "Press Play"
      LabMessage = Message$
      LabMessage.Refresh
      Message$ = ""
   End If
   Form1.LabPB(0) = 0
   Form1.LabPB(1) = NFirstMoves
   aBusy = False
   Enabler
End Sub

Private Sub optUPDN_Click(Index As Integer)
' Mate search direction
   UDSearch = Index
End Sub

Private Sub FindMate(WB$, NMates As Long)
Dim MM As Long  ' Returned Mate in MM moves
Dim PieceTaken$
      SetStartUp
      mnuComputer_Click 0
      
      Screen.MousePointer = vbHourglass
      MoveTot = 0
      
      TEST_FOR_MATE WB$, NMates, MM
      
      RestoreSavedPosition Form1, 1 ' Restore display from Board 1
      CopyMemory bRCBoard(1, 1, 0), bRCBoard(1, 1, 1), 64   ' Make Board 0 the same
      CopyALLBOOLS 1, 0 ' 1 -> 0
      Form1.LabMvString = MoveString$
      Form1.ListTheMove ' Does HalfMove = ListMoves.NewIndex
      Select Case WorBsMove$
      Case "W": Form1.mnuMoves_Click (1)  ' Now "B"
      Case "B": Form1.mnuMoves_Click (0)  ' Now "W"
      Case "A": Form1.mnuMoves_Click (2)
      End Select
      If InStr(4, MoveString, "x") <> 0 Then ' PieceTaken
         ' DestPieceNum -> PieceTaken$
         ConvPNtoPNDescrip DestPieceNum, PieceTaken$
         ShowPieceTaken PieceTaken$
      End If
      
      Screen.MousePointer = vbDefault
      Play 0
      If Message$ <> "" Then
         If InStr(1, Message$, "CHECKMATE") Then
            Play 2
            Message$ = "CHECKMATE" ' Re-inforce ??!!
            LabMessage = Message$
            LabMessage.Refresh
            InitArrays
            SavePosition Form1, 0      ' Display to bRCBoard(1-8, 1-8, 0) & show
         ElseIf Message$ = "CHECK" Then
            Play 1
            Message$ = "CHECK" ' Re-inforce ??!!
            LabMessage = Message$
            LabMessage.Refresh
         ElseIf InStr(1, Message$, "Key move") Then 'Message$ = Str$(MM) & " Key move = " & MoveString$
            LabMessage = Message$
            LabMessage.Refresh
         ElseIf Message$ = "STALEMATE" Then
            Message$ = "STALEMATE"   ' Re-inforce ??!!
            LabMessage = Message$
            LabMessage.Refresh
         ElseIf Message = "HALTED" Then
            LabMessage = Message$
            LabMessage.Refresh
         End If
      Else
         RestoreSavedPosition Form1, 0 ' Restore display from Board 0
         ListMoves.Clear
         HalfMove = 0
         Message$ = "NO MATE FOUND IN" & Str$(MM) & " MOVES"  ' Else "HALTED"
         LabMessage = Message$
         LabMessage.Refresh
      End If
      CountPieces 0
      LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
      LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
End Sub

Private Sub mnuFile_Click(Index As Integer)
If aBusy Then Exit Sub
Dim G As Long
Dim WB$

   LabMessage = ""
   
   Select Case Index
   Case 0, 2  ' Open setup FEN file, ReLoad
      If Index = 2 And LoadFENSpec$ = "" Then
         Message$ = "No board loaded yet"
         LabMessage = Message$
         LabMessage.Refresh
         Exit Sub
      End If
      If Open_FEN_File(Form1, Index) Then
         mnuMoves_Click 2
         SavePosition Form1, 0      ' Display to bRCBoard(1-8, 1-8, 0) & show
         If Not TestValidStartPosition Then
            LabMessage = Message$
            LabMessage.Refresh
            SavePosition Form1, 0      ' bRCBoard(1-8, 1-8, 0) & show
            aHold = True
            optSetUp(2).Value = True   ' Back to Setup
            aHold = False
         Else
            Reset_ALL_BOOLS
            SavePosition Form1, 0      ' Display to bRCBoard(1-8, 1-8, 0)
            CountPieces 0
            LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
            LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
            ReDim BeginBoard(1 To 8, 1 To 8)
            CopyMemory BeginBoard(1, 1), bRCBoard(1, 1, 0), 64
            SavePosition Form1, 1      ' Display to bRCBoard(1-8, 1-8, 1)
            SetStartUp
            ClearTakenPieces Form1
            SwapPieceList False
            ListMoves.Clear
            aHold = True
            optSetUp(3).Value = True   ' Play
            aSetUp = False
            aHold = False
            mnuMov.Enabled = False
            mnuComp.Enabled = False
            
            Caption = "~RRChess~  " & FindName$(LoadFENSpec$)
         End If
      End If
   Case 1   ' Save setup FEN file
      If Not Save_FEN_File(Form1) Then
         MsgBox "FEN Save Error or Cancelled", vbCritical, "Saving"
      Else
         Caption = "~RRChess~  " & FindName$(SaveFENSpec$)
      End If
   Case 3   ' Break
   
   Case 5   ' Save game chg file
      ' Wind back to start
      cmdStepThruGame_MouseUp 2, 0, 0, 0, 0 '  >| (Index,,,,)
      GameOffset = -1
      ' Saves postion from main display to board index
      SavePosition Form1, 0
      ' FEN color for MakeFENString
      WB$ = ListMoves.List(ListMoves.ListCount - 1)
      WB$ = Mid$(WB$, 4)
      WB$ = LTrim$(WB$)
      WB$ = Left$(WB$, 1)
      If WB$ = "W" Then WB$ = "B" Else WB$ = "W"
      MakeFENString WB$
      If Not Save_chg_File(Form1, ListMoves, WB$) Then
         MsgBox "CHG Save Error or Cancelled", vbCritical, "Saving"
      Else
         Caption = "~RRChess~  " & FindName$(SaveCHGSpec$)
      End If
   Case 4, 6 ' Load game chg file, Reload game
      If Open_chg_File(Form1, ListMoves, Index - 4) Then ' 0 or 1
         SetStartUp
         Reset_ALL_BOOLS
         ClearTakenPieces Form1
         SwapPieceList False
         LabMvString = ""
         'FullSetUp Form1
         SavePosition Form1, 0 ' Display to Board 0
         CopyMemory BeginBoard(1, 1), bRCBoard(1, 1, 0), 64
         
         aHold = True
         optSetUp(3).Value = True   ' Play
         aSetUp = False
         aHold = False
         
         Caption = "~RRChess~  " & FindName$(LoadCHGSpec$)
         
         cmdStepThruGame_MouseUp 1, 0, 0, 0, 0 '  >| (Index,,,,)
         G = HalfMove
         ' Find which color to move next
         WB$ = Trim$(ListMoves.List(ListMoves.ListCount - 1))
         WB$ = UCase$(Mid$(WB$, (InStr(1, WB$, " ") + 1), 1))
         
         If WB$ = "W" Then
            WorBsMove$ = "B"
         Else
            WorBsMove$ = "W"
         End If
         
         If WorBsMove$ <> "A" Then
            If WorBsMove$ = "B" Then
               mnuMoves_Click 1  ' Set B to move
            Else
               mnuMoves_Click 0  ' Set W to move
            End If
         End If
         mnuComputer_Click 0     ' No computer
      End If
   Case 7   ' Break
   End Select
   aDraggedOut = False
   aMouseDown = False
   CountPieces 0
   LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
   LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
   
   Me.SetFocus
End Sub

Public Sub Flip()
If aBusy Then Exit Sub
Dim k As Long
Dim ixc As Long, iyc As Long
   aBlackAtTop = Not aBlackAtTop
   ixc = picBoard.Width \ 2
   iyc = picBoard.Height \ 2
   For k = 0 To 63
      IM(k).Left = 2 * ixc - IM(k).Left - 2 * margin - SQ \ 2
      IM(k).Top = 2 * iyc - IM(k).Top - 2 * margin - SQ \ 2
   Next k
   CheckerPic Form1, picBoard, LabCH(0), LabCul(0), LabMvIndicator(0)
   If aBlackAtTop Then
      LabCul(1) = "Black"
      LabCul(0) = "White"
      If WorBsMove$ = "W" Then
         If CorHsMove$ = "CW" Then
            LabCH(0) = "COMPUTER"
            LabCH(1) = "HUMAN"
         ElseIf CorHsMove$ = "CB" Then
            LabCH(1) = "COMPUTER"
            LabCH(0) = "HUMAN"
         Else
            LabCH(0) = "HUMAN"
            LabCH(1) = "HUMAN"
         End If
      ElseIf WorBsMove$ = "B" Then
         If CorHsMove$ = "CW" Then
            LabCH(0) = "COMPUTER"
            LabCH(1) = "HUMAN"
         ElseIf CorHsMove$ = "CB" Then
            LabCH(1) = "COMPUTER"
            LabCH(0) = "HUMAN"
         Else
            LabCH(0) = "HUMAN"
            LabCH(1) = "HUMAN"
         End If
      End If
   ElseIf Not aBlackAtTop Then ' Black at bottom
      LabCul(1) = "White"
      LabCul(0) = "Black"
      If WorBsMove$ = "W" Then
         If CorHsMove$ = "CW" Then
            LabCH(1) = "COMPUTER"
            LabCH(0) = "HUMAN"
         ElseIf CorHsMove$ = "CB" Then
            LabCH(0) = "COMPUTER"
            LabCH(1) = "HUMAN"
         Else
            LabCH(0) = "HUMAN"
            LabCH(1) = "HUMAN"
         End If
      ElseIf WorBsMove$ = "B" Then
         If CorHsMove$ = "CW" Then
            LabCH(1) = "COMPUTER"
            LabCH(0) = "HUMAN"
         ElseIf CorHsMove$ = "CB" Then
            LabCH(0) = "COMPUTER"
            LabCH(1) = "HUMAN"
         Else
            LabCH(0) = "HUMAN"
            LabCH(1) = "HUMAN"
         End If
      End If
   End If
   
   LabMvIndicator(0).Visible = True
   LabMvIndicator(1).Visible = True
   Select Case WorBsMove$
   Case "W": mnuMoves_Click (0)  ' W
   Case "B": mnuMoves_Click (1)  ' B
   Case Else: mnuMoves_Click (2) ' Any
   End Select
   picBoard.SetFocus
End Sub

Public Sub mnuMoves_Click(Index As Integer)  ' 0 W , 1 B, 2 Any
If aBusy Then Exit Sub
Dim k As Long
   For k = 0 To 2
      mnuMoves(k).Checked = False
   Next k
   mnuMoves(Index).Checked = True
   LabMvIndicator(0).Visible = True ' to move Bottom
   LabMvIndicator(1).Visible = True ' to move Top
   Select Case Index
   Case 0   ' "W"
      WorBsMove$ = "W"
      If aBlackAtTop Then
         LabMvIndicator(1).Visible = False
      Else
         LabMvIndicator(0).Visible = False
      End If
      mnuComputer(1).Enabled = True
      mnuComputer(2).Enabled = True
      mnuComputer(3).Enabled = True
   Case 1   ' "B"
      WorBsMove$ = "B"
      If aBlackAtTop Then
         LabMvIndicator(0).Visible = False
      Else
         LabMvIndicator(1).Visible = False
      End If
      mnuComputer(1).Enabled = True
      mnuComputer(2).Enabled = True
      mnuComputer(3).Enabled = True
   Case 2   ' "A"
      WorBsMove$ = "A"
      LabMvIndicator(1).Visible = True
      LabMvIndicator(0).Visible = True
      mnuComputer_Click 0  ' No computer
      mnuComputer(1).Enabled = False
      mnuComputer(2).Enabled = False
      mnuComputer(3).Enabled = False
      LabCH(1) = "HUMAN"
      LabCH(0) = "HUMAN"
   End Select
   LabMvIndicator(0).Refresh
   LabMvIndicator(1).Refresh
End Sub

Private Sub mnuComputer_Click(Index As Integer)
If aBusy Then Exit Sub
Dim k As Long
   For k = 0 To 3
      mnuComputer(k).Checked = False
   Next k
   mnuComputer(Index).Checked = True
   Select Case Index
   Case 0   ' No computer
      CorHsMove$ = "H"
      LabCH(1) = "HUMAN"
      LabCH(0) = "HUMAN"
      If WorBsMove$ = "W" Then
         mnuMoves_Click 0 ' 0 W , 1 B, 2 Any
      ElseIf WorBsMove$ = "B" Then
         mnuMoves_Click 1 ' 0 W , 1 B, 2 Any
      End If
   Case 1   ' Computer plays Black
      CorHsMove$ = "CB"
      If aBlackAtTop Then
         LabCH(1) = "COMPUTER"
         LabCH(0) = "HUMAN"
      Else
         LabCH(1) = "HUMAN"
         LabCH(0) = "COMPUTER"
      End If
      aHold = True
      optSetUp(2).Value = True
      aHold = False
   
   Case 2   ' Computer plays White
      CorHsMove$ = "CW"
      If aBlackAtTop Then
         LabCH(1) = "HUMAN"
         LabCH(0) = "COMPUTER"
      Else
         LabCH(1) = "COMPUTER"
         LabCH(0) = "HUMAN"
      End If
      aHold = True
      optSetUp(2).Value = True
      aHold = False
   Case 3
      If WorBsMove$ = "W" Then
         CorHsMove$ = "CCW"
      Else
         CorHsMove$ = "CCB"
      End If
      LabCH(1) = "COMPUTER"
      LabCH(0) = "COMPUTER"
      aHold = True
      optSetUp(2).Value = True
      aHold = False
   End Select
End Sub

Private Sub mnuOptions_Click(Index As Integer)
If aBusy Then Exit Sub
   
   Select Case Index
   Case 0 To 3
      aHold = True
      optSetUp(Index).Value = True
      aHold = False
   End Select
   
   Select Case Index
   Case 0: optSetUp_Click 0  ' Clear board
   Case 1: optSetUp_Click 1  ' Full board
   Case 2: optSetUp_Click 2  ' Setup
   Case 3: optSetUp_Click 3  ' Play
   Case 4   ' Break
   Case 5   ' Break
   Case 6: Flip
   Case 7   ' Break
   Case 9   ' Sound
      mnuOptions(9).Checked = Not mnuOptions(9).Checked
      aSound = -mnuOptions(9).Checked
   End Select
   Me.SetFocus
End Sub


Private Sub optPLY_Click(Index As Integer)
   PLY = Index + 4
   LabMessage.Caption = Str$(PLY) & " PLY"
End Sub

Private Sub optSetUp_Click(Index As Integer)
   If aBusy Then
      Refresh
      DoEvents
      Exit Sub
   End If

   If aHold Then Exit Sub
   
   
   
   If Index = 0 Then ' Clear board
      ClearBoard Form1
      SetStartUp
   End If
   
   Select Case Index
   Case 0, 2 ' Clear board & SetUp
      mnuMov.Enabled = True
      mnuComp.Enabled = True
      fraMates.Enabled = True
      fraMates.BackColor = MainColor
      ClearTakenPieces Form1
      SavePosition Form1, 0
      aHold = True
      optSetUp(2).Value = True   ' To SetUp
      aSetUp = True
      aHold = False
      ListMoves.Clear
      LabMvString = ""
      SwapPieceList True
      With fraPieces
         .Left = 464
         .Top = 132
         .Visible = True
      End With
      Reset_ALL_BOOLS
      LabMessage = ""
      'mnuMoves_Click 2  ' WorBsMove$ = "A" HUMAN HUMAN
   Case 1   ' Full board
      mnuMov.Enabled = True
      mnuComp.Enabled = True
      fraMates.Enabled = True
      fraMates.BackColor = MainColor
      ClearTakenPieces Form1
      aHold = True
      optSetUp(2).Value = True   ' SetUp
      aSetUp = True
      aHold = False
      ListMoves.Clear
      LabMvString = ""
      SwapPieceList True
      With fraPieces
         .Left = 464
         .Top = 132
         .Visible = True
      End With
      
      SetStartUp
      FullSetUp Form1   ' Also does SavePosition Form1, 0  'Display to Board 0
      SavePosition Form1, 0      ' Display to bRCBoard(1-8, 1-8, 0)
      CopyMemory BeginBoard(1, 1), bRCBoard(1, 1, 0), 64
      HalfMove = -1
      Reset_ALL_BOOLS
   Case 3   ' Play
      If PieceCount > 0 Then
         ' Check if acceptable start position on Play
         ' Test for missing kings, king next to king,
         ' king in check & pawns in unacceptable positions.
         If Not TestValidStartPosition Then
            LabMessage = Message$
            LabMessage.Refresh
            SavePosition Form1, 0      ' bRCBoard(1-8, 1-8, 0) & show
            aHold = True
            aSetUp = True
            aHold = False
            ListMoves.Clear
            HalfMove = 0
            mnuMoves_Click 2  ' WorBsMove$ = "A" HUMAN HUMAN
         Else  ' START PLAY
            MaxTime = 0
            FirstColor$ = WorBsMove$
            MakeFENString WorBsMove$
            NumRepPositions = 1
            OpenString$ = ""
            GameOffset = 0
            aSetUp = False
            SetStartUp
            ClearTakenPieces Form1
            SavePosition Form1, 0      ' Display to Board 0
            ''If CorHsMove$ = "H" Then
            CopyMemory BeginBoard(1, 1), bRCBoard(1, 1, 0), 64
            ''End If
            SwapPieceList False  ' Swap in ListMoves etc
            If CorHsMove$ <> "H" Then
               aHold = True
               optSetUp(3).Value = True   ' Play
               aHold = False
            End If
            If WorBsMove$ = "W" And CorHsMove$ = "CW" Then
               CheckForCompMove 1
            ElseIf WorBsMove$ = "B" And CorHsMove$ = "CB" Then
               CheckForCompMove 1
            ElseIf CorHsMove$ = "CCW" Then
               WorBsMove$ = "W"
               CheckForCompMove 1
            ElseIf CorHsMove$ = "CCB" Then
               WorBsMove$ = "B"
               CheckForCompMove 1
            End If
         End If
      Else  ' PieceCount = 0
         SavePosition Form1, 0      ' bRCBoard(1-8, 1-8, 0) & show
         optSetUp(2).Value = True
         ListMoves.Clear
         SwapPieceList True
      End If
   End Select
   
   aDraggedOut = False
   aMouseDown = False
   
   CountPieces 0
   LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
   LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
   
   Me.SetFocus
End Sub

Private Function TestValidStartPosition() As Boolean
Dim C As Long
      
   SavePosition Form1, 0      ' bRCBoard(1-8, 1-8, 0)
   CountPieces 0
   LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
    LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
   TestValidStartPosition = False
   Message$ = "Missing Kings"
   If Not FindKingRC("W", rwk, cwk, 0) And _
      Not FindKingRC("B", rbk, cbk, 0) Then Exit Function
   
   Message$ = " Missing White King"
   If Not FindKingRC("W", rwk, cwk, 0) Then Exit Function
   Message$ = " Missing Black King"
   If Not FindKingRC("B", rbk, cbk, 0) Then Exit Function
   ' Have 2 kings
   ' Test if K next to K
   Message$ = " King next to King"
   If CLng(Sqr((rwk - rbk) ^ 2 + (cwk - cbk) ^ 2)) = 1 Then Exit Function
   ' Test if K in check
   Message$ = "White King in check!"
   If RC_Targetted(rwk, cwk, "B", 0) Then Exit Function ' Test if WK in check
   Message$ = "Black King in check!"
   If RC_Targetted(rbk, cbk, "W", 0) Then Exit Function  ' Test if BK in check
   ' Test if any pawns on 1st or 8th row
   For C = 56 To 63 ' row 1 & row 8
      If (IM(C).Tag = "WP" Or IM(C).Tag = "BP") Or _
         (IM(C - 56).Tag = "WP" Or IM(C - 56).Tag = "BP") Then
         Exit For
      End If
   Next C
   Message$ = "Pawn(s) in unacceptable start position"
   If C < 64 Then Exit Function
   
   CountPieces 0
   LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
   LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
   Message$ = "More than one White King"
   If NumWK > 1 Then Exit Function
   Message$ = "More than one Black King"
   If NumBK > 1 Then Exit Function
   
   Message$ = ""
   TestValidStartPosition = True
End Function

Private Sub IMO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If aBusy Then Exit Sub
' SetUp - Pick up piece
   If aSetUp Then
      IMOIndex = Index
      IMO(Index).Drag 1
   End If
End Sub

Private Sub IM_DblClick(Index As Integer)
If aBusy Then Exit Sub
' Clear piece in Setup
   If aSetUp Then
      If IM(Index).Tag <> "" Then
         IM(Index).Picture = LoadPicture
         IM(Index).Tag = ""
         IM(Index).DragIcon = LoadPicture
         PieceCount = PieceCount - 1
         'LabPC = "PieceCount =" & Str$(PieceCount)
      End If
   End If
End Sub

Private Sub IM_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If aBusy Then Exit Sub
   If Not aSetUp Then
      If (WorBsMove$ = "W" And CorHsMove$ = "CW") Or _
         (WorBsMove$ = "B" And CorHsMove$ = "CB") Then
         LabMessage = ""
         LabMessage.Refresh
      End If
   End If
End Sub

Private Sub IM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If aBusy Then Exit Sub
' Mouse down on a piece
   
   'If frmMsg.Visible Then Unload frmMsg
   LabMessage = ""
   LabMessage.Refresh
   If IM(Index).Tag = "" Then Exit Sub
   
   IMIndex = Index
   ' Save original position in case piece
   ' moved off boad or illegal move
   IMLeft = IM(Index).Left
   IMTop = IM(Index).Top
   
   
   If Not aSetUp Then
      If WorBsMove$ = "W" And CorHsMove$ = "CW" Then CheckForCompMove 1
   End If
   
   'If aCheckmate Then Exit Sub   ' TODO Board/Moves after Checkmate
   
   aDraggedOut = False
   aMouseDown = True
   IMIndex = Index
   IM(Index).Drag 1
   
   ' Clear piece square so that only that piece visibly moving
   If Not aSetUp Then IM(Index).Visible = False
   
   ' Save original position in case piece
   ' moved off boad or illegal move
   IMLeft = IM(Index).Left
   IMTop = IM(Index).Top
   
   ' Show Start Square a1 -> h8
   ShowSquare IM(Index), MoveString$
   LabMvString = MoveString$
End Sub

Private Sub IM_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
If aBusy Then Exit Sub
' Drop piece on board
Dim R As Long, C As Long
Dim a$
   
   ' Setup
   
   If aSetUp Then
      DoSetUp Index
      Exit Sub
   End If
   
   ' Play
   
   If Index = IMIndex Then ' D-Click on same square
      IM(IMIndex).Visible = True
      Exit Sub
   End If
   
   PlayColor$ = Left$(IM(IMIndex).Tag, 1)
   If PlayColor$ = "W" Then OppColor$ = "B" Else OppColor$ = "W"
   
   
   If PlayColor$ <> WorBsMove$ Then
      If WorBsMove$ <> "A" Then ' Wrong color
         ResetForIllegalMove IM(IMIndex)
         Exit Sub
      End If
   End If
   
   If Not aSetUp Then
      If HalfMove < ListMoves.ListCount - 1 Then
         HalfMove = ListMoves.ListCount - 1
         
   ListMoves.Selected(HalfMove) = True
   If ListMoves.ListCount > 10 Then
      ListMoves.TopIndex = ListMoves.TopIndex + 1
   End If
         
         Source.Move IMLeft, IMTop
         Source.Visible = True
         'Exit Sub
      End If
   End If
   
   ' Convert Drop index to a1 -> h8 notation
   R = 8 - Index \ 8
   C = Index - (64 - 8 * R) + 1
   a$ = Chr$(C + 96) & Trim$(Str$(R))
   
   If IM(Index).Tag = "" Then
      a$ = "-" & a$
   Else
      a$ = "x" & a$
   End If

   ' Show Move a1-h8
   MoveString$ = MoveString$ & a$
   LabMvString = MoveString$
   
   SavePosition Form1, 0   ' Display to bRCBoard(1-8, 1-8, 0)
   
   'IM(IMIndex).Visible = True
   aLegalMove = LegalMove(MoveString$)
   
   If Not aLegalMove Then   ' Return piece to original position
      ResetForIllegalMove IM(IMIndex)
      If Message$ <> "" Then
         LabMessage = Message$
         LabMessage.Refresh
      End If
      ResetForIllegalMove IM(IMIndex)
   ElseIf (PlayColor$ = "W" And aWKingInCheck) Or (PlayColor$ = "B" And aBKingInCheck) Then
      ResetForIllegalMove IM(IMIndex)
      If Message$ <> "" Then
         LabMessage = Message$
         LabMessage.Refresh
      End If
      ResetForIllegalMove IM(IMIndex)
   Else  ' Legal
      If Message$ <> "" Then
         LabMessage = Message$
         LabMessage.Refresh
      End If
      SetPiece Index
   End If
End Sub

Private Sub picBoard_DragDrop(Source As Control, x As Single, y As Single)
If aBusy Then Exit Sub
' Drop piece on board
' Source IM(Index)
Dim ixc As Long, iyr As Long
Dim LandIndex As Integer
Dim R As Long, C As Long
Dim a$
   
   If Not aSetUp Then
      If HalfMove < ListMoves.ListCount - 1 Then
         Source.Move IMLeft, IMTop
         Source.Visible = True
         Exit Sub
      End If
   End If
   
   If Not aSetUp And aDraggedOut Then ' Off board
      Source.Move IMLeft, IMTop
      Source.Visible = True
      'Mouse Up to clear off-board piece image
      mouse_event &H4, 0, 0, 0, 0
      MoveString$ = ""
      LabMvString = MoveString$
      Exit Sub
   End If
   
   
   ' Get 0,0 - 7,7 (SQ=42)
   ixc = x \ SQ
   iyr = y \ SQ
   LandIndex = ixc + 8 * iyr
   If LandIndex < 0 Then
      LandIndex = 0
   End If
   If LandIndex > 63 Then
      LandIndex = 63
   End If
   If Not aBlackAtTop Then
      LandIndex = 63 - LandIndex
   End If
   
   ' Setup

   If aSetUp Then
      DoSetUp LandIndex
      Exit Sub
   End If
   
   ' Play
   
   If Left$(IM(IMIndex).Tag, 1) <> WorBsMove$ Then
      If WorBsMove$ <> "A" Then
         ResetForIllegalMove IM(IMIndex)
         Exit Sub
      End If
   End If
   
   ' Convert Drop LandIndex to a1 -> h8 notation
   R = 8 - LandIndex \ 8
   C = LandIndex - (64 - 8 * R) + 1
   a$ = Chr$(C + 96) & Trim$(Str$(R))
   
   If IM(LandIndex).Tag = "" Then
      a$ = "-" & a$
   Else
      a$ = "x" & a$
   End If
         
   ' Show Move a1-h8
   MoveString$ = MoveString$ & a$
   LabMvString = MoveString$
   
   SavePosition Form1, 0      ' bRCBoard(1-8, 1-8, 0)
   
   aLegalMove = LegalMove(MoveString$)
         
   If Not aLegalMove Then  ' Illegal
      ResetForIllegalMove Source
      If Message$ <> "" Then
         LabMessage = Message$
         LabMessage.Refresh
      End If
      ResetForIllegalMove Source
   ElseIf (PlayColor$ = "W" And aWKingInCheck) Or (PlayColor$ = "B" And aBKingInCheck) Then
      ResetForIllegalMove IM(IMIndex)
      If Message$ <> "" Then
         LabMessage = Message$
         LabMessage.Refresh
      End If
      ResetForIllegalMove IM(IMIndex)
   Else
      If Message$ <> "" Then
         LabMessage = Message$
         LabMessage.Refresh
      End If
      SetPiece LandIndex
   End If
End Sub

Private Sub DoSetUp(i As Integer)
   If Not aMouseDown Then
      SetUp_PlacePiece IM(i), IMO(IMOIndex)
   Else  ' Moving one piece on top of another in SetUp
         ' or just moving a piece into an empty image box
      SetUp_MovePiece IM(i), IM(IMIndex)
   End If
   aMouseDown = False
   ' Show Start Square a1 -> h8
   ShowSquare IM(i), MoveString$
   LabMvString = MoveString$
   fraOptions.SetFocus
End Sub

Private Function SetPiece(i As Integer) As Boolean
Dim a$
Dim PieceTaken$
   SetPiece = False
   ippt = -1 ' piece num promoted or taken
   
   SavePosition Form1, 0   ' Display -> Board 0
   
   CountPieces 0
   LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
   LabPC(1) = " W pieces =" & Str$(WhitePieceCount)

   PieceTaken$ = Left$(IM(i).Tag, 2)  ' Before moving piece
                                      ' for ShowPieceTaken when all OK
   ' MAKE THE MOVE
   ' Show piece at new position
   IM(i).Picture = IM(IMIndex).Picture
   IM(i).Tag = IM(IMIndex).Tag
   IM(i).DragIcon = IM(IMIndex).DragIcon
   IM(i).Visible = True
   ' i is now the destination
   ' Clear start position
   IM(IMIndex).Picture = LoadPicture
   IM(IMIndex).Tag = ""
   IM(IMIndex).DragIcon = LoadPicture
   IM(IMIndex).Visible = True
   ' IMIndex is now blank
   
   DoEvents
   
   ' New position
   SavePosition Form1, 1   ' Display -> Board 1
         'also returns rwk,cwk & rbk,cbk king positions

   ' Test pawn promotion
   If (Left$(IM(i).Tag, 2) = "WP" And WEPP(WEPProm, 0) = 1) Or _
      (Left$(IM(i).Tag, 2) = "BP" And BEPP(WEPProm, 0) = 1) Then
      ' If OK returns ippt the promotion piece number
      If Not TestPromotion(i) Then
         RestoreSavedPosition Form1, 0 ' from bRCBoard(1-8, 1-8, 0)
         Exit Function ' Promotion puts King in check
      Else
         ' ippt promotion piece number returned
      End If
   End If
   
   ' Test castling
   aCastling(0) = 0
   If Mid$(IM(i).Tag, 1, 2) = "WK" Then ' "WK"
      If Cast(WKSCastOK, 0) = 1 Or Cast(WQSCastOK, 0) = 1 Then ' aKSideCastling or aQSideCastling
            DoCastling (i)
      End If
   ElseIf Mid$(IM(i).Tag, 1, 2) = "BK" Then ' "BK"
      If Cast(BKSCastOK, 0) = 1 Or Cast(BQSCastOK, 0) = 1 Then ' aKSideCastling or aQSideCastling
            DoCastling (i)
      End If
   End If
   
   ' Test en passant
   If Mid$(IM(i).Tag, 1, 2) = "WP" Then  ' "WP" or "BP"
      If ColS <> ColE And WEPP(WEPOK, 0) = 1 Then
         PieceTaken$ = "BP"   ' For ShowPieceTaken
         IM(i + 8).Picture = LoadPicture
         IM(i + 8).Tag = ""
         IM(i + 8).DragIcon = LoadPicture
         WEPP(WEPOK, 0) = 0
      End If
      ' One chance for en passant
    ElseIf Mid$(IM(i).Tag, 1, 2) = "BP" Then
      If ColS <> ColE And BEPP(BEPOK, 0) = 1 Then
         PieceTaken$ = "WP"   ' For ShowPieceTaken
         IM(i - 8).Picture = LoadPicture
         IM(i - 8).Tag = ""
         IM(i - 8).DragIcon = LoadPicture
         BEPP(BEPOK, 0) = 0
      End If
   End If
   
   ' New position. Saves any promotion, castling or en passant
   
   SavePosition Form1, 1   ' Display -> Board 1
   ' returns rwk,cwk & rbk,cbk king positions
   CopyALLBOOLS 0, 1
   CountPieces 1
   LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
   LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
   
   ' Test for Checks & CHECKMATE
   
   If Mid$(IM(i).Tag, 1, 1) = "W" Then ' White moved
      If RC_Targetted(rbk, cbk, "W", 1) Then ' Test if BK in check by White
         If TestForCheckMate("B", 1) Then   ' BK @ rbk,cbk
            aCheckmate = True
            ListTheMove
            Play 2
            Message$ = "BLACK IS CHECKMATED !"
            LabMessage = Message$
            LabMessage.Refresh
            Message$ = "BLACK IS CHECKMATED !"  ' Re-inforce ??!!
            LabMessage = Message$
            LabMessage.Refresh
            
         aBusy = False
         mnuMoves_Click (2)   ' "A"
         CorHsMove$ = "H"
         Enabler
            
            
            SavePosition Form1, 0   ' Display -> Board 0
            SetPiece = True
            Exit Function     ' TODO WRAP
         Else
            Play 1   ' "W"
            Message$ = "Black King in CHECK!"
            LabMessage = Message$
            LabMessage.Refresh
            SavePosition Form1, 0   ' Display -> Board 0
            SavePosition Form1, 1   ' Display -> Board 1
         End If
      End If
   
   Else  ' Black move
      If RC_Targetted(rwk, cwk, "B", 1) Then ' Test if WK in check by Black
         If TestForCheckMate("W", 1) Then   ' WK @ rwk,cwk
            aCheckmate = True
            ListTheMove
            Play 4
            Message$ = "WHITE IS CHECKMATED !"
            LabMessage = Message$
            LabMessage.Refresh
            Message$ = "WHITE IS CHECKMATED !"  ' Re-inforce ??!!
            LabMessage = Message$
            LabMessage.Refresh
            
         aBusy = False
         mnuMoves_Click (2)   ' "A"
         CorHsMove$ = "H"
         Enabler
            
            SavePosition Form1, 0   ' Display -> Board 0
            SetPiece = True
            Exit Function  ' TODO WRAP
         Else
            Play 3
            Message$ = "White King in CHECK!"
            LabMessage = Message$
            LabMessage.Refresh
            SavePosition Form1, 0   ' Display -> Board 0
            SavePosition Form1, 1   ' Display -> Board 1
         End If
      End If
   
   End If
   
   ShowPieceTaken PieceTaken$
   ' Returns ippt piece taken number &
   ' W/BPTaken = W/BPTaken + 1
   ' aW/BPieceTaken = True
   CountPieces 1
   
   LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
   LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
   ListTheMove ' Also sets HalfMove
   
   IM(IMIndex).Visible = True
   Me.SetFocus
   SetPiece = True
   ' Swap move color
   Select Case WorBsMove$
   Case "W"
      mnuMoves_Click (1)  ' Now WorBsMove$ = "B"
   Case "B"
      mnuMoves_Click (0)  ' Now WorBsMove$ = "W"
   Case "A"
      mnuMoves_Click (2)  ' Now WorBsMove$ = "A"
   End Select
   Play 0
   a$ = LabMvString
   If WorBsMove$ = "W" And CorHsMove$ = "CW" Then
      CheckForCompMove 1
   ElseIf WorBsMove$ = "B" And CorHsMove$ = "CB" Then
      CheckForCompMove 1
   End If
End Function

Public Sub ListTheMove()
Dim k As Long, N$
   k = ListMoves.ListCount
   N$ = Str$(k + 1 + GameOffset)
   If k + 1 + GameOffset > 9 Then N$ = Trim$(N$) ' > 99 not dealt with
   ListMoves.AddItem N$ & " " & LabMvString  ' eg WP e2-e4 P-K4
   If HalfMove < 10 Then
      OpenString$ = OpenString$ & Mid$(LabMvString, 4, 5) & " "
   End If
   HalfMove = ListMoves.NewIndex
   ListMoves.Selected(HalfMove) = True
   If k > 10 Then ' k = ListMoves.ListCount
      ListMoves.TopIndex = ListMoves.TopIndex + 1
   End If
End Sub

Private Sub ResetForIllegalMove(Source As Control)
   ' Illegal or off board
   ' Reset to original position
   Source.Move IMLeft, IMTop
   Source.Visible = True
   'Mouse Up to clear off-board piece image
   mouse_event &H4, 0, 0, 0, 0
   MoveString$ = ""
   If ListMoves.ListCount > 0 Then
      MoveString$ = ListMoves.List(HalfMove)
   End If
   LabMvString = Mid$(MoveString$, 3)
End Sub

Private Sub ShowPieceTaken(PieceTaken$)
   
   'If Left$(IM(i).Tag, 2) <> "" Then   ' Taking a piece
   If PieceTaken$ <> "" Then   ' Taking a piece
      With frmPieces
         If Left$(PieceTaken$, 1) = "W" Then
            IMWSP(WPTaken + 16).Picture = LoadPicture
            Select Case PieceTaken$ ' Which W PieceTaken?
            Case "WR": IMWSP(WPTaken + 16).Picture = .IMW(1).Picture: ippt = WRn
            Case "WN": IMWSP(WPTaken + 16).Picture = .IMW(2).Picture: ippt = WNn
            Case "WB": IMWSP(WPTaken + 16).Picture = .IMW(3).Picture: ippt = WBn
            Case "WQ": IMWSP(WPTaken + 16).Picture = .IMW(4).Picture: ippt = WQn
            Case "WP": IMWSP(WPTaken + 16).Picture = .IMW(6).Picture: ippt = WPn
            End Select
            WPTaken = WPTaken + 1
            If WPTaken > 15 Then WPTaken = 0
         
         Else  ' BPiece
            
            IMWSP(BPTaken).Picture = LoadPicture
            Select Case PieceTaken$ ' Which B PieceTaken?
            Case "BR": IMWSP(BPTaken).Picture = .IMW(7).Picture: ippt = BRn
            Case "BN": IMWSP(BPTaken).Picture = .IMW(8).Picture: ippt = BNn
            Case "BB": IMWSP(BPTaken).Picture = .IMW(9).Picture: ippt = BBn
            Case "BQ": IMWSP(BPTaken).Picture = .IMW(10).Picture: ippt = BQn
            Case "BP": IMWSP(BPTaken).Picture = .IMW(12).Picture: ippt = BPn
            End Select
            BPTaken = BPTaken + 1
            If BPTaken > 15 Then BPTaken = 0
         End If
      End With
   End If
End Sub

Private Function TestPromotion(i As Integer) As Boolean
'Publc ippt As Long ' = piece num promoted or taken
Dim a$
   TestPromotion = False
   a$ = IM(i).Tag
   ippt = -1  ' To skip when not promotion
   If Left$(IM(i).Tag, 2) = "WP" Then  ' WP = WQ,WR,WB or WN ?
      If WEPP(WEPProm, 0) = 1 Then ' IMIndex -> i aWPawnPromotion(0) = True
         ' Test if WK attacked by Black piece
         If RC_Targetted(rwk, cwk, "B", 1) Then
            'Message$ = "White King put in check!"
            WEPP(WEPProm, 0) = 0 ' aWPawnPromotion(0) = False
            Exit Function
         End If
         Message$ = "Select promtion piece!"
         Message$ = "Select promtion piece!"
         LabMessage = Message$
         LabMessage.Refresh
'##########################################
' fraWPP
         Disabler
         PromPiece$ = ""
         fraWPP.Visible = True
         Play 8
         Do
            If PromPiece$ <> "" Then
               Select Case PromPiece$
               Case "WR": ippt = WRn ': NumWR = NumWR + 1
               Case "WN": ippt = WNn ': NumWN = NumWN + 1
               Case "WB": ippt = WBn ': NumWB = NumWB + 1
               Case "WQ": ippt = WQn ': NumWQ = NumWQ + 1
               Case Else   ' Default
                  ippt = WQn
                  'NumWQ = NumWQ + 1
               End Select
               WEPP(WEPProm, 0) = 0  ' aWPawnPromotion(0) = False
               LabMvString = LabMvString & "=" & PromPiece$
               fraWPP.Visible = False
               Enabler
               LabMessage = ""
               LabMessage.Refresh
               Exit Do
            End If
            DoEvents
         Loop
      End If
   
   ElseIf Left$(IM(i).Tag, 2) = "BP" Then ' BP = BQ,BR,BB or BN ?
      If BEPP(BEPProm, 0) = 1 Then   ' IMIndex -> i aBPawnPromotion(0) = True
         ' Test if BK attacked by White piece
         If RC_Targetted(rbk, cbk, "W", 1) Then
            'Message$ = "Black King put in check!"
            BEPP(BEPProm, 0) = 0 ' aBPawnPromotion(0) = False
            Exit Function
         End If
         Message$ = "Select promtion piece!"
         LabMessage = Message$
         LabMessage.Refresh
'###############################################
' fraBPP
         Play 8
         
         Disabler
         PromPiece$ = ""
         fraBPP.Visible = True
         Do
            If PromPiece$ <> "" Then
               Select Case PromPiece$
               Case "BR": ippt = BRn ': NumBR = NumBR + 1
               Case "BN": ippt = BNn ': NumBN = NumBN + 1
               Case "BB": ippt = BBn ': NumBB = NumBB + 1
               Case "BQ": ippt = BQn ': NumBQ = NumBQ + 1
               Case Else   ' Default
                  ippt = BQn
                  'NumBQ = NumBQ + 1
               End Select
               BEPP(BEPProm, 0) = 0  ' aBPawnPromotion(0) = False
               LabMvString = LabMvString & "=" & PromPiece$
               fraBPP.Visible = False
               Enabler
               LabMessage = ""
               LabMessage.Refresh
               Exit Do
            End If
            DoEvents
         Loop
      End If
   End If
         
   If ippt <> -1 Then
   If ippt <> 0 Then
      ' Promote WP or BP
      IM(i).Picture = IMO(ippt).Picture
      IM(i).Tag = IMO(ippt).Tag
      IM(i).DragIcon = IMO(ippt).DragIcon
      IM(i).Visible = True
   End If
   End If
   TestPromotion = True
End Function

Private Sub IMBPP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      Select Case Index
      Case 1: PromPiece$ = "BR"
      Case 2: PromPiece$ = "BN"
      Case 3: PromPiece$ = "BB"
      Case 4: PromPiece$ = "BQ"
      End Select
   End If
End Sub

Private Sub IMWPP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      Select Case Index
      Case 1: PromPiece$ = "WR"
      Case 2: PromPiece$ = "WN"
      Case 3: PromPiece$ = "WB"
      Case 4: PromPiece$ = "WQ"
      End Select
   End If
End Sub

Private Sub DoCastling(i As Integer)
   If Left$(IM(i).Tag, 1) = "W" Then   ' 61,63
      If Cast(WKSCastOK, 0) Then       ' W king side castling
         IM(61).Picture = IM(63).Picture
         IM(61).Tag = IM(63).Tag
         IM(61).DragIcon = IM(63).DragIcon
         IM(61).Visible = True
         ' Clear WKR position
         IM(63).Picture = LoadPicture
         IM(63).Tag = ""
         IM(63).DragIcon = LoadPicture
         WKRR(WKMoved, 0) = 2          ' WK moved
         Cast(WKSCastOK, 0) = 0        ' W KSideCastling = False
      ElseIf Cast(WQSCastOK, 0) Then   ' W queen side castling
         IM(59).Picture = IM(56).Picture
         IM(59).Tag = IM(56).Tag
         IM(59).DragIcon = IM(56).DragIcon
         IM(59).Visible = True
         ' Clear QKR position
         IM(56).Picture = LoadPicture
         IM(56).Tag = ""
         IM(56).DragIcon = LoadPicture
         WKRR(WKMoved, 0) = 3          ' WK moved
         Cast(WQSCastOK, 0) = 0        ' W QSideCastling = False
      End If
   
   ElseIf Left$(IM(i).Tag, 1) = "B" Then
      If Cast(BKSCastOK, 0) Then       ' B king side castling   ' 5,7
         IM(5).Picture = IM(7).Picture
         IM(5).Tag = IM(7).Tag
         IM(5).DragIcon = IM(7).DragIcon
         IM(5).Visible = True
         ' Clear BKR position
         IM(7).Picture = LoadPicture
         IM(7).Tag = ""
         IM(7).DragIcon = LoadPicture
         'IM(7).Visible = False
         BKRR(BKMoved, 0) = 2          ' BK moved
         Cast(BKSCastOK, 0) = 0        ' B KSideCastling = False
      ElseIf Cast(BQSCastOK, 0) Then   ' B queen side castling
         IM(3).Picture = IM(0).Picture
         IM(3).Tag = IM(0).Tag
         IM(3).DragIcon = IM(0).DragIcon
         IM(3).Visible = True
         ' Clear QKR position
         IM(0).Picture = LoadPicture
         IM(0).Tag = ""
         IM(0).DragIcon = LoadPicture
         BKRR(BKMoved, 0) = 3          ' BK moved
         Cast(BQSCastOK, 0) = 0        ' B QSideCastling = False
      End If
   End If
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
If aBusy Then Exit Sub
   If Screen.MousePointer <> vbDefault Then Exit Sub
   If Not aSetUp Then
      ' Piece moved off board
      aDraggedOut = True
      picBoard_DragDrop Source, x, y
   Else
   End If
End Sub

Private Sub Form_Click()
If aBusy Then Exit Sub
   If Screen.MousePointer <> vbDefault Then Exit Sub
End Sub

Private Sub Form_DblClick()
If aBusy Then Exit Sub
   If Screen.MousePointer <> vbDefault Then Exit Sub
End Sub

Private Sub ListMoves_DblClick()
If aBusy Then Exit Sub
Dim G As Long
Dim k As Long
   G = ListMoves.ListIndex  ' = HalfMove
   CopyMemory bRCBoard(1, 1, 0), BeginBoard(1, 1), 64
   SetStartUp
   Reset_ALL_BOOLS   ' to 0 = False
   ClearTakenPieces Form1
   RestoreSavedPosition Form1, 0 ' Restore display from Board 0
   If ListMoves.ListCount > 0 Then
      For k = ListMoves.ListCount - 1 To G + 1 Step -1
         ListMoves.RemoveItem k
      Next k
      HalfMove = G
      For k = 0 To HalfMove 'To ListMoves.ListCount - 1
         StepThruGame k
         SavePosition Form1, 0
         ListMoves.Selected(k) = True
         Sleep 50
      Next k
      ListMoves.Selected(ListMoves.ListCount - 1) = True
      Play 0
   End If
   
   If Left$(LabMvString, 1) = "W" Then
      If WorBsMove$ <> "A" Then mnuMoves_Click 1 ' Switch to black
      If CorHsMove$ = "CB" Then
         NumRepPositions = 1
         CheckForCompMove 1
      End If
   ElseIf Left$(LabMvString, 1) = "B" Then
      If WorBsMove$ <> "A" Then mnuMoves_Click 0 ' Switch to white
      If CorHsMove$ = "CW" Then
         NumRepPositions = 1
         CheckForCompMove 1
      End If
   End If
End Sub

Private Sub cmdStepThruGame_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If aBusy Then Exit Sub
Dim k As Long
Dim svStepGame As Long
Dim svWorBsMove$
   
   svWorBsMove$ = WorBsMove$
   mnuMoves_Click 2  ' WorBsMove$ = "A" HUMAN HUMAN

   Select Case Index
   Case 0      ' >
      If HalfMove >= 0 And HalfMove <= ListMoves.ListCount - 1 Then
         ListMoves.Selected(HalfMove) = True
         LabMvString = ListMoves.List(HalfMove)
         StepThruGame HalfMove
         SavePosition Form1, 0
         Sleep 100
         Play 0
         If HalfMove < ListMoves.ListCount - 1 Then
            HalfMove = HalfMove + 1
         End If
      ElseIf HalfMove < 0 And ListMoves.ListCount > 0 Then
            HalfMove = 0
            Reset_ALL_BOOLS   ' to 0 = False
            ListMoves.Selected(0) = True
            LabMvString = ListMoves.List(0)
            StepThruGame HalfMove
            Play 0
            SavePosition Form1, 0
         If HalfMove < ListMoves.ListCount - 1 Then
            HalfMove = HalfMove + 1
         End If
      End If
      
   Case 1      '>|
      If HalfMove < 0 And ListMoves.ListCount > 0 Then
         HalfMove = 0
         Reset_ALL_BOOLS   ' to 0 = False
      End If
      
      If HalfMove >= 0 And HalfMove < ListMoves.ListCount - 1 Then
         For k = HalfMove To ListMoves.ListCount - 1
            StepThruGame k
            SavePosition Form1, 0
            ListMoves.Selected(k) = True
            Sleep 50
         Next k
         ListMoves.Selected(ListMoves.ListCount - 1) = True
         HalfMove = ListMoves.ListCount - 1
         Play 0
      End If
   
   Case 2   ' |<
         HalfMove = 0
         CopyMemory bRCBoard(1, 1, 0), BeginBoard(1, 1), 64
         RestoreBeginBoard Form1 ' KeepLIst
         SetStartUp
         Reset_ALL_BOOLS   ' to 0 = False
         ClearTakenPieces Form1
         RestoreSavedPosition Form1, 0 ' Restore display from Board 0
         
         If ListMoves.ListCount > 0 Then
            If HalfMove >= 0 Then
               ListMoves.Selected(HalfMove) = True
            End If
            LabMvString = ListMoves.List(0)
            ListMoves.Selected(0) = True
            ListMoves.Selected(0) = False
         End If
         SavePosition Form1, 0
         HalfMove = -1
         
         Play 0
         
   Case 3   ' <  Step back 1 But uses Step forward from begining
            '    to HalfMove-1
      If HalfMove > 0 Then
         svStepGame = HalfMove
         CopyMemory bRCBoard(1, 1, 0), BeginBoard(1, 1), 64
         RestoreSavedPosition Form1, 0 ' Restore display from Board 0
         SetStartUp
         ClearTakenPieces Form1
         HalfMove = svStepGame
         HalfMove = HalfMove - 1
         Reset_ALL_BOOLS   ' to 0 = False
         
         StepThruBoard HalfMove, 0  ' Step thru from beginning on Board 0
         
         RestoreSavedPosition Form1, 0 ' Restore display from Board 0
         SavePosition Form1, 0   ' Display to Board 0
         If ListMoves.ListCount > 0 Then
            If HalfMove >= 0 Then
               ListMoves.Selected(HalfMove) = True
               LabMvString = ListMoves.List(HalfMove)
            Else
               ListMoves.Selected(0) = True
               LabMvString = ListMoves.List(0)
            End If
         End If
         DoEvents
         Play 0
      Else  ' HalfMove <= 0
         CopyMemory bRCBoard(1, 1, 0), BeginBoard(1, 1), 64
         RestoreBeginBoard Form1 ' KeepLIst
         optSetUp_Click 3  ' Play
         If ListMoves.ListCount > 0 Then
            ListMoves.Selected(0) = False
         End If
         HalfMove = -1
         DoEvents
         Play 0
      End If
   End Select
   CountPieces 0
   LabPC(0) = " B pieces =" & Str$(BlackPieceCount)
   LabPC(1) = " W pieces =" & Str$(WhitePieceCount)
   WorBsMove$ = svWorBsMove$
   mnuMoves_Click 2  ' WorBsMove$ = "A" HUMAN HUMAN
End Sub

Public Sub Reset_ALL_BOOLS()
   ReDim WKRR(0 To 3, 0 To MaxIndex)
   ReDim WEPP(0 To 3, 0 To MaxIndex)
   ReDim BKRR(0 To 3, 0 To MaxIndex)
   ReDim BEPP(0 To 3, 0 To MaxIndex)
   ReDim Cast(0 To 3, 0 To MaxIndex)
End Sub

Public Sub StepThruBoard(s As Long, Index As Integer)
'Called when  <  Step forward from begining ie

'         For k = 0 To HalfMove - 1  ' on Board 0
'            StepThruBoard k, 0
'         Next k

' S = ListMove index (0 to HalfMove)
Dim k As Long
Dim P As Long
Dim a$, Piece$, Locat$
Dim EndPiece$
Dim RowS As Long, ColS As Long        ' Row,Col start
Dim RowE As Long, ColE As Long        ' Row,Col end
Dim PN As Long, DPN As Long

ReDim MOV$(0 To s)
   For k = 0 To s
      MOV$(k) = ListMoves.List(k)
   Next k
   
On Error Resume Next
   
   For k = 0 To s
   
      a$ = MOV$(k) 'ListMoves.List(k)  ' eg s1 WP e2-e4
      DoEvents
      
      a$ = Trim$(a$)                ' remove leading space for move numbers 1-9
      P = InStr(1, a$, " ") + 1     ' -> Piece descrip
      a$ = Mid$(a$, P)                     ' eg a$ = WP e2-e4
      Piece$ = UCase$(Left$(a$, 2))        ' eg Piece$ = WP
      ' Get start square
      Locat$ = LCase$(Mid$(a$, 4, 2))      ' eg Locat$ = e2  anything after 5th char ignored
      ColS = (Asc(Left$(Locat$, 1)) - 96)  ' eg 5
      RowS = Val((Right$(Locat$, 1)))      ' eg 2
      PN = bRCBoard(RowS, ColS, Index)     ' eg 6  Piece number
      ' Get landing quare
      Locat$ = LCase$(Mid$(a$, 7, 2))      ' eg Locat$ = e4   end postion
      ColE = (Asc(Left$(Locat$, 1)) - 96)  ' eg 5
      RowE = Val((Right$(Locat$, 1)))      ' eg 4
      DPN = bRCBoard(RowE, ColE, Index)    ' eg 0  Destination piece number
   
      ' Promotion ?
      If InStr(1, a$, "=") <> 0 Then
         PromPiece$ = Right$(a$, 2)
         ConvPNDescriptoPN PromPiece$, PN
      Else
         ConvPNtoPNDescrip DPN, EndPiece$  ' eg 0 = ""
      End If
   
      ' Make the move
      bRCBoard(RowE, ColE, Index) = PN
      bRCBoard(RowS, ColS, Index) = 0
      
      Select Case PN
      Case Not WPn, Not BPn
         ReDim WPawn(1 To 8, 0 To 3)
         ReDim BPawn(1 To 8, 0 To 3)
      Case WRn
         If ColS = 8 Then WKRR(1, 0) = 1   ' WKR
         If ColS = 1 Then WKRR(2, 0) = 1   ' WQR
      Case BRn
         If ColS = 8 Then BKRR(1, 0) = 1   ' BKR
         If ColS = 1 Then BKRR(2, 0) = 1   ' BQR
      Case WKn
         WKRR(WKMoved, Index) = 1 ' aWKMoved(0) = True
         If ColE - ColS = 2 Then ' W kingside castling
            WKRR(WKMoved, Index) = 2 ' aWKMoved(0) = True
            bRCBoard(1, 6, Index) = WRn
            bRCBoard(1, 8, Index) = 0
         ElseIf ColS - ColE = 2 Then   ' W queenside castling
            WKRR(WKMoved, Index) = 3 ' aWKMoved(0) = True
            bRCBoard(1, 4, Index) = WRn
            bRCBoard(1, 1, Index) = 0
         End If
      Case BKn
         BKRR(BKMoved, Index) = 1 ' aBKMoved(0) = True
         If ColE - ColS = 2 Then ' B kingside castling
            BKRR(BKMoved, Index) = 2 ' aBKMoved(0) = True
            bRCBoard(8, 6, Index) = BRn
            bRCBoard(8, 8, Index) = 0
         ElseIf ColS - ColE = 2 Then   ' W queenside castling
            BKRR(BKMoved, Index) = 3 ' aBKMoved(0) = True
            bRCBoard(8, 4, Index) = BRn
            bRCBoard(8, 1, Index) = 0
         End If
      Case WPn
            bRCBoard(RowS, ColS, Index) = 0
            'DPN = 0
            If ColS <> ColE Then    ' Diagonal, take or en passant
               If BPawn(ColE, Index) = 2 And DPN = 0 Then ' B 2 down sets it
                  ' DPN will be = 0
                  bRCBoard(RowS, ColE, Index) = 0
                  EndPiece$ = "BP"
                  ReDim BPawn(1 To 8, 0 To 3)
                  ReDim WPawn(1 To 8, 0 To 3)
               Else  ' A take
                  ConvPNtoPNDescrip DPN, EndPiece$
               End If
            Else  ' ColS=ColE
               If RowS = 2 And RowE = 4 Then
                  ReDim WPawn(1 To 8, 0 To 3)
                  ReDim BPawn(1 To 8, 0 To 3)
                  WPawn(ColS, 0) = 2
               Else
                  ReDim WPawn(1 To 8, 0 To 3)
                  ReDim BPawn(1 To 8, 0 To 3)
               End If
            End If
      Case BPn
            bRCBoard(RowS, ColS, Index) = 0
            'DPN = 0
            If ColS <> ColE Then    ' Diagonal, take or en passant
               If WPawn(ColE, Index) = 2 And DPN = 0 Then ' W 2 forward sets it
                  ' DPN will be = 0
                  bRCBoard(RowS, ColE, Index) = 0
                  EndPiece$ = "WP"
                  ReDim BPawn(1 To 8, 0 To 3)
                  ReDim WPawn(1 To 8, 0 To 3)
               Else  ' A take
                  ConvPNtoPNDescrip DPN, EndPiece$
               End If
            Else  ' ColS=ColE
               If RowS = 7 And RowE = 5 Then
                  ReDim WPawn(1 To 8, 0 To 3)
                  ReDim BPawn(1 To 8, 0 To 3)
                  BPawn(ColS, Index) = 2
               Else
                  ReDim WPawn(1 To 8, 0 To 3)
                  ReDim BPawn(1 To 8, 0 To 3)
               End If
            End If
      End Select
   
      If InStr(1, a$, "=") = 0 Then   ' Not Promotion
         If EndPiece$ <> "" Then
            ShowPieceTaken EndPiece$ '( PieceTaken$)
         End If
      Else  ' Promotion
      End If
      Refresh
      LabMvString = a$  ' Last move
   Next k
   If Left$(LabMvString, 1) = "W" Then
      If WorBsMove$ = "W" Then mnuMoves_Click 1 ' Switch to black
   ElseIf Left$(LabMvString, 1) = "B" Then
      If WorBsMove$ = "B" Then mnuMoves_Click 0 ' Switch to white
   End If
On Error GoTo 0
End Sub

Public Sub StepThruGame(s As Long)
' Called when <, >|, >
' S = ListMove index (0 to HalfMove)
Dim P As Long
Dim a$, b$, Piece$, Locat$
Dim EndPiece$
Dim StartIndex As Long, EndIndex As Integer
Dim RowS As Long, ColS As Long        ' Row,Col start
Dim RowE As Long, ColE As Long        ' Row,Col end
   On Error Resume Next
   
   a$ = ListMoves.List(s)
   b$ = a$
   
   a$ = Trim$(a$)                      ' remove leading space for move numbers 1-9
   P = InStr(1, a$, " ") + 1           ' -> Piece descrip
   a$ = Mid$(a$, P)
   Piece$ = UCase$(Left$(a$, 2))       ' eg WP e2-e4 RowS=2 ColS=5, RowE=4 ColE=5
   
   Locat$ = LCase$(Mid$(a$, 4, 2))     ' End position anything after 5th char ignored
   ColS = (Asc(Left$(Locat$, 1)) - 96) ' Start loc
   RowS = Val((Right$(Locat$, 1)))
   StartIndex = 64 - ((9 - ColS) + 8 * (RowS - 1))  ' IM start index
   
   Locat$ = LCase$(Mid$(a$, 7, 2))     ' End position
   ColE = (Asc(Left$(Locat$, 1)) - 96) ' End loc
   RowE = Val((Right$(Locat$, 1)))
   EndIndex = 64 - ((9 - ColE) + 8 * (RowE - 1))    ' IM end index
   
   EndPiece$ = IM(EndIndex).Tag
   
   If Left$(EndPiece$, 1) = Left$(Piece$, 1) Then Exit Sub
   
   IM(StartIndex).Picture = LoadPicture
   IM(StartIndex).DragIcon = LoadPicture
   IM(StartIndex).Tag = ""
      
    
   ' Promotion ?
   If InStr(1, a$, "=") <> 0 Then
      Piece$ = Right$(a$, 2)
      EndPiece$ = Right$(a$, 2)
   Else
      EndPiece$ = IM(EndIndex).Tag
   End If
   
   Select Case Piece$
   Case "WR": IMOIndex = 1
   Case "WN": IMOIndex = 2
   Case "WB": IMOIndex = 3
   Case "WQ": IMOIndex = 4
   Case "WK": IMOIndex = 5
   Case "WP": IMOIndex = 6

   Case "BR": IMOIndex = 7
   Case "BN": IMOIndex = 8
   Case "BB": IMOIndex = 9
   Case "BQ": IMOIndex = 10
   Case "BK": IMOIndex = 11
   Case "BP": IMOIndex = 12
   Case Else   ' Shouldn't come here
      Exit Sub
   End Select
   
   Select Case EndPiece$
   Case "WR": ippt = 1
   Case "WN": ippt = 2
   Case "WB": ippt = 3
   Case "WQ": ippt = 4
   Case "WK": ippt = 5
   Case "WP": ippt = 6

   Case "BR": ippt = 7
   Case "BN": ippt = 8
   Case "BB": ippt = 9
   Case "BQ": ippt = 10
   Case "BK": ippt = 11
   Case "BP": ippt = 12
   End Select
   
   
   IM(EndIndex).Tag = Piece$
   IM(EndIndex).Picture = IMO(IMOIndex).Picture
   IM(EndIndex).DragIcon = IMO(IMOIndex).DragIcon
   IM(EndIndex).Visible = True
   ' en passant ?
   ' WP @ row 5 col c, to row 6 col c+-1 a blank sqaure AND
   ' BP @ row 5 col c+-1 then en passant & blank row5 col c+-1
   
   ' BP @ row 4 col c, to row 3 col c+-1 a blank sqaure AND
   ' WP @ row 4 col c+-1 then en passant & blank row5 col c+-1
    
   If EndPiece$ = "" Then
   
      If Piece$ = "WP" Then
         If RowS = 5 And RowE = 6 Then
            If ColE = ColS - 1 Then    ' Take or enpassant
               If IM(StartIndex - 1).Tag = "BP" Then  ' en passant tl
                  IM(StartIndex - 1).Picture = LoadPicture
                  IM(StartIndex - 1).DragIcon = LoadPicture
                  IM(StartIndex - 1).Tag = ""
               End If
               WPawn(ColS, 0) = 0
            ElseIf ColE = ColS + 1 Then   ' Take or enpassant
               If IM(StartIndex + 1).Tag = "BP" Then  ' en passant tr
                  IM(StartIndex + 1).Picture = LoadPicture
                  IM(StartIndex + 1).DragIcon = LoadPicture
                  IM(StartIndex + 1).Tag = ""
               End If
               WPawn(ColS, 0) = 0
            End If
         ElseIf RowS = 2 And RowE = 4 Then
            WPawn(ColS, 0) = 2
         Else
            WPawn(ColS, 0) = 0
         End If
      
      ElseIf Piece$ = "BP" Then
         If RowS = 4 And RowE = 3 Then
            If ColE = ColS - 1 Then    ' Take or enpassant
               If IM(StartIndex - 1).Tag = "WP" Then  ' en passant bl
                  IM(StartIndex - 1).Picture = LoadPicture
                  IM(StartIndex - 1).DragIcon = LoadPicture
                  IM(StartIndex - 1).Tag = ""
               End If
               BPawn(ColS, 0) = 0
            ElseIf ColE = ColS + 1 Then   ' Take or enpassant
               If IM(StartIndex + 1).Tag = "WP" Then   ' en passant br
                  IM(StartIndex + 1).Picture = LoadPicture
                  IM(StartIndex + 1).DragIcon = LoadPicture
                  IM(StartIndex + 1).Tag = ""
               End If
               BPawn(ColS, 0) = 0
            End If
         ElseIf RowS = 7 And RowE = 5 Then
            BPawn(ColS, 0) = 2
         Else
            BPawn(ColS, 0) = 0
         End If
      End If
   End If
   
   ' Castling
   
   If (Piece$ = "WR" Or Piece$ = "BR") Then
      If Piece$ = "WR" Then
         If RowS = 1 And ColS = 8 Then WKRR(1, 0) = 1   ' WKR
         If RowS = 1 And ColS = 1 Then WKRR(2, 0) = 1   ' WQR
      Else  ' "BR   "
         If RowS = 8 And ColS = 8 Then BKRR(1, 0) = 1   ' BKR
         If RowS = 8 And ColS = 1 Then BKRR(2, 0) = 1   ' BQR
      End If
   End If
   
   If (Piece$ = "WK" Or Piece$ = "BK") Then
      If Piece$ = "WK" Then WKRR(0, 0) = 1
      If Piece$ = "BK" Then BKRR(0, 0) = 1
      If Abs(EndIndex - StartIndex) = 2 Then
         If EndIndex > StartIndex Then  ' KingSide castling
            IM(EndIndex + 1).Picture = LoadPicture
            IM(EndIndex + 1).DragIcon = LoadPicture
            IM(EndIndex + 1).Tag = ""
            If Piece$ = "WK" Then
               IM(EndIndex - 1).Tag = "WR"
               IM(EndIndex - 1).Picture = IMO(1).Picture
               IM(EndIndex - 1).DragIcon = IMO(1).DragIcon
               WKRR(WKMoved, 0) = 2
            Else  ' Piece$ = "BK"
               IM(EndIndex - 1).Tag = "BR"
               IM(EndIndex - 1).Picture = IMO(7).Picture
               IM(EndIndex - 1).DragIcon = IMO(7).DragIcon
               BKRR(BKMoved, 0) = 2
            End If
            IM(EndIndex - 1).Visible = True
         Else  ' QueenSide castling
            IM(EndIndex - 2).Picture = LoadPicture
            IM(EndIndex - 2).DragIcon = LoadPicture
            IM(EndIndex - 2).Tag = ""
            If Piece$ = "WK" Then
               IM(EndIndex + 1).Tag = "WR"
               IM(EndIndex + 1).Picture = IMO(1).Picture
               IM(EndIndex + 1).DragIcon = IMO(1).DragIcon
               WKRR(WKMoved, 0) = 3
            Else  ' Piece$ = "BK"
               IM(EndIndex + 1).Tag = "BR"
               IM(EndIndex + 1).Picture = IMO(7).Picture
               IM(EndIndex + 1).DragIcon = IMO(7).DragIcon
               BKRR(BKMoved, 0) = 3
            End If
            IM(EndIndex + 1).Visible = True
            Refresh
         End If '''
      End If
   End If
   
   If InStr(1, a$, "=") = 0 Then   ' Not Promotion
      If EndPiece$ <> "" Then
         ShowPieceTaken EndPiece$ '( PieceTaken$)
      End If
   Else  ' Promotion
   End If
   Refresh
   
   LabMvString = a$  ' Last move
   If Left$(LabMvString, 1) = "W" Then
      If WorBsMove$ <> "A" Then mnuMoves_Click 1 ' Switch to black
   ElseIf Left$(LabMvString, 1) = "B" Then
      If WorBsMove$ <> "A" Then mnuMoves_Click 0 ' Switch to white
   End If
   On Error GoTo 0
' Called when <, >|, >
End Sub

Private Sub SwapPieceList(aTF As Boolean)
   fraPieces.Visible = aTF       'True/False
   fraSwap.Visible = Not aTF
End Sub


Private Sub mnuHelp_Click()
   App.HelpFile = PathSpec$ & "RRChess.chm"
   SendKeys "{f1}"
End Sub

Private Sub mnuPrintCB_Click(Index As Integer)
' Print Clipboard
   ' SendKeys "%{PRTSC}", True   ' DOESN'T WORK
   ' Me.PrintForm    ' DOESN'T WORK
Dim res As Long
   On Error GoTo PCBError
   
   res = MsgBox("NB. If ALT-PrintScrn not done click No  " & vbCrLf & "IS PRINTER LIVE!", vbQuestion + vbYesNo, "Print Clipboard")
   If res = vbNo Then Exit Sub
   
   picCB.Picture = LoadPicture
   picCB.Picture = Clipboard.GetData(2)
   picCB.Refresh
   Clipboard.Clear
   Printer.PaintPicture picCB.Picture, 50, 50
   Printer.EndDoc
PCBError:
   Printer.KillDoc
   picCB.Picture = LoadPicture
   picCB.Width = 4
   picCB.Height = 4
   picBoard.SetFocus
   On Error GoTo 0
End Sub

Private Sub TileForm1()
Dim ix As Long, iy As Long
Dim TileSize As Long
   TileSize = 64
   For iy = 0 To Me.Height / STY Step TileSize
   For ix = 0 To Me.Width / STX Step TileSize
      BitBlt Me.hDC, ix, iy, TileSize, TileSize, Me.hDC, 0, 0, vbSrcCopy
   Next ix
   Next iy
   Me.Refresh
End Sub

Private Sub mnuExit_Click()
   Form_Unload 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim fnum As Long
   aBusy = True
   aExit = True
   StopPlay
   Printer.KillDoc
   fnum = FreeFile
   Open PathSpec$ & "RRChessInfo.txt" For Output As #fnum
   Write #fnum, aSound         ' #TRUE# or #FALSE#
   Close #fnum
   
   Call Unload(frmPieces)
   Set frmPieces = Nothing
   Unload Form1
   End
End Sub

