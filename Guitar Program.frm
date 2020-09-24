VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wilksey's Guitar Chord Finder Program - ©2002 Wilksey - Licensed To: Release Version 2.0"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "Guitar Program.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCount 
      Enabled         =   0   'False
      Left            =   5040
      Top             =   2880
   End
   Begin VB.CheckBox chkStrummed 
      Caption         =   "Strummed"
      Height          =   255
      Left            =   5625
      TabIndex        =   11
      Top             =   2235
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Chord"
      Height          =   375
      Left            =   4425
      TabIndex        =   10
      Top             =   2190
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   2175
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   2175
      Width           =   855
   End
   Begin VB.ComboBox cboChord 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   21
      Left            =   8445
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   20
      Left            =   8055
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   19
      Left            =   7665
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   18
      Left            =   7290
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   17
      Left            =   6915
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   16
      Left            =   6525
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   15
      Left            =   6150
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   14
      Left            =   5775
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   13
      Left            =   5385
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   12
      Left            =   4995
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   21
      Left            =   8445
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   20
      Left            =   8055
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   19
      Left            =   7680
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   18
      Left            =   7290
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   17
      Left            =   6915
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   16
      Left            =   6525
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   15
      Left            =   6150
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   14
      Left            =   5775
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   13
      Left            =   5385
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   12
      Left            =   4995
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   21
      Left            =   8445
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   20
      Left            =   8055
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   19
      Left            =   7680
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   18
      Left            =   7290
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   17
      Left            =   6915
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   16
      Left            =   6525
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   15
      Left            =   6150
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   14
      Left            =   5775
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   13
      Left            =   5400
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   12
      Left            =   5010
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   21
      Left            =   8445
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   20
      Left            =   8055
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   19
      Left            =   7680
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   18
      Left            =   7305
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   17
      Left            =   6915
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   16
      Left            =   6540
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   15
      Left            =   6150
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   14
      Left            =   5760
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   13
      Left            =   5385
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   12
      Left            =   5010
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   21
      Left            =   8445
      Top             =   1365
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   20
      Left            =   8055
      Top             =   1365
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   19
      Left            =   7665
      Top             =   1365
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   18
      Left            =   7290
      Top             =   1365
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   17
      Left            =   6930
      Top             =   1380
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   16
      Left            =   6525
      Top             =   1380
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   15
      Left            =   6150
      Top             =   1380
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   14
      Left            =   5775
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   13
      Left            =   5400
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   12
      Left            =   5010
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   21
      Left            =   8445
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   20
      Left            =   8055
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   19
      Left            =   7680
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   18
      Left            =   7305
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   17
      Left            =   6915
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   16
      Left            =   6540
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   15
      Left            =   6165
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   14
      Left            =   5775
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   13
      Left            =   5400
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   12
      Left            =   5010
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblPlayChord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Play Chord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      TabIndex        =   12
      Top             =   2190
      Width           =   1140
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   11
      Left            =   4620
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   10
      Left            =   4230
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   9
      Left            =   3855
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   8
      Left            =   3480
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   7
      Left            =   3105
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   6
      Left            =   2715
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   5
      Left            =   2340
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   11
      Left            =   4620
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   10
      Left            =   4230
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   9
      Left            =   3855
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   8
      Left            =   3480
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   7
      Left            =   3105
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   6
      Left            =   2715
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   5
      Left            =   2340
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   11
      Left            =   4620
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   10
      Left            =   4230
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   9
      Left            =   3855
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   8
      Left            =   3465
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   7
      Left            =   3105
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   6
      Left            =   2730
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   5
      Left            =   2340
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   11
      Left            =   4620
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   10
      Left            =   4245
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   9
      Left            =   3855
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   8
      Left            =   3465
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   7
      Left            =   3105
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   6
      Left            =   2715
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   5
      Left            =   2340
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   11
      Left            =   4620
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   10
      Left            =   4245
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   9
      Left            =   3855
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   8
      Left            =   3480
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   7
      Left            =   3105
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   6
      Left            =   2715
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   5
      Left            =   2340
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   11
      Left            =   4635
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   10
      Left            =   4245
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   9
      Left            =   3855
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   8
      Left            =   3480
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   7
      Left            =   3105
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   6
      Left            =   2730
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   5
      Left            =   2340
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblShowChord 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Chord:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   9
      Top             =   2205
      Width           =   1365
   End
   Begin VB.Image imgMiddle 
      Height          =   405
      Left            =   4560
      Picture         =   "Guitar Program.frx":0BC2
      Top             =   2880
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgBottom 
      Height          =   210
      Left            =   4080
      Picture         =   "Guitar Program.frx":0F97
      Top             =   3000
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgTop 
      Height          =   405
      Left            =   3720
      Picture         =   "Guitar Program.frx":1371
      Top             =   2880
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgX6 
      Height          =   210
      Left            =   240
      Top             =   105
      Width           =   255
   End
   Begin VB.Image imgX5 
      Height          =   210
      Left            =   240
      Top             =   450
      Width           =   255
   End
   Begin VB.Image imgX4 
      Height          =   210
      Left            =   240
      Top             =   765
      Width           =   255
   End
   Begin VB.Image imgX3 
      Height          =   210
      Left            =   240
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image imgX2 
      Height          =   210
      Left            =   240
      Top             =   1410
      Width           =   255
   End
   Begin VB.Image imgX1 
      Height          =   210
      Left            =   240
      Top             =   1740
      Width           =   255
   End
   Begin VB.Image imgX 
      Height          =   210
      Left            =   3360
      Picture         =   "Guitar Program.frx":1775
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   420
      Width           =   105
   End
   Begin VB.Label lblG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   765
      Width           =   120
   End
   Begin VB.Label lblD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   120
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   1395
      Width           =   105
   End
   Begin VB.Label lblE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   1725
      Width           =   105
   End
   Begin VB.Label lblE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   105
      Width           =   90
   End
   Begin VB.Image imgNothing 
      Height          =   15
      Left            =   4080
      Picture         =   "Guitar Program.frx":1AF1
      Top             =   2280
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgNote 
      Height          =   210
      Left            =   3000
      Picture         =   "Guitar Program.frx":1B2E
      Top             =   3000
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   4
      Left            =   1980
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   3
      Left            =   1575
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   2
      Left            =   1200
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   1
      Left            =   825
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img6 
      Height          =   210
      Index           =   0
      Left            =   420
      Top             =   105
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   4
      Left            =   1980
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   3
      Left            =   1575
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   2
      Left            =   1200
      Top             =   420
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   1
      Left            =   825
      Top             =   435
      Width           =   255
   End
   Begin VB.Image img5 
      Height          =   210
      Index           =   0
      Left            =   420
      Top             =   435
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   4
      Left            =   1980
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   3
      Left            =   1575
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   2
      Left            =   1185
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   1
      Left            =   825
      Top             =   720
      Width           =   255
   End
   Begin VB.Image img4 
      Height          =   210
      Index           =   0
      Left            =   420
      Top             =   735
      Width           =   270
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   4
      Left            =   1980
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   3
      Left            =   1575
      Top             =   1095
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   2
      Left            =   1185
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   1
      Left            =   825
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img3 
      Height          =   210
      Index           =   0
      Left            =   420
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   4
      Left            =   1980
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   3
      Left            =   1575
      Top             =   1395
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   2
      Left            =   1185
      Top             =   1410
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   1
      Left            =   825
      Top             =   1410
      Width           =   255
   End
   Begin VB.Image img2 
      Height          =   210
      Index           =   0
      Left            =   420
      Top             =   1410
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   4
      Left            =   1980
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   3
      Left            =   1575
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   2
      Left            =   1185
      Top             =   1695
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   1
      Left            =   825
      Top             =   1695
      Width           =   255
   End
   Begin VB.Image img1 
      Height          =   210
      Index           =   0
      Left            =   420
      Top             =   1695
      Width           =   255
   End
   Begin VB.Image imgFretboard 
      Height          =   2145
      Left            =   0
      Picture         =   "Guitar Program.frx":1F1C
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Wilksey's Guitar Chord Finding Program
'©2002 Wilksey - Full Release Version 2.0
'Any questions, Commments please email me: Wilksey@Softhome.net  thank you!
'Please vote for this version fairly, I noticed a couple of unfair votes before.
'How many open source programs like this and as well commented have u seen on PSC?
'
Option Explicit
'--------Variables and Types--------
Dim Msg As String
'Chord type using standard tuning EADGBe or e2 as E has already been used
Private Type Chord
    E As String
    A As String
    D As String
    G As String
    B As String
    e2 As String
    Name As String
    Barre As Boolean
End Type
'This defines a standard Note, Notes being the note played on each string(0 to 5)=6, 6 strings, Instrument, this is the MIDI patch that is sent to the MIDI Device
Private Type Note
    Notes(5) As Integer
    Instrument As Integer
End Type
'Declare variables using our newly created types
Dim Chords() As Chord
Dim ANote() As Note
'Declaration for Database
Dim DB_DEFS As Database
Dim RS_DEFS As Recordset
'Midi Declarations
Dim Channel As Integer
Dim Volume As Integer
Dim hWndMidi As Long
Dim BaseNote As Integer
Dim MidiMsg As Long
Dim rc As Long
Dim HasMIDI As Boolean
'String Declarations
Dim StringE As Integer, StringA As Integer, StringD As Integer, StringG As Integer, StringB As Integer, StringE2 As Integer
'Timing Declarations
Dim Counter As Integer


Private Sub cboChord_Change()
'--Variables--
Dim i As Integer
'Error trapping
On Error Resume Next
'Check for Error and Empty
If cboChord.Text = "Error" Then Exit Sub Else
If cboChord.Text = "" Then cboChord.Text = cboChord.List(0): Exit Sub Else
'Start a loop between 0 and Chords max index
For i = 0 To UBound(Chords)
'If we find the chord exit the sub
If cboChord.Text = Chords(i).Name Then
    Exit Sub
End If
Next i
    
    'If we cant find the chord tell the user
    Msg = "Invalid Chord Name"
    MsgBox Msg, vbOKOnly + vbInformation, "Error..."
End Sub

Private Sub cboChord_Click()
'Call our 2 main subs for displaying the chords
    ClearChordView   'Clear the current chord display(if any)
    ShowChord (cboChord.Text)  'Show the chord by name
End Sub

Private Sub cboChord_KeyPress(KeyAscii As Integer)
Dim i As Integer
If KeyAscii = 13 Then
For i = 0 To UBound(Chords)
'If we find the chord exit the sub
If cboChord.Text = Chords(i).Name Then
    'call external sub
    Call cboChord_Click
    'exit subroutine
    Exit Sub
End If
Next i
End If
End Sub

Private Sub cmdAbout_Click()
    'A Simple message box
    Msg = "Wilksey's Guitar Chord Finder Program" & vbCrLf & "©2002 Wilksey" & vbCrLf & "Licensed To: Release Version 2.0"
    MsgBox Msg, vbOKOnly + vbInformation, "About..."
End Sub

Private Sub cmdExit_Click()
    'calls the Form_Unload sub
    Unload Me
End Sub



Private Sub cmdPlay_Click()
    'Disable self plus the checkbox, so we dont cause errors, plus we know when the MIDI has stopped!
    cmdPlay.Enabled = False
    chkStrummed.Enabled = False
    'This sends our instrument setting through to the MIDI Device telling it what PATCH to use.
    ChangeInstrument ANote(cboChord.ListIndex).Instrument
    'This gets the Notes from the chord Index
    NoteFromChordIndex cboChord.ListIndex
    'If The value is not -1( 'X') Then play the note
    If ANote(cboChord.ListIndex).Notes(0) > -1 Then
        'This sends the note to the MIDI Device
        Play NoteFromString(ANote(cboChord.ListIndex).Notes(0), 1), 50, chkStrummed.Value
    'End condition
    End If
    'Same as above, just for other strings
    If ANote(cboChord.ListIndex).Notes(1) > -1 Then
        Play NoteFromString(ANote(cboChord.ListIndex).Notes(1), 2), 50, chkStrummed.Value
    End If
    If ANote(cboChord.ListIndex).Notes(2) > -1 Then
        Play NoteFromString(ANote(cboChord.ListIndex).Notes(2), 3), 50, chkStrummed.Value
    End If
    If ANote(cboChord.ListIndex).Notes(3) > -1 Then
        Play NoteFromString(ANote(cboChord.ListIndex).Notes(3), 4), 50, chkStrummed.Value
    End If
    If ANote(cboChord.ListIndex).Notes(4) > -1 Then
        Play NoteFromString(ANote(cboChord.ListIndex).Notes(4), 5), 50, chkStrummed.Value
    End If
    If ANote(cboChord.ListIndex).Notes(5) > -1 Then
        Play NoteFromString(ANote(cboChord.ListIndex).Notes(5), 6), 50, chkStrummed.Value
    End If
        'If we have not strummed the chord
    If chkStrummed.Value = 0 Then
        'Enable Play button and Strummed Checkbox
        cmdPlay.Enabled = True
        chkStrummed.Enabled = True
    'End if condition
    End If
    'Exit from subroutine, so we dont go into the error trap below.
    Exit Sub
NoteError:
'Message box
    Msg = "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & "Source:" & "cmdPlay_Click()"
    MsgBox Msg, vbOKOnly + vbExclamation, "An Error Occured..."
    'Exit from sub
    Exit Sub
End Sub

Private Sub Form_Load()
    'Set Up Base Strings Notes
    StringE = 4
    StringA = 9
    StringD = 14
    StringG = 19
    StringB = 23
    StringE2 = 28
    'Set local variable Count to 1
    Counter = 1
    'Loads chords from database into the Chords() memory
    LoadChordsFromDatabase App.Path + "\Guitar Chord Finder.mdb"
    'Sets Up MIDI Mapper as our default MIDI Device and sets its properties
    SetupMIDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close the MIDI Device
    rc = midiOutClose(hWndMidi)
    'end application
    End
End Sub

Private Sub ClearChordView()    'This user sub is to clear the contents of the chord view
'--Variables--
Dim i As Integer
    'Sets the images of Open chord(X) to Nothing/Blank picture
    imgX1.Picture = imgNothing.Picture
    imgX2.Picture = imgNothing.Picture
    imgX3.Picture = imgNothing.Picture
    imgX4.Picture = imgNothing.Picture
    imgX5.Picture = imgNothing.Picture
    imgX6.Picture = imgNothing.Picture
    'Loops from 0 to max index of Img1
    For i = 0 To img1.UBound
        'sets img1(index) to nothing/blank
        img1(i).Picture = imgNothing.Picture
    'continue loop
    Next i
    'same as above'
    For i = 0 To img2.UBound
        img2(i).Picture = imgNothing.Picture
    Next i
    
    For i = 0 To img3.UBound
        img3(i).Picture = imgNothing.Picture
    Next i
    
    For i = 0 To img4.UBound
        img4(i).Picture = imgNothing.Picture
    Next i
    
    For i = 0 To img5.UBound
        img5(i).Picture = imgNothing.Picture
    Next i
    
    For i = 0 To img6.UBound
        img6(i).Picture = imgNothing.Picture
    Next i
    
End Sub

Private Sub ShowChord(ChordName As String)  'This is the MAIN sub to display the chord
'--Variables--
Dim i As Integer
'Error Trapping
On Error GoTo ShowChordError
    'Loop from 0 to max index of Chords
    For i = 0 To UBound(Chords)
    'If we find the chord
    If Chords(i).Name = ChordName Then
        'If the integer value of Chords(index).E property is greater than 0 or the Chords(i).E is not set to X
        If UCase$(Chords(i).E) <> "X" And Val(Chords(i).E) > 0 Then
            'Set the note image at the apropriate place.
            'Note: we -1 from the value as the img starts from 0, and we want realtime so on the diagram 0 would be fret 1, and 1 would be fret 2 etc
            img1(Val(Chords(i).E) - 1).Picture = imgNote.Picture
        'if it is 0 and chords(i).E=X
        ElseIf Val(Chords(i).E) = 0 And UCase$(Chords(i).E) = "X" Then
            'set the String to Open(X)
            imgX1.Picture = imgX.Picture
        End If
        'Same as above, except for the A property....Note: this happens six times in total: EADGBe 6 strings remember :)
        
        If UCase$(Chords(i).A) <> "X" And Val(Chords(i).A) > 0 Then
            img2(Val(Chords(i).A) - 1).Picture = imgNote.Picture
        ElseIf Val(Chords(i).A) = 0 And UCase$(Chords(i).A) = "X" Then
            imgX2.Picture = imgX.Picture
        End If
        
        If UCase$(Chords(i).D) <> "X" And Val(Chords(i).D) > 0 Then
            img3(Val(Chords(i).D) - 1).Picture = imgNote.Picture
        ElseIf Val(Chords(i).D) = 0 And UCase$(Chords(i).D) = "X" Then
            imgX3.Picture = imgX.Picture
        End If
        
        If UCase$(Chords(i).G) <> "X" And Val(Chords(i).G) > 0 Then
            img4(Val(Chords(i).G) - 1).Picture = imgNote.Picture
        ElseIf Val(Chords(i).G) = 0 And UCase$(Chords(i).G) = "X" Then
            imgX4.Picture = imgX.Picture
        End If
        
        If UCase$(Chords(i).B) <> "X" And Val(Chords(i).B) > 0 Then
            img5(Val(Chords(i).B) - 1).Picture = imgNote.Picture
        ElseIf Val(Chords(i).B) = 0 And UCase$(Chords(i).B) = "X" Then
            imgX5.Picture = imgX.Picture
        End If
        
        If UCase$(Chords(i).e2) <> "X" And Val(Chords(i).e2) > 0 Then
            img6(Val(Chords(i).e2) - 1).Picture = imgNote.Picture
        ElseIf Val(Chords(i).e2) = 0 And UCase$(Chords(i).e2) = "X" Then
            imgX6.Picture = imgX.Picture
        End If
        'This is for barre Chords
        If Chords(i).Barre = True Then
        'Checks if the E string is equal to E2(e) string and if the fret is higher than 0
        If Val(Chords(i).E) = Val(Chords(i).e2) And Val(Chords(i).E) > 0 And Val(Chords(i).e2) > 0 Then
        'Sets the images
            img6(Val(Chords(i).E) - 1).Picture = imgTop.Picture
            img1(Val(Chords(i).E) - 1).Picture = imgBottom.Picture
            img2(Val(Chords(i).E) - 1).Picture = imgMiddle.Picture
            img3(Val(Chords(i).E) - 1).Picture = imgMiddle.Picture
            img4(Val(Chords(i).e2) - 1).Picture = imgMiddle.Picture
            img5(Val(Chords(i).e2) - 1).Picture = imgMiddle.Picture
        End If
        'Checks if the A string is equal to E2(e) string and if the fret is higher than 0
        If Val(Chords(i).A) = Val(Chords(i).e2) And Val(Chords(i).A) > 0 And Val(Chords(i).e2) > 0 Then
        'Sets the images
            img6(Val(Chords(i).e2) - 1).Picture = imgTop.Picture
            img2(Val(Chords(i).A) - 1).Picture = imgBottom.Picture
            img3(Val(Chords(i).A) - 1).Picture = imgMiddle.Picture
            img4(Val(Chords(i).e2) - 1).Picture = imgMiddle.Picture
            img5(Val(Chords(i).e2) - 1).Picture = imgMiddle.Picture
        End If
        End If
        'Exit sub
        Exit Sub
    'end IF condition checking
    End If
    'Continues looping
    Next i
        'Message box
        Msg = "Sorry, Chord Not Found..."
        MsgBox Msg, vbOKOnly + vbInformation, "Cannot Locate Chord..."
        'exit from sub
        Exit Sub
ShowChordError:
    'Message box
    Msg = "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & "Source:" & "ShowChord()"
    MsgBox Msg, vbOKOnly + vbExclamation, "An Error Occured..."
    'Clear any chords previously shown...also reset the Combo index back.
    ClearChordView
    cboChord.ListIndex = 0
    'Exit from sub
    Exit Sub
End Sub


Private Sub LoadChordsFromDatabase(DatabaseFile As String)  'This user sub loads the chords from the Database sepcified
'Definitions
Dim Index As Integer
'Check if the database exists
If Check4Database(DatabaseFile) = "" Then
    'Display error box
    Msg = "Error, Cannot find database"
    MsgBox Msg, vbOKOnly + vbCritical, "Error loading database"
    'quit subroutine
    Exit Sub
End If
'Error Trapping
On Error GoTo LoadChordError
    'Open the database
    Set DB_DEFS = OpenDatabase(DatabaseFile)
    'Open table 'Definitions'
    Set RS_DEFS = DB_DEFS.OpenRecordset("Definitions")
    'Move to the first record in the recordset
    RS_DEFS.MoveFirst
    'Set variable
    Index = 0
    'Give the properties of variable to VB
    With RS_DEFS
    'Redefine the Chords variable with RecordCount, keeping the values already in the array.
    ReDim Preserve Chords(.RecordCount)
    'Redefine the Notes variable with RecordCount, keeping the values already in the array.
    ReDim Preserve ANote(.RecordCount)
    'Do until there are no more records
    Do While Not (.EOF)
        'Set variable to current record field
        Chords(Index).Name = .Fields("Chord Name").Value
        'Split the chords into single string variables and apply to array Chords()
        SplitChords .Fields("Define Chord").Value, Index
        'Set the Barre Property on the chord from the Database
        Chords(Index).Barre = .Fields("Barred").Value
        'This tells the program what instrument to use when playing, this is a index from the patches list
        ANote(Index).Instrument = .Fields("Instrument").Value
        'Add the item to the Combo box
        cboChord.AddItem Chords(Index).Name
        'Move to the next record in the database
        .MoveNext
        'Increment Index
        Index = Index + 1
    Loop
    'Release the properties of Variable
    End With
    'Close recordset
    RS_DEFS.Close
    'Close Database
    DB_DEFS.Close
    'Set the combo box to the first Chord in the list
    cboChord.Text = cboChord.List(0)
    'Sets the combobox's List Index property to 0
    cboChord.ListIndex = 0
    'Show the chord
    ShowChord cboChord.Text
    'quit from subroutine
    Exit Sub
LoadChordError:
    'Message box
    Msg = "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & "Source:" & "LoadChordsFromDatabase()"
    MsgBox Msg, vbOKOnly + vbExclamation, "An Error Occured..."
    'Set text of ComboBox to 'Error'
    cboChord.Text = "Error"
    'Exit from sub
    Exit Sub
End Sub

Private Function Check4Database(DatabaseFile As String) As String   'This user function checks if the database exists
    'Returns Filename if it is present
    Check4Database = Dir(DatabaseFile)
End Function

Private Sub SplitChords(ChordStr As String, Index As Integer)   'This user Sub splits the 'lump' chord string into seperate single strings and assigns them to the chord definition
'Declarations
Dim tmpArray() As String
    'Split string into array using delimeter ":" and length 1, and doing a text compare.
    tmpArray() = Split(ChordStr, ":", 8, vbTextCompare)
    'Set variables from array
    Chords(Index).E = tmpArray(0)
    Chords(Index).A = tmpArray(1)
    Chords(Index).D = tmpArray(2)
    Chords(Index).G = tmpArray(3)
    Chords(Index).B = tmpArray(4)
    Chords(Index).e2 = tmpArray(5)
End Sub

Private Sub Play(Note2Play As Integer, SleepTime As Integer, Strummed As Boolean)    'This user sub will play the Note through the MIDI device
'Error Trap
On Error GoTo ErrPlay
'Checks if we have a MIDI Device available
If HasMIDI = True Then
'Check for invalid Note
If Note2Play = -1 Then
    'display error message
    Msg = "Error, Invalid Note."
    MsgBox Msg, vbOKOnly + vbExclamation, "Error..."
    'exit subroutine
    Exit Sub
End If
'If used wants chord Picked, not strummed
If Strummed = False Then
    'Send the note to midi device, converting the String note to a Integer representative.
    PlayNote (Note2Play)
    'Sleep for a number of Miliseconds
    Sleep SleepTime
    'Send stop command to Midi Device, converting the String Note to Integer
    StopNote (Note2Play)
    'Sleep for number of miliseconds
    Sleep SleepTime / 2
    'Exit Subroutine
    Exit Sub
'End if condition
End If
'If user wants chord to be strummed
If Strummed = True Then
    'Send note to the midi Device, converting the String note to Integer
    PlayNote (Note2Play)
    'Set the timer interval to 1.5 seconds(1500 milliseconds)
    tmrCount.Interval = 1500
    'Enable the timer
    tmrCount.Enabled = True
    'Exit Subroutine
    Exit Sub
End If
End If
    'Exit Subroutine
    Exit Sub
ErrPlay:
'Message box
    Msg = "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & "Source:" & "Play()"
    MsgBox Msg, vbOKOnly + vbExclamation, "An Error Occured..."
    'Exit from sub
    Exit Sub
End Sub
Private Function NoteFromString(Note As Integer, WhichString As Integer) As Integer 'This gets the full note, including the base string setting
'Work with the value of WhichString variable
Select Case WhichString
'The value is 1(String 1, E)
Case 1
    'Return String Base Note, plus the Note passed to the function, Forming a complete note
    NoteFromString = StringE + Note
    'This is the same except for other Strings
Case 2
    NoteFromString = StringA + Note
Case 3
    NoteFromString = StringD + Note
Case 4
    NoteFromString = StringG + Note
Case 5
    NoteFromString = StringB + Note
Case 6
    NoteFromString = StringE2 + Note
Case Else
    'Return -1 ( Error )
    NoteFromString = -1
End Select
End Function
Private Sub NoteFromChordIndex(ChordIndex As Integer)  'This puts the notes to be played, derived from the chord fret numbers, into an array
    'Check Uppercase value of Chords(ChordIndex).E for 'X' Not played
    If UCase$(Chords(ChordIndex).E) = "X" Then
        'Set value to -1, meaning not played
        ANote(ChordIndex).Notes(0) = -1
    'It is played
    Else
        'If it is played, then put the value of E into the array
        ANote(ChordIndex).Notes(0) = Chords(ChordIndex).E
    'End condition statement
    End If
    'Below is the same, just for the other strings
    If UCase$(Chords(ChordIndex).A) = "X" Then
        ANote(ChordIndex).Notes(1) = -1
    Else
        ANote(ChordIndex).Notes(1) = Chords(ChordIndex).A
    End If
    If UCase$(Chords(ChordIndex).D) = "X" Then
        ANote(ChordIndex).Notes(2) = -1
    Else
        ANote(ChordIndex).Notes(2) = Chords(ChordIndex).D:
    End If
    If UCase$(Chords(ChordIndex).G) = "X" Then
        ANote(ChordIndex).Notes(3) = -1
    Else
        ANote(ChordIndex).Notes(3) = Chords(ChordIndex).G
    End If
    If UCase$(Chords(ChordIndex).B) = "X" Then
        ANote(ChordIndex).Notes(4) = -1
    Else
        ANote(ChordIndex).Notes(4) = Chords(ChordIndex).B
    End If
    If UCase$(Chords(ChordIndex).e2) = "X" Then
        ANote(ChordIndex).Notes(5) = -1
    Else
        ANote(ChordIndex).Notes(5) = Chords(ChordIndex).e2
    End If
End Sub

Private Sub PlayNote(Note As Integer)   'This user sub sends the note as a message to the midi device
    'This is the message it will send to the MIDI Device
    MidiMsg = &H90 + ((BaseNote + Note) * &H100) + (Volume * &H10000) + Channel
    'This is the API to send the message to the Handle of the midi device
    midiOutShortMsg hWndMidi, MidiMsg
End Sub
Private Sub StopNote(Note As Integer)   'This user Sub sends the stop command to the midi device
    'This is the message it will send to the MIDI Device
    MidiMsg = &H80 + ((BaseNote + Note) * &H100) + Channel
    'This is the API to send the message to the Handle of the midi device
    midiOutShortMsg hWndMidi, MidiMsg
End Sub

Private Sub ChangeInstrument(Instrument As Integer) 'This sends a message to the MIDI device telling it what instrument patch to use
    'This is the message it will send to the MIDI Device
    MidiMsg = (Instrument * 256) + &HC0 + Channel + (0 * 256) * 256
    'This is the API to send the message to the Handle of the midi device
    midiOutShortMsg hWndMidi, MidiMsg
End Sub
Private Sub SetupMIDI() 'This user sub sets the MIDI configuration for the default MIDI Mapper
    'return error code from function, 0=All clear
    rc = midiOutOpen(hWndMidi, -1, 0, 0, 0)
'If rc isnt 0
If rc <> 0 Then
    'Display error
    Msg = "Error, couldnt open Midi Device: MIDI Mapper"
    MsgBox Msg, vbOKOnly + vbCritical, "Error opening MIDI Device..."
    'Tells the program No MIDI device available
    HasMIDI = False
    'Exit Subroutine
    Exit Sub
End If
    'Set up the standard Variables
    Volume = 127
    'This base note was matched against my guitar in standard tuning
    BaseNote = 36
    Channel = 1
    'Tells the program a MIDI device is available
    HasMIDI = True
End Sub

Private Sub Sleep(MilliSeconds As Integer)   'This user sub makes the program wait for a number of seconds
    'Set Timer Interval to 1(1 millisecond)
    tmrCount.Interval = 1
    'Set the Enabled property to true
    tmrCount.Enabled = True
'Carry on until Counter = Seconds
Do Until Counter > MilliSeconds
    'Let the computer carry on with things, instead of locking the progrm like the Sleep API does!!
    DoEvents
'Loop until condition met
Loop
    'Disable the timer
    tmrCount.Enabled = False
    'Reset the Counter to 1
    Counter = 1
End Sub



Private Sub tmrCount_Timer()
Dim i As Integer
'If chord is not being strummed
If chkStrummed.Value = 0 Then
    'Increment Counter
    Counter = Counter + 1
Else    'If chord IS being strummed
    'Loop whilst incrementing I
    For i = 0 To 65 'to make sure we clear up any left overs.
        'Stop playing note i
        StopNote i
    'continue loop
    Next i
    'Enable Play button and Strummed Checkbox
    cmdPlay.Enabled = True
    chkStrummed.Enabled = True
    'Disable the timer
    tmrCount.Enabled = False
'end if condition
End If
End Sub
