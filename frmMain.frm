VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Penitentiary Horse Race"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Place Bet: "
      Height          =   1215
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   3255
      Begin VB.CommandButton cmdbet 
         Caption         =   "Bet"
         Height          =   315
         Left            =   510
         TabIndex        =   13
         Top             =   810
         Width           =   1905
      End
      Begin VB.OptionButton opt4 
         Caption         =   "Fulla Heart"
         Height          =   285
         Left            =   1530
         TabIndex        =   12
         Top             =   510
         Width           =   1245
      End
      Begin VB.OptionButton opt3 
         Caption         =   "Girls Best Friend"
         Height          =   285
         Left            =   1530
         TabIndex        =   11
         Top             =   240
         Width           =   1665
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Puppy Toes"
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Top             =   510
         Width           =   1245
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Spadey Lady"
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.ListBox lstCards 
      Height          =   1230
      Left            =   9060
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H002E8553&
      Height          =   2655
      Left            =   30
      ScaleHeight     =   2595
      ScaleWidth      =   10545
      TabIndex        =   0
      Top             =   1290
      Width           =   10605
      Begin VB.PictureBox pbDiamond 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         ScaleHeight     =   540
         ScaleWidth      =   720
         TabIndex        =   4
         Top             =   2010
         Width           =   720
      End
      Begin VB.PictureBox pbClub 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         ScaleHeight     =   540
         ScaleWidth      =   720
         TabIndex        =   3
         Top             =   1350
         Width           =   720
      End
      Begin VB.PictureBox pbHeart 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         ScaleHeight     =   540
         ScaleWidth      =   720
         TabIndex        =   2
         Top             =   690
         Width           =   720
      End
      Begin VB.PictureBox pbSpade 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   0
         ScaleHeight     =   540
         ScaleWidth      =   720
         TabIndex        =   1
         Top             =   60
         Width           =   720
      End
      Begin VB.Line Line14 
         X1              =   0
         X2              =   10560
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Line Line13 
         X1              =   0
         X2              =   10560
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Line Line12 
         X1              =   0
         X2              =   10560
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9120
         TabIndex        =   6
         Top             =   90
         Width           =   255
      End
      Begin VB.Line Line11 
         X1              =   9240
         X2              =   9240
         Y1              =   0
         Y2              =   2640
      End
      Begin VB.Line Line10 
         X1              =   8400
         X2              =   8400
         Y1              =   0
         Y2              =   2670
      End
      Begin VB.Line Line9 
         X1              =   7560
         X2              =   7560
         Y1              =   0
         Y2              =   2610
      End
      Begin VB.Line Line8 
         X1              =   6720
         X2              =   6720
         Y1              =   0
         Y2              =   2640
      End
      Begin VB.Line Line7 
         X1              =   5880
         X2              =   5880
         Y1              =   0
         Y2              =   2670
      End
      Begin VB.Line Line6 
         X1              =   5040
         X2              =   5040
         Y1              =   0
         Y2              =   2670
      End
      Begin VB.Line Line5 
         X1              =   4200
         X2              =   4200
         Y1              =   0
         Y2              =   2580
      End
      Begin VB.Line Line4 
         X1              =   3360
         X2              =   3360
         Y1              =   0
         Y2              =   2610
      End
      Begin VB.Line Line3 
         X1              =   2520
         X2              =   2520
         Y1              =   0
         Y2              =   2610
      End
      Begin VB.Line Line2 
         X1              =   1680
         X2              =   1680
         Y1              =   0
         Y2              =   2580
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   750
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.Line Line1 
         X1              =   840
         X2              =   840
         Y1              =   0
         Y2              =   2610
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4530
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":798B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B771
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFlip 
      Caption         =   "Flip Next Card"
      Height          =   435
      Left            =   600
      TabIndex        =   14
      Top             =   360
      Width           =   1785
   End
   Begin VB.Label lblBet 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   510
      TabIndex        =   16
      Top             =   60
      Width           =   2115
   End
   Begin VB.Label lblNext 
      Caption         =   "Next Card is a "
      Height          =   255
      Left            =   210
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function IsWinner() As Boolean
    IsWinner = False
    If pbClub.Left > 9000 And opt2.Value = True Then IsWinner = True
    If pbSpade.Left > 9000 And opt1.Value = True Then IsWinner = True
    If pbHeart.Left > 9000 And opt4.Value = True Then IsWinner = True
    If pbDiamond.Left > 9000 And opt3.Value = True Then IsWinner = True
End Function

Private Sub FlipNextCard()
    Dim nCount As Integer, nNext As Integer
    nCount = lstCards.ListCount - 1
    Randomize Timer
    nNext = Int(Rnd * (nCount - 0) + 1)
    Select Case lstCards.List(nNext)
        Case "Spade"
            MoveSpade
            lblNext.Caption = "Card is a spade."
        Case "Club"
            MoveClub
            lblNext.Caption = "Card is a club."
        Case "Diamond"
            MoveDiamond
            lblNext.Caption = "Card is a diamond."
        Case "Heart"
            MoveHeart
            lblNext.Caption = "Card is a heart."
    End Select
    lstCards.RemoveItem nNext
    lstCards.Refresh
End Sub

Private Sub MoveClub()
    If pbClub.Left = 0 Then
        pbClub.Left = 900
    Else
        pbClub.Left = pbClub.Left + 840
    End If
    If pbClub.Left > 9000 Then
        Dim nString As String
        cmdFlip.Enabled = False
        nString = "PUPPY TOES WINS!" & Chr$(13)
        If IsWinner = True Then
            nString = nString & "YOU WIN!"
        Else
            nString = nString & "YOU LOSE!"
        End If
        MsgBox nString, vbExclamation + vbOKOnly, "Penitentiary Horse Race"
        Unload Me
        End
    End If
End Sub

Private Sub MoveHeart()
    If pbHeart.Left = 0 Then
        pbHeart.Left = 900
    Else
        pbHeart.Left = pbHeart.Left + 840
    End If
    If pbHeart.Left > 9000 Then
        cmdFlip.Enabled = False
        Dim nString As String
        nString = "FULLA HEART WINS!" & Chr$(13)
        If IsWinner = True Then
            nString = nString & "YOU WIN!"
        Else
            nString = nString & "YOU LOSE!"
        End If
        MsgBox nString, vbExclamation + vbOKOnly, "Penitentiary Horse Race"
        Unload Me
        End
    End If
End Sub

Private Sub MoveDiamond()
    If pbDiamond.Left = 0 Then
        pbDiamond.Left = 900
    Else
        pbDiamond.Left = pbDiamond.Left + 840
    End If
    If pbDiamond.Left > 9000 Then
        cmdFlip.Enabled = False
        Dim nString As String
        nString = "GIRL'S BEST FRIEND WINS!" & Chr$(13)
        If IsWinner = True Then
            nString = nString & "YOU WIN!"
        Else
            nString = nString & "YOU LOSE!"
        End If
        MsgBox nString, vbExclamation + vbOKOnly, "Penitentiary Horse Race"
        Unload Me
        End
    End If
End Sub

Private Sub MoveSpade()
    If pbSpade.Left = 0 Then
        pbSpade.Left = 900
    Else
        pbSpade.Left = pbSpade.Left + 840
    End If
    If pbSpade.Left > 9000 Then
        cmdFlip.Enabled = False
        Dim nString As String
        nString = "SPADEY LADY WINS!" & Chr$(13)
        If IsWinner = True Then
            nString = nString & "YOU WIN!"
        Else
            nString = nString & "YOU LOSE!"
        End If
        MsgBox nString, vbExclamation + vbOKOnly, "Penitentiary Horse Race"
        Unload Me
        End
    End If
End Sub

Private Sub cmdbet_Click()
    If opt1.Value = False And opt2.Value = False And opt3.Value = False And opt4.Value = False Then
        MsgBox "Please place your bet.", vbCritical + vbOKOnly, "Penitentiary Horse Race"
        Exit Sub
    Else
        Frame1.Visible = False
        lblBet.Caption = "You bet on "
        If opt1.Value = True Then lblBet.Caption = lblBet.Caption & "Spadey Lady."
        If opt2.Value = True Then lblBet.Caption = lblBet.Caption & "Puppy Toes."
        If opt3.Value = True Then lblBet.Caption = lblBet.Caption & "Girl's best friend."
        If opt4.Value = True Then lblBet.Caption = lblBet.Caption & "Fulla Heart."
    End If
End Sub

Private Sub cmdFlip_Click()
    FlipNextCard
    lblNext.Visible = True
End Sub

Private Sub Form_Load()
    pbSpade.Picture = ImageList1.ListImages(1).Picture
    pbClub.Picture = ImageList1.ListImages(2).Picture
    pbDiamond.Picture = ImageList1.ListImages(3).Picture
    pbHeart.Picture = ImageList1.ListImages(4).Picture
    lblNext.Visible = False
    pbSpade.Left = 0
    pbClub.Left = 0
    pbHeart.Left = 0
    pbDiamond.Left = 0
    pbSpade.AutoSize = True
    Label1.Caption = "S" & Chr$(13) & "T" & Chr$(13) & "A" & Chr$(13) & "R" & Chr$(13) & "T"
    Label1.AutoSize = True
    Label2.Caption = "F" & Chr$(13) & "I" & Chr$(13) & "N" & Chr$(13) & "I" & Chr$(13) & "S" & Chr$(13) & "H"
    Label2.AutoSize = True
    Dim I As Integer
    lstCards.Clear
    For I = 1 To 13
        lstCards.AddItem "Spade"
        lstCards.AddItem "Heart"
        lstCards.AddItem "Club"
        lstCards.AddItem "Diamond"
    Next I
    Frame1.Visible = True
    cmdFlip.Enabled = True
End Sub
