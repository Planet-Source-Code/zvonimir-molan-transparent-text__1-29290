VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Transparent Text"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Change text color"
      Height          =   405
      Left            =   5235
      TabIndex        =   4
      Top             =   4230
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Picture"
      Height          =   405
      Left            =   45
      TabIndex        =   1
      Top             =   4230
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change cursor color"
      Height          =   405
      Left            =   3600
      TabIndex        =   3
      Top             =   4230
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change cursor blinking rate"
      Height          =   405
      Left            =   1425
      TabIndex        =   2
      Top             =   4230
      Width           =   2160
   End
   Begin VB.Timer Timer1 
      Left            =   6045
      Top             =   165
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4110
      Left            =   45
      Picture         =   "Main.frx":0000
      ScaleHeight     =   4050
      ScaleWidth      =   6510
      TabIndex        =   0
      Top             =   45
      Width           =   6570
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5940
         Top             =   630
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "GIF|*.gif|JPG|*.jpg|BMP|*.bmp"
      End
      Begin VB.Label Txt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Run the program and select the TXT object with mouse to edit TEXT"
         Height          =   195
         Index           =   8
         Left            =   60
         MousePointer    =   3  'I-Beam
         TabIndex        =   9
         Top             =   1365
         Width           =   5040
      End
      Begin VB.Label Txt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Change Fonts, setting of fonts as you wish"
         Height          =   195
         Index           =   7
         Left            =   60
         MousePointer    =   3  'I-Beam
         TabIndex        =   8
         Top             =   1065
         Width           =   3165
      End
      Begin VB.Label Txt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Copy and paste as many TXT object as you wish"
         Height          =   195
         Index           =   5
         Left            =   60
         MousePointer    =   3  'I-Beam
         TabIndex        =   7
         Top             =   765
         Width           =   3630
      End
      Begin VB.Label TxtInit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Init"
         Height          =   195
         Left            =   6045
         TabIndex        =   6
         Top             =   1215
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Shape Cur 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         DrawMode        =   14  'Copy Pen
         Height          =   210
         Left            =   6150
         Top             =   1530
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.Label Txt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transparent Text (example)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   75
         MousePointer    =   3  'I-Beam
         TabIndex        =   5
         Top             =   75
         Width           =   5850
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Pos As Integer
Private PosEnd As Integer
Private TXTPos As Integer
Private CurBRate As Long
Public Sub TXTInitSet(ByVal Index As Long)
    On Error GoTo TXTError
    'set the font settings of TXTInit object equal to TXT object
    TxtInit.Font = Txt(Index).Font
    TxtInit.Font.Bold = Txt(Index).Font.Bold
    TxtInit.Font.Italic = Txt(Index).Font.Italic
    TxtInit.Font.Size = Txt(Index).Font.Size
    TxtInit.Font.Strikethrough = Txt(Index).Font.Strikethrough
    TxtInit.Font.Underline = Txt(Index).Font.Underline
    'set the text of TXTInit equal to TXT object
    If Pos = PosEnd Then TxtInit = Txt(Index)
    If Pos <> PosEnd And Pos <> 0 Then TxtInit = Mid(Txt(Index), 1, Pos)
    'set the cursor position
    Cur.Width = 30
    Cur.Height = Txt(Index).Height
    If Pos <> 0 Then Cur.Left = Txt(Index).Left + TxtInit.Width Else Cur.Left = Txt(Index).Left
    Cur.Top = Txt(Index).Top
    Timer1.Interval = CurBRate
    TXTPos = Index 'set the TXTPos to index value
    Exit Sub
    
TXTError:
    TXTPos = -1 'set TXTPos to -1 because there is no object with Index value
End Sub

Private Sub Command1_Click()
    CurBRate = Val(InputBox("Set the new cursor blinking rate in 1/1000 sec.", "Cursor blinking rate", CurBRate))
    If CurBRate = 10 Then CurBRate = 11
End Sub

Private Sub Command2_Click()
    'Open the DialogBox with colors
    On Error GoTo ErrorOpen
    CommonDialog1.ShowColor
    Cur.BackColor = CommonDialog1.Color
    Exit Sub
    
ErrorOpen:
    'Cancel was selected
End Sub

Private Sub Command3_Click()
    'Open the new picture in PictureBox
    On Error GoTo ErrorOpen
    CommonDialog1.ShowOpen
    Picture1.Picture = LoadPicture(CommonDialog1.FileName)
    Exit Sub
    
ErrorOpen:
    'Cancel was selected
End Sub


Private Sub Command4_Click()
    Dim I As Integer
    'Open the DialogBox with colors
    On Error GoTo ErrorOpen
    CommonDialog1.ShowColor
    On Error Resume Next
    For I = 0 To Txt.UBound
        Txt(I).ForeColor = CommonDialog1.Color
    Next
    Exit Sub
    
ErrorOpen:
    'Cancel was selected
End Sub

Private Sub Form_Load()
    CurBRate = 400 'Set the cursor blinking rate
End Sub

Private Sub Picture1_GotFocus()
    Dim I As Integer
    
    'Set Focus to first Txt object
    For I = Txt.LBound To Txt.UBound
        TXTInitSet (I)
        If TXTPos <> -1 Then Exit For
    Next
    Pos = Len(Txt(TXTPos)): PosEnd = Pos 'position of the cursor and length of the text
    TXTInitSet (I)
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    If TXTPos <= 0 Then Exit Sub
    Dim I As Integer
    
    Select Case KeyCode
    Case vbKeyBack
        If PosEnd > 0 And Pos > 0 Then
            PosEnd = PosEnd - 1: Pos = Pos - 1
            If Pos = PosEnd Then
                Txt(TXTPos) = Mid(Txt(TXTPos), 1, PosEnd)
            Else
                Txt(TXTPos) = Mid(Txt(TXTPos), 1, Pos) + Right(Txt(TXTPos), Len(Txt(TXTPos)) - Pos - 1)
            End If
            TXTInitSet (TXTPos)
        End If
    Case vbKeyDelete
        If Pos <> PosEnd Then
            PosEnd = PosEnd - 1
            Txt(TXTPos) = Mid(Txt(TXTPos), 1, Pos) + Right(Txt(TXTPos), Len(Txt(TXTPos)) - Pos - 1)
            TXTInitSet (TXTPos)
        End If
    Case vbKeyReturn, vbKeyDown
        If Txt.Count <= 1 Then Exit Sub
        'Set Focus to next Txt object, if the current is the last then set to first
        If TXTPos <> Txt.UBound Then
            For I = TXTPos + 1 To Txt.UBound
                TXTInitSet (I)
                If TXTPos <> -1 Then Exit For
            Next
        Else
            For I = Txt.LBound To Txt.UBound
                TXTInitSet (I)
                If TXTPos <> -1 Then Exit For
            Next
        End If
        Pos = Len(Txt(TXTPos)): PosEnd = Pos 'position of the cursor and length of the text
        TXTInitSet (I)
    Case vbKeyUp
        If Txt.Count <= 1 Then Exit Sub
        'Set Focus to previous Txt object, if the current is the first then set to last
        If TXTPos <> Txt.LBound Then
            For I = TXTPos - 1 To Txt.LBound Step -1
                TXTInitSet (I)
                If TXTPos <> -1 Then Exit For
            Next
        Else
            For I = Txt.UBound To Txt.LBound Step -1
                TXTInitSet (I)
                If TXTPos <> -1 Then Exit For
            Next
        End If
        Pos = Len(Txt(TXTPos)): PosEnd = Pos 'position of the cursor and length of the text
        TXTInitSet (I)
    Case vbKeyLeft
        If Pos > 0 Then
            Pos = Pos - 1
            TXTInitSet (TXTPos)
        End If
    Case vbKeyRight
        If Pos < PosEnd Then
            Pos = Pos + 1
            TXTInitSet (TXTPos)
        End If
    End Select
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    If TXTPos <= 0 Or KeyAscii < 31 Then Exit Sub
    
    'Add char code to Txt object
    If Pos = PosEnd Then
        Txt(TXTPos) = Txt(TXTPos) + Chr(KeyAscii)
    Else
        Txt(TXTPos) = Mid(Txt(TXTPos), 1, Pos) + Chr(KeyAscii) + Right(Txt(TXTPos), Len(Txt(TXTPos)) - Pos)
    End If
    PosEnd = PosEnd + 1: Pos = Pos + 1
    TXTInitSet (TXTPos)
End Sub


Private Sub Picture1_LostFocus()
    Timer1.Interval = 10 'Clear the cursor
    TXTPos = 0 'Set the TXTPos to 0, because there is no active Txt object
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Interval = 10 'Clear the cursor
    TXTPos = 0 'Set the TXTPos to 0, because there is no active Txt object
End Sub

Private Sub Timer1_Timer()
    If Timer1.Interval = 10 Then Cur.Visible = False: Timer1.Interval = 0: Exit Sub

    If Cur.Visible = True Then Cur.Visible = False: Exit Sub
    Cur.Visible = True
End Sub


Private Sub Txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    Dim Temp As Long
    
    Picture1.SetFocus ' Focus must be on PictureBox to change the Txt objects
    Pos = Len(Txt(Index)): PosEnd = Pos 'position of the cursor and length of the text
    TXTInitSet (Index) 'set the cursor position
    
    'Find the position within the text
    For I = 1 To Len(Txt(Index))
        TxtInit = Mid(Txt(Index), 1, I)
        If TxtInit.Width > X Then
            TxtInit = Mid(Txt(Index), 1, I - 1)
            If I = 1 Then Temp = 0 Else Temp = TxtInit.Width
            TxtInit = Mid(Txt(Index), I, 1)
            If Temp + (TxtInit.Width / 2) > X Then TxtInit = Mid(Txt(Index), 1, I - 1) Else TxtInit = Mid(Txt(Index), 1, I)
            Exit For
        End If
    Next
    Pos = Len(TxtInit)
    TXTInitSet (Index)
End Sub


