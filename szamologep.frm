VERSION 5.00
Begin VB.Form Szamologep 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hangos számológép - Muráti Ákos"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "szamologep.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton felolv 
      Caption         =   "C<<"
      Height          =   975
      Left            =   120
      TabIndex        =   20
      ToolTipText     =   "[Page Up]"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer megjelenit 
      Interval        =   1
      Left            =   5160
      Top             =   0
   End
   Begin VB.CommandButton egyenlo 
      Caption         =   "="
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3720
      TabIndex        =   19
      ToolTipText     =   "[ENTER]"
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton torol 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "[ESC]"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton torol 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   1320
      TabIndex        =   17
      ToolTipText     =   "[Backspace]"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton muvelet 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "[+]"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton muvelet 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "[-]"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton muvelet 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "[*]"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton muvelet 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "[/]"
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton tizedes 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   12
      ToolTipText     =   "[,]"
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton elojel 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   11
      ToolTipText     =   "[Page Down]"
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   2520
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   1320
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   2520
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton szam 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox kijelzo 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,000000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1038
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   4575
   End
   Begin VB.Timer szunet 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   600
      Top             =   240
   End
End
Attribute VB_Name = "Szamologep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'DirectX és DS életrehívása
Dim Dx As New DirectX7
Dim Ds As DirectSound

Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX

Dim hangok(0 To 50) As DirectSoundBuffer 'A 51 jelentése "semmi"
Dim jatszando(0 To 100) As Byte
Dim mutato As Byte

'Számológép változói
Dim a As Double, kep As Double, muv As Byte

Private Sub egyenlo_Click()
On Error GoTo hiba
    hangok(42).Stop
    hangok(42).SetCurrentPosition 0
    
    hangok(40).Play DSBPLAY_DEFAULT
    Select Case muv
        Case 0
            kep = a / kep
        Case 1
            kep = a * kep
        Case 2
            kep = a - kep
        Case 3
            kep = a + kep
        Case 4
            felolv_Click
    End Select
    kijelzo.Text = 0
    kijelzo.Text = kep
    If kep = 0 Then
        Szamot_hangga (0)
    End If
    muv = 4
    alapgombok
Exit Sub
hiba:
    Szamot_hangga (0)
    kep = 0
    muv = 4
    alapgombok
End Sub

Private Sub elojel_Click()
    kijelzo.Text = -1 * CDbl(kijelzo.Text)
    egyenlo.SetFocus
End Sub


Private Sub felolv_Click()
    Szamot_hangga CDbl(kijelzo.Text) '(kep) 'CDbl(kijelzo.Text))
    egyenlo.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    hangok(42).Stop
    hangok(42).SetCurrentPosition 0
    Select Case KeyCode
        Case 13
            egyenlo_Click
        Case 33
            felolv_Click
        Case 34
            elojel_Click
        Case 112
            hangok(42).Play DSBPLAY_DEFAULT
        Case 67
            If Shift = 2 Then
                Clipboard.SetText kijelzo.Text
            End If
        'Case Else
        '    MsgBox KeyCode & " " & Shift
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    hangok(42).Stop
    hangok(42).SetCurrentPosition 0
    Select Case KeyAscii
        Case 48 To 57
            szam_Click (CByte(Chr(KeyAscii)))
        Case 13
            egyenlo_Click
        Case 8
            torol_Click (0)
        Case 27
            torol_Click (1)
        Case 47
            muvelet_Click (0)
        Case 42
            muvelet_Click (1)
        Case 43
            muvelet_Click (3)
        Case 44
            tizedes_Click
        Case 45
            muvelet_Click (2)
        Case 3
            
        Case Else
           'MsgBox KeyAscii
           Form_KeyDown 112, 0
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo hiba
    Dim i As Integer
    'Olvaso.Show
    'Olvaso.SetFocus
    'Olvaso.ZOrder 0
    
    For i = 0 To 9
        szam(i).Caption = i
        szam(i).ToolTipText = "[" & i & "]"
    Next i
    kijelzo.Text = 0
    'kep = 0
    muv = 4
    'Hangvezérlés töltése
    Set Ds = Dx.DirectSoundCreate("")

    Ds.SetCooperativeLevel Szamologep.hWnd, DSSCL_NORMAL

    DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    DsWave.nFormatTag = WAVE_FORMAT_PCM
    DsWave.nChannels = 2
    DsWave.lSamplesPerSec = 22050
    DsWave.nBitsPerSample = 16
    DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
    DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign
    szunet.Interval = 550
    
    'Hangok betöltése a memóriába
    Set hangok(0) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\nulla.wav", DsDesc, DsWave)
    Set hangok(1) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\egy.wav", DsDesc, DsWave)
    Set hangok(2) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\ketto.wav", DsDesc, DsWave)
    Set hangok(3) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\harom.wav", DsDesc, DsWave)
    Set hangok(4) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\negy.wav", DsDesc, DsWave)
    Set hangok(5) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\ot.wav", DsDesc, DsWave)
    Set hangok(6) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\hat.wav", DsDesc, DsWave)
    Set hangok(7) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\het.wav", DsDesc, DsWave)
    Set hangok(8) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\nyolc.wav", DsDesc, DsWave)
    Set hangok(9) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\kilenc.wav", DsDesc, DsWave)
    Set hangok(10) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tiz.wav", DsDesc, DsWave)
    Set hangok(11) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tizen.wav", DsDesc, DsWave)
    Set hangok(12) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\husz.wav", DsDesc, DsWave)
    Set hangok(13) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\huszon.wav", DsDesc, DsWave)
    Set hangok(14) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\harminc.wav", DsDesc, DsWave)
    Set hangok(15) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\negyven.wav", DsDesc, DsWave)
    Set hangok(16) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\otven.wav", DsDesc, DsWave)
    Set hangok(17) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\hatvan.wav", DsDesc, DsWave)
    Set hangok(18) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\hetven.wav", DsDesc, DsWave)
    Set hangok(19) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\nyolcvan.wav", DsDesc, DsWave)
    Set hangok(20) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\kilencven.wav", DsDesc, DsWave)
    Set hangok(21) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\szaz.wav", DsDesc, DsWave)
    Set hangok(22) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\ezer.wav", DsDesc, DsWave)
    Set hangok(23) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\millio.wav", DsDesc, DsWave)
    Set hangok(24) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\billio.wav", DsDesc, DsWave)
    Set hangok(25) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\minusz.wav", DsDesc, DsWave)
    Set hangok(26) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\egesz.wav", DsDesc, DsWave)
    Set hangok(27) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tized.wav", DsDesc, DsWave)
    Set hangok(28) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\szazad.wav", DsDesc, DsWave)
    Set hangok(29) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\ezred.wav", DsDesc, DsWave)
    Set hangok(30) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tizezred.wav", DsDesc, DsWave)
    Set hangok(31) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\szazezred.wav", DsDesc, DsWave)
    Set hangok(32) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\milliomod.wav", DsDesc, DsWave)
    Set hangok(33) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tizmilliomod.wav", DsDesc, DsWave)
    Set hangok(34) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\szazmilliomod.wav", DsDesc, DsWave)
    Set hangok(35) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\billiomod.wav", DsDesc, DsWave)
    Set hangok(36) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\tizbilliomod.wav", DsDesc, DsWave)
    Set hangok(37) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\meg.wav", DsDesc, DsWave)
    Set hangok(38) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\szor.wav", DsDesc, DsWave)
    Set hangok(39) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\oszt.wav", DsDesc, DsWave)
    Set hangok(40) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\egyenlo.wav", DsDesc, DsWave)
    Set hangok(41) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\nemtom.wav", DsDesc, DsWave)
    Set hangok(42) = Ds.CreateSoundBufferFromFile(App.Path & "\hangok\sugo.wav", DsDesc, DsWave)
Exit Sub
hiba:
    Select Case Err.Number
        Case 432
            MsgBox "Az egyik hang fájl nem található. Kérem telepítsen újra!", vbInformation, "Inicializálási hiba"
            Unload Me
        Case Else
            MsgBox Err.Description, vbInformation, "A rendszer hibaüzenete: " & Err.Number
            Unload Me
    End Select
    End
End Sub
Public Function Kovetkezo()
    If mutato > 0 Then
            hangok(jatszando(mutato)).Play DSBPLAY_DEFAULT
            mutato = mutato - 1
        Else
            szunet.Enabled = False
            'Unload Me
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    For i = 0 To 30
        Set hangok(i) = Nothing
    Next i
End Sub

Private Sub kijelzo_Change()
On Error Resume Next
    'If CDbl(kijelzo.Text) <> 0 Or (kijelzo.Text = 0 And muv = 4) Then
        If Len(kijelzo.Text) > 1 And Mid(kijelzo.Text, 1, 1) = "0" And Mid(kijelzo.Text, 2, 1) <> "," Then
            kijelzo.Text = Mid(kijelzo.Text, 2, Len(kijelzo.Text) - 1)
        End If
    
   lecsap
    
    szunet.Enabled = False
    Szamot_hangga (kijelzo.Text)
    
    'End If
End Sub

Private Sub kijelzo_Click()
    'MsgBox CLng(kijelzo.Text)
End Sub

Private Sub megjelenit_Timer()
On Error Resume Next
   kep = CDbl(kijelzo.Text)
End Sub

Private Sub muvelet_Click(Index As Integer)
    'If a <> kep Then
    a = kep
    muv = Index
    'kep
    kijelzo.Text = ""
    alapgombok
    muvelet(Index).BackColor = &H80000014
    Select Case Index
        Case 0
            hangok(39).Play DSBPLAY_DEFAULT
        Case 1
            hangok(38).Play DSBPLAY_DEFAULT
        Case 2
            hangok(25).Play DSBPLAY_DEFAULT
        Case 3
            hangok(37).Play DSBPLAY_DEFAULT
    End Select
    egyenlo.SetFocus
End Sub

Private Sub szam_Click(Index As Integer)
    'kep
    kijelzo.Text = kijelzo.Text & Index
    egyenlo.SetFocus
End Sub

Private Sub szunet_Timer()
    If jatszando(mutato) <> 51 Then
            Kovetkezo
        Else
            mutato = mutato - 1
            szunet_Timer
    End If
End Sub
Public Sub Szamot_hangga(szam As Double)
On Error GoTo hiba:
    'Form_Load
    'Olvaso.Cls
    'Olvaso.Print (szam)
    
    Dim egesz_resz As String, tort_resz As String
    Dim tort_str As String, i As Integer
    
    TorolJatszando
    tort_str = CStr(szam)
    egesz_resz = tort_str
    
    For i = 1 To Len(tort_str)
        If Mid(tort_str, i, 1) = "." Or Mid(tort_str, i, 1) = "," Then
            egesz_resz = CLng(Mid(tort_str, 1, i - 1))
            tort_resz = Mid(tort_str, i + 1, Len(tort_str) - i)
            GoTo kilepes
        End If
    Next i
kilepes:
    If Len(tort_resz) > 0 And Len(tort_resz) <= 10 Then
        UjJatszando (26 + Len(tort_resz))
        egesz (tort_resz)
        UjJatszando (26)
    Else
        If Len(tort_resz) > 10 Then GoTo hiba
    End If
    egesz (egesz_resz)
    If szam < 0 Then UjJatszando (25)
    
    szunet.Enabled = True
    mutato = jatszando(0)
    Exit Sub
hiba:
    'MsgBox "Túl nagy a szám!", vbInformation, "Longint Túlcsordulás"
    hangok(41).Play DSBPLAY_DEFAULT
    'Unload Me
End Sub
Private Sub UjJatszando(Index As Byte)
    jatszando(0) = jatszando(0) + 1
    jatszando(jatszando(0)) = Index
End Sub
Private Sub TorolJatszando()
    For i = 0 To 100
        jatszando(i) = 0
    Next i
End Sub
Private Function UtolsoJatszando() As Byte
    UtolsoJatszando = jatszando(jatszando(0))
End Function
Private Sub UjEgyes(szam)
    If szam <> 0 Then
            UjJatszando (szam)
        Else
            UjJatszando 51
    End If
End Sub
Private Sub egesz(szam As Long)
    On Error GoTo hiba
    Dim i As Integer
    Dim szam_str As String ', szoveg As String, kj As String
    Dim jegyek(1 To 10) As Byte ', atalakitott(1 To 10) As String, minus As String
    'Hibajelzés kikapcsolása
    hangok(41).Stop
    hangok(41).SetCurrentPosition 0
    
    szam_str = CStr(Abs(szam))
    'TorolJatszando
    
    'a számjegyek sorrendjének felcserélése
    For i = 1 To 10
        jegyek(i) = 0
    Next i
    
    For i = 1 To Len(szam_str)
        jegyek(i) = Mid(szam_str, Len(szam_str) - i + 1, 1)
    Next i
    
    'Helyiértékes vizsgálat
    For i = 1 To Len(szam_str)
        Select Case i
            Case 1
                If jegyek(i) = 0 And Len(szam_str) = 1 Then
                       UjJatszando (0)
                    Else
                        UjEgyes (jegyek(i))
                End If
            
            Case 2, 5, 8
                If UtolsoJatszando = 51 And (jegyek(i) = 1 Or jegyek(i) = 2) Then
                        If jegyek(i) = 1 Then
                                UjJatszando (10)
                            Else
                                UjJatszando (12)
                        End If
                    Else
                        Select Case jegyek(i)
                            Case 0
                                UjJatszando (51)
                            Case 1
                                UjJatszando (11)
                            Case 2
                                UjJatszando (13)
                            Case Else
                                UjJatszando (11 + jegyek(i))
                        End Select
                End If
            Case 3, 6, 9
                If jegyek(i) = 1 Then
                        UjJatszando (21)
                    Else
                        If jegyek(i) <> 0 Then
                            UjJatszando (21)
                            UjEgyes jegyek(i)
                        End If
                End If
            Case 4
                If jegyek(i) <> 0 Or jegyek(i + 1) <> 0 Or jegyek(i + 2) <> 0 Then
                    UjJatszando (22)
                End If
                
                If (jegyek(i) = 1 And Len(szam_str) > 4) Or jegyek(i) <> 1 Then  'jegyek(i + 1) = 0 And jegyek(i + 2) = 0 Then
                       UjEgyes jegyek(i)
                End If
            Case 7
                If jegyek(i) <> 0 Or jegyek(i + 1) <> 0 Or jegyek(i + 2) <> 0 Then
                    UjJatszando (23)
                End If
                UjEgyes jegyek(i)
            Case 10
                UjJatszando (24)
                UjEgyes jegyek(i)
        End Select
kov:
    Next i
    
    'A szám negatív
    'If szam < 0 Then UjJatszando (25)
    
    'szunet.Enabled = True
    'mutato = jatszando(0)
    Exit Sub
hiba:
    MsgBox Err.Description, vbInformation, "A rendszer hibaüzenete:"
End Sub


Private Sub tizedes_Click()
    If InStr(kijelzo.Text, ",") = 0 Then
        kijelzo.Text = kijelzo.Text & ","
    End If
End Sub

Private Sub torol_Click(Index As Integer)
    If Index = 1 Then
            'kep = 0
            kijelzo.Text = 0
            a = 0
            muv = 4
            alapgombok
        Else
            If Len(CStr(kijelzo.Text)) > 1 Then
                    'kep
                    kijelzo.Text = CDbl(Mid(kijelzo.Text, 1, Len(kijelzo.Text) - 1))
                Else
                    kijelzo.Text = 0
            End If
    End If
    egyenlo.SetFocus
End Sub
Private Sub alapgombok()
On Error Resume Next
    For i = 0 To 3
        muvelet(i).BackColor = vbButtonFace
    Next i
End Sub

Private Sub lecsap()
On Error GoTo hiba
     Dim egesz As String, tt As String, i As Integer
    
    For i = 1 To Len(kijelzo.Text)
        If Mid(kijelzo.Text, i, 1) = "." Or Mid(kijelzo.Text, i, 1) = "," Then
            egesz = Mid(kijelzo.Text, 1, i - 1)
            tt = Mid(kijelzo.Text, i + 1, Len(kijelzo.Text) - i)
        End If
    Next i
    If Len(tt) > 9 Then tt = Mid(tt, 1, 9)
    If Len(tt) <> 0 Then kijelzo.Text = egesz & "," & tt
Exit Sub
hiba:
    
End Sub
