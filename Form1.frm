VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo -=[ cHttpClient ]=-"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmEnvoi 
      Caption         =   "Envoyer: "
      Height          =   1635
      Left            =   60
      TabIndex        =   8
      Top             =   6840
      Width           =   7935
      Begin VB.TextBox txtPage 
         Height          =   315
         Left            =   4080
         TabIndex        =   16
         Text            =   "/chttpclient/page_test.php"
         Top             =   540
         Width           =   3735
      End
      Begin VB.TextBox txtAdresse 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Text            =   "www.open-design.be"
         Top             =   540
         Width           =   3735
      End
      Begin VB.TextBox txtDate 
         Height          =   315
         Left            =   2040
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtNom 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Envoyer infos en POST"
         Height          =   375
         Left            =   5700
         TabIndex        =   9
         Top             =   1140
         Width           =   2115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Page:"
         Height          =   195
         Left            =   4080
         TabIndex        =   17
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Adresse:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Une date de naissance:"
         Height          =   195
         Left            =   2040
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Un nom:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   600
      End
   End
   Begin VB.Frame frmRecData 
      Caption         =   "Donnees recues: "
      Height          =   4455
      Left            =   60
      TabIndex        =   6
      Top             =   2280
      Width           =   7935
      Begin VB.TextBox txtRecData 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   4095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Headers recus"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   3915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Donnees envoyees"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   3915
   End
   Begin VB.Frame frmHeaders 
      Caption         =   "Headers recus: "
      Height          =   1575
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   7935
      Begin VB.ListBox lstRecHeaders 
         Height          =   1230
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Frame frmEnv 
      Caption         =   "Donnees envoyees: "
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      Begin VB.TextBox txtEnvData 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   7695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim WithEvents cH As cHttpClient
Attribute cH.VB_VarHelpID = -1
'
Const AppCaption = "Demo -=[ cHttpClient ]=-"
'
Dim FichierId As Integer
Dim Dossier As String
Dim Nom As String
'
'socket connecte
Private Sub cH_connection()
    '
    Me.Caption = AppCaption & " connecte"
    '
    ouvrirFichier
    '
End Sub
'
'erreur
Private Sub cH_erreur(Data As String)
    '
    Me.Caption = AppCaption & " erreur: " & Data
    '
End Sub
'
'socket ferme
Private Sub cH_fermeture()
    '
    If cH.RecTailleData = cH.RecTailleTotaleData And cH.ChunkRecFin = True Then
        '
        Me.Caption = AppCaption & "deconnecte: toutes les donnees ont ete recues: " & cH.RecTailleData & "/" & cH.RecTailleTotaleData
        '
    Else
        '
        Me.Caption = AppCaption & "deconnecte: donnees recues: " & cH.RecTailleData & "/" & cH.RecTailleTotaleData
        '
    End If
    '
    fermerFichier
    '
End Sub
'
'reception des headers
Private Sub cH_headers(Nom As String, Data As String)
    '
    lstRecHeaders.AddItem Nom & ": " & Data
    '
End Sub
'
'reception des donnees apres les headers
Private Sub cH_reception(Data As String, TailleRecu As Long, TailleTotaleRecu As Long, Fini As Boolean)
    '
    txtRecData = txtRecData & Data
    '
    Put #FichierId, , Data
    '
    'on verifie si on a tout recu
    'TailleRecu = la taille des donnees recues
    'TailleTotaleRecu = la taille totale des donnees a recevoir (peut varier...)
    'Fini = specifie si la taille totale des donnees a recevoir est definitive
    'varie aussi mais lorsqu elle est a TRUE, elle le reste
    'Ceci est du au fait qu un transfert peut etre decoupe en chunk
    'donc pour verifier si toutes les donnees ont ete recues:
    If TailleRecu = TailleTotaleRecu And Fini = True Then
        '
        Me.Caption = AppCaption & " toutes les donnees ont ete recues: " & TailleRecu & "/" & TailleTotaleRecu
        '
    Else
        '
        Me.Caption = AppCaption & " donnees recues: " & TailleRecu & "/" & TailleTotaleRecu
        '
    End If
    '
End Sub
'
Private Sub cH_timeout()
    '
    Me.Caption = AppCaption & " Time Out !"
    '
End Sub
'
Private Sub Command1_Click()
    '
    Me.txtEnvData = ""
    Me.txtRecData = ""
    Me.lstRecHeaders.Clear
    '
    cH.initVars
    cH.ajouterFormData "nom", txtNom.Text
    cH.ajouterFormData "date_naissance", txtDate.Text
    cH.ajouterFormData "localisation", "planet-source-code"
    'cH.ajouterFormData "ze_fichier", "", "c:\log.txt", "text/plain"
    '
    If cH.connecter(txtAdresse.Text, 80, txtPage.Text, "POST") = True Then
        '
        Me.Caption = AppCaption & " en attente de connexion"
        '
    Else
        '
        Me.Caption = AppCaption & " erreur lors de la connexion"
        '
    End If
    '
End Sub
'
Private Sub Command2_Click()
    '
    frmEnv.Visible = True
    frmHeaders.Visible = False
    '
End Sub

Private Sub Command3_Click()
    '
    frmEnv.Visible = False
    frmHeaders.Visible = True
    '
End Sub

Private Sub Form_Load()
    '
    Set cH = New cHttpClient
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    Set cH = Nothing
    '
End Sub
'
'ouvre le fichier
Public Sub ouvrirFichier()
    '
    If FichierId = 0 Then
        '
        Nom = "reception.txt"
        '
        'on determine le dossier de sauvegarde
        Dossier = App.Path
        '
        FichierId = FreeFile
        '
        Open Dossier & "\" & Nom For Binary As #FichierId
        '
    End If
    '
End Sub
'
'ferme le fichier
Public Sub fermerFichier()
    '
    'On Error Resume Next
    '
    If FichierId > 0 Then Close #FichierId
    '
    FichierId = 0
    '
    'If Err Then log Err.Description & " " & Err.Number
    '
End Sub
