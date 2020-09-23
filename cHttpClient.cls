VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Classe cHttpClient
' (c) Thomas John (NoRabbit)
'
'
'classe CSocket remplaçant le contrôle winsock (http://www.vbip.com)
Public WithEvents cs As CSocket
Attribute cs.VB_VarHelpID = -1
'
'classe minuteur reception des donnees
Public WithEvents csTmrReception As Minuteur
Attribute csTmrReception.VB_VarHelpID = -1
'
'classe minuteur time out connexion
Public WithEvents csTmrTimeOut As Minuteur
Attribute csTmrTimeOut.VB_VarHelpID = -1
'
'
'contient le "reste" d'une commande (voir le traîtement des données)
Dim ResteRecptTemp As String
'
'le nombre max de donnees a envoyer en une fois (octets)
Const MaxDonneesEnvoi = 4
'
'spécifie si le socket est prêt pour envoyer des données
Dim PretEnvoiSz As Boolean
'
'specifie si on peut fermer le socket apres la fin d un envoi de donnees
Dim FermerSocketApresEnvoi As Boolean
'
'
'serveur
Dim ServeurSz As String
'
'port
Dim PortSz As Long
'
'methode de la requete (GET, POST,...)
Dim MethodeRequete As String
'
'page appellee
Dim PageRequete As String
'
'
'headers a envoyer
Dim EnvHeaders() As String
'
'headers recus
Dim RecHeaders() As String
'
'totalite des headers recus ?
Dim bRecHeaders As Boolean
'
'
'reponse du serveur (200, 404, ...)
Dim ReponseHTTP As String
'
'message de la reponse du serveur (OK, ...)
Dim ReponseMsg As String
'
'tranfer encoding (chunked, ...)
Dim TransferEncoding As String
'
'
'version du protocole (HTTP/1.1)
Dim VersionProtocole As String
'
'FORM DATA
'contient le nom des data qu on envoie
Dim FormDataNom() As String
'
'contient le nom du fichier
Dim FormDataFichier() As String
'
'contient le type de donnees
Dim FormDataType() As String
'
'contient les donnees
Dim FormDataDonnees() As String
'
'contient le tout a envoyer
Dim FormDataEnvoi() As String
'
'
'taille des donnees recues
Public RecTailleData As Long
'
'taille des donnees totales a recevoir
Public RecTailleTotaleData As Long
'
'
'taille du chunk en cour
Dim ChunkTaille As Long
'taille totale des chunks connus reunis
Dim ChunkTailleTotale As Long
'taille des donnees recues
Dim ChunkTailleRecu As Long
'taille des donnees restantes du chunk en cour
Dim ChunkReste As Long
'les donnees restantes du chunk en cour
Dim ChunkResteData As String
'fin de la reception des chunks ?
Public ChunkRecFin As Boolean
'
'
'occupe a traiter les donnees ?
Public traiterDonneesOccupe As Boolean
'
'
'les etats de notre classe
Public Enum etats
    '
    SOCKET_FERME = 0
    SOCKET_OUVERT = 1
    SOCKET_RECEPTION = 2
    '
End Enum
'
'l etat de notre classe
Dim etat As etats
'
'
'EVENEMENTS DE LA CLASSE
Event reception(Data As String, TailleRecu As Long, TailleTotaleRecu As Long, Fini As Boolean)
Event connection()
Event fermeture()
Event erreur(Data As String)
Event timeout()
Event headers(Nom As String, Data As String)
'
'
'************************************************
'* EVENEMENTS class
'************************************************
'
'initialisation
Private Sub Class_Initialize()
    '
    'on initialise quelques variables
    Set cs = New CSocket
    '
    Set csTmrReception = New Minuteur
    Set csTmrTimeOut = New Minuteur
    '
    csTmrReception.Intervalle = 100
    csTmrTimeOut.Intervalle = 5000
    '
    initVars
    '
End Sub
'
'de-initialisation
Private Sub Class_Terminate()
    '
    On Error Resume Next
    '
    cs.CloseSocket
    '
    Set cs = Nothing
    '
    Set csTmrReception = Nothing
    Set csTmrTimeOut = Nothing
    '
End Sub
'
'************************************************
'* EVENEMENTS class
'************************************************
'
'
'initialise les variables de base
Public Sub initVars()
    '
    MethodeRequete = "GET"
    VersionProtocole = "HTTP/1.1"
    '
    FermerSocketApresEnvoi = False
    bRecHeaders = False
    '
    ReDim EnvHeaders(0 To 0)
    ReDim RecHeaders(0 To 0)
    '
    ReDim FormDataNom(0 To 0)
    ReDim FormDataDonnees(0 To 0)
    ReDim FormDataType(0 To 0)
    ReDim FormDataFichier(0 To 0)
    ReDim FormDataEnvoi(1 To 2)
    '
    etat = SOCKET_FERME
    '
    'on ajoute qques headers de base
    'host sera remplace par apres, je le met ici pour qu il soit en debut de liste
    ajouterHeader "Host", ""
    ajouterHeader "User-Agent", "cHttpClient"
    'ajouterHeader "Accept", "text/plain"
    'ajouterHeader "Accept-Language", "en-us,en;q=0.5"
    'ajouterHeader "Accept-Encoding", "gzip,deflate"
    'ajouterHeader "Accept-Charset", "ISO-8859-1,utf-8;q=0.7,*;q=0.7"
    'ajouterHeader "Keep-Alive", "300"
    'ajouterHeader "Connection", "keep-alive"
    ajouterHeader "Connection", "Close"
    '
End Sub
'
'
'************************************************
'* FONCTIONS header envoi
'************************************************
'
'ajoute un header a envoyer
Public Sub ajouterHeader(Nom As String, Data As String, Optional retourLigne As Boolean = True)
    '
    Dim iH As Integer
    '
    'on verifie si ce nom existe deja dans la liste
    iH = rechHeader(Nom)
    '
    If iH <= 0 Then
        '
        'non, on cree une nouvelle entree
        iH = UBound(EnvHeaders) + 1
        '
        ReDim Preserve EnvHeaders(LBound(EnvHeaders) To iH)
        '
    End If
    '
    'on insere les donnees dans le tableau a l endroit specifie
    'on ajoute le retour de ligne a la fin s il le faut
    EnvHeaders(iH) = Nom & ": " & Data & IIf(retourLigne = True, vbCrLf, "")
    '
End Sub
'
'recherche si un header est deja present dans la liste
Private Function rechHeader(Nom As String) As Integer
    '
    Dim i As Integer
    Dim iH As Integer
    '
    iH = UBound(EnvHeaders)
    '
    'on verifie si ce nom existe dans la liste
    For i = 1 To iH
        '
        If LCase(Left(EnvHeaders(i), Len(Nom))) = LCase(Nom) Then
            '
            'oui, on renvoie l index
            rechHeader = i
            '
            'et on quitte la fonction ici
            Exit Function
            '
        End If
        '
    Next
    '
    'n existe pas, on renvoie 0
    rechHeader = 0
    '
End Function
'
'envoie les headers
Private Sub envoiHeaders()
    '
    Dim i As Integer
    Dim iH As Integer
    '
    iH = UBound(EnvHeaders)
    '
    'on parcourt la liste
    For i = 1 To iH
        '
        Env EnvHeaders(i)
        '
    Next
    '
End Sub
'
'************************************************
'* FIN FONCTIONS header envoi
'************************************************
'
'
'************************************************
'* FONCTIONS form data
'************************************************
'
'ajoute un form data a envoyer
Public Sub ajouterFormData(Nom As String, Data As String, Optional Fichier As String = "", Optional TypeData As String = "")
    '
    Dim iH As Integer
    '
    'on verifie si ce nom existe deja dans la liste
    iH = rechFormData(Nom)
    '
    If iH <= 0 Then
        '
        'non, on cree une nouvelle entree
        iH = UBound(FormDataNom) + 1
        '
        ReDim Preserve FormDataNom(LBound(FormDataNom) To iH)
        ReDim Preserve FormDataDonnees(LBound(FormDataDonnees) To iH)
        ReDim Preserve FormDataFichier(LBound(FormDataFichier) To iH)
        ReDim Preserve FormDataType(LBound(FormDataType) To iH)
        '
    End If
    '
    'on insere les donnees dans les tableaux
    FormDataNom(iH) = Nom
    FormDataDonnees(iH) = Data
    FormDataFichier(iH) = Fichier
    FormDataType(iH) = TypeData
    '
End Sub
'
'recherche si un form data est deja present dans la liste
Private Function rechFormData(Nom As String) As Integer
    '
    Dim i As Integer
    Dim iH As Integer
    '
    iH = UBound(FormDataNom)
    '
    'on verifie si ce nom existe dans la liste
    For i = 1 To iH
        '
        If LCase(FormDataNom(i)) = LCase(Nom) Then
            '
            'oui, on renvoie l index
            rechFormData = i
            '
            'et on quitte la fonction ici
            Exit Function
            '
        End If
        '
    Next
    '
    'n existe pas, on renvoie 0
    rechFormData = 0
    '
End Function
'
'prepare les form data pour envoi
Private Sub prepareFormDatas()
    '
    Dim i As Integer
    Dim iH As Integer
    Dim Boundary As String
    Dim l As Long
    Dim gSz As String
    '
    iH = UBound(FormDataNom)
    '
    'on verifie qu on a qque chose
    If iH <= 0 Then Exit Sub
    '
    'ok, on cree notre "Boundary" qui va servir a delimiter les donnees qu on envoie
    Boundary = creerBoundary(14)
    '
    'les guillemets
    gSz = Chr(34)
    '
    'la taille des donnees
    l = 0
    '
    'on parcourt la liste
    For i = 1 To iH
        '
        DoEvents
        '
        'notre premier boundary pour ces donnees, on ajoute "--" devant...
        l = l + ajoutFormDataEnvoi("--" & Boundary & vbCrLf)
        '
        'les infos de ces donnees
        l = l + ajoutFormDataEnvoi("Content-Disposition: form-data; name=" & gSz & FormDataNom(i) & gSz)
        '
        'on verifie s il ne faut pas envoyer un fichier
        If Len(FormDataFichier(i)) > 0 Then
            '
            'oui, on rajoute son nom a la suite
            l = l + ajoutFormDataEnvoi("; filename=" & gSz & FormDataFichier(i) & gSz & vbCrLf)
            '
            'ensuite le type de donnees
            l = l + ajoutFormDataEnvoi("Content-Type: " & FormDataType(i) & vbCrLf & vbCrLf)
            '
            'et enfin les donnees
            l = l + ajoutFormDataEnvoi("fichier: " & FormDataFichier(i))
            '
        Else
            '
            'non, on ajoute simplement les donnees a la suite
            l = l + ajoutFormDataEnvoi(vbCrLf & vbCrLf & FormDataDonnees(i) & vbCrLf)
            '
        End If
        '
        'on verifie s il s agit du dernier element de nos tableau
        If i = iH Then
            '
            'oui, on termine par notre boundary de fin, on ajoute "--" devant et derriere
            l = l + ajoutFormDataEnvoi("--" & Boundary & "--" & vbCrLf)
            '
        End If
        '
    Next
    '
    'on ajoute les infos dans l entete pour "presenter" les donnees "form data"
    FormDataEnvoi(1) = "Content-Type: multipart/form-data; boundary=" & Boundary & vbCrLf
    '
    'la taille de ces donnees
    FormDataEnvoi(2) = "Content-Length: " & l & vbCrLf & vbCrLf
    '
End Sub
'
'ajoute les donnees dans le tableau FormDataEnvoi
Private Function ajoutFormDataEnvoi(Data As String) As Long
    '
    Dim iH As Integer
    Dim l As Long
    '
    iH = UBound(FormDataEnvoi) + 1
    '
    'les 2 premier index de ce tableau sont reserves
    'on redimensionne
    ReDim Preserve FormDataEnvoi(1 To iH)
    '
    FormDataEnvoi(iH) = Data
    '
    'on verifie s il ne s agit pas d un fichier
    If Left(Data, 9) = "fichier: " Then
        '
        'oui, on recupere sa taille + la taille de vbcrlf
        l = FileLen(Mid(Data, 10)) + Len(vbCrLf)
        '
    Else
        '
        'non, pas de fichier
        l = Len(Data)
        '
    End If
    '
    'on renvoie la longueur de ces donnees
    ajoutFormDataEnvoi = l
    '
End Function
'
'envoie les form data
Private Sub envoiFormDatas()
    '
    Dim i As Integer
    Dim iH As Integer
    '
    iH = UBound(FormDataEnvoi)
    '
    'on verifie qu on a qque chose en plus que les 2 premiers headers de presentation
    'de ces form data
    If iH <= 2 Then Exit Sub
    '
    'on parcourt la liste
    For i = 1 To iH
        '
        'on verifie qu il ne faut pas envoyer un fichier
        If Left(FormDataEnvoi(i), 9) = "fichier: " Then
            '
            'si, on l envoie
            envoiFichier Mid(FormDataEnvoi(i), 10)
            '
            'et le retour de ligne de fin
            Env vbCrLf
            '
        Else
            '
            'non, on envoie les infos du tableau
            Env FormDataEnvoi(i)
            '
        End If
        '
    Next
    '
End Sub
'
'envoie un fichier
Private Sub envoiFichier(Fichier As String)
    '
    Dim FichSz As Integer
    Dim sData As String
    Dim TailleTotale As Long
    Dim tailleRestante As Long
    Dim tailleEnvoi As Long
    '
    FichSz = FreeFile
    '
    Open Fichier For Binary As FichSz
    '
    'on recupere la taille totale et la taille restante
    TailleTotale = LOF(FichSz)
    tailleRestante = TailleTotale
    '
    'on envoie les donnees par bout
    Do
        '
        DoEvents
        '
        'on specifie la taille des donnees a envoyer
        tailleEnvoi = MaxDonneesEnvoi
        '
        'on verifie si ce n est pas trop grand
        If tailleEnvoi > tailleRestante Then tailleEnvoi = tailleRestante
        '
        'on recupere les donnees
        sData = Space(tailleEnvoi)
        Get FichSz, , sData
        '
        'et on les envoie
        Env sData
        '
        'on soustrait la taille des donnees envoyees a la taille des donnees restantes
        tailleRestante = tailleRestante - tailleEnvoi
        '
        'on verifie si on a termine
        If tailleRestante <= 0 Then Exit Do
        '
    Loop
    '
    Close FichSz
    '
End Sub
'
'cree un boundary aleatoire
Private Function creerBoundary(Max As Integer)
    '
    Dim i As Integer
    Dim s As String
    '
    Randomize
    '
    For i = 1 To Max
        '
        s = s & Round(Rnd() * 9)
        '
    Next
    '
    creerBoundary = "---------------------------" & s
    '
End Function
'
'************************************************
'* FIN FONCTIONS form data
'************************************************
'
'
'PROCEDURE PRINCIPALE (CONNEXION)
Public Function connecter(Serveur As String, Port As Long, Page As String, Methode As String) As Boolean
    '
    'on verifie si on est deja connecte
    If etat <> SOCKET_FERME Then
        '
        connecter = False
        '
        Exit Function
        '
    End If
    '
    'on stocke les infos nécessaires dans les variables
    ServeurSz = Serveur
    PortSz = Port
    '
    PageRequete = Page
    MethodeRequete = Methode
    '
    'on initialise qques variables
    ChunkTaille = 0
    ChunkTailleTotale = 0
    ChunkTailleRecu = 0
    ChunkReste = 0
    ChunkResteData = ""
    ChunkRecFin = False
    '
    RecTailleTotaleData = 0
    RecTailleData = 0
    '
    ResteRecptTemp = ""
    '
    FermerSocketApresEnvoi = False
    '
    'on se connecte
    cs.Connect Serveur, Port
    '
    'on lance le timer timeout
    csTmrTimeOut.Actif = True
    '
    connecter = True
    '
End Function
'
'deconnexion
Public Sub deconnecter()
    '
    'on stoppe les timers
    csTmrReception.Actif = False
    csTmrTimeOut.Actif = False
    '
    cs.CloseSocket
    '
    etat = SOCKET_FERME
    '
    RaiseEvent fermeture
    '
End Sub
'
'
'************************************************
'* EVENEMENTS CSocket
'************************************************
'
Private Sub Cs_OnClose()
    '
    'on specifie que notre socket est ferme
    'il se peut que notre timer de reception des donnees soit toujours actif
    'c est la fonction de traitement des donnees qui arretera ou pas ce timer
    'et enverra l evenement "fermeture" pour eviter les double fermetures
    etat = SOCKET_FERME
    '
End Sub
'
Private Sub Cs_OnConnect()
    '
    'on stoppe le timer timeout
    csTmrTimeOut.Actif = False
    '
    'on est connecté
    etat = SOCKET_OUVERT
    '
    RaiseEvent connection
    '
    'on n a pas encore recu les headers
    bRecHeaders = False
    '
    'on est pret a envoyer
    PretEnvoiSz = True
    '
    'on n est pas occupe a traiter les donnees
    traiterDonneesOccupe = False
    '
    'on envoie la premiere ligne de la requete
    Env MethodeRequete & " " & PageRequete & " " & VersionProtocole & vbCrLf
    '
    'on ajoute le header "host"
    ajouterHeader "Host", ServeurSz
    '
    'on envoie les headers
    envoiHeaders
    '
    'on prepare les form data
    prepareFormDatas
    '
    'et on les envoie
    envoiFormDatas
    '
    Env vbCrLf
    '
    'FermerSocketApresEnvoi = True
    '
End Sub
'
Private Sub Cs_OnConnectionRequest(ByVal requestID As Long)
    '
    'pas utilisé ici
    '
End Sub
'
Private Sub Cs_OnDataArrival(ByVal bytesTotal As Long)
    '
    etat = SOCKET_RECEPTION
    '
    Dim DonnSz As String
    '
    cs.GetData DonnSz, , bytesTotal
    '
    ResteRecptTemp = ResteRecptTemp & DonnSz
    '
    'on active notre timer de traitement des donnees
    csTmrReception.Actif = True
    '
End Sub
'
Private Sub Cs_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '
    cs.CloseSocket
    '
    etat = SOCKET_FERME
    '
    RaiseEvent erreur(Description)
    '
End Sub
'
Private Sub Cs_OnSendComplete()
    '
    'le socket est à nouveau prêt à envoyer des données
    PretEnvoiSz = True
    '
    'on verifie s il ne faut pas fermer le socket
    'If FermerSocketApresEnvoi = True Then Cs.CloseSocket
    '
End Sub
'
Private Sub Cs_OnSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    '
    'pas utilisé
    '
End Sub
'
'************************************************
'* FIN DES EVENEMENTS CSocket
'************************************************
'
'
'************************************************
'* TRAITEMENT DES DONNEES RECUES
'************************************************
'
'(premiere etape) on traite les donnees brutes recues
Public Sub traiterDonneesBrutes()
    '
    'occupe ?
    If traiterDonneesOccupe = True Then Exit Sub
    '
    'maintenant on l est
    traiterDonneesOccupe = True
    '
    '**********************************************************
    '
    'C est ici qu on separe les "headers" du reste des donnees.
    'Les "headers" (entetes) d un message html sont separes par
    'deux vbcrlf.
    'Une fois ces deux vbcrlf detectes, on traite le reste des
    'donnees separement
    '
    '**********************************************************
    '
    'on verifie si on est connecte, sinon on a rien a faire ici
    'If etat = SOCKET_FERME Then Exit Sub
    '
    Dim sTmp As String
    '
    Dim Bloc1 As String
    Dim Bloc2 As String
    Dim pos As Long
    '
    Dim SeparateurSz As String
    '
    SeparateurSz = vbCrLf & vbCrLf
    '
    'on recupere les donnees sauvegardees
    Bloc1 = ResteRecptTemp
    '
    'on efface le reste
    ResteRecptTemp = ""
    '
    'on verifie si on a bien qque chose a traiter, sinon on passe a la fin
    If Len(Bloc1) <= 0 Then GoTo VERIF_REC_TOT
    '
    'on verifie si on a deja recu la totalite des headers
    If bRecHeaders = False Then
        '
        'non, on recherche notre separateur
        pos = InStr(1, Bloc1, SeparateurSz)
        '
        If pos > 0 Then
            '
            'ok, on l a, on separe ces headers du reste des donnees
            Bloc2 = Left(Bloc1, pos - 1)
            Bloc1 = Right(Bloc1, Len(Bloc1) - pos - Len(SeparateurSz) + 1)
            '
            bRecHeaders = True
            '
            'et on les envoie pour traitement
            'on rajoute un vbcrlf pour que notre procedure detecte plus facilement
            'le dernier header
            traiterHeaders Bloc2 & vbCrLf
            '
        Else
            '
            'aucun séparateur n'a été trouvé, on sauvegarde les donnees recues
            'avec celles sauvegardees entre-temps
            ResteRecptTemp = Bloc1 & ResteRecptTemp
            '
            'et on quite la procédure
            GoTo FIN_PROCEDURE
            '
        End If
        '
    End If
    '
    'les headers ont ete recus, on traite ces donnees en fonction de differents facteurs
    If bRecHeaders = True Then
        '
        If TransferEncoding = "chunked" Then
            '
            'on recoit les donnees en mode "chunked", cad bout par bout.
            'chaque bout commence par une ligne specifiant la taille de ce bout
            '
            'on verifie si on a recu tout le chunk et plus
            If Len(Bloc1) > ChunkReste And ChunkReste > 0 Then
                '
                'on recupere la fin du chunk
                sTmp = Left(Bloc1, ChunkReste)
                '
                'le reste des donnees (le chunk suivant)
                Bloc1 = Right(Bloc1, Len(Bloc1) - ChunkReste)
                '
                'on reinitialise qques variables
                ChunkReste = 0
                ChunkTaille = 0
                '
            End If
            '
            'on verifie s il faut traiter un nouveau chunk
            If ChunkReste = 0 Then
                '
                Bloc1 = recChunkTaille(Bloc1)
                '
                'on verifie si ca n est pas la fin...
                If ChunkRecFin = True Then
                    '
                    'Exit Sub
                    '
                End If
                '
            End If
            '
            'on verifie s il faut comptabiliser le reste des donnees a recevoir
            If ChunkReste > 0 Then
                '
                'pas encore recu toutes les donnees, on deduit la taille des donnees
                'recues du reste du chunk a recevoir
                ChunkReste = ChunkReste - Len(Bloc1)
                '
            End If
            '
            'on verifie s il ne faut pas ajouter la fin du chunk precedent a ce debut de chunk
            If Len(sTmp) > 0 Then Bloc1 = sTmp & Bloc1
            '
        End If
        '
        'on comptabilise la taille totale des donnees recues jusqu a present
        RecTailleData = RecTailleData + Len(Bloc1)
        '
        'on envoie les donnees traitees
        RaiseEvent reception(Bloc1, RecTailleData, RecTailleTotaleData, ChunkRecFin)
        '
        '
        '
VERIF_REC_TOT:
        '
        'on verifie si on a encore des donnees a traiter
        If Len(ResteRecptTemp) <= 0 Then
            '
            'non, on verifie si on a recu toutes les donnees
            If RecTailleData = RecTailleTotaleData And ChunkRecFin = True Then
                '
                'oui, on ferme notre socket ici
                cs.CloseSocket
                '
                etat = SOCKET_FERME
                '
                'on stoppe le timer
                csTmrReception.Actif = False
                '
                'on informe de la fin
                RaiseEvent fermeture
                '
            Else
                '
                'on verifie si notre socket est en reception
                If etat <> SOCKET_RECEPTION Then
                    '
                    'non, on stoppe le timer
                    csTmrReception.Actif = False
                    '
                    'on informe de la fin
                    RaiseEvent fermeture
                    '
                End If
                '
            End If
            '
            GoTo FIN_PROCEDURE
            '
        End If
        '
        '
        '
    End If
    '
FIN_PROCEDURE:
    '
    'on n est plus occupe
    traiterDonneesOccupe = False
    '
End Sub
'
'recupere la taille du chunk et renvoie les donnees restantes
Private Function recChunkTaille(Data As String) As String
    '
    Dim sRetour As String
    Dim tmp As String
    Dim Taille As Long
    '
    Dim Bloc2 As String
    Dim pos As Long
    '
    'on enleve les caracteres indesirables du debut de la chaine
    'on ne prend que les 50 premiers caracteres de la chaine pour
    'eviter un depassement...
    Data = lTrimS(Left(Data, 50), vbCrLf & vbTab & " ") & Mid(Data, 51)
    '
    'on verifie si on a un retour de ligne
    pos = InStr(1, Data, vbCrLf)
    '
    If pos > 0 Then
        '
        'oui, on separe la taille du chunk du reste des donnees
        Bloc2 = Left(Data, pos - 1)
        Data = Right(Data, Len(Data) - pos - Len(vbCrLf) + 1)
        '
        'on epure notre taille de tout caractere inutile (Cr, Lf, espaces...)
        Bloc2 = lTrimS(Bloc2, vbCrLf & vbTab & " ")
        Bloc2 = rTrimS(Bloc2, vbCrLf & vbTab & " ")
        '
        'on converti le code hex en dec
        ChunkTaille = hex2dec(Bloc2)
        '
        'Debug.Print "chunk: " & ChunkTaille
        '
        'on verifie la taille de ce chunk
        If ChunkTaille <= 0 Then
            '
            'fini, plus de chunk
            ChunkRecFin = True
            '
        Else
            '
            'un nouveau chunk
            '
            'on met a jour la taille totale des donnees a recevoir
            RecTailleTotaleData = RecTailleTotaleData + ChunkTaille
            '
            'on verifie que ce nouveau chunk n en contient pas un autre
            If Len(Data) > ChunkTaille Then
                '
                'si, on recupere le chunk connu
                sRetour = Left(Data, ChunkTaille)
                '
                'et on sauve l autre partie dans notre variable globale
                ResteRecptTemp = Right(Data, Len(Data) - ChunkTaille) & ResteRecptTemp
                '
                'on reinitialise qques variables
                ChunkReste = 0
                ChunkTaille = 0
                '
            Else
                '
                'pas d autre chunk,on recupere les donnees
                sRetour = Data
                '
                'on definit le reste a recuperer
                ChunkReste = ChunkTaille
                '
            End If
            '
        End If
        '
    Else
        '
        'non, on sauve le reste
        ResteRecptTemp = Data & ResteRecptTemp
        '
    End If
    '
    recChunkTaille = sRetour
    '
End Function
'
'convertion de hex vers dec
Public Function hex2dec(Data As String) As Long
    '
    Dim total As Long
    Dim l As Long
    Dim i As Integer
    '
    For i = 1 To Len(Data)
        '
        'DoEvents
        '
        total = total * 16
        '
        'on recupere la position de ce caractere dans la table ascii
        l = Asc(Mid(Data, i, 1))
        '
        'on verifie si c est un chiffre
        If l < 58 Then
            '
            'oui, on soustrait 48 de sorte a retrouver le nombre initial
            '1 dans la table ascii se trouve a la position 49
            l = l - 48
            '
        Else
            '
            'non, c est une lettre (a,b,c,d,e,f)
            'soustraction pour retrouver la valeur correspondante en hex
            l = l - 87
            '
        End If
        '
        'on additionne
        total = total + l
        '
    Next
    '
    'et on renvoie la valeur
    hex2dec = total
    '
End Function
'
'on separe les differents headers et on traite les infos recues
Private Sub traiterHeaders(Data As String, Optional SeparateurSz As String = vbCrLf)
'    '
    '**********************************************************
    '
    'on separe les differents headers et on les stocke dans un
    'tableau
    '
    '*********************************************************
    '
    Dim Bloc1 As String
    Dim Bloc2 As String
    Dim pos As Long
    Dim iH As Integer
    '
    'les donnees a traiter
    Bloc1 = Data
    '
    'on fait une boucle jusqu'à ce qu'on arrive à la fin du bloc
    pos = 1
    Do Until pos = 0
        '
        'DoEvents
        '
        'on cherche la position de notre séparateur
        pos = InStr(1, Bloc1, SeparateurSz)
        '
        'si la position est > 0 c'est qu'il l'a trouvé
        If pos > 0 Then
            '
            'on sépare la première partie du reste du bloc
            Bloc2 = Left(Bloc1, pos - 1)
            Bloc1 = Right(Bloc1, Len(Bloc1) - pos - Len(SeparateurSz) + 1)
            '
            'on ajoute un element a notre tableau
            iH = UBound(RecHeaders) + 1
            '
            ReDim Preserve RecHeaders(LBound(RecHeaders) To iH)
            RecHeaders(iH) = Bloc2
            '
            'on traite le premier header (reponse code et msg)
            If iH = 1 Then
                '
                traiterCodeReponse Bloc2
                '
            Else
                '
                'on traite ce header normalement
                traiterSimpleHeader Bloc2
                '
            End If
            '
        Else
            '
            'on a fini
            Exit Do
            '
        End If
        '
    Loop
    '
End Sub
'
'traite le premier header (reponse code et msg)
Private Sub traiterCodeReponse(Data As String)
    '
    'on recupere le code de reponse et le msg http
    Dim sTmp() As String
    '
    sTmp = Split(Data, " ", 3)
    '
    If UBound(sTmp) > 0 Then
        '
        ReponseHTTP = sTmp(1)
        '
    End If
    
    If UBound(sTmp) > 1 Then
        '
        ReponseMsg = sTmp(2)
        '
    End If
    '
    'on envoie ces infos
    RaiseEvent headers("c_http_client_code", ReponseHTTP)
    RaiseEvent headers("c_http_client_msg", ReponseMsg)
    '
End Sub
'
'traite un header simple
Private Sub traiterSimpleHeader(Data As String)
    '
    Dim sTmp() As String
    Dim pos As Long
    '
    sTmp = Split(Data, ":")
    '
    'on verifie qu on a bien 2 elements (au-)
    If UBound(sTmp) > 0 Then
        '
        'on epure les infos
        sTmp(0) = Trim(sTmp(0))
        sTmp(1) = Trim(sTmp(1))
        '
        'on traite ce header
        Select Case sTmp(0)
            '
            'la taille totale des donnees a recevoir
            Case "Content-Length"
                '
                RecTailleTotaleData = CLng(sTmp(1))
                '
                'specifie que la taille totale est definitive
                ChunkRecFin = True
                '
            'type de transfert
            Case "Transfer-Encoding"
                '
                TransferEncoding = sTmp(1)
                '
                'specifie que la taille totale n est pas definitive
                ChunkRecFin = False
                '
            '
        End Select
        '
        'on informe de la reception de ce header
        RaiseEvent headers(sTmp(0), sTmp(1))
        '
    End If
    '
End Sub
'
'recupere la valeur d un header
Public Function recValHeader(Nom As String) As String
    '
    Dim i As Integer
    Dim iH As Integer
    '
    iH = UBound(RecHeaders)
    '
    For i = 1 To iH
        '
        If LCase(Left(RecHeaders(i), Len(Nom))) = LCase(Nom) Then
            '
            recValHeader = Mid(RecHeaders(i), Len(Nom) + 3)
            '
            Exit Function
            '
        End If
        '
    Next
    '
    recValHeader = ""
    '
End Function
'
'ENVOI DES DONNEES
Private Sub Env(Data As String)
    '
    Form1.txtEnvData = Form1.txtEnvData & Data
    '
    'Exit Sub
    '
    'on vérifie si le socket n'est pas fermé
    If cs.State = sckClosed Then Exit Sub
    If cs.State = sckClosing Then Exit Sub
    '
    'on vérifie si le socket est prêt à envoyer des données
    If PretEnvoiSz = False Then
        '
        'non, donc on attend
        '
        Do Until PretEnvoiSz = True
            '
            DoEvents
            '
            'on vérifie quand même si entre-temps, le socket ne s'est pas fermé
            If cs.State = sckClosed Then Exit Sub
            If cs.State = sckClosing Then Exit Sub
            '
        Loop
        '
    End If
    '
    PretEnvoiSz = False
    '
    'on envoie les données
    cs.SendData Data
    '
End Sub
'
'on traite les donnees
Private Sub csTmrReception_Action()
    '
    traiterDonneesBrutes
    '
End Sub
'
'timeout
Private Sub csTmrTimeOut_Action()
    '
    'on stoppe les timers
    csTmrReception.Actif = False
    csTmrTimeOut.Actif = False
    '
    'on ferme la connexion
    cs.CloseSocket
    '
    etat = SOCKET_FERME
    '
    RaiseEvent timeout
    '
End Sub
