Attribute VB_Name = "ModMinuteur"
Option Explicit
'
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'
'DECLARATION DE LA COLLECTION QUI VA STOCKER L'ID DE CHAQUE CLASSE QUI APPELLERA UN MINUTEUR
Public IDClsMinuteur As New Collection
'
Public Sub AjoutColl(IDClasse As Long, IDMinuteur As Long)
    '
    'on ajoue un �l�ment dans la liste qui va permettre d'identifier la classe
    IDClsMinuteur.Add IDClasse, "M" & IDMinuteur
    '
End Sub

Public Sub EnleveColl(IDClasse As Long)
    '
    'on enl�ve un �l�ment
    '
    Dim i As Integer
    '
    For i = 1 To IDClsMinuteur.Count
        If Str(IDClsMinuteur.Item(i)) = Str(IDClasse) Then
            IDClsMinuteur.Remove i
            Exit For
        End If
    Next
    '
End Sub

Public Sub MinuteurProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal iEvent As Long, ByVal iTime As Long)
    '
    On Error Resume Next
    '
    Dim cM As Minuteur
    Dim hLng As Long
    '
    'on r�cup�re le num�ro d'identification de la classe correspondante
    hLng = CLng(IDClsMinuteur.Item("M" & iEvent))
    '
    If hLng = 0 Then Exit Sub
    '
    'on cr�e une copie parfaite de la classe voulue
    'la copie et l'originale d�pendent l'une de l'autre, si un changement est effectu� � l'une, il est d'office effectu� � l'autre
    'c'est gr�ce � cela que l'on va pouvoir lancer une fontion (ou proc�dure) de cette classe
    CopyMemory cM, hLng, 4&
    '
    'on lace la proc�dure proc�dure
    cM.LancementAction
    '
    'on "efface" la copie en inscrivant rien
    CopyMemory cM, 0&, 4
    '
    If Err Then Exit Sub
    '
End Sub
