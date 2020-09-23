Attribute VB_Name = "mStrings"
'
' qques fonctionalites en + pour les strings
'
'************************************************************************
'  fonction: lTrimS
'------------------------------------------------------------------------
' Enleve les caracteres de dTrim de la chaine Data (cote gauche)
' Les caracteres de dTrim sont compares un a un
' avec les caracteres de Data.
'
' ex:
' lTrimS("aacaToto", "ca")
'
' renvoie:
' "Toto"
'------------------------------------------------------------------------
Public Function lTrimS(Data As String, dTrim As String)
    '
    Dim l As Long
    Dim l2 As Long
    '
    l = Len(Data)
    l2 = Len(dTrim)
    '
    If l <= 0 Or l2 <= 0 Then Exit Function
    '
    Dim i As Integer
    Dim i2 As Integer
    '
    Dim b As Boolean
    '
    Dim s As String
    '
    'on verifie chaque caractere de Data
    For i = 1 To l
        '
        DoEvents
        '
        b = False
        '
        'qu on compare avec chaque caractere de dTrim
        For i2 = 1 To l2
            '
            DoEvents
            '
            If Mid(Data, i, 1) = Mid(dTrim, i2, 1) Then
                '
                'ok, on en a trouve au moins un, on quitte cette boucle ici
                b = True
                '
                Exit For
                '
            End If
            '
        Next
        '
        'en a t on trouve un ? si non, on quitte cette boucle aussi
        If b = False Then Exit For
        '
    Next
    '
    If l - i > 0 Then
        '
        s = Mid(Data, i, l - i + 1)
        '
    End If
    '
    lTrimS = s
    '
End Function
'
'
'************************************************************************
'  fonction: rTrimS
'------------------------------------------------------------------------
' Enleve les caracteres de dTrim de la chaine Data (cote droite).
' Les caracteres de dTrim sont compares un a un
' avec les caracteres de Data.
'
' ex:
' rTrimS("Totoaaca", "ca")
'
' renvoie:
' "Toto"
'------------------------------------------------------------------------
Public Function rTrimS(Data As String, dTrim As String)
    '
    Dim l As Long
    Dim l2 As Long
    '
    l = Len(Data)
    l2 = Len(dTrim)
    '
    If l <= 0 Or l2 <= 0 Then Exit Function
    '
    Dim i As Integer
    Dim i2 As Integer
    '
    Dim b As Boolean
    '
    Dim s As String
    '
    i = l
    '
    'on verifie chaque caractere de Data
    Do
        '
        DoEvents
        '
        b = False
        '
        'qu on compare avec chaque caractere de dTrim
        For i2 = 1 To l2
            '
            DoEvents
            '
            If Mid(Data, i, 1) = Mid(dTrim, i2, 1) Then
                '
                'ok, on en a trouve au moins un, on quitte cette boucle ici
                b = True
                '
                Exit For
                '
            End If
            '
        Next
        '
        'en a t on trouve un ? si non, on quitte cette boucle aussi
        If b = False Then Exit Do
        '
        i = i - 1
        '
        If i <= 0 Then Exit Do
        '
    Loop
    '
    If i > 0 Then
        '
        s = Left(Data, i)
        '
    End If
    '
    rTrimS = s
    '
End Function
