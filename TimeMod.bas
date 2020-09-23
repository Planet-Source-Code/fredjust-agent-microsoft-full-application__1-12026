Attribute VB_Name = "TimeMod"
'==================================================================================
'   Réalisation de Frédéric Just
'   Commentaires remarques et critiques :
'
'   adresse en cours    : fred.just@free.fr
'   site actuel         : http://www.fredjust.com
'   adresse de secours  : fredjust@hotmail.com
'==================================================================================


Option Explicit

Public ChaineTime As String

Public merlin As IAgentCtlCharacterEx

Public Frequence As Long

Public NextMessageDate As Variant

Public ChaineBonjour As String

Public ChaineFin As String

Public FRMbulleVisible As Boolean

Public FRMmessageVisible As Boolean

Public chaineagent As String

Public lastleft As Long
Public lasttop As Long

Public InMem As Boolean

Public IndexLigne As Long

Public SaveAs As String

Public Const CouleurBulle = &HCCFFFF

Public ShowAbout As Boolean


'==================================================================================
' maximun de 2 nombres
Public Function Max(ByVal a, ByVal b As Long) As Long

    If a > b Then Max = a Else Max = b
End Function

'==================================================================================
' minimun de 2 nombres
Public Function Min(ByVal a, ByVal b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function


'==================================================================================
'   CREER UN FICHIER TEXTE
'   RETOURNE 0 SI TOUT C EST BIEN PASSER LE NUMERO DE L ERR SINON
'==================================================================================
Public Function CreerFichierTexte(ByVal NomFichier As String, ByRef TextStream) As Long

Dim FileSystemObject, LeFile

On Error GoTo gestion_erreur

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set TextStream = FileSystemObject.CreateTextFile(NomFichier)
    CreerFichierTexte = 0
    
Exit Function

gestion_erreur:
    CreerFichierTexte = Err.Number
End Function

'==================================================================================
'   OUVRE UN FICHIER TEXTE EXISTANT
'   EN ECRITURE PAR DEFAUT
'   PARAMETRES :
'       NOM DU FICHIER AVEC SON CHEMIN
'       TEXTSTREAM DANS LEQUEL SERA PLACER LE CONTENU DU FICHIER
'   PARAMETRES OPTIONELS :
'       LECTURESEUL SI TRUE ECRITURE IMPOSIBLE
'       POURAJOUR  SI TRUE SE PLACE A LA FIN
'   RETOURNE 0 SI TOUT C EST BIEN PASSER LE NUMERO DE L ERR SINON
'==================================================================================
Public Function OuvreFichierTexte(ByVal NomFichier As String, ByRef TextStream, _
                        Optional ByVal LectureSeul As Boolean, _
                        Optional ByVal PourAjout As Boolean) As Long
Dim FileSystemObject, LeFile

On Error GoTo gestion_erreur

    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set LeFile = FileSystemObject.GetFile(NomFichier)
    If LectureSeul Then
        Set TextStream = LeFile.OpenAsTextStream(1)
    Else
        If PourAjout Then
            Set TextStream = LeFile.OpenAsTextStream(8)
        Else
            Set TextStream = LeFile.OpenAsTextStream(2)
        End If
    End If
    
    OuvreFichierTexte = 0
Exit Function

gestion_erreur:
    OuvreFichierTexte = Err.Number
End Function


'==================================================================================
'   FERME LE FICHIER TEXTSTREAM
'   RETOURNE 0 SI TOUT C EST BIEN PASSER LE NUMERO DE L ERR SINON
'==================================================================================
Public Function FermeFichier(ByRef TextStream) As Long
On Error GoTo gestion_erreur

    TextStream.Close
    FermeFichier = 0
    
Exit Function
gestion_erreur:
    FermeFichier = Err.Number
End Function

'==================================================================================
'   ENREGISTRE LE CONTENU D UNE LISTE VIEW DANS UN FICHIER TEXTE
'   CHAQUE COLONNE EST SEPARE PAR UNE TABULATION
'   EXCEL PEUT ENREGISTRER ET OUVRIR SOUS CE FORMAT !
'   RETOURNE 0 SI TOUT C EST BIEN PASSER LE NUMERO DE L ERR SINON
'==================================================================================
Public Function EcritContenuListViewDansFichier(ByVal ListView As Object, ByRef NomFichier As String) As Long
    Dim i As Long
    Dim j As Long
    Dim tempo As String
    Dim Result As Long
    Dim TextStream As TextStream

    On Error GoTo gestion_erreur

    Result = CreerFichierTexte(NomFichier, TextStream)
    TextStream.WriteLine "[TimeAgent Version 1.0]"
    If Result = 0 Then
        For i = 1 To ListView.ListItems.Count
            tempo = ListView.ListItems.Item(i).Text
            For j = 1 To ListView.ColumnHeaders.Count - 1
                tempo = tempo & Chr(9) & ListView.ListItems.Item(i).SubItems(j)
            Next j
            'tempo = tempo & Chr(9) & IIf(ListView.ListItems.Item(i).Checked, "1", "0")
            TextStream.WriteLine tempo
        Next i

        EcritContenuListViewDansFichier = FermeFichier(TextStream)

    Else
        EcritContenuListViewDansFichier = Result
    End If

    Exit Function
gestion_erreur:
    EcritContenuListViewDansFichier = Err.Number
End Function

'==================================================================================
'   REMPLIS UNE LISTVIEW AVEC LE CONTENU D UN FICHIER TEXTE SEPARATEUR TABULATION
'   RETOURNE 0 SI TOUT C EST BIEN PASSER LE NUMERO DE L ERR SINON
'==================================================================================
Public Function LisContenuLISTVIEWdepuisFichier(ListView As Object, ByRef NomFichier As String)
    Dim i As Long
    Dim j As Long
    Dim TextStream 'as textstream
    Dim Result As Long
    Dim NbLigne As Long
    Dim NbColonne As Long
    Dim LigneTexte
    Dim ColonneTexte As String

    On Error GoTo gestion_erreur

    ListView.ListItems.Clear
    Result = OuvreFichierTexte(NomFichier, TextStream, True)

    If Result = 0 Then
        i = 1
        j = 1
        If TextStream.ReadLine <> "[TimeAgent Version 1.0]" Then
            merlin.Speak "Sorry It's not a valid TimeAgent file !"
            MsgBox "Sorry It's not a valid TimeAgent file !", vbCritical, "TimeAgent Error !"
            Exit Function
        End If
        Do
            LigneTexte = TextStream.ReadLine
            ColonneTexte = Mid(LigneTexte, 1, InStr(1, LigneTexte, Chr(9), vbTextCompare) - 1)
            LigneTexte = Replace(LigneTexte, ColonneTexte + Chr(9), "", , 1)
            ListView.ListItems.Add , , ColonneTexte

            While InStr(1, LigneTexte, Chr(9), vbTextCompare) <> 0
                If ListView.ColumnHeaders.Count <= j Then ListView.ColumnHeaders.Add , , j
                ColonneTexte = Mid(LigneTexte, 1, InStr(1, LigneTexte, Chr(9), vbTextCompare) - 1)
                ListView.ListItems.Item(i).SubItems(j) = ColonneTexte
                LigneTexte = Replace(LigneTexte, ColonneTexte + Chr(9), "", , 1)
                j = j + 1
            Wend

            If ListView.ColumnHeaders.Count <= j Then ListView.ColumnHeaders.Add , , j
            ListView.ListItems.Item(i).SubItems(j) = LigneTexte
            ListView.ListItems.Item(i).Checked = IIf(LigneTexte = "1", True, False)

            j = 1
            i = i + 1

        Loop Until TextStream.AtEndOfStream

        LisContenuLISTVIEWdepuisFichier = FermeFichier(TextStream)

    Else
        LisContenuLISTVIEWdepuisFichier = Result
    End If
    FRMbulle.ChercheMessageSuivant
    Exit Function
gestion_erreur:
    LisContenuLISTVIEWdepuisFichier = Err.Number
End Function

'==================================================================================
'
'==================================================================================
Public Sub SelectAll(TXT As TextBox)
    TXT.SelStart = 0
    TXT.SelLength = Len(TXT)
End Sub


'==================================================================================
'   classe les colonnes d'une ListView en fonction de la colonne clicker
'   a appeler dans l'evenement ColumnClick
'   ATTENTION utilise le TAG de la ListView
Public Sub ClasseLesColonnes(UneListView As ListView, COLONNE As MSComctlLib.ColumnHeader)
    With UneListView
        .Sorted = True
        If .Tag = COLONNE.index Then   'si on click sur la meme colonne
            If .SortOrder = lvwAscending Then 'inversion de l'ordre de classement
               .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
            .SortKey = COLONNE.index - 1
        Else
            .SortOrder = lvwAscending 'classe sur cette colonne et par ordre
            .SortKey = COLONNE.index - 1
        End If
        .Tag = COLONNE.index   'stock la derniere colonne clicker dans le TAG
    End With
End Sub


'==================================================================================
'   AFFICHE LA BOITE D OUVERTURE DE FICHIER
'   RENVOIE LE NOM AINSI QUE LE CHEMIN DU FICHIER CHOISI
'   EN CAS D ANNULATION RENVOIE ""
'   PARAMETRES OPTIONELS :
'       LES FILTRES : *.* PAR DEFAUT, A PASSER AINSI : "TXT HTML BAT"
'       LE NOM DU FICHIER SANS LE CHEMIN
'       LE TITRE DE LA BOITE D OUVERTURE
'       NUMERO DE L ERREUR QUI C EST PRODUIT
'==================================================================================
Public Function OuvrirFichierExistant(Optional ByVal TypeFile As String, _
                                    Optional File As String, _
                                    Optional ByVal DialogTitle As String, _
                                    Optional ErrNumber As Long) As String
Dim tempo As String
Dim Extention As String
Dim BoiteOuverture As CommonDialog

Set BoiteOuverture = FRMmerlin.CommonDialog1

On Error GoTo ErrHandler
    BoiteOuverture.Flags = cdlOFNHideReadOnly
    BoiteOuverture.Filter = ""
    BoiteOuverture.CancelError = True
    
    If DialogTitle <> "" Then BoiteOuverture.DialogTitle = DialogTitle
    
    If TypeFile = "" Then
        BoiteOuverture.Filter = "Tous les fichiers (*.*)|*.*"
    Else
        If InStr(1, TypeFile, " ", vbTextCompare) = 0 Then
            BoiteOuverture.Filter = "Fichiers " + UCase(TypeFile) + "|*." + UCase(TypeFile)
        Else
            tempo = UCase(TypeFile)
            While InStr(1, tempo, " ", vbTextCompare) <> 0
                Extention = Mid(tempo, 1, InStr(1, tempo, " ", vbTextCompare) - 1)
                BoiteOuverture.Filter = BoiteOuverture.Filter & "|Fichiers " + Extention + "|*." + Extention
                tempo = Replace(tempo, Extention + " ", "")
            Wend
            BoiteOuverture.Filter = BoiteOuverture.Filter & "|Fichiers " + tempo + "|*." + tempo
            BoiteOuverture.Filter = Mid(BoiteOuverture.Filter, 2)
        End If
    End If
    
    BoiteOuverture.FilterIndex = 1
    BoiteOuverture.ShowOpen
    OuvrirFichierExistant = BoiteOuverture.FileName
    File = BoiteOuverture.FileTitle
    ErrNumber = 0
    Exit Function
ErrHandler:
    OuvrirFichierExistant = ""
    ErrNumber = Err.Number
End Function

'==================================================================================
'   AFFICHE LA BOITE DE SAUVEGARDE D UN FICHIER
'   RENVOIE LE NOM AINSI QUE LE CHEMIN DU FICHIER CIBLE
'   EN CAS D ANNULATION RENVOIE ""
'   PARAMETRES OPTIONELS :
'       LES FILTRES : *.* PAR DEFAUT, A PASSER AINSI : "TXT HTML BAT"
'       LE NOM DU FICHIER SANS LE CHEMIN
'       LE TITRE DE LA BOITE DE SAUVEGARDE
'       NUMERO DE L ERREUR QUI C EST PRODUIT
'==================================================================================
Public Function EnregistrerSous(Optional ByVal TypeFile As String, _
                            Optional File As String, _
                            Optional ByVal DialogTitle As String, _
                            Optional ErrNumber As Long) As String
Dim tempo As String
Dim Extention As String
Dim BoiteOuverture As CommonDialog

Set BoiteOuverture = FRMmerlin.CommonDialog1

On Error GoTo ErrHandler
    BoiteOuverture.Flags = cdlOFNHideReadOnly
    BoiteOuverture.Filter = ""
    BoiteOuverture.CancelError = True
    
    If DialogTitle <> "" Then BoiteOuverture.DialogTitle = DialogTitle
    If TypeFile = "" Then
        BoiteOuverture.Filter = "Tous les fichiers (*.*)|*.*"
    Else
        If InStr(1, TypeFile, " ", vbTextCompare) = 0 Then
            BoiteOuverture.Filter = "Fichiers " + UCase(TypeFile) + "|*." + UCase(TypeFile)
        Else
            tempo = UCase(TypeFile)
            While InStr(1, tempo, " ", vbTextCompare) <> 0
                Extention = Mid(tempo, 1, InStr(1, tempo, " ", vbTextCompare) - 1)
                BoiteOuverture.Filter = BoiteOuverture.Filter & "|Fichiers " + Extention + "|*." + Extention
                tempo = Replace(tempo, Extention + " ", "")
            Wend
            BoiteOuverture.Filter = BoiteOuverture.Filter & "|Fichiers " + tempo + "|*." + tempo
            BoiteOuverture.Filter = Mid(BoiteOuverture.Filter, 2)
        End If
    End If
    
    BoiteOuverture.FilterIndex = 1
    BoiteOuverture.ShowSave
    EnregistrerSous = BoiteOuverture.FileName
    File = BoiteOuverture.FileTitle
    ErrNumber = 0
    Exit Function
ErrHandler:
    EnregistrerSous = ""
    ErrNumber = Err.Number
End Function



