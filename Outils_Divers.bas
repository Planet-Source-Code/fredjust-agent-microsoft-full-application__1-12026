Attribute VB_Name = "Outils_Divers"
Option Explicit

'**********************************************************************************
'   FONCTION SUR DES OBJETS VISUELS
'**********************************************************************************

'==================================================================================
'   selectionne tout le contenu d'une zone text
'   à appeler dans l'evenement gotfocus
Public Sub SelectAll(ByRef ZoneText As TextBox)
    ZoneText.SelStart = 0
    ZoneText.SelLength = Len(ZoneText)
End Sub

'**********************************************************************************
'   FONCTIONS MATHEMATIQUES
'**********************************************************************************

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
' retourne TRUE si la chaine peut etre converti en un long
' il est possible d'obtenir directement le resultat dans le deuxieme paramètre
Public Function IsLong(ByVal chaine As String, Optional EntierLong As Long) As Boolean
On Error GoTo non
    EntierLong = CLng(chaine)
    IsLong = True
    Exit Function
non:
    IsLong = False
End Function

'==================================================================================
' retourne TRUE si la chaine peut etre converti en un double
' il est possible d'obtenir directement le resultat dans le deuxieme paramètre
Public Function IsDouble(ByVal chaine As String, Optional NBdouble As Double) As Boolean
On Error GoTo non
    NBdouble = CDbl(chaine)
    IsDouble = True
    Exit Function
non:
    IsDouble = False
End Function
'==================================================================================
'
'==================================================================================
Public Function IsCur(ByVal chaine As String, Optional Montant As Currency) As Boolean
    Dim tempo As Currency
On Error GoTo gestion_erreur
    IsCur = True
    Montant = CCur(chaine)
    Exit Function
gestion_erreur:
    IsCur = False
End Function



