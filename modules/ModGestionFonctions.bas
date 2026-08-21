Attribute VB_Name = "ModGestionFonctions"
'******************************************************************************
'***                                                                           ***
'******************************************************************************
'***    FORM:                                                                                              ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    PROGRAMMEUR:                                                                              ***
'***      Royer Jean-Laurent                                                                         ***
'******************************************************************************

'******************************************************************************
'***    MODIF :                                                                                            ***
'******************************************************************************
Option Explicit

'******************************************************************************
'***   Declaration De Constante Public                                                         ***
'******************************************************************************

'******************************************************************************
'***    Declaration De Fonction Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function InStrRev(ByVal StrChaine As String, ByVal StrChaineRechercher As String, Optional DebutRechercher, Optional Compare) As Integer
   Dim IntChaineLongueur As Integer
   Dim IntChaineRechercherLongueur As Integer
   Dim IntPosRechercher As Integer
   Dim IntDebutRechercher As Integer
   Dim IntCompare As Integer

   On Error GoTo Err_InStrRev

   IntChaineLongueur = Len(StrChaine)

   IntChaineRechercherLongueur = Len(StrChaineRechercher)

   Select Case IsMissing(DebutRechercher)
      Case True

         IntDebutRechercher = IntChaineLongueur
      Case False

         Select Case VarType(DebutRechercher)
            Case vbByte, vbInteger, vbLong

               IntDebutRechercher = DebutRechercher

            Case Else

               IntDebutRechercher = IntChaineLongueur

         End Select
   End Select

   Select Case IsMissing(Compare)
      Case True

         IntCompare = vbEmpty
      Case False

         Select Case VarType(Compare)
            Case vbByte, vbInteger, vbLong

               Select Case Compare
                  Case Is < vbEmpty

                     MsgBox "Parametre Compare Incorrect", , "InStrRev"

                     GoTo Exit_InStrRev

                  Case Is > 2

                     MsgBox "Parametre Compare Incorrect", , "InStrRev"

                     GoTo Exit_InStrRev

                  Case Else

                     IntCompare = Compare

                  End Select
            Case Else

               IntCompare = vbEmpty

         End Select
   End Select

   Select Case IntCompare
      Case vbEmpty

      Case 1

         StrChaine = UCase(StrChaine)

         StrChaineRechercher = UCase(StrChaineRechercher)

      Case 2

         ' Il Faut Recuperer Le Type De Comparaisons Indiquer Par Les Parametres De Comparaison De La Base
         ' A faire ulterieurement
   End Select

   Select Case IntChaineLongueur
      Case vbEmpty

         InStrRev = vbEmpty

         GoTo Exit_InStrRev

      Case Is < IntDebutRechercher

         InStrRev = vbEmpty

         GoTo Exit_InStrRev

      Case Else
         Select Case IntChaineRechercherLongueur
            Case vbEmpty

               InStrRev = IntDebutRechercher

               GoTo Exit_InStrRev

            Case Is > IntChaineLongueur

               InStrRev = vbEmpty

               GoTo Exit_InStrRev

            Case Is > IntDebutRechercher

               InStrRev = vbEmpty

               GoTo Exit_InStrRev

            Case Else

               For IntPosRechercher = IntDebutRechercher To 1 Step -1

                  Select Case Mid(StrChaine, IntPosRechercher, IntChaineRechercherLongueur)
                     Case StrChaineRechercher

                        InStrRev = IntPosRechercher

                        GoTo Exit_InStrRev

                     Case Else

                  End Select
               Next

               InStrRev = vbEmpty

            End Select
   End Select

Exit_InStrRev:

   Exit Function

Err_InStrRev:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "InStrRev"

   Resume Exit_InStrRev
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function RemplaceChr(ByVal StrChaine As String, ByVal StrChaineRechercher As String, ByVal strChaineRemplace As String, Optional Compare) As String
   Dim StrChaineResultat As String
   Dim IntPos As Integer
   Dim IntCompare As Integer

   On Error GoTo Err_RemplaceChr

   If StrChaine = vbNullString Then

       RemplaceChr = vbNullString

       GoTo Exit_RemplaceChr

   End If

   Select Case IsMissing(Compare)
      Case True

         Compare = 1
      Case False

         Select Case VarType(Compare)
            Case vbByte, vbInteger, vbLong

                  Select Case Compare
                     Case Is < vbEmpty

                        MsgBox "Parametre Compare Incorrect", , "RemplaceChr"

                        GoTo Exit_RemplaceChr

                     Case Is > 2

                        MsgBox "Parametre Compare Incorrect", , "RemplaceChr"

                        GoTo Exit_RemplaceChr


                     Case Else

                        IntCompare = Compare

                  End Select
            Case Else

               IntCompare = vbEmpty

         End Select
   End Select

   StrChaineResultat = vbNullString

   IntPos = InStr(1, StrChaine, StrChaineRechercher, IntCompare)

   While (IntPos > vbEmpty)

       StrChaineResultat = StrChaineResultat + Left(StrChaine, IntPos - 1) & strChaineRemplace

       StrChaine = Right(StrChaine, Len(StrChaine) - Len(StrChaineRechercher) - IntPos + 1)

       IntPos = InStr(1, StrChaine, StrChaineRechercher, IntCompare)

   Wend

   RemplaceChr = StrChaineResultat + StrChaine

Exit_RemplaceChr:

   Exit Function

Err_RemplaceChr:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "RemplaceChr"

   Resume Exit_RemplaceChr
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function Split(ByVal StrChaine As String, ByVal StrChaineDelimiter As String) As Variant
   Dim StrChaineTableau() As String
   Dim IntChainePosition As Integer
   Dim IntTableauIndex As Integer

   On Error GoTo Err_Split

   IntTableauIndex = 0

   IntChainePosition = InStr(1, StrChaine, StrChaineDelimiter)

   Select Case IntChainePosition
      Case 0

         Split = StrChaine

      Case Else

         Do

            ReDim Preserve StrChaineTableau(0 To IntTableauIndex)

            StrChaineTableau(IntTableauIndex) = Left(StrChaine, IntChainePosition - 1)

            StrChaine = Mid(StrChaine, IntChainePosition + 1)

            IntChainePosition = InStr(1, StrChaine, StrChaineDelimiter)

            IntTableauIndex = IntTableauIndex + 1

            Select Case IntChainePosition
               Case 0

                  ReDim Preserve StrChaineTableau(0 To IntTableauIndex)

                  StrChaineTableau(IntTableauIndex) = StrChaine

               Case Else

            End Select

         Loop Until IntChainePosition = 0

   End Select

   Split = StrChaineTableau

Exit_Split:

   Exit Function

Err_Split:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Split"

   Resume Exit_Split
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function CreeRepertoire(ByVal StrRepertoire As String) As Boolean
   Dim StrRepertoireEnCours As String
   Dim IntPos As Integer

   On Error GoTo Err_CreeRepertoire

   CreeRepertoire = True

   IntPos = 1

   Do While Dir(StrRepertoire, vbDirectory) = vbNullString

      IntPos = InStr(IntPos + 1, StrRepertoire, "\")

      Select Case IntPos
         Case 0

            MkDir StrRepertoire

         Case Else

            Select Case Dir(Left(StrRepertoire, IntPos - 1), vbDirectory)
               Case vbNullString

                  MkDir Left(StrRepertoire, IntPos - 1)

               Case Else

            End Select
      End Select
   Loop

Exit_CreeRepertoire:

   Exit Function

Err_CreeRepertoire:

   CreeRepertoire = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeRepertoire"

   Resume Exit_CreeRepertoire
End Function
