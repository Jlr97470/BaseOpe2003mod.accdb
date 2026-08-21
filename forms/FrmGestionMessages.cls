Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'***    Declaration De Variable Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable De Controle De Mode De Saisie                                            ***
'******************************************************************************
Public GIntMode As Integer
Public GIntModeOld As Integer

'******************************************************************************
'***    Declaration De Variable Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable                                                                                           ***
'******************************************************************************
Private MColObject As New Collection

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
Public Function TesteMessagerie() As Boolean
   Dim RsMachinesMessages As DAO.Recordset
   Dim StrEval As String
   Dim StrEvalResultat As String
   Dim StrSQLReponse As String

   On Error GoTo Err_TesteMessagerie

   TesteMessagerie = True

   Set RsMachinesMessages = CurrentDb.OpenRecordset("SELECT TBLMACHINESMESSAGES.* FROM TBLMACHINESMESSAGES WHERE (MacMesUtilisateur='" & _
      GStrUtilisateurNom & "' OR MacMesUtilisateur Is Null) AND (MacMesDomaine='" & _
      GStrOrdinateurDomaine & "' OR MacMesDomaine Is Null) AND (MacMesNom='" & _
      GStrOrdinateurNom & "' OR MacMesNom Is Null) AND ((((Now()-[MacMesDate]-" & _
      "[MacMesHeure])>#12/30/1899#)=True) AND (((Now()-[MacMesDate]-[MacMesHeure])<#12/30/1899 0:0:03#)=True));")

   Select Case RsMachinesMessages.EOF
      Case True

      Case False

         Select Case RsMachinesMessages!MacMesAgent
            Case True

               MsgBox Nz(RsMachinesMessages!MacMesSuject, vbNullString), , Nz(RsMachinesMessages!MacMesObject, vbNullString)

            Case False

               Me.Visible = True

               Me.Filter = "MacMesNumero=" & RsMachinesMessages!MacMesNumero

         End Select

         Select Case RsMachinesMessages!MacMesCommande
            Case 0

            Case 1

               StrEval = RsMachinesMessages!MacMesSuject

               StrEval = RemplaceChr(StrEval, "'", "'''")

               StrEval = RemplaceChr(StrEval, Chr$(34), "'")

               Select Case Left(StrEval, 3)
                  Case "Set"

                     StrEval = Mid(StrEval, 3)

                  Case Else

                     On Error Resume Next

                     StrEvalResultat = Eval(StrEval)

                     Select Case StrEvalResultat
                        Case vbNullString

                           StrEvalResultat = "AUCUNNE REPONSE"

                        Case Else

                     End Select

                     StrSQLReponse = "INSERT INTO TBLMACHINESMESSAGES ( MacMesExpUtilisateur, MacMesExpNom, MacMesExpDomaine,  MacMesUtilisateur, " & _
                                                  "MacMesNom, MacMesDomaine, MacMesObject, MacMesSuject, MacMesDate, MacMesHeure, MacMesAgent ) VALUES ('" & _
                                                  GStrUtilisateurNom & "','" & GStrOrdinateurNom & "','" & GStrOrdinateurDomaine & "','" & Nz(RsMachinesMessages!MacMesExpUtilisateur, "Null") & _
                                                  "','" & Nz(RsMachinesMessages!MacMesExpNom, "Null") & "','" & Nz(RsMachinesMessages!MacMesExpDomaine, "Null") & "','EVAL RESULTAT:" & _
                                                  "','" & StrEvalResultat & "',Date(),Time()+CDate('00:00:03'),False);"

                     CurrentDb.Execute StrSQLReponse

                     On Error GoTo Err_TesteMessagerie

               End Select

         End Select

   End Select

   RsMachinesMessages.Close

   Set RsMachinesMessages = Nothing

Exit_TesteMessagerie:

   Exit Function

Err_TesteMessagerie:

   TesteMessagerie = False

   MsgBox Err.Number & " " & Err.Description, 0, "TesteMessagerie"

   Resume Exit_TesteMessagerie
End Function

'******************************************************************************
'***    Declaration De Procedure Priver                                                       ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                   ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub Rafraichir(Optional ByVal BytModeRafraichir As Byte)
   Dim CtlActive As Access.Control

   On Error GoTo Err_Rafraichir

   Select Case (Me.RecordsetClone.RecordCount = 0) And (Me.NewRecord = False)
      Case True

      Case False

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

              On Error Resume Next

               Set CtlActive = Me.ActiveControl

               Select Case Err.Number
                  Case 0

                  Case Else

                      Err.Clear

                      Set CtlActive = LstMacMesNom

              End Select

              On Error GoTo Err_Rafraichir

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT, RAFRAICHIRCHOIX

               LstMacMesNom.Requery

               LstMacMesUtilisateur.Requery

               LstMacMesDomaine.Requery

            Case Else

         End Select


         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

              CtlActive.SetFocus

              Set CtlActive = Nothing

            Case Else

         End Select
   End Select

Exit_Rafraichir:

   Exit Sub

Err_Rafraichir:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Rafraichir"

   Resume Exit_Rafraichir
End Sub

'******************************************************************************
'***    Declaration D'Evenement Priver                                                        ***
'******************************************************************************

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   DoCmd.RunCommand acCmdWindowCascade

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Open"

   Resume Exit_Form_Open
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   TxtMacMesExpUtilisateur.DefaultValue = GStrUtilisateurNom

   TxtMacMesExpNom.DefaultValue = GStrOrdinateurNom

   TxtMacMesExpDomaine.DefaultValue = GStrOrdinateurDomaine

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Me.Visible = False

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Load"

   Resume Exit_Form_Load
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

   On Error GoTo Err_Form_Unload

   Select Case GIntMode
      Case MODEDEPLACEMENT, MODENULL

      Case Else

         Cancel = True

         MsgBox "Vous Devez Validez Ou Annulez La Modification En Cours !", vbInformation, Me.Caption

   End Select

Exit_Form_Unload:

   Exit Sub

Err_Form_Unload:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Unload"

   Resume Exit_Form_Unload
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_Form_BeforeUpdate

   Select Case GIntMode
      Case MODEVALIDE

      Case Else

         Cancel = True

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(MacMesExpDate)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Machine Message Expedition Date Est Invalide !", , Me.Caption

         TxtMacMesExpDate.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(MacMesExpHeure)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Machine Message Expedition Heure Est Invalide !", , Me.Caption

         TxtMacMesExpHeure.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(MacMesDate)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Machine Message Date Est Invalide !", , Me.Caption

         TxtMacMesDate.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(MacMesHeure)
      Case True

         MacMesHeure = IIf(MacMesHeure > Time, MacMesHeure, Time) + CDate("00:00:06")

      Case False

         Cancel = True

         MsgBox "Le Champ Machine Message Heure Est Invalide !", , Me.Caption

         TxtMacMesDate.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(MacMesObject)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Machine Message Object Est Obligatoire !", , Me.Caption

         TxtMacMesObject.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case LstMacMesUtilisateur.ItemsSelected.Count <> 0 Or LstMacMesNom.ItemsSelected.Count <> 0 Or LstMacMesDomaine.ItemsSelected.Count <> 0
      Case True

         Select Case LstMacMesNom.ItemsSelected.Count
            Case vbEmpty

            Case Else

               MacMesNom = LstMacMesNom.ItemData(LstMacMesNom.ItemsSelected(0))

         End Select

         Select Case LstMacMesUtilisateur.ItemsSelected.Count
           Case vbEmpty

           Case Else

               MacMesUtilisateur = LstMacMesUtilisateur.ItemData(LstMacMesUtilisateur.ItemsSelected(0))

         End Select

         Select Case LstMacMesDomaine.ItemsSelected.Count
           Case vbEmpty

           Case Else

             MacMesDomaine = LstMacMesDomaine.ItemData(LstMacMesDomaine.ItemsSelected(0))

         End Select

      Case False

         Cancel = True

         MsgBox "Aucun Déstinataire Selectionée !!!", , Me.Caption

         LstMacMesNom.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

Exit_Form_BeforeUpdate:

   Exit Sub

Err_Form_BeforeUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_BeforeUpdate"

   Resume Exit_Form_BeforeUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_AfterUpdate()

   On Error GoTo Err_Form_AfterUpdate

   LstMacMesNom.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   LstMacMesNom.SetFocus

   Rafraichir (RAFRAICHIRVISIBLE)

Exit_Form_AfterUpdate:

   Exit Sub

Err_Form_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_AfterUpdate"

   Resume Exit_Form_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Current()

   On Error GoTo Exit_Form_Current

   Rafraichir (RAFRAICHIRDEPLACEMENT)

Exit_Form_Current:

   Exit Sub

Err_Form_Current:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Current"

   Resume Exit_Form_Current
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Activate()

   On Error GoTo Exit_Form_Activate

   Rafraichir (RAFRAICHIRVISIBLE)

Exit_Form_Activate:

   Exit Sub

Err_Form_Activate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Activate"

   Resume Exit_Form_Activate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouter_Click()
   Dim TxtMessageDate As String

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         LstMacMesNom.SetFocus

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtMessageDate = TxtMacMesDate

   TxtMacMesDate = "01/01/2003"

   TxtMacMesDate = TxtMessageDate

   TxtMacMesExpUtilisateur = GStrUtilisateurNom

   TxtMacMesExpNom = GStrOrdinateurNom

   TxtMacMesExpDomaine = GStrOrdinateurDomaine

   LstMacMesNom.SetFocus

Exit_CmdAjouter_Click:

   Exit Sub

Err_CmdAjouter_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouter_Click"

   Resume Exit_CmdAjouter_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdModifier_Click()
   Dim TxtMessageDate As String

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         LstMacMesNom.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtMessageDate = TxtMacMesDate

         TxtMacMesDate = "01/01/2003"

         TxtMacMesDate = TxtMessageDate

         LstMacMesNom.SetFocus

   End Select

Exit_CmdModifier_Click:

   Exit Sub
Err_CmdModifier_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdModifier_Click"

   Resume Exit_CmdModifier_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdSupprimer_Click()
   Dim LngCadAchSocNum As Long
   Dim LngCadAchNum As Long

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         LstMacMesNom.SetFocus

         ChangeBoutonMode Me, MODESUPPRIME

         On Error Resume Next

         DoCmd.RunCommand acCmdDeleteRecord

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               ChangeBoutonMode Me, GIntModeOld

         End Select

         On Error GoTo Err_CmdSupprimer_Click

         ChangeBoutonMode Me, MODEDEPLACEMENT

         Select Case Me.RecordsetClone.RecordCount
            Case vbEmpty

            Case Else

               LstMacMesNom.SetFocus

         End Select
   End Select

Exit_CmdSupprimer_Click:

   Exit Sub

Err_CmdSupprimer_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSupprimer_Click"

   Resume Exit_CmdSupprimer_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   LstMacMesNom.SetFocus

   ChangeBoutonMode Me, MODEVALIDE

   LstMacMesNom.SetFocus

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Err.Clear

         ChangeBoutonMode Me, GIntModeOld

   End Select

   On Error GoTo Err_CmdValider_Click

Exit_CmdValider_Click:

   Exit Sub

Err_CmdValider_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdValider_Click"

   Resume Exit_CmdValider_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAnnuler_Click()

   On Error GoTo Err_CmdAnnuler_Click

   LstMacMesNom.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         LstMacMesNom.SetFocus

   End Select

Exit_CmdAnnuler_Click:

   Exit Sub
Err_CmdAnnuler_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAnnuler_Click"

   Resume Exit_CmdAnnuler_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   On Error Resume Next

   Me.Visible = False

   On Error GoTo Err_CmdFermer_Click

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAfficherTous_Click()

   On Error Resume Next

   Select Case Me.Filter
      Case vbNullString

      Case Else

         Me.Filter = vbNullString

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear
         End Select

         Me.Requery

   End Select

Exit_CmdAfficherTous_Click:

   Exit Sub

Err_CmdAfficherTous_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherTous_Click"

   Resume Exit_CmdAfficherTous_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdRechercher_Click()

   On Error GoTo Exit_CmdRechercher_Click

   OuvreRecherches 3

Exit_CmdRechercher_Click:

   Exit Sub

Err_CmdRechercher_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdRechercher_Click"

   Resume Exit_CmdRechercher_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstMacMesNom_Click()

   On Error GoTo Err_LstMacMesNom_Click

   LstMacMesDomaine = vbNullString

   LstMacMesUtilisateur = vbNullString

   LstMacMesDomaine.Requery

   LstMacMesUtilisateur.Requery

Exit_LstMacMesNom_Click:

    Exit Sub

Err_LstMacMesNom_Click:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "LstMacMesNom_Click"

    Resume Exit_LstMacMesNom_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstMacMesUtilisateur_Click()

   On Error GoTo Err_LstMacMesUtilisateur_Click

   LstMacMesDomaine = vbNullString

   LstMacMesNom = vbNullString

   LstMacMesDomaine.Requery

   LstMacMesNom.Requery

Exit_LstMacMesUtilisateur_Click:

    Exit Sub

Err_LstMacMesUtilisateur_Click:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "LstMacMesUtilisateur_Click"

    Resume Exit_LstMacMesUtilisateur_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstMacMesDomaine_Click()

   On Error GoTo Err_LstMacMesDomaine_Click

   LstMacMesNom = vbNullString

   LstMacMesUtilisateur = vbNullString

   LstMacMesNom.Requery

   LstMacMesUtilisateur.Requery

Exit_LstMacMesDomaine_Click:

    Exit Sub

Err_LstMacMesDomaine_Click:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "LstMacMesDomaine_Click"

    Resume Exit_LstMacMesDomaine_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Timer()
   Dim LngEnregistrementNumero As Long
   Dim RstEnregistrement As Recordset

   On Error GoTo Err_Form_Timer

   LngEnregistrementNumero = Nz(Me.MacMesNumero)

   Screen.MousePointer = ccArrow

   Me.Requery

   Select Case Me.AllowAdditions = True Or Me.AllowEdits = True Or Me.RecordsetClone.RecordCount = 0
         Case True

         Case False

             Select Case TesteMessagerie
               Case True

               Case False

            End Select

            Select Case LngEnregistrementNumero
               Case vbEmpty

               Case Else

                  Set RstEnregistrement = Me.RecordsetClone

                  RstEnregistrement.FindFirst "MacMesNumero=" & LngEnregistrementNumero

                  Me.Bookmark = RstEnregistrement.Bookmark

                  Set RstEnregistrement = Nothing

            End Select

   End Select

   DoCmd.Hourglass False

Exit_Form_Timer:

    Exit Sub

Err_Form_Timer:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Timer"

    Resume Exit_Form_Timer
End Sub
