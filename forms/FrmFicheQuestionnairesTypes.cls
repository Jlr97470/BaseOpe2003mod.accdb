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
'***    Variable De Controle De Mode De Saisie                                            ***
'******************************************************************************
Private MLngPerSocQstTypNum As Long
Private MLngPerSocQstTypNbResultat As Long

'******************************************************************************
'***    Declaration De Procedure Privée                                                       ***
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

               Set CtlActive = ActiveControl

               Select Case Err.Number
                  Case 0

                  Case Else

                      Err.Clear

                      Set CtlActive = TxtPerSocQstTypCode

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         SFrmFicheQuestionnairesTypesQuestionnairesTypesID.Requery

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
'***    Declaration D'Evenement Privée                                                        ***
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

   ChangeBoutonMode Me, MODEDEPLACEMENT

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

   MLngPerSocQstTypNum = Nz(PerSocQstTypNum)

   MLngPerSocQstTypNbResultat = Nz(PerSocQstTypNbResultat)

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
      Case MODEDEPLACEMENT

         Select Case Me.RecordsetClone.RecordCount
            Case 0

            Case Else

                Select Case PerSocQstTypNbResultat
                  Case Is = SFrmFicheQuestionnairesTypesQuestionnairesTypesID.Form.RecordsetClone.RecordCount

                  Case Else

                     Cancel = True

                     MsgBox "Attention Le Nombre De Résultat Definie Attendu Est Incorrect !!!", , Me.Caption

                End Select

         End Select

      Case Else

         Cancel = True

         MsgBox "Vous Devez Validez Ou Annulez La Modification En Cours !", , Me.Caption

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

   Select Case IsNull(PerSocQstTypCode)
      Case False

         TxtPerSocQstTypCode = UCase(TxtPerSocQstTypCode)

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaire Type Code Est Obligatoire !", , Me.Caption

         TxtPerSocQstTypCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(PerSocQstTypDateDebut)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Questionnaire Type Date Debut Est Obligatoire !", , Me.Caption

         TxtPerSocQstTypDateDebut.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(PerSocQstTypDateFin)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Questionnaire Type Date Fin Est Obligatoire !", , Me.Caption

         TxtPerSocQstTypDateFin.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case PerSocQstTypDateDebut <= PerSocQstTypDateFin
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Questionnaire Type Date Debut Doit Etre Inferieur A Questionnaires Types Date Fin !", , Me.Caption

         TxtPerSocQstTypDateDebut.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocQstTypLibelle)
      Case False

         TxtPerSocQstTypLibelle = UCase(TxtPerSocQstTypLibelle)

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaire Type Libelle Est Obligatoire !", , Me.Caption

         TxtPerSocQstTypLibelle.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocQstTypNbResultat) Or PerSocQstTypNbResultat = 0
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaire Type Nombre Resultats Est Invalide !", , Me.Caption

         TxtPerSocQstTypNbResultat.SetFocus

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

   MLngPerSocQstTypNum = PerSocQstTypNum

   MLngPerSocQstTypNbResultat = PerSocQstTypNbResultat

   TxtPerSocQstTypCode.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   TxtPerSocQstTypCode.SetFocus

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
   Dim RsQuestionnairesTypes As DAO.Recordset
   Dim RsQuestionnairesTypesID As DAO.Recordset
   Dim StrQuestionnairesTypesID As String

   On Error GoTo Err_Form_Current

   Select Case GIntMode
      Case MODEDEPLACEMENT

         StrQuestionnairesTypesID = "SELECT TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPESID.* " & _
            "FROM TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPESID WHERE " & "PerSocQstTypIDPerSocQstTypNum=" & MLngPerSocQstTypNum & ";"

         Set RsQuestionnairesTypesID = CurrentDb.OpenRecordset(StrQuestionnairesTypesID)

         Select Case RsQuestionnairesTypesID.EOF
            Case True

            Case False

               RsQuestionnairesTypesID.MoveLast

         End Select

         Select Case MLngPerSocQstTypNbResultat
            Case Is = RsQuestionnairesTypesID.RecordCount

               MLngPerSocQstTypNum = PerSocQstTypNum

               MLngPerSocQstTypNbResultat = PerSocQstTypNbResultat

            Case Else

               Select Case PerSocQstTypNum
                  Case MLngPerSocQstTypNum

                  Case Else

                     Set RsQuestionnairesTypes = Me.RecordsetClone

                     RsQuestionnairesTypes.FindFirst "PerSocQstTypNum=" & MLngPerSocQstTypNum

                     Me.Bookmark = RsQuestionnairesTypes.Bookmark

                     MsgBox "Attention Le Nombre De Résultat Definie Attendu Est Incorrect !!!", , Me.Caption

                     RsQuestionnairesTypes.Close

               End Select

         End Select

         RsQuestionnairesTypesID.Close

      Case Else

   End Select

   Rafraichir (RAFRAICHIRDEPLACEMENT)

Exit_Form_Current:

   Set RsQuestionnairesTypes = Nothing

   Set RsQuestionnairesTypesID = Nothing

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
   Dim TxtQstTypDateDebut As String

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         TxtPerSocQstTypCode.SetFocus

         Select Case PerSocQstTypNbResultat
           Case Is = SFrmFicheQuestionnairesTypesQuestionnairesTypesID.Form.RecordsetClone.RecordCount

               Case Else

               MsgBox "Attention Le Nombre De Résultat Definie Attendu Est Incorrect !!!", , Me.Caption

               GoTo Exit_CmdAjouter_Click

          End Select

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtQstTypDateDebut = TxtPerSocQstTypDateDebut

   TxtPerSocQstTypDateDebut = "01/01/2003"

   TxtPerSocQstTypDateDebut = TxtQstTypDateDebut

   TxtPerSocQstTypCode.SetFocus

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
   Dim TxtQstTypDateDebut As String

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerSocQstTypCode.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtQstTypDateDebut = TxtPerSocQstTypDateDebut

         TxtPerSocQstTypDateDebut = "01/01/2003"

         TxtPerSocQstTypDateDebut = TxtQstTypDateDebut

         TxtPerSocQstTypCode.SetFocus

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

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerSocQstTypCode.SetFocus

         ChangeBoutonMode Me, MODESUPPRIME

         On Error Resume Next

         DoCmd.RunCommand acCmdDeleteRecord

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Select Case Err.Number
                  Case 2501

                  Case Else

                     MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSupprimer_Click"

               End Select

               Err.Clear

         End Select

         On Error GoTo Err_CmdSupprimer_Click

         ChangeBoutonMode Me, MODEDEPLACEMENT

          Select Case Me.RecordsetClone.RecordCount
             Case vbEmpty

             Case Else

                TxtPerSocQstTypCode.SetFocus

          End Select

   End Select

   MLngPerSocQstTypNum = Nz(PerSocQstTypNum)

   MLngPerSocQstTypNbResultat = Nz(PerSocQstTypNbResultat)

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

   TxtPerSocQstTypCode.SetFocus

   ChangeBoutonMode Me, MODEVALIDE

   On Error Resume Next

   DoCmd.RunCommand acCmdSaveRecord

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Select Case Err.Number
            Case 2501, 2759

            Case Else

               MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdValider_Click"

         End Select

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

   TxtPerSocQstTypCode.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerSocQstTypCode.SetFocus

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

   DoCmd.Close acForm, Me.Name, acSavePrompt

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Err.Clear

   End Select

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

   On Error GoTo Err_CmdRechercher_Click

   OuvreRecherches 2

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
Private Sub CmdMinimize_Click()

   On Error GoTo Err_CmdMinimize_Click

   DoCmd.Minimize

Exit_CmdMinimize_Click:

   Exit Sub

Err_CmdMinimize_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdMinimize_Click"

   Resume Exit_CmdMinimize_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocQstTypNbResultat_Exit(Cancel As Integer)

    On Error GoTo Err_TxtPerSocQstTypNbResultat_Exit

    Select Case PerSocQstTypNbResultat
      Case Is >= SFrmFicheQuestionnairesTypesQuestionnairesTypesID.Form.RecordsetClone.RecordCount

      Case Else

         Cancel = True

         MsgBox "Attention Le Nombre De Résultat Definie Attendu Est Superieur !!!", vbCritical, Me.Caption

    End Select

Exit_TxtPerSocQstTypNbResultat_Exit:

    Exit Sub

Err_TxtPerSocQstTypNbResultat_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocQstTypNbResultat_Exit"

    Resume Exit_TxtPerSocQstTypNbResultat_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocQstTypCode_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocQstTypCode_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocQstTypCode_KeyPress:

    Exit Sub

Err_TxtPerSocQstTypCode_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocQstTypCode_KeyPress"

    Resume Exit_TxtPerSocQstTypCode_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocQstTypLibelle_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocQstTypLibelle_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocQstTypLibelle_KeyPress:

    Exit Sub

Err_TxtPerSocQstTypLibelle_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocQstTypLibelle_KeyPress"

    Resume Exit_TxtPerSocQstTypLibelle_KeyPress
End Sub
