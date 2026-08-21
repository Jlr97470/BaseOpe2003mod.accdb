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
'***    Variable De Controle De Mode Saisie                                                 ***
'******************************************************************************
Public GIntMode As Integer
Public GIntModeOld As Integer

'******************************************************************************
'***    Declaration De Variable Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Pour Le Rafraichissement                                                    ***
'******************************************************************************
Private MTxtPerCivCode As String
Private MTxtPerVilNum As String

'******************************************************************************
'***    Declaration Procedure Priver                                                             ***
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
   Dim StrSQLPerCivCode As String
   Dim StrSQLPerVilNum As String
   Dim StrSQLPerTel1Libelle As String
   Dim StrSQLPerTel2Libelle As String
   Dim CtlActive As Control

   On Error GoTo Err_Rafraichir

   StrSQLPerCivCode = "SELECT CivCode, CivLibelle FROM TBLCIVILITES ORDER BY CivCode;"

   StrSQLPerVilNum = "SELECT VilNum, VilCodePostal, VilNom, DepNom, DepPays FROM SelVillesDepartementsCodes LEFT JOIN TBLDEPARTEMENTS ON SelVillesDepartementsCodes.VilDepCode=TBLDEPARTEMENTS.DepCode ORDER BY VilCodePostal;"

   StrSQLPerTel1Libelle = "SELECT DISTINCT PerTel1Libelle FROM TBLPERSONNELS ORDER BY PerTel1Libelle;"

   StrSQLPerTel2Libelle = "SELECT DISTINCT PerTel2Libelle FROM TBLPERSONNELS ORDER BY PerTel2Libelle;"

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

                      Set CtlActive = TxtPerNum

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               CmbPerCivCode.RowSource = StrSQLPerCivCode

               Select Case MTxtPerCivCode
                  Case vbNullString

                     CmbPerCivCode.Requery

                  Case Else

                     CmbPerCivCode.Requery

                     CmbPerCivCode.Text = MTxtPerCivCode

                     MTxtPerCivCode = vbNullString

               End Select

               CmbPerVilNum.RowSource = StrSQLPerVilNum

               Select Case MTxtPerVilNum
                  Case vbNullString

                     CmbPerVilNum.Requery

                  Case True

                     CmbPerVilNum.Requery

                     CmbPerVilNum.Text = MTxtPerVilNum

                     MTxtPerVilNum = vbNullString

               End Select

            Case Else

         End Select

         Select Case OgtFichePersonnelsDetails
            Case 0

            Case 1

               Select Case BytModeRafraichir
                  Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

                     LstAnniversaires.Requery

                  Case Else

               End Select

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
'***    Declaration Procedure Public                                                            ***
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
Public Sub ChangeSFrmButtonMode(ByRef FrmSousForm As Form)

   On Error GoTo Err_Rafraichir

   CmdSFrmAjouter.Enabled = FrmSousForm.CmdAjouter.Enabled

   CmdSFrmModifier.Enabled = FrmSousForm.CmdModifier.Enabled

   CmdSFrmSupprimer.Enabled = FrmSousForm.CmdSupprimer.Enabled

   CmdSFrmAnnuler.Enabled = FrmSousForm.CmdAnnuler.Enabled

   CmdSFrmValider.Enabled = FrmSousForm.CmdValider.Enabled

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

   OgtFichePersonnelsDetails_Change

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

   Select Case IsDate(PerDateCre)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Personnels Date De Creation Est Invalide !", , Me.Caption

         GoTo Exit_Form_BeforeUpdate

   End Select

   PerDateModif = Date

   Select Case IsNull(PerCivCode)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Personnels Civilitées Est Obligatoire !", , Me.Caption

         CmbPerCivCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerNom)
      Case False

         TxtPerNom = UCase(TxtPerNom)

      Case True

         Cancel = True

         MsgBox "Le Champ Personnels Nom Est Obligatoire !", , Me.Caption

         TxtPerNom.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerPrenom)
      Case False

         TxtPerPrenom = UCase(TxtPerPrenom)

      Case True

         Cancel = True

         MsgBox "Le Champ Personnels Prenom Est Obligatoire !", , Me.Caption

         TxtPerPrenom.SetFocus

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

   TxtPerNum.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   TxtPerNum.SetFocus

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
   Dim TxtDateCreer As String

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         TxtPerNum.SetFocus

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateCreer = TxtPerDateCre

   TxtPerDateCre = "01/01/2003"

   TxtPerDateCre = TxtDateCreer

   TxtPerNum.SetFocus

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
   Dim TxtDateCreer As String

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerNum.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateCreer = TxtPerDateCre

         TxtPerDateCre = "01/01/2003"

         TxtPerDateCre = TxtDateCreer

         TxtPerNum.SetFocus

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

         TxtPerNum.SetFocus

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

                TxtPerNum.SetFocus

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

   TxtPerNum.SetFocus

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

   TxtPerNum.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerNum.SetFocus

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
Private Sub CmdAfficherSocietes_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherSocietes_Click

   DonneEtatFormulaire "FrmFicheSocietes", IntEtat

   Select Case IntEtat
      Case VUEFORMULAIRE

         Forms!FrmFicheSocietes.SetFocus

         DoCmd.ApplyFilter , "SocNum=" & SFrmFichePersonnelsPersonnelsSocietes.Form!PerSocSocNum

         DoCmd.RunCommand acCmdWindowCascade

      Case Else

         DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & SFrmFichePersonnelsPersonnelsSocietes.Form!PerSocSocNum, acFormEdit

   End Select

Exit_CmdAfficherSocietes_Click:

   Exit Sub

Err_CmdAfficherSocietes_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherSocietes_Click"

   Resume Exit_CmdAfficherSocietes_Click
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
Private Sub CmdSFrmAjouter_Click()

   On Error GoTo Err_CmdSFrmAjouter_Click

   Select Case OgtFichePersonnelsDetails
      Case 0

         SFrmFichePersonnelsPersonnelsSocietes.SetFocus

         SFrmFichePersonnelsPersonnelsSocietes.Form.CmdAjouter_Click

         ChangeSFrmButtonMode SFrmFichePersonnelsPersonnelsSocietes.Form

      Case Else

   End Select

Exit_CmdSFrmAjouter_Click:

   Exit Sub

Err_CmdSFrmAjouter_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSFrmAjouter_Click"

   Resume Exit_CmdSFrmAjouter_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdSFrmAnnuler_Click()

   On Error GoTo Err_CmdSFrmAnnuler_Click

   Select Case OgtFichePersonnelsDetails
      Case 0

         SFrmFichePersonnelsPersonnelsSocietes.SetFocus

         SFrmFichePersonnelsPersonnelsSocietes.Form.CmdAnnuler_Click

         ChangeSFrmButtonMode SFrmFichePersonnelsPersonnelsSocietes.Form

      Case Else

   End Select

Exit_CmdSFrmAnnuler_Click:

   Exit Sub

Err_CmdSFrmAnnuler_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSFrmAnnuler_Click"

   Resume Exit_CmdSFrmAnnuler_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdSFrmModifier_Click()

   On Error GoTo Err_CmdSFrmModifier_Click

   Select Case OgtFichePersonnelsDetails
      Case 0

         SFrmFichePersonnelsPersonnelsSocietes.SetFocus

         SFrmFichePersonnelsPersonnelsSocietes.Form.CmdModifier_Click

         ChangeSFrmButtonMode SFrmFichePersonnelsPersonnelsSocietes.Form

      Case Else

   End Select

Exit_CmdSFrmModifier_Click:

   Exit Sub

Err_CmdSFrmModifier_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSFrmModifier_Click"

   Resume Exit_CmdSFrmModifier_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdSFrmSupprimer_Click()

   On Error GoTo Err_CmdSFrmSupprimer_Click

   Select Case OgtFichePersonnelsDetails
      Case 0

         SFrmFichePersonnelsPersonnelsSocietes.SetFocus

         SFrmFichePersonnelsPersonnelsSocietes.Form.CmdSupprimer_Click

         ChangeSFrmButtonMode SFrmFichePersonnelsPersonnelsSocietes.Form

      Case Else

   End Select

Exit_CmdSFrmSupprimer_Click:

   Exit Sub

Err_CmdSFrmSupprimer_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSFrmSupprimer_Click"

   Resume Exit_CmdSFrmSupprimer_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdSFrmValider_Click()

   On Error GoTo Err_CmdSFrmValider_Click

   Select Case OgtFichePersonnelsDetails
      Case 0

         SFrmFichePersonnelsPersonnelsSocietes.SetFocus

         SFrmFichePersonnelsPersonnelsSocietes.Form.CmdValider_Click

         ChangeSFrmButtonMode SFrmFichePersonnelsPersonnelsSocietes.Form

      Case Else

   End Select

Exit_CmdSFrmValider_Click:

   Exit Sub

Err_CmdSFrmValider_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSFrmValider_Click"

   Resume Exit_CmdSFrmValider_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouterCivilites_Click()

   On Error GoTo Err_CmdAjouterCivilites_Click

   AjouteCivilites

Exit_CmdAjouterCivilites_Click:

   Exit Sub

Err_CmdAjouterCivilites_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterCivilites_Click"

   Resume Exit_CmdAjouterCivilites_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdAjouterVilles_Click()

   On Error GoTo Err_CmdAjouterCivilites_Click

   AjouteVilles

Exit_CmdAjouterCivilites_Click:

   Exit Sub

Err_CmdAjouterCivilites_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAjouterCivilites_Click"

   Resume Exit_CmdAjouterCivilites_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerCivCode_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_CmbPerCivCode_NotInList

   AjouteCivilites (NewData)

   Response = acDataErrContinue

   MTxtPerCivCode = CmbPerCivCode.Text

   CmbPerCivCode.Undo

Exit_CmbPerCivCode_NotInList:

   Exit Sub

Err_CmbPerCivCode_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerCivCode_NotInList"

   Resume Exit_CmbPerCivCode_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerVilNum_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_CmbPerVilNum_NotInList

   AjouteVilles (NewData)

   Response = acDataErrContinue

   MTxtPerVilNum = CmbPerVilNum.Text

   CmbPerVilNum.Undo

Exit_CmbPerVilNum_NotInList:

   Exit Sub

Err_CmbPerVilNum_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerVilNum_NotInList"

   Resume Exit_CmbPerVilNum_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub OgtFichePersonnelsDetails_Change()
   Dim StrSQLAnniversaires As String
   Dim StrSFrmFichePersonnelsPersonnelsSocietes As String
   Dim StrSFrmFichePersonnelsOperations As String
   Dim StrSFrmFichePersonnelsQuestionnaires As String

   On Error GoTo Err_OgtFichePersonnelsDetails_Change

   StrSQLAnniversaires = "SELECT PerSocAnnNum, PerSocPerNum, PerSocAnnAnnee, PerSocAnnDateSaisie FROM SelPersonnelsSocietesAnniversaires WHERE SelPersonnelsSocietesAnniversaires.PerSocPerNum=[TxtPerNum];"

   StrSFrmFichePersonnelsPersonnelsSocietes = "SFrmFichePersonnelsPersonnelsSocietes"

   StrSFrmFichePersonnelsOperations = "SFrmFichePersonnelsOperations"

   StrSFrmFichePersonnelsQuestionnaires = "SFrmFichePersonnelsQuestionnaires"

   Select Case OgtFichePersonnelsDetails
      Case 0

         SFrmFichePersonnelsPersonnelsSocietes.SourceObject = StrSFrmFichePersonnelsPersonnelsSocietes

         LstAnniversaires.RowSource = vbNullString

         SFrmFichePersonnelsOperations.SourceObject = vbNullString

         SFrmFichePersonnelsQuestionnaires.SourceObject = vbNullString

         CmdSFrmAjouter.Visible = True

         CmdSFrmModifier.Visible = True

         CmdSFrmSupprimer.Visible = True

      Case 1

         SFrmFichePersonnelsPersonnelsSocietes.SourceObject = vbNullString

         LstAnniversaires.RowSource = StrSQLAnniversaires

         SFrmFichePersonnelsOperations.SourceObject = vbNullString

         SFrmFichePersonnelsQuestionnaires.SourceObject = vbNullString

         CmdSFrmAjouter.Visible = False

         CmdSFrmModifier.Visible = False

         CmdSFrmSupprimer.Visible = False

      Case 2

         SFrmFichePersonnelsPersonnelsSocietes.SourceObject = vbNullString

         LstAnniversaires.RowSource = vbNullString

         SFrmFichePersonnelsOperations.SourceObject = StrSFrmFichePersonnelsOperations

         SFrmFichePersonnelsQuestionnaires.SourceObject = vbNullString

         CmdSFrmAjouter.Visible = False

         CmdSFrmModifier.Visible = False

         CmdSFrmSupprimer.Visible = False

      Case 3

         SFrmFichePersonnelsPersonnelsSocietes.SourceObject = vbNullString

         LstAnniversaires.RowSource = vbNullString

         SFrmFichePersonnelsOperations.SourceObject = vbNullString

         SFrmFichePersonnelsQuestionnaires.SourceObject = StrSFrmFichePersonnelsQuestionnaires

         CmdSFrmAjouter.Visible = False

         CmdSFrmModifier.Visible = False

         CmdSFrmSupprimer.Visible = False

   End Select

   Rafraichir (RAFRAICHIRCHOIX)

Exit_OgtFichePersonnelsDetails_Change:

   Exit Sub

Err_OgtFichePersonnelsDetails_Change:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OgtFichePersonnelsDetails_Change"

   Resume Exit_OgtFichePersonnelsDetails_Change
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerNom_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerNom_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerNom_KeyPress:

    Exit Sub

Err_TxtPerNom_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerNom_KeyPress"

    Resume Exit_TxtPerNom_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerPrenom_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerPrenom_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerPrenom_KeyPress:

    Exit Sub

Err_TxtPerPrenom_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerPrenom_KeyPress"

    Resume Exit_TxtPerPrenom_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerCivCode_GotFocus()

    On Error GoTo Err_CmbPerCivCode_GotFocus

    CmbPerCivCode.Dropdown

Exit_CmbPerCivCode_GotFocus:

    Exit Sub

Err_CmbPerCivCode_GotFocus:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerCivCode_GotFocus"

    Resume Exit_CmbPerCivCode_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerVilNum_GotFocus()

    On Error GoTo Err_CmbPerVilNum_GotFocus

    CmbPerVilNum.Dropdown

Exit_CmbPerVilNum_GotFocus:

    Exit Sub

Err_CmbPerVilNum_GotFocus:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerVilNum_GotFocus"

    Resume Exit_CmbPerVilNum_GotFocus
End Sub
