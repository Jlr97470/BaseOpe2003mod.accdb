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
Public GBlnModif As Boolean

'******************************************************************************
'***    Declaration De Variable Privée                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Du Numero D'Operations En Cours                                       ***
'******************************************************************************
Private MLngPerSocOpeNum As Long

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
   Dim StrSQLPerSocOpePerSocNum As String
   Dim StrSQLPerSocOpePerSocOpeTypCode As String
   Dim CtlActive As Control

   On Error GoTo Err_Rafraichir

   StrSQLPerSocOpePerSocOpeTypCode = "SELECT PerSocOpeTypNum, PerSocOpeTypCode, PerSocOpeTypLibelle, PerSocOpeTypDateDebut, PerSocOpeTypDateFin, PerSocOpeTypNbResultat " & _
      " FROM SelPersonnelsSocietesOperationsTypes ;"

   StrSQLPerSocOpePerSocNum = "SELECT PerSocNum , PerNum, PerCivCode, PerNom, PerPrenom, PerSocPerSocTypCode, PerSocTypLibelle, PerSocActif, " & _
      " PerSocAdherent, SocNum, SocCode, SocEnseigne, SocRaison, SocAdr, SocVilCodePostal, SocVilNom FROM SelPersonnelsSocietesPersonnelsSocietesTypesActif WHERE PerSocPerSocTypCode In ('VND','AMB') ;"

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

                      Set CtlActive = CmbPerSocOpePerSocOpeTypCode

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               CmbPerSocOpePerSocOpeTypCode.RowSource = StrSQLPerSocOpePerSocOpeTypCode

               CmbPerSocOpePerSocOpeTypCode.Requery

               CmbPerSocOpePerSocNum.RowSource = StrSQLPerSocOpePerSocNum

               CmbPerSocOpePerSocNum.Requery

               SFrmSaisiesOperationsTypesOperationsResultats.Requery

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

   MLngPerSocOpeNum = Nz(PerSocOpeNum)

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
   Dim RsOperationsValeur As Recordset

   On Error GoTo Err_Form_Unload

   Select Case GIntMode
      Case MODEDEPLACEMENT

         Select Case Me.RecordsetClone.RecordCount
            Case 0

            Case Else

               Set RsOperationsValeur = SFrmSaisiesOperationsTypesOperationsResultats.Form.RecordsetClone

               RsOperationsValeur.FindFirst "PerSocOpeTypIDObligatoire=True AND IsNull([PerSocOpeResValeur])=True"

               Select Case RsOperationsValeur.NoMatch
                  Case True

                  Case False

                     Cancel = True

                     MsgBox RsOperationsValeur!PerSocOpeTypID & " - " & RsOperationsValeur!PerSocOpeTypIDLibelle & " n'a pas ete repondue !", , Me.Caption

                     SFrmSaisiesOperationsTypesOperationsResultats.Form.Bookmark = RsOperationsValeur.Bookmark

                     SFrmSaisiesOperationsTypesOperationsResultats.SetFocus

               End Select

               RsOperationsValeur.Close

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

   Select Case IsNull(PerSocOpePerSocOpeTypNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Operation Type Code Est Obligatoire !", , Me.Caption

         CmbPerSocOpePerSocOpeTypCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocOpePerSocNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Operation Personnel Numero Est Obligatoire !", , Me.Caption

         CmbPerSocOpePerSocNum.SetFocus

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
   Dim StrSQLOperations As String

   On Error GoTo Err_Form_AfterUpdate

   StrSQLOperations = "DELETE TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.* " & _
                                 " FROM TBLPERSONNELSSOCIETESOPERATIONSRESULTATS LEFT JOIN " & _
                                 " SelPersonnelsSocietesOperationsTypesID ON (TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.PerSocOpeResPerSocNum = " & _
                                 " SelPersonnelsSocietesOperationsTypesID.PerSocOpePerSocNum) AND (TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.PerSocOpeResPerSocOpeTypIDNum = " & _
                                 " SelPersonnelsSocietesOperationsTypesID.PerSocOpeTypIDNum) WHERE (((SelPersonnelsSocietesOperationsTypesID.PerSocOpePerSocNum) Is Null));"

   CurrentDb.Execute StrSQLOperations

   StrSQLOperations = "INSERT INTO TBLPERSONNELSSOCIETESOPERATIONSRESULTATS ( PerSocOpeResPerSocNum, " & _
                                 "  PerSocOpeResPerSocOpeTypIDNum ) SELECT SelPersonnelsSocietesOperationsTypesID.PerSocOpePerSocNum, " & _
                                 " SelPersonnelsSocietesOperationsTypesID.PerSocOpeTypIDNum FROM SelPersonnelsSocietesOperationsTypesID " & _
                                 " WHERE SelPersonnelsSocietesOperationsTypesID.PerSocOpePerSocNum=" & PerSocOpePerSocNum & _
                                 " AND SelPersonnelsSocietesOperationsTypesID.PerSocOpeTypIDPerSocOpeTypNum=" & PerSocOpePerSocOpeTypNum & ";"

   CurrentDb.Execute StrSQLOperations

   CmbPerSocOpePerSocOpeTypCode.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   CmbPerSocOpePerSocOpeTypCode.SetFocus

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
   Dim RsOperations As DAO.Recordset
   Dim RsOperationsValeur As DAO.Recordset
   Dim RsFormOperationsValeur As DAO.Recordset
   Dim StrOperationsValeur As String

   On Error GoTo Err_Form_Current

   Select Case GIntMode
      Case MODEDEPLACEMENT

         StrOperationsValeur = "SELECT PerSocOpeNum, PerSocOpeTypID, PerSocOpeTypIDObligatoire, PerSocOpeTypIDLibelle, PerSocOpeResValeur " & _
               " FROM TBLPERSONNELSSOCIETESOPERATIONSTYPES INNER JOIN (TBLPERSONNELSSOCIETESOPERATIONSTYPESID INNER JOIN (TBLPERSONNELSSOCIETESOPERATIONSRESULTATS INNER JOIN TBLPERSONNELSSOCIETESOPERATIONS ON TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.PerSocOpeResPerSocNum = TBLPERSONNELSSOCIETESOPERATIONS.PerSocOpePerSocNum) ON TBLPERSONNELSSOCIETESOPERATIONSTYPESID.PerSocOpeTypIDNum = TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.PerSocOpeResPerSocOpeTypIDNum) ON (TBLPERSONNELSSOCIETESOPERATIONSTYPES.PerSocOpeTypNum = TBLPERSONNELSSOCIETESOPERATIONS.PerSocOpePerSocOpeTypNum) AND (TBLPERSONNELSSOCIETESOPERATIONSTYPES.PerSocOpeTypNum = TBLPERSONNELSSOCIETESOPERATIONSTYPESID.PerSocOpeTypIDPerSocOpeTypNum) " & _
               " WHERE PerSocOpeNum=" & MLngPerSocOpeNum & ";"

         Set RsOperationsValeur = CurrentDb.OpenRecordset(StrOperationsValeur)

         RsOperationsValeur.FindFirst "PerSocOpeTypIDObligatoire=True AND IsNull([PerSocOpeResValeur])=True"

         Select Case RsOperationsValeur.NoMatch
            Case True

               MLngPerSocOpeNum = PerSocOpeNum

            Case False

               Select Case MLngPerSocOpeNum
                  Case PerSocOpeNum

                  Case Else

                     Set RsOperations = Me.RecordsetClone

                     RsOperations.FindFirst "PerSocOpeNum=" & MLngPerSocOpeNum

                     Me.Bookmark = RsOperations.Bookmark

                     MsgBox RsOperationsValeur!PerSocOpeTypID & " - " & RsOperationsValeur!PerSocOpeTypIDLibelle & " n'a pas ete repondue !", , Me.Caption

                     Set RsFormOperationsValeur = SFrmSaisiesOperationsTypesOperationsResultats.Form.RecordsetClone

                     RsFormOperationsValeur.FindFirst "PerSocOpeTypID=" & RsOperationsValeur!PerSocOpeTypID

                     SFrmSaisiesOperationsTypesOperationsResultats.Form.Bookmark = RsFormOperationsValeur.Bookmark

                     SFrmSaisiesOperationsTypesOperationsResultats.SetFocus

                     SFrmSaisiesOperationsTypesOperationsResultats.Form!CmbPerSocOpeResValeur.SetFocus

                     RsOperations.Close

                     RsFormOperationsValeur.Close

               End Select

         End Select

         RsOperationsValeur.Close

      Case Else

   End Select

   Rafraichir (RAFRAICHIRDEPLACEMENT)

Exit_Form_Current:

   Set RsOperations = Nothing

   Set RsOperationsValeur = Nothing

   Set RsFormOperationsValeur = Nothing

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
   Dim TxtDateSaisie As String
   Dim RsOperationsValeur As Recordset

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         CmbPerSocOpePerSocOpeTypCode.SetFocus

         CmbPerSocOpePerSocOpeTypCode.DefaultValue = "'" & CmbPerSocOpePerSocOpeTypCode.Value & "'"

         Set RsOperationsValeur = SFrmSaisiesOperationsTypesOperationsResultats.Form.RecordsetClone

         RsOperationsValeur.FindFirst "PerSocOpeTypIDObligatoire=True AND IsNull([PerSocOpeResValeur])=True"

         Select Case RsOperationsValeur.NoMatch
            Case True

            Case False

               MsgBox RsOperationsValeur!PerSocOpeTypID & " - " & RsOperationsValeur!PerSocOpeTypIDLibelle & " n'a pas ete repondue !", , Me.Caption

               SFrmSaisiesOperationsTypesOperationsResultats.Form.Bookmark = RsOperationsValeur.Bookmark

               SFrmSaisiesOperationsTypesOperationsResultats.SetFocus

               GoTo Exit_CmdAjouter_Click

         End Select

         RsOperationsValeur.Close

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateSaisie = TxtPerSocOpeDateSaisie

   TxtPerSocOpeDateSaisie = "01/01/2003"

   TxtPerSocOpeDateSaisie = TxtDateSaisie

   CmbPerSocOpePerSocOpeTypCode.SetFocus

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
   Dim TxtDateSaisie As String

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocOpePerSocOpeTypCode.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateSaisie = TxtPerSocOpeDateSaisie

         TxtPerSocOpeDateSaisie = "01/01/2003"

         TxtPerSocOpeDateSaisie = TxtDateSaisie

         CmbPerSocOpePerSocOpeTypCode.SetFocus

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
   Dim StrSQLOperations As String

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocOpePerSocOpeTypCode.SetFocus

         ChangeBoutonMode Me, MODESUPPRIME

         On Error Resume Next

         DoCmd.RunCommand acCmdDeleteRecord

         Select Case Err.Number
            Case vbEmpty

               StrSQLOperations = "DELETE TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.* " & _
                                                   " FROM TBLPERSONNELSSOCIETESOPERATIONSRESULTATS LEFT JOIN " & _
                                                   " SelPersonnelsSocietesOperationsTypesID ON (TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.PerSocOpeResPerSocNum = " & _
                                                   " SelPersonnelsSocietesOperationsTypesID.PerSocOpePerSocNum) AND (TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.PerSocOpeResPerSocOpeTypIDNum = " & _
                                                   " SelPersonnelsSocietesOperationsTypesID.PerSocOpeTypIDNum) WHERE (((SelPersonnelsSocietesOperationsTypesID.PerSocOpePerSocNum) Is Null));"

               CurrentDb.Execute StrSQLOperations

            Case Else

               Select Case Err.Number
                  Case 2501

                  Case Else

                     MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdSupprimer_Click"

               End Select

               Err.Clear

         End Select

         ChangeBoutonMode Me, MODEDEPLACEMENT

          Select Case Me.RecordsetClone.RecordCount
             Case vbEmpty

             Case Else

                CmbPerSocOpePerSocOpeTypCode.SetFocus

          End Select

         On Error GoTo Err_CmdSupprimer_Click

   End Select

   MLngPerSocOpeNum = Nz(PerSocOpeNum)

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

   CmbPerSocOpePerSocOpeTypCode.SetFocus

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

   CmbPerSocOpePerSocOpeTypCode.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocOpePerSocOpeTypCode.SetFocus

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
Private Sub CmdAfficherPersonnels_Click()
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherPersonnels_Click

   DonneEtatFormulaire "FrmFichePersonnels", IntEtat

   Select Case IntEtat
      Case VUEFORMULAIRE

         Forms!FrmFichePersonnels.SetFocus

         DoCmd.ApplyFilter , "PerNum=" & CmbPerSocOpePerSocNum.Column(1)

         DoCmd.RunCommand acCmdWindowCascade

      Case Else

         DoCmd.OpenForm "FrmFichePersonnels", acNormal, , "PerNum=" & CmbPerSocOpePerSocNum.Column(1), acFormEdit

   End Select

Exit_CmdAfficherPersonnels_Click:

   Exit Sub

Err_CmdAfficherPersonnels_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherPersonnels_Click"

   Resume Exit_CmdAfficherPersonnels_Click
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

         DoCmd.ApplyFilter , "SocNum=" & CmbPerSocOpePerSocNum.Column(9)

         DoCmd.RunCommand acCmdWindowCascade

      Case Else

         DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & CmbPerSocOpePerSocNum.Column(9), acFormEdit

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
Private Sub SFrmSaisiesOperationsTypesOperationsResultats_Enter()

   On Error GoTo Err_SFrmSaisiesOperationsTypesOperationsResultats_Enter

   SFrmSaisiesOperationsTypesOperationsResultats.Form.AllowEdits = True

Exit_SFrmSaisiesOperationsTypesOperationsResultats_Enter:

    Exit Sub

Err_SFrmSaisiesOperationsTypesOperationsResultats_Enter:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "SFrmSaisiesOperationsTypesOperationsResultats_Enter"

    Resume Exit_SFrmSaisiesOperationsTypesOperationsResultats_Enter
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub SFrmSaisiesOperationsTypesOperationsResultats_Exit(Cancel As Integer)

   On Error GoTo Err_SFrmSaisiesOperationsTypesOperationsResultats_Exit

   SFrmSaisiesOperationsTypesOperationsResultats.Form.AllowEdits = False

Exit_SFrmSaisiesOperationsTypesOperationsResultats_Exit:

    Exit Sub

Err_SFrmSaisiesOperationsTypesOperationsResultats_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "SFrmSaisiesOperationsTypesOperationsResultats_Exit"

    Resume Exit_SFrmSaisiesOperationsTypesOperationsResultats_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocOpePerSocOpeTypCode_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_CmbPerSocOpePerSocOpeTypCode_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_CmbPerSocOpePerSocOpeTypCode_KeyPress:

    Exit Sub

Err_CmbPerSocOpePerSocOpeTypCode_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocOpePerSocOpeTypCode_KeyPress"

    Resume Exit_CmbPerSocOpePerSocOpeTypCode_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocOpePerSocNum_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_CmbPerSocOpePerSocNum_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_CmbPerSocOpePerSocNum_KeyPress:

    Exit Sub

Err_CmbPerSocOpePerSocNum_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocOpePerSocNum_KeyPress"

    Resume Exit_CmbPerSocOpePerSocNum_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocOpePerSocOpeTypCode_GotFocus()

    On Error GoTo Err_CmbPerSocOpePerSocOpeTypCode_GotFocus

  ' CmbPerSocOpePerSocOpeTypCode.Dropdown

Exit_CmbPerSocOpePerSocOpeTypCode_GotFocus:

    Exit Sub

Err_CmbPerSocOpePerSocOpeTypCode_GotFocus:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocOpePerSocOpeTypCode_GotFocus"

    Resume Exit_CmbPerSocOpePerSocOpeTypCode_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocOpePerSocNum_GotFocus()

    On Error GoTo Err_CmbPerSocOpePerSocNum_GotFocus

    CmbPerSocOpePerSocNum.Dropdown

Exit_CmbPerSocOpePerSocNum_GotFocus:

    Exit Sub

Err_CmbPerSocOpePerSocNum_GotFocus:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocOpePerSocNum_GotFocus"

    Resume Exit_CmbPerSocOpePerSocNum_GotFocus
End Sub
