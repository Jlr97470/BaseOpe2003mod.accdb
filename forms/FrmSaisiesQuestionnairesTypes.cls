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
'***    Variable Du Numero De Questionnaires En Cours                               ***
'******************************************************************************
Private MLngPerSocQstNum As Long

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
   Dim StrSQLPerSocQstPerSocNum As String
   Dim StrSQLPerSocQstPerSocQstTypCode As String
   Dim CtlActive As Control

   On Error GoTo Err_Rafraichir

   StrSQLPerSocQstPerSocQstTypCode = "SELECT PerSocQstTypNum, PerSocQstTypCode, PerSocQstTypLibelle, PerSocQstTypDateDebut, PerSocQstTypDateFin, PerSocQstTypNbResultat " & _
      " FROM SelPersonnelsSocietesQuestionnairesTypes ;"

   StrSQLPerSocQstPerSocNum = "SELECT PerSocNum , PerNum, PerCivCode, PerNom, PerPrenom, PerSocPerSocTypCode, PerSocTypLibelle, PerSocActif, " & _
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

                      Set CtlActive = CmbPerSocQstPerSocQstTypCode

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               CmbPerSocQstPerSocQstTypCode.RowSource = StrSQLPerSocQstPerSocQstTypCode

               CmbPerSocQstPerSocQstTypCode.Requery

               CmbPerSocQstPerSocNum.RowSource = StrSQLPerSocQstPerSocNum

               CmbPerSocQstPerSocNum.Requery

               SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Requery

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

   MLngPerSocQstNum = Nz(PerSocQstNum)

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
   Dim RsQuestionnairesValeur As Recordset

   On Error GoTo Err_Form_Unload

   Select Case GIntMode
      Case MODEDEPLACEMENT

         Select Case Me.RecordsetClone.RecordCount
            Case 0

            Case Else

               Set RsQuestionnairesValeur = SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form.RecordsetClone

               RsQuestionnairesValeur.FindFirst "PerSocQstTypIDObligatoire=True AND IsNull([PerSocQstResValeur])=True"

               Select Case RsQuestionnairesValeur.NoMatch
                  Case True

                  Case False

                     Cancel = True

                     MsgBox RsQuestionnairesValeur!PerSocQstTypID & " - " & RsQuestionnairesValeur!PerSocQstTypIDLibelle & " n'a pas ete repondue !", , Me.Caption

                     SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form.Bookmark = RsQuestionnairesValeur.Bookmark

                     SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.SetFocus

                     SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form!CmbPerSocQstResValeur.SetFocus

               End Select

               RsQuestionnairesValeur.Close

         End Select

      Case Else

         Cancel = True

         MsgBox "Vous Devez Validez Ou Annulez La Modification En Cours !", , Me.Caption

   End Select

Exit_Form_Unload:

   Set RsQuestionnairesValeur = Nothing

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

   Select Case IsNull(PerSocQstPerSocQstTypNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaire Type Code Est Obligatoire !", , Me.Caption

         CmbPerSocQstPerSocQstTypCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocQstPerSocNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaire Personnel Numero Est Obligatoire !", , Me.Caption

         CmbPerSocQstPerSocNum.SetFocus

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
   Dim StrSQLQuestionnaires As String

   On Error GoTo Err_Form_AfterUpdate

   MLngPerSocQstNum = PerSocQstNum

   StrSQLQuestionnaires = "DELETE TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.* " & _
                                       " FROM TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS LEFT JOIN " & _
                                       " SelPersonnelsSocietesQuestionnairesTypesID ON (TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.PerSocQstResPerSocNum = " & _
                                       " SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstPerSocNum) AND (TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.PerSocQstResPerSocQstTypIDNum = " & _
                                       " SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstTypIDNum) WHERE (((SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstPerSocNum) Is Null));"

   CurrentDb.Execute StrSQLQuestionnaires

   StrSQLQuestionnaires = "INSERT INTO TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS ( PerSocQstResPerSocNum, " & _
                                        "  PerSocQstResPerSocQstTypIDNum ) SELECT SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstPerSocNum, " & _
                                        " SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstTypIDNum FROM SelPersonnelsSocietesQuestionnairesTypesID " & _
                                        " WHERE SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstPerSocNum=" & PerSocQstPerSocNum & _
                                        " AND SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstTypIDPerSocQstTypNum=" & PerSocQstPerSocQstTypNum & ";"

   CurrentDb.Execute StrSQLQuestionnaires

   CmbPerSocQstPerSocQstTypCode.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   CmbPerSocQstPerSocQstTypCode.SetFocus

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
   Dim RsQuestionnaires As DAO.Recordset
   Dim RsQuestionnairesValeur As DAO.Recordset
   Dim RsFormQuestionnairesValeur As DAO.Recordset
   Dim StrQuestionnairesValeur As String

   On Error GoTo Err_Form_Current

   Select Case GIntMode
      Case MODEDEPLACEMENT

         StrQuestionnairesValeur = "SELECT PerSocQstNum, PerSocQstTypID, PerSocQstTypIDObligatoire, PerSocQstTypIDLibelle, PerSocQstResValeur " & _
               " FROM TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPES INNER JOIN (TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPESID INNER JOIN (TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS INNER JOIN TBLPERSONNELSSOCIETESQUESTIONNAIRES ON TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.PerSocQstResPerSocNum = TBLPERSONNELSSOCIETESQUESTIONNAIRES.PerSocQstPerSocNum) ON TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPESID.PerSocQstTypIDNum = TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.PerSocQstResPerSocQstTypIDNum) ON (TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPES.PerSocQstTypNum = TBLPERSONNELSSOCIETESQUESTIONNAIRES.PerSocQstPerSocQstTypNum) AND (TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPES.PerSocQstTypNum = TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPESID.PerSocQstTypIDPerSocQstTypNum) " & _
               " WHERE PerSocQstNum=" & MLngPerSocQstNum & ";"

         Set RsQuestionnairesValeur = CurrentDb.OpenRecordset(StrQuestionnairesValeur)

         RsQuestionnairesValeur.FindFirst "PerSocQstTypIDObligatoire=True AND IsNull([PerSocQstResValeur])=True"

         Select Case RsQuestionnairesValeur.NoMatch
            Case True

               MLngPerSocQstNum = PerSocQstNum

               Rafraichir (RAFRAICHIRDEPLACEMENT)

            Case False

               Select Case MLngPerSocQstNum
                  Case PerSocQstNum

                  Case Else

                     Set RsQuestionnaires = Me.RecordsetClone

                     RsQuestionnaires.FindFirst "PerSocQstNum=" & MLngPerSocQstNum

                     Me.Bookmark = RsQuestionnaires.Bookmark

                     MsgBox RsQuestionnairesValeur!PerSocQstTypID & " - " & RsQuestionnairesValeur!PerSocQstTypIDLibelle & " n'a pas ete repondue !", , Me.Caption

                     Set RsFormQuestionnairesValeur = SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form.RecordsetClone

                     RsFormQuestionnairesValeur.FindFirst "PerSocQstTypID=" & RsQuestionnairesValeur!PerSocQstTypID

                     SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form.Bookmark = RsFormQuestionnairesValeur.Bookmark

                     SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.SetFocus

                     SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form!CmbPerSocQstResValeur.SetFocus

                     RsQuestionnaires.Close

                     RsFormQuestionnairesValeur.Close

               End Select

         End Select

         RsQuestionnairesValeur.Close

      Case Else

   End Select

Exit_Form_Current:

   Set RsQuestionnaires = Nothing

   Set RsQuestionnairesValeur = Nothing

   Set RsFormQuestionnairesValeur = Nothing

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
   Dim RsQuestionnairesValeur As Recordset

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         CmbPerSocQstPerSocQstTypCode.SetFocus

         CmbPerSocQstPerSocQstTypCode.DefaultValue = "'" & CmbPerSocQstPerSocQstTypCode.Value & "'"

         Set RsQuestionnairesValeur = SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form.RecordsetClone

         RsQuestionnairesValeur.FindFirst "PerSocQstTypIDObligatoire=True AND IsNull([PerSocQstResValeur])=True"

         Select Case RsQuestionnairesValeur.NoMatch
            Case True

            Case False

               MsgBox RsQuestionnairesValeur!PerSocQstTypID & " - " & RsQuestionnairesValeur!PerSocQstTypIDLibelle & " n'a pas ete repondue !", , Me.Caption

               SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form.Bookmark = RsQuestionnairesValeur.Bookmark

               SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.SetFocus

               GoTo Exit_CmdAjouter_Click

         End Select

         RsQuestionnairesValeur.Close

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateSaisie = TxtPerSocQstDateSaisie

   TxtPerSocQstDateSaisie = "01/01/2003"

   TxtPerSocQstDateSaisie = TxtDateSaisie

   CmbPerSocQstPerSocQstTypCode.SetFocus

Exit_CmdAjouter_Click:

   Set RsQuestionnairesValeur = Nothing

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

         CmbPerSocQstPerSocQstTypCode.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateSaisie = TxtPerSocQstDateSaisie

         TxtPerSocQstDateSaisie = "01/01/2003"

         TxtPerSocQstDateSaisie = TxtDateSaisie

         CmbPerSocQstPerSocQstTypCode.SetFocus

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
   Dim StrSQLQuestionnaires As String

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocQstPerSocQstTypCode.SetFocus

         ChangeBoutonMode Me, MODESUPPRIME

         On Error Resume Next

         DoCmd.RunCommand acCmdDeleteRecord

         Select Case Err.Number
            Case vbEmpty

               StrSQLQuestionnaires = "DELETE TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.* " & _
                                                   " FROM TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS LEFT JOIN " & _
                                                   " SelPersonnelsSocietesQuestionnairesTypesID ON (TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.PerSocQstResPerSocNum = " & _
                                                   " SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstPerSocNum) AND (TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.PerSocQstResPerSocQstTypIDNum = " & _
                                                   " SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstTypIDNum) WHERE (((SelPersonnelsSocietesQuestionnairesTypesID.PerSocQstPerSocNum) Is Null));"

               CurrentDb.Execute StrSQLQuestionnaires

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

                CmbPerSocQstPerSocQstTypCode.SetFocus

          End Select

         On Error GoTo Err_CmdSupprimer_Click

   End Select

   MLngPerSocQstNum = Nz(PerSocQstNum)

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

   CmbPerSocQstPerSocQstTypCode.SetFocus

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

   CmbPerSocQstPerSocQstTypCode.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocQstPerSocQstTypCode.SetFocus

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

         DoCmd.ApplyFilter , "PerNum=" & CmbPerSocQstPerSocNum.Column(1)

         DoCmd.RunCommand acCmdWindowCascade

      Case Else

         DoCmd.OpenForm "FrmFichePersonnels", acNormal, , "PerNum=" & CmbPerSocQstPerSocNum.Column(1), acFormEdit

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

         DoCmd.ApplyFilter , "SocNum=" & CmbPerSocQstPerSocNum.Column(9)

         DoCmd.RunCommand acCmdWindowCascade

      Case Else

         DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & CmbPerSocQstPerSocNum.Column(9), acFormEdit

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
Private Sub SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Enter()

   On Error GoTo Err_SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Enter

   SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form.AllowEdits = True

Exit_SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Enter:

    Exit Sub

Err_SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Enter:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Enter"

    Resume Exit_SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Enter
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Exit(Cancel As Integer)

   On Error GoTo Err_SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Exit

   SFrmSaisiesQuestionnairesTypesQuestionnairesResultats.Form.AllowEdits = False

Exit_SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Exit:

    Exit Sub

Err_SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Exit"

    Resume Exit_SFrmSaisiesQuestionnairesTypesQuestionnairesResultats_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocQstPerSocQstTypCode_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_CmbPerSocQstPerSocQstTypCode_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_CmbPerSocQstPerSocQstTypCode_KeyPress:

    Exit Sub

Err_CmbPerSocQstPerSocQstTypCode_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocQstPerSocQstTypCode_KeyPress"

    Resume Exit_CmbPerSocQstPerSocQstTypCode_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocQstPerSocNum_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_CmbPerSocQstPerSocNum_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_CmbPerSocQstPerSocNum_KeyPress:

    Exit Sub

Err_CmbPerSocQstPerSocNum_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocQstPerSocNum_KeyPress"

    Resume Exit_CmbPerSocQstPerSocNum_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocQstPerSocQstTypCode_GotFocus()

    On Error GoTo Err_CmbPerSocQstPerSocQstTypCode_GotFocus

  ' CmbPerSocQstPerSocQstTypCode.Dropdown

Exit_CmbPerSocQstPerSocQstTypCode_GotFocus:

    Exit Sub

Err_CmbPerSocQstPerSocQstTypCode_GotFocus:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocQstPerSocQstTypCode_GotFocus"

    Resume Exit_CmbPerSocQstPerSocQstTypCode_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocQstPerSocNum_GotFocus()

    On Error GoTo Err_CmbPerSocQstPerSocNum_GotFocus

    CmbPerSocQstPerSocNum.Dropdown

Exit_CmbPerSocQstPerSocNum_GotFocus:

    Exit Sub

Err_CmbPerSocQstPerSocNum_GotFocus:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocQstPerSocNum_GotFocus"

    Resume Exit_CmbPerSocQstPerSocNum_GotFocus
End Sub
