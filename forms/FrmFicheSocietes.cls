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
'***     Declaration De Variable Public                                                         ***
'******************************************************************************

'******************************************************************************
'***     Variable De Controle De Mode Saisie                                                ***
'******************************************************************************
Public GIntMode As Integer
Public GIntModeOld As Integer

'******************************************************************************
'***     Declaration De Variable Priver                                                         ***
'******************************************************************************

'******************************************************************************
'***     Variable Pour Le Rafraichissement                                                   ***
'******************************************************************************
Private MTxtSocVilNum As String

'******************************************************************************
'***    Declaration De Procedure Priver                                                        ***
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
   Dim StrSQLSocVilNum As String
   Dim CtlActive As Control

   On Error GoTo Err_Rafraichir

   StrSQLSocVilNum = "SELECT VilNum, VilCodePostal, VilNom, DepNom, DepPays FROM SelVillesDepartementsCodes LEFT JOIN TBLDEPARTEMENTS ON SelVillesDepartementsCodes.VilDepCode=TBLDEPARTEMENTS.DepCode ORDER BY VilCodePostal;"

   Select Case (Me.RecordsetClone.RecordCount = 0) And (Me.NewRecord = False)
      Case True

      Case False

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               On Error Resume Next

               Set CtlActive = Me.ActiveControl

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                      Err.Clear

                      Set CtlActive = TxtSocCode

               End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               CmbSocVilNum.RowSource = StrSQLSocVilNum

               Select Case MTxtSocVilNum
                  Case vbNullString

                     CmbSocVilNum.Requery

                  Case Else

                     CmbSocVilNum.Requery

                     Select Case CmbSocVilNum.Text
                        Case vbNullString

                           CmbSocVilNum.Text = MTxtSocVilNum

                        Case Else

                     End Select

                     MTxtSocVilNum = vbNullString

               End Select

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
   Dim BlnAdhEtat As Variant

   On Error GoTo Err_Form_BeforeUpdate

   Select Case GIntMode
      Case MODEVALIDE

      Case Else

         Cancel = True

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(SocDateCre)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Sociétés Date De Creation Est Invalide !", , Me.Caption

         TxtSocDateCre.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   SocDateMaj = Date

   Select Case IsNull(SocRaison)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Sociétés Raison Sociale Est Obligatoire !", , Me.Caption

         TxtSocDateCre.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(SocAdr)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Sociétés Adresse Est Obligatoire !", , Me.Caption

         TxtSocAdresse.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(SocVilNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Sociétés Ville Est Obligatoire", , Me.Caption

         CmbSocVilNum.SetFocus

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

   TxtSocCode.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   TxtSocCode.SetFocus

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
   Dim RsSocPerRVRDVR As DAO.Recordset
   Dim StrSQLSocPerRVRDVR As String

   On Error GoTo Err_Form_Current

   StrSQLSocPerRVRDVR = "SELECT TBLPERSONNELSSOCIETESCODES.PerSocCodCode AS RVRCode, TBLPERSONNELS.PerNom AS RVRNom, TBLPERSONNELS.PerPrenom AS RVRPrenom, TBLSOCIETESSECTEURS.SocSecLibelle, TBLPERSONNELSSOCIETESCODES_1.PerSocCodCode AS DVRCode, TBLPERSONNELS_1.PerNom AS DVRNom, TBLPERSONNELS_1.PerPrenom AS DVRPrenom, TBLSOCIETESSECTEURSREGIONS.SocSecRegLibelle " & _
         " FROM ((TBLSOCIETESSECTEURSREGIONS INNER JOIN (((TBLPERSONNELS RIGHT JOIN TBLPERSONNELSSOCIETESCODES ON TBLPERSONNELS.PerNum = TBLPERSONNELSSOCIETESCODES.PerSocCodPerNum) RIGHT JOIN TBLSOCIETESSECTEURS ON TBLPERSONNELSSOCIETESCODES.PerSocCodSocSecNum = TBLSOCIETESSECTEURS.SocSecNum) INNER JOIN TBLSOCIETESSECTEURSCODES ON TBLSOCIETESSECTEURS.SocSecNum = TBLSOCIETESSECTEURSCODES.SocSecCodSocSecNum) ON TBLSOCIETESSECTEURSREGIONS.SocSecRegNum = TBLSOCIETESSECTEURS.SocSecSocSecRegNum) LEFT JOIN TBLPERSONNELSSOCIETESCODES AS TBLPERSONNELSSOCIETESCODES_1 ON TBLSOCIETESSECTEURSREGIONS.SocSecRegNum = TBLPERSONNELSSOCIETESCODES_1.PerSocCodSocSecRegNum) LEFT JOIN TBLPERSONNELS AS TBLPERSONNELS_1 ON TBLPERSONNELSSOCIETESCODES_1.PerSocCodPerNum = TBLPERSONNELS_1.PerNum " & _
         " WHERE (((TBLSOCIETESSECTEURSCODES.SocSecCodCode)=Left('" & CmbSocVilNum.Column(1) & "',Len([SocSecCodCode]))));"

   Set RsSocPerRVRDVR = CurrentDb.OpenRecordset(StrSQLSocPerRVRDVR)

   Select Case RsSocPerRVRDVR.EOF
      Case True

         TxtSocPerRVR = vbNullString

      Case False

         TxtSocPerRVR = "RVR : " & RsSocPerRVRDVR!RVRCode & " - " & RsSocPerRVRDVR!RVRNom & ", " & RsSocPerRVRDVR!RVRPrenom & " (" & RsSocPerRVRDVR!SocSecLibelle & ") | DVR : " & RsSocPerRVRDVR!DVRCode & " - " & RsSocPerRVRDVR!DVRNom & ", " & RsSocPerRVRDVR!DVRPrenom & " (" & RsSocPerRVRDVR!SocSecRegLibelle & ")"

   End Select

   RsSocPerRVRDVR.Close

   EtiPreChanel.Visible = (Nz(SocSocResNum, 0) = 1)

   EtiPrePrecision.Visible = (Nz(SocSocResNum, 0) = 1)

   ImgPreChanel.Visible = (Nz(SocSocResNum, 0) = 1)

   Rafraichir (RAFRAICHIRDEPLACEMENT)

Exit_Form_Current:

   Set RsSocPerRVRDVR = Nothing

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

         TxtSocCode.SetFocus

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateCreer = TxtSocDateCre

   TxtSocDateCre = "01/01/2003"

   TxtSocDateCre = TxtDateCreer

   TxtSocCode.SetFocus

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

         TxtSocCode.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateCreer = TxtSocDateCre

         TxtSocDateCre = "01/01/2003"

         TxtSocDateCre = TxtDateCreer

         TxtSocCode.SetFocus

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

         TxtSocCode.SetFocus

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

               TxtSocCode.SetFocus

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

   TxtSocCode.SetFocus

   ChangeBoutonMode Me, MODEVALIDE

   DoCmd.RunCommand acCmdSaveRecord

   On Error Resume Next

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Select Case Err.Number
            Case 2501

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

   TxtSocCode.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtSocCode.SetFocus

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

   DoCmd.Close acForm, Me.Name, acSavePrompt

   On Error Resume Next

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
Private Sub CmdRechercher_Click()

   On Error GoTo Exit_CmdRechercher_Click

   OuvreRecherches 1

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
Private Sub CmdAfficherContacts_Click()
'   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherContacts_Click

'   DonneEtatForm "FrmFicheContacts", IntEtat

'   Select Case IntEtat
'      Case ETATFORMULAIRE

'         Forms!FrmFicheContacts.SetFocus

'         DoCmd.ApplyFilter , "ConSocNum=" & SocNum

'         DoCmd.RunCommand acCmdWindowCascade

'      Case Else

'         DoCmd.OpenForm "FrmFicheContacts", acNormal, , "ConSocNum=" & SocNum, acFormEdit

'   End Select

'   Forms!FrmFicheContacts.LstConSocNum.DefaultValue = SocNum

'   Forms!FrmFicheContacts.LstConSocNum.DefaultValue = IIf(IsNull(SocNum), vbNullString, SocNum)

'   Forms!FrmFicheContacts.TxtConAdresse.DefaultValue = IIf(IsNull(SocAdresse.Value), vbNullString, """" & SocAdresse & """")

'   Forms!FrmFicheContacts.LstConVilNum.DefaultValue = IIf(IsNull(SocVilNum), vbNullString, SocVilNum)

'   Forms!FrmFicheContacts.TxtConFax.DefaultValue = IIf(IsNull(SocFax), vbNullString, """" & SocFax & """")

Exit_CmdAfficherContacts_Click:

   Exit Sub

Err_CmdAfficherContacts_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherContacts_Click"

   Resume Exit_CmdAfficherContacts_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdNouvellesVilles_Click()

   On Error GoTo Err_CmdNouvellesVilles_Click

   AjouteVilles

Exit_CmdNouvellesVilles_Click:

   Exit Sub

Err_CmdNouvellesVilles_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdNouvellesVilles_Click"

   Resume Exit_CmdNouvellesVilles_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbSocVilNum_NotInList(NewData As String, Response As Integer)

   On Error GoTo Err_CmbSocVilNum_NotInList

   AjouteVilles (NewData)

   Response = acDataErrContinue

   MTxtSocVilNum = CmbSocVilNum.Text

   CmbSocVilNum.Undo

Exit_CmbSocVilNum_NotInList:

   Exit Sub

Err_CmbSocVilNum_NotInList:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocVilNum_NotInList"

   Resume Exit_CmbSocVilNum_NotInList
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbSocSocTypNum_AfterUpdate()

   On Error GoTo Err_CmbSocSocTypNum_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

Exit_CmbSocSocTypNum_AfterUpdate:

   Exit Sub

Err_CmbSocSocTypNum_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocSocTypNum_AfterUpdate"

   Resume Exit_CmbSocSocTypNum_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtSocRaison_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtSocRaison_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtSocRaison_KeyPress:

    Exit Sub

Err_TxtSocRaison_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocRaison_KeyPress"

    Resume Exit_TxtSocRaison_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbSocEnseigne_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_CmbSocEnseigne_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_CmbSocEnseigne_KeyPress:

    Exit Sub

Err_CmbSocEnseigne_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocEnseigne_KeyPress"

    Resume Exit_CmbSocEnseigne_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbSocTel1Libelle_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_CmbSocTel1Libelle_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_CmbSocTel1Libelle_KeyPress:

    Exit Sub

Err_CmbSocTel1Libelle_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocTel1Libelle_KeyPress"

    Resume Exit_CmbSocTel1Libelle_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbSocTel2Libelle_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_CmbSocTel2Libelle_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_CmbSocTel2Libelle_KeyPress:

    Exit Sub

Err_CmbSocTel2Libelle_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbSocTel2Libelle_KeyPress"

    Resume Exit_CmbSocTel2Libelle_KeyPress
End Sub
