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
Private MLngPerSocOpeTypNum As Long
Private MLngPerSocOpeTypNbResultat As Long

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

                      Set CtlActive = TxtPerSocOpeTypCode

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         SFrmFicheOperationsTypesOperationsTypesID.Requery

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

   MLngPerSocOpeTypNum = Nz(PerSocOpeTypNum)

   MLngPerSocOpeTypNbResultat = Nz(PerSocOpeTypNbResultat)

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

                Select Case PerSocOpeTypNbResultat
                  Case Is = SFrmFicheOperationsTypesOperationsTypesID.Form.RecordsetClone.RecordCount

                  Case Else

                     Cancel = True

                     MsgBox "Attention Le Nombre De Résultat Definie Attendu Est Incorrect !!!", vbCritical, Me.Caption

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

   Select Case IsNull(PerSocOpeTypCode)
      Case False

         TxtPerSocOpeTypCode = UCase(TxtPerSocOpeTypCode)

      Case True

         Cancel = True

         MsgBox "Le Champ Operation Type Code Est Obligatoire !", , Me.Caption

         TxtPerSocOpeTypCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(PerSocOpeTypDateDebut)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Operation Type Date Debut Est Obligatoire !", , Me.Caption

         TxtPerSocOpeTypDateDebut.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(PerSocOpeTypDateFin)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Operation Type Date Fin Est Obligatoire !", , Me.Caption

         TxtPerSocOpeTypDateFin.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case PerSocOpeTypDateDebut <= PerSocOpeTypDateFin
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Operation Type Date Debut Doit Etre Inferieur A Personnels Sociétés Operations Types Date Fin !", , Me.Caption

         TxtPerSocOpeTypDateDebut.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocOpeTypLibelle)
      Case False

         TxtPerSocOpeTypLibelle = UCase(TxtPerSocOpeTypLibelle)

      Case True

         Cancel = True

         MsgBox "Le Champ Operation Type Libelle Est Obligatoire !", , Me.Caption

         TxtPerSocOpeTypLibelle.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocOpeTypNbResultat) Or PerSocOpeTypNbResultat <= 0
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Operation Type Nombre Résultats Est Invalide !", , Me.Caption

         TxtPerSocOpeTypNbResultat.SetFocus

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

   MLngPerSocOpeTypNum = PerSocOpeTypNum

   MLngPerSocOpeTypNbResultat = PerSocOpeTypNbResultat

   TxtPerSocOpeTypCode.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   TxtPerSocOpeTypCode.SetFocus

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
   Dim RstOperationsTypes As DAO.Recordset
   Dim RstOperationsTypesID As DAO.Recordset
   Dim StrOperationsTypesID As String

   On Error GoTo Exit_Form_Current

   Select Case GIntMode
      Case MODEDEPLACEMENT

         StrOperationsTypesID = "SELECT TBLPERSONNELSSOCIETESOPERATIONSTYPESID.* " & _
            "FROM TBLPERSONNELSSOCIETESOPERATIONSTYPESID WHERE " & "PerSocOpeTypIDPerSocOpeTypNum=" & MLngPerSocOpeTypNum & ";"

         Set RstOperationsTypesID = CurrentDb.OpenRecordset(StrOperationsTypesID)

         Select Case RstOperationsTypesID.EOF
            Case True

            Case False

               RstOperationsTypesID.MoveLast

         End Select

         Select Case MLngPerSocOpeTypNbResultat
            Case Is = RstOperationsTypesID.RecordCount

               MLngPerSocOpeTypNum = PerSocOpeTypNum

               MLngPerSocOpeTypNbResultat = PerSocOpeTypNbResultat

               Rafraichir (RAFRAICHIRDEPLACEMENT)

            Case Else

               Select Case PerSocOpeTypNum
                  Case MLngPerSocOpeTypNum

                  Case Else

                     Set RstOperationsTypes = Me.RecordsetClone

                     RstOperationsTypes.FindFirst "PerSocOpeTypNum=" & MLngPerSocOpeTypNum

                     Me.Bookmark = RstOperationsTypes.Bookmark

                     MsgBox "Attention Le Nombre De Résultat Definie Attendu Est Incorrect !!!", , Me.Caption

                     RstOperationsTypes.Close

               End Select

         End Select

         RstOperationsTypesID.Close

      Case Else

   End Select

Exit_Form_Current:

   Set RstOperationsTypes = Nothing

   Set RstOperationsTypesID = Nothing

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
   Dim TxtOpeTypDateDebut As String

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         TxtPerSocOpeTypCode.SetFocus

         Select Case PerSocOpeTypNbResultat
           Case Is = SFrmFicheOperationsTypesOperationsTypesID.Form.RecordsetClone.RecordCount

           Case Else

              MsgBox "Attention Le Nombre De Résultat Definie Attendu Est Incorrect !!!", vbCritical, Me.Caption

              GoTo Exit_CmdAjouter_Click

         End Select

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtOpeTypDateDebut = TxtPerSocOpeTypDateDebut

   TxtPerSocOpeTypDateDebut = "01/01/2003"

   TxtPerSocOpeTypDateDebut = TxtOpeTypDateDebut

   TxtPerSocOpeTypCode.SetFocus

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
   Dim TxtOpeTypDateDebut As String

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerSocOpeTypCode.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtOpeTypDateDebut = TxtPerSocOpeTypDateDebut

         TxtPerSocOpeTypDateDebut = "01/01/2003"

         TxtPerSocOpeTypDateDebut = TxtOpeTypDateDebut

         TxtPerSocOpeTypCode.SetFocus

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

         TxtPerSocOpeTypCode.SetFocus

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

                TxtPerSocOpeTypCode.SetFocus

          End Select

   End Select

   MLngPerSocOpeTypNum = Nz(PerSocOpeTypNum)

   MLngPerSocOpeTypNbResultat = Nz(PerSocOpeTypNbResultat)

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

   TxtPerSocOpeTypCode.SetFocus

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

   TxtPerSocOpeTypCode.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerSocOpeTypCode.SetFocus

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
Private Sub TxtPerSocOpeTypNbResultat_Exit(Cancel As Integer)

    On Error GoTo Err_TxtPerSocOpeTypNbResultat_Exit

    Select Case PerSocOpeTypNbResultat
      Case Is >= SFrmFicheOperationsTypesOperationsTypesID.Form.RecordsetClone.RecordCount

      Case Else

         Cancel = True

         MsgBox "Attention Le Nombre De Resultat Définie " & PerSocOpeTypNbResultat & " Est Inferieur Au Nombre D'Identifiant Définie " & SFrmFicheOperationsTypesOperationsTypesID.Form.RecordsetClone.RecordCount & "!!!", vbCritical, Me.Caption

   End Select

Exit_TxtPerSocOpeTypNbResultat_Exit:

    Exit Sub

Err_TxtPerSocOpeTypNbResultat_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocOpeTypNbResultat_Exit"

    Resume Exit_TxtPerSocOpeTypNbResultat_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocOpeTypCode_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocOpeTypCode_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocOpeTypCode_KeyPress:

    Exit Sub

Err_TxtPerSocOpeTypCode_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocOpeTypCode_KeyPress"

    Resume Exit_TxtPerSocOpeTypCode_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocOpeTypLibelle_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocOpeTypLibelle_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocOpeTypLibelle_KeyPress:

    Exit Sub

Err_TxtPerSocOpeTypLibelle_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocOpeTypLibelle_KeyPress"

    Resume Exit_TxtPerSocOpeTypLibelle_KeyPress
End Sub
