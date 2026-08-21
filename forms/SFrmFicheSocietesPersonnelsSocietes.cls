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
Public GBlnModif As Boolean

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
   Dim StrSQLPerSocPerNum As String
   Dim StrSQLPerSocPerSocTypCode As String
   Dim CtlActive As Control

   On Error GoTo Err_Rafraichir

   StrSQLPerSocPerNum = "SELECT TBLPERSONNELS.PerNum, TBLPERSONNELS.PerCivCode, TBLPERSONNELS.PerNom, TBLPERSONNELS.PerPrenom FROM TBLPERSONNELS ORDER BY TBLPERSONNELS.PerNum;"

   StrSQLPerSocPerSocTypCode = "SELECT SelPersonnelsSocietesTypes.* FROM SelPersonnelsSocietesTypes ORDER BY SelPersonnelsSocietesTypes.PerSocTypRang; "

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

                      Set CtlActive = CmbPerSocPerSocTypCode

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX, RAFRAICHIRDEPLACEMENT

               CmbPerSocPerNum.RowSource = StrSQLPerSocPerNum

               CmbPerSocPerSocTypCode.RowSource = StrSQLPerSocPerSocTypCode

               CmbPerSocPerNum.Requery

               CmbPerSocPerSocTypCode.Requery

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

   On Error Resume Next

   Select Case IsObject(Me.Parent)
      Case True

      Case False

         Cancel = True

   End Select

   Select Case Err.Number
      Case vbEmpty

      Case Else

         Err.Clear

         Cancel = True

   End Select

   On Error GoTo Err_Form_Open

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

   ChangeBoutonMode Me, MODEDEPLACEMENT

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

   Select Case MsgBox("Voulez-Vous Mettre à Jour Le Profil ?", vbYesNo, Me.Caption)
      Case vbYes

      Case vbNo

         Cancel = True

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(PerSocDateCre)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Personnels Sociétés Date De Creation Est Invalide !", , Me.Caption

         GoTo Exit_Form_BeforeUpdate

   End Select

   PerSocDateMaj = Date

   Select Case IsNull(PerSocSocNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Personnels Sociétés Sociétés Numero Est Obligatoire !", , Me.Caption

         CmbPerSocPerNum.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocPerSocTypCode)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Personnels Sociétés Personnels Sociétés Types Code Est Obligatoire !", , Me.Caption

         CmbPerSocPerSocTypCode.SetFocus

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
   Dim BlnAdhEtat As Variant
   Dim BlnActEtat As Variant

   On Error GoTo Err_Form_AfterUpdate

   CmbPerSocPerNum.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   CmbPerSocPerNum.SetFocus

   Me.Parent.ChangeSFrmButtonMode Me

   BlnAdhEtat = DLookup("PerSocAdhEtat", "TBLPERSONNELSSOCIETESADHERENTS", "PerSocAdhPerSocNum=" & PerSocNum)

   Select Case IsNull(BlnAdhEtat)
      Case True

         CurrentDb.Execute "INSERT INTO TBLPERSONNELSSOCIETESADHERENTS (PerSocAdhPerSocNum,PerSocAdhDate,PerSocAdhEtat) VALUES (" & PerSocNum & ",Date() ," & PerSocAdherent & ") ;"

      Case False

         Select Case CBool(PerSocAdherent)
            Case Is <> CBool(BlnAdhEtat)

               CurrentDb.Execute "INSERT INTO TBLPERSONNELSSOCIETESADHERENTS (PerSocAdhPerSocNum,PerSocAdhDate,PerSocAdhEtat) VALUES (" & PerSocNum & ",Date()," & PerSocAdherent & ") ;"

            Case Else

         End Select

   End Select

   BlnActEtat = DLookup("PerSocActEtat", "TBLPERSONNELSSOCIETESACTIFS", "PerSocActPerSocNum=" & PerSocNum)

   Select Case IsNull(BlnActEtat)
      Case True

         CurrentDb.Execute "INSERT INTO TBLPERSONNELSSOCIETESACTIFS (PerSocActPerSocNum,PerSocActDate,PerSocActEtat) VALUES (" & PerSocNum & ",Date()," & PerSocActif & ") ;"

      Case False

         Select Case CBool(PerSocActif)
            Case Is <> CBool(BlnActEtat)

               CurrentDb.Execute "INSERT INTO TBLPERSONNELSSOCIETESACTIFS (PerSocActPerSocNum,PerSocActDate,PerSocActEtat) VALUES (" & PerSocNum & ",Date()," & PerSocActif & ") ;"

            Case Else

         End Select

   End Select

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
Public Sub CmdAjouter_Click()
   Dim TxtDateCreer As String

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         CmbPerSocPerNum.SetFocus

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateCreer = TxtPerSocActDate

   TxtPerSocActDate = "01/01/2003"

   TxtPerSocActDate = TxtDateCreer

   CmbPerSocPerNum.SetFocus

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
Public Sub CmdModifier_Click()
   Dim TxtDateCreer As String

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocPerNum.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateCreer = TxtPerSocActDate

         TxtPerSocActDate = "01/01/2003"

         TxtPerSocActDate = TxtDateCreer

         CmbPerSocPerNum.SetFocus

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
Public Sub CmdSupprimer_Click()

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocPerNum.SetFocus

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

         ChangeBoutonMode Me, MODEDEPLACEMENT

          Select Case Me.RecordsetClone.RecordCount
             Case vbEmpty

             Case Else

                CmbPerSocPerNum.SetFocus

          End Select

         On Error GoTo Err_CmdSupprimer_Click
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
Public Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   CmbPerSocPerNum.SetFocus

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
Public Sub CmdAnnuler_Click()

   On Error GoTo Err_CmdAnnuler_Click

   CmbPerSocPerNum.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocPerNum.SetFocus

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
Private Sub CmdAfficheTous_Click()

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

Exit_CmdAfficheTous_Click:

   Exit Sub

Err_CmdAfficheTous_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficheTous_Click"

   Resume Exit_CmdAfficheTous_Click
End Sub
