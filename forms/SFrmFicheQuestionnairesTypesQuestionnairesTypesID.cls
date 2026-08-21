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
   Dim CtlActive As Control

   On Error GoTo Err_Rafraichir

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

                      Set CtlActive = TxtPerSocQstTypID

              End Select

            Case Else

         End Select

         On Error GoTo Err_Rafraichir

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               CmbPerSocQstTypIDTypNum.Requery

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRDEPLACEMENT

               Select Case IsNull(PerSocQstTypIDTypNum)
                  Case False

                     Select Case PerSocQstTypIDTypNum
                        Case vbBoolean

                           TxtPerSocQstTypIDMin.Enabled = False

                           TxtPerSocQstTypIDMax.Enabled = False

                        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

                           TxtPerSocQstTypIDMin.Enabled = True

                           TxtPerSocQstTypIDMax.Enabled = True

                     End Select

                  Case True

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
   Dim VarValeur As Variant

   On Error GoTo Err_Form_BeforeUpdate

   Select Case GIntMode
      Case MODEVALIDE

      Case Else

         Cancel = True

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocQstTypID)
      Case False

         TxtPerSocQstTypID = UCase(TxtPerSocQstTypID)

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaires Types ID Est Obligatoire", , Me.Caption

         TxtPerSocQstTypID.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocQstTypIDNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaires Types ID Types Est Obligatoire", , Me.Caption

         CmbPerSocQstTypIDTypNum.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocQstTypIDLibelle)
      Case False

         TxtPerSocQstTypIDLibelle = UCase(TxtPerSocQstTypIDLibelle)

      Case True

         Cancel = True

         MsgBox "Le Champ Questionnaires Types ID Libelle Est Obligatoire", , Me.Caption

         TxtPerSocQstTypIDLibelle.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocQstTypIDMin)
      Case True

      Case False

         Select Case PerSocQstTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

               On Error Resume Next

               Select Case PerSocQstTypIDTypNum
                  Case vbByte

                     VarValeur = CByte(PerSocQstTypIDMin)

                  Case vbInteger

                     VarValeur = CInt(PerSocQstTypIDMin)

                  Case vbLong

                     VarValeur = CLng(PerSocQstTypIDMin)

                  Case vbSingle

                     VarValeur = CSng(PerSocQstTypIDMin)

                  Case vbDouble

                     VarValeur = CDbl(PerSocQstTypIDMin)

                  Case vbCurrency

                     VarValeur = CCur(PerSocQstTypIDMin)

                  Case vbDate

                     VarValeur = CDate(PerSocQstTypIDMin)

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

                     Cancel = True

                     MsgBox "Le Champ Questionnaires Types ID Min Est Invalide", , Me.Caption

                     TxtPerSocQstTypIDMin.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

               On Error GoTo Err_Form_BeforeUpdate

            Case vbBoolean

               PerSocQstTypIDMin = Null

         End Select

   End Select

   Select Case IsNull(PerSocQstTypIDMax)
      Case True

      Case False

         Select Case PerSocQstTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate

               On Error Resume Next

               Select Case PerSocQstTypIDTypNum
                  Case vbByte

                     VarValeur = CByte(PerSocQstTypIDMax)

                  Case vbInteger

                     VarValeur = CInt(PerSocQstTypIDMax)

                  Case vbLong

                     VarValeur = CLng(PerSocQstTypIDMax)

                  Case vbSingle

                     VarValeur = CSng(PerSocQstTypIDMax)

                  Case vbDouble

                     VarValeur = CDbl(PerSocQstTypIDMax)

                  Case vbCurrency

                     VarValeur = CCur(PerSocQstTypIDMax)

                  Case vbDate

                     VarValeur = CDate(PerSocQstTypIDMax)

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

                     Cancel = True

                     MsgBox "Le Champ Questionnaires Types ID Max Est Invalide", , Me.Caption

                     TxtPerSocQstTypIDMax.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

               On Error GoTo Err_Form_BeforeUpdate

            Case vbBoolean

               PerSocQstTypIDMax = Null

         End Select

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

   CreeAnalysePersonnelsSocietesQuestionnaires PerSocQstTypIDPerSocQstTypNum

   TxtPerSocQstTypID.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   TxtPerSocQstTypID.SetFocus

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
   Dim BolIDObligatoire As Boolean

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.Parent!PerSocQstTypNbResultat
      Case Is > Me.RecordsetClone.RecordCount

         Select Case Me.RecordsetClone.RecordCount
            Case vbEmpty

               CmdImprimer.SetFocus

            Case Else

               TxtPerSocQstTypID.SetFocus

         End Select

         ChangeBoutonMode Me, MODEAJOUT

         DoCmd.RunCommand acCmdRecordsGoToNew

         BolIDObligatoire = CbxPerSocQstTypIDObligatoire

         CbxPerSocQstTypIDObligatoire = True

         CbxPerSocQstTypIDObligatoire = BolIDObligatoire

         TxtPerSocQstTypID.SetFocus

      Case Else

         MsgBox "Attention Le Nombre De Résultat Attendu Est Atteint !!!", vbCritical, Me.Caption

   End Select

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
   Dim BolIDObligatoire As Boolean

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerSocQstTypID.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         BolIDObligatoire = CbxPerSocQstTypIDObligatoire

         CbxPerSocQstTypIDObligatoire = True

         CbxPerSocQstTypIDObligatoire = BolIDObligatoire

         TxtPerSocQstTypID.SetFocus

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

         TxtPerSocQstTypID.SetFocus

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

                TxtPerSocQstTypID.SetFocus

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
Private Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   TxtPerSocQstTypID.SetFocus

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

   TxtPerSocQstTypID.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerSocQstTypID.SetFocus

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
Private Sub CmbPerSocQstTypIDTypNum_AfterUpdate()

   On Error GoTo Err_CmbPerSocQstTypIDTypNum_AfterUpdate

   Select Case IsNull(PerSocQstTypIDTypNum)
      Case False

         Select Case PerSocQstTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

               TxtPerSocQstTypIDMin.Enabled = True

               TxtPerSocQstTypIDMax.Enabled = True

            Case vbBoolean

               TxtPerSocQstTypIDMin.Enabled = False

               TxtPerSocQstTypIDMax.Enabled = False

               TxtPerSocQstTypIDMin = Null

               TxtPerSocQstTypIDMax = Null

         End Select

      Case True

         MsgBox "Le Champ Questionnaires Types ID Types Est Obligatoire !", , Me.Caption

         CmbPerSocQstTypIDTypNum.SetFocus

   End Select

Exit_CmbPerSocQstTypIDTypNum_AfterUpdate:

    Exit Sub

Err_CmbPerSocQstTypIDTypNum_AfterUpdate:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocQstTypIDTypNum_AfterUpdate"

    Resume Exit_CmbPerSocQstTypIDTypNum_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocQstTypIDMin_Exit(Cancel As Integer)
   Dim VarValeur As Variant

   On Error GoTo Err_TxtPerSocQstTypIDMin_Exit

   Select Case IsNull(PerSocQstTypIDMin)
      Case True

      Case False

         Select Case IsNull(PerSocQstTypIDTypNum)
            Case False

               Select Case PerSocQstTypIDTypNum
                  Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

                     On Error Resume Next

                     Select Case PerSocQstTypIDTypNum
                        Case vbByte

                           VarValeur = CByte(PerSocQstTypIDMin)

                        Case vbInteger

                           VarValeur = CInt(PerSocQstTypIDMin)

                        Case vbLong

                           VarValeur = CLng(PerSocQstTypIDMin)

                        Case vbSingle

                           VarValeur = CSng(PerSocQstTypIDMin)

                        Case vbDouble

                           VarValeur = CDbl(PerSocQstTypIDMin)

                        Case vbCurrency

                           VarValeur = CCur(PerSocQstTypIDMin)

                        Case vbDate

                           VarValeur = CDate(PerSocQstTypIDMin)

                     End Select

                     Select Case Err.Number
                        Case vbEmpty

                        Case Else

                           Err.Clear

                           MsgBox "Le Champ Questionnaires Types ID Min Est Invalide !", , Me.Caption

                           Cancel = True

                     End Select

                     On Error GoTo Err_TxtPerSocQstTypIDMin_Exit

                  Case vbBoolean

                     PerSocQstTypIDMin = Null

               End Select

            Case True

               MsgBox "Le Champ Questionnaires Types ID Types Est Obligatoire !", , Me.Caption

               CmbPerSocQstTypIDTypNum.SetFocus

         End Select

   End Select

Exit_TxtPerSocQstTypIDMin_Exit:

    Exit Sub

Err_TxtPerSocQstTypIDMin_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocQstTypIDMin_Exit"

    Resume Exit_TxtPerSocQstTypIDMin_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocQstTypIDMax_Exit(Cancel As Integer)
   Dim VarValeur As Variant

   On Error GoTo Err_TxtPerSocQstTypIDMax_Exit

   Select Case IsNull(PerSocQstTypIDMax)
      Case True

      Case False

         Select Case IsNull(PerSocQstTypIDTypNum)
            Case False

               Select Case PerSocQstTypIDTypNum
                  Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

                     On Error Resume Next

                     Select Case PerSocQstTypIDTypNum
                        Case vbByte

                           VarValeur = CByte(PerSocQstTypIDMax)

                        Case vbInteger

                           VarValeur = CInt(PerSocQstTypIDMax)

                        Case vbLong

                           VarValeur = CLng(PerSocQstTypIDMax)

                        Case vbSingle

                           VarValeur = CSng(PerSocQstTypIDMax)

                        Case vbDouble

                           VarValeur = CDbl(PerSocQstTypIDMax)

                        Case vbCurrency

                           VarValeur = CCur(PerSocQstTypIDMax)

                        Case vbDate

                           VarValeur = CDate(PerSocQstTypIDMax)

                     End Select

                     Select Case Err.Number
                        Case vbEmpty

                        Case Else

                           Err.Clear

                           MsgBox "Le Champ Questionnaires Types ID Max Est Invalide !", , Me.Caption

                           TxtPerSocQstTypIDMax.SetFocus

                     End Select

                     On Error GoTo Err_TxtPerSocQstTypIDMax_Exit

                  Case vbBoolean

                     PerSocQstTypIDMax = Null

               End Select

            Case True

               MsgBox "Le Champ Questionnaires Types ID Types Est Obligatoire !", , Me.Caption

               CmbPerSocQstTypIDTypNum.SetFocus

         End Select

   End Select

Exit_TxtPerSocQstTypIDMax_Exit:

    Exit Sub

Err_TxtPerSocQstTypIDMax_Exit:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocQstTypIDMax_Exit"

    Resume Exit_TxtPerSocQstTypIDMax_Exit
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocQstTypID_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocQstTypID_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocQstTypID_KeyPress:

    Exit Sub

Err_TxtPerSocQstTypID_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocQstTypID_KeyPress"

    Resume Exit_TxtPerSocQstTypID_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocQstTypIDLibelle_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocQstTypIDLibelle_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocQstTypIDLibelle_KeyPress:

    Exit Sub

Err_TxtPerSocQstTypIDLibelle_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocQstTypIDLibelle_KeyPress"

    Resume Exit_TxtPerSocQstTypIDLibelle_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocQstTypIDMin_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocQstTypIDMin_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocQstTypIDMin_KeyPress:

    Exit Sub

Err_TxtPerSocQstTypIDMin_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocQstTypIDMin_KeyPress"

    Resume Exit_TxtPerSocQstTypIDMin_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocQstTypIDMax_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_TxtPerSocQstTypIDMax_KeyPress

   Select Case PerSocQstTypIDTypNum
      Case vbString

      Case Else

         KeyAscii = Asc(UCase((Chr$(KeyAscii))))

   End Select

Exit_TxtPerSocQstTypIDMax_KeyPress:

    Exit Sub

Err_TxtPerSocQstTypIDMax_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocQstTypIDMax_KeyPress"

    Resume Exit_TxtPerSocQstTypIDMax_KeyPress
End Sub
