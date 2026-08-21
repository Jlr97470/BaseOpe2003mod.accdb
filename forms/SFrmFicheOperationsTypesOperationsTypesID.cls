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
   Dim CtlActive As Control

   On Error GoTo Err_Rafraichir

   Select Case (Me.RecordsetClone.RecordCount = 0) And (Me.NewRecord = False)
      Case True

      Case False

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               On Error Resume Next

               Set CtlActive = ActiveControl

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                      Err.Clear

                      Set CtlActive = TxtPerSocOpeTypID

              End Select

            Case Else

         End Select

         On Error GoTo Err_Rafraichir

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               CmbPerSocOpeTypIDTypNum.Requery

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRDEPLACEMENT

               Select Case IsNull(PerSocOpeTypIDTypNum)
                  Case False

                     Select Case PerSocOpeTypIDTypNum
                        Case vbBoolean

                           TxtPerSocOpeTypIDMin.Enabled = False

                           TxtPerSocOpeTypIDMax.Enabled = False

                        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

                           TxtPerSocOpeTypIDMin.Enabled = True

                           TxtPerSocOpeTypIDMax.Enabled = True

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

   Select Case IsNull(PerSocOpeTypID)
      Case False

         TxtPerSocOpeTypID = UCase(TxtPerSocOpeTypID)

      Case True

         Cancel = True

         MsgBox "Le Champ Operations Types ID Est Obligatoire !", , Me.Caption

         TxtPerSocOpeTypID.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocOpeTypIDNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Operations Types ID Types Est Obligatoire !", , Me.Caption

         CmbPerSocOpeTypIDTypNum.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocOpeTypIDLibelle)
      Case False

         TxtPerSocOpeTypIDLibelle = UCase(TxtPerSocOpeTypIDLibelle)

      Case True

         Cancel = True

         MsgBox "Le Champ Operations Types ID Libelle Est Obligatoire !", , Me.Caption

         TxtPerSocOpeTypIDLibelle.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocOpeTypIDMin)
      Case True

      Case False

         Select Case PerSocOpeTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate

               On Error Resume Next

               Select Case PerSocOpeTypIDTypNum
                  Case vbByte

                     VarValeur = CByte(PerSocOpeTypIDMin)

                  Case vbInteger

                     VarValeur = CInt(PerSocOpeTypIDMin)

                  Case vbLong

                     VarValeur = CLng(PerSocOpeTypIDMin)

                  Case vbSingle

                     VarValeur = CSng(PerSocOpeTypIDMin)

                  Case vbDouble

                     VarValeur = CDbl(PerSocOpeTypIDMin)

                  Case vbCurrency

                     VarValeur = CCur(PerSocOpeTypIDMin)

                  Case vbDate

                     VarValeur = CDate(PerSocOpeTypIDMin)

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

                     Cancel = True

                     MsgBox "Le Champ Operations Types ID Min Est Invalide !", , Me.Caption

                     TxtPerSocOpeTypIDMin.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

               On Error GoTo Err_Form_BeforeUpdate

            Case vbBoolean

               PerSocOpeTypIDMin = Null

         End Select

   End Select

   Select Case IsNull(PerSocOpeTypIDMax)
      Case True

      Case False

         Select Case PerSocOpeTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate

               On Error Resume Next

               Select Case PerSocOpeTypIDTypNum
                  Case vbByte

                     VarValeur = CByte(PerSocOpeTypIDMax)

                  Case vbInteger

                     VarValeur = CInt(PerSocOpeTypIDMax)

                  Case vbLong

                     VarValeur = CLng(PerSocOpeTypIDMax)

                  Case vbSingle

                     VarValeur = CSng(PerSocOpeTypIDMax)

                  Case vbDouble

                     VarValeur = CDbl(PerSocOpeTypIDMax)

                  Case vbCurrency

                     VarValeur = CCur(PerSocOpeTypIDMax)

                  Case vbDate

                     VarValeur = CDate(PerSocOpeTypIDMax)

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

                     Cancel = True

                     MsgBox "Le Champ Operations Types ID Max Est Invalide !", , Me.Caption

                     TxtPerSocOpeTypIDMax.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

               On Error GoTo Err_Form_BeforeUpdate

            Case vbBoolean

               PerSocOpeTypIDMax = Null

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

   CreeAnalysePersonnelsSocietesOperations PerSocOpeTypIDPerSocOpeTypNum

   TxtPerSocOpeTypID.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   TxtPerSocOpeTypID.SetFocus

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

   Select Case Me.Parent!PerSocOpeTypNbResultat
      Case Is > Me.RecordsetClone.RecordCount

         Select Case Me.RecordsetClone.RecordCount
            Case vbEmpty

               CmdImprimer.SetFocus

            Case Else

               TxtPerSocOpeTypID.SetFocus

         End Select

         ChangeBoutonMode Me, MODEAJOUT

         DoCmd.RunCommand acCmdRecordsGoToNew

         BolIDObligatoire = CbxPerSocOpeTypIDObligatoire

         CbxPerSocOpeTypIDObligatoire = True

         CbxPerSocOpeTypIDObligatoire = BolIDObligatoire

         TxtPerSocOpeTypID.SetFocus

      Case Else

         MsgBox "Attention Le Nombre D'Identifient De Résultat Attendu Est Atteint !!!", vbCritical, Me.Caption

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

         TxtPerSocOpeTypID.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         BolIDObligatoire = CbxPerSocOpeTypIDObligatoire

         CbxPerSocOpeTypIDObligatoire = True

         CbxPerSocOpeTypIDObligatoire = BolIDObligatoire

         TxtPerSocOpeTypID.SetFocus

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

         TxtPerSocOpeTypID.SetFocus

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

                TxtPerSocOpeTypID.SetFocus

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

   TxtPerSocOpeTypID.SetFocus

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

   TxtPerSocOpeTypID.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         TxtPerSocOpeTypID.SetFocus

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
Private Sub CmbPerSocOpeTypIDTypNum_AfterUpdate()

   On Error GoTo Err_CmbPerSocOpeTypIDTypNum_AfterUpdate

   Select Case IsNull(PerSocOpeTypIDTypNum)
      Case False

         Select Case PerSocOpeTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

               TxtPerSocOpeTypIDMin.Enabled = True

               TxtPerSocOpeTypIDMax.Enabled = True

            Case vbBoolean

               TxtPerSocOpeTypIDMin.Enabled = False

               TxtPerSocOpeTypIDMax.Enabled = False

               TxtPerSocOpeTypIDMin = Null

               TxtPerSocOpeTypIDMax = Null

         End Select

      Case True

         MsgBox "Le Champ Operations Types ID Types Est Obligatoire !", , Me.Caption

         CmbPerSocOpeTypIDTypNum.SetFocus

   End Select

Exit_CmbPerSocOpeTypIDTypNum_AfterUpdate:

    Exit Sub

Err_CmbPerSocOpeTypIDTypNum_AfterUpdate:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocOpeTypIDTypNum_AfterUpdate"

    Resume Exit_CmbPerSocOpeTypIDTypNum_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocOpeTypIDMin_AfterUpdate()
   Dim VarValeur As Variant

   On Error GoTo Err_TxtPerSocOpeTypIDMin_AfterUpdate

   Select Case IsNull(PerSocOpeTypIDMin)
      Case True

      Case False

         Select Case IsNull(PerSocOpeTypIDTypNum)
            Case False

               Select Case PerSocOpeTypIDTypNum
                  Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

                     On Error Resume Next

                     Select Case PerSocOpeTypIDTypNum
                        Case vbByte

                           VarValeur = CByte(PerSocOpeTypIDMin)

                        Case vbInteger

                           VarValeur = CInt(PerSocOpeTypIDMin)

                        Case vbLong

                           VarValeur = CLng(PerSocOpeTypIDMin)

                        Case vbSingle

                           VarValeur = CSng(PerSocOpeTypIDMin)

                        Case vbDouble

                           VarValeur = CDbl(PerSocOpeTypIDMin)

                        Case vbCurrency

                           VarValeur = CCur(PerSocOpeTypIDMin)

                        Case vbDate

                           VarValeur = CDate(PerSocOpeTypIDMin)

                     End Select

                     Select Case Err.Number
                        Case vbEmpty

                        Case Else

                           Err.Clear

                           MsgBox "Le Champ Operations Types ID Min Est Invalide !", , Me.Caption

                           TxtPerSocOpeTypIDMin.SetFocus

                     End Select

                     On Error GoTo Err_TxtPerSocOpeTypIDMin_AfterUpdate

                  Case vbBoolean

                     PerSocOpeTypIDMin = Null

               End Select

            Case True

               MsgBox "Le Champ Operations Types ID Types Est Obligatoire !", , Me.Caption

               CmbPerSocOpeTypIDTypNum.SetFocus

         End Select

   End Select

Exit_TxtPerSocOpeTypIDMin_AfterUpdate:

    Exit Sub

Err_TxtPerSocOpeTypIDMin_AfterUpdate:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocOpeTypIDMin_AfterUpdate"

    Resume Exit_TxtPerSocOpeTypIDMin_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocOpeTypIDMax_AfterUpdate()
   Dim VarValeur As Variant

   On Error GoTo Err_TxtPerSocOpeTypIDMax_AfterUpdate

   Select Case IsNull(PerSocOpeTypIDMax)
      Case True

      Case False

         Select Case IsNull(PerSocOpeTypIDTypNum)
            Case False

               Select Case PerSocOpeTypIDTypNum
                  Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbString

                     On Error Resume Next

                     Select Case PerSocOpeTypIDTypNum
                        Case vbByte

                           VarValeur = CByte(PerSocOpeTypIDMax)

                        Case vbInteger

                           VarValeur = CInt(PerSocOpeTypIDMax)

                        Case vbLong

                           VarValeur = CLng(PerSocOpeTypIDMax)

                        Case vbSingle

                           VarValeur = CSng(PerSocOpeTypIDMax)

                        Case vbDouble

                           VarValeur = CDbl(PerSocOpeTypIDMax)

                        Case vbCurrency

                           VarValeur = CCur(PerSocOpeTypIDMax)

                        Case vbDate

                           VarValeur = CDate(PerSocOpeTypIDMax)

                     End Select

                     Select Case Err.Number
                        Case vbEmpty

                        Case Else

                           Err.Clear

                           MsgBox "Le Champ Operations Types ID Max Est Invalide !", , Me.Caption

                           TxtPerSocOpeTypIDMax.SetFocus

                     End Select

                     On Error GoTo Err_TxtPerSocOpeTypIDMax_AfterUpdate

                  Case vbBoolean

                     PerSocOpeTypIDMax = Null

               End Select

            Case True

               MsgBox "Le Champ Operations Types ID Types Est Obligatoire !", , Me.Caption

               CmbPerSocOpeTypIDTypNum.SetFocus

         End Select

   End Select

Exit_TxtPerSocOpeTypIDMax_AfterUpdate:

    Exit Sub

Err_TxtPerSocOpeTypIDMax_AfterUpdate:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocOpeTypIDMax_AfterUpdate"

    Resume Exit_TxtPerSocOpeTypIDMax_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocOpeTypID_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocOpeTypID_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocOpeTypID_KeyPress:

    Exit Sub

Err_TxtPerSocOpeTypID_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocOpeTypID_KeyPress"

    Resume Exit_TxtPerSocOpeTypID_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocOpeTypIDLibelle_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocOpeTypIDLibelle_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocOpeTypIDLibelle_KeyPress:

    Exit Sub

Err_TxtPerSocOpeTypIDLibelle_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocOpeTypIDLibelle_KeyPress"

    Resume Exit_TxtPerSocOpeTypIDLibelle_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocOpeTypIDMin_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtPerSocOpeTypIDMin_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtPerSocOpeTypIDMin_KeyPress:

    Exit Sub

Err_TxtPerSocOpeTypIDMin_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocOpeTypIDMin_KeyPress"

    Resume Exit_TxtPerSocOpeTypIDMin_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtPerSocOpeTypIDMax_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_TxtPerSocOpeTypIDMax_KeyPress

   Select Case PerSocOpeTypIDTypNum
      Case vbString

      Case Else

         KeyAscii = Asc(UCase((Chr$(KeyAscii))))

   End Select

Exit_TxtPerSocOpeTypIDMax_KeyPress:

    Exit Sub

Err_TxtPerSocOpeTypIDMax_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtPerSocOpeTypIDMax_KeyPress"

    Resume Exit_TxtPerSocOpeTypIDMax_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocOpeTypIDTypNum_GotFocus()

   On Error GoTo Err_CmbPerSocOpeTypIDTypNum_GotFocus

   CmbPerSocOpeTypIDTypNum.Dropdown

Exit_CmbPerSocOpeTypIDTypNum_GotFocus:

    Exit Sub

Err_CmbPerSocOpeTypIDTypNum_GotFocus:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocOpeTypIDTypNum_GotFocus"

    Resume Exit_CmbPerSocOpeTypIDTypNum_GotFocus
End Sub
