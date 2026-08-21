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
'***    Declaration de constante privee                                                        ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE :                                                                                  ***
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

                      Set CtlActive = LstParType

              End Select

              On Error GoTo Err_Rafraichir

            Case Else

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               LstParType.Requery

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo Err_Form_BeforeUpdate

    Select Case IsNull(ParType)
      Case True

         Cancel = True

         MsgBox "Le Champ Paramétre Type Est Obligatoire !", , Me.Caption

         LstParType.SetFocus

         GoTo Exit_Form_BeforeUpdate
      Case False

   End Select

   Select Case IsNull(ParCode)
      Case True

         Cancel = True

         MsgBox "Le Champ Paramétre Code Est Obligatoire !", , Me.Caption

         TxtParCode.SetFocus

         GoTo Exit_Form_BeforeUpdate
      Case False

   End Select

   Select Case IsNull(ParValeur)
      Case True

         Cancel = True

         MsgBox "Le Champ Paramétre Valeur Est Obligatoire !", , Me.Caption

         TxtParValeur.SetFocus

         GoTo Exit_Form_BeforeUpdate
      Case False

   End Select

   LstParType = UCase(LstParType)

   TxtParCode = UCase(TxtParCode)

   TxtParValeur = UCase(TxtParValeur)

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
Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   DoCmd.RunCommand acCmdWindowCascade

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Open"

   Resume Err_Form_Open
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_Activate()

   On Error GoTo Err_Form_Activate

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_Current()

   On Error GoTo Err_Form_Current

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub LstParType_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_LstParType_KeyPress

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit_LstParType_KeyPress:

   Exit Sub

Err_LstParType_KeyPress:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstParType_KeyPress"

   Resume Exit_LstParType_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub TxtParCode_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_TxtParCode_KeyPress

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit_TxtParCode_KeyPress:

   Exit Sub

Err_TxtParCode_KeyPress:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtParCode_KeyPress"

   Resume Exit_TxtParCode_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub TxtParValeur_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_TxtParValeur_KeyPress

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit_TxtParValeur_KeyPress:

   Exit Sub

Err_TxtParValeur_KeyPress:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtParValeur_KeyPress"

   Resume Exit_TxtParValeur_KeyPress
End Sub
