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
'***    Declaration De Constante Privée                                                       ***
'******************************************************************************

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

               Set CtlActive = Me.ActiveControl

               Select Case Err.Number
                  Case 0

                  Case Else

                      Err.Clear

                      Set CtlActive = Me.TxtVilCodePostal
              End Select

              On Error GoTo Err_Rafraichir
         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               TxtVilDepNom.Requery

               TxtVilDepPays.Requery

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
Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   DoCmd.RunCommand acCmdWindowCascade

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Load"

   Resume Err_Form_Load
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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_BeforeUpdate(Cancel As Integer)

    On Error GoTo Err_Form_BeforeUpdate

   Select Case IsNull(VilNom)
      Case True

         Cancel = True

         MsgBox "Le Champ Ville Nom Est Obligatoire !", , Me.Caption

         TxtVilNom.SetFocus

         GoTo Exit_Form_BeforeUpdate
      Case False

   End Select

   Select Case IsNull(VilCodePostal)
      Case True

         Cancel = True

         MsgBox "Le Champ Ville Code Postal Est Obligatoire !", , Me.Caption

         TxtVilCodePostal.SetFocus

         GoTo Exit_Form_BeforeUpdate
      Case False

   End Select

   TxtVilNom = UCase(TxtVilNom)

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub TxtVilNom_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_TxtVilNom_KeyPress

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_TxtVilNom_KeyPress:

    Exit Sub

Err_TxtVilNom_KeyPress:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtVilNom_KeyPress"

    Resume Exit_TxtVilNom_KeyPress
End Sub
