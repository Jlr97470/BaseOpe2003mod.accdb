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
'***    Declaration De Constante Privée                                                        ***
'******************************************************************************

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub Form_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_Form_BeforeUpdate

   Select Case IsNull(SocTypNum)
      Case True

         Cancel = True

         MsgBox "Le Champ Sociétés Type Code Est Obligatoire !", , Me.Caption

         TxtSocTypNum.SetFocus

         GoTo Exit_Form_BeforeUpdate
      Case False

   End Select

   Select Case IsNull(SocTypLibelle)
     Case True

        Cancel = True

        MsgBox "Le Champ Sociétés Type Libelle Est Obligatoire !", , Me.Caption

        TxtSocTypLibelle.SetFocus

        GoTo Exit_Form_BeforeUpdate
     Case False

   End Select

   TxtSocTypLibelle = StrConv(TxtSocTypLibelle, vbProperCase)

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
Private Sub TxtSocTypCode_KeyPress(KeyAscii As Integer)

    On Error GoTo Err_Form_BeforeUpdate

    KeyAscii = Asc(UCase((Chr$(KeyAscii))))

Exit_Form_BeforeUpdate:

    Exit Sub

Err_Form_BeforeUpdate:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtSocTypCode_KeyPress"

    Resume Exit_Form_BeforeUpdate
End Sub
