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
'***  Declaration De Procedure Privée                                                         ***
'******************************************************************************

'******************************************************************************
'***  PROCEDURE:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
'******************************************************************************
Private Sub Rafraichir(Optional ByVal BytModeRafraichir As Byte)

   On Error GoTo Err_Rafraichir

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE

         LstLstSocietes.Requery

      Case Else

   End Select

Exit_Rafraichir:

   Exit Sub

Err_Rafraichir:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Rafraichir"

   Resume Exit_Rafraichir
End Sub

'******************************************************************************
'***  PROCEDURE:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
'******************************************************************************
Private Sub ValideSocietes()
   Dim IntEtat As Integer

   On Error GoTo Err_ValideSocietes

   Select Case IsNull(LstLstSocietes.Column(0))
      Case True

      Case False

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
            Case VUEFORMULAIRE

               Forms!FrmFicheSocietes.SetFocus

               DoCmd.ApplyFilter , "SocNum=" & LstLstSocietes.Column(0)

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

              DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & LstLstSocietes.Column(0), acFormEdit

         End Select
   End Select

Exit_ValideSocietes:

   Exit Sub

Err_ValideSocietes:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ValideSocietes"

   Resume Exit_ValideSocietes
End Sub

'******************************************************************************
'***  Declaration D'Evenement Privée                                                          ***
'******************************************************************************

'******************************************************************************
'***  EVENEMENT:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
'******************************************************************************
Private Sub Form_Load()
   Dim StrSQLSocietesDebut As String
   Dim StrSQLSocietesFin As String

   On Error GoTo Err_Form_Load

   StrSQLSocietesDebut = "SELECT SocNum, SocResLibelle, SocCode, " & _
                                       "SocRaison, SocActif, SocAdr, SocVilCodePostal, " & _
                                       "SocVilNom, SocTel1, SocTel2 FROM SelSocietesVilles "

   StrSQLSocietesFin = ";"

   Select Case GStrSocietesFiltre
      Case vbNullString

         LstLstSocietes.RowSource = StrSQLSocietesDebut & StrSQLSocietesFin

      Case Else

         LstLstSocietes.RowSource = StrSQLSocietesDebut & "WHERE SelSocietesVilles." & GStrSocietesFiltre & StrSQLSocietesFin

   End Select

   DoCmd.RunCommand acCmdWindowCascade

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Load"

   Resume Exit_Form_Load
End Sub

'******************************************************************************
'***  EVENEMENT:                                                                                      ***
'******************************************************************************
'***  FONCTION:                                                                                         ***
'******************************************************************************
'***  ENTREE:                                                                                            ***
'***  SORTIE:                                                                                             ***
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
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   ValideSocietes

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
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

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

   OuvreRecherches

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
Private Sub LstLstSocietes_DblClick(Cancel As Integer)

   On Error GoTo Err_LstLstSocietes_DblClick

   ValideSocietes

Exit_LstLstSocietes_DblClick:

   Exit Sub

Err_LstLstSocietes_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstLstSocietes_DblClick"

   Resume Exit_LstLstSocietes_DblClick
End Sub
