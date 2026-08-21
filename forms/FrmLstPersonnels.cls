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

         LstLstPersonnels.Requery

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
Private Sub ValidePersonnels()
   Dim IntEtat As Integer

   On Error GoTo Err_ValidePersonnels

   Select Case IsNull(LstLstPersonnels.Column(0))
      Case True

      Case False

         DonneEtatFormulaire "FrmFichePersonnels", IntEtat

         Select Case IntEtat
            Case VUEFORMULAIRE

               Forms!FrmFichePersonnels.SetFocus

               DoCmd.ApplyFilter , "PerNum=" & LstLstPersonnels.Column(0)

               DoCmd.RunCommand acCmdWindowCascade
           Case Else

               DoCmd.OpenForm "FrmFichePersonnels", acNormal, , "PerNum=" & LstLstPersonnels.Column(0), acFormEdit
      End Select
   End Select

Exit_ValidePersonnels:

   Exit Sub

Err_ValidePersonnels:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ValidePersonnels"

   Resume Exit_ValidePersonnels
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
   Dim StrSQLPersonnelsDebut As String
   Dim StrSQLPersonnelsFin As String

   On Error GoTo Err_Form_Load

   StrSQLPersonnelsDebut = "SELECT PerNum, PerCivCode, PerNom, " & _
                                       "PerPrenom, PerAdr, PerVilCodePostal, " & _
                                       "PerVilNom, PerTel1, PerTel2, PerFax " & _
                                       "FROM SelPersonnelsVilles "

   StrSQLPersonnelsFin = ";"

   Select Case GStrPersonnelsFiltre
      Case vbNullString

         LstLstPersonnels.RowSource = StrSQLPersonnelsDebut & StrSQLPersonnelsFin

      Case Else

         LstLstPersonnels.RowSource = StrSQLPersonnelsDebut & "WHERE SelPersonnelsVilles." & GStrPersonnelsFiltre & StrSQLPersonnelsFin

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

   ValidePersonnels

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
Private Sub LstLstPersonnels_DblClick(Cancel As Integer)

   On Error GoTo Err_LstLstPersonnels_DblClick

   ValidePersonnels

Exit_LstLstPersonnels_DblClick:

   Exit Sub

Err_LstLstPersonnels_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstLstPersonnels_DblClick"

   Resume Exit_LstLstPersonnels_DblClick
End Sub
