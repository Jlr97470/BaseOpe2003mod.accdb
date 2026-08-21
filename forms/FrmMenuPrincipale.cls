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
'***    Declaration De Procedure Priver                                                        ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub Rafraichir(Optional ByVal BytModeRafraichir As Byte)

   On Error GoTo Err_Rafraichir

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

          LstMenu.Requery

      Case Else

   End Select

Exit_Rafraichir:

   Exit Sub

Err_Rafraichir:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Rafraichir"

   Resume Exit_Rafraichir
End Sub

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub ValideMenu()

    On Error GoTo Err_ValideMenu

    Select Case IsNull(LstMenu.Column(3))
        Case True

        Case False

            DoCmd.RunMacro LstMenu.Column(3)

    End Select

Exit_ValideMenu:

    Exit Sub

Err_ValideMenu:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "ValideMenu"

    Resume Exit_ValideMenu
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
   Dim RsFichier As DAO.Recordset

   On Error GoTo Err_Form_Open

   Set RsFichier = CurrentDb.OpenRecordset("SELECT SelFichiersDetailler.* FROM SelFichiersDetailler WHERE FicType='FICCHM' AND FicCode LIKE '*=*' AND FicValide=True")

   Select Case RsFichier.EOF
      Case True

      Case False

         Me.HelpFile = RsFichier!FicValeur & "\" & RsFichier!FicCode2

   End Select

   RsFichier.Close

Exit_Form_Open:

    Set RsFichier = Nothing

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

   Quit

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
Private Sub ImgMenuPrincipale_Click()

    On Error GoTo Err_ImgMenuPrincipale_Click

    DoCmd.OpenForm "FrmAbout", acNormal, , , acFormReadOnly, acDialog

Exit_ImgMenuPrincipale_Click:

    Exit Sub

Err_ImgMenuPrincipale_Click:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "ImgMenuPrincipale_Click"

    Resume Exit_ImgMenuPrincipale_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TbsMenu_Change()

    On Error GoTo Err_TbsMenu_Change

    TxtGroupeNumVal = TbsMenu.Value + 1

    Rafraichir (RAFRAICHIRCHOIX)

Exit_TbsMenu_Change:

    Exit Sub

Err_TbsMenu_Change:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "TbsMenu_Change"

    Resume Exit_TbsMenu_Change
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

    ValideMenu

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
Private Sub CmdQuitter_Click()

   On Error GoTo Err_CmdQuitter_Click

   TermineMessagerie

Exit_CmdQuitter_Click:

   Exit Sub

Err_CmdQuitter_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdQuitter_Click"

   Resume Exit_CmdQuitter_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstMenu_DblClick(Cancel As Integer)

    On Error GoTo Err_LstMenu_DblClick

    ValideMenu

Exit_LstMenu_DblClick:

    Exit Sub

Err_LstMenu_DblClick:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "LstMenu_DblClick"

    Resume Exit_LstMenu_DblClick
End Sub
