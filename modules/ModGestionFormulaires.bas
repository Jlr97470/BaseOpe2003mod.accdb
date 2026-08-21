Attribute VB_Name = "ModGestionFormulaires"
'******************************************************************************
'***                                                                           ***
'******************************************************************************
'***    MODULE:                                                                                          ***
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
'***    Declaration De Constante Public                                                       ***
'******************************************************************************

'******************************************************************************
'***   Constante Pour L'Etat D'Un Formulaire                                               ***
'******************************************************************************
Public Const VUECREATION = 0
Public Const VUEFORMULAIRE = 1
Public Const VUEFEUILLEDONNER = 2
Public Const VUEFERMER = 3

'******************************************************************************
'***    Constante Pour Le Mode De Saisie                                                    ***
'******************************************************************************
Public Const MODENULL = 0
Public Const MODEAJOUT = 1
Public Const MODEMODIF = 2
Public Const MODESUPPRIME = 3
Public Const MODEVALIDE = 4
Public Const MODEDEPLACEMENT = 5

'******************************************************************************
'***    Constante Pour Le Rafraichissement                                                 ***
'******************************************************************************
Public Const RAFRAICHIRVISIBLE = 0
Public Const RAFRAICHIRCHOIX = 1
Public Const RAFRAICHIRDEPLACEMENT = 2

'******************************************************************************
'***    Declaration De Variable Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    Variable Pour Les Filtres                                                                   ***
'******************************************************************************
Public GStrSocietesFiltre As String
Public GStrPersonnelsFiltre As String

'******************************************************************************
'***    Declaration De Fonction Public                                                          ***
'******************************************************************************

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function DonneFiltre(ByVal StrFiltre As String) As String
   Dim RsFiltre As DAO.Recordset

   On Error GoTo Err_DonneFiltre

   DonneFiltre = vbNullString

   StrFiltre = Mid(StrFiltre, InStr(1, StrFiltre, "In") + 4, Len(StrFiltre) - InStr(1, StrFiltre, "In") - 4)

   Set RsFiltre = CurrentDb.OpenRecordset(StrFiltre)

   Select Case RsFiltre.EOF
      Case True

      Case False

         Do Until RsFiltre.EOF = True

            DonneFiltre = DonneFiltre & RsFiltre.Fields(0) & ";"

            RsFiltre.MoveNext

         Loop

            DonneFiltre = Mid(DonneFiltre, 1, Len(DonneFiltre) - 1)

   End Select

   RsFiltre.Close

Exit_DonneFiltre:

   Set RsFiltre = Nothing

   Exit Function

Err_DonneFiltre:

   DonneFiltre = vbNullString

   MsgBox Err.Number & " " & Err.Description, vbCritical, "DonneFiltre"

   Resume Exit_DonneFiltre
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function OuvreRecherches(Optional ByVal bytTypeOuvre As Byte) As Boolean

   On Error GoTo Err_OuvreRecherches

   FermeFormulaire "FrmMenuPrincipale"

   DoCmd.OpenForm "FrmRecherche", acNormal, , , acFormEdit

   Select Case bytTypeOuvre
      Case 0

      Case Else

         Forms!FrmRecherche!TbsRecherche.Tabs(bytTypeOuvre).Selected = True

   End Select

Exit_OuvreRecherches:

   Exit Function

Err_OuvreRecherches:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "OuvreRecherches"

   Resume Exit_OuvreRecherches
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function AjouteCivilites(Optional ByVal StrCivilite As String) As Boolean
   Dim IntEtat As Integer

   On Error GoTo Err_AjouteCivilites

   AjouteCivilites = True

   Select Case MsgBox("Voulez Vous Créer Une Nouvelle Civilités ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DonneEtatFormulaire "FrmLstCivilites", IntEtat

         Select Case IntEtat
           Case VUEFORMULAIRE

              Forms!FrmLstCivilites.SetFocus

              DoCmd.ApplyFilter , "CivCode=" & StrCivilite

              DoCmd.RunCommand acCmdRecordsGoToNew

              DoCmd.RunCommand acCmdWindowCascade

           Case Else

              DoCmd.OpenForm "FrmLstCivilites", acNormal, , , acFormAdd

         End Select

         Select Case StrCivilite
            Case vbNullString

            Case Else

               Forms!FrmLstCivilites.TxtCivCode = StrCivilite

         End Select
      Case vbNo

   End Select

Exit_AjouteCivilites:

   Exit Function

Err_AjouteCivilites:

   AjouteCivilites = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteCivilites"

   Resume Exit_AjouteCivilites
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function AjouteVilles(Optional ByVal strCodePostal As String) As Boolean
   Dim IntEtat As Integer

   On Error GoTo Err_AjouteVilles

   AjouteVilles = True

   Select Case MsgBox("Voulez Vous Créer Une Nouvelle Villes ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DonneEtatFormulaire "FrmLstVilles", IntEtat

         Select Case IntEtat
            Case VUEFORMULAIRE

               Forms!FrmLstVilles.SetFocus

               DoCmd.ApplyFilter , "VilCodePostal=" & strCodePostal

               DoCmd.RunCommand acCmdRecordsGoToNew

               DoCmd.RunCommand acCmdWindowCascade

            Case Else

               DoCmd.OpenForm "FrmLstVilles", acNormal, , , acFormAdd

         End Select

         Select Case strCodePostal
            Case vbNullString

            Case Else

               Forms!FrmLstVilles.TxtVilCodePostal = strCodePostal

         End Select
      Case vbNo

   End Select

Exit_AjouteVilles:

   Exit Function

Err_AjouteVilles:

   AjouteVilles = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteVilles"

   Resume Exit_AjouteVilles
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function AjouteFonctions(Optional ByVal strFonction As String) As Boolean
   Dim IntEtat As Integer

   On Error GoTo Err_AjouteFonctions

   AjouteFonctions = True

   Select Case MsgBox("Voulez Vous Créer Une Nouvelle Fonctions ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DonneEtatFormulaire "FrmLstContactsFonctions", IntEtat

         Select Case IntEtat
           Case VUEFORMULAIRE

              Forms!FrmLstContactsFonctions.SetFocus

              DoCmd.ApplyFilter , "ConFonCode=" & strFonction

              DoCmd.RunCommand acCmdRecordsGoToNew

              DoCmd.RunCommand acCmdWindowCascade

           Case Else

               DoCmd.OpenForm "FrmLstContactsFonctions", acNormal, , , acFormAdd

         End Select

         Select Case strFonction
            Case vbNullString

            Case Else

               Forms!FrmLstContactsFonctions.TxtConFonCode = strFonction

         End Select
      Case vbNo

   End Select

Exit_AjouteFonctions:

   Exit Function

Err_AjouteFonctions:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteFonctions"

   Resume Exit_AjouteFonctions
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function AjouteSocietes(Optional ByVal StrRaison As String) As Boolean
   Dim IntEtat As Integer

   On Error GoTo Err_AjouteSocietes

   AjouteSocietes = True

   Select Case MsgBox("Voulez Vous Créer Une Nouvelle Sociétés ? ", vbYesNo, Screen.ActiveForm.Caption)
      Case vbYes

         DonneEtatFormulaire "FrmFicheSocietes", IntEtat

         Select Case IntEtat
           Case VUEFORMULAIRE

              Forms!FrmFicheSocietes.SetFocus

              DoCmd.ApplyFilter , "SocNum=0"

              DoCmd.RunCommand acCmdWindowCascade

           Case Else

              DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=0", acFormEdit

       End Select

         Select Case StrRaison
            Case vbNullString

            Case Else

               Forms!FrmFicheSocietes.TxtSocRaison.DefaultValue = """ & StrRaison & """

         End Select
      Case vbNo

   End Select

Exit_AjouteSocietes:

   Exit Function

Err_AjouteSocietes:

   AjouteSocietes = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "AjouteSocietes"

   Resume Exit_AjouteSocietes
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function ChangeBoutonMode(ByRef FrmForm As Form, ByVal IntMode As Integer) As Boolean

   On Error GoTo Err_ChangeBoutonMode

   ChangeBoutonMode = True

   FrmForm.GIntModeOld = FrmForm.GIntMode

   FrmForm.GIntMode = IntMode

   Select Case IntMode
      Case MODEAJOUT

         FrmForm.AllowAdditions = True

         FrmForm.AllowDeletions = False

         FrmForm.AllowEdits = False

         FrmForm.NavigationButtons = False

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdValider.Enabled = True

         FrmForm.CmdAnnuler.Enabled = True

         FrmForm.CmdFermer.Enabled = False

         FrmForm.CmdAfficherTous.Enabled = False

         FrmForm.RecButton.BackStyle = 1

         FrmForm.RecButton.BackColor = 65280

      Case MODEMODIF

         FrmForm.AllowAdditions = False

         FrmForm.AllowDeletions = False

         FrmForm.AllowEdits = True

         FrmForm.NavigationButtons = False

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdValider.Enabled = True

         FrmForm.CmdAnnuler.Enabled = True

         FrmForm.CmdFermer.Enabled = False

         FrmForm.CmdAfficherTous.Enabled = False

         FrmForm.RecButton.BackStyle = 1

         FrmForm.RecButton.BackColor = 16711680

      Case MODESUPPRIME

         FrmForm.AllowAdditions = False

         FrmForm.AllowDeletions = True

         FrmForm.AllowEdits = False

         FrmForm.NavigationButtons = False

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdValider.Enabled = False

         FrmForm.CmdAnnuler.Enabled = False

         FrmForm.CmdFermer.Enabled = False

         FrmForm.CmdAfficherTous.Enabled = False

         FrmForm.RecButton.BackStyle = 1

         FrmForm.RecButton.BackColor = 255

      Case MODEDEPLACEMENT

         FrmForm.AllowAdditions = False

         FrmForm.AllowDeletions = False

         FrmForm.AllowEdits = False

         FrmForm.NavigationButtons = True

         FrmForm.CmdAjouter.Enabled = True

         FrmForm.CmdModifier.Enabled = True

         FrmForm.CmdSupprimer.Enabled = True

         FrmForm.CmdValider.Enabled = False

         FrmForm.CmdAnnuler.Enabled = False

         FrmForm.CmdFermer.Enabled = True

         FrmForm.CmdAfficherTous.Enabled = True

         FrmForm.RecButton.BackStyle = vbEmpty

         FrmForm.RecButton.BackColor = 255

      Case MODEVALIDE

         FrmForm.NavigationButtons = False

         FrmForm.CmdAjouter.Enabled = False

         FrmForm.CmdModifier.Enabled = False

         FrmForm.CmdSupprimer.Enabled = False

         FrmForm.CmdValider.Enabled = False

         FrmForm.CmdAnnuler.Enabled = False

         FrmForm.CmdFermer.Enabled = False

         FrmForm.CmdAfficherTous.Enabled = False

         FrmForm.RecButton.BackStyle = 1

         FrmForm.RecButton.BackColor = 65535

   End Select

Exit_ChangeBoutonMode:

   Exit Function

Err_ChangeBoutonMode:

   ChangeBoutonMode = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ChangeBoutonMode"

   Resume Exit_ChangeBoutonMode
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function FermeFormulaire(Optional ByVal StrFormName As String) As Boolean
   Dim FrmForm As Form

   On Error GoTo Err_FermeFormulaire

   For Each FrmForm In Forms

      Select Case FrmForm.Name
         Case StrFormName

         Case Else

            DoCmd.Close acForm, FrmForm.Name, acSavePrompt

            FermeFormulaire = True

      End Select

   Next

Exit_FermeFormulaire:

   Exit Function

Err_FermeFormulaire:

   FermeFormulaire = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "FermeFormulaire"

   Resume Exit_FermeFormulaire
End Function


'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Public Function DonneEtatFormulaire(ByVal StrFormName As String, ByRef intFormEtat As Integer) As Boolean
   Dim FrmDataBase As Forms
   Dim FrmForm As Form

   On Error GoTo Exit_DonneEtatFormulaire

   Set FrmDataBase = Forms

   DonneEtatFormulaire = False

   For Each FrmForm In FrmDataBase

      Select Case FrmForm.Name
         Case StrFormName

            intFormEtat = FrmForm.CurrentView

            DonneEtatFormulaire = True

            GoTo Exit_DonneEtatFormulaire

         Case Else

      End Select
   Next

   intFormEtat = VUEFERMER

   DonneEtatFormulaire = False

Exit_DonneEtatFormulaire:

   Exit Function

Err_DonneEtatFormulaire:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "DonneEtatFormulaire"

   Resume Exit_DonneEtatFormulaire
End Function

'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
