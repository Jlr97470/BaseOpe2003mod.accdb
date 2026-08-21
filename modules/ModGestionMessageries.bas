Attribute VB_Name = "ModGestionMessageries"
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
'***   Declaration De Variable Public                                                           ***
'******************************************************************************

'******************************************************************************
'***   Variable De Controle De L'Utilisateur                                                 ***
'******************************************************************************
Public GStrOrdinateurNom As String
Public GStrOrdinateurDomaine As String
Public GStrUtilisateurNom As String

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
Public Function InitialiseMessagerie() As Boolean
   Dim IntEtat As Integer

   On Error GoTo Err_InitialiseMessagerie

   InitialiseMessagerie = True

   MessagerieConnexion

   DonneEtatFormulaire "FrmGestionMessages", IntEtat

   Select Case IntEtat
     Case VUEFORMULAIRE

        Forms!FrmGestionMessages.Visible = False

     Case Else

        DoCmd.OpenForm "FrmGestionMessages", acNormal, , , acFormEdit

        Forms!FrmGestionMessages.Visible = False

   End Select

Exit_InitialiseMessagerie:

   Exit Function

Err_InitialiseMessagerie:

   InitialiseMessagerie = False

   MsgBox Err.Number & " " & Err.Description, 0, "InitialiseMessagerie"

   Resume Exit_InitialiseMessagerie
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
Public Function TermineMessagerie() As Boolean

   On Error GoTo Err_InitialiseMessagerie

   TermineMessagerie = True

   MessagerieDeConnexion

   FermeFormulaire "FrmMenuPrincipale'"

Exit_InitialiseMessagerie:

   Exit Function

Err_InitialiseMessagerie:

   TermineMessagerie = False

   MsgBox Err.Number & " " & Err.Description, 0, "InitialiseMessagerie"

   Resume Exit_InitialiseMessagerie
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
Public Function MessagerieConnexion() As Boolean
   Dim NetWsc As New WshNetwork

   On Error GoTo Err_MessagerieConnexion

   MessagerieConnexion = True

   GStrUtilisateurNom = UCase(NetWsc.UserName)

   GStrOrdinateurNom = UCase(NetWsc.ComputerName)

   GStrOrdinateurDomaine = UCase(NetWsc.UserDomain)

   CurrentDb.Execute "INSERT INTO TBLMACHINES ( MacNom, MacUtilisateur, MacDomaine ) SELECT '" & GStrOrdinateurNom & "', '" & GStrUtilisateurNom & "', '" & GStrOrdinateurDomaine & "';"

   Set NetWsc = Nothing

Exit_MessagerieConnexion:

   Exit Function

Err_MessagerieConnexion:

   MessagerieConnexion = False

   MsgBox Err.Number & " " & Err.Description, 0, "MessagerieConnexion"

   Resume Exit_MessagerieConnexion
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
Public Function MessagerieDeConnexion() As Boolean

   On Error GoTo Err_MessagerieDeConnexion

   MessagerieDeConnexion = True

   CurrentDb.Execute "DELETE * FROM TBLMACHINES WHERE TBLMACHINES.MacNom='" & GStrOrdinateurNom & "' AND TBLMACHINES.MacUtilisateur='" & GStrUtilisateurNom & "' AND TBLMACHINES.MacDomaine='" & GStrOrdinateurDomaine & "';"

Exit_MessagerieDeConnexion:

   Exit Function

Err_MessagerieDeConnexion:

   MessagerieDeConnexion = False

   MsgBox Err.Number & " " & Err.Description, 0, "MessagerieDeConnexion"

   Resume Exit_MessagerieDeConnexion
End Function
