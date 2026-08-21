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
   Dim StrSQLFichiersPreference As String
   Dim RsMachines As DAO.Recordset
   Dim CtlActive As Access.Control

   On Error GoTo Err_Rafraichir

   StrSQLFichiersPreference = "SELECT FicCode1, FicValeur FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe] UNION SELECT TBLMACHINES.MacNom AS FicCode1, '' AS FicValeur FROM TBLMACHINES GROUP BY TBLMACHINES.MacNom, '' HAVING TBLMACHINES.MacNom not In (SELECT FicCode1 FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe];) UNION SELECT TBLMACHINES.MacUtilisateur AS FicCode1,'' As FicValeur FROM TBLMACHINES GROUP BY TBLMACHINES.MacUtilisateur,'' HAVING TBLMACHINES.MacUtilisateur not In (SELECT FicCode1 FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe];) UNION SELECT TBLMACHINES.MacDomaine AS FicCode1,'' As FicValeur FROM TBLMACHINES GROUP BY TBLMACHINES.MacDomaine,'' HAVING TBLMACHINES.MacDomaine not In (SELECT FicCode1 FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe];);"

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE

         On Error Resume Next

         Set CtlActive = Me.ActiveControl

         Select Case Err.Number
            Case 0

            Case Else

                Err.Clear

                Set CtlActive = CmbFichiersListe

        End Select

        On Error GoTo Err_Rafraichir

   End Select

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE, RAFRAICHIRCHOIX

         On Error Resume Next

         Set RsMachines = CurrentDb.OpenRecordset("SELECT * FROM TBLMACHINES;", dbOpenForwardOnly)

         Select Case RsMachines.RecordCount
            Case vbEmpty

               CmbFichiersPreference.RowSource = "SELECT FicCode1, FicValeur FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe] ;"

            Case Else

               CmbFichiersPreference.RowSource = StrSQLFichiersPreference

         End Select

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               CmbFichiersPreference.RowSource = "SELECT FicCode1, FicValeur FROM SelFichiersDetailler WHERE SelFichiersDetailler.FicCode2=[CmbFichiersListe] ;"

         End Select

         RsMachines.Close

         On Error GoTo Err_Rafraichir

         CmbFichiersListe.Requery

         CmbFichiersPreference.Requery

         LstFichiersManquant.Requery

         LstFichiersValide.Requery

      Case Else

   End Select

   Select Case BytModeRafraichir
      Case RAFRAICHIRVISIBLE

         CtlActive.SetFocus

         Set CtlActive = Nothing

      Case Else

   End Select

Exit_Rafraichir:

   Set RsMachines = Nothing

   Exit Sub

Err_Rafraichir:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Rafraichir"

   Resume Exit_Rafraichir
End Sub

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
Private Sub ChangeEtatBouton()
   Dim CtlActive As Control

   On Error GoTo Err_ChangeEtatBouton

   On Error Resume Next

   Set CtlActive = Me.ActiveControl

   Select Case Err.Number
      Case 0

      Case Else

          Err.Clear

          Set CtlActive = CmbFichiersListe
  End Select

  On Error GoTo Err_ChangeEtatBouton

   Select Case Not IsNull(CmbFichiersListe) And Not IsNull(CmbFichiersPreference)
      Case True

         CmdValider.Enabled = True

         CmdSupprimer.Enabled = True

      Case False

         CmdValider.Enabled = False

         CmdSupprimer.Enabled = False

   End Select

   CtlActive.SetFocus

   Set CtlActive = Nothing

Exit_ChangeEtatBouton:

   Exit Sub

Err_ChangeEtatBouton:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ChangeEtatBouton"

   Resume Exit_ChangeEtatBouton
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
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub Form_Load()

   On Error GoTo Err_Form_Open

   Rafraichir (RAFRAICHIRVISIBLE)

   ChangeEtatBouton

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
Private Sub CmbFichiersListe_AfterUpdate()

   On Error GoTo Err_CmbFichiersListe_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   CmbFichiersPreference = Null

   TxtFichiersRepertoire = vbNullString

   ChangeEtatBouton

Exit_CmbFichiersListe_AfterUpdate:

    Exit Sub

Err_CmbFichiersListe_AfterUpdate:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbFichiersListe_AfterUpdate"

   Resume Exit_CmbFichiersListe_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmbFichiersPreference_AfterUpdate()

   On Error GoTo Err_CmbFichiersPreference_AfterUpdate

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CmbFichiersPreference.Column(1)
      Case vbNullString

      Case Else

         TxtFichiersRepertoire = CmbFichiersPreference.Column(1)

   End Select

   ChangeEtatBouton

Exit_CmbFichiersPreference_AfterUpdate:

    Exit Sub

Err_CmbFichiersPreference_AfterUpdate:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbFichiersPreference_AfterUpdate"

   Resume Exit_CmbFichiersPreference_AfterUpdate
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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmdRepertoireFichier_Click()
   Dim FsoSystemFichier As New FileSystemObject
   Dim FilFichier As File
   Dim CdgFichiers As Office.FileDialog

   On Error GoTo Err_CmdRepertoireFichier_Click

   Set CdgFichiers = Application.FileDialog(msoFileDialogOpen)

   CdgFichiers.Title = "CP 2003 - Indiquer L'Emplacement Du Fichier " & CmbFichiersListe

   CdgFichiers.Filters.Add TxtExtension, "*." & TxtExtension & "", 1

   CdgFichiers.InitialFileName = IIf(IsNull(CmbFichiersListe), vbNullString, IIf(TxtFichiersRepertoire = vbNullString, "", TxtFichiersRepertoire & "\") & Nz(CmbFichiersListe))

   CdgFichiers.Show

   Select Case FsoSystemFichier.FileExists(CdgFichiers.SelectedItems(1))
      Case True

         Set FilFichier = FsoSystemFichier.GetFile(CdgFichiers.SelectedItems(1))

         CmbFichiersListe = UCase(FilFichier.Name)

         TxtFichiersRepertoire = UCase(FilFichier.ParentFolder.Path)

      Case False

         MsgBox "Aucun Fichier N'a Ete Choisi !", vbInformation, Me.Caption

   End Select

   Rafraichir (RAFRAICHIRCHOIX)

   Select Case CmbFichiersPreference.ListIndex
      Case -1

         CmbFichiersPreference = "DEFAULT"

      Case Else

   End Select

   ChangeEtatBouton

Exit_CmdRepertoireFichier_Click:

   Set FilFichier = Nothing

   Set FsoSystemFichier = Nothing

   Exit Sub

Err_CmdRepertoireFichier_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdRepertoireFichier_Click"

   Resume Exit_CmdRepertoireFichier_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmdValider_Click()
   Dim FsoSystemFichier As New Scripting.FileSystemObject

   On Error GoTo Err_CmdValider_Click

   Select Case FsoSystemFichier.FileExists(TxtFichiersRepertoire & "\" & CmbFichiersListe)
      Case True

           CurrentDb.Execute "UPDATE TBLFICHIERS SET FicValide=False WHERE FicType='FIC" & TxtExtension & "' AND FicCode LIKE '*=" & CmbFichiersListe & "' ;"

           Select Case IsNull(DLookup("FicType", "TBLFICHIERS", "FicType='FIC" & TxtExtension & "' AND FicCode LIKE '" & CmbFichiersPreference & "=" & CmbFichiersListe & "'"))
              Case True

                 CurrentDb.Execute "INSERT INTO TBLFICHIERS (FicType,FicCode,FicValeur,FicValide) VALUES('FIC" & UCase(Right(CmbFichiersListe, Len(CmbFichiersListe) - InStrRev(CmbFichiersListe, "."))) & "','" & CmbFichiersPreference & "=" & CmbFichiersListe & "','" & TxtFichiersRepertoire & "',True) ;"

              Case False

                 CurrentDb.Execute "UPDATE TBLFICHIERS SET FicValeur='" & TxtFichiersRepertoire & "', FicValide=True WHERE FicType='FIC" & UCase(Right(CmbFichiersListe, Len(CmbFichiersListe) - InStrRev(CmbFichiersListe, "."))) & "' AND FicCode='" & CmbFichiersPreference & "=" & CmbFichiersListe & "' ;"

           End Select

           ChangeTableInformationConnexion

        Case False

            MsgBox "Chemin Invalide !", , Me.Caption
   End Select

   Rafraichir (RAFRAICHIRCHOIX)

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmdSupprimer_Click()

   On Error GoTo Err_CmdSupprimer_Click

   CurrentDb.Execute "DELETE * FROM  TBLFICHIERS WHERE FicType='FIC" & TxtExtension & "' AND FicCode LIKE '" & CmbFichiersPreference & "=" & CmbFichiersListe & "' ;"

   Rafraichir (RAFRAICHIRCHOIX)

   CmbFichiersListe = Null

   CmbFichiersPreference = Null

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
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmbFichiersListe_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_CmbFichiersListe_KeyPress

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit_CmbFichiersListe_KeyPress:

   Exit Sub

Err_CmbFichiersListe_KeyPress:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbFichiersListe_KeyPress"

   Resume Exit_CmbFichiersListe_KeyPress
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                           ***
'***    SORTIE:                                                                                            ***
'******************************************************************************
Private Sub CmbFichiersPreference_KeyPress(KeyAscii As Integer)

   On Error GoTo Err_CmbFichiersPreference_KeyPress

   KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit_CmbFichiersPreference_KeyPress:

   Exit Sub

Err_CmbFichiersPreference_KeyPress:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbFichiersListe_KeyPress"

   Resume Exit_CmbFichiersPreference_KeyPress
End Sub
