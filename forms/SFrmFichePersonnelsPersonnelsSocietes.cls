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
'***    Variable De Controle De Mode Saisie                                                 ***
'******************************************************************************
Public GIntMode As Integer
Public GIntModeOld As Integer
Public GBlnModif As Boolean

'******************************************************************************
'***    Declaration Procedure Priver                                                             ***
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
   Dim StrSQLPerSocSocNum As String
   Dim StrSQLPerSocPerSocTypCode As String
   Dim CtlActive As Control

   On Error GoTo Err_Rafraichir

   StrSQLPerSocSocNum = "SELECT SocNum, SocCode, SocEnseigne, SocRaison, SocResLibelle, SocTypLibelle, SocAdr, SocVilCodePostal, SocVilNom, SocSocTypNum, SocActif, SocAdhIpe, SocAdhConcours " & _
                                       " FROM SelSocietesVilles ;"

   StrSQLPerSocPerSocTypCode = "SELECT PerSocTypCode, PerSocTypLibelle, PerSocTypRang FROM SelPersonnelsSocietesTypes " & _
                                                   " ORDER BY SelPersonnelsSocietesTypes.PerSocTypRang;"

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

                      Set CtlActive = CmbPerSocPerSocTypCode

              End Select

              On Error GoTo Err_Rafraichir

         End Select

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               CmbPerSocSocNum.RowSource = StrSQLPerSocSocNum

               CmbPerSocSocNum.Requery

               CmbPerSocPerSocTypCode.RowSource = StrSQLPerSocPerSocTypCode

               CmbPerSocPerSocTypCode.Requery

               TxtPerSocActDate.Requery

               TxtPerSocAdhDate.Requery

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
Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Me.Parent.ChangeSFrmButtonMode Me

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
   Dim IntNbAdherents As Integer
   Dim IntNbActifs As Integer

   On Error GoTo Err_Form_BeforeUpdate

   Select Case MsgBox("Voulez-Vous Mettre à Jour Le Profil ?", vbYesNo, Me.Caption)
      Case vbYes

      Case vbNo

         Cancel = True

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsDate(PerSocDateCre)
      Case True

      Case False

         Cancel = True

         MsgBox "Le Champ Personnels Sociétés Date De Creation Est Invalide !", , Me.Caption

         GoTo Exit_Form_BeforeUpdate

   End Select

   PerSocDateMaj = Date

   Select Case IsNull(PerSocSocNum)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Personnels Sociétés Sociétés Numero Est Obligatoire !", , Me.Caption

         CmbPerSocSocNum.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case IsNull(PerSocPerSocTypCode)
      Case False

      Case True

         Cancel = True

         MsgBox "Le Champ Personnels Sociétés Personnels Sociétés Types Code Est Obligatoire !", , Me.Caption

         CmbPerSocPerSocTypCode.SetFocus

         GoTo Exit_Form_BeforeUpdate

   End Select

   Select Case PerSocActif
      Case True

         Select Case PerSocPerSocTypCode
            Case TYPERESPONSABLE

            Case TYPEMARKETING, TYPEDIRECTEURVENTE, TYPERESPONSABLEVENTE, TYPEFORMATEUR, TYPEAMBASSADRICE, TYPEVENDEUR

               Select Case CmbPerSocPerSocTypCode.Column(10)
                  Case True

                  Case False

                     Cancel = True

                     MsgBox "Attention !! le point de vente n'est pas actif !!", , Me.Caption

                     CbxPerSocAdherent.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

         End Select

         Select Case PerSocPerSocTypCode
            Case TYPERESPONSABLE

            Case TYPEMARKETING, TYPEDIRECTEURVENTE, TYPERESPONSABLEVENTE, TYPEFORMATEUR

               Select Case CmbPerSocPerSocTypCode.Column(9)
                  Case TYPESIEGE

                  Case Else

                     Cancel = True

                     MsgBox "Attention !! les responsables de marketing, les directeur de ventes regional, responsable des ventes par region et les formateurs ne peuvent etre adherent que sur le siege chanel !!", , Me.Caption

                     CbxPerSocAdherent.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

            Case TYPEAMBASSADRICE, TYPEVENDEUR

         End Select

         Select Case PerSocPerSocTypCode
            Case TYPERESPONSABLE

            Case TYPEMARKETING, TYPEDIRECTEURVENTE, TYPERESPONSABLEVENTE, TYPEFORMATEUR, TYPEAMBASSADRICE, TYPEVENDEUR

               IntNbActifs = DCount("PerSocNum", "TBLPERSONNELSSOCIETES", "PerSocNum<>" & PerSocNum & " AND PerSocPerSocTypCode In ('AMB','VND') AND PerSocActif=True")

               Select Case IntNbActifs
                  Case 0

                  Case Else

                     Cancel = True

                     MsgBox "Attention !! un autre profil est actif !! Desactiver le !!", , Me.Caption

                     CbxPerSocActif.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

         End Select

         Select Case PerSocAdherent
            Case True

               Select Case PerSocPerSocTypCode
                  Case TYPERESPONSABLE

                     Cancel = True

                     MsgBox "Attention !! les responsables ne peuvent etre adherent !!", , Me.Caption

                     CbxPerSocAdherent.SetFocus

                     GoTo Exit_Form_BeforeUpdate

                  Case TYPEMARKETING, TYPEDIRECTEURVENTE, TYPERESPONSABLEVENTE, TYPETECHNICO, TYPEFORMATEUR, TYPEAMBASSADRICE, TYPEVENDEUR

                     Select Case CmbPerSocPerSocTypCode.Column(11)
                        Case True

                        Case False

                           Cancel = True

                           MsgBox "Attention !! le point de vente n'est pas adherents instants pour elle !!", , Me.Caption

                           CbxPerSocAdherent.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

               Select Case PerSocPerSocTypCode
                  Case TYPERESPONSABLE

                  Case TYPEAMBASSADRICE, TYPEVENDEUR

                  Case TYPEMARKETING, TYPEDIRECTEURVENTE, TYPERESPONSABLEVENTE, TYPETECHNICO, TYPEFORMATEUR

                     Select Case CmbPerSocPerSocTypCode.Column(9)
                        Case TYPESIEGE

                        Case Else

                           Cancel = True

                           MsgBox "Attention !! les responsables de marketing, les directeur de ventes regional, responsable des ventes par region, les technico et les formateurs ne peuvent etre adherent que sur le siege chanel !!", , Me.Caption

                           CbxPerSocAdherent.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

               IntNbAdherents = DCount("PerSocNum", "TBLPERSONNELSSOCIETES", "PerSocNum<>" & PerSocNum & " AND PerSocPerSocTypCode In ('AMB','VND') AND PerSocAdherent=True")

               Select Case IntNbAdherents
                  Case 0

                  Case Else

                     Cancel = True

                     MsgBox "Attention !! un autre profil est adherent !! Desadherer le !!", , Me.Caption

                     CbxPerSocAdherent.SetFocus

                     GoTo Exit_Form_BeforeUpdate

               End Select

            Case False

         End Select

      Case False

         Select Case PerSocAdherent
            Case True

               Cancel = True

               MsgBox "Attention !! le profil n'est pas actif !!", , Me.Caption

               CbxPerSocActif.SetFocus

               GoTo Exit_Form_BeforeUpdate

            Case False

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
   Dim BlnAdhEtat As Variant
   Dim BlnActEtat As Variant
   Dim DatAdh As Date
   Dim DatAct As Date

   On Error GoTo Err_Form_AfterUpdate

   BlnAdhEtat = DLookup("PerSocAdhEtat", "TBLPERSONNELSSOCIETESADHERENTS", "PerSocAdhPerSocNum=" & PerSocNum)

   Select Case IsNull(BlnAdhEtat)
      Case True

         CurrentDb.Execute "INSERT INTO TBLPERSONNELSSOCIETESADHERENTS (PerSocAdhPerSocNum,PerSocAdhDate,PerSocAdhEtat) VALUES (" & PerSocNum & ",Date() ," & PerSocAdherent & ") ;"

      Case False

         Select Case CBool(PerSocAdherent)
            Case Is <> CBool(BlnAdhEtat)

               DatAdh = DLookup("PerSocAdhDate", "TBLPERSONNELSSOCIETESADHERENTS", "PerSocAdhPerSocNum=" & PerSocNum)

               Select Case DatAdh
                  Case Date

                     CurrentDb.Execute "UPDATE TBLPERSONNELSSOCIETESADHERENTS SET TBLPERSONNELSSOCIETESADHERENTS.PerSocAdhEtat = " & PerSocAdherent & " WHERE TBLPERSONNELSSOCIETESADHERENTS.PerSocAdhPerSocNum=" & PerSocNum & " AND TBLPERSONNELSSOCIETESADHERENTS.PerSocAdhDate=Date();"

                  Case Else

                     CurrentDb.Execute "INSERT INTO TBLPERSONNELSSOCIETESADHERENTS (PerSocAdhPerSocNum,PerSocAdhDate,PerSocAdhEtat) VALUES (" & PerSocNum & ",Date()," & PerSocAdherent & ") ;"

               End Select

            Case Else

         End Select

   End Select

   BlnActEtat = DLookup("PerSocActEtat", "TBLPERSONNELSSOCIETESACTIFS", "PerSocActPerSocNum=" & PerSocNum)

   Select Case IsNull(BlnActEtat)
      Case True

         CurrentDb.Execute "INSERT INTO TBLPERSONNELSSOCIETESACTIFS (PerSocActPerSocNum,PerSocActDate,PerSocActEtat) VALUES (" & PerSocNum & ",Date()," & PerSocActif & ") ;"

      Case False

         Select Case CBool(PerSocActif)
            Case Is <> CBool(BlnActEtat)

               DatAdh = DLookup("PerSocAdhDate", "TBLPERSONNELSSOCIETESADHERENTS", "PerSocAdhPerSocNum=" & PerSocNum)

               Select Case DatAdh
                  Case Date

                     CurrentDb.Execute "UPDATE TBLPERSONNELSSOCIETESACTIFS SET TBLPERSONNELSSOCIETESACTIFS.PerSocActEtat = " & PerSocActif & " WHERE TBLPERSONNELSSOCIETESACTIFS.PerSocActPerSocNum=" & PerSocNum & " AND TBLPERSONNELSSOCIETESACTIFS.PerSocActDate=Date();"

                  Case Else

                     CurrentDb.Execute "INSERT INTO TBLPERSONNELSSOCIETESACTIFS (PerSocActPerSocNum,PerSocActDate,PerSocActEtat) VALUES (" & PerSocNum & ",Date()," & PerSocActif & ") ;"

               End Select

            Case Else

         End Select

   End Select

   CmbPerSocSocNum.SetFocus

   ChangeBoutonMode Me, MODEDEPLACEMENT

   CmbPerSocSocNum.SetFocus

   Me.Parent.ChangeSFrmButtonMode Me

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
Public Sub CmdAjouter_Click()
   Dim TxtDateCreer As String

   On Error GoTo Err_CmdAjouter_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

         CmdImprimer.SetFocus

      Case Else

         CmbPerSocSocNum.SetFocus

   End Select

   ChangeBoutonMode Me, MODEAJOUT

   DoCmd.RunCommand acCmdRecordsGoToNew

   TxtDateCreer = TxtPerSocDateCre

   TxtPerSocDateCre = "01/01/2003"

   TxtPerSocDateCre = TxtDateCreer

   CmbPerSocSocNum.SetFocus

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
Public Sub CmdModifier_Click()
   Dim TxtDateCreer As String

   On Error GoTo Err_CmdModifier_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocSocNum.SetFocus

         ChangeBoutonMode Me, MODEMODIF

         TxtDateCreer = TxtPerSocDateCre

         TxtPerSocDateCre = "01/01/2003"

         TxtPerSocDateCre = TxtDateCreer

         CmbPerSocSocNum.SetFocus

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
Public Sub CmdSupprimer_Click()

   On Error GoTo Err_CmdSupprimer_Click

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocSocNum.SetFocus

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

         On Error GoTo Err_CmdSupprimer_Click

         ChangeBoutonMode Me, MODEDEPLACEMENT

          Select Case Me.RecordsetClone.RecordCount
             Case vbEmpty

             Case Else

                CmbPerSocSocNum.SetFocus

          End Select

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
Public Sub CmdValider_Click()

   On Error GoTo Err_CmdValider_Click

   CmbPerSocSocNum.SetFocus

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
Public Sub CmdAnnuler_Click()

   On Error GoTo Err_CmdAnnuler_Click

   CmbPerSocSocNum.SetFocus

   Me.Undo

   ChangeBoutonMode Me, MODEDEPLACEMENT

   Select Case Me.RecordsetClone.RecordCount
      Case vbEmpty

      Case Else

         CmbPerSocSocNum.SetFocus

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

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficheTous_Click"

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
Private Sub CmbPerSocSocNum_DblClick(Cancel As Integer)
   Dim IntEtat As Integer

   On Error GoTo Err_CmdAfficherSocietes_Click

   DonneEtatFormulaire "FrmFicheSocietes", IntEtat

   Select Case IntEtat
      Case VUEFORMULAIRE

         Forms!FrmFicheSocietes.SetFocus

         DoCmd.ApplyFilter , "SocNum=" & PerSocSocNum

         DoCmd.RunCommand acCmdWindowCascade

      Case Else

         DoCmd.OpenForm "FrmFicheSocietes", acNormal, , "SocNum=" & PerSocSocNum, acFormEdit

   End Select

Exit_CmdAfficherSocietes_Click:

   Exit Sub

Err_CmdAfficherSocietes_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdAfficherSocietes_Click"

   Resume Exit_CmdAfficherSocietes_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CbxPerSocActif_BeforeUpdate(Cancel As Integer)
   Dim IntNbActifs As Integer

   On Error GoTo Err_CbxPerSocActif_BeforeUpdate

   Select Case PerSocActif
      Case True

         Select Case PerSocPerSocTypCode
            Case TYPEAMBASSADRICE, TYPEVENDEUR

               IntNbActifs = DCount("PerSocNum", "TBLPERSONNELSSOCIETES", "PerSocNum<>" & PerSocNum & " AND PerSocPerSocTypCode In ('AMB','VND') AND PerSocActif=True")

               Select Case IntNbActifs
                  Case 0

                  Case Else

                     Cancel = True

                     MsgBox "Attention !! un autre profil est actif !! Desactiver le !!", , Me.Caption

                     GoTo Exit_CbxPerSocActif_BeforeUpdate

               End Select

            Case Else

         End Select

      Case False

         Select Case PerSocAdherent
            Case True

               Cancel = True

               MsgBox "Attention !! le profil est adherent !! Desadherer le !!", , Me.Caption

               GoTo Exit_CbxPerSocActif_BeforeUpdate

            Case False

         End Select

   End Select

Exit_CbxPerSocActif_BeforeUpdate:

   Exit Sub

Err_CbxPerSocActif_BeforeUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CbxPerSocActif_BeforeUpdate"

   Resume Exit_CbxPerSocActif_BeforeUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CbxPerSocAdherent_BeforeUpdate(Cancel As Integer)
   Dim IntNbAdherents As Integer

   On Error GoTo Err_CbxPerSocAdherent_BeforeUpdate

   Select Case PerSocAdherent
      Case True

         Select Case PerSocActif
            Case True

               IntNbAdherents = DCount("PerSocNum", "TBLPERSONNELSSOCIETES", "PerSocNum<>" & PerSocNum & " AND PerSocPerSocTypCode In ('AMB','VND') AND PerSocAdherent=True")

               Select Case IntNbAdherents
                  Case 0

                  Case Else

                     Cancel = True

                     MsgBox "Attention !! un autre profil est adherent !! Desadherer le !!", , Me.Caption

                     GoTo Exit_CbxPerSocAdherent_BeforeUpdate

               End Select

            Case False

               Cancel = True

               MsgBox "Attention !! le profil n'est pas actif !!", , Me.Caption

               GoTo Exit_CbxPerSocAdherent_BeforeUpdate

         End Select

      Case False

   End Select

Exit_CbxPerSocAdherent_BeforeUpdate:

   Exit Sub

Err_CbxPerSocAdherent_BeforeUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CbxPerSocAdherent_BeforeUpdate"

   Resume Exit_CbxPerSocAdherent_BeforeUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocSocNum_GotFocus()

   On Error GoTo Err_CmbPerSocSocNum_GotFocus

'   CmbPerSocSocNum.Dropdown

Exit_CmbPerSocSocNum_GotFocus:

   Exit Sub

Err_CmbPerSocSocNum_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocSocNum_GotFocus"

   Resume Exit_CmbPerSocSocNum_GotFocus
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbPerSocPerSocTypCode_GotFocus()

   On Error GoTo Err_CmbPerSocPerSocTypCode_GotFocus

 ' CmbPerSocPerSocTypCode.Dropdown

Exit_CmbPerSocPerSocTypCode_GotFocus:

   Exit Sub

Err_CmbPerSocPerSocTypCode_GotFocus:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocPerSocTypCode_GotFocus"

   Resume Exit_CmbPerSocPerSocTypCode_GotFocus
End Sub
