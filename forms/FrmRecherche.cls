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
'***    Declaration De Variable Priver                                                          ***
'******************************************************************************

'******************************************************************************
'***    PROCEDURE:                                                                                    ***
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

         EtiValeurRecherche.Caption = "Valeur Recherchée : " & TbsRecherche.Pages(TbsRecherche.Value).Name

         Select Case TbsRecherche.Value + 1
            Case 1

               EtiTypeRecherche1.Caption = "Recherche Alpha-Numerique"

               EtiTypeRecherche2.Caption = "Recherche Numerique"

               EtiTypeRecherche3.Caption = "Recherche Date"

               EtiTypeRecherche4.Caption = "Recherche Code STA"

               EtiTypeRecherche5.Caption = "Recherche Raison Sociale"

               EtiTypeRecherche6.Caption = "Recherche Code Postal"

               EtiTypeRecherche7.Caption = "Recherche Telephone"

               EtiTypeRecherche3.Visible = True

               EtiTypeRecherche4.Visible = True

               EtiTypeRecherche5.Visible = True

               EtiTypeRecherche6.Visible = True

               EtiTypeRecherche7.Visible = True

               CbxTypeRecherche3.Visible = True

               CbxTypeRecherche4.Visible = True

               CbxTypeRecherche5.Visible = True

               CbxTypeRecherche6.Visible = True

               CbxTypeRecherche7.Visible = True

            Case 2

               EtiTypeRecherche1.Caption = "Recherche Alpha-Numerique"

               EtiTypeRecherche2.Caption = "Recherche Numerique"

               EtiTypeRecherche3.Caption = "Recherche Date"

               EtiTypeRecherche4.Caption = "Recherche Code Postal"

               EtiTypeRecherche5.Caption = "Recherche Telephone"

               EtiTypeRecherche6.Caption = vbNullString

               EtiTypeRecherche7.Caption = vbNullString

               EtiTypeRecherche3.Visible = True

               EtiTypeRecherche4.Visible = True

               EtiTypeRecherche5.Visible = True

               EtiTypeRecherche6.Visible = False

               EtiTypeRecherche7.Visible = False

               CbxTypeRecherche3.Visible = True

               CbxTypeRecherche4.Visible = True

               CbxTypeRecherche5.Visible = True

               CbxTypeRecherche6.Visible = False

               CbxTypeRecherche7.Visible = False
            Case 3

               EtiTypeRecherche1.Caption = "Recherche Alpha-Numerique"

               EtiTypeRecherche2.Caption = "Recherche Numerique"

               EtiTypeRecherche3.Caption = "Recherche Date"

               EtiTypeRecherche4.Caption = vbNullString

               EtiTypeRecherche5.Caption = vbNullString

               EtiTypeRecherche6.Caption = vbNullString

               EtiTypeRecherche7.Caption = vbNullString

               EtiTypeRecherche3.Visible = True

               EtiTypeRecherche4.Visible = False

               EtiTypeRecherche5.Visible = False

               EtiTypeRecherche6.Visible = False

               EtiTypeRecherche7.Visible = False

               CbxTypeRecherche3.Visible = True

               CbxTypeRecherche4.Visible = False

               CbxTypeRecherche5.Visible = False

               CbxTypeRecherche6.Visible = False

               CbxTypeRecherche7.Visible = False

            Case 4

               EtiTypeRecherche1.Caption = "Recherche Alpha-Numerique"

               EtiTypeRecherche2.Caption = "Recherche Numerique"

               EtiTypeRecherche3.Caption = "Recherche Date"

               EtiTypeRecherche4.Caption = vbNullString

               EtiTypeRecherche5.Caption = vbNullString

               EtiTypeRecherche6.Caption = vbNullString

               EtiTypeRecherche7.Caption = vbNullString

               EtiTypeRecherche3.Visible = True

               EtiTypeRecherche4.Visible = False

               EtiTypeRecherche5.Visible = False

               EtiTypeRecherche6.Visible = False

               EtiTypeRecherche7.Visible = False

               CbxTypeRecherche3.Visible = True

               CbxTypeRecherche4.Visible = False

               CbxTypeRecherche5.Visible = False

               CbxTypeRecherche6.Visible = False

               CbxTypeRecherche7.Visible = False

         End Select

      Case Else

   End Select

Exit_Rafraichir:

   Exit Sub

Err_Rafraichir:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Rafraichir"

   Resume Exit_Rafraichir
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
Private Sub Form_Activate()

   On Error GoTo Err_Form_Activate

   Rafraichir (RAFRAICHIRVISIBLE)

Exit_Form_Activate:

   Exit Sub

Err_Form_Activate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Activate"

   Resume Err_Form_Activate
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
   Dim IntEtat As Integer
   Dim RsSocietesVilles As DAO.Recordset
   Dim RsPersonnelsVilles As DAO.Recordset
   Dim FldChamp As DAO.Field

   On Error GoTo Err_CmdValider_Click

   Select Case TbsRecherche.SelectedItem.Index
      Case 1

         GStrSocietesFiltre = vbNullString

         Set RsSocietesVilles = CurrentDb.OpenRecordset("SelSocietesVilles")

         GStrSocietesFiltre = "SocNum In (SELECT SelSocietesVilles.SocNum FROM SelSocietesVilles WHERE "

         Select Case CadTypeRecherche
            Case 1

               For Each FldChamp In RsSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText, dbMemo

                        GStrSocietesFiltre = GStrSocietesFiltre & "SelSocietesVilles." & FldChamp.Name & " LIKE '*" & TxtValeurRecherche & "*' OR "

                     Case Else

                  End Select

               Next

            Case 2

               For Each FldChamp In RsSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbByte, dbCurrency, dbDouble, dbFloat, dbLong, dbSingle

                        GStrSocietesFiltre = GStrSocietesFiltre & "SelSocietesVilles." & FldChamp.Name & "=" & RemplaceChr(TxtValeurRecherche, ",", ".") & " OR "

                     Case Else

                  End Select

               Next

            Case 3

               For Each FldChamp In RsSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbDate, dbTime, dbTimeStamp

                        GStrSocietesFiltre = GStrSocietesFiltre & "SelSocietesVilles." & FldChamp.Name & "=CDate('" & TxtValeurRecherche & "') OR "

                     Case Else

                  End Select

               Next

            Case 4

               GStrSocietesFiltre = GStrSocietesFiltre & "SelSocietesVilles.SocSTA LIKE '*" & TxtValeurRecherche & "*' OR "

            Case 5

               GStrSocietesFiltre = GStrSocietesFiltre & "SelSocietesVilles.SocRaison LIKE '*" & TxtValeurRecherche & "*' OR "

            Case 6

               For Each FldChamp In RsSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText

                        Select Case FldChamp.Size
                           Case 5

                              GStrSocietesFiltre = GStrSocietesFiltre & "SelSocietesVilles." & FldChamp.Name & " LIKE '*" & TxtValeurRecherche & "*' OR "

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

            Case 7

               For Each FldChamp In RsSocietesVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText

                        Select Case FldChamp.Size
                           Case 10

                              GStrSocietesFiltre = GStrSocietesFiltre & "SelSocietesVilles." & FldChamp.Name & " LIKE '*" & TxtValeurRecherche & "*' OR "

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

         End Select

         GStrSocietesFiltre = Mid(GStrSocietesFiltre, 1, Len(GStrSocietesFiltre) - 4) & ")"

         RsSocietesVilles.Filter = GStrSocietesFiltre

         Set RsSocietesVilles = RsSocietesVilles.OpenRecordset

         Select Case RsSocietesVilles.EOF
            Case True

            Case False

               RsSocietesVilles.MoveLast

         End Select

         Select Case RsSocietesVilles.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdValider_Click

            Case 1

               DonneEtatFormulaire "FrmFicheSocietes", IntEtat

               Select Case IntEtat
                  Case VUEFORMULAIRE

                     Forms!FrmFicheSocietes.SetFocus

                     DoCmd.ApplyFilter , GStrSocietesFiltre

                     DoCmd.RunCommand acCmdWindowCascade

                  Case Else

                     DoCmd.OpenForm "FrmFicheSocietes", acNormal, , GStrSocietesFiltre, acFormEdit

               End Select

            Case Else

               DoCmd.Close acForm, "FrmLstSocietes", acSavePrompt

               DoCmd.OpenForm "FrmLstSocietes", acNormal, , , acFormEdit, acWindowNormal, GStrSocietesFiltre

         End Select

         RsSocietesVilles.Close

      Case 2

         GStrPersonnelsFiltre = vbNullString

         Set RsPersonnelsVilles = CurrentDb.OpenRecordset("SelPersonnelsVilles")

         GStrPersonnelsFiltre = "ConNum In (SELECT SelPersonnelsVilles.ConNum FROM SelPersonnelsVilles WHERE "

         Select Case CadTypeRecherche
            Case 1

               For Each FldChamp In RsPersonnelsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText, dbMemo

                        GStrPersonnelsFiltre = GStrPersonnelsFiltre & "SelPersonnelsVilles." & FldChamp.Name & " LIKE '*" & TxtValeurRecherche & "*' OR "

                     Case Else

                  End Select

               Next

            Case 2

               For Each FldChamp In RsPersonnelsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbByte, dbCurrency, dbDouble, dbFloat, dbLong, dbSingle

                        GStrPersonnelsFiltre = GStrPersonnelsFiltre & "SelPersonnelsVilles." & FldChamp.Name & "=" & RemplaceChr(TxtValeurRecherche, ",", ".") & " OR "

                     Case Else

                  End Select

               Next

            Case 3

               For Each FldChamp In RsPersonnelsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbDate, dbTime, dbTimeStamp

                        GStrPersonnelsFiltre = GStrPersonnelsFiltre & "SelPersonnelsVilles." & FldChamp.Name & "=CDate('" & TxtValeurRecherche & "') OR "

                     Case Else

                  End Select

               Next

            Case 4

               For Each FldChamp In RsPersonnelsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText

                        Select Case FldChamp.Size
                           Case 5

                              GStrPersonnelsFiltre = GStrPersonnelsFiltre & "SelPersonnelsVilles." & FldChamp.Name & " LIKE '*" & TxtValeurRecherche & "*' OR "

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

            Case 5

               For Each FldChamp In RsPersonnelsVilles.Fields

                  Select Case FldChamp.Type
                     Case dbText

                        Select Case FldChamp.Size
                           Case 10

                              GStrPersonnelsFiltre = GStrPersonnelsFiltre & "SelPersonnelsVilles." & FldChamp.Name & " LIKE '*" & TxtValeurRecherche & "*' OR "

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

         End Select

         GStrPersonnelsFiltre = Mid(GStrPersonnelsFiltre, 1, Len(GStrPersonnelsFiltre) - 4) & ")"

         RsPersonnelsVilles.Filter = GStrPersonnelsFiltre

         Set RsPersonnelsVilles = RsPersonnelsVilles.OpenRecordset

         Select Case RsPersonnelsVilles.EOF
            Case True

            Case False

               RsPersonnelsVilles.MoveLast

         End Select

         Select Case RsPersonnelsVilles.RecordCount
            Case 0

               MsgBox "Aucune fiche pour cette selection !!", , Me.Caption

               GoTo Exit_CmdValider_Click

            Case 1

               DonneEtatFormulaire "FrmFichePersonnels", IntEtat

               Select Case IntEtat
                  Case VUEFORMULAIRE

                     Forms!FrmFichePersonnels.SetFocus

                     DoCmd.ApplyFilter , GStrPersonnelsFiltre

                     DoCmd.RunCommand acCmdWindowCascade

                  Case Else

                     DoCmd.OpenForm "FrmFichePersonnels", acNormal, , GStrPersonnelsFiltre, acFormEdit

               End Select

            Case Else

               DoCmd.Close acForm, "FrmLstPersonnels", acSavePrompt

               DoCmd.OpenForm "FrmLstPersonnels", acNormal, , , acFormEdit, acWindowNormal, GStrPersonnelsFiltre

         End Select

         RsPersonnelsVilles.Close

   End Select

   DoCmd.Close acForm, Me.Name, acSavePrompt

Exit_CmdValider_Click:

   Set RsSocietesVilles = Nothing

   Set RsPersonnelsVilles = Nothing

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
Private Sub TbsRecherche_Change()

   On Error GoTo Err_TbsRecherche_Change

   Rafraichir RAFRAICHIRCHOIX

Exit_TbsRecherche_Change:

   Exit Sub

Err_TbsRecherche_Change:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TbsRecherche_Change"

   Resume Exit_TbsRecherche_Change
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub TxtValeurRecherche_Change()

   On Error GoTo Err_TxtValeurRecherche

   Select Case Len(TxtValeurRecherche.Text)
      Case Is > vbEmpty

         CmdValider.Enabled = True
      Case Else

         CmdValider.Enabled = False
   End Select

Exit_TxtValeurRecherche:

   Exit Sub

Err_TxtValeurRecherche:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "TxtValeurRecherche_Change"

   Resume Exit_TxtValeurRecherche
End Sub
