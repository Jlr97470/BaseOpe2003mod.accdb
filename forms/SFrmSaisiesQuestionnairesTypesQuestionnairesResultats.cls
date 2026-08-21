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
'***    Declaration De Procedure Privée                                                        ***
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
   Dim CtlActive As Control
   Dim VarValeurTableau As Variant
   Dim VarValeur As Variant
   Dim StrValeurListe As String
   Dim LngValeurCompteur As Long
   Dim IntValeurMin As Integer
   Dim IntValeurMax As Integer

   On Error GoTo Err_Rafraichir

   Select Case (Me.RecordsetClone.RecordCount = 0) And (Me.NewRecord = False)
      Case True

      Case False

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE

               On Error Resume Next

               Set CtlActive = ActiveControl

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

                     Set CtlActive = CmbPerSocQstResValeur

              End Select

            Case Else

         End Select

         On Error GoTo Err_Rafraichir

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               StrValeurListe = "'Resultats Valeur';"

               Select Case PerSocQstTypIDTypNum
                  Case vbBoolean

                     StrValeurListe = StrValeurListe & "'Vrai';'Faux'"

                     CmbPerSocQstResValeur.InputMask = vbNullString

                     CmbPerSocQstResValeur.Format = vbNullString

                     CmbPerSocQstResValeur.RowSource = StrValeurListe

                  Case vbByte, vbInteger

                     Select Case PerSocQstTypIDTypNum
                        Case vbByte

                           IntValeurMin = 0

                           IntValeurMax = 255

                        Case vbInteger

                           IntValeurMin = -32768

                           IntValeurMax = 32767

                     End Select

                     Select Case Nz(PerSocQstTypIDMax, IntValeurMax) - Nz(PerSocQstTypIDMin, IntValeurMin)
                        Case Is < 50

                           For LngValeurCompteur = Nz(PerSocQstTypIDMin, IntValeurMin) To Nz(PerSocQstTypIDMax, IntValeurMax)

                              StrValeurListe = StrValeurListe & LngValeurCompteur & ";"

                           Next

                           CmbPerSocQstResValeur.RowSource = StrValeurListe

                        Case Else

                           CmbPerSocQstResValeur.RowSource = vbNullString

                     End Select

                     StrValeurListe = Left(StrValeurListe, Len(StrValeurListe) - 1)

                     CmbPerSocQstResValeur.InputMask = vbNullString

                     CmbPerSocQstResValeur.Format = vbNullString

                     CmbPerSocQstResValeur.RowSource = StrValeurListe

                  Case vbLong, vbSingle, vbDouble, vbCurrency

                     CmbPerSocQstResValeur.InputMask = vbNullString

                     CmbPerSocQstResValeur.Format = vbNullString

                     CmbPerSocQstResValeur.RowSource = vbNullString

                  Case vbDate

                     CmbPerSocQstResValeur.InputMask = "99/99/0000;;_"

                     CmbPerSocQstResValeur.Format = "Date, abrégé"

                     CmbPerSocQstResValeur.RowSource = vbNullString

                  Case vbString

                     Select Case IsNull(PerSocQstTypIDMin)
                        Case True

                           CmbPerSocQstResValeur.InputMask = Nz(PerSocQstTypIDMax, vbNullString)

                           CmbPerSocQstResValeur.Format = vbNullString

                           CmbPerSocQstResValeur.RowSource = vbNullString

                        Case False

                           VarValeurTableau = Split(PerSocQstTypIDMin, ";")

                           For Each VarValeur In VarValeurTableau

                              StrValeurListe = StrValeurListe & VarValeur & ";"

                           Next

                           StrValeurListe = Left(StrValeurListe, Len(StrValeurListe) - 1)

                           CmbPerSocQstResValeur.InputMask = Nz(PerSocQstTypIDMax, vbNullString)

                           CmbPerSocQstResValeur.Format = vbNullString

                           CmbPerSocQstResValeur.RowSource = StrValeurListe

                     End Select

               End Select

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
'***    Declaration d'Evenement Privée                                                        ***
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
Private Sub Form_BeforeUpdate(Cancel As Integer)
   Dim VarValeur As Variant
   Dim VarMax As Variant
   Dim VarMin As Variant

   On Error GoTo Err_Form_BeforeUpdate

   Select Case PerSocQstTypIDObligatoire = False And IsEmpty(CmbPerSocQstResValeur)
      Case True

      Case False

         On Error Resume Next

         Select Case PerSocQstTypIDTypNum
            Case vbBoolean

               Select Case CmbPerSocQstResValeur
                  Case "Vrai"

                  Case "Faux"

                  Case Else

                     Err.Raise 13

               End Select

            Case vbByte

               VarValeur = CByte(CmbPerSocQstResValeur)

               Select Case IsNull(PerSocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CByte(PerSocQstTypIDMax)

               End Select

               Select Case IsNull(PerSocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CByte(PerSocQstTypIDMin)

               End Select

            Case vbInteger

               VarValeur = CInt(CmbPerSocQstResValeur)

               Select Case IsNull(PerSocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CInt(PerSocQstTypIDMax)

               End Select

               Select Case IsNull(PerSocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CInt(PerSocQstTypIDMin)

               End Select

            Case vbLong

               VarValeur = CLng(CmbPerSocQstResValeur)

               Select Case IsNull(PerSocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CLng(PerSocQstTypIDMax)

               End Select

               Select Case IsNull(PerSocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CLng(PerSocQstTypIDMin)

               End Select

            Case vbSingle

               VarValeur = CSng(CmbPerSocQstResValeur)

               Select Case IsNull(PerSocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CSng(PerSocQstTypIDMax)

               End Select

               Select Case IsNull(PerSocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CSng(PerSocQstTypIDMin)

               End Select

            Case vbDouble

               VarValeur = CDbl(CmbPerSocQstResValeur)

               Select Case IsNull(PerSocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CDbl(PerSocQstTypIDMax)

               End Select

               Select Case IsNull(PerSocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CDbl(PerSocQstTypIDMin)

               End Select

            Case vbCurrency

               VarValeur = CCur(CmbPerSocQstResValeur)

               Select Case IsNull(PerSocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CCur(PerSocQstTypIDMax)

               End Select

               Select Case IsNull(PerSocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CCur(PerSocQstTypIDMin)

               End Select

            Case vbDate

               VarValeur = CDate(CmbPerSocQstResValeur)

               Select Case IsNull(PerSocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CDate(PerSocQstTypIDMax)

               End Select

               Select Case IsNull(PerSocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CDate(PerSocQstTypIDMin)

               End Select

            Case vbString

               Select Case Not IsNull(PerSocQstTypIDMin) And IsNull(PerSocQstTypIDMax)
                  Case True

                     VarValeur = UCase(CStr(CmbPerSocQstResValeur))

                  Case False

                     VarValeur = CStr(CmbPerSocQstResValeur)

               End Select

               Select Case VarValeur
                  Case vbNullString

                     PerSocQstResValeur = Null

                  Case Else

                     PerSocQstResValeur = VarValeur

               End Select

               Select Case IsNull(PerSocQstTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CStr(PerSocQstTypIDMax)

               End Select

               Select Case IsNull(PerSocQstTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CStr(PerSocQstTypIDMin)

               End Select

         End Select

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               Cancel = True

               MsgBox "Le Champ Questionnaire Resultat Valeur Est Invalide !", , Me.Caption

               CmbPerSocQstResValeur.SetFocus

               GoTo Exit_Form_BeforeUpdate

         End Select

         On Error GoTo Err_Form_BeforeUpdate

         Select Case PerSocQstTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate

               Select Case IsNull(VarMin)
                  Case True

                  Case False

                     Select Case VarMin <= VarValeur
                        Case True

                        Case False

                           Cancel = True

                           MsgBox "Le Champ Questionnaire Resultat Valeur Doit Etre Superieur Ou Egal A " & VarMin & " !", , Me.Caption

                           CmbPerSocQstResValeur.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

               Select Case IsNull(VarMax)
                  Case True

                  Case False

                     Select Case VarMax >= VarValeur
                        Case True

                        Case False

                           Cancel = True

                           MsgBox "Le Champ Questionnaire Resultat Valeur Doit Etre Inferieur Ou Egal A " & VarMax & " !", , Me.Caption

                           CmbPerSocQstResValeur.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

            Case vbString

               Select Case IsNull(VarMin)
                  Case True

                  Case False

                     Select Case CmbPerSocQstResValeur.ListIndex
                        Case Is >= 0

                        Case Else

                           Cancel = True

                           MsgBox "Le Champ Questionnaire Resultat Valeur Est Invalide !", , Me.Caption

                           CmbPerSocQstResValeur.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

         End Select

   End Select

   PerSocQstResPerSocNum = Me.Parent.CmbPerSocQstPerSocNum

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

   On Error GoTo Err_Form_AfterUpdate

   CmbPerSocQstResValeur.SetFocus

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
Private Sub CmbPerSocQstResValeur_DblClick(Cancel As Integer)

   On Error GoTo Err_CmbPerSocQstResValeur_DblClick

   DoCmd.OpenForm "FrmNotes", acNormal, , , acFormEdit, acDialog, "FrmSaisiesQuestionnairesTypes:SFrmSaisiesQuestionnairesTypesQuestionnairesResultats:CmbPerSocQstResValeur:"

Exit_CmbPerSocQstResValeur_DblClick:

   Exit Sub

Err_CmbPerSocQstResValeur_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocQstResValeur_DblClick"

   Resume Exit_CmbPerSocQstResValeur_DblClick
End Sub
