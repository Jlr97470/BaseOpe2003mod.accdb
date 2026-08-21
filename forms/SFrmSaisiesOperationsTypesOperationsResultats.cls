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

                     Set CtlActive = CmbPerSocOpeResValeur

              End Select

            Case Else

         End Select

         On Error GoTo Err_Rafraichir

         Select Case BytModeRafraichir
            Case RAFRAICHIRVISIBLE, RAFRAICHIRDEPLACEMENT

               StrValeurListe = "'Resultats Valeur';"

               Select Case PerSocOpeTypIDTypNum
                  Case vbBoolean

                     StrValeurListe = StrValeurListe & "'Vrai';'Faux'"

                     CmbPerSocOpeResValeur.InputMask = vbNullString

                     CmbPerSocOpeResValeur.Format = vbNullString

                     CmbPerSocOpeResValeur.RowSource = StrValeurListe

                  Case vbByte, vbInteger

                     Select Case PerSocOpeTypIDTypNum
                        Case vbByte

                           IntValeurMin = 0

                           IntValeurMax = 255

                        Case vbInteger

                           IntValeurMin = -32768

                           IntValeurMax = 32767

                     End Select

                     Select Case Nz(PerSocOpeTypIDMax, IntValeurMax) - Nz(PerSocOpeTypIDMin, IntValeurMin)
                        Case Is < 50

                           For LngValeurCompteur = Nz(PerSocOpeTypIDMin, IntValeurMin) To Nz(PerSocOpeTypIDMax, IntValeurMax)

                              StrValeurListe = StrValeurListe & LngValeurCompteur & ";"

                           Next

                           CmbPerSocOpeResValeur.RowSource = StrValeurListe

                        Case Else

                           CmbPerSocOpeResValeur.RowSource = vbNullString

                     End Select

                     StrValeurListe = Left(StrValeurListe, Len(StrValeurListe) - 1)

                     CmbPerSocOpeResValeur.InputMask = vbNullString

                     CmbPerSocOpeResValeur.Format = vbNullString

                     CmbPerSocOpeResValeur.RowSource = StrValeurListe

                  Case vbLong, vbSingle, vbDouble, vbCurrency

                     CmbPerSocOpeResValeur.InputMask = vbNullString

                     CmbPerSocOpeResValeur.Format = vbNullString

                     CmbPerSocOpeResValeur.RowSource = vbNullString

                  Case vbDate

                     CmbPerSocOpeResValeur.InputMask = "99/99/0000;;_"

                     CmbPerSocOpeResValeur.Format = "Date, abrégé"

                     CmbPerSocOpeResValeur.RowSource = vbNullString

                  Case vbString

                     Select Case IsNull(PerSocOpeTypIDMin)
                        Case True

                           CmbPerSocOpeResValeur.InputMask = Nz(PerSocOpeTypIDMax, vbNullString)

                           CmbPerSocOpeResValeur.Format = vbNullString

                           CmbPerSocOpeResValeur.RowSource = vbNullString

                        Case False

                           VarValeurTableau = Split(PerSocOpeTypIDMin, ";")

                           For Each VarValeur In VarValeurTableau

                              StrValeurListe = StrValeurListe & VarValeur & ";"

                           Next

                           StrValeurListe = Left(StrValeurListe, Len(StrValeurListe) - 1)

                           CmbPerSocOpeResValeur.InputMask = Nz(PerSocOpeTypIDMax, vbNullString)

                           CmbPerSocOpeResValeur.Format = vbNullString

                           CmbPerSocOpeResValeur.RowSource = StrValeurListe

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

   Select Case PerSocOpeTypIDObligatoire = False And IsNull(PerSocOpeResValeur)
      Case True

      Case False

         On Error Resume Next

         Select Case PerSocOpeTypIDTypNum
            Case vbBoolean

               Select Case PerSocOpeResValeur
                  Case "Vrai"

                  Case "Faux"

                  Case Else

                     Err.Raise 13

               End Select

            Case vbByte

               VarValeur = CByte(PerSocOpeResValeur)

               Select Case IsNull(PerSocOpeTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CByte(PerSocOpeTypIDMax)

               End Select

               Select Case IsNull(PerSocOpeTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CByte(PerSocOpeTypIDMin)

               End Select

            Case vbInteger

               VarValeur = CInt(PerSocOpeResValeur)

               Select Case IsNull(PerSocOpeTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CInt(PerSocOpeTypIDMax)

               End Select

               Select Case IsNull(PerSocOpeTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CInt(PerSocOpeTypIDMin)

               End Select

            Case vbLong

               VarValeur = CLng(PerSocOpeResValeur)

               Select Case IsNull(PerSocOpeTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CLng(PerSocOpeTypIDMax)

               End Select

               Select Case IsNull(PerSocOpeTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CLng(PerSocOpeTypIDMin)

               End Select

            Case vbSingle

               VarValeur = CSng(PerSocOpeResValeur)

               Select Case IsNull(PerSocOpeTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CSng(PerSocOpeTypIDMax)

               End Select

               Select Case IsNull(PerSocOpeTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CSng(PerSocOpeTypIDMin)

               End Select

            Case vbDouble

               VarValeur = CDbl(PerSocOpeResValeur)

               Select Case IsNull(PerSocOpeTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CDbl(PerSocOpeTypIDMax)

               End Select

               Select Case IsNull(PerSocOpeTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CDbl(PerSocOpeTypIDMin)

               End Select

            Case vbCurrency

               VarValeur = CCur(PerSocOpeResValeur)

               Select Case IsNull(PerSocOpeTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CCur(PerSocOpeTypIDMax)

               End Select

               Select Case IsNull(PerSocOpeTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CCur(PerSocOpeTypIDMin)

               End Select

            Case vbDate

               VarValeur = CDate(PerSocOpeResValeur)

               Select Case IsNull(PerSocOpeTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CDate(PerSocOpeTypIDMax)

               End Select

               Select Case IsNull(PerSocOpeTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CDate(PerSocOpeTypIDMin)

               End Select

            Case vbString

               Select Case Not IsNull(PerSocOpeTypIDMin) And IsNull(PerSocOpeTypIDMax)
                  Case True

                     VarValeur = UCase(CStr(PerSocOpeResValeur))

                  Case False

                     VarValeur = CStr(PerSocOpeResValeur)

               End Select

               Select Case VarValeur
                  Case vbNullString

                     PerSocOpeResValeur = Null

                  Case Else

                     PerSocOpeResValeur = VarValeur

               End Select

               Select Case IsNull(PerSocOpeTypIDMax)
                  Case True

                     VarMax = Null

                  Case False

                     VarMax = CStr(PerSocOpeTypIDMax)

               End Select

               Select Case IsNull(PerSocOpeTypIDMin)
                  Case True

                     VarMin = Null

                  Case False

                     VarMin = CStr(PerSocOpeTypIDMin)

               End Select

         End Select

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               Cancel = True

               MsgBox "Le Champ Operations Resultats Valeur Est Invalide !", , Me.Caption

               CmbPerSocOpeResValeur.SetFocus

               GoTo Exit_Form_BeforeUpdate

         End Select

         On Error GoTo Err_Form_BeforeUpdate

         Select Case PerSocOpeTypIDTypNum
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate

               Select Case IsNull(VarMin)
                  Case True

                  Case False

                     Select Case VarMin <= VarValeur
                        Case True

                        Case False

                           Cancel = True

                           MsgBox "Le Champ Operations Resultats Valeur Doit Etre Superieur Ou Egal A " & VarMin & " !", , Me.Caption

                           CmbPerSocOpeResValeur.SetFocus

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

                           MsgBox "Le Champ Operations Resultats Valeurs Doit Etre Inferieur Ou Egal A " & VarMax & " !", , Me.Caption

                           CmbPerSocOpeResValeur.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

            Case vbString

               Select Case IsNull(VarMin)
                  Case True

                  Case False

                     Select Case CmbPerSocOpeResValeur.ListIndex
                        Case Is >= 0

                        Case Else

                           Cancel = True

                           MsgBox "Le Champ Operation Resultat Valeur Est Invalide !", , Me.Caption

                           CmbPerSocOpeResValeur.SetFocus

                           GoTo Exit_Form_BeforeUpdate

                     End Select

               End Select

         End Select

   End Select

   PerSocOpeResPerSocNum = Me.Parent.CmbPerSocOpePerSocNum

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

   CmbPerSocOpeResValeur.SetFocus

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
Private Sub CmbPerSocOpeResValeur_DblClick(Cancel As Integer)

   On Error GoTo Err_CmbPerSocOpeResValeur_DblClick

   DoCmd.OpenForm "FrmNotes", acNormal, , , acFormEdit, acDialog, "FrmSaisiesOperationsTypes:SFrmSaisiesOperationsTypesOperationsResultats:CmbPerSocOpeResValeur:"

Exit_CmbPerSocOpeResValeur_DblClick:

   Exit Sub

Err_CmbPerSocOpeResValeur_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbPerSocOpeResValeur_DblClick"

   Resume Exit_CmbPerSocOpeResValeur_DblClick
End Sub
