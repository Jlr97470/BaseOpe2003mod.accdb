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
Private Sub CreeListeTableRequete()
   Dim TdfTable As DAO.TableDef
   Dim QdfRequete As DAO.QueryDef
   Dim IntTableRequeteIndex As Integer

   On Error GoTo Err_CreeListeTableRequete

   Select Case CmbTableRequete.ListCount
      Case 0

      Case Else

         For IntTableRequeteIndex = 0 To CmbTableRequete.ListCount - 1

            CmbTableRequete.RemoveItem IntTableRequeteIndex

         Next

   End Select

   For Each TdfTable In CurrentDb.TableDefs

      Select Case (TdfTable.Attributes And dbSystemObject) <> 0 Or (TdfTable.Attributes And dbHiddenObject) <> 0
         Case True

         Case False

            On Error Resume Next

            CmbTableRequete.AddItem TdfTable.Name & ";" & TdfTable.Properties("Description") & ";" & "TABLE"

            Select Case Err.Number
               Case vbEmpty

               Case Else

                  Err.Clear

                  CmbTableRequete.AddItem TdfTable.Name & ";" & ";" & "TABLE"

            End Select

            On Error GoTo Err_CreeListeTableRequete

      End Select

   Next

   For Each QdfRequete In CurrentDb.QueryDefs

      Select Case Left(QdfRequete.Name, 4) = "~sq_" Or QdfRequete.Type = dbQProcedure Or QdfRequete.Type = dbQAction Or QdfRequete.Type = dbQAppend Or QdfRequete.Type = dbQDelete Or QdfRequete.Type = dbQMakeTable Or QdfRequete.Type = dbQUpdate
         Case True

         Case False

            On Error Resume Next

            CmbTableRequete.AddItem QdfRequete.Name & ";" & QdfRequete.Properties("Description") & ";" & "REQUETE"

            Select Case Err.Number
               Case vbEmpty

               Case Else

                  Err.Clear

                  CmbTableRequete.AddItem QdfRequete.Name & ";" & ";" & "REQUETE"

            End Select

            On Error GoTo Err_CreeListeTableRequete

      End Select

   Next

Exit_CreeListeTableRequete:

   Set TdfTable = Nothing

   Exit Sub

Err_CreeListeTableRequete:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeListeTableRequete"

   Resume Exit_CreeListeTableRequete
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
Private Sub CreeListeChampTableRequete()
   Dim TdsTable As DAO.TableDefs
   Dim TdfTable As DAO.TableDef
   Dim QdsRequete As DAO.QueryDefs
   Dim QdfRequete As DAO.QueryDef
   Dim FldChamp As DAO.Field

   On Error GoTo Err_CreeListeChampTableRequete

   LstChamp.RowSource = vbNullString

   Select Case CmbTableRequete.Column(0)
      Case vbNullString

      Case Else

         Select Case CmbTableRequete.Column(2)
            Case "TABLE"

               Set TdsTable = CurrentDb.TableDefs

               Set TdfTable = TdsTable(CmbTableRequete.Column(0))

               For Each FldChamp In TdfTable.Fields

                  On Error Resume Next

                  LstChamp.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description") & ";" & FldChamp.Properties("Type")

                  Select Case Err.Number
                     Case vbEmpty

                     Case Else

                        Err.Clear

                        LstChamp.AddItem FldChamp.Name & ";"

                  End Select

                  On Error GoTo Err_CreeListeChampTableRequete

               Next

            Case "REQUETE"

               Set QdsRequete = CurrentDb.QueryDefs

               Set QdfRequete = QdsRequete(CmbTableRequete.Column(0))

               For Each FldChamp In QdfRequete.Fields

                  On Error Resume Next

                  LstChamp.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description") & ";" & FldChamp.Properties("Type")

                  Select Case Err.Number
                     Case vbEmpty

                     Case Else

                        Err.Clear

                        LstChamp.AddItem FldChamp.Name & ";"

                  End Select

                  On Error GoTo Err_CreeListeChampTableRequete

               Next
      End Select

   End Select

Exit_CreeListeChampTableRequete:

   Exit Sub

Err_CreeListeChampTableRequete:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeListeChampTableRequete"

   Resume Exit_CreeListeChampTableRequete
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

   CmbTableRequete.RowSource = vbNullString

   LstChamp.RowSource = vbNullString

   LstChampSelectionner.RowSource = vbNullString

   CreeListeTableRequete

   CmbChampCondition1 = vbNullString

   CmbChampCondition2 = vbNullString

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
Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name, acSavePrompt

   On Error Resume Next

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
Private Sub CmdMinimize_Click()

   On Error GoTo Err_CmdMinimize_Click

   DoCmd.Minimize

Exit_CmdMinimize_Click:

   Exit Sub

Err_CmdMinimize_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdMinimize_Click"

   Resume Exit_CmdMinimize_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdExporteExcel_Click()
   Dim StrSQLRequeteTable As String
   Dim StrSQLRequeteTableChamp As String
   Dim StrSQLRequeteTableWhere As String
   Dim StrSQLRequeteTableOrder As String
   Dim StrCondition As String
   Dim StrValeur As String
   Dim StrOrdre() As String
   Dim IntChamp As Integer

   On Error GoTo Err_CmdExporteExcel_Click

   Select Case CmbTableRequete.ListIndex
      Case Is < 0

      Case Else

         Select Case LstChampSelectionner.ListCount
            Case 0

            Case Else

               ReDim StrOrdre(LstChampSelectionner.ListCount - 1)

               StrSQLRequeteTable = "SELECT "

               StrSQLRequeteTableChamp = vbNullString

               StrSQLRequeteTableWhere = vbNullString

               StrSQLRequeteTableOrder = vbNullString

               For IntChamp = 0 To LstChampSelectionner.ListCount - 1

                  StrSQLRequeteTableChamp = StrSQLRequeteTableChamp & LstChampSelectionner.Column(0, IntChamp) & ", "

                  Select Case (InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") <= 0 Or IsNull(InStr(1, LstChampSelectionner.Column(3, IntChamp), " ")))
                     Case True

                     Case False

                        StrCondition = vbNullString

                        StrValeur = vbNullString

                        Select Case Left(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") - 1)
                           Case "Comme"

                              StrCondition = "Like"

                              StrValeur = "'" & Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1) & "'"

                           Case "Pas_Comme"

                              StrCondition = "Not Like"

                              StrValeur = "'" & Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1) & "'"

                           Case Else

                              StrCondition = Left(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") - 1)

                              Select Case LstChampSelectionner.Column(2, IntChamp)
                                 Case dbBoolean

                                    Select Case Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1)
                                       Case "Vrai"

                                          StrValeur = "True"

                                       Case "Faux"

                                          StrValeur = "False"

                                    End Select

                                 Case dbDate, dbTime, dbTimeStamp

                                    StrValeur = "CDate('" & Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1) & "')"

                                 Case dbChar, dbText, dbMemo

                                    StrValeur = "'" & Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1) & "'"

                                 Case Else

                                    StrValeur = Mid(LstChampSelectionner.Column(3, IntChamp), InStr(1, LstChampSelectionner.Column(3, IntChamp), " ") + 1)

                              End Select

                        End Select

                        StrSQLRequeteTableWhere = StrSQLRequeteTableWhere & LstChampSelectionner.Column(0, IntChamp) & " " & StrCondition & " " & StrValeur & " AND "

                  End Select

                  Select Case (InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") <= 0) Or IsNull(InStr(1, LstChampSelectionner.Column(4, IntChamp), " "))
                     Case True

                     Case False

                        StrCondition = vbNullString

                        StrValeur = vbNullString

                        Select Case Left(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") - 1)
                           Case "Comme"

                              StrCondition = "Like"

                              StrValeur = "'" & Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1) & "'"

                           Case "Pas_Comme"

                              StrCondition = "Not Like"

                              StrValeur = "'" & Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1) & "'"

                           Case Else

                              StrCondition = Left(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") - 1)

                              Select Case LstChampSelectionner.Column(2, IntChamp)
                                 Case dbBoolean

                                    Select Case Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1)
                                       Case "Vrai"

                                          StrValeur = "True"

                                       Case "Faux"

                                          StrValeur = "False"

                                    End Select

                                 Case dbDate, dbTime, dbTimeStamp

                                    StrValeur = "CDate('" & Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1) & "')"

                                 Case dbChar, dbText, dbMemo

                                    StrValeur = "'" & Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1) & "'"

                                 Case Else

                                    StrValeur = Mid(LstChampSelectionner.Column(4, IntChamp), InStr(1, LstChampSelectionner.Column(4, IntChamp), " ") + 1)

                              End Select

                        End Select

                        StrSQLRequeteTableWhere = StrSQLRequeteTableWhere & LstChampSelectionner.Column(0, IntChamp) & " " & StrCondition & " " & StrValeur & " AND "

                  End Select

                  Select Case (InStr(1, LstChampSelectionner.Column(5, IntChamp), " ") <= 0) Or IsNull(InStr(1, LstChampSelectionner.Column(5, IntChamp), " "))
                     Case True

                     Case False

                        Select Case Mid(LstChampSelectionner.Column(5, IntChamp), InStr(1, LstChampSelectionner.Column(5, IntChamp), " ") + 1)
                           Case "Croisant"

                              StrOrdre(Left(LstChampSelectionner.Column(5, IntChamp), InStr(1, LstChampSelectionner.Column(5, IntChamp), " ") - 1)) = LstChampSelectionner.Column(0, IntChamp) & " ASC"

                           Case "Decroisant"

                              StrOrdre(Left(LstChampSelectionner.Column(5, IntChamp), InStr(1, LstChampSelectionner.Column(5, IntChamp), " ") - 1)) = LstChampSelectionner.Column(0, IntChamp) & " DESC"

                        End Select

                  End Select

               Next

               StrSQLRequeteTableChamp = Left(StrSQLRequeteTableChamp, Len(StrSQLRequeteTableChamp) - 2)

               StrSQLRequeteTable = StrSQLRequeteTable & " " & StrSQLRequeteTableChamp & " FROM " & CmbTableRequete.Column(0)

               Select Case StrSQLRequeteTableWhere
                  Case vbNullString

                  Case Else

                     StrSQLRequeteTableWhere = Left(StrSQLRequeteTableWhere, Len(StrSQLRequeteTableWhere) - 5)

                     StrSQLRequeteTable = StrSQLRequeteTable & " WHERE " & StrSQLRequeteTableWhere

               End Select

               For IntChamp = 0 To LstChampSelectionner.ListCount - 1

                  Select Case StrOrdre(IntChamp)
                     Case vbNullString

                     Case Else

                        StrSQLRequeteTableOrder = StrSQLRequeteTableOrder & StrOrdre(IntChamp) & ", "

                  End Select

               Next

               Select Case StrSQLRequeteTableOrder
                  Case vbNullString

                  Case Else

                     StrSQLRequeteTableOrder = Left(StrSQLRequeteTableOrder, Len(StrSQLRequeteTableOrder) - 2)

                     StrSQLRequeteTable = StrSQLRequeteTable & " ORDER BY " & StrSQLRequeteTableOrder

               End Select

               StrSQLRequeteTable = StrSQLRequeteTable & ";"

               ExporteExcelRequete (StrSQLRequeteTable)

         End Select

   End Select

Exit_CmdExporteExcel_Click:

   Exit Sub

Err_CmdExporteExcel_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdExporteExcel_Click"

   Resume Exit_CmdExporteExcel_Click
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
   Dim IntIndex As Integer
   Dim IntOrdre As Integer
   Dim TxtChamp(2) As String
   Dim TxtCondition1Valeur As String
   Dim TxtCondition2Valeur As String
   Dim TxtOrdreValeur As String

   On Error GoTo Err_CmdValider_Click

   Select Case LstChampSelectionner.ListIndex
      Case Is >= 0

         IntOrdre = 0

         TxtCondition1Valeur = vbNullString

         TxtCondition2Valeur = vbNullString

         TxtOrdreValeur = vbNullString

         Select Case InStr(1, LstChampSelectionner.Column(5), " ")
            Case Is > 0

               IntOrdre = CInt(Left(LstChampSelectionner.Column(5), InStr(1, LstChampSelectionner.Column(5), " ") - 1))

            Case Else

               For IntIndex = 0 To LstChampSelectionner.ListCount - 1

                  Select Case InStr(1, LstChampSelectionner.Column(5, IntIndex), " ")
                     Case Is > 0

                        Select Case CInt(Left(LstChampSelectionner.Column(5, IntIndex), InStr(1, LstChampSelectionner.Column(5, IntIndex), " ") - 1))
                           Case Is > 0

                              IntOrdre = CInt(Left(LstChampSelectionner.Column(5, IntIndex), InStr(1, LstChampSelectionner.Column(5, IntIndex), " ") - 1))

                           Case Else

                        End Select

                     Case Else

                  End Select

               Next

               IntOrdre = IntOrdre + 1

         End Select

         IntIndex = LstChampSelectionner.ListIndex

         TxtChamp(0) = LstChampSelectionner.Column(0)

         TxtChamp(1) = LstChampSelectionner.Column(1)

         TxtChamp(2) = LstChampSelectionner.Column(2)

         Select Case CmbChampCondition1.ListIndex
            Case Is < 0

            Case Else

               TxtCondition1Valeur = CmbChampCondition1.Column(0) & " " & CmbChampValeurCondition1

         End Select

         Select Case CmbChampCondition2.ListIndex
            Case Is < 0

            Case Else

               TxtCondition2Valeur = CmbChampCondition2.Column(0) & " " & CmbChampValeurCondition2

         End Select

         Select Case CmbChampOrdre.ListIndex
            Case Is < 0

            Case Else

               TxtOrdreValeur = IntOrdre & " " & CmbChampOrdre

         End Select

         LstChampSelectionner.RemoveItem IntIndex

         Select Case LstChampSelectionner.ListCount
            Case IntIndex

               LstChampSelectionner.AddItem TxtChamp(0) & ";" & TxtChamp(1) & ";" & TxtChamp(2) & ";" & TxtCondition1Valeur & ";" & TxtCondition2Valeur & ";" & TxtOrdreValeur

            Case Else

               LstChampSelectionner.AddItem TxtChamp(0) & ";" & TxtChamp(1) & ";" & TxtChamp(2) & ";" & TxtCondition1Valeur & ";" & TxtCondition2Valeur & ";" & TxtOrdreValeur, IntIndex

         End Select

      Case Else

   End Select

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
Private Sub LstChamp_DblClick(Cancel As Integer)

   On Error GoTo Err_LstChamp_DblClick

   Select Case LstChamp.ListIndex
      Case Is >= 0

         LstChampSelectionner.AddItem LstChamp.Column(0) & ";" & LstChamp.Column(1) & ";" & LstChamp.Column(2) & ";;;;"

         LstChamp.RemoveItem LstChamp.ListIndex

      Case Else

   End Select

Exit_LstChamp_DblClick:

   Exit Sub

Err_LstChamp_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChamp_DblClick"

   Resume Exit_LstChamp_DblClick
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChampSelectionner_DblClick(Cancel As Integer)

   On Error GoTo Err_LstChampSelectionner_DblClick

   Select Case LstChampSelectionner.ListIndex
      Case Is >= 0

         LstChamp.AddItem LstChampSelectionner.Column(0) & ";" & LstChampSelectionner.Column(1) & ";" & LstChamp.Column(2)

         LstChampSelectionner.RemoveItem LstChampSelectionner.ListIndex

      Case Else

   End Select

Exit_LstChampSelectionner_DblClick:

   Exit Sub

Err_LstChampSelectionner_DblClick:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChampSelectionner_DblClick"

   Resume Exit_LstChampSelectionner_DblClick
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbTableRequete_AfterUpdate()

   On Error GoTo Err_CmbTableRequete_AfterUpdate

   CreeListeChampTableRequete

   LstChampSelectionner.RowSource = vbNullString

   CmbChampValeurCondition1.RowSource = vbNullString

   CmbChampValeurCondition1 = vbNullString

   CmbChampCondition1 = vbNullString

   CmbChampCondition1.Enabled = True

   CmbChampCondition1.Locked = False

   CmbChampValeurCondition2.RowSource = vbNullString

   CmbChampValeurCondition2 = vbNullString

   CmbChampCondition2 = vbNullString

   CmbChampCondition2.Enabled = True

   CmbChampCondition2.Locked = False

   CmbChampOrdre = vbNullString

   CmbChampOrdre.Enabled = True

   CmbChampOrdre.Locked = False

Exit_CmbTableRequete_AfterUpdate:

   Exit Sub

Err_CmbTableRequete_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbTableRequete_AfterUpdate"

   Resume Exit_CmbTableRequete_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub LstChampSelectionner_AfterUpdate()

   On Error GoTo Err_LstChampSelectionner_AfterUpdate

   CmbChampCondition1.Enabled = True

   CmbChampCondition1.Locked = False

   CmbChampValeurCondition1.RowSource = "SELECT DISTINCT " & TxtChampSelectionner & " FROM " & CmbTableRequete.Column(0) & " ORDER BY " & TxtChampSelectionner & ";"

   Select Case InStr(1, LstChampSelectionner.Column(3), " ")
      Case Is > 0

         CmbChampValeurCondition1.Enabled = True

         CmbChampValeurCondition1.Locked = False

         CmbChampCondition1 = Left(LstChampSelectionner.Column(3), InStr(1, LstChampSelectionner.Column(3), " ") - 1)

         CmbChampValeurCondition1 = Mid(LstChampSelectionner.Column(3), InStr(1, LstChampSelectionner.Column(3), " ") + 1)

      Case Else

         CmbChampValeurCondition1.Enabled = False

         CmbChampValeurCondition1.Locked = True

         CmbChampCondition1 = vbNullString

         CmbChampValeurCondition1 = vbNullString

   End Select

   CmbChampCondition2.Enabled = True

   CmbChampCondition2.Locked = False

   CmbChampValeurCondition2.RowSource = "SELECT DISTINCT " & TxtChampSelectionner & " FROM " & CmbTableRequete.Column(0) & " ORDER BY " & TxtChampSelectionner & ";"

   Select Case InStr(1, LstChampSelectionner.Column(4), " ")
      Case Is > 0

         CmbChampValeurCondition2.Enabled = True

         CmbChampValeurCondition2.Locked = False

         CmbChampCondition2 = Left(LstChampSelectionner.Column(4), InStr(1, LstChampSelectionner.Column(4), " ") - 1)

         CmbChampValeurCondition2 = Mid(LstChampSelectionner.Column(4), InStr(1, LstChampSelectionner.Column(4), " ") + 1)

      Case Else

         CmbChampValeurCondition2.Enabled = False

         CmbChampValeurCondition2.Locked = True

         CmbChampCondition2 = vbNullString

         CmbChampValeurCondition2 = vbNullString

   End Select

   CmbChampOrdre.Enabled = True

   CmbChampOrdre.Locked = False

   Select Case InStr(1, LstChampSelectionner.Column(5), " ")
      Case Is > 0

         CmbChampOrdre = Mid(LstChampSelectionner.Column(5), InStr(1, LstChampSelectionner.Column(5), " ") + 1)

      Case Else

         CmbChampOrdre = vbNullString

   End Select

Exit_LstChampSelectionner_AfterUpdate:

   Exit Sub

Err_LstChampSelectionner_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LstChampSelectionner_AfterUpdate"

   Resume Exit_LstChampSelectionner_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbChampCondition1_AfterUpdate()

   On Error GoTo Err_CmbChampCondition1_AfterUpdate

   Select Case CmbChampCondition1.ListIndex
      Case Is <= 0

         CmbChampValeurCondition1 = vbNullString

         CmbChampValeurCondition1.Enabled = False

         CmbChampValeurCondition1.Locked = True

      Case Else

         CmbChampValeurCondition1.Enabled = True

         CmbChampValeurCondition1.Locked = False

   End Select

Exit_CmbChampCondition1_AfterUpdate:

   Exit Sub

Err_CmbChampCondition1_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbChampCondition1_AfterUpdate"

   Resume Exit_CmbChampCondition1_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbChampCondition2_AfterUpdate()

   On Error GoTo Err_CmbChampCondition2_AfterUpdate

   Select Case CmbChampCondition2.ListIndex
      Case Is <= 0

         CmbChampValeurCondition2 = vbNullString

         CmbChampValeurCondition2.Enabled = False

         CmbChampValeurCondition2.Locked = True

      Case Else

         CmbChampValeurCondition2.Enabled = True

         CmbChampValeurCondition2.Locked = False

   End Select

Exit_CmbChampCondition2_AfterUpdate:

   Exit Sub

Err_CmbChampCondition2_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbChampCondition2_AfterUpdate"

   Resume Exit_CmbChampCondition2_AfterUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbChampValeurCondition1_BeforeUpdate(Cancel As Integer)
   Dim VarValeur As Variant

   On Error GoTo Err_CmbChampValeurCondition1_BeforeUpdate

   Select Case CmbChampCondition1.Column(0)
      Case "<", "<=", "=", ">=", ">", "<>"

         On Error Resume Next

         Select Case LstChampSelectionner.Column(2)
            Case dbByte

               VarValeur = CByte(CmbChampValeurCondition1)

            Case dbChar, dbText

               VarValeur = CStr(CmbChampValeurCondition1)

            Case dbCurrency

               VarValeur = CCur(CmbChampValeurCondition1)

            Case dbDate, dbTime, dbTimeStamp

               VarValeur = CDate(CmbChampValeurCondition1)

            Case dbDecimal

               VarValeur = CDec(CmbChampValeurCondition1)

            Case dbDouble

               VarValeur = CDbl(CmbChampValeurCondition1)

            Case dbSingle

               VarValeur = CSng(CmbChampValeurCondition1)

            Case dbInteger

               VarValeur = CInt(CmbChampValeurCondition1)

            Case dbLong

               VarValeur = CLng(CmbChampValeurCondition1)

            Case Else

         End Select

         Select Case Err.Number
            Case 0

            Case Else

               Err.Clear

               MsgBox "La valeur saisie est invalide"

               Cancel = True

         End Select

         On Error GoTo Err_CmbChampValeurCondition1_BeforeUpdate

      Case Else

   End Select

Exit_CmbChampValeurCondition1_BeforeUpdate:

   Exit Sub

Err_CmbChampValeurCondition1_BeforeUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbChampValeurCondition1_BeforeUpdate"

   Resume Exit_CmbChampValeurCondition1_BeforeUpdate
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmbChampValeurCondition2_BeforeUpdate(Cancel As Integer)
   Dim VarValeur As Variant

   On Error GoTo Err_CmbChampValeurCondition2_BeforeUpdate

   Select Case CmbChampCondition2.Column(0)
      Case "<", "<=", "=", ">=", ">", "<>"

         On Error Resume Next

         Select Case LstChampSelectionner.Column(2)
            Case dbByte

               VarValeur = CByte(CmbChampValeurCondition2)

            Case dbChar, dbText

               VarValeur = CStr(CmbChampValeurCondition2)

            Case dbCurrency

               VarValeur = CCur(CmbChampValeurCondition2)

            Case dbDate, dbTime, dbTimeStamp

               VarValeur = CDate(CmbChampValeurCondition2)

            Case dbDecimal

               VarValeur = CDec(CmbChampValeurCondition2)

            Case dbDouble

               VarValeur = CDbl(CmbChampValeurCondition2)

            Case dbSingle

               VarValeur = CSng(CmbChampValeurCondition2)

            Case dbInteger

               VarValeur = CInt(CmbChampValeurCondition2)

            Case dbLong

               VarValeur = CLng(CmbChampValeurCondition2)

            Case Else

         End Select

         Select Case Err.Number
            Case 0

            Case Else

               Err.Clear

               MsgBox "La valeur saisie est invalide"

               Cancel = True

         End Select

         On Error GoTo Err_CmbChampValeurCondition2_BeforeUpdate

      Case Else

   End Select

Exit_CmbChampValeurCondition2_BeforeUpdate:

   Exit Sub

Err_CmbChampValeurCondition2_BeforeUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbChampValeurCondition2_BeforeUpdate"

   Resume Exit_CmbChampValeurCondition2_BeforeUpdate
End Sub
