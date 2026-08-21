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

                  LstChamp.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description")

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

                  LstChamp.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description")

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
   Dim IntChamp As Integer

   On Error GoTo Err_CmdExporteExcel_Click

   Select Case CmbTableRequete.ListIndex
      Case Is < 0

      Case Else

         Select Case LstChampSelectionner.ListCount
            Case 0

            Case Else

               StrSQLRequeteTable = "SELECT "

               For IntChamp = 0 To LstChampSelectionner.ListCount - 1

                  StrSQLRequeteTable = StrSQLRequeteTable & LstChampSelectionner.Column(0, IntChamp) & ", "

               Next

               StrSQLRequeteTable = Left(StrSQLRequeteTable, Len(StrSQLRequeteTable) - 2) & " FROM " & CmbTableRequete.Column(0) & ";"

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
Private Sub LstChamp_DblClick(Cancel As Integer)

   On Error GoTo Err_LstChamp_DblClick

   Select Case LstChamp.ListIndex
      Case Is >= 0

         LstChampSelectionner.AddItem LstChamp.Column(0) & ";" & LstChamp.Column(1)

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

         LstChamp.AddItem LstChampSelectionner.Column(0) & ";" & LstChampSelectionner.Column(1)

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

Exit_CmbTableRequete_AfterUpdate:

   Exit Sub

Err_CmbTableRequete_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbTableRequete_AfterUpdate"

   Resume Exit_CmbTableRequete_AfterUpdate
End Sub
