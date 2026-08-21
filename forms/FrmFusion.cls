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

   CmbTableRequete.RowSource = vbNullString

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

   Set QdfRequete = Nothing

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

   LstChampTableRequete.RowSource = vbNullString

   Select Case CmbTableRequete.Column(0)
      Case vbNullString

      Case Else

         Select Case CmbTableRequete.Column(2)
            Case "TABLE"

               Set TdsTable = CurrentDb.TableDefs

               Set TdfTable = TdsTable(CmbTableRequete.Column(0))

               For Each FldChamp In TdfTable.Fields

                  On Error Resume Next

                  LstChampTableRequete.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description")

                  Select Case Err.Number
                     Case vbEmpty

                     Case Else

                        Err.Clear

                        LstChampTableRequete.AddItem FldChamp.Name & ";"

                  End Select

                  On Error GoTo Err_CreeListeChampTableRequete

               Next

            Case "REQUETE"

               Set QdsRequete = CurrentDb.QueryDefs

               Set QdfRequete = QdsRequete(CmbTableRequete.Column(0))

               For Each FldChamp In QdfRequete.Fields

                  On Error Resume Next

                  LstChampTableRequete.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description")

                  Select Case Err.Number
                     Case vbEmpty

                     Case Else

                        Err.Clear

                        LstChampTableRequete.AddItem FldChamp.Name & ";"

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
'***    PROCEDURE:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
'***    EXEMPLE:                                                                                        ***
'******************************************************************************
Private Sub CreeListeChampDocument()
   Dim WshShellSystem As New WshShell
   Dim AppWord As Word.Application
   Dim DocWord As Word.Document
   Dim FldWord As Word.Field
   Dim IntChampNom As Integer
   Dim VarChampNomKey() As Variant
   Dim StrChampNom As String
   Dim DicChampNom As New Scripting.Dictionary
   Dim StrCheminFichierAppWord As String
   Dim StrCheminFichierWord As String

   On Error GoTo Err_CreeListeChampDocument

   LstChampDocument.RowSource = vbNullString

   Set AppWord = CreateObject("Word.Application")

   StrCheminFichierWord = CmbDocument.Column(1) & "\" & CmbDocument.Column(0)

   Select Case StrCheminFichierWord
      Case vbNullString

      Case Else

         On Error Resume Next

         AppWord.Documents.Open (StrCheminFichierWord)

         Select Case Err.Number
            Case 0

            Case 5981

               Err.Clear

            Case Else

               GoTo Err_CreeListeChampDocument

         End Select

         On Error GoTo Err_CreeListeChampDocument

         Set DocWord = AppWord.Documents(StrCheminFichierWord)

         For Each FldWord In DocWord.Fields

            StrChampNom = RemplaceChr(RemplaceChr(FldWord.Code.Text, "MERGEFIELD ", vbNullString), Chr(34), vbNullString)

            Select Case DicChampNom.Exists(StrChampNom)
               Case True

               Case False

                  DicChampNom.Add StrChampNom, StrChampNom

            End Select

         Next

         VarChampNomKey = DicChampNom.Keys

         For IntChampNom = 0 To DicChampNom.Count - 1

            LstChampDocument.AddItem VarChampNomKey(IntChampNom) & ";"

         Next

         DocWord.Close 0

         AppWord.Quit

   End Select

Exit_CreeListeChampDocument:

   Set DocWord = Nothing

   Set AppWord = Nothing

   Set DicChampNom = Nothing

   Exit Sub

Err_CreeListeChampDocument:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeListeChampDocument"

   Resume Exit_CreeListeChampDocument
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

   LstChampTableRequete.RowSource = vbNullString

   LstChampDocument.RowSource = vbNullString

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
Private Sub CmdChampTableRequeteMonter_Click()
   Dim IntIndex As Integer
   Dim StrValeur As String

   On Error GoTo Err_CmdChampTableRequeteMonter_Click

   Select Case LstChampTableRequete.ListIndex
      Case Is <= 0

      Case Else

         IntIndex = LstChampTableRequete.ListIndex

         StrValeur = LstChampTableRequete.Column(0, LstChampTableRequete.ListIndex) & ";" & LstChampTableRequete.Column(1, LstChampTableRequete.ListIndex)

         LstChampTableRequete.RemoveItem IntIndex

         LstChampTableRequete.AddItem StrValeur, IntIndex - 1

         LstChampTableRequete.Selected(IntIndex - 1) = True

   End Select

Exit_CmdChampTableRequeteMonter_Click:

   Exit Sub

Err_CmdChampTableRequeteMonter_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdChampTableRequeteMonter_Click"

   Resume Exit_CmdChampTableRequeteMonter_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdChampTableRequeteDescendre_Click()
   Dim IntIndex As Integer
   Dim StrValeur As String

   On Error GoTo Err_CmdChampTableRequeteDescendre_Click

   Select Case LstChampTableRequete.ListIndex
      Case Is < 0

      Case Is = LstChampTableRequete.ListCount - 1

      Case Else

         IntIndex = LstChampTableRequete.ListIndex

         StrValeur = LstChampTableRequete.Column(0, LstChampTableRequete.ListIndex + 1) & ";" & LstChampTableRequete.Column(1, LstChampTableRequete.ListIndex + 1)

         LstChampTableRequete.RemoveItem IntIndex + 1

         LstChampTableRequete.AddItem StrValeur, IntIndex

         LstChampTableRequete.Selected(IntIndex + 1) = True

   End Select

Exit_CmdChampTableRequeteDescendre_Click:

   Exit Sub

Err_CmdChampTableRequeteDescendre_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdChampTableRequeteDescendre_Click"

   Resume Exit_CmdChampTableRequeteDescendre_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdChampDocumentDescendre_Click()
   Dim IntIndex As Integer
   Dim StrValeur As String

   On Error GoTo Err_CmdChampDocumentDescendre_Click

   Select Case LstChampDocument.ListIndex
      Case Is < 0

      Case Is = LstChampDocument.ListCount - 1

      Case Else

         IntIndex = LstChampDocument.ListIndex

         StrValeur = LstChampDocument.Column(0, LstChampDocument.ListIndex + 1) & ";" & LstChampDocument.Column(1, LstChampDocument.ListIndex + 1)

         LstChampDocument.RemoveItem IntIndex + 1

         LstChampDocument.AddItem StrValeur, IntIndex

         LstChampDocument.Selected(IntIndex + 1) = True

   End Select

Exit_CmdChampDocumentDescendre_Click:

   Exit Sub

Err_CmdChampDocumentDescendre_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdChampDocumentDescendre_Click"

   Resume Exit_CmdChampDocumentDescendre_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdChampDocumentMonter_Click()
   Dim IntIndex As Integer
   Dim StrValeur As String

   On Error GoTo Err_CmdChampDocumentMonter_Click

   Select Case LstChampDocument.ListIndex
      Case Is <= 0

      Case Else

         IntIndex = LstChampDocument.ListIndex

         StrValeur = LstChampDocument.Column(0, LstChampDocument.ListIndex) & ";" & LstChampDocument.Column(1, LstChampDocument.ListIndex)

         LstChampDocument.RemoveItem IntIndex

         LstChampDocument.AddItem StrValeur, IntIndex - 1

         LstChampDocument.Selected(IntIndex - 1) = True

   End Select

Exit_CmdChampDocumentMonter_Click:

   Exit Sub

Err_CmdChampDocumentMonter_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdChampDocumentMonter_Click"

   Resume Exit_CmdChampDocumentMonter_Click
End Sub

'******************************************************************************
'***    EVENEMENT:                                                                                    ***
'******************************************************************************
'***    FONCTION:                                                                                       ***
'******************************************************************************
'***    ENTREE:                                                                                          ***
'***    SORTIE:                                                                                           ***
'******************************************************************************
Private Sub CmdExporteWord_Click()
   Dim StrSQLRequeteTable As String
   Dim IntChamp As Integer

   On Error GoTo Err_CmdExporteWord_Click

   Select Case CmbTableRequete.ListIndex
      Case Is < 0

      Case Else

         Select Case CmbDocument.ListIndex
            Case Is < 0

            Case Else

               Select Case LstChampDocument.ListCount
                  Case 0

                  Case Else

                     StrSQLRequeteTable = "SELECT "

                     For IntChamp = 0 To LstChampDocument.ListCount - 1

                        StrSQLRequeteTable = StrSQLRequeteTable & LstChampTableRequete.Column(0, IntChamp) & " AS " & LstChampDocument.Column(0, IntChamp) & ", "

                     Next

                     StrSQLRequeteTable = Left(StrSQLRequeteTable, Len(StrSQLRequeteTable) - 2) & " FROM " & CmbTableRequete.Column(0) & ";"

                     FusionWordRequete StrSQLRequeteTable, CmbDocument.Column(0)

               End Select

         End Select

   End Select

Exit_CmdExporteWord_Click:

   Exit Sub

Err_CmdExporteWord_Click:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmdExporteWord_Click"

   Resume Exit_CmdExporteWord_Click
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
Private Sub CmbDocument_AfterUpdate()

   On Error GoTo Err_CmbDocument_AfterUpdate

   CreeListeChampDocument

Exit_CmbDocument_AfterUpdate:

   Exit Sub

Err_CmbDocument_AfterUpdate:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CmbDocument_AfterUpdate"

   Resume Exit_CmbDocument_AfterUpdate
End Sub
