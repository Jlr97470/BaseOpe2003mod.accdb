Attribute VB_Name = "ModGestionImportsExports"
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
Public Function ExporteExcelRequete(ByVal StrSQL As String) As Boolean
   Dim AppExcel As Excel.Application
   Dim WbkClasseur As Excel.Workbook
   Dim WstFeuille As Excel.Worksheet
   Dim RsRequete As DAO.Recordset

   On Error GoTo Err_ExporteExcelRequete

   ExporteExcelRequete = True

   DoCmd.Hourglass True

   DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   GEtiFonTraitements.Caption = "ExporteExcelRequete"

   GEtiObjTraitements.Caption = StrSQL

   DoEvents

   Set RsRequete = CurrentDb.OpenRecordset(StrSQL)

   Select Case RsRequete.RecordCount
      Case vbEmpty

      Case Else

         On Error Resume Next

         GEtiTitTraitements.Caption = RsRequete.Properties("Description").Value

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               GEtiTitTraitements.Caption = "Exportation Excel"

         End Select

         On Error GoTo Err_ExporteExcelRequete

         DoEvents

         On Error Resume Next

         Set AppExcel = GetObject(, "Excel.Application")

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               Set AppExcel = CreateObject("Excel.Application")

         End Select

         On Error GoTo Err_ExporteExcelRequete

         Set WbkClasseur = AppExcel.Workbooks.Add

         Set WstFeuille = WbkClasseur.Worksheets(1)

         On Error Resume Next

         WstFeuille.Name = Left(RsRequete.Properties("Description").Value, 31)

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               WstFeuille.Name = "Exportation Excel"

         End Select

         WstFeuille.Rows.RowHeight = 14

         ExporteExcelRequeteDonner WstFeuille, RsRequete, 1, 1

         AppExcel.Visible = True

   End Select

   RsRequete.Close

Exit_ExporteExcelRequete:

   DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

   DoCmd.Hourglass False

   Set RsRequete = Nothing

   Set WstFeuille = Nothing

   Set WbkClasseur = Nothing

   Set AppExcel = Nothing

   Exit Function

Err_ExporteExcelRequete:

   ExporteExcelRequete = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ExporteExcelRequete"

   Resume Exit_ExporteExcelRequete
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
Public Function ExporteExcelRequeteDonner(ByRef WstFeuille As Worksheet, ByRef RsRequete As DAO.Recordset, ByVal LngPosCol As Long, ByVal LngPosLig As Long) As Boolean
   Dim FldChamp As DAO.Field
   Dim HypChamp As Excel.Hyperlinks
   Dim LngPosColChamp As Long
   Dim LngPosLigChamp As Long

   On Error GoTo Err_ExporteExcelRequeteDonner

   ExporteExcelRequeteDonner = True

   DoCmd.Hourglass True

   DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   GEtiFonTraitements.Caption = "ExporteExcelRequeteDonner"

   GEtiObjTraitements.Caption = RsRequete.Name

   DoEvents

   With WstFeuille

         On Error Resume Next

         .Cells(LngPosLig, LngPosCol).Value = RsRequete.Properties("Description").Value

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               .Cells(LngPosLig, LngPosCol).Value = "Exportation Excel"

         End Select

         On Error GoTo Err_ExporteExcelRequeteDonner

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).MergeCells = True

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).HorizontalAlignment = -4108

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).Borders.LineStyle = 1

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).Interior.Color = RGB(100, 100, 100)

         LngPosColChamp = vbEmpty

         GPbrTraitements.Min = 0

         GPbrTraitements.Value = 1

         GPbrTraitements.Max = RsRequete.Fields.Count

         DoEvents

         For Each FldChamp In RsRequete.Fields

            On Error Resume Next

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Properties("Description").Value

            Select Case Err.Number
               Case vbEmpty

               Case Else

                  Err.Clear

                  .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Name

            End Select

            On Error GoTo Err_ExporteExcelRequeteDonner

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Borders.LineStyle = 1

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Borders.Weight = 3

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).HorizontalAlignment = -4108

            LngPosColChamp = LngPosColChamp + 1

            GEtiObjTraitements.Caption = "Exportation Des Libelles Des Champs : " & FldChamp.Name

            GPbrTraitements.Value = LngPosColChamp

            DoEvents

         Next

         GEtiObjTraitements.Caption = "Exportation Des Données"

         DoEvents

         .Cells(LngPosLig + 2, LngPosCol).CopyFromRecordset RsRequete

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders.LineStyle = 1

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(8).Weight = 3

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(9).Weight = 3

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(7).Weight = 3

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(10).Weight = 3

         LngPosColChamp = vbEmpty

         GEtiObjTraitements.Caption = "Mise En Forme De La Feuille Excel"

         GPbrTraitements.Min = 0

         GPbrTraitements.Max = RsRequete.Fields.Count

         GPbrTraitements.Value = 1

         DoEvents

         For Each FldChamp In RsRequete.Fields

            On Error Resume Next

            Select Case FldChamp.Properties("Format")
               Case vbNullString

               Case "Percent"

                  .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "0.00%"

               Case "Short Date"

                  .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "DD/MM/YYYY"

               Case "General Date"

                  .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "DD/MM/YYYY HH:MM"

               Case Else

            End Select

            On Error GoTo Err_ExporteExcelRequeteDonner

            Select Case InStr(1, FldChamp.Name, "EMail")
               Case 0

               Case Else

                  For LngPosLigChamp = 0 To RsRequete.RecordCount - 1

                     Select Case InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#")
                        Case 0

                        Case Else

                           .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value = Left(.Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#") - 1)

                     End Select

                     Set HypChamp = .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Hyperlinks

                     HypChamp.Add .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp), "mailto:" & .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value

                  Next

            End Select

            Select Case InStr(1, FldChamp.Name, "Web")
               Case 0

               Case Else

                  For LngPosLigChamp = 0 To RsRequete.RecordCount - 1

                     Select Case InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#")
                        Case 0

                        Case Else

                           .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value = Left(.Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#") - 1)

                     End Select

                     Set HypChamp = .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Hyperlinks

                     HypChamp.Add .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp), "http://" & .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value

                  Next

            End Select

            .Columns(LngPosCol + LngPosColChamp).AutoFit

            LngPosColChamp = LngPosColChamp + 1

            GPbrTraitements.Value = LngPosColChamp

            DoEvents

         Next

         .PageSetup.Orientation = 2

         .PageSetup.Zoom = False

         .PageSetup.FitToPagesWide = 1

         .PageSetup.FitToPagesTall = 10

   End With

Exit_ExporteExcelRequeteDonner:

   DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

   DoCmd.Hourglass False

   Set HypChamp = Nothing

   Exit Function

Err_ExporteExcelRequeteDonner:

   ExporteExcelRequeteDonner = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ExporteExcelRequeteDonner"

   Resume Exit_ExporteExcelRequeteDonner
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
Public Function FusionWordRequete(ByVal StrSQL As String, ByVal StrFichier As String) As Boolean
   Dim WshShellSystem As New IWshRuntimeLibrary.WshShell
   Dim AppWord As Word.Application
   Dim DocWord As Word.Document
   Dim MerWord As Word.MailMerge
   Dim TdfRequete As DAO.QueryDef
   Dim TdsRequete As DAO.QueryDefs
   Dim StrRequeteNom As String
   Dim StrCheminFichierAppWord As String
   Dim StrCheminFichierWord As String

   On Error GoTo Err_FusionWordRequete

   Set AppWord = CreateObject("Word.Application")

   For Each DocWord In AppWord.Documents

      DocWord.Close 0

   Next

   AppWord.Visible = True

   StrCheminFichierWord = Nz(DLookup("FicValeur", "SelFichiersDetailler", "FicType='FICDOC' AND FicCode LIKE '*=" & StrFichier & "' AND FicValide=True"))

   Select Case StrCheminFichierWord
      Case vbNullString

      Case Else

         On Error Resume Next

         AppWord.Documents.Open (StrCheminFichierWord & "\" & StrFichier)

         Select Case Err.Number
            Case 0

            Case 5981

               Err.Clear

            Case Else

               GoTo Err_FusionWordRequete

         End Select

         On Error GoTo Err_FusionWordRequete

         Select Case Left(StrSQL, 6)
            Case "SELECT"

            Case Else

               StrSQL = "SELECT " & StrSQL & ".* FROM " & StrSQL & ";"

         End Select

         Set DocWord = AppWord.Documents(StrFichier)

         Set MerWord = DocWord.MailMerge

         Set TdsRequete = CurrentDb.QueryDefs

         StrRequeteNom = "SelFusionWord" & Int(Rnd() * 10) & Int(Rnd() * 10)

         Set TdfRequete = CurrentDb.CreateQueryDef(StrRequeteNom, StrSQL)

         TdsRequete.Refresh

         MerWord.OpenDataSource Name:=CurrentDb.Name, Connection:="DSN=MS Access Databases;DBQ=" & CurrentDb.Name & ";", SQLStatement:="SELECT " & StrRequeteNom & ".* FROM " & StrRequeteNom & ";"

         MerWord.Destination = 0

         DocWord.MailMerge.Execute

         DocWord.Close 0

         Set TdfRequete = Nothing

         TdsRequete.Delete (StrRequeteNom)

   End Select

Exit_FusionWordRequete:

   Set WshShellSystem = Nothing

   Set MerWord = Nothing

   Set DocWord = Nothing

   Set AppWord = Nothing

   On Error Resume Next

   Set TdfRequete = Nothing

   TdsRequete.Delete (StrRequeteNom)

   Set TdsRequete = Nothing

   Exit Function

Err_FusionWordRequete:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "FusionWordRequete"

   Resume Exit_FusionWordRequete
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
Public Function ExporteMailingRequete(ByVal StrEmailExpediteur As String, ByVal StrEmailCopieCacher As String, ByVal StrSQLRequete As String, ByVal StrChampDestination As String, ByVal StrSujet As String, ByVal StrFichierBody As String, ByRef StrFichierAttache() As String, ByVal BlnSimulation As Boolean) As Boolean

   On Error GoTo Err_ExporteMailingRequete

   Select Case DLookup("ParValeur", "SelParametresDetailler", "ParType='USER' ANd ParCode='MAILING'")
      Case "JAVAMAIL"

         ExporteMailingRequete = MailingJavaMailRequete(StrEmailExpediteur, StrEmailCopieCacher, StrSQLRequete, StrChampDestination, StrSujet, StrFichierBody, StrFichierAttache(), BlnSimulation)

      Case "VBSENDMAIL"

         ExporteMailingRequete = MailingSendMailRequete(StrEmailExpediteur, StrEmailCopieCacher, StrSQLRequete, StrChampDestination, StrSujet, StrFichierBody, StrFichierAttache(), BlnSimulation)

   End Select

Exit_ExporteMailingRequete:

   Exit Function

Err_ExporteMailingRequete:

   ExporteMailingRequete = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ExporteMailingRequete"

   Resume Exit_ExporteMailingRequete
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
Public Function MailingJavaMailRequete(ByVal StrEmailExpediteur As String, ByVal StrEmailCopieCacher As String, ByVal StrSQLRequete As String, ByVal StrChampDestination As String, ByVal StrSujet As String, ByVal StrFichierBody As String, ByRef StrFichierAttache() As String, ByVal BlnSimulation As Boolean) As Boolean
   Dim FsoSystem As Scripting.FileSystemObject
   Dim FilFichier As Scripting.File
   Dim JavaMsg As Object 'jmail.Message
   Dim HtmObjectDocument As MSHTML.HTMLDocument
   Dim HtmDocument As MSHTML.HTMLDocument
   Dim RsRequete As DAO.Recordset
   Dim RsFichier As DAO.Recordset
   Dim StrEmailFrom As String
   Dim StrAttacheID As String
   Dim StrAttacheFichier As String
   Dim StrDocument As String
   Dim StrCheminFichierBody As String
   Dim StrChampNom As String
   Dim IntNb As Integer
   Dim IntPosDebut As Integer
   Dim IntPosFin As Integer

   On Error GoTo Err_MailingJavaMailRequete

   MailingJavaMailRequete = True

   DoCmd.Hourglass True

   DoCmd.OpenForm "FrmNotes", acNormal, , , acFormEdit, acWindowNormal

   DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   GEtiTitTraitements.Caption = "Mailing : " & StrSujet

   GEtiFonTraitements.Caption = "MailingJavaMailRequete"

   GEtiObjTraitements.Caption = StrSQLRequete

   DoEvents

   Set RsFichier = CurrentDb.OpenRecordset("SELECT SelFichiersDetailler.* FROM SelFichiersDetailler WHERE FicCode2='" & StrFichierBody & "' AND FicValide=True")

   Select Case RsFichier.EOF
      Case True

      Case False

         StrCheminFichierBody = RsFichier!FicValeur & "\" & RsFichier!FicCode2

         RsFichier.Close

         Set RsRequete = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case RsRequete.EOF
            Case True

            Case False

               Set HtmObjectDocument = New MSHTML.HTMLDocument

               Set HtmDocument = HtmObjectDocument.createDocumentFromUrl(StrCheminFichierBody, vbNullString)

               Do Until HtmDocument.readyState = "complete"

                  DoEvents

               Loop

              ' Set JavaMsg = New jmail.Message

               GEtiObjTraitements.Caption = "CONNEXION " & vbCrLf

               JavaMsg.from = StrEmailExpediteur

               JavaMsg.Subject = StrSujet

               JavaMsg.AppendText ("Apparemment, votre logiciel d'email ne supporte pas le format HTML.")

               JavaMsg.Priority = 3

               JavaMsg.ReturnReceipt = True

               JavaMsg.Logging = True

               JavaMsg.MailServerUserName = vbNullString

               JavaMsg.MailServerPassWord = vbNullString

               Set FsoSystem = New Scripting.FileSystemObject

               Select Case HtmDocument.images.Length
                  Case 0

                  Case Else

                     For IntNb = 0 To HtmDocument.images.Length - 1

                        StrAttacheFichier = RemplaceChr(RemplaceChr(RemplaceChr(HtmDocument.images(IntNb).src, "file:///", vbNullString), "%20", " "), "/", "\")

                        Select Case FsoSystem.FileExists(StrAttacheFichier)
                           Case True

                              StrAttacheID = JavaMsg.AddAttachment(StrAttacheFichier, True)

                              Set FilFichier = FsoSystem.GetFile(StrAttacheFichier)

                              HtmDocument.images(IntNb).src = "cid:" & StrAttacheID

                           Case False

                        End Select

                     Next

               End Select

               On Error Resume Next

               Select Case StrFichierAttache(LBound(StrFichierAttache))
                  Case vbNullString

                  Case Else

                     For IntNb = LBound(StrFichierAttache) To UBound(StrFichierAttache)

                        StrAttacheFichier = StrFichierAttache(IntNb)

                        Select Case FsoSystem.FileExists(StrAttacheFichier)
                           Case True

                              JavaMsg.AddCustomAttachment StrAttacheFichier, Right(StrAttacheFichier, Len(StrAttacheFichier) - InStrRev(StrAttacheFichier, "\")), False

                           Case False

                        End Select

                     Next

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

               End Select

               On Error GoTo Err_MailingJavaMailRequete

               RsRequete.MoveLast

               RsRequete.MoveFirst

               GPbrTraitements.Min = 0

               GPbrTraitements.Value = 0

               GPbrTraitements.Max = RsRequete.RecordCount

               Do Until RsRequete.EOF = True

                  JavaMsg.ClearRecipients

                  Select Case BlnSimulation
                        Case True

                        Case False

                           JavaMsg.AddRecipientBCC StrEmailCopieCacher

                  End Select

                  Select Case IsNull(RsRequete(StrChampDestination))
                     Case False

                        Select Case InStr(1, RsRequete(StrChampDestination), "mailto:")
                           Case Is <= 0

                              Select Case InStr(1, RsRequete(StrChampDestination), "#")
                                 Case Is <= 0

                                    StrEmailFrom = RsRequete(StrChampDestination)

                                 Case Else

                                    StrEmailFrom = Left(RsRequete(StrChampDestination), InStr(1, RsRequete(StrChampDestination), "#") - 1)

                              End Select

                           Case Else

                              StrEmailFrom = RemplaceChr(Mid(RsRequete(StrChampDestination), InStr(1, RsRequete(StrChampDestination), "mailto:") + 7), "#", vbNullString)

                        End Select

                        Select Case BlnSimulation
                           Case True

                              JavaMsg.AddRecipient StrEmailExpediteur

                              JavaMsg.Subject = StrSujet & " (" & StrEmailFrom & ")"

                           Case False

                              JavaMsg.AddRecipient StrEmailFrom

                        End Select

                        StrDocument = "<HTML>" & vbCrLf & HtmDocument.documentElement.innerHTML & vbCrLf & "</HTML>"

                        IntPosDebut = 1

                        Do Until InStr(IntPosDebut, StrDocument, "<D VALUE=") <= 0

                           IntPosDebut = InStr(IntPosDebut, StrDocument, "<D VALUE=")

                           IntPosFin = InStr(IntPosDebut, StrDocument, ">")

                           StrChampNom = RemplaceChr(Mid(StrDocument, IntPosDebut + 10, IntPosFin - IntPosDebut - 11), Chr$(34), "")

                           IntPosFin = InStr(IntPosDebut, StrDocument, "</D>")

                           On Error Resume Next

                           StrDocument = Left(StrDocument, IntPosDebut - 1) & RsRequete(StrChampNom) & Mid(StrDocument, IntPosFin + 4)

                           Select Case Err.Number
                              Case vbEmpty

                              Case Else

                                 Err.Clear

                                 IntPosDebut = IntPosFin

                           End Select

                           On Error GoTo Err_MailingJavaMailRequete

                        Loop

                        JavaMsg.HTMLBody = StrDocument

                     Case True

                  End Select

                  GPbrTraitements.Value = GPbrTraitements.Value + 1

                 Select Case BlnSimulation
                     Case True

                        GEtiObjTraitements.Caption = GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailExpediteur & "(" & StrEmailFrom & ")" & vbCrLf

                    '    TxtNotesTraitements = TxtNotesTraitements & GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailExpediteur & "(" & StrEmailFrom & ")" & vbCrLf

                     Case False

                        GEtiObjTraitements.Caption = GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailFrom & vbCrLf

                     '   TxtNotesTraitements = TxtNotesTraitements & GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailFrom & vbCrLf

                  End Select

                  Select Case JavaMsg.Send("172.23.4.201")
                     Case True

                     '   TxtNotesTraitements = TxtNotesTraitements & JavaMsg.Log & vbCrLf

                      '  TxtNotesTraitements = TxtNotesTraitements & "Reussi" & vbCrLf

                     Case False

                     '   TxtNotesTraitements = TxtNotesTraitements & JavaMsg.Log & vbCrLf

                     '   TxtNotesTraitements = TxtNotesTraitements & "Echec" & vbCrLf

                  End Select

                  DoEvents

                  RsRequete.MoveNext

               Loop

         End Select

         RsRequete.Close

   End Select

Exit_MailingJavaMailRequete:

   DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

   DoCmd.Hourglass False

   Set FsoSystem = Nothing

   Set FilFichier = Nothing

   Set HtmObjectDocument = Nothing

   Set HtmDocument = Nothing

   Set JavaMsg = Nothing

   Set RsFichier = Nothing

   Set RsRequete = Nothing

   Exit Function

Err_MailingJavaMailRequete:

   MailingJavaMailRequete = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "MailingJavaMailRequete"

   Resume Exit_MailingJavaMailRequete
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
Public Function MailingSendMailRequete(ByVal StrEmailExpediteur As String, ByVal StrEmailCopieCacher As String, ByVal StrSQLRequete As String, ByVal StrChampDestination As String, ByVal StrSujet As String, ByVal StrFichierBody As String, ByRef StrFichierAttache() As String, ByVal BlnSimulation As Boolean) As Boolean
   Dim FsoSystem As Scripting.FileSystemObject
   Dim FilFichier As Scripting.File
   Dim SndMsg As Object 'clsSendMail
   Dim SndMsgStatus As ClsSendMailStatus
   Dim HtmObjectDocument As MSHTML.HTMLDocument
   Dim HtmDocument As MSHTML.HTMLDocument
   Dim RsRequete As DAO.Recordset
   Dim RsFichier As DAO.Recordset
   Dim StrEmailFrom As String
   Dim StrAttache As String
   Dim StrAttacheFichier As String
   Dim StrDocument As String
   Dim StrCheminFichierBody As String
   Dim StrChampNom As String
   Dim IntNb As Integer
   Dim IntPosDebut As Integer
   Dim IntPosFin As Integer

   On Error GoTo Err_MailingSendMailRequete

   MailingSendMailRequete = True

   DoCmd.Hourglass True

   DoCmd.OpenForm "FrmNotes", acNormal, , , acFormEdit, acWindowNormal

   DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   GEtiTitTraitements.Caption = "Mailing : " & StrSujet

   GEtiFonTraitements.Caption = "MailingSendMailRequete"

   GEtiObjTraitements.Caption = StrSQLRequete

   DoEvents

   Set RsFichier = CurrentDb.OpenRecordset("SELECT SelFichiersDetailler.* FROM SelFichiersDetailler WHERE FicCode2='" & StrFichierBody & "' AND FicValide=True")

   Select Case RsFichier.EOF
      Case True

      Case False

         StrCheminFichierBody = RsFichier!FicValeur & "\" & RsFichier!FicCode2

         RsFichier.Close

         Set RsRequete = CurrentDb.OpenRecordset(StrSQLRequete)

         Select Case RsRequete.EOF
            Case True

            Case False

               Set HtmObjectDocument = New MSHTML.HTMLDocument

               Set HtmDocument = HtmObjectDocument.createDocumentFromUrl(StrCheminFichierBody, vbNullString)

               Do Until HtmDocument.readyState = "complete"

                  DoEvents

               Loop

           '    Set SndMsg = New clsSendMail

              ' SndMsg.SMTPHostValidation = VALIDATE_NONE

             '  SndMsg.EmailAddressValidation = VALIDATE_SYNTAX

               SndMsg.Delimiter = ";"

               SndMsg.SMTPHost = "172.23.4.201"

               GEtiObjTraitements.Caption = "CONNEXION A " & SndMsg.SMTPHost & vbCrLf

               SndMsg.from = StrEmailExpediteur

               SndMsg.ReplyToAddress = StrEmailExpediteur

               Select Case BlnSimulation
                     Case True

                     Case False

                        SndMsg.BccRecipient = StrEmailCopieCacher

               End Select

               SndMsg.Subject = StrSujet

               SndMsg.AsHTML = True

               SndMsg.ContentBase = vbNullString

          '     SndMsg.EncodeType = MIME_ENCODE

          '     SndMsg.Priority = NORMAL_PRIORITY

               SndMsg.Receipt = True

               SndMsg.UseAuthentication = False

               SndMsg.UsePopAuthentication = False

               SndMsg.UserName = vbNullString

               SndMsg.Password = vbNullString

               SndMsg.POP3Host = vbNullString

               SndMsg.MaxRecipients = 100

               StrAttache = vbNullString

               Set FsoSystem = New Scripting.FileSystemObject

               Select Case HtmDocument.images.Length
                  Case 0

                  Case Else

                     For IntNb = 0 To HtmDocument.images.Length - 1

                        StrAttacheFichier = RemplaceChr(RemplaceChr(RemplaceChr(HtmDocument.images(IntNb).src, "file:///", vbNullString), "%20", " "), "/", "\")

                        Select Case FsoSystem.FileExists(StrAttacheFichier)
                           Case True

                              StrAttache = StrAttache & StrAttacheFichier & SndMsg.Delimiter

                              Set FilFichier = FsoSystem.GetFile(StrAttacheFichier)

                              HtmDocument.images(IntNb).src = "cid:" & FilFichier.Name

                           Case False

                        End Select

                     Next

               End Select

               On Error Resume Next

               Select Case StrFichierAttache(LBound(StrFichierAttache))
                  Case vbNullString

                  Case Else

                     For IntNb = LBound(StrFichierAttache) To UBound(StrFichierAttache)

                        StrAttacheFichier = StrFichierAttache(IntNb)

                        Select Case FsoSystem.FileExists(StrAttacheFichier)
                           Case True

                              StrAttache = StrAttache & StrAttacheFichier & SndMsg.Delimiter

                           Case False

                        End Select

                     Next

               End Select

               Select Case Err.Number
                  Case vbEmpty

                  Case Else

                     Err.Clear

               End Select

               On Error GoTo Err_MailingSendMailRequete

               Select Case StrAttache
                  Case vbNullString

                  Case Else

                     StrAttache = Left(StrAttache, Len(StrAttache) - 1)

               End Select

               SndMsg.Attachment = StrAttache

               Set SndMsgStatus = New ClsSendMailStatus

               Set SndMsgStatus.Mail = SndMsg

               RsRequete.MoveLast

               RsRequete.MoveFirst

               GPbrTraitements.Min = 0

               GPbrTraitements.Value = 0

               GPbrTraitements.Max = RsRequete.RecordCount

               Select Case SndMsg.Connect
                  Case True

                     Do Until RsRequete.EOF = True

                        Select Case IsNull(RsRequete(StrChampDestination))
                           Case False

                              Select Case InStr(1, RsRequete(StrChampDestination), "mailto:")
                                 Case Is <= 0

                                    Select Case InStr(1, RsRequete(StrChampDestination), "#")
                                       Case Is <= 0

                                          StrEmailFrom = RsRequete(StrChampDestination)

                                       Case Else

                                          StrEmailFrom = Left(RsRequete(StrChampDestination), InStr(1, RsRequete(StrChampDestination), "#") - 1)

                                    End Select

                                 Case Else

                                    StrEmailFrom = RemplaceChr(Mid(RsRequete(StrChampDestination), InStr(1, RsRequete(StrChampDestination), "mailto:") + 7), "#", vbNullString)

                              End Select

                              Select Case BlnSimulation
                                 Case True

                                    SndMsg.Recipient = StrEmailExpediteur

                                    SndMsg.Subject = StrSujet & " (" & StrEmailFrom & ")"

                                 Case False

                                    SndMsg.Recipient = StrEmailFrom

                              End Select

                              StrDocument = "<HTML>" & vbCrLf & HtmDocument.documentElement.innerHTML & vbCrLf & "</HTML>"

                              IntPosDebut = 1

                              Do Until InStr(IntPosDebut, StrDocument, "<D VALUE=") <= 0

                                 IntPosDebut = InStr(IntPosDebut, StrDocument, "<D VALUE=")

                                 IntPosFin = InStr(IntPosDebut, StrDocument, ">")

                                 StrChampNom = RemplaceChr(Mid(StrDocument, IntPosDebut + 10, IntPosFin - IntPosDebut - 11), Chr$(34), "")

                                 IntPosFin = InStr(IntPosDebut, StrDocument, "</D>")

                                 On Error Resume Next

                                 StrDocument = Left(StrDocument, IntPosDebut - 1) & RsRequete(StrChampNom) & Mid(StrDocument, IntPosFin + 4)

                                 Select Case Err.Number
                                    Case vbEmpty

                                    Case Else

                                       Err.Clear

                                       IntPosDebut = IntPosFin

                                 End Select

                                 On Error GoTo Err_MailingSendMailRequete

                              Loop

                              SndMsg.Message = StrDocument

                           Case True

                        End Select

                        GPbrTraitements.Value = GPbrTraitements.Value + 1

                        Select Case BlnSimulation
                            Case True

                               GEtiObjTraitements.Caption = GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailExpediteur & "(" & StrEmailFrom & ")" & vbCrLf

                        '       TxtNotesTraitements = TxtNotesTraitements & GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailExpediteur & "(" & StrEmailFrom & ")" & vbCrLf

                            Case False

                               GEtiObjTraitements.Caption = GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailFrom & vbCrLf

                          '     TxtNotesTraitements = TxtNotesTraitements & GPbrTraitements.Value & "/" & RsRequete.RecordCount & " Envoie Mail a  : " & StrEmailFrom & vbCrLf

                         End Select

                        SndMsg.Send

                        DoEvents

                        RsRequete.MoveNext

                     Loop

                     SndMsg.Disconnect

               End Select

               RsRequete.Close

         End Select

   End Select

Exit_MailingSendMailRequete:

   DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

   DoCmd.Hourglass False

   Set FsoSystem = Nothing

   Set FilFichier = Nothing

   Set HtmObjectDocument = Nothing

   Set HtmDocument = Nothing

   Set SndMsg = Nothing

   Set SndMsgStatus = Nothing

   Set RsFichier = Nothing

   Set RsRequete = Nothing

   Exit Function

Err_MailingSendMailRequete:

   MailingSendMailRequete = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "MailingSendMailRequete"

   Resume Exit_MailingSendMailRequete
End Function
