Attribute VB_Name = "ModGestionFichiers"
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
Public Function LectureTableInformationConnexion(ByVal TdfTable As TableDef, ByRef StrType As String, ByRef StrParametre, ByRef StrChemin As String, ByRef StrFichier As String) As Boolean
   Dim StrConnexion As String
   Dim IntConnexionLongueur As Integer
   Dim StrCheminFichier As String
   Dim IntCheminFichierLongueur As Integer

   On Error GoTo Err_LectureTableInformationConnexion

   StrConnexion = TdfTable.Connect

   IntConnexionLongueur = Len(StrConnexion)

   Select Case IntConnexionLongueur
      Case vbEmpty

         StrType = vbNullString

         StrParametre = vbNullString

         StrChemin = vbNullString

         StrFichier = vbNullString

         LectureTableInformationConnexion = False

         GoTo Exit_LectureTableInformationConnexion

      Case Else

   End Select

   StrType = Mid(StrConnexion, 1, InStr(1, StrConnexion, ";") - 1)

   StrCheminFichier = Mid(StrConnexion, InStr(1, StrConnexion, "DATABASE=") + 9)

   StrParametre = RemplaceChr(RemplaceChr(StrConnexion, StrType & ";", vbNullString), "DATABASE=" & StrCheminFichier, vbNullString)

   IntCheminFichierLongueur = Len(StrCheminFichier)

   Select Case IntCheminFichierLongueur
      Case vbEmpty

         StrChemin = vbNullString

         StrFichier = vbNullString

         LectureTableInformationConnexion = False

      Case Else

         Select Case StrType
            Case vbNullString, "Excel 5.0"

               StrChemin = Mid(StrCheminFichier, 1, InStrRev(StrCheminFichier, "\") - 1)

               StrFichier = Mid(StrCheminFichier, InStrRev(StrCheminFichier, "\") + 1)

            Case "Text"

               StrChemin = StrCheminFichier

               StrFichier = TdfTable.SourceTableName

            Case Else

         End Select

         LectureTableInformationConnexion = True

   End Select

Exit_LectureTableInformationConnexion:

   Exit Function

Err_LectureTableInformationConnexion:

   LectureTableInformationConnexion = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "LectureTableInformationConnexion"

   Resume Exit_LectureTableInformationConnexion
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
Public Function SauveTableInformationConnexion(ByVal TdfTable As TableDef, ByRef StrType As String, ByRef StrParametre, ByRef StrChemin As String, ByRef StrFichier As String) As Boolean
   Dim strConnect As String

   On Error GoTo Err_SauveTableInformationConnexion

   SauveTableInformationConnexion = True

   Select Case StrType
      Case vbNullString

         strConnect = ";DATABASE=" & StrChemin & "\" & StrFichier

      Case "Excel 5.0"

         strConnect = StrType & ";" & StrParametre & ";DATABASE=" & StrChemin & "\" & StrFichier

      Case "Text"

         strConnect = StrType & ";" & StrParametre & ";DATABASE=" & StrChemin

      Case Else

   End Select

   TdfTable.Connect = strConnect

   TdfTable.RefreshLink

Exit_SauveTableInformationConnexion:

   Exit Function

Err_SauveTableInformationConnexion:

   SauveTableInformationConnexion = False

   MsgBox Err.Number & " " & Err.Description

   Resume Exit_SauveTableInformationConnexion
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
Public Function VerifieTableInformationConnexion() As Boolean
   Dim FsoFichierSystem As New Scripting.FileSystemObject
   Dim TdsTableListe As DAO.TableDefs
   Dim TdfTable As DAO.TableDef
   Dim BolConnectionSauve As Boolean
   Dim StrConnection As String
   Dim StrType As String
   Dim StrParametre As String
   Dim StrFichier As String
   Dim StrExtension As String
   Dim StrChemin As String
   Dim StrCheminDefault As String

   On Error GoTo Err_VerifieTableInformationConnexion

   Set TdsTableListe = CurrentDb.TableDefs

   VerifieTableInformationConnexion = True

   DoCmd.Hourglass True

   DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   GEtiTitTraitements.Caption = "Verification Des Fichiers Des Tables Attachées"

   GEtiFonTraitements.Caption = "VerifieTableInformationConnexion"

   GPbrTraitements.Min = 0

   GPbrTraitements.Max = TdsTableListe.Count

   GEtiObjTraitements.Caption = "Nombre De Tables : " & TdsTableListe.Count

   DoEvents

   For Each TdfTable In TdsTableListe

      Select Case LectureTableInformationConnexion(TdfTable, StrType, StrParametre, StrChemin, StrFichier)
         Case True

            BolConnectionSauve = False

            StrConnection = GetSetting(CurrentDb.Name, "TABLECONNECTION", TdfTable.Name, vbNullString)

            Select Case StrConnection
               Case vbNullString

               Case Else

                  Select Case TdfTable.Connect
                     Case StrConnection

                        BolConnectionSauve = False

                     Case Else

                        BolConnectionSauve = True

                  End Select

                  TdfTable.Connect = StrConnection

            End Select

            LectureTableInformationConnexion TdfTable, StrType, StrParametre, StrChemin, StrFichier

            StrExtension = UCase(Right(StrFichier, Len(StrFichier) - InStrRev(StrFichier, ".")))

            CurrentDb.Execute "UPDATE TBLFICHIERS SET FicValide=False WHERE FicType='FIC" & StrExtension & "' AND FicCode LIKE '*=" & StrFichier & "' ;"

            Select Case FsoFichierSystem.FileExists(StrChemin & "\" & StrFichier)
               Case True

                  StrCheminDefault = Nz(DLookup("FicValeur", "TBLFICHIERS", "FicType='FIC" & StrExtension & "' AND FicCode='DEFAULT=" & StrFichier & "'"))

                  Select Case StrCheminDefault
                     Case vbNullString

                        CurrentDb.Execute "INSERT INTO TBLFICHIERS (FicType,FicCode,FicValeur,FicValide) VALUES('FIC" & StrExtension & "','DEFAULT=" & StrFichier & "','" & StrChemin & "',True) ;"

                     Case Else

                        CurrentDb.Execute "UPDATE TBLFICHIERS SET FicValide=True WHERE FicType='FIC" & StrExtension & "' AND FicCode LIKE '*=" & StrFichier & "' AND FicValeur='" & StrChemin & "' ;"

                  End Select

                  SaveSetting CurrentDb.Name, "TABLECONNECTION", TdfTable.Name, TdfTable.Connect

                  Select Case BolConnectionSauve
                     Case True

                        SauveTableInformationConnexion TdfTable, StrType, StrParametre, StrChemin, StrFichier

                     Case False

                  End Select

               Case False

                  VerifieTableInformationConnexion = False

            End Select

         Case False

      End Select

      GPbrTraitements.Value = GPbrTraitements.Value + 1

      DoEvents

   Next

Exit_VerifieTableInformationConnexion:

   DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

   DoCmd.Hourglass False

   Set TdsTableListe = Nothing

   Set TdfTable = Nothing

   Exit Function

Err_VerifieTableInformationConnexion:

   VerifieTableInformationConnexion = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "VerifieTableInformationConnexion"

   Resume Exit_VerifieTableInformationConnexion
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
Public Function ChangeTableInformationConnexion() As Boolean
   Dim FsoFichierSystem As New Scripting.FileSystemObject
   Dim TdsTableListe As DAO.TableDefs
   Dim TdfTable As DAO.TableDef
   Dim StrConnection As String
   Dim StrType As String
   Dim StrParametre As String
   Dim StrFichier As String
   Dim StrExtension As String
   Dim StrChemin As String
   Dim StrCheminDefault As String

   On Error GoTo Err_ChangeTableInformationConnexion

   Set TdsTableListe = CurrentDb.TableDefs

   ChangeTableInformationConnexion = True

   DoCmd.Hourglass True

   DoCmd.OpenForm "FrmGestionTraitements", acNormal, , , acFormEdit, acWindowNormal

   GEtiTitTraitements.Caption = "Chagement Des Fichiers Des Tables Attachées"

   GEtiFonTraitements.Caption = "ChangeTableInformationConnexion"

   GPbrTraitements.Min = 0

   GPbrTraitements.Max = TdsTableListe.Count

   GEtiObjTraitements.Caption = "Nombre De Tables : " & TdsTableListe.Count

   DoEvents

   For Each TdfTable In TdsTableListe

      Select Case LectureTableInformationConnexion(TdfTable, StrType, StrParametre, StrChemin, StrFichier)
         Case True

            StrExtension = UCase(Right(StrFichier, Len(StrFichier) - InStrRev(StrFichier, ".")))

            StrCheminDefault = Nz(DLookup("FicValeur", "TBLFICHIERS", "FicType='FIC" & StrExtension & "' AND FicCode LIKE '*=" & StrFichier & "' AND FicValide=True"), vbNullString)

            Select Case StrCheminDefault
               Case StrChemin

               Case Else

                  Select Case FsoFichierSystem.FileExists(StrCheminDefault & "\" & StrFichier)
                     Case True

                        SauveTableInformationConnexion TdfTable, StrType, StrParametre, StrCheminDefault, StrFichier

                        SaveSetting CurrentDb.Name, "TABLECONNECTION", TdfTable.Name, TdfTable.Connect

                     Case False

                        ChangeTableInformationConnexion = False

                  End Select

            End Select
         Case False

      End Select

      GPbrTraitements.Value = GPbrTraitements.Value + 1

   Next

   DoCmd.Close acForm, "FrmGestionTraitements", acSavePrompt

   DoCmd.Hourglass False

Exit_ChangeTableInformationConnexion:

   Set TdsTableListe = Nothing

   Set TdfTable = Nothing

   Exit Function

Err_ChangeTableInformationConnexion:

   ChangeTableInformationConnexion = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "ChangeTableInformationConnexion"

   Resume Exit_ChangeTableInformationConnexion
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
Public Function VerifieFichiers() As Boolean
   Dim FsoFichierSystem As New Scripting.FileSystemObject
   Dim RstFichiersDetailler As DAO.Recordset
   Dim StrSQLFichiersDetailler As String
   Dim StrSQLFichiers As String

   On Error GoTo Err_VerifieFichiers

   StrSQLFichiersDetailler = "SELECT SelFichiersDetailler.* FROM SelFichiersDetailler WHERE FicValide=True;"

   Set RstFichiersDetailler = CurrentDb.OpenRecordset(StrSQLFichiersDetailler)

   Do Until RstFichiersDetailler.EOF = True

      Select Case FsoFichierSystem.FileExists(RstFichiersDetailler!FicValeur & "\" & RstFichiersDetailler!FicCode2)
        Case True

        Case False

            StrSQLFichiers = "UPDATE TBLFICHIERS SET FicValide=False WHERE FicCode LIKE '*=" _
               & RstFichiersDetailler!FicCode2 & "' AND FicValeur='" & RstFichiersDetailler!FicValeur & "';"

            CurrentDb.Execute StrSQLFichiers

      End Select

      RstFichiersDetailler.MoveNext

   Loop

   VerifieFichiers = VerifieTableInformationConnexion

   Select Case VerifieFichiers
      Case False

         DoCmd.OpenForm "FrmGestionFichiers", acNormal, , , acFormEdit, acDialog

         VerifieFichiers = VerifieTableInformationConnexion

      Case Else

   End Select

Exit_VerifieFichiers:

   Exit Function

Err_VerifieFichiers:

   VerifieFichiers = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, "VerifieFichiers"

   Resume Exit_VerifieFichiers
End Function
