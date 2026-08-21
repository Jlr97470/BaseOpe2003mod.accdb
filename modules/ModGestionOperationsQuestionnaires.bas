Attribute VB_Name = "ModGestionOperationsQuestionnaires"
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
'***    Declaration De Procedure Public                                                       ***
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
Public Sub CreeAnalysePersonnelsSocietesQuestionnaires(ByVal IntQuestionnairesNumero As Integer)
   Dim StrAnalyseQuestionnaires As String
   Dim StrSQLAnalyseQuestionnairesDebut As String
   Dim StrSQLAnalyseQuestionnairesWhere As String
   Dim StrSQLAnalyseQuestionnairesFin As String
   Dim StrSQLAnalyseQuestionnairesChamp As String
   Dim StrSQLQuestionnairesTypeID As String
   Dim RsQuestionnairesTypesID As DAO.Recordset
   Dim QdfTable As DAO.QueryDef

   On Error GoTo Err_CreeAnalyseQuestionnaires

   StrSQLQuestionnairesTypeID = "SELECT * FROM TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPESID WHERE PerSocQstTypIDPerSocQstTypNum=" & IntQuestionnairesNumero & ";"

   Set RsQuestionnairesTypesID = CurrentDb.OpenRecordset(StrSQLQuestionnairesTypeID)

   Select Case RsQuestionnairesTypesID.EOF
      Case True

      Case False

         StrAnalyseQuestionnaires = "AnaPersonnelsSocietesQst" & IntQuestionnairesNumero

         StrSQLAnalyseQuestionnairesDebut = "TRANSFORM First(PerSocQstResValeur) AS PerSocQstResValeurValeur " & _
               " SELECT PerSocQstNum, PerSocQstPerSocQstTypNum, PerSocQstPerSocNum, PerSocQstDateSaisie, PerSocQstNotes " & _
               " FROM ((TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPES INNER JOIN TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPESID ON TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPES.PerSocQstTypNum = TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPESID.PerSocQstTypIDPerSocQstTypNum) INNER JOIN TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS ON TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPESID.PerSocQstTypIDNum = TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.PerSocQstResPerSocQstTypIDNum) INNER JOIN TBLPERSONNELSSOCIETESQUESTIONNAIRES ON (TBLPERSONNELSSOCIETESQUESTIONNAIRESRESULTATS.PerSocQstResPerSocNum = TBLPERSONNELSSOCIETESQUESTIONNAIRES.PerSocQstPerSocNum) AND (TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPES.PerSocQstTypNum = TBLPERSONNELSSOCIETESQUESTIONNAIRES.PerSocQstPerSocQstTypNum) "

         StrSQLAnalyseQuestionnairesWhere = " WHERE TBLPERSONNELSSOCIETESQUESTIONNAIRESTYPES.PerSocQstTypNum = " & IntQuestionnairesNumero

         StrSQLAnalyseQuestionnairesFin = " GROUP BY PerSocQstNum, PerSocQstPerSocQstTypNum, PerSocQstPerSocNum, PerSocQstDateSaisie, PerSocQstNotes, PerSocQstTypCode, PerSocQstTypLibelle, PerSocQstTypDateDebut, PerSocQstTypDateFin, PerSocQstTypNbResultat PIVOT 'PerSocQstResValeur' & [PerSocQstTypID] In "

         StrSQLAnalyseQuestionnairesChamp = "("

         Do Until RsQuestionnairesTypesID.EOF = True

            StrSQLAnalyseQuestionnairesChamp = StrSQLAnalyseQuestionnairesChamp & "'PerSocQstResValeur" & RsQuestionnairesTypesID!PerSocQstTypID & "',"

            RsQuestionnairesTypesID.MoveNext

         Loop

         StrSQLAnalyseQuestionnairesChamp = Left(StrSQLAnalyseQuestionnairesChamp, Len(StrSQLAnalyseQuestionnairesChamp) - 1) & ")"

         On Error Resume Next

         Set QdfTable = CurrentDb.QueryDefs(StrAnalyseQuestionnaires)

         Select Case Err.Number
            Case 0

               QdfTable.SQL = StrSQLAnalyseQuestionnairesDebut & StrSQLAnalyseQuestionnairesWhere & StrSQLAnalyseQuestionnairesFin & StrSQLAnalyseQuestionnairesChamp & ";"

            Case Else

               Err.Clear

               Set QdfTable = CurrentDb.CreateQueryDef(StrAnalyseQuestionnaires, StrSQLAnalyseQuestionnairesDebut & StrSQLAnalyseQuestionnairesWhere & StrSQLAnalyseQuestionnairesFin & StrSQLAnalyseQuestionnairesChamp & ";")

               CurrentDb.QueryDefs.Append QdfTable

         End Select

         On Error GoTo Err_CreeAnalyseQuestionnaires

      End Select

Exit_CreeAnalyseQuestionnaires:

   Set QdfTable = Nothing

   Set RsQuestionnairesTypesID = Nothing

   Exit Sub

Err_CreeAnalyseQuestionnaires:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeAnalyseQuestionnaires"

   Resume Exit_CreeAnalyseQuestionnaires
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
Public Sub CreeAnalysePersonnelsSocietesOperations(ByVal IntOperationsNumero As Integer)
   Dim StrAnalyseOperations As String
   Dim StrSQLAnalyseOperationsDebut As String
   Dim StrSQLAnalyseOperationsWhere As String
   Dim StrSQLAnalyseOperationsFin As String
   Dim StrSQLAnalyseOperationsChamp As String
   Dim StrSQLOperationsTypeID As String
   Dim RsOperationsTypesID As DAO.Recordset
   Dim QdfTable As DAO.QueryDef

   On Error GoTo Err_CreeAnalyseOperations

   StrSQLOperationsTypeID = "SELECT * FROM TBLPERSONNELSSOCIETESOPERATIONSTYPESID WHERE PerSocOpeTypIDPerSocOpeTypNum=" & IntOperationsNumero & ";"

   Set RsOperationsTypesID = CurrentDb.OpenRecordset(StrSQLOperationsTypeID)

   Select Case RsOperationsTypesID.EOF
      Case True

      Case False

         StrAnalyseOperations = "AnaPersonnelsSocietesOpe" & IntOperationsNumero

         StrSQLAnalyseOperationsDebut = "TRANSFORM First(PerSocOpeResValeur) AS PerSocOpeResValeurValeur " & _
               " SELECT PerSocOpeNum, PerSocOpePerSocOpeTypNum, PerSocOpePerSocNum, PerSocOpeDateSaisie, PerSocOpeNotes " & _
               " FROM ((TBLPERSONNELSSOCIETESOPERATIONSTYPES INNER JOIN TBLPERSONNELSSOCIETESOPERATIONSTYPESID ON TBLPERSONNELSSOCIETESOPERATIONSTYPES.PerSocOpeTypNum = TBLPERSONNELSSOCIETESOPERATIONSTYPESID.PerSocOpeTypIDPerSocOpeTypNum) INNER JOIN TBLPERSONNELSSOCIETESOPERATIONSRESULTATS ON TBLPERSONNELSSOCIETESOPERATIONSTYPESID.PerSocOpeTypIDNum = TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.PerSocOpeResPerSocOpeTypIDNum) INNER JOIN TBLPERSONNELSSOCIETESOPERATIONS ON (TBLPERSONNELSSOCIETESOPERATIONSRESULTATS.PerSocOpeResPerSocNum = TBLPERSONNELSSOCIETESOPERATIONS.PerSocOpePerSocNum) AND (TBLPERSONNELSSOCIETESOPERATIONSTYPES.PerSocOpeTypNum = TBLPERSONNELSSOCIETESOPERATIONS.PerSocOpePerSocOpeTypNum) "

         StrSQLAnalyseOperationsWhere = " WHERE TBLPERSONNELSSOCIETESOPERATIONSTYPES.PerSocOpeTypNum = " & IntOperationsNumero

         StrSQLAnalyseOperationsFin = " GROUP BY PerSocOpeNum, PerSocOpePerSocOpeTypNum, PerSocOpePerSocNum, PerSocOpeDateSaisie, PerSocOpeNotes, PerSocOpeTypCode, PerSocOpeTypLibelle, PerSocOpeTypDateDebut, PerSocOpeTypDateFin, PerSocOpeTypNbResultat PIVOT 'PerSocOpeResValeur' & [PerSocOpeTypID] In "

         StrSQLAnalyseOperationsChamp = "("

         Do Until RsOperationsTypesID.EOF = True

            StrSQLAnalyseOperationsChamp = StrSQLAnalyseOperationsChamp & "'PerSocOpeResValeur" & RsOperationsTypesID!PerSocOpeTypID & "',"

            RsOperationsTypesID.MoveNext

         Loop

         StrSQLAnalyseOperationsChamp = Left(StrSQLAnalyseOperationsChamp, Len(StrSQLAnalyseOperationsChamp) - 1) & ")"

         On Error Resume Next

         Set QdfTable = CurrentDb.QueryDefs(StrAnalyseOperations)

         Select Case Err.Number
            Case 0

               QdfTable.SQL = StrSQLAnalyseOperationsDebut & StrSQLAnalyseOperationsWhere & StrSQLAnalyseOperationsFin & StrSQLAnalyseOperationsChamp & ";"

            Case Else

               Err.Clear

               Set QdfTable = CurrentDb.CreateQueryDef(StrAnalyseOperations, StrSQLAnalyseOperationsDebut & StrSQLAnalyseOperationsWhere & StrSQLAnalyseOperationsFin & StrSQLAnalyseOperationsChamp & ";")

               CurrentDb.QueryDefs.Append QdfTable

         End Select

         On Error GoTo Err_CreeAnalyseOperations

      End Select

Exit_CreeAnalyseOperations:

   Set QdfTable = Nothing

   Set RsOperationsTypesID = Nothing

   Exit Sub

Err_CreeAnalyseOperations:

   MsgBox Err.Number & " " & Err.Description, vbCritical, "CreeAnalyseOperations"

   Resume Exit_CreeAnalyseOperations
End Sub
