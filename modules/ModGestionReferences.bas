Attribute VB_Name = "ModGestionReferences"
''******************************************************************************
''***                                                                           ***
''******************************************************************************
''***    FORM:                                                                                              ***
''******************************************************************************
''***    FONCTION:                                                                                       ***
''******************************************************************************
''***    PROGRAMMEUR:                                                                              ***
''***      Royer Jean-Laurent                                                                         ***
''******************************************************************************
'
''******************************************************************************
''***    MODIF :                                                                                            ***
''******************************************************************************
'Option Explicit
'
''******************************************************************************
''***    Declaration De Constante Priver                                                        ***
''******************************************************************************
'
''******************************************************************************
''***    Constante De Controle Du Type De Referencement                            ***
''******************************************************************************
'Private Const REFERENCEFICHIER = 0
'Private Const REFERENCEGUID = 1
'
''******************************************************************************
''***    Constante De Controle De L'Action A Utiliser                                    ***
''******************************************************************************
'Public Const REFERENCEINITIALISE = 0
'Public Const REFERENCETERMINE = 1
'
''******************************************************************************
''***    Declaration De Fonction Public                                                          ***
''******************************************************************************
'
''******************************************************************************
''***    FONCTION:                                                                                       ***
''******************************************************************************
''***    FONCTION:                                                                                       ***
''******************************************************************************
''***    ENTREE:                                                                                          ***
''***    SORTIE:                                                                                           ***
''******************************************************************************
''***    EXEMPLE:                                                                                        ***
''******************************************************************************
'Public Function LectureVbaReferenceInformation(ByVal StrName As String, ByRef BolEtat As Boolean, ByRef StrCheminFichier As String, ByRef StrGuid As String, ByRef BytMajor As Byte, ByRef BytMinor As Byte) As Boolean
'   Dim RefVbaReference As Reference
'
'   On Error GoTo Err_LectureVbaReferenceInformation
'
'   LectureVbaReferenceInformation = True
'
'   For Each RefVbaReference In References
'
'      Select Case RefVbaReference.Name
'         Case StrName
'
'            BolEtat = RefVbaReference.IsBroken
'
'            StrGuid = RefVbaReference.Guid
'
'            BytMajor = RefVbaReference.Major
'
'            BytMinor = RefVbaReference.Minor
'
'            StrCheminFichier = RefVbaReference.FullPath
'
'            GoTo Exit_LectureVbaReferenceInformation
'
'         Case Else
'
'      End Select
'   Next
'
'   LectureVbaReferenceInformation = False
'
'Exit_LectureVbaReferenceInformation:
'
'   Set RefVbaReference = Nothing
'
'   Exit Function
'
'Err_LectureVbaReferenceInformation:
'
'   LectureVbaReferenceInformation = False
'
'   MsgBox Err.Number & " " & Err.Description, vbCritical, "LectureVbaReferenceInformation"
'
'   Resume Exit_LectureVbaReferenceInformation
'End Function
'
''******************************************************************************
''***    FONCTION:                                                                                       ***
''******************************************************************************
''***    FONCTION:                                                                                       ***
''******************************************************************************
''***    ENTREE:                                                                                          ***
''***    SORTIE:                                                                                           ***
''******************************************************************************
''***    EXEMPLE:                                                                                        ***
''******************************************************************************
'Public Function SauveVbaReferenceInformation(ByVal StrName As String, ByVal BytReferenceType As Byte, ByVal StrReferenceIdentifiant As String, Optional ByVal BytMajor As Byte, Optional ByVal BytMinor As Byte) As Boolean
'   Dim RefVbaReference As Reference
'
'   On Error GoTo Err_SauveVbaReferenceInformation
'
'   SauveVbaReferenceInformation = True
'
'   For Each RefVbaReference In References
'
'      Select Case RefVbaReference.Name
'         Case StrName
'
'            On Error Resume Next
'
'            References.Remove RefVbaReference
'
'            Select Case Err.Number
'               Case vbEmpty
'
'               Case 57101
'
'                  Err.Clear
'
'                  GoTo Exit_SauveVbaReferenceInformation
'
'               Case Else
'
'                  GoTo Err_SauveVbaReferenceInformation
'
'            End Select
'
'            On Error GoTo Err_SauveVbaReferenceInformation
'
'         Case Else
'
'      End Select
'
'   Next
'
'   On Error Resume Next
'
'   Select Case BytReferenceType
'      Case REFERENCEFICHIER
'
'         References.AddFromFile StrReferenceIdentifiant
'
'      Case REFERENCEGUID
'
'         References.AddFromGuid StrReferenceIdentifiant, BytMajor, BytMinor
'
'   End Select
'
'   Select Case Err.Number
'      Case 0
'
'      Case Else
'
'         Err.Clear
'
'         SauveVbaReferenceInformation = False
'
'   End Select
'
'Exit_SauveVbaReferenceInformation:
'
'   Set RefVbaReference = Nothing
'
'   Exit Function
'
'Err_SauveVbaReferenceInformation:
'
'   SauveVbaReferenceInformation = False
'
'   MsgBox Err.Number & " " & Err.Description, vbCritical, "SauveVbaReferenceInformation"
'
'   Resume Exit_SauveVbaReferenceInformation
'End Function
'
''******************************************************************************
''***    FONCTION:                                                                                       ***
''******************************************************************************
''***    FONCTION:                                                                                       ***
''******************************************************************************
''***    ENTREE:                                                                                          ***
''***    SORTIE:                                                                                           ***
''******************************************************************************
''***    EXEMPLE:                                                                                        ***
''******************************************************************************
'Public Function VerifieVbaReferenceInformation(Optional ByVal BytVerifieType As Byte) As Boolean
'   Dim DBDataBase As Object
'   Dim RsParametres As Object
'   Dim RefVbaReference As Access.Reference
'   Dim BolVbaReferenceInscrit As Boolean
'   Dim StrCheminFichierLocal As String
'   Dim StrCheminLocal As String
'   Dim StrFichierLocal As String
'   Dim StrCheminFichierServeur As String
'   Dim StrCheminServeur As String
'   Dim StrFichierServeur As String
'   Dim BolLicencesValide As Boolean
'
'   On Error GoTo Err_VerifieVbaReferenceInformation
'
'   VerifieVbaReferenceInformation = True
'
'   BolLicencesValide = True
'
'   Set DBDataBase = CurrentDb
'
'   StrCheminServeur = Left(DBDataBase.Name, InStrRev(DBDataBase.Name, "\") - 1) & "\DATA\BIN"
'
'   Select Case Dir(StrCheminServeur, vbDirectory)
'      Case vbNullString
'
'         CreeRepertoire StrCheminServeur
'
'      Case Else
'
'   End Select
'
'   Select Case BytVerifieType
'      Case REFERENCEINITIALISE
'
'         Set RsParametres = DBDataBase.OpenRecordset("SELECT * FROM SelParametresDetailler WHERE ParType='VBAREF' ; ")
'
'         Select Case RsParametres.EOF
'            Case True
'
'            Case False
'
'               Do Until RsParametres.EOF = True
'
'                  StrFichierLocal = RsParametres!ParValeur
'
'                  StrFichierServeur = StrFichierLocal
'
'                  StrCheminFichierServeur = StrCheminServeur & "\" & StrFichierLocal
'
'                  Select Case SauveVbaReferenceInformation(RsParametres!ParCode1, REFERENCEGUID, RsParametres!ParCode2, RsParametres!ParCode3, RsParametres!ParCode4)
'                     Case True
'
'                        StrCheminFichierLocal = RsParametres!ParCode5
'
'                        Select Case Dir(StrCheminFichierServeur)
'                           Case vbNullString
'
'                              Select Case Dir(StrCheminFichierLocal)
'                                 Case vbNullString
'
'                                 Case Else
'
'                                    FileCopy StrCheminFichierLocal, StrCheminFichierServeur
'
'                              End Select
'
'
'                           Case Else
'
'                        End Select
'
'                     Case False
'
'                        SauveVbaReferenceInformation RsParametres!ParCode1, REFERENCEFICHIER, StrCheminFichierServeur
'
'                        BolVbaReferenceInscrit = Shell("REGSVR32.EXE " & Chr$(34) & StrCheminFichierServeur & Chr$(34) & " /S", vbHide)
'
'                        BolLicencesValide = False
'
'                  End Select
'
'                  RsParametres.MoveNext
'
'               Loop
'
'         End Select
'
'         RsParametres.Close
'
'         Set RsParametres = Nothing
'
'      Case REFERENCETERMINE
'
'         CurrentDb.Execute "DELETE * FROM TBLPARAMETRES WHERE ParType='VBAREF' ;"
'
'         For Each RefVbaReference In References
'
'            StrCheminFichierLocal = RefVbaReference.FullPath
'
'            StrCheminLocal = Left(StrCheminFichierLocal, InStrRev(StrCheminFichierLocal, "\") - 1)
'
'            StrFichierLocal = Mid(StrCheminFichierLocal, InStrRev(StrCheminFichierLocal, "\") + 1)
'
'            StrFichierServeur = StrFichierLocal
'
'            StrCheminFichierServeur = StrCheminServeur & "\" & StrFichierLocal
'
'            CurrentDb.Execute "INSERT INTO TBLPARAMETRES (ParType,ParCode,ParValeur,ParValide) VALUES ('VBAREF','" & RefVbaReference.Name & "=" & RefVbaReference.Guid & "=" & RefVbaReference.Major & "=" & RefVbaReference.Minor & "=" & RefVbaReference.FullPath & "','" & StrFichierLocal & "',True) ;"
'
'            Select Case Dir(StrCheminFichierLocal)
'               Case vbNullString
'
'                  Select Case Dir(StrCheminLocal, vbDirectory)
'                     Case vbNullString
'
'                        CreeRepertoire StrCheminLocal
'
'                     Case Else
'
'                  End Select
'
'                  FileCopy StrCheminFichierServeur, StrCheminFichierLocal
'
'                  BolVbaReferenceInscrit = Shell("REGSVR32.EXE " & Chr$(34) & StrCheminFichierLocal & Chr$(34) & " /S", vbHide)
'
'                  BolLicencesValide = False
'
'               Case Else
'
'            End Select
'
'            Select Case Dir(StrCheminFichierServeur)
'               Case vbNullString
'
'                     FileCopy StrCheminFichierLocal, StrCheminFichierServeur
'
'               Case Else
'
'            End Select
'
'            On Error Resume Next
'
'            References.Remove RefVbaReference
'
'            Select Case Err.Number
'               Case 0
'
'               Case 57101
'
'                  Err.Clear
'
'               Case Else
'
'                  GoTo Err_VerifieVbaReferenceInformation
'
'            End Select
'
'            On Error GoTo Err_VerifieVbaReferenceInformation
'
'         Next
'
'   End Select
'
'   Select Case BolLicencesValide
'      Case True
'
'      Case False
'
'         Shell "REGEDIT.EXE " & Chr$(34) & StrCheminServeur & "\LICENCES.REG" & Chr$(34)
'
'   End Select
'
'Exit_VerifieVbaReferenceInformation:
'
'   DBDataBase.Close
'
'   Set DBDataBase = Nothing
'
'   Set RefVbaReference = Nothing
'
'   Exit Function
'
'Err_VerifieVbaReferenceInformation:
'
'   VerifieVbaReferenceInformation = False
'
'   MsgBox Err.Number & " " & Err.Description, vbCritical, "VerifieVbaReferenceInformation"
'
'   Resume Exit_VerifieVbaReferenceInformation
'End Function
'
