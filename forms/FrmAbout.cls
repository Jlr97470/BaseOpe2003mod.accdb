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
'***    Variable                                                                                            ***
'******************************************************************************
Private MIntSeconde As Integer

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
    Dim DBDataBase As DAO.Database
    Dim CtrDataBase As DAO.Container
    Dim DocDataBase As DAO.Document
    Dim PrpDocument As DAO.Property                                                                     ' Une proprieter du document

    On Error Resume Next

    MIntSeconde = vbEmpty

    Set DBDataBase = CurrentDb

    Set CtrDataBase = DBDataBase.Containers("Databases")

    Set DocDataBase = CtrDataBase.Documents("MSysDb")

    EtiAccessVersion.Caption = DocDataBase.Properties("AccessVersion").Value

    EtiAccessVersionBuild.Caption = DocDataBase.Properties("Build").Value

    Set DocDataBase = CtrDataBase.Documents("SummaryInfo")

    EtiSocieteVal.Caption = DocDataBase.Properties("Company").Value

    EtiTitreVal.Caption = DocDataBase.Properties("Title").Value

    EtiSujetVal.Caption = DocDataBase.Properties("Subject").Value

    EtiAutherVal.Caption = DocDataBase.Properties("Author").Value

    EtiDirecteurVal.Caption = DocDataBase.Properties("Manager").Value

    EtiUtilisateurVal.Caption = DocDataBase.Properties("UserName").Value

    For Each CtrDataBase In DBDataBase.Containers

        Select Case CtrDataBase.Name
            Case "Databases", "SysRel"

            Case Else

                For Each DocDataBase In CtrDataBase.Documents

                    Select Case Left(DocDataBase.Name, 3)
                        Case "~sq", "MSy"

                        Case Else

                            Select Case DocDataBase.Properties("DateCreated").Value
                                Case Is < CDate(EtiDateCreationVal.Caption)

                                    EtiDateCreationVal.Caption = DocDataBase.Properties("DateCreated").Value

                                    EtiDateCreationNom.Caption = DocDataBase.Name

                                Case Else

                            End Select

                            Select Case DocDataBase.Properties("LastUpdated").Value
                                Case Is > CDate(EtiDateModificationVal.Caption)

                                    EtiDateModificationVal.Caption = DocDataBase.Properties("LastUpdated").Value

                                    EtiDateModificationNom.Caption = DocDataBase.Name

                                Case Else

                            End Select
                    End Select
                Next
        End Select
    Next

   EtiUtilisateurNom.Caption = GStrUtilisateurNom

   EtiOrdinateurNom.Caption = GStrOrdinateurNom

   EtiOrdinateurDomaine.Caption = GStrOrdinateurDomaine

   GStrAccessVersion = EtiAccessVersion.Caption

   GStrAccessVersionBuild = EtiAccessVersionBuild.Caption

   DBDataBase.Properties("AppIcon") = Left(DBDataBase.Name, InStrRev(DBDataBase.Name, "\") - 1) & "\DATA\ICONES\LOGO.ICO"

   DBDataBase.Close

Exit_Form_Open:

    Set DocDataBase = Nothing

    Set CtrDataBase = Nothing

    Set DBDataBase = Nothing

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
Private Sub Form_Timer()

   On Error GoTo Err_Form_Timer

   MIntSeconde = MIntSeconde + 1

   Select Case MIntSeconde
      Case Is >= 1

         DoCmd.Close acForm, Me.Name, acSavePrompt

      Case Else

   End Select

Exit_Form_Timer:

    Exit Sub

Err_Form_Timer:

    MsgBox Err.Number & " " & Err.Description, vbCritical, "Form_Timer"

    Resume Exit_Form_Timer
End Sub
