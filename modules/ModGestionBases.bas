Attribute VB_Name = "ModGestionBases"
Option Compare Database
Option Explicit
' ##MODULE_SUMMARY Module de gestion de l'aide contextuel de la base.
' ##PROJECT_TITLE DeltaInformatique (Frontale)
' ##PROJECT_SUMMARY Interface DeltaInformatique
Private Const ModName = "ModGestionBases"

Public Const TYPEMARKETING = "RSM"
Public Const TYPEDIRECTEURVENTE = "DVR"
Public Const TYPERESPONSABLEVENTE = "RVR"
Public Const TYPERESPONSABLE = "RES"
Public Const TYPEFORMATEUR = "FOR"
Public Const TYPETECHNICO = "TEC"
Public Const TYPEAMBASSADRICE = "AMB"
Public Const TYPEVENDEUR = "VND"

Public Const TYPESIEGE = 1
Public Const TYPEPDVDEPENDANT = 2
Public Const TYPEPDVINDEPENDANT = 3

Public GStrAccessVersion As String
Public GStrAccessVersionBuild As String
' ##SUMMARY Fonction qui affiche l'aide contextuel du control actif.
' ##RETURNS Retourne vrai si l'aide s'est bien afficher.
' ##PARAM
Public Function AfficheAide() As Boolean
'   Dim FrmDataBase As Access.Form
'   Dim CtlDataBase As Access.Control
'   Dim CdgBoite As MSComDlg.CommonDialog
'
'   On Error GoTo Exit_AfficheAide
'
'   AfficheAide = True
'
'   Set CdgBoite = Forms!FrmMenuPrincipale.CdgBoite
'
'   Select Case IsObject(Screen.ActiveControl)
'      Case True
'
'         Set CtlDataBase = Screen.ActiveControl
'
'         CdgBoite.HelpKey = CtlDataBase.HelpContextId
'
'         CdgBoite.HelpFile = CtlDataBase.Parent.HelpFile
'
'         CdgBoite.HelpCommand = cdlHelpContextPopup
'
'         CdgBoite.ShowHelp
'
'      Case False
'
'         Select Case IsObject(Screen.ActiveForm)
'            Case True
'
'               Set FrmDataBase = Screen.ActiveForm
'
'               CdgBoite.HelpKey = CtlDataBase.HelpContextId
'
'               CdgBoite.HelpFile = CtlDataBase.Parent.HelpFile
'
'               CdgBoite.HelpCommand = cdlHelpContextPopup
'
'               CdgBoite.ShowHelp
'
'            Case False
'
'         End Select
'
'   End Select

Exit_AfficheAide:

'   Set CdgBoite = Nothing

   Exit Function

Err_AfficheAide:

   AfficheAide = False

   MsgBox Err.Number & " " & Err.Description, vbCritical, ModName & ".AfficheAide()"

   Resume Exit_AfficheAide:
End Function
