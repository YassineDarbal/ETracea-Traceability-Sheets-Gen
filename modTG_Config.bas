Attribute VB_Name = "modTG_Config"
Option Explicit

'===============================================================================
' CONFIGURATION CENTRALE
'
' Modifiez uniquement ce module si les noms des feuilles, les cellules,
' les colonnes ou le chemin de destination changent.
'===============================================================================

'-------------------------------------------------------------------------------
' Feuilles du classeur source / modèle
'-------------------------------------------------------------------------------

Public Const TG_DASHBOARD_SHEET As String = "Dashboard"
Public Const TG_SOURCE_SHEET As String = "OFs"

Public Const TG_SCREEN_VIEW_SHEET As String = "SCREEN VIEW"
Public Const TG_INFORMATION_SHEET As String = "INFORMATIONS"
Public Const TG_TRACEABILITY_SHEET As String = "TRACEABILITY"
Public Const TG_HISTORY_SHEET As String = "HISTORY"

'-------------------------------------------------------------------------------
' Structure de la feuille OFs
'-------------------------------------------------------------------------------

' Première ligne contenant les données
Public Const TG_FIRST_SOURCE_ROW As Long = 3

' Colonne contenant le numéro d'OF
Public Const TG_OF_COLUMN As Long = 1

' Première colonne à exporter : colonne B
Public Const TG_FIRST_DATA_COLUMN As Long = 2

' Dernière colonne à exporter : colonne AC
Public Const TG_LAST_DATA_COLUMN As Long = 29

'-------------------------------------------------------------------------------
' Destination dans la feuille TRACEABILITY
'-------------------------------------------------------------------------------

' Première ligne d'insertion
Public Const TG_FIRST_TRACEABILITY_ROW As Long = 3

' Première colonne d'insertion : colonne A
Public Const TG_FIRST_TRACEABILITY_COLUMN As Long = 1

'-------------------------------------------------------------------------------
' Cellules de destination dans la feuille INFORMATIONS
'-------------------------------------------------------------------------------

Public Const TG_INFORMATION_TITLE_CELL As String = "B1"
Public Const TG_INFORMATION_OF_CELL As String = "B3"

'-------------------------------------------------------------------------------
' Configuration dynamique depuis la feuille Dashboard
'-------------------------------------------------------------------------------

' Noms définis Excel (portée : classeur)
' TG_FileName      = nom de base du fichier, sans l'OF
' TG_OutputFolder  = dossier de génération
Public Const TG_FILE_NAME_RANGE As String = "TG_FileName"
Public Const TG_OUTPUT_FOLDER_RANGE As String = "TG_OutputFolder"

'-------------------------------------------------------------------------------
' Format de sortie
'-------------------------------------------------------------------------------

Public Const TG_OUTPUT_EXTENSION As String = ".xlsx"
