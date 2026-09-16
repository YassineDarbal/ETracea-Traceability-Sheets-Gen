Attribute VB_Name = "modTG_Helpers"
Option Explicit

'===============================================================================
' OUTILS DE VALIDATION ET DE TRAITEMENT
'===============================================================================

'-------------------------------------------------------------------------------
' Vérifie que toutes les feuilles nécessaires existent dans le classeur modèle.
'-------------------------------------------------------------------------------
Public Sub ValidateTemplateWorkbook(ByVal wb As Workbook)

    Dim requiredSheets As Variant
    Dim sheetName As Variant
    Dim missingSheets As String

    requiredSheets = Array( _
        TG_DASHBOARD_SHEET, _
        TG_SOURCE_SHEET, _
        TG_SCREEN_VIEW_SHEET, _
        TG_INFORMATION_SHEET, _
        TG_TRACEABILITY_SHEET, _
        TG_HISTORY_SHEET)

    For Each sheetName In requiredSheets

        If Not WorksheetExists(CStr(sheetName), wb) Then
            missingSheets = missingSheets & vbCrLf & _
                            "- " & CStr(sheetName)
        End If

    Next sheetName

    If Len(missingSheets) > 0 Then

        Err.Raise _
            Number:=vbObjectError + 2000, _
            Source:="ValidateTemplateWorkbook", _
            Description:= _
                "Les feuilles suivantes sont absentes du classeur modèle :" & _
                missingSheets

    End If

    If TG_LAST_DATA_COLUMN < TG_FIRST_DATA_COLUMN Then

        Err.Raise _
            Number:=vbObjectError + 2001, _
            Source:="ValidateTemplateWorkbook", _
            Description:= _
                "La plage de colonnes source configurée est invalide."

    End If

End Sub

'-------------------------------------------------------------------------------
' Lit les paramètres de génération depuis les cellules nommées du Dashboard.
'
' Les deux noms doivent être définis avec une portée "Classeur" :
' - TG_FileName
' - TG_OutputFolder
'-------------------------------------------------------------------------------
Public Sub GetGenerationSettings( _
    ByVal wb As Workbook, _
    ByRef fileNameBase As String, _
    ByRef outputFolder As String _
)

    fileNameBase = GetRequiredNamedCellText( _
        wb:=wb, _
        definedName:=TG_FILE_NAME_RANGE, _
        expectedSheetName:=TG_DASHBOARD_SHEET, _
        settingLabel:="Nom de base des fichiers" _
    )

    outputFolder = GetRequiredNamedCellText( _
        wb:=wb, _
        definedName:=TG_OUTPUT_FOLDER_RANGE, _
        expectedSheetName:=TG_DASHBOARD_SHEET, _
        settingLabel:="Dossier de destination" _
    )

End Sub

'-------------------------------------------------------------------------------
' Retourne le texte d'une cellule nommée obligatoire.
'
' La cellule doit :
' - exister comme nom défini au niveau du classeur ;
' - pointer vers une seule cellule ;
' - appartenir à la feuille attendue ;
' - contenir une valeur non vide et sans erreur Excel.
'-------------------------------------------------------------------------------
Public Function GetRequiredNamedCellText( _
    ByVal wb As Workbook, _
    ByVal definedName As String, _
    ByVal expectedSheetName As String, _
    ByVal settingLabel As String _
) As String

    Dim workbookName As Name
    Dim targetCell As Range
    Dim textValue As String

    On Error Resume Next
    Set workbookName = wb.Names(definedName)
    On Error GoTo 0

    If workbookName Is Nothing Then

        Err.Raise _
            Number:=vbObjectError + 2050, _
            Source:="GetRequiredNamedCellText", _
            Description:= _
                "Le nom défini '" & definedName & _
                "' est introuvable." & vbCrLf & vbCrLf & _
                "Créez-le avec une portée 'Classeur' et associez-le " & _
                "à la cellule correspondante de la feuille '" & _
                expectedSheetName & "'."

    End If

    On Error Resume Next
    Set targetCell = workbookName.RefersToRange
    On Error GoTo 0

    If targetCell Is Nothing Then

        Err.Raise _
            Number:=vbObjectError + 2051, _
            Source:="GetRequiredNamedCellText", _
            Description:= _
                "Le nom défini '" & definedName & _
                "' ne référence pas une cellule valide."

    End If

    If targetCell.Cells.CountLarge <> 1 Then

        Err.Raise _
            Number:=vbObjectError + 2052, _
            Source:="GetRequiredNamedCellText", _
            Description:= _
                "Le nom défini '" & definedName & _
                "' doit référencer une seule cellule."

    End If

    If StrComp( _
        targetCell.Worksheet.Name, _
        expectedSheetName, _
        vbBinaryCompare _
    ) <> 0 Then

        Err.Raise _
            Number:=vbObjectError + 2053, _
            Source:="GetRequiredNamedCellText", _
            Description:= _
                "Le nom défini '" & definedName & _
                "' doit pointer vers la feuille '" & _
                expectedSheetName & "'."

    End If

    If IsError(targetCell.Value) Then

        Err.Raise _
            Number:=vbObjectError + 2054, _
            Source:="GetRequiredNamedCellText", _
            Description:= _
                "Le paramètre '" & settingLabel & _
                "' contient une erreur Excel."

    End If

    textValue = Trim$(CStr(targetCell.Value2))

    If Len(textValue) = 0 Then

        Err.Raise _
            Number:=vbObjectError + 2055, _
            Source:="GetRequiredNamedCellText", _
            Description:= _
                "Le paramètre '" & settingLabel & _
                "' est vide dans la feuille '" & _
                expectedSheetName & "'."

    End If

    GetRequiredNamedCellText = textValue

End Function

'-------------------------------------------------------------------------------
' Vérifie si une feuille existe dans un classeur donné.
'-------------------------------------------------------------------------------
Public Function WorksheetExists( _
    ByVal sheetName As String, _
    ByVal wb As Workbook _
) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    WorksheetExists = Not ws Is Nothing

End Function

'-------------------------------------------------------------------------------
' Analyse la feuille OFs et regroupe les numéros de lignes par OF.
'
' Exemple :
'
' OF 000013827960
'     Ligne 3
'     Ligne 4
'
' OF 000013827961
'     Ligne 5
'
' orderedOFs conserve l'ordre d'apparition des OFs dans la feuille.
'-------------------------------------------------------------------------------
Public Sub BuildOFIndex( _
    ByVal wsSource As Worksheet, _
    ByRef rowsByOF As Object, _
    ByRef orderedOFs As Collection _
)

    Dim lastRow As Long
    Dim sourceRow As Long

    Dim ofNumber As String
    Dim rowList As Collection

    Dim originalDescription As String

    Set rowsByOF = CreateObject("Scripting.Dictionary")
    rowsByOF.CompareMode = vbBinaryCompare

    Set orderedOFs = New Collection

    lastRow = wsSource.Cells( _
        wsSource.Rows.Count, _
        TG_OF_COLUMN _
    ).End(xlUp).Row

    If lastRow < TG_FIRST_SOURCE_ROW Then
        Exit Sub
    End If

    For sourceRow = TG_FIRST_SOURCE_ROW To lastRow

        If IsError(wsSource.Cells(sourceRow, TG_OF_COLUMN).Value) Then

            Err.Raise _
                Number:=vbObjectError + 2002, _
                Source:="BuildOFIndex", _
                Description:= _
                    "La cellule A" & sourceRow & _
                    " contient une erreur Excel."

        End If

        If Len(Trim$(CStr( _
            wsSource.Cells(sourceRow, TG_OF_COLUMN).Value2 _
        ))) > 0 Then

            On Error GoTo InvalidOF

            ofNumber = GetFormattedOF( _
                wsSource.Cells(sourceRow, TG_OF_COLUMN) _
            )

            On Error GoTo 0

            If Not rowsByOF.Exists(ofNumber) Then

                Set rowList = New Collection

                rowsByOF.Add ofNumber, rowList
                orderedOFs.Add ofNumber

            Else

                Set rowList = rowsByOF.Item(ofNumber)

            End If

            rowList.Add sourceRow

        End If

    Next sourceRow

    Exit Sub

InvalidOF:

    originalDescription = Err.Description

    On Error GoTo 0

    Err.Raise _
        Number:=vbObjectError + 2003, _
        Source:="BuildOFIndex", _
        Description:= _
            "OF invalide à la ligne " & sourceRow & "." & _
            vbCrLf & originalDescription

End Sub

'-------------------------------------------------------------------------------
' Transforme un OF en texte de 12 chiffres.
'
' Exemple :
' 13827960 devient 000013827960
'
' Cette fonction accepte :
' - un OF stocké comme nombre ;
' - un OF stocké comme texte ;
' - un OF contenant déjà des zéros à gauche.
'-------------------------------------------------------------------------------
Public Function GetFormattedOF(ByVal ofCell As Range) As String

    Dim rawValue As Variant
    Dim cleanedText As String

    Dim numericValue As Double

    Dim characterIndex As Long
    Dim currentCharacter As String

    rawValue = ofCell.Value2

    If IsError(rawValue) Then

        Err.Raise _
            Number:=vbObjectError + 2010, _
            Source:="GetFormattedOF", _
            Description:= _
                "La cellule contient une erreur Excel."

    End If

    cleanedText = Trim$(CStr(rawValue))

    If Len(cleanedText) = 0 Then
        Exit Function
    End If

    '---------------------------------------------------------------------------
    ' Cas où Excel stocke l'OF comme nombre.
    '---------------------------------------------------------------------------
    If IsNumeric(rawValue) Then

        numericValue = CDbl(rawValue)

        If numericValue < 0 Or numericValue > 999999999999# Then

            Err.Raise _
                Number:=vbObjectError + 2011, _
                Source:="GetFormattedOF", _
                Description:= _
                    "L'OF doit être compris entre 0 et 999999999999."

        End If

        If numericValue <> Fix(numericValue) Then

            Err.Raise _
                Number:=vbObjectError + 2012, _
                Source:="GetFormattedOF", _
                Description:= _
                    "L'OF ne peut pas contenir de décimales."

        End If

        GetFormattedOF = Format$( _
            numericValue, _
            "000000000000" _
        )

        Exit Function

    End If

    '---------------------------------------------------------------------------
    ' Cas où l'OF est stocké sous forme de texte.
    '---------------------------------------------------------------------------

    ' Supprime les espaces normaux et insécables.
    cleanedText = Replace(cleanedText, ChrW(160), vbNullString)
    cleanedText = Replace(cleanedText, " ", vbNullString)

    ' Vérifie que chaque caractère est un chiffre.
    For characterIndex = 1 To Len(cleanedText)

        currentCharacter = Mid$( _
            cleanedText, _
            characterIndex, _
            1 _
        )

        If currentCharacter < "0" Or currentCharacter > "9" Then

            Err.Raise _
                Number:=vbObjectError + 2013, _
                Source:="GetFormattedOF", _
                Description:= _
                    "L'OF doit contenir uniquement des chiffres."

        End If

    Next characterIndex

    If Len(cleanedText) > 12 Then

        Err.Raise _
            Number:=vbObjectError + 2014, _
            Source:="GetFormattedOF", _
            Description:= _
                "L'OF contient plus de 12 chiffres."

    End If

    ' Ajoute automatiquement les zéros nécessaires à gauche.
    GetFormattedOF = Right$( _
        String$(12, "0") & cleanedText, _
        12 _
    )

End Function

'-------------------------------------------------------------------------------
' Crée le dossier de destination s'il n'existe pas.
'
' Les dossiers parents sont également créés si nécessaire.
'-------------------------------------------------------------------------------
Public Sub EnsureFolderExists(ByVal folderPath As String)

    Dim fileSystem As Object
    Dim parentFolder As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    If fileSystem.FolderExists(folderPath) Then
        Exit Sub
    End If

    parentFolder = fileSystem.GetParentFolderName(folderPath)

    If Len(parentFolder) = 0 Then

        Err.Raise _
            Number:=vbObjectError + 2020, _
            Source:="EnsureFolderExists", _
            Description:= _
                "Le chemin de destination est invalide :" & _
                vbCrLf & folderPath

    End If

    If Not fileSystem.FolderExists(parentFolder) Then
        EnsureFolderExists parentFolder
    End If

    fileSystem.CreateFolder folderPath

End Sub

'-------------------------------------------------------------------------------
' Remplace les caractères interdits dans les noms de fichiers Windows.
'-------------------------------------------------------------------------------
Public Function CleanFileName(ByVal fileName As String) As String

    Dim forbiddenCharacters As Variant
    Dim currentCharacter As Variant

    forbiddenCharacters = Array( _
        "\", _
        "/", _
        ":", _
        "*", _
        "?", _
        """", _
        "<", _
        ">", _
        "|" _
    )

    CleanFileName = Trim$(fileName)

    For Each currentCharacter In forbiddenCharacters

        CleanFileName = Replace( _
            CleanFileName, _
            CStr(currentCharacter), _
            "_" _
        )

    Next currentCharacter

    ' Windows n'accepte pas un nom se terminant par un point ou un espace.
    Do While Len(CleanFileName) > 0 And _
             (Right$(CleanFileName, 1) = "." Or _
              Right$(CleanFileName, 1) = " ")

        CleanFileName = Left$( _
            CleanFileName, _
            Len(CleanFileName) - 1 _
        )

    Loop

    If Len(CleanFileName) = 0 Then

        Err.Raise _
            Number:=vbObjectError + 2030, _
            Source:="CleanFileName", _
            Description:= _
                "Le nom du fichier généré est vide ou invalide."

    End If

End Function

'-------------------------------------------------------------------------------
' Assemble un chemin de dossier et un nom de fichier.
'-------------------------------------------------------------------------------
Public Function CombinePath( _
    ByVal folderPath As String, _
    ByVal fileName As String _
) As String

    If Right$(folderPath, 1) = "\" Then

        CombinePath = folderPath & fileName

    Else

        CombinePath = folderPath & "\" & fileName

    End If

End Function

'-------------------------------------------------------------------------------
' Supprime le fichier existant avant de créer la nouvelle version.
'
' Si le fichier est ouvert ou verrouillé, une erreur compréhensible est générée.
'-------------------------------------------------------------------------------
Public Sub DeleteExistingFile(ByVal filePath As String)

    If Len(Dir$( _
        filePath, _
        vbNormal Or vbHidden Or vbSystem Or vbReadOnly _
    )) = 0 Then

        Exit Sub

    End If

    On Error Resume Next
    SetAttr filePath, vbNormal
    On Error GoTo DeleteError

    Kill filePath

    Exit Sub

DeleteError:

    Err.Raise _
        Number:=vbObjectError + 2040, _
        Source:="DeleteExistingFile", _
        Description:= _
            "Impossible de remplacer le fichier existant :" & _
            vbCrLf & filePath & _
            vbCrLf & vbCrLf & _
            "Vérifiez que le fichier n'est pas déjà ouvert."

End Sub

'-------------------------------------------------------------------------------
' Supprime les anciennes données présentes à partir de la ligne 3 dans
' TRACEABILITY.
'
' Les formats de cellules sont conservés.
'-------------------------------------------------------------------------------
Public Sub ClearTraceabilityData( _
    ByVal wsTraceability As Worksheet _
)

    Dim lastUsedCell As Range
    Dim lastDataRow As Long
    Dim exportedColumnCount As Long

    exportedColumnCount = _
        TG_LAST_DATA_COLUMN - TG_FIRST_DATA_COLUMN + 1

    Set lastUsedCell = wsTraceability.Range( _
        wsTraceability.Cells( _
            TG_FIRST_TRACEABILITY_ROW, _
            TG_FIRST_TRACEABILITY_COLUMN _
        ), _
        wsTraceability.Cells( _
            wsTraceability.Rows.Count, _
            TG_FIRST_TRACEABILITY_COLUMN + exportedColumnCount - 1 _
        ) _
    ).Find( _
        What:="*", _
        After:=wsTraceability.Cells( _
            TG_FIRST_TRACEABILITY_ROW, _
            TG_FIRST_TRACEABILITY_COLUMN _
        ), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )

    If lastUsedCell Is Nothing Then
        Exit Sub
    End If

    lastDataRow = lastUsedCell.Row

    If lastDataRow < TG_FIRST_TRACEABILITY_ROW Then
        Exit Sub
    End If

    wsTraceability.Range( _
        wsTraceability.Cells( _
            TG_FIRST_TRACEABILITY_ROW, _
            TG_FIRST_TRACEABILITY_COLUMN _
        ), _
        wsTraceability.Cells( _
            lastDataRow, _
            TG_FIRST_TRACEABILITY_COLUMN + exportedColumnCount - 1 _
        ) _
    ).ClearContents

End Sub

'-------------------------------------------------------------------------------
' Limite la taille des détails affichés dans une MsgBox.
'-------------------------------------------------------------------------------
Public Function ShortenForMessageBox( _
    ByVal textValue As String, _
    Optional ByVal maximumLength As Long = 1200 _
) As String

    If Len(textValue) <= maximumLength Then

        ShortenForMessageBox = textValue

    Else

        ShortenForMessageBox = _
            Left$(textValue, maximumLength) & _
            vbCrLf & _
            "... liste tronquée ..."

    End If

End Function
