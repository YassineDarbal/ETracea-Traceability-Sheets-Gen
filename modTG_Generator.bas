Attribute VB_Name = "modTG_Generator"
Option Explicit

'===============================================================================
' GÉNÉRATEUR PRINCIPAL
'===============================================================================

'-------------------------------------------------------------------------------
' Procédure principale.
'
' Pour chaque OF unique :
' 1. Copie les quatre feuilles du modèle.
' 2. Copie les lignes correspondantes de OFs!B:AC.
' 3. Colle les données dans TRACEABILITY!A3:AB...
' 4. Inscrit l'OF dans INFORMATIONS!B3.
' 5. Inscrit le titre dans INFORMATIONS!B1.
' 6. Enregistre le fichier au format XLSX.
'-------------------------------------------------------------------------------
Public Sub GenerateTraceabilityFiles()

    Dim wbTemplate As Workbook
    Dim wsSource As Worksheet

    Dim rowsByOF As Object
    Dim orderedOFs As Collection
    Dim sourceRows As Collection

    Dim ofNumber As String
    Dim fileNameBase As String
    Dim outputFolder As String
    Dim savedFilePath As String
    Dim GenerationError As String
    Dim failureDetails As String

    Dim generatedCount As Long
    Dim failedCount As Long
    Dim ofIndex As Long

    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousDisplayAlerts As Boolean

    Dim applicationSettingsChanged As Boolean

    On Error GoTo FatalError

    Set wbTemplate = ThisWorkbook

    ' Vérifie la structure du classeur modèle.
    ValidateTemplateWorkbook wbTemplate

    ' Lit la configuration définie dans le Dashboard.
    GetGenerationSettings _
        wb:=wbTemplate, _
        fileNameBase:=fileNameBase, _
        outputFolder:=outputFolder

    ' Crée le dossier de destination s'il n'existe pas.
    EnsureFolderExists outputFolder

    Set wsSource = wbTemplate.Worksheets(TG_SOURCE_SHEET)

    ' Regroupe les lignes appartenant au même OF.
    BuildOFIndex _
        wsSource:=wsSource, _
        rowsByOF:=rowsByOF, _
        orderedOFs:=orderedOFs

    If orderedOFs.Count = 0 Then

        MsgBox _
            Prompt:= _
                "Aucun OF n'a été trouvé dans la colonne A de la feuille '" & _
                TG_SOURCE_SHEET & "', à partir de la ligne " & _
                TG_FIRST_SOURCE_ROW & ".", _
            Buttons:=vbExclamation, _
            Title:="Génération annulée"

        Exit Sub

    End If

    ' Sauvegarde les paramètres actuels d'Excel.
    previousCalculation = Application.Calculation
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousDisplayAlerts = Application.DisplayAlerts

    ' Désactive temporairement certaines fonctions pour accélérer le traitement.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    applicationSettingsChanged = True

    For ofIndex = 1 To orderedOFs.Count

        ofNumber = CStr(orderedOFs.Item(ofIndex))

        Set sourceRows = rowsByOF.Item(ofNumber)

        Application.StatusBar = _
            "Génération du fichier " & _
            ofIndex & " / " & orderedOFs.Count & _
            " - OF " & ofNumber

        savedFilePath = vbNullString
        GenerationError = vbNullString

        If GenerateOneOFFile( _
            wbTemplate:=wbTemplate, _
            wsSource:=wsSource, _
            ofNumber:=ofNumber, _
            sourceRows:=sourceRows, _
            fileNameBase:=fileNameBase, _
            outputFolder:=outputFolder, _
            savedFilePath:=savedFilePath, _
            errorMessage:=GenerationError _
        ) Then

            generatedCount = generatedCount + 1

        Else

            failedCount = failedCount + 1

            failureDetails = _
                failureDetails & _
                vbCrLf & _
                "- OF " & ofNumber & _
                " : " & GenerationError

        End If

    Next ofIndex

CleanExit:

    Application.StatusBar = False

    ' Restaure les paramètres Excel.
    If applicationSettingsChanged Then

        Application.Calculation = previousCalculation
        Application.DisplayAlerts = previousDisplayAlerts
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating

    End If

    If failedCount = 0 Then

        MsgBox _
            Prompt:= _
                generatedCount & _
                " fichier(s) généré(s) avec succès." & _
                vbCrLf & vbCrLf & _
                "Dossier de destination :" & _
                vbCrLf & outputFolder, _
            Buttons:=vbInformation, _
            Title:="Génération terminée"

    Else

        MsgBox _
            Prompt:= _
                generatedCount & _
                " fichier(s) généré(s)." & _
                vbCrLf & _
                failedCount & _
                " fichier(s) en erreur." & _
                vbCrLf & vbCrLf & _
                "Détails :" & _
                ShortenForMessageBox(failureDetails), _
            Buttons:=vbExclamation, _
            Title:="Génération terminée avec erreurs"

    End If

    Exit Sub

FatalError:

    GenerationError = Err.Description

    Application.StatusBar = False

    If applicationSettingsChanged Then

        On Error Resume Next

        Application.Calculation = previousCalculation
        Application.DisplayAlerts = previousDisplayAlerts
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating

        On Error GoTo 0

    End If

    MsgBox _
        Prompt:= _
            "La génération a été interrompue." & _
            vbCrLf & vbCrLf & _
            GenerationError, _
        Buttons:=vbCritical, _
        Title:="Erreur de génération"

End Sub

'-------------------------------------------------------------------------------
' Génère un seul fichier pour un OF donné.
'
' Les quatre feuilles du modèle peuvent être masquées ou xlSheetVeryHidden.
' Leur état d'origine est sauvegardé, elles sont rendues visibles uniquement
' pendant la copie, puis leur visibilité d'origine est immédiatement restaurée.
' Les copies présentes dans le fichier généré sont toujours rendues visibles.
'
' La fonction retourne True si le fichier a été généré correctement.
' Elle retourne False si une erreur s'est produite.
'-------------------------------------------------------------------------------
Private Function GenerateOneOFFile( _
    ByVal wbTemplate As Workbook, _
    ByVal wsSource As Worksheet, _
    ByVal ofNumber As String, _
    ByVal sourceRows As Collection, _
    ByVal fileNameBase As String, _
    ByVal outputFolder As String, _
    ByRef savedFilePath As String, _
    ByRef errorMessage As String _
) As Boolean

    Dim wbOutput As Workbook
    Dim wsInformation As Worksheet
    Dim wsTraceability As Worksheet

    Dim outputData() As Variant

    Dim sourceRowIndex As Long
    Dim dataColumnIndex As Long
    Dim exportedColumnCount As Long

    Dim sourceRowNumber As Long

    Dim documentTitle As String
    Dim outputFileName As String

    Dim templateSheetNames As Variant
    Dim originalVisibility() As XlSheetVisibility
    Dim templateSheetIndex As Long
    Dim sourceVisibilityNeedsRestore As Boolean

    On Error GoTo GenerationError

    '---------------------------------------------------------------------------
    ' Copie uniquement les quatre feuilles devant apparaître dans le fichier
    ' final.
    '
    ' Les feuilles peuvent rester masquées dans le classeur modèle. Pour rendre
    ' la copie robuste quel que soit leur état, on sauvegarde leur visibilité,
    ' on les affiche temporairement, puis on restaure immédiatement leur état.
    '---------------------------------------------------------------------------

    templateSheetNames = Array( _
        TG_SCREEN_VIEW_SHEET, _
        TG_INFORMATION_SHEET, _
        TG_TRACEABILITY_SHEET, _
        TG_HISTORY_SHEET _
    )

    ReDim originalVisibility( _
        LBound(templateSheetNames) To UBound(templateSheetNames) _
    )

    For templateSheetIndex = _
        LBound(templateSheetNames) To UBound(templateSheetNames)

        originalVisibility(templateSheetIndex) = _
            wbTemplate.Worksheets( _
                CStr(templateSheetNames(templateSheetIndex)) _
            ).Visible

    Next templateSheetIndex

    sourceVisibilityNeedsRestore = True

    For templateSheetIndex = _
        LBound(templateSheetNames) To UBound(templateSheetNames)

        wbTemplate.Worksheets( _
            CStr(templateSheetNames(templateSheetIndex)) _
        ).Visible = xlSheetVisible

    Next templateSheetIndex

    wbTemplate.Worksheets(templateSheetNames).Copy

    Set wbOutput = ActiveWorkbook

    ' Restaure immédiatement l'état des feuilles dans le classeur modèle.
    For templateSheetIndex = _
        LBound(templateSheetNames) To UBound(templateSheetNames)

        wbTemplate.Worksheets( _
            CStr(templateSheetNames(templateSheetIndex)) _
        ).Visible = originalVisibility(templateSheetIndex)

    Next templateSheetIndex

    sourceVisibilityNeedsRestore = False

    ' Les feuilles générées doivent être visibles dans le fichier final.
    For templateSheetIndex = _
        LBound(templateSheetNames) To UBound(templateSheetNames)

        wbOutput.Worksheets( _
            CStr(templateSheetNames(templateSheetIndex)) _
        ).Visible = xlSheetVisible

    Next templateSheetIndex

    Set wsInformation = _
        wbOutput.Worksheets(TG_INFORMATION_SHEET)

    Set wsTraceability = _
        wbOutput.Worksheets(TG_TRACEABILITY_SHEET)

    ' Supprime les éventuelles anciennes données présentes dans le modèle.
    ClearTraceabilityData wsTraceability

    exportedColumnCount = _
        TG_LAST_DATA_COLUMN - TG_FIRST_DATA_COLUMN + 1

    ReDim outputData( _
        1 To sourceRows.Count, _
        1 To exportedColumnCount _
    )

    '---------------------------------------------------------------------------
    ' Charge les données B:AC dans un tableau mémoire.
    '
    ' Cela évite de copier les cellules une par une vers le nouveau classeur.
    '---------------------------------------------------------------------------

    For sourceRowIndex = 1 To sourceRows.Count

        sourceRowNumber = CLng( _
            sourceRows.Item(sourceRowIndex) _
        )

        For dataColumnIndex = 1 To exportedColumnCount

            outputData( _
                sourceRowIndex, _
                dataColumnIndex _
            ) = wsSource.Cells( _
                sourceRowNumber, _
                TG_FIRST_DATA_COLUMN + dataColumnIndex - 1 _
            ).Value2

        Next dataColumnIndex

    Next sourceRowIndex

    '---------------------------------------------------------------------------
    ' Colle toutes les lignes de l'OF dans TRACEABILITY en une seule opération.
    '
    ' OFs!B:AC devient TRACEABILITY!A:AB
    '---------------------------------------------------------------------------

    wsTraceability.Cells( _
        TG_FIRST_TRACEABILITY_ROW, _
        TG_FIRST_TRACEABILITY_COLUMN _
    ).Resize( _
        sourceRows.Count, _
        exportedColumnCount _
    ).Value2 = outputData

    '---------------------------------------------------------------------------
    ' Renseigne la feuille INFORMATIONS.
    '---------------------------------------------------------------------------

    documentTitle = fileNameBase & " " & ofNumber

    With wsInformation

        ' Force B3 au format texte pour conserver les zéros à gauche.
        .Range(TG_INFORMATION_OF_CELL).NumberFormat = "@"
        .Range(TG_INFORMATION_OF_CELL).Value2 = ofNumber

        .Range(TG_INFORMATION_TITLE_CELL).Value2 = documentTitle

    End With

    '---------------------------------------------------------------------------
    ' Construit le nom et le chemin du fichier.
    '---------------------------------------------------------------------------

    outputFileName = _
        CleanFileName( _
            CStr( _
                wsInformation.Range( _
                    TG_INFORMATION_TITLE_CELL _
                ).Value2 _
            ) _
        ) & TG_OUTPUT_EXTENSION

    savedFilePath = CombinePath( _
        outputFolder, _
        outputFileName _
    )

    ' Remplace automatiquement le fichier s'il existe déjà.
    DeleteExistingFile savedFilePath

    '---------------------------------------------------------------------------
    ' Enregistre le fichier généré au format XLSX.
    '
    ' Le fichier final ne contient aucune macro.
    '---------------------------------------------------------------------------

    wbOutput.SaveAs _
        fileName:=savedFilePath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False, _
        Local:=True

    wbOutput.Close SaveChanges:=False
    Set wbOutput = Nothing

    GenerateOneOFFile = True

    Exit Function

GenerationError:

    errorMessage = _
        "Erreur " & Err.Number & _
        " - " & Err.Description

    On Error Resume Next

    ' Même en cas d'erreur pendant la copie, le classeur modèle retrouve
    ' toujours exactement l'état de visibilité qu'il avait au départ.
    If sourceVisibilityNeedsRestore Then

        For templateSheetIndex = _
            LBound(templateSheetNames) To UBound(templateSheetNames)

            wbTemplate.Worksheets( _
                CStr(templateSheetNames(templateSheetIndex)) _
            ).Visible = originalVisibility(templateSheetIndex)

        Next templateSheetIndex

    End If

    If Not wbOutput Is Nothing Then
        wbOutput.Close SaveChanges:=False
    End If

    On Error GoTo 0

    GenerateOneOFFile = False

End Function
