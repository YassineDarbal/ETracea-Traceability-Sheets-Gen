Attribute VB_Name = "modTG_Dashboard"
Option Explicit

'===============================================================================
' DASHBOARD ETRACEA
'===============================================================================

Private Const TG_STATUS_READY As String = "READY"
Private Const TG_STATUS_RUNNING As String = "RUNNING"
Private Const TG_STATUS_SUCCESS As String = "SUCCESS"
Private Const TG_STATUS_WARNING As String = "WARNING"
Private Const TG_STATUS_ERROR As String = "ERROR"

'-------------------------------------------------------------------------------
' Construit ou reconstruit complètement la feuille Dashboard.
'
' Les valeurs existantes des deux noms définis sont conservées :
'   TG_FileName     -> Dashboard!F8
'   TG_OutputFolder -> Dashboard!F10
'-------------------------------------------------------------------------------
Public Sub BuildDashboard()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileNameValue As String
    Dim outputFolderValue As String
    Dim previousScreenUpdating As Boolean

    On Error GoTo BuildError

    Set wb = ThisWorkbook

    fileNameValue = ReadExistingNamedValue(wb, TG_FILE_NAME_RANGE)
    outputFolderValue = ReadExistingNamedValue(wb, TG_OUTPUT_FOLDER_RANGE)

    If Len(Trim$(fileNameValue)) = 0 Then
        fileNameValue = "Traçabilité"
    End If

    If Len(Trim$(outputFolderValue)) = 0 Then
        outputFolderValue = wb.Path
    End If

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set ws = GetOrCreateDashboardSheet(wb)

    PrepareDashboardCanvas ws
    BuildDashboardHeader ws
    BuildConfigurationCard ws, fileNameValue, outputFolderValue
    BuildPrimaryAction ws
    BuildStatusCard ws
    BuildDashboardFooter ws

    CreateOrReplaceWorkbookName _
        wb:=wb, _
        definedName:=TG_FILE_NAME_RANGE, _
        targetCell:=ws.Range("F8")

    CreateOrReplaceWorkbookName _
        wb:=wb, _
        definedName:=TG_OUTPUT_FOLDER_RANGE, _
        targetCell:=ws.Range("F10")

    UpdateDashboardStatus _
        messageText:="Configuration prête. Renseignez les paramètres puis lancez la génération.", _
        statusType:=TG_STATUS_READY

    ws.Activate
    ws.Range("A1").Select

    Application.ScreenUpdating = previousScreenUpdating

    MsgBox _
        Prompt:= _
            "Le Dashboard ETracea a été généré avec succès." & _
            vbCrLf & vbCrLf & _
            "Les valeurs existantes de TG_FileName et TG_OutputFolder ont été conservées.", _
        Buttons:=vbInformation, _
        Title:="Dashboard ETracea"

    Exit Sub

BuildError:

    Application.ScreenUpdating = previousScreenUpdating

    MsgBox _
        Prompt:= _
            "Impossible de générer le Dashboard." & _
            vbCrLf & vbCrLf & _
            Err.Description, _
        Buttons:=vbCritical, _
        Title:="Dashboard ETracea"

End Sub

'-------------------------------------------------------------------------------
' Ouvre un sélecteur de dossier et alimente TG_OutputFolder.
'-------------------------------------------------------------------------------
Public Sub SelectOutputFolder()

    Dim folderPicker As FileDialog
    Dim targetCell As Range
    Dim initialFolder As String

    On Error GoTo FolderError

    Set targetCell = _
        ThisWorkbook.Names(TG_OUTPUT_FOLDER_RANGE).RefersToRange

    initialFolder = Trim$(CStr(targetCell.Value2))

    Set folderPicker = _
        Application.FileDialog(msoFileDialogFolderPicker)

    With folderPicker

        .Title = "Sélectionnez le dossier de sortie"
        .AllowMultiSelect = False

        If Len(initialFolder) > 0 Then
            .InitialFileName = initialFolder
        End If

        If .Show <> -1 Then
            Exit Sub
        End If

        targetCell.NumberFormat = "@"
        targetCell.Value2 = CStr(.SelectedItems(1))

    End With

    UpdateDashboardStatus _
        messageText:="Dossier de sortie sélectionné. Le générateur est prêt.", _
        statusType:=TG_STATUS_READY

    Exit Sub

FolderError:

    MsgBox _
        Prompt:= _
            "Impossible de sélectionner le dossier de sortie." & _
            vbCrLf & vbCrLf & _
            Err.Description, _
        Buttons:=vbExclamation, _
        Title:="Dossier de sortie"

End Sub

'-------------------------------------------------------------------------------
' Supprime toutes les lignes de la feuille OFs à partir de la ligne 3 incluse.
' Les lignes 1 et 2 sont toujours conservées.
'-------------------------------------------------------------------------------
Public Sub ClearOFs()

    Dim ws As Worksheet
    Dim lastCell As Range
    Dim lastRow As Long
    Dim answer As VbMsgBoxResult
    Dim previousScreenUpdating As Boolean

    On Error GoTo ClearError

    Set ws = ThisWorkbook.Worksheets(TG_SOURCE_SHEET)

    Set lastCell = ws.Cells.Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False _
    )

    If lastCell Is Nothing Then

        UpdateDashboardStatus _
            messageText:="La feuille OFs est déjà vide.", _
            statusType:=TG_STATUS_READY

        MsgBox _
            Prompt:="La feuille '" & TG_SOURCE_SHEET & "' est déjà vide.", _
            Buttons:=vbInformation, _
            Title:="Nettoyage des OF"

        Exit Sub

    End If

    lastRow = lastCell.Row

    If lastRow < TG_FIRST_SOURCE_ROW Then

        UpdateDashboardStatus _
            messageText:="Aucune donnée OF à supprimer.", _
            statusType:=TG_STATUS_READY

        MsgBox _
            Prompt:= _
                "Aucune donnée à supprimer à partir de la ligne " & _
                TG_FIRST_SOURCE_ROW & ".", _
            Buttons:=vbInformation, _
            Title:="Nettoyage des OF"

        Exit Sub

    End If

    answer = MsgBox( _
        Prompt:= _
            "Voulez-vous supprimer toutes les lignes de la feuille '" & _
            TG_SOURCE_SHEET & "' à partir de la ligne " & _
            TG_FIRST_SOURCE_ROW & " ?" & _
            vbCrLf & vbCrLf & _
            "Cette action supprimera les données actuellement chargées.", _
        Buttons:=vbYesNo + vbQuestion + vbDefaultButton2, _
        Title:="Confirmer le nettoyage" _
    )

    If answer <> vbYes Then
        Exit Sub
    End If

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ws.Rows( _
        TG_FIRST_SOURCE_ROW & ":" & lastRow _
    ).Delete Shift:=xlUp

    Application.ScreenUpdating = previousScreenUpdating

    UpdateDashboardStatus _
        messageText:="La feuille OFs a été vidée. Vous pouvez charger un nouveau lot.", _
        statusType:=TG_STATUS_READY

    MsgBox _
        Prompt:= _
            "Les données de la feuille '" & TG_SOURCE_SHEET & _
            "' ont été supprimées.", _
        Buttons:=vbInformation, _
        Title:="Nettoyage terminé"

    Exit Sub

ClearError:

    On Error Resume Next
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0

    UpdateDashboardStatus _
        messageText:="Erreur pendant le nettoyage de la feuille OFs.", _
        statusType:=TG_STATUS_ERROR

    MsgBox _
        Prompt:= _
            "Impossible de vider la feuille '" & TG_SOURCE_SHEET & "'." & _
            vbCrLf & vbCrLf & _
            Err.Description, _
        Buttons:=vbCritical, _
        Title:="Erreur de nettoyage"

End Sub

'-------------------------------------------------------------------------------
' Met à jour la zone de statut du Dashboard.
' statusType : READY / RUNNING / SUCCESS / WARNING / ERROR
'-------------------------------------------------------------------------------
Public Sub UpdateDashboardStatus( _
    ByVal messageText As String, _
    Optional ByVal statusType As String = TG_STATUS_READY _
)

    Dim ws As Worksheet
    Dim titleCell As Range
    Dim messageCell As Range
    Dim statusTitle As String
    Dim statusFill As Long
    Dim statusFont As Long

    On Error GoTo SafeExit

    If Not WorksheetExists(TG_DASHBOARD_SHEET, ThisWorkbook) Then
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(TG_DASHBOARD_SHEET)
    Set titleCell = ws.Range("C18")
    Set messageCell = ws.Range("C19")

    Select Case UCase$(Trim$(statusType))

        Case TG_STATUS_RUNNING
            statusTitle = "EN COURS"
            statusFill = RGB(227, 242, 253)
            statusFont = RGB(21, 101, 192)

        Case TG_STATUS_SUCCESS
            statusTitle = "TERMINÉ"
            statusFill = RGB(232, 245, 233)
            statusFont = RGB(27, 94, 32)

        Case TG_STATUS_WARNING
            statusTitle = "ATTENTION"
            statusFill = RGB(255, 248, 225)
            statusFont = RGB(181, 101, 0)

        Case TG_STATUS_ERROR
            statusTitle = "ERREUR"
            statusFill = RGB(255, 235, 238)
            statusFont = RGB(183, 28, 28)

        Case Else
            statusTitle = "PRÊT"
            statusFill = RGB(244, 246, 248)
            statusFont = RGB(30, 111, 140)

    End Select

    With titleCell.MergeArea
        .Interior.Color = statusFill
        .Font.Color = statusFont
        .Font.Bold = True
    End With

    titleCell.Value2 = statusTitle

    With messageCell.MergeArea
        .Interior.Color = statusFill
        .Font.Color = RGB(31, 41, 51)
    End With

    messageCell.Value2 = messageText

SafeExit:

End Sub

'===============================================================================
' CONSTRUCTION DE L'INTERFACE
'===============================================================================

Private Function GetOrCreateDashboardSheet( _
    ByVal wb As Workbook _
) As Worksheet

    If WorksheetExists(TG_DASHBOARD_SHEET, wb) Then

        Set GetOrCreateDashboardSheet = _
            wb.Worksheets(TG_DASHBOARD_SHEET)

    Else

        Set GetOrCreateDashboardSheet = _
            wb.Worksheets.Add(Before:=wb.Worksheets(1))

        GetOrCreateDashboardSheet.Name = TG_DASHBOARD_SHEET

    End If

End Function

Private Sub PrepareDashboardCanvas(ByVal ws As Worksheet)

    Dim shapeIndex As Long

    ws.Cells.UnMerge
    ws.Cells.Clear

    For shapeIndex = ws.Shapes.Count To 1 Step -1
        ws.Shapes(shapeIndex).Delete
    Next shapeIndex

    With ws.Range("A1:L28")
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Aptos"
        .Font.Size = 11
        .Font.Color = RGB(31, 41, 51)
    End With

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 3
    ws.Columns("C").ColumnWidth = 17
    ws.Columns("D").ColumnWidth = 17
    ws.Columns("E").ColumnWidth = 17
    ws.Columns("F").ColumnWidth = 17
    ws.Columns("G").ColumnWidth = 17
    ws.Columns("H").ColumnWidth = 17
    ws.Columns("I").ColumnWidth = 17
    ws.Columns("J").ColumnWidth = 3
    ws.Columns("K:L").ColumnWidth = 3

    ws.Rows("1").RowHeight = 10
    ws.Rows("2:3").RowHeight = 34
    ws.Rows("4").RowHeight = 24
    ws.Rows("5:6").RowHeight = 12
    ws.Rows("7").RowHeight = 26
    ws.Rows("8").RowHeight = 30
    ws.Rows("9").RowHeight = 22
    ws.Rows("10").RowHeight = 30
    ws.Rows("11:12").RowHeight = 14
    ws.Rows("13").RowHeight = 12
    ws.Rows("14:16").RowHeight = 24
    ws.Rows("17").RowHeight = 16
    ws.Rows("18").RowHeight = 25
    ws.Rows("19:21").RowHeight = 23
    ws.Rows("22:23").RowHeight = 12
    ws.Rows("24").RowHeight = 18

    ws.Tab.Color = RGB(17, 47, 70)

End Sub

Private Sub BuildDashboardHeader(ByVal ws As Worksheet)

    With ws.Range("B2:B4")
        .Merge
        .Interior.Color = RGB(208, 0, 0)
    End With

    With ws.Range("C2:J4")
        .Interior.Color = RGB(17, 47, 70)
    End With

    With ws.Range("C2:J3")
        .Merge
        .Value2 = "Générateur ETracea"
        .Font.Name = "Aptos Display"
        .Font.Size = 26
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("C4:J4")
        .Merge
        .Value2 = "Génération en masse des gammes de traçabilité"
        .Font.Size = 11
        .Font.Color = RGB(220, 228, 234)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

End Sub

Private Sub BuildConfigurationCard( _
    ByVal ws As Worksheet, _
    ByVal fileNameValue As String, _
    ByVal outputFolderValue As String _
)

    With ws.Range("C7:I12")
        .Interior.Color = RGB(248, 250, 252)
        .Borders.Color = RGB(216, 222, 228)
        .Borders.Weight = xlThin
    End With

    With ws.Range("C7:I7")
        .Merge
        .Value2 = "CONFIGURATION"
        .Interior.Color = RGB(17, 47, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("C8:E8")
        .Merge
        .Value2 = "Nom des fichiers"
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("F8:I8")
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.Color = RGB(190, 199, 207)
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "@"
    End With

    ws.Range("F8").Value2 = fileNameValue

    With ws.Range("F9:I9")
        .Merge
        .Value2 = "Le numéro d'OF est ajouté automatiquement au nom du fichier."
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(103, 114, 125)
        .HorizontalAlignment = xlLeft
    End With

    With ws.Range("C10:E10")
        .Merge
        .Value2 = "Dossier de sortie"
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("F10:H10")
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.Color = RGB(190, 199, 207)
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .NumberFormat = "@"
        .ShrinkToFit = True
    End With

    ws.Range("F10").Value2 = outputFolderValue

    CreateDashboardButton _
        ws:=ws, _
        shapeName:="tg_btnBrowse", _
        caption:="PARCOURIR", _
        targetArea:=ws.Range("I10"), _
        macroName:="SelectOutputFolder", _
        fillColor:=RGB(244, 246, 248), _
        fontColor:=RGB(17, 47, 70), _
        borderColor:=RGB(17, 47, 70), _
        fontSize:=9

End Sub

Private Sub BuildPrimaryAction(ByVal ws As Worksheet)

    CreateDashboardButton _
        ws:=ws, _
        shapeName:="tg_btnClearOFs", _
        caption:="VIDER LES OFs", _
        targetArea:=ws.Range("C14:C16"), _
        macroName:="ClearOFs", _
        fillColor:=RGB(255, 255, 255), _
        fontColor:=RGB(183, 28, 28), _
        borderColor:=RGB(183, 28, 28), _
        fontSize:=10

    CreateDashboardButton _
        ws:=ws, _
        shapeName:="tg_btnGenerate", _
        caption:="GÉNÉRER LES FICHIERS", _
        targetArea:=ws.Range("D14:H16"), _
        macroName:="GenerateTraceabilityFiles", _
        fillColor:=RGB(30, 111, 140), _
        fontColor:=RGB(255, 255, 255), _
        borderColor:=RGB(17, 47, 70), _
        fontSize:=16

End Sub

Private Sub BuildStatusCard(ByVal ws As Worksheet)

    With ws.Range("C18:I21")
        .Interior.Color = RGB(244, 246, 248)
        .Borders.Color = RGB(216, 222, 228)
        .Borders.Weight = xlThin
    End With

    With ws.Range("C18:I18")
        .Merge
        .Value2 = "PRÊT"
        .Font.Bold = True
        .Font.Color = RGB(30, 111, 140)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("C19:I21")
        .Merge
        .Value2 = "Configuration prête."
        .Font.Size = 10
        .Font.Color = RGB(31, 41, 51)
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

End Sub

Private Sub BuildDashboardFooter(ByVal ws As Worksheet)

    With ws.Range("C24:I24")
        .Merge
        .Value2 = "ETracea Traceability Sheets Generator"
        .Font.Size = 9
        .Font.Color = RGB(130, 140, 150)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

End Sub

Private Sub CreateDashboardButton( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal targetArea As Range, _
    ByVal macroName As String, _
    ByVal fillColor As Long, _
    ByVal fontColor As Long, _
    ByVal borderColor As Long, _
    ByVal fontSize As Single _
)

    Dim buttonShape As Shape

    Set buttonShape = ws.Shapes.AddShape( _
        Type:=msoShapeRoundedRectangle, _
        Left:=targetArea.Left, _
        Top:=targetArea.Top, _
        Width:=targetArea.Width, _
        Height:=targetArea.Height _
    )

    With buttonShape

        .Name = shapeName
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macroName

        .Fill.ForeColor.RGB = fillColor
        .Fill.Solid

        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 1.5

        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = fontSize
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fontColor
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.MarginLeft = 6
        .TextFrame2.MarginRight = 6
        .TextFrame2.MarginTop = 3
        .TextFrame2.MarginBottom = 3

        .Placement = xlMoveAndSize

    End With

End Sub

'===============================================================================
' NOMS DÉFINIS / CONSERVATION DES VALEURS
'===============================================================================

Private Function ReadExistingNamedValue( _
    ByVal wb As Workbook, _
    ByVal definedName As String _
) As String

    On Error Resume Next

    ReadExistingNamedValue = _
        Trim$(CStr(wb.Names(definedName).RefersToRange.Value2))

    On Error GoTo 0

End Function

Private Sub CreateOrReplaceWorkbookName( _
    ByVal wb As Workbook, _
    ByVal definedName As String, _
    ByVal targetCell As Range _
)

    On Error Resume Next
    wb.Names(definedName).Delete
    On Error GoTo 0

    wb.Names.Add _
        Name:=definedName, _
        RefersTo:="=" & targetCell.Address( _
            RowAbsolute:=True, _
            ColumnAbsolute:=True, _
            ReferenceStyle:=xlA1, _
            External:=True _
        )

End Sub
