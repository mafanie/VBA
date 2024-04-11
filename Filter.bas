' FILTER
'
' Some useful macros related to filters
'
' - UseTextFilter: Ctrl + E
' - TODO: RemoveAllFilterWorksheet/Workbook

' Determine filter range and use (fast) text filter via InputBox
' TODO:
'   - Support two search texts (for mmore than two explicit selection is necessary)
'   - Use ## to remove all filter on worksheet and ### to remove all filter in workbook
Sub UseTextFilter()
 
    Dim ws As Worksheet: Set ws = ActiveSheet
    'Debug.Print "ws.Name: " & ws.Name
 
    Dim rngActive As Range
    Dim rngCurrent As Range
    Dim rngAutoFilter As Range
    Dim rngFiler As Range
 
    Dim intResponse As Integer
    Dim strSearch As String
 
    Set rngActive = ActiveCell
    'Debug.Print "rngActive.Address : " & rngActive.Address
 
    ' If selection not in list object
    If Selection.ListObject Is Nothing Then
        Set rngCurrent = ActiveCell.CurrentRegion
        'Debug.Print "rngCurrent.Address: " & rngCurrent.Address
        ' If auto filter exists
        If ws.AutoFilterMode Then
            Set rngAutoFilter = ws.AutoFilter.Range
            'Debug.Print "rngAutoFilter.Address: " & rAutoFilter.Address
            ' If selection is not in auto filter range: ask to remove auto filter and use current region to filter
            If Application.Intersect(rngActive, rngAutoFilter) Is Nothing Then
                ' English version
                Response = MsgBox("Selection is outside the AutoFilter range." & vbNewLine & _
                        "The existing AutoFilter will be removed and the current region will be used.", _
                        vbOKCancel, _
                        "Quick Search in Tables: Filter Area")
                ' German version
                'Response = MsgBox("Auswahl liegt außerhalb des AutoFilter-Bereichs." & vbNewLine & _
                '        "Der bestehende AutoFilter wird entfernt und die aktuelle Region verwendet.", _
                '        vbOKCancel, _
                '        "Schnelle Suche in Tabellen: Filterbereich")
                ' If 'OK': remove auto filter
                If Response = vbOK Then
                    ws.AutoFilterMode = False
                    Set rngFilter = rngCurrent
                ' Else = 'Cancel': exit
                Else
                    Exit Sub
                End If
            ' Else = If selection is in auto filter range: use that range to filter
            Else
                Set rngFilter = rngAutoFilter
            End If
        ' If no auto filter exists: use current region to filter
        Else
            Set rngFilter = rngCurrent
        End If
    ' Else = If selection is in list object: use range of list object to filter
    Else
        'Debug.Print "ListObject.Address " & Selection.ListObject.Range.Address
        Set rngFilter = Selection.ListObject.Range
    End If
    'Debug.Print "rngFilter.Address " & rngFilter.Address
 
    Dim RelColField As Integer
    RelColField = ActiveCell.Column - rngFilter.Column + 1
    'Debug.Print "RelColField: " & RelColField
    ' English version
    Search = InputBox( _
             "Please enter search text." & vbNewLine & _
             "'?': Wildcard for a character" & vbNewLine & _
             "'*': Wildcard for any characters" & vbNewLine & _
             "'~': Placeholder for value of active cell" & vbNewLine & _
             "'#': Deletes the filters from the table", _
             "Quick search in tables: Search text")
    ' German version
    'Search = InputBox( _
    '        "Bitte Suchtext eingeben. " & vbNewLine & _
    '        "'?': Platzhalter für ein Zeichen" & vbNewLine & _
    '        "'*': Platzhalter für beliebige Zeichen" & vbNewLine & _
    '        "'~': Platzhalter für Wert der aktiven Zelle" & vbNewLine & _
    '        "'#': Löscht die Filter aus der Tabelle", _
    '        "Schnelle Suche in Tabellen: Suchtext")
    'Debug.Print "Search: " & Search
 
    Search = Replace(Search, "~", ActiveCell.Value)
    'Debug.Print "Search: " & Search
 
    ' "": Noting
    If Search = "" Then
        Exit Sub
    ' "*": Remove filter in column
    ElseIf Search = "*" Or Search = "" Then
        rngFilter.AutoFilter Field:=RelColField
    ' "#": Remove all filters
    ElseIf Search = "#" Then
        If ws.FilterMode Then
            ws.ShowAllData
        End If
    ' Else: Filter
    Else
        rngFilter.AutoFilter Field:=RelColField, Criteria1:=Search
    End If
 
End Sub
