VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayWithNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CLASS: ArrayWithNames
'
' DESCRIPTION:
'
' AUTHOR: Martin Trinks
'
' HISTORY:
' - To 2025-04: Initial version created
'
' TODO:
' - RangeWithNames (reading from and writing to sheets via Names)
'

Option Explicit

Private m_Data() As Variant
Private m_RowNames As Object ' Scripting.Dictionary
Private m_ColNames As Object ' Scripting.Dictionary
Private m_HasRowNames As Boolean
Private m_HasColNames As Boolean

Public Function GetRowIndex(RowKey As Variant) As Long

    ' RowKey is a number and therefore already the row index
    If VarType(RowKey) = vbInteger Or VarType(RowKey) = vbLong Then
        GetRowIndex = CLng(RowKey)
        If GetRowIndex < 1 Or GetRowIndex > RowCount Then
            Err.Raise 9, "NamedArray.GetRowIndex", "Row index is out of range."
            'Err.Raise 9, "NamedArray.GetRowIndex", "Zeilenindex außerhalb des gültigen Bereichs."
        End If
    ' Array does not have row names
    ElseIf Not m_HasRowNames Then
        Err.Raise 13, "NamedArray.GetRowIndex", "Row index (number)expectd, no row names are used."
        'Err.Raise 13, "NamedArray.GetRowIndex", "Zeilenindex (Zahl) erwartet, es werden keine Zeilennamen verwendet."
    ' RowKey is not a string and therefore not a valid row name
    ElseIf Not VarType(RowKey) = vbString Then
        Err.Raise 13, "NamedArray.GetRowIndex", "Invalide data type for RowKey."
        'Err.Raise 13, "NamedArray.GetRowIndex", "Ungültiger Datentyp für RowKey."
    ' There is no row with that row name
    ElseIf Not m_RowNames.Exists(Trim(RowKey)) Then
        Err.Raise 13, "NamedArray.GetRowIndex", "Row names not found: " & RowKey
        'Err.Raise 13, "NamedArray.GetRowIndex", "Zeilenname nicht gefunden: " & RowKey
    ' Get row index
    Else
        GetRowIndex = m_RowNames(Trim(RowKey))
    End If
        
End Function

Public Function GetColIndex(ColKey As Variant) As Long

    ' ColKey is a number and therefore already the col index
    If VarType(ColKey) = vbInteger Or VarType(ColKey) = vbLong Then
        GetColIndex = CLng(ColKey)
        If GetColIndex < 1 Or GetColIndex > ColCount Then
            Err.Raise 9, "NamedArray.GetColIndex", "Column index is out of ramge."
            'Err.Raise 9, "NamedArray.GetColIndex", "Spaltenindex ist außerhalb des gültigen Bereichs."
        End If
    ' Array does not have col names
    ElseIf Not m_HasColNames Then
        Err.Raise 13, "NamedArray.GetColIndex", "Column index (number) expected, no column names are used."
        ' Err.Raise 13, "NamedArray.GetColIndex", "Zeilenindex (Zahl) erwartet, es werden keine Zeilennamen verwendet."
    ' ColKey is not a string and therefore not a valid col name
    ElseIf Not VarType(ColKey) = vbString Then
        Err.Raise 13, "NamedArray.GetColIndex", "Invalide data type for ColKey."
        'Err.Raise 13, "NamedArray.GetColIndex", "Ungültiger Datentyp für ColKey."
    ' There is no col with that col name
    ElseIf Not m_ColNames.Exists(Trim(ColKey)) Then
        Err.Raise 13, "NamedArray.GetColIndex", "Column name not found: " & ColKey
        'Err.Raise 13, "NamedArray.GetColIndex", "Spaltenname nicht gefunden: " & ColKey
    ' Get col index
    Else
        GetColIndex = m_ColNames(Trim(ColKey))
    End If
        
End Function


' Propertys

Public Property Get Item(RowKey As Variant, Optional ColKey As Variant) As Variant
    Dim rowIndex As Long
    Dim colIndex As Long

    rowIndex = GetRowIndex(RowKey)
    colIndex = GetColIndex(ColKey)

    'Debug.Print "rowIndex: " & rowIndex & ", colIndex: " & colIndex

    Item = m_Data(rowIndex, colIndex)

End Property

Public Property Let Item(RowKey As Variant, ColKey As Variant, Value As Variant)
    Dim rowIndex As Long
    Dim colIndex As Long

    rowIndex = GetRowIndex(RowKey)
    colIndex = GetColIndex(ColKey)

    'Debug.Print "rowIndex: " & rowIndex & ", colIndex: " & colIndex
    
    m_Data(rowIndex, colIndex) = Value

End Property

' Methods

Public Sub Initialize(SourceRange As Range, Optional HasRowNames As Boolean = True, Optional HasColNames As Boolean = True)
    
    Dim Col As Long
    Dim Row As Long
    Dim Name As String

    m_HasRowNames = HasRowNames
    m_HasColNames = HasColNames
    Debug.Print "m_HasRowNames: " & m_HasRowNames & ", m_HasColNames: " & m_HasColNames
    
    Set m_RowNames = CreateObject("Scripting.Dictionary")
    Set m_ColNames = CreateObject("Scripting.Dictionary")

    Dim RowCount As Long: RowCount = SourceRange.Rows.Count - IIf(m_HasColNames, 1, 0)
    Dim ColCount As Long: ColCount = SourceRange.Columns.Count - IIf(m_HasRowNames, 1, 0)
    Debug.Print "RowCount: " & RowCount & ", ColCount: " & ColCount

    ' Dimension data array taking into account the names
    ReDim m_Data(IIf(m_HasColNames, 0, 1) To RowCount, IIf(m_HasRowNames, 0, 1) To ColCount)

    ' Write data to array
    For Row = IIf(m_HasColNames, 0, 1) To RowCount
        For Col = IIf(m_HasRowNames, 0, 1) To ColCount
            m_Data(Row, Col) = SourceRange.Cells(IIf(m_HasColNames, 1, 0) + Row, IIf(m_HasRowNames, 1, 0) + Col).Value
        Next Col
    Next Row

    ' Read row names (values of 0. column) and remove trailing or leading blanks
    If m_HasRowNames Then
        For Row = 1 To RowCount
            Name = Trim(CStr(m_Data(Row, 0)))
            'Debug.Print "Name: " & Name
            If Not m_RowNames.Exists(Name) Then
                m_RowNames.Add Name, Row
            Else
                Debug.Print "Warning: Duplicate row name '" & Name & "' in row " & Col & ". Only the first entry is considerted."
                'Debug.Print "Warnung: Doppelter Zeilenname '" & Name & "' in Zeile " & Col & ". Nur der erste Eintrag wird berücksichtigt."
            End If
        Next Row
    End If

    ' Read column names (values of 0. row) and remove trailing or leading blanks
    If m_HasColNames Then
        For Col = 1 To ColCount
            Name = Trim(CStr(m_Data(0, Col)))
            If Not m_ColNames.Exists(Name) Then
                m_ColNames.Add Name, Col
            Else
                Debug.Print "Warning: Duplicate column name '" & Name & "' in columns " & Col & ". Only the first entry is considerted."
                'Debug.Print "Warnung: Doppelter Spaltenname '" & Name & "' in Spalte " & Col & ". Nur der erste Eintrag wird berücksichtigt."
            End If
        Next Col
    End If

End Sub

Public Property Get RowCount() As Long
    RowCount = UBound(m_Data, 1)
End Property

Public Property Get ColCount() As Long
    ColCount = UBound(m_Data, 2)
End Property

