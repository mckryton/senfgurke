VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TDataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_column_names As Collection
Dim m_table_rows As Collection

Private Sub Class_Terminate()
    Set m_column_names = Nothing
    Set m_table_rows = Nothing
End Sub

Public Sub add_row(feature_line As String)
        
    Dim column_name As Variant
    Dim column_value As Variant
    Dim column_values As Variant
    Dim column_index As Integer
    Dim cell_text As String
    Dim table_row As Collection

    feature_line = Trim(feature_line)
    column_values = Split(feature_line, "|")
    If UBound(column_values) - 1 <> Me.column_names.Count Then
        TError.raise ERR_ID_STEP_SYNTAX_TABLE_COLUMN_COUNT_MISMATCH, "TDataTable.add_row", Array(Trim(feature_line))
    End If
    Set table_row = New Collection
    For column_index = 1 To Me.column_names.Count
        column_name = Me.column_names(column_index)
        cell_text = Trim(column_values(column_index))
        table_row.Add cell_text, column_name
    Next
    Me.table_rows.Add table_row
End Sub

Public Sub set_header(feature_line As String)

    Dim column_name As String
    Dim column_names As Variant
    Dim column_index As Integer

    feature_line = Trim(feature_line)
    column_names = Split(feature_line, "|")
    For column_index = 1 To UBound(column_names) - 1
        column_name = Trim(CStr(column_names(column_index)))
        If ExtraVBA.collection_has_value(column_name, Me.column_names) Then
            TError.raise ERR_ID_STEP_SYNTAX_TABLE_DUPLICATE_COLUMN, _
                "TDataTable.set_header", _
                Array(column_name)
        End If
        Me.column_names.Add column_name
    Next
End Sub

Public Function add_column_name(new_name As String) As Boolean

    If ExtraVBA.collection_has_value(new_name, Me.column_names) Then
        'indicate that adding a new column name has failed
        add_column_name = False
    Else
        Me.column_names.Add new_name
        add_column_name = True
    End If
End Function

Public Property Get column_names() As Collection
    If m_column_names Is Nothing Then Set m_column_names = New Collection
    Set column_names = m_column_names
End Property

Public Property Get table_rows() As Collection
    If m_table_rows Is Nothing Then Set m_table_rows = New Collection
    Set table_rows = m_table_rows
End Property

