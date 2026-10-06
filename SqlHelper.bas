'
' SqlHelper
' Copyright (c) 2020 Timothy Baxendale (pcluddite@outlook.com)
'
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
'
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
'
Option Explicit

Function CreateInsert(ByVal TableName As String, ByRef RowRange As Variant) As String
    Dim Row As Range
    Dim nCol As Long
    Set Row = GetRange(RowRange)
    CreateInsert = "INSERT INTO " & TableName & " VALUES ( "
    For nCol = 1 To Row.Columns.Count
        Dim Cell As Range
        Set Cell = Row.Cells(1, nCol)
        If TypeName(Cell.Value2) = "String" Then
            CreateInsert = CreateInsert & "'" & Cell.Value2 & "', "
        Else
            CreateInsert = CreateInsert & Cell.Value2 & ", "
        End If
    Next nCol
    CreateInsert = Mid(CreateInsert, 1, Len(CreateInsert) - 2) & " )"
End Function

Function CreateDelete(ByVal TableName As String, ByRef RowRange As Variant, ByRef KeyColumns As Variant) As String
    Dim Row As Range, Keys As Variant
    Dim x As Long
    
    Set Row = GetRange(RowRange)
    
    If TypeName(KeyColumns) = "Range" Then
        Dim KeyArr() As Variant
        ReDim KeyArr(1 To KeyColumns.Columns.Count)
        For x = LBound(KeyArr) To UBound(KeyArr)
            KeyArr(x) = KeyColumns.Columns(x).Column
        Next x
        Keys = KeyArr
    Else
        Keys = KeyColumns
    End If
    
    CreateDelete = "DELETE FROM " & TableName & " WHERE "
    For x = LBound(Keys) To UBound(Keys)
        Dim Value As Variant, Header As String
        With Row.Worksheet
            Value = .Cells(Row.Row, Keys(x)).Value2
            Header = .Cells(1, Keys(x)).Value
        End With
        If TypeName(Value) = "String" Then
            CreateDelete = CreateDelete & Header & " = '" & Value & "' AND "
        Else
            CreateDelete = CreateDelete & Header & " = " & Value & " AND "
        End If
    Next x
    
    CreateDelete = Mid(CreateDelete, 1, Len(CreateDelete) - 5)
    
End Function
