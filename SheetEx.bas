'
' SheetEx
' Copyright (c) 2016-2017 Timothy Baxendale (pcluddite@outlook.com)
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
'
' Last update: 11/17/17
' Requires: N/A
' Description: Miscelaneous functions for interacting with sheets and ranges
'
' Change the SHEETEX_VER constant to keep certain backward compatability
' Revisions stack on top of one another, i.e. If you roll back to Rev 1000, you will revert
' any changes made after that version. New functions may still exist when reset
' to an earlier version, but this is not guaranteed.
'
' Revisions:
'   1000 - use exBoth for `Shrink` default in ShrinkToFit/ResizeToFit
'   1001 - use False for `KeepFormula` default in ClearTable
'   1002 - use `AllowTables` param name instead of `StopAtTables` for FindLastRow/FindLastColumn
'   1004 - use exRow for `Shrink` default in ShrinkToFit/ResizeToFit
'   1005 - use empty string default value for JoinReplace
'   1007 - use True for `IncludeHeader` default in GetRange and do not check if range object fits ListObject.DataBodyRange
'   1009 - do not use clipboard by default
'   1010 - CURRENT VERSION
'
#Const SHEETEX_VER = 1010
#Const SAFE_RANGE_ENABLED = True ' set to True for bounds checking on ranges

' integer constants
Public Const MAX_BYTE As Byte = 255
Public Const MIN_BYE As Byte = 0
Public Const MAX_INT As Integer = 32767
Public Const MIN_INT As Integer = -32768
Public Const MAX_LONG As Long = 2147483647
Public Const MIN_LONG As Long = -2147483648#

'
' Clear value types
'
Enum ExClearType
    exClearContents = 1
    exClearFormats = 2
    exClearComments = 4
    exClearHLinks = 8
    exClearNotes = 16
    exClearOutline = 32
    exClearAll = 63 ' this is the sum of all the previous
End Enum

'
' Column/Row or both
'
Enum ExColumnRow
    exColumn = 1
    exRow = 2
    exBoth = 3
End Enum

'
' Appends a new row of values to a range
'
Function AppendRow(ByRef Range As Variant, ParamArray Values() As Variant) As Range
    Values = GetArray(Values)
    Set AppendRow = RangeBelow(Range, FindLast:=True, Height:=1)
    With AppendRow
        Dim i As Long, Value As Variant
        i = 1
        For Each Value In Values
            .Columns(i).Value2 = Value
            i = i + 1
        Next Value
    End With
End Function

'
' Inserts a value in an array
'
Sub ArrayInsert(ByRef arr As Variant, ByVal Index As Long, ByRef Value As Variant)
    Dim newarr() As Variant, x As Long
    ReDim newarr(LBound(arr) To UBound(arr) + 1)
    
    For x = 0 To Index - 1
        LetOrSet newarr(x), arr(x)
    Next x
    
    LetOrSet newarr(Index), Value
    
    For x = Index + 1 To UBound(newarr)
        LetOrSet newarr(x), arr(x - 1)
    Next x
    
    arr = newarr
End Sub

'
' Converts a single dimentional array to a list
'
Function ArrayToList(ByRef arr As Variant) As Collection
    Dim x As Long
    Set ArrayToList = New Collection
    For x = LBound(arr) To UBound(arr)
        ArrayToList.Add arr(x)
    Next x
End Function


'
' Builds a formula with given ranges
'
' Ranges - a range or array of ranges that will be used as an argument for this formula
' Formula - the formula that should be applied to the array. Arguments can be used in the formula string
'           in the format %1 is arg1, %2 is arg2, etc.
'
Function BuildFormula(ByVal Ranges As Variant, ByVal Formula As String) As String
    Dim i As Long
    
    If Not IsArray(Ranges) Or TypeName(Ranges) = "Range" Then Ranges = Array(Ranges)
        
    For i = LBound(Ranges) To UBound(Ranges)
        Dim arg As String, argRange As Range
        Set argRange = GetRange(Ranges(i)).Resize(RowSize:=1)
        arg = argRange.Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Formula = Replace(Formula, "%" & (i + 1), arg)
    Next i
    
    BuildFormula = Formula
End Function

'
' Clears all tables in a worksheet
'
Sub ClearAllTables(ByRef Worksheet As Worksheet, Optional ByVal KeepFormula As Boolean = False)
    Dim n As Long
    With Worksheet
        Dim loTable As ListObject
        For Each loTable In Worksheet.ListObjects
            ClearTable loTable, NewHeight:=1, KeepFormula:=KeepFormula
        Next loTable
    End With
End Sub

'
' Clears a range
'
Private Sub ClearEx(ByRef Range As Range, ByVal ClearType As ExClearType)
    
    If HasFlag(ClearType, exClearAll) Then
        Range.Clear
        Exit Sub
    End If
    
    If HasFlag(ClearType, exClearContents) Then Range.ClearContents
    If HasFlag(ClearType, exClearComments) Then Range.ClearComments
    If HasFlag(ClearType, exClearFormats) Then Range.ClearFormats
    If HasFlag(ClearType, exClearHLinks) Then Range.ClearHyperlinks
    If HasFlag(ClearType, exClearNotes) Then Range.ClearNotes
    If HasFlag(ClearType, exClearOutline) Then Range.ClearOutline

End Sub

'
' Clears the filters for a given table
'
Sub ClearFilters(ByRef Table As Variant)
    Dim loTable As ListObject
    Set loTable = TryGetTable(Table)
    If loTable Is Nothing Then
        With GetRange(Table, IncludeHeader:=True)
            If .Worksheet.AutoFilterMode Then .AutoFilter
        End With
    Else
        loTable.AutoFilter.ShowAllData
    End If
End Sub

'
' Clear data from a Range
' Range - the Range to clear
' KeepFormula - whether or not to keep formulas
' ClearType - the type of data to clear
'
Sub ClearRange(ByRef Range As Variant, Optional ByVal KeepFormula As Boolean = True, Optional ByVal ClearType As ExClearType = ExClearType.exClearContents)
    Dim nCol As Long
    Dim rSrc As Range, rCol As Range
    Set rSrc = GetRange(Range)
    
    If Not KeepFormula Then
        ClearEx rSrc, ClearType
    Else
        For Each rCol In rSrc.Columns
            If Not rCol.HasFormula Then ClearEx rCol, ClearType
        Next rCol
    End If
End Sub

'
' Clears the data from a table
' Table - the table name, or a ListObject with the table
' NewRows - resize the table to this number of Rows. values less than 1 will not resize the table
' KeepFormula - whether or not to keep formulas. this is false by default.
'
#If SHEETEX_VER > 1001 Then
Sub ClearTable(ByRef Table As Variant, Optional ByVal NewHeight As Long = 1, Optional ByVal KeepFormula As Boolean = True)
#Else
Sub ClearTable(ByRef Table As Variant, Optional ByVal NewHeight As Long = 1, Optional ByVal KeepFormula As Boolean = False)
#End If
    With GetTable(Table)
        ClearFilters .Range
        ' trim any extra Rows at the bottom and resize the table
        If NewHeight > 0 Then
            If Not (.DataBodyRange Is Nothing) Then
                If NewHeight < .DataBodyRange.Rows.Count Then RShrink(.DataBodyRange, Height:=-NewHeight).Clear
                If NewHeight = 1 And Not KeepFormula Then .DataBodyRange.Delete
            End If
#If SAFE_RANGE_ENABLED Then
            .Resize ROffset(.Range, Height:=NewHeight + 1)
#Else
            .Resize .Range.Resize(RowSize:=NewHeight + 1)
#End If
        End If
        
        If Not (.DataBodyRange Is Nothing) Then
            ClearRange .DataBodyRange, KeepFormula
        End If
    End With
End Sub

'
' Performs an index match (hlookup)
' LookupValue - the key to search for in the LookupRange
' LookupRange - the range of keys in which to look for the LookupValue
' ValueRange - the range of values corresponding to the keys in the LookupRange
' DefaultValue - the default value used in case a match was not found
'
Function CLookup(ByRef LookupValue As Variant, ByRef LookupRange As Variant, ByRef ValueRange As Variant, Optional ByRef DefaultValue As Variant) As Variant
    Dim nMatch As Variant
    
    If IsMissing(DefaultValue) Then DefaultValue = CVErr(xlErrNA)
    
    nMatch = MMatch(LookupValue, GetRange(LookupRange))
    If IsError(nMatch) Then
        CLookup = DefaultValue
    Else
        CLookup = GetRange(ValueRange).Columns(nMatch).Value2
    End If
End Function

'
' Union a variable number of ranges, or an array of ranges.
' Use the built in Union() function if the following can be guaranteed:
'     - the number of ranges is known at compile time
'     - all objects are of type "Range"
'     - the order of the ranges does not matter
'
Function CombineRange(ParamArray Ranges() As Variant) As Range
    Dim n As Long
    
    Ranges = GetArray(Ranges)
    Set CombineRange = GetRange(Ranges(LBound(Ranges)))
    
    For n = LBound(Ranges) + 1 To UBound(Ranges)
        Set CombineRange = Union(CombineRange, GetRange(Ranges(n)))
    Next n
    
    SortRange CombineRange
End Function

Private Function ConcatRangeAddr(ByRef Ranges As Variant, Optional ByRef nRows As Long, Optional ByVal Absolute As ExColumnRow = exColumn) As String
    If Not IsArray(Ranges) Or TypeName(Ranges) = "Range" Then
        With GetRange(Ranges)
            ConcatRangeAddr = .Cells(1, 1).Address(RowAbsolute:=HasFlag(Absolute, exRow), ColumnAbsolute:=HasFlag(Absolute, exColumn), External:=True)
            nRows = .Rows.Count
        End With
    Else
        Dim addstrs() As String
        Dim i As Long, curr As Range
        ReDim addstrs(UBound(Ranges))
        nRows = 0
        For i = LBound(Ranges) To UBound(Ranges)
            If TypeName(Ranges(i)) = "String" Then
                addstrs(i) = Ranges(i)
            Else
                With GetRange(Ranges(i))
                    addstrs(i) = .Cells(1, 1).Address(RowAbsolute:=HasFlag(Absolute, exRow), ColumnAbsolute:=HasFlag(Absolute, exColumn), External:=True)
                    nRows = WorksheetFunction.Max(nRows, .Rows.Count)
                End With
            End If
        Next i
        
        ConcatRangeAddr = Join(addstrs, "&")
    End If
End Function

'
' Gets a true or false value indicating whether a given workbook contains a sheet
'
Function ContainsSheet(ByRef Workbook As Workbook, ByRef Index As Variant) As Boolean
    Dim Sheet As Worksheet
    On Error GoTo catch
    Set Sheet = Workbook.Worksheets(Index)
    On Error GoTo 0
    ContainsSheet = True
    Exit Function
catch:
    ContainsSheet = False
End Function

'
' Copies the widths of each column from one range to another
'
Function CopyColumnWidth(ByRef Source As Variant, ByRef Destination As Variant)
    Dim rSrc As Range, rDest As Range
    Dim rCol As Range, n As Long
    
    If TypeName(Source) = "Workbook" And TypeName(Destination) = "Workbook" Then
        Dim SrcSheet As Worksheet, DestSheet As Worksheet
        Dim wbDest As Workbook
        Set wbDest = Destination
        For Each SrcSheet In Source.Worksheets
            Set DestSheet = TryGetSheet(wbDest, SrcSheet.Name)
            If Not (DestSheet Is Nothing) Then CopyColumnWidth SrcSheet, DestSheet
        Next SrcSheet
    Else
        Set rSrc = GetRange(Source)
        Set rDest = GetRange(Destination)
        
        If rSrc.Columns.Count <> rDest.Columns.Count Then Set rDest = ROffset(rDest, Width:=rSrc.Columns.Count)
        
        For Each rCol In rDest.Columns
            n = n + 1
            rCol.EntireColumn.ColumnWidth = rSrc.Columns(n).EntireColumn.ColumnWidth
        Next rCol
    End If
End Function

'
' Copies one or more ranges to given destinations
' Ranges - an array of pairs, the first element the source and the second the destination
' returns - the most Rows copied
'
Function CopyMult(ByRef Ranges As Variant, Optional ByVal CellType As XlCellType, _
                     Optional ByVal Paste As XlPasteType = xlPasteValues, _
                     Optional ByVal Append As Boolean = False, Optional ByVal ForceColumnVisible As Boolean = False, Optional ByVal FindLast As Boolean = False) As Long
    Dim n As Long
    For n = LBound(Ranges) To UBound(Ranges) Step 2
        CopyMult = WorksheetFunction.Max(CopySingle(Ranges(n), Ranges(n + 1), CellType, Paste, Append, ForceColumnVisible, FindLast), CopyMult)
    Next n
End Function

'
' Copies the widths of each row from one range to another
'
Function CopyRowHeight(ByRef Source As Variant, ByRef Destination As Variant)
    Dim rSrc As Range, rDest As Range
    Dim rRow As Range, n As Long
    
    If TypeName(Source) = "Workbook" And TypeName(Destination) = "Workbook" Then
        Dim SrcSheet As Worksheet, DestSheet As Worksheet
        Dim wbDest As Workbook
        Set wbDest = Destination
        For Each SrcSheet In Source.Worksheets
            Set DestSheet = TryGetSheet(wbDest, SrcSheet.Name)
            If Not (DestSheet Is Nothing) Then CopyRowHeight SrcSheet, DestSheet
        Next SrcSheet
    Else
        Set rSrc = GetRange(Source)
        Set rDest = GetRange(Destination)
        
        If rSrc.Rows.Count <> rDest.Rows.Count Then Set rDest = ROffset(rDest, Width:=rSrc.Rows.Count)
        
        For Each rRow In rDest.Rows
            n = n + 1
            rRow.EntireRow.RowHeight = rSrc.Rows(n).EntireRow.RowHeight
        Next rRow
    End If
End Function

'
' Copies a single range to a destination
'
Function CopySingle(ByRef Source As Variant, ByRef Dest As Variant, _
                    Optional ByVal CellType As XlCellType, Optional ByVal Paste As XlPasteType = xlPasteValues, _
                    Optional Append As Boolean = False, Optional ByVal ForceColumnVisible As Boolean = False, Optional ByVal FindLast As Boolean = True) As Long
    Dim rSrc As Range, rDest As Range
    Dim Hidden() As Boolean, LastRow As Long
    Dim Area As Range, Column As Range, n As Long
    
    Set rSrc = GetRange(Source)
    
    SortRange rSrc
    
    If CellType <> Empty Then
        Set rSrc = SafeSpecial(rSrc, CellType)
        If rSrc Is Nothing Then Exit Function
    End If
    
    If Append Then
        Set rDest = RangeBelow(Dest, FindLast:=FindLast).Resize(RowSize:=CountRows(rSrc, True), ColumnSize:=CountColumns(rSrc, True))
    Else
        Set rDest = GetRange(Dest).Resize(RowSize:=CountRows(rSrc, True), ColumnSize:=CountColumns(rSrc, True))
    End If
    
    If ForceColumnVisible Then
        Dim Size As Long
        Let n = 0
        For Each Area In rSrc.Areas
            Size = Size + Area.Columns.Count
            ReDim Preserve Hidden(Size)
            For Each Column In Area.Columns
                Hidden(n) = ShowColumn(Column)
                n = n + 1
            Next Column
        Next Area
    End If
    
    rSrc.Copy
    rDest.PasteSpecial Paste
    
    If ForceColumnVisible Then
        Let n = 0
        For Each Area In rSrc.Areas
            For Each Column In Area.Columns
                If Hidden(n) Then HideColumn Column
                n = n + 1
            Next Column
        Next Area
    End If
    
    If Append Then EnsureTableSize rSrc, Dest, rDest
    CopySingle = CountRows(rSrc, AssumeOrder:=True)
End Function

'
' Gets the number of columns from a multi-area range
'
' AssumeOrder - Assumes that the columns of each area are in ascending order.
'               This is faster but won't be correct unless the areas in the range are in order.
'
Property Get CountColumns(ByVal Range As Variant, Optional ByVal AssumeOrder As Boolean = False) As Long
    Dim rArea As Range
    With GetRange(Range)
        If .Areas.Count = 1 Then
            CountColumns = .Columns.Count
        ElseIf AssumeOrder Then
            Dim rLast As Range
            Set rLast = .Areas(1)
            CountColumns = rLast.Columns.Count
            For Each rArea In .Areas
                If rArea.Column >= rLast.Columns.Count + rLast.Column Then
                    CountColumns = CountColumns + rArea.Columns.Count
                    Set rLast = rArea
                End If
            Next rArea
        Else
            Dim dict As New Collection, rCol As Range
            For Each rArea In .Areas
                For Each rCol In rArea.Columns
                    On Error Resume Next
                    dict.Add rCol.Column, CStr(rCol.Column)
                    On Error GoTo 0
                Next rCol
            Next rArea
            CountColumns = dict.Count
        End If
    End With
End Property

'
' Gets the number of rows from a multi-area range
'
' AssumeOrder - Assumes that the rows of each area are in ascending order.
'               This is faster but won't be correct unless the areas in the range are in order.
'
Property Get CountRows(ByVal Range As Variant, Optional ByVal AssumeOrder As Boolean = False) As Long
    Dim rArea As Range
    With GetRange(Range)
        If .Areas.Count = 1 Then ' most likely case
            CountRows = .Rows.Count
        ElseIf AssumeOrder Then
            Dim rLast As Range
            Set rLast = .Areas(1)
            CountRows = rLast.Rows.Count
            For Each rArea In .Areas
                If rArea.Row >= rLast.Rows.Count + rLast.Row Then
                    CountRows = CountRows + rArea.Rows.Count
                    Set rLast = rArea
                End If
            Next rArea
        Else
            Dim dict As New Collection, rRow As Range
            For Each rArea In .Areas
                For Each rRow In rArea.Rows
                    On Error Resume Next
                    dict.Add rRow.Row, CStr(rRow.Row)
                    On Error GoTo 0
                Next rRow
            Next rArea
            CountRows = dict.Count
        End If
    End With
End Property

Private Function CreateArrayLookup(ByRef SourceKeys As Variant, ByRef SourceValues As Variant, ByRef DestKeys As Variant) As String
    Dim Matches() As String
    Dim n As Long
    
    ' SourceKeys and DestKeys MUST have the same bounds
    Debug.Assert LBound(SourceKeys) = LBound(DestKeys) And UBound(SourceKeys) = UBound(DestKeys)
    
    ReDim Matches(LBound(SourceKeys) To UBound(SourceKeys))
    
    For n = LBound(SourceKeys) To UBound(SourceKeys)
        Matches(n) = "(" & GetRange(SourceKeys(n)).Address(External:=True) & "=" & GetRange(DestKeys(n)).Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False, External:=True) & ")"
    Next n
    CreateArrayLookup = Join(Matches, "*")
    
    CreateArrayLookup = "INDEX(" & GetRange(SourceValues).Address(External:=True) & ", MATCH(1, " & CreateArrayLookup & ", 0))"
End Function

'
' Creates an INDEX/MATCH from the given criteria
' if the source keys are an array, the destination keys must also be an array of the same size.
' if the destinations keys are an array but the source is not, the destination ranges are concatenated to form a key
'
Function CreateLookup(ByRef SourceKeys As Variant, ByRef SourceValues As Variant, ByRef DestKeys As Variant, Optional ByVal KeyAbsolute As ExColumnRow = exColumn) As String
    If IsArrayNotRange(SourceKeys) And IsArrayNotRange(DestKeys) Then
        CreateLookup = CreateArrayLookup(SourceKeys, SourceValues, DestKeys)
    Else
        Dim rSrcKeys As Range, rSrcValues As Range
        Dim MatchAddr As String
        
        Set rSrcKeys = GetRange(SourceKeys)
        Set rSrcValues = GetRange(SourceValues)
        
        If IsArrayNotRange(DestKeys) Then
            MatchAddr = ConcatRangeAddr(DestKeys, Absolute:=KeyAbsolute)
        Else
            MatchAddr = GetRange(DestKeys).Range("A1").Address(RowAbsolute:=HasFlag(KeyAbsolute, exRow), ColumnAbsolute:=HasFlag(KeyAbsolute, exColumn), External:=True)
        End If
        
        CreateLookup = "INDEX(" & rSrcValues.Address(RowAbsolute:=True, ColumnAbsolute:=True, External:=True) & ", MATCH(" & MatchAddr & ", " & rSrcKeys.Address(RowAbsolute:=True, ColumnAbsolute:=True, External:=True) & ", 0))"
    End If
End Function

'
' Creates an INDEX/MATCH with two lookups (matches) for row and column respectively
'
Function CreateDualLookup(ByRef SourceRange As Variant, ByRef SourceRows As Variant, ByRef SourceColumns As Variant, ByRef DestRow As Variant, ByRef DestColumn As Variant)
    Dim rSrc As Range, rSrcRows As Range, rSrcCols As Range
    Dim rDestRow As Range, rDestCol As Range
    
    Set rSrc = GetRange(rSrc)
    Set rSrcRows = GetRange(SourceRows)
    Set rSrcCols = GetRange(SourceColumns)
    Set rDestRow = GetRange(DestRow)
    Set rDestCol = GetRange(DestColumn)
    
    CreateDualLookup = "=INDEX(" & rSrc.Address(External:=True) & _
                       ", MATCH(" & rDestRow.Address(RowAbsolute:=False) & ", " & rSrcRows.Address(External:=True) & ", 0)" & _
                       ", MATCH(" & rDestRow.Address(ColumnAbsolute:=False) & ", " & rSrcCols.Address(External:=True) & ", 0))"
End Function

Private Sub EnsureTableSize(ByRef rSrc As Range, ByRef Dest As Variant, ByRef rDest As Range)
    Dim loTable As ListObject, TableLast As Long, DestLast As Long
    Set loTable = TryGetTable(Dest)
    If Not (loTable Is Nothing) Then
        TableLast = FindLastRow(GetBodyRange(loTable))
        DestLast = rDest.Row + rDest.Rows.Count - 1
        If DestLast > TableLast Then
            ResizeTable loTable, loTable.Range.Resize(RowSize:=DestLast - TableLast + 1)
        End If
    End If
End Sub

'
' Exclude a range from another range (the inverse of Intersect)
' DO NOT USE THIS ON CELLS THAT HAVE VALIDATION
'
Function Exclude(ByRef Range1 As Variant, ByRef Range2 As Variant) As Range
    Dim R1 As Range, R2 As Range, Overlap As Range
    Set R1 = GetRange(Range1)
    Set R2 = GetRange(Range2)
    Debug.Assert SafeSpecial(R1, xlCellTypeSameValidation) Is Nothing And SafeSpecial(R2, xlCellTypeSameValidation) Is Nothing
    
    With Union(R1, R2).Validation
        .Add xlValidateInputOnly, 1
        Intersect(R1, R2).Validation.Delete
        Set Exclude = SafeSpecial(R1, xlCellTypeAllValidation)
        .Delete
    End With
    
End Function

'
' Copies ranges without putting the copied data on the clipboard. With a large number of areas
' in a multi-area range, CopyMult() may be faster.
'
' Ranges - expected to be an even numbered array, with every even index containing the source range
'          and every odd index containing the destination for that source
'
' returns - the largest number of Rows copied
'
Function FastCopy(ParamArray Ranges() As Variant) As Long
    FastCopy = FastCopyEx(GetArray(Ranges))
End Function

'
' Copies ranges without putting the copied data on the clipboard. With a large number of areas
' in a multi-area range, CopyMult() may be faster.
'
' Ranges - expected to be an even numbered array, with every even index containing the source range
'          and every odd index containing the destination for that source
' CellType - the type of cell to copy
' NumberFormat - the number format to use
' Append - place the copied values at the bottom of the destination range rather than overwriting it
'
' returns - the largest number of Rows copied
'
Function FastCopyEx(ByRef Ranges As Variant, Optional ByVal CellType As XlCellType, Optional ByVal NumberFormat As String, Optional ByVal Append As Boolean = False, Optional ByVal ForceColumnVisible As Boolean = False, Optional ByVal StackSources As Boolean = True) As Long
    Dim n As Long
    
    For n = LBound(Ranges) To UBound(Ranges) Step 2
        Dim rSrc As Range, rDest As Range
        Dim rSrcArea As Range, rDestArea As Range
        Dim srcRows As Long, bHidden As Boolean
        
        Set rSrc = ShrinkToFit(Ranges(n), exRow)
        Set rDest = GetRange(Ranges(n + 1))
        
        SortRange rSrc
        
        If Append Then Set rDest = RangeBelow(rDest, FindLast:=False)
        If ForceColumnVisible Then bHidden = ShowColumn(rSrc)
        
        If CellType <> Empty Then
            Set rSrc = SafeSpecial(rSrc, CellType)
            If rSrc Is Nothing Then GoTo continue
        End If
        
        Let srcRows = CountRows(rSrc, AssumeOrder:=True)
        rDest.ClearContents
        If srcRows <> rDest.Rows.Count Then Set rDest = rDest.Resize(RowSize:=srcRows)
        
        For Each rSrcArea In rSrc.Areas
            If rDestArea Is Nothing Then
                Set rDestArea = rDest.Resize(RowSize:=rSrcArea.Rows.Count, ColumnSize:=rSrcArea.Columns.Count)
            Else
                If StackSources Then
                    Set rDestArea = rDestArea.Offset(RowOffset:=rDestArea.Rows.Count).Resize(rSrcArea.Rows.Count, rSrcArea.Columns.Count)
                Else
                    Set rDestArea = rDestArea.Offset(ColumnOffset:=rDestArea.Columns.Count).Resize(rSrcArea.Rows.Count, rSrcArea.Columns.Count)
                End If
            End If
#If SAFE_RANGE_ENABLED Then
            SafeSetValue rSrcArea, rDestArea
#Else
            rDestArea.Value2 = rSrcArea.Value2
#End If
        Next rSrcArea
        
        If NumberFormat <> Empty Then rDest.NumberFormat = NumberFormat
        
        Let FastCopyEx = WorksheetFunction.Max(srcRows, FastCopyEx)
        
        If ForceColumnVisible And bHidden Then HideColumn rSrc
        If Append Then EnsureTableSize rSrc, Ranges(n + 1), rDest
continue:
        Set rDestArea = Nothing
    Next n
End Function

'
' Fills a formula in a given range and then saves the values
'
' Range - the range to fill in with the formula
' Formula - the formula
' Calculate - whether or not the worksheet should be calculated after the values are filled
' Clipboard - whether or not to use the clipboad. Using the clipboard may be slower, but it will ensure the source formatting and data types
' ArrayFormula - The formula given is an array formula
'
#If SHEETEX_VER < 1010 Then
Function FillValues(ByRef Range As Variant, ByVal Formula As String, Optional ByVal Calculate As Boolean = True, Optional ByVal Clipboard As Boolean = False, Optional ByVal ArrayFormula As Boolean = False)
#Else
Function FillValues(ByRef Range As Variant, ByVal Formula As String, Optional ByVal Calculate As Boolean = True, Optional ByVal Clipboard As Boolean = True, Optional ByVal ArrayFormula As Boolean = False)
#End If
    With GetRange(Range)
        With .Range("A1")
            If ArrayFormula Then
                .FormulaArray = Formula
            Else
                .Formula = Formula
            End If
        End With
        If .Rows.Count > 1 Then .FillDown
        If .Columns.Count > 1 Then .FillRight
        
        If Calculate Then .Worksheet.Calculate
        If Clipboard Then
            CopySingle Range, Range, Paste:=xlPasteValues
        Else
            .Value2 = .Value2
        End If
    End With
End Function

'
' Applies a fiilter to a pivot table on a given field
'
Sub FilterPivot(ByRef Pivot As Variant, ByRef Field As Variant, ByRef Criteria As Variant)
    Dim PivotTable As PivotTable, Item As PivotItem
    Dim x As Long, Match As Boolean, CriteriaArray As Variant
        
    If IsArrayNotRange(Criteria) Then
        CriteriaArray = Criteria
    Else
        CriteriaArray = Array(Criteria)
    End If
    
    Set PivotTable = GetPivot(Pivot)
    
    For Each Item In PivotTable.PivotFields(Field).PivotItems
        Match = False
        For x = LBound(CriteriaArray) To UBound(CriteriaArray)
            Match = CriteriaArray(x) Like Item.Name
            If Match Then Exit For
        Next x
        On Error Resume Next
        Item.Visible = Match
        On Error GoTo 0
    Next Item
End Sub


'
' Applies a filter to the specified header/filter.
'
Sub FilterTable(ByRef Range As Variant, ByRef Field As Variant, ByRef Criteria1 As Variant, Optional ByVal Operator As XlAutoFilterOperator = xlAnd, Optional ByRef Criteria2 As Variant, Optional ByRef VisibleDropDown As Variant)
    With GetRange(Range, IncludeHeader:=True)
'        If .Worksheet.AutoFilterMode Then
'            .Worksheet.AutoFilterMode = False
'            .AutoFilter
'        End If
        .AutoFilter GetColumnRange(Range, Field).Column - .Column + 1, Criteria1, Operator, Criteria2, VisibleDropDown
    End With
End Sub

'
' Finds the last column in a given row. If no row is supplied, finds the last column in the range.
' NOTE - merged cells are not supported and could give unexpected results
'
' Range - the range to find the last row
' RowIndex - the index of the row. If this is not specified, all rows are counted
' StopAtTables - Whether or not to stop when a table is reached (i.e., count tables as non-empty cells)
'
' returns - the (absolute) index of last column
'
#If SHEETEX_VER > 1002 Then
Function FindLastColumn(ByRef Range As Variant, Optional ByRef RowIndex As Variant, Optional ByVal StopAtTables As Boolean = False) As Long
#Else
Function FindLastColumn(ByRef Range As Variant, Optional ByRef RowIndex As Variant, Optional ByVal AllowTables As Boolean = False) As Long
    Dim StopAtTables As Boolean
    StopAtTables = AllowTables
#End If
    Dim CurrCell As Range, Hidden As Boolean
    
    If IsMissing(RowIndex) Then
        Dim n As Long, rng As Range
        Set rng = GetRange(Range)
        FindLastColumn = 1
        If rng.Columns.Count >= rng.Rows.Count Then
            If rng.Worksheet.Columns("A").Hidden Then Hidden = ShowColumn(rng.Worksheet.Columns("A"))
            For n = 1 To rng.Rows.Count
                Set CurrCell = rng.Worksheet.Cells(rng.Rows(n).Row, rng.Column + rng.Columns.Count - 1)
                While IsEmpty(CurrCell) And CurrCell.Column <> 1 And (Not StopAtTables Or CurrCell.ListObject Is Nothing)
                    Set CurrCell = CurrCell.End(xlToLeft)
                Wend
                FindLastColumn = WorksheetFunction.Max(FindLastColumn, CurrCell.Column)
            Next n
            If Hidden Then HideColumn rng.Worksheet.Columns("A")
        Else
            For n = rng.Columns.Count To 1 Step -1
                Dim LastRow As Long
                LastRow = FindLastRow(rng, n, StopAtTables)
                Set CurrCell = rng.Cells(LastRow, n)
                If LastRow < rng.Row Or (Not IsEmpty(CurrCell) Or (StopAtTables And Not (CurrCell.ListObject Is Nothing))) Then
                    FindLastColumn = n + rng.Column - 1
                    Exit For
                End If
            Next n
        End If
    Else
        With GetRange(Range)
            Set CurrCell = .Worksheet.Cells(.Rows(RowIndex).Row, .Column + .Columns.Count - 1)
            If .Worksheet.Columns("A").Hidden Then Hidden = ShowColumn(.Worksheet.Columns("A"))
            While IsEmpty(CurrCell) And CurrCell.Column <> 1 And (Not StopAtTables Or CurrCell.ListObject Is Nothing)
                Set CurrCell = CurrCell.End(xlToLeft)
            Wend
            FindLastColumn = CurrCell.Column
            If Hidden Then HideColumn .Worksheet.Columns("A")
        End With
    End If
End Function

'
' Finds the last non-empty row in a range
' NOTE - merged cells are not supported and could give unexpected results
'
' Range - the range to find the last row
' ColumnIndex - the index of the column. If this is not specified, all columns are counted
' StopAtTables - Whether or not to stop when a table is reached (i.e., count tables as non-empty cells)
'
' returns - the (absolute) index of the last row
'
#If SHEETEX_VER > 1002 Then
Function FindLastRow(ByRef Range As Variant, Optional ByRef ColumnIndex As Variant, Optional ByVal StopAtTables As Boolean = False) As Long
#Else
Function FindLastRow(ByRef Range As Variant, Optional ByRef ColumnIndex As Variant, Optional ByVal AllowTables As Boolean = False) As Long
    Dim StopAtTables As Boolean
    StopAtTables = AllowTables
#End If
    Dim CurrCell As Range
    
    If IsMissing(ColumnIndex) Then
        Dim n As Long, rng As Range
        Set rng = GetRange(Range)
        FindLastRow = 1
        If rng.Rows.Count >= rng.Columns.Count Then
            For n = 1 To rng.Columns.Count
                Set CurrCell = rng.Worksheet.Cells(rng.Row + rng.Rows.Count - 1, rng.Columns(n).Column)
                While IsEmpty(CurrCell) And CurrCell.Row <> 1 And (Not StopAtTables Or CurrCell.ListObject Is Nothing)
                    Set CurrCell = CurrCell.End(xlUp)
                Wend
                FindLastRow = WorksheetFunction.Max(FindLastRow, CurrCell.Row)
            Next n
        Else
            For n = rng.Rows.Count To 1 Step -1
                Dim LastColumn As Long
                LastColumn = RelativeColumn(rng, FindLastColumn(rng, n, StopAtTables))
                Set CurrCell = rng.Cells(n, LastColumn)
                If LastColumn < rng.Column Or (Not IsEmpty(CurrCell) Or (StopAtTables And Not (CurrCell.ListObject Is Nothing))) Then
                    FindLastRow = n + rng.Row - 1
                    Exit For
                End If
            Next n
        End If
    Else
        Dim Hidden As Boolean
        With GetRange(Range)
            Set CurrCell = .Worksheet.Cells(.Row + .Rows.Count - 1, .Columns(ColumnIndex).Column)
            Hidden = .Worksheet.Rows(1).Hidden
            If Hidden Then .Worksheet.Rows(1).Hidden = False
            While IsEmpty(CurrCell) And CurrCell.Row <> 1 And (Not StopAtTables Or CurrCell.ListObject Is Nothing)
                Set CurrCell = CurrCell.End(xlUp)
            Wend
            FindLastRow = CurrCell.Row
            .Worksheet.Rows(1).Hidden = Hidden
        End With
    End If
End Function

'
' Finds the first column in a given row. If no row is supplied, finds the first column in the range.
' NOTE - merged cells are not supported and could give unexpected results
'
' Range - the range to find the last row
' RowIndex - the index of the row. If this is not specified, all rows are counted
' StopAtTables - Whether or not to stop when a table is reached (i.e., count tables as non-empty cells)
'
' returns - the (absolute) index of first column
'
Function FindFirstColumn(ByRef Range As Variant, Optional ByRef RowIndex As Variant, Optional ByVal StopAtTables As Boolean = False) As Long
    Dim CurrCell As Range, Hidden As Boolean
    
    If IsMissing(RowIndex) Then
        Dim n As Long, rng As Range
        Set rng = GetRange(Range)
        FindFirstColumn = rng.Columns.Count
        If rng.Columns.Count >= rng.Rows.Count Then
            If rng.Worksheet.Columns(rng.Worksheet.Columns.Count).Hidden Then Hidden = ShowColumn(rng.Worksheet.Columns(rng.Worksheet.Columns.Count))
            For n = 1 To rng.Rows.Count
                Set CurrCell = rng.Worksheet.Cells(rng.Rows(n).Row, rng.Column)
                While IsEmpty(CurrCell) And CurrCell.Column <> rng.Columns.Count And (Not StopAtTables Or CurrCell.ListObject Is Nothing)
                    Set CurrCell = CurrCell.End(xlToRight)
                Wend
                FindFirstColumn = WorksheetFunction.Min(FindFirstColumn, CurrCell.Column)
            Next n
            If Hidden Then HideColumn rng.Worksheet.Columns("A")
        Else
            For n = 1 To rng.Columns.Count
                Dim LastRow As Long
                LastRow = FindFirstRow(rng, n, StopAtTables)
                Set CurrCell = rng.Cells(LastRow, n)
                If LastRow < rng.Row Or (Not IsEmpty(CurrCell) Or (StopAtTables And Not (CurrCell.ListObject Is Nothing))) Then
                    FindFirstColumn = n + rng.Column - 1
                    Exit For
                End If
            Next n
        End If
    Else
        With GetRange(Range)
            Set CurrCell = .Worksheet.Cells(.Rows(RowIndex).Row, .Column)
            If .Worksheet.Columns(.Worksheet.Columns.Count).Hidden Then Hidden = ShowColumn(.Worksheet.Columns(.Worksheet.Columns.Count))
            While IsEmpty(CurrCell) And CurrCell.Column <> .Worksheet.Columns.Count And (Not StopAtTables Or CurrCell.ListObject Is Nothing)
                Set CurrCell = CurrCell.End(xlToRight)
            Wend
            FindFirstColumn = CurrCell.Column
            If Hidden Then HideColumn .Worksheet.Columns(.Worksheet.Columns.Count)
        End With
    End If
End Function

'
' Finds the first non-empty row in a range
' NOTE - merged cells are not supported and could give unexpected results
'
' Range - the range to find the first row
' ColumnIndex - the index of the column. If this is not specified, all columns are counted
' StopAtTables - Whether or not to stop when a table is reached (i.e., count tables as non-empty cells)
'
' returns - the (absolute) index of the first row
'
Function FindFirstRow(ByRef Range As Variant, Optional ByRef ColumnIndex As Variant, Optional ByVal StopAtTables As Boolean = False) As Long
    Dim CurrCell As Range
    
    If IsMissing(ColumnIndex) Then
        Dim n As Long, rng As Range
        Set rng = GetRange(Range)
        FindFirstRow = rng.Worksheet.Rows.Count
        If rng.Rows.Count >= rng.Columns.Count Then
            For n = 1 To rng.Columns.Count
                Set CurrCell = rng.Worksheet.Cells(rng.Row, rng.Columns(n).Column)
                While IsEmpty(CurrCell) And CurrCell.Row <> rng.Worksheet.Rows.Count And (Not StopAtTables Or CurrCell.ListObject Is Nothing)
                    Set CurrCell = CurrCell.End(xlDown)
                Wend
                FindFirstRow = WorksheetFunction.Min(FindFirstRow, CurrCell.Row)
            Next n
        Else
            For n = 1 To rng.Rows.Count
                Dim FirstColumn As Long
                FirstColumn = RelativeColumn(rng, FindFirstColumn(rng, n, StopAtTables))
                Set CurrCell = rng.Cells(n, FirstColumn)
                If FirstColumn < rng.Column Or (Not IsEmpty(CurrCell) Or (StopAtTables And Not (CurrCell.ListObject Is Nothing))) Then
                    FindFirstRow = n + rng.Row - 1
                    Exit For
                End If
            Next n
        End If
    Else
        Dim Hidden As Boolean
        With GetRange(Range)
            Set CurrCell = .Worksheet.Cells(.Row, .Columns(ColumnIndex).Column)
            Hidden = .Worksheet.Rows(1).Hidden
            If Hidden Then .Worksheet.Rows(1).Hidden = False
            While IsEmpty(CurrCell) And CurrCell.Row <> .Worksheet.Rows.Count And (Not StopAtTables Or CurrCell.ListObject Is Nothing)
                Set CurrCell = CurrCell.End(xlDown)
            Wend
            FindFirstRow = CurrCell.Row
            .Worksheet.Rows(1).Hidden = Hidden
        End With
    End If
End Function

Private Function GenerateTableName(ByRef wb As Workbook) As String
    Dim loTable As ListObject, ws As Worksheet
    Dim n As Long, Found As Boolean
    Let n = 1
    Do
        For Each ws In wb.Worksheets
            For Each loTable In ws.ListObjects
                If StrComp(loTable.Name, "Table" & n, vbTextCompare) = 0 Then
                    Found = True
                    GoTo Break_Out
                End If
            Next loTable
        Next ws
Break_Out:
        If Not Found Then GenerateTableName = "Table" & n
        n = n + 1
        Found = False
    Loop While GenerateTableName = ""
End Function

'
' Returns an array in the case that a ParamArray function was passed an array
'
Function GetArray(ParamArray arr() As Variant) As Variant
    If UBound(arr) < LBound(arr) Then Exit Function
    
    If IsArrayNotRange(arr(LBound(arr))) Then
        Dim newarr As Variant
        LetOrSet newarr, arr(LBound(arr))
        If LBound(newarr) = UBound(newarr) Then
            If IsArrayNotRange(newarr(LBound(newarr))) Then
                LetOrSet GetArray, newarr(LBound(newarr))
                Exit Function
            End If
        End If
        LetOrSet GetArray, newarr
    Else
        GetArray = arr
    End If
End Function

Private Function GetBodyRange(ByRef Range As Variant) As Range
    If Range.DataBodyRange Is Nothing Then
        Set GetBodyRange = Range.Range.Resize(Range.Range.Rows.Count - 1).Offset(RowOffset:=1) ' if no data body, use range below the header
    Else
        Set GetBodyRange = Range.DataBodyRange
    End If
End Function

'
' Gets a column as a range
' You can use a column header in the case of a list object, or a column letter or number.
' Column headers will take precedence over column letters
'     (e.g. a table with a column named PO will take precedence over $PO:$PO)
'
Function GetColumnRange(ByRef Range As Variant, ByRef Column As Variant) As Range
    If TypeName(Column) = "String" Then
        Set GetColumnRange = ListObjCol(GetRange(Range), Column)
        If GetColumnRange Is Nothing Then
            Set GetColumnRange = GetRange(Range).Columns(Column)
        End If
    ElseIf TypeName(Column) = "Range" Then
        Set GetColumnRange = Column
    ElseIf IsNumeric(Column) Then
        Set GetColumnRange = GetRange(Range).Columns(Column)
    Else
        Err.Raise 5
    End If
End Function

'
' Alias for GetColumnRange
'
Function GCR(ByRef Range As Variant, ByRef Column As Variant) As Range
    Set GCR = GetColumnRange(Range, Column)
End Function

'
' Converts the given value into a PivotTable object
'
Function GetPivot(ByRef Pivot As Variant) As PivotTable
    ' this is in the order that I think is most likely to happen
    Select Case TypeName(Pivot)
        Case Is = "PivotTable"
            Set GetPivot = Pivot
        Case Is = "Range"
            Set GetPivot = Pivot.PivotTable
        Case Is = "String"
            Set GetPivot = ThisWorkbook.PivotTables(Pivot)
        Case Else
            Set GetPivot = Nothing
    End Select
End Function

'
' Converts a String, ListObject or Worksheet to a Range
'
#If SHEETEX_VER > 1007 Then
Function GetRange(ByRef Range As Variant, Optional ByVal IncludeHeader As Boolean = False) As Range
#Else
Function GetRange(ByRef Range As Variant, Optional ByVal IncludeHeader As Boolean = True) As Range
#End If
    Set GetRange = TryGetRange(Range, IncludeHeader)
    ' check if the Range object could be converted
    If GetRange Is Nothing Then Err.Raise 5, Source:="Range", Description:="Argument of type '" & TypeName(Range) & "' cannot be converted to a Range."
End Function

'
' Gets a table from the workbook
' table - the table to look for. if this is a string, the table with that name will be returned
'         if this is a ListObject, the same object is returned. otherwise an error is raised
'
Function GetTable(ByRef Table As Variant) As ListObject
    Set GetTable = TryGetTable(Table)
    If GetTable Is Nothing Then Err.Raise 5, "Table", "Could not convert " & TypeName(Table) & " to a ListObject"
End Function

Private Function HasFlag(ByVal Value As Variant, ByVal Flag As Variant)
    HasFlag = (Value And Flag) = Flag
End Function

'
' Hides a single column or the entire group that the column is in
'
Function HideColumn(ByRef Column As Variant) As Boolean
    With GetRange(Column).EntireColumn
        If Not .Hidden Then
            If .OutlineLevel > 1 Then
                .ShowDetail = False
            Else
                .Hidden = True
            End If
            HideColumn = True
        Else
            HideColumn = False
        End If
    End With
End Function

Private Property Get IsArrayNotRange(ByRef Range As Variant) As Boolean
    IsArrayNotRange = IsArray(Range) And TypeName(Range) <> "Range"
End Property

'
' Keeps the current format of a range and removes conditional formats associated with it
'
Sub KeepCurrentFormat(ByRef Range As Variant)
    Dim Cell As Range
    With GetRange(Range)
        For Each Cell In .Cells
            With Cell
                .Interior.Color = .DisplayFormat.Interior.Color
                .Font.FontStyle = .DisplayFormat.Font.FontStyle
                .Font.Strikethrough = .DisplayFormat.Font.Strikethrough
                .Font.Color = .DisplayFormat.Font.Color
                .Font.Bold = .DisplayFormat.Font.Bold
                .Font.Italic = .DisplayFormat.Font.Italic
                .Font.Underline = .DisplayFormat.Font.Underline
            End With
        Next Cell

        .FormatConditions.Delete
    End With
End Sub

Private Function ListObjCol(ByRef Range As Range, ByRef Column As Variant) As Range
    Dim loTable As ListObject
    Set loTable = TryGetTable(Range)
    If loTable Is Nothing Then GoTo catch
    On Error GoTo catch
    Set ListObjCol = loTable.ListColumns(Column).DataBodyRange
    On Error GoTo 0
    If ListObjCol Is Nothing Then
        Set ListObjCol = loTable.ListColumns(Column).Range
        Set ListObjCol = ListObjCol.Offset(1).Resize(1)
    End If
    With ListObjCol
        Dim oRange As Range, StartRow As Long
        Set oRange = GetRange(Range, IncludeHeader:=False)
        Let StartRow = WorksheetFunction.Max(oRange.Row, .Row)
#If SAFE_RANGE_ENABLED Then
        Set ListObjCol = SafeRange(.Worksheet, StartRow, .Column, Height:=oRange.Row + oRange.Rows.Count - StartRow)
#Else
        Set ListObjCol = .Worksheet.Cells(StartRow, .Column).Resize(RowSize:=oRange.Row + oRange.Rows.Count - StartRow)
#End If
    End With
    Exit Function
catch:
    Set ListObjCol = Nothing
End Function

'
' Gets the largest number in a series of numbers
'
Property Get Max(ParamArray args() As Variant) As Variant
    If UBound(args) = LBound(args) + 1 Then  ' most likely case
        Max = WorksheetFunction.Max(args(LBound(args)), args(UBound(args)))
    Else
        Dim i As Integer
        If UBound(args) < LBound(args) Then Exit Property
        args = GetArray(args)
        If UBound(args) < LBound(args) Then Exit Property
        
        Let Max = args(LBound(args))
        
        For i = LBound(args) + 1 To UBound(args)
            Max = WorksheetFunction.Max(args(i), Max)
        Next i
    End If
End Property

'
' Gets the maximum number of columns in a range with one or more areas.
' If no range is supplied, the current sheet is used
'
Property Get MaxColumn(Optional ByRef Range As Variant) As Long
    If IsMissing(Range) Then
        Static MaxStatic As Long
        If MaxStatic = Empty Then MaxStatic = ActiveSheet.Range("1:1").Columns.Count
        MaxColumn = MaxStatic
    Else
        Dim Area As Range
        For Each Area In GetRange(Range).Areas
            MaxColumn = WorksheetFunction.Max(MaxColumn, Area.Columns.Count)
        Next Area
    End If
End Property

'
' Gets the maximum number of rows in a range with one or more areas.
' If no range is supplied, the current sheet is used
'
Property Get MaxRow(Optional ByRef Range As Variant) As Long
    If IsMissing(Range) Then
        Static MaxStatic As Long
        If MaxStatic = Empty Then MaxStatic = ActiveSheet.Range("A:A").Rows.Count
        MaxRow = MaxStatic
    Else
        Dim Area As Range
        For Each Area In GetRange(Range).Areas
            MaxRow = WorksheetFunction.Max(MaxRow, Area.Rows.Count)
        Next Area
    End If
End Property

'
' Gets the smallest number in a series of numbers
'
Property Get Min(ParamArray args() As Variant) As Variant
    If UBound(args) = LBound(args) + 1 Then  ' most likely case
        Min = WorksheetFunction.Min(args(LBound(args)), args(UBound(args)))
    Else
        Dim i As Integer
        If UBound(args) < LBound(args) Then Exit Property
        args = GetArray(args)
        If UBound(args) < LBound(args) Then Exit Property
        
        Let Min = args(LBound(args))
        
        For i = LBound(args) + 1 To UBound(args)
            Min = WorksheetFunction.Min(args(i), Min)
        Next i
    End If
End Property

'
' Gets the Missing value
'
Property Get Missing() As Variant
    Missing = SheetXArg
End Property

Private Property Get SheetXArg(Optional ByRef arg As Variant) As Variant
    SheetXArg = arg
End Property

'
' Finds the first index of a given value in a range
'
Function MMatch(ByRef Value As Variant, ByRef Range As Variant, Optional ByRef DefaultValue As Variant) As Variant
    If IsMissing(DefaultValue) Then DefaultValue = CVErr(xlErrNA)
    
    On Error GoTo catch
    MMatch = WorksheetFunction.Match(Value, GetRange(Range), 0)
    On Error GoTo 0
    
    If IsError(MMatch) Then GoTo catch
    Exit Function
catch:
    LetOrSet MMatch, DefaultValue
End Function

'
' Creates an array with numbers through
'
Function NumberArray(ByRef Lower As Variant, ByRef Upper As Variant, Optional ByRef Increment As Variant) As Variant
    Dim x As Variant, NewArray() As Variant
    Dim Index As Long, Count As Variant, ArrSize As Long
    If IsMissing(Increment) Then Increment = IIf(Upper < Lower, -1, 1)
    Count = Abs(Upper - Lower) + 1
    ArrSize = WorksheetFunction.Ceiling(Abs(Count / Increment), 1)
    If Count = 0 Or ArrSize <= 0 Then
        NumberArray = Array()
    Else
        ReDim NewArray(0 To ArrSize - 1)
        Index = 0
        For x = Lower To Upper Step Increment
            NewArray(Index) = x
            Index = Index + 1
        Next x
        NumberArray = NewArray
    End If
End Function

Private Sub QuickSort(ByRef arr As Variant, ByVal Low As Long, ByVal High As Long)
    Dim pvt As Variant
    Dim tmpLo As Long, tmpHi As Long
    tmpLo = Low
    tmpHi = High
    LetOrSet pvt, arr((Low + High) \ 2)
    While tmpLo <= tmpHi
        While arr(tmpLo) < pvt And tmpLo < High
            tmpLo = tmpLo + 1
        Wend
        While arr(tmpHi) > pvt And tmpHi > Low
            tmpHi = tmpHi - 1
        Wend
        If tmpLo <= tmpHi Then
            Swap arr(tmpLo), arr(tmpHi)
            tmpLo = tmpLo + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    
    If Low < tmpHi Then QuickSort arr, Low, tmpHi
    If tmpLo < High Then QuickSort arr, tmpLo, High
End Sub

Private Sub QuickSortRev(ByRef arr As Variant, ByVal Low As Long, ByVal High As Long)
    Dim pvt As Variant
    Dim tmpLo As Long, tmpHi As Long
    tmpLo = Low
    tmpHi = High
    LetOrSet pvt, arr((Low + High) \ 2)
    While tmpLo <= tmpHi
        While arr(tmpLo) > pvt And tmpLo < High
            tmpLo = tmpLo + 1
        Wend
        While arr(tmpHi) < pvt And tmpHi > Low
            tmpHi = tmpHi - 1
        Wend
        If tmpLo <= tmpHi Then
            Swap arr(tmpHi), arr(tmpLo)
            tmpLo = tmpLo + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    
    If Low < tmpHi Then QuickSortRev arr, Low, tmpHi
    If tmpLo < High Then QuickSortRev arr, tmpLo, High
End Sub

'
' Gets a range immediately below another range
'
' FindLast - whether or not to find the first non empty row in the range
'
Function RangeBelow(ByRef Range As Variant, Optional ByVal FindLast As Boolean = True, Optional ByVal Height As Long) As Range
    Dim rSrc As Range
    Dim nHeight As Long
    If FindLast Then
        Dim LastRow As Long
        Set rSrc = GetRange(Range)
        LastRow = FindLastRow(rSrc)
        nHeight = rSrc.Rows.Count
        Set rSrc = ROffset(rSrc, RowOffset:=LastRow - rSrc.Row, Height:=1)
    Else
        Set rSrc = GetRange(Range)
        nHeight = rSrc.Rows.Count
    End If
    If Height = Empty Then Height = nHeight
    Set RangeBelow = ROffset(rSrc, RowOffset:=rSrc.Rows.Count, Height:=Height)
End Function

'
' Converts a Range object into an array with bounds 1 to Range.Cells.Count
'
Function RangeToArray(ByRef Range As Variant, Optional ByVal SkipBlanks As Boolean = False) As Variant
    Dim arr() As Variant, n As Long, nIdx As Long
    With GetRange(Range)
        ReDim arr(1 To .Cells.Count)
        nIdx = LBound(arr)
        For n = LBound(arr) To UBound(arr)
            If Not IsEmpty(.Cells(n)) Or Not SkipBlanks Then
                arr(nIdx) = .Cells(n).Value
                nIdx = nIdx + 1
            End If
        Next n
        If UBound(arr) <> nIdx - 1 Then ReDim Preserve arr(1 To nIdx - 1)
    End With
    RangeToArray = arr
End Function

'
' Gets the relative row from an absolute row
'
Property Get RelativeColumn(ByRef RelativeTo As Variant, ByRef ColumnAbsolute As Variant) As Long
    With GetRange(RelativeTo)
        RelativeColumn = GetColumnRange(.Worksheet, ColumnAbsolute).Column - .Column + 1
    End With
End Property

'
' Gets the relative row from an absolute row
'
Property Get RelativeRow(ByRef RelativeTo As Variant, ByRef RowAbsolute As Variant) As Long
    With GetRange(RelativeTo)
        RelativeRow = .Worksheet.Rows(RowAbsolute).Row - .Row + 1
    End With
End Property

'
' Remove an element from an array at a given index
'
Function RemoveAt(ByRef arr As Variant, ByVal Index As Long) As Variant
    Dim NewArray() As Variant, x As Long
    ReDim NewArray(LBound(arr) To UBound(arr) - 1)
    
    For x = LBound(arr) To Index - 1
        LetOrSet NewArray(x), arr(x)
    Next x
    For x = Index + 1 To UBound(arr)
        LetOrSet NewArray(x - 1), arr(x)
    Next x
    
    RemoveAt = NewArray
End Function

'
' Resizes a table and deletes any rows that are no longer in the table
' Headers must be in the same row and the resulting table must overlap the original table range.
'
Sub ResizeTable(ByRef Table As Variant, ByRef NewRange As Variant)
    Dim NewHeight As Long, NewWidth As Long
    Dim rng As Range
    
    Set rng = GetRange(NewRange, IncludeHeader:=True)
    NewHeight = rng.Rows.Count
    NewWidth = rng.Columns.Count
    
    
    With GetTable(Table)
        If rng.Row <> .HeaderRowRange.Row Or Intersect(rng, .Range) Is Nothing Then
            Err.Raise 1004, Description:="Headers must be in the same row and the resulting table must overlap the original table range."
        End If
        
        If NewHeight > 0 Then
            If Not (.DataBodyRange Is Nothing) Then
                If NewHeight < .DataBodyRange.Rows.Count Then RShrink(.DataBodyRange, Height:=-NewHeight).Clear
            End If
        End If
        If NewWidth > 0 Then
            If Not (.DataBodyRange Is Nothing) Then
                If NewWidth < .DataBodyRange.Columns.Count Then RShrink(.DataBodyRange, Width:=-NewWidth).Clear
            End If
        End If
        .Resize rng
    End With
End Sub

'
' Resises a table or a worksheet by removing empty rows or columns
'
#If SHEETEX_VER > 1000 And SHEETEX_VER < 1005 Then
Function ResizeToFit(ByRef Range As Variant, Optional ByVal Shrink As ExColumnRow = exRow) As Range
#Else
Function ResizeToFit(ByRef Range As Variant, Optional ByVal Shrink As ExColumnRow = exBoth) As Range
#End If
    Set ResizeToFit = ShrinkToFit(Range, Shrink)

    Select Case TypeName(Range)
        Case Is = "ListObject"
            If ResizeToFit.Rows.Count > 1 Then
                ResizeTable Range, ResizeToFit
            Else
                ResizeTable Range, Range.HeaderRowRange.Resize(RowSize:=2)
            End If
        Case Is = "Worksheet"
            Dim ws As Worksheet
            Dim LastRow As Long, LastCol As Long
            Set ws = Range
            With ResizeToFit
                LastRow = .Rows.Count + .Row - 1
                LastCol = .Columns.Count + .Column - 1
            End With
                        
            If ws.UsedRange.Rows.Count > LastRow Then
                ws.Range(ws.Cells(LastRow + 1, 1), ws.Cells(ws.UsedRange.Rows.Count, 1)).EntireRow.Delete
            End If
            
            If ws.UsedRange.Columns.Count > LastCol Then
                With ws.UsedRange
                    ws.Range(ws.Cells(1, LastCol + 1), ws.Cells(.Rows.Count, .Columns.Count)).EntireColumn.Delete xlToLeft
                End With
            End If
    End Select
End Function

'
' Expands a given range
' Height - the number of Rows to expand. Use negative numbers to expand upward.
' Width - the number of columns to expand. Use negative numbers to expand left.
'
Function RExpand(ByRef Range As Variant, Optional ByVal Width As Long, Optional ByVal Height As Long) As Range
    Dim nRow As Long, nCol As Long, oRange As Range
    Set oRange = GetRange(Range)
    With oRange
        Let nRow = IIf(Height < 0, Height, 0)
        Let nCol = IIf(Width < 0, Width, 0)
#If SAFE_RANGE_ENABLED Then
        Set RExpand = ROffsetInternal(oRange, RowOffset:=nRow, ColumnOffset:=nCol, Width:=Math.Abs(Width) + .Columns.Count, Height:=Math.Abs(Height) + .Rows.Count)
#Else
        Set RExpand = .Offset(nRow, nCol).Resize(ColumnSize:=Math.Abs(Width) + .Columns.Count, RowSize:=Math.Abs(Height) + .Rows.Count)
#End If
    End With
End Function

'
' Copies a column based on a key, replacing any values that were there previously
'
' SourceKeys   - the range that contains keys on the source table
' SourceValues - the range of values to copy to the destination
' DestKeys     - the range that contains keys on the destination table
' DestValues   - the range to place the copied values
'
#If SHEETEX_VER < 1010 Then
Sub JoinReplace(ByRef SourceKeys As Variant, ByRef SourceValues As Variant, ByRef DestKeys As Variant, ByRef DestValues As Variant, Optional ByVal DefaultValue As Variant, Optional ByVal Calculate As Boolean = True, Optional ByVal Clipboard As Boolean = False)
#Else
Sub JoinReplace(ByRef SourceKeys As Variant, ByRef SourceValues As Variant, ByRef DestKeys As Variant, ByRef DestValues As Variant, Optional ByVal DefaultValue As Variant, Optional ByVal Calculate As Boolean = True, Optional ByVal Clipboard As Boolean = True)
#End If
    Dim Formula As String, ArrayFormula As Boolean
    
    Formula = CreateLookup(SourceKeys, SourceValues, DestKeys)
    ArrayFormula = IsArrayNotRange(SourceKeys) And IsArrayNotRange(DestKeys)
    
    If IsMissing(DefaultValue) Then
#If SHEETEX_VER > 1005 Then
        Formula = "=" & Formula
#Else
        Formula = "=IFNA(" & Formula & ", """")"
#End If
    Else
        If TypeName(DefaultValue) = "String" Then DefaultValue = """" & Replace(DefaultValue, """", """""") & """"
        Formula = "=IFNA(" & Formula & ", " & DefaultValue & ")"
    End If
    
    FillValues DestValues, Formula, Calculate, Clipboard, ArrayFormula
End Sub

'
' Performs an index match (vlookup)
' LookupValue - the key to search for in the LookupRange
' LookupRange - the range of keys in which to look for the LookupValue
' ValueRange - the range of values corresponding to the keys in the LookupRange
' DefaultValue - the default value used in case a match was not found
'
Function RLookup(ByRef LookupValue As Variant, ByRef LookupRange As Variant, ByRef ValueRange As Variant, Optional ByRef DefaultValue As Variant) As Variant
    Dim nMatch As Variant
    
    nMatch = MMatch(LookupValue, GetRange(LookupRange))
    If IsError(nMatch) Then
        RLookup = DefaultValue
    Else
        RLookup = GetRange(ValueRange).Rows(nMatch).Value2
    End If
End Function

'
' Gets the offset of a given range
' RowOffset - the number of Rows from the topmost cells to offset
' ColumnOffset - the number of columns from the leftmost cells to offset
' Width - the number of columns in the resulting range
' Height - the number of Rows in the resulting range
'
Function ROffset(ByRef Range As Variant, Optional ByVal RowOffset As Long, Optional ColumnOffset As Long, Optional ByVal Width As Long, Optional ByVal Height As Long) As Range
    Dim oRange As Range, Area As Range
    Set oRange = GetRange(Range)
    For Each Area In oRange.Areas
        If ROffset Is Nothing Then
            Set ROffset = ROffsetInternal(Area, RowOffset, ColumnOffset, Width, Height)
        Else
            Set ROffset = Union(ROffset, ROffsetInternal(Area, RowOffset, ColumnOffset, Width, Height))
        End If
    Next Area
End Function

Private Function ROffsetInternal(ByRef Range As Range, ByVal RowOffset As Long, ByVal ColumnOffset As Long, ByVal Width As Long, ByVal Height As Long)
    With Range
        If Width = Empty Then Width = .Columns.Count
        If Height = Empty Then Height = .Rows.Count
#If SAFE_RANGE_ENABLED Then
        Set ROffsetInternal = SafeRange(.Worksheet, .Row + RowOffset, .Column + ColumnOffset, Width, Height)
#Else
        Set ROffsetInternal = .Resize(Height, Width).Offset(RowOffset, ColumnOffset)
#End If
    End With
End Function

'
' Converts a collection to an array. Keys are ignored.
'
Function ListToArray(ByRef List As Collection) As Variant
    Dim arr() As Variant, n As Long
    ReDim arr(1 To List.Count)
    For n = LBound(arr) To UBound(arr)
        LetOrSet arr(n), List(n)
    Next n
    ListToArray = arr
End Function

'
' Shrinks a given range
' Height - the number of Rows to shrink. Use negative numbers to shrink from top.
' Width - the number of columns to expand. Use negative numbers to shrink left.
'
Function RShrink(ByRef Range As Variant, Optional ByVal Width As Long, Optional ByVal Height As Long) As Range
    Dim nRow As Long, nCol As Long, oRange As Range
    Set oRange = GetRange(Range)
    With oRange
        Let nRow = IIf(Height < 0, -Height, 0)
        Let nCol = IIf(Width < 0, -Width, 0)
#If SAFE_RANGE_ENABLED Then
        Set RShrink = ROffsetInternal(oRange, RowOffset:=nRow, ColumnOffset:=nCol, Width:=.Columns.Count - Math.Abs(Width), Height:=.Rows.Count - Math.Abs(Height))
#Else
        Set RShrink = .Offset(nRow, nCol).Resize(ColumnSize:=.Columns.Count - Math.Abs(Width), RowSize:=.Rows.Count - Math.Abs(Height))
#End If
    End With
End Function

'
' Takes a set number of rows/columns from a given range
' Rows - the number of rows to take from the range. Use negative numbers to shrink from the bottom.
' Columns - the number of columns to take from the range. Use negative numbers to take from the left.
'
Function RTake(ByRef Range As Variant, Optional ByVal Rows As Long, Optional ByVal Columns As Long) As Range
    Set RTake = GetRange(Range)
    With RTake
        If Rows < 0 Then
            Set RTake = .Rows(.Rows.Count).Offset(RowOffset:=Rows + 1).Resize(RowSize:=Math.Abs(Rows))
        ElseIf Rows > 0 Then
            Set RTake = .Resize(RowSize:=Rows)
        End If
    End With
    With RTake
        If Columns < 0 Then
            Set RTake = .Columns(.Columns.Count).Offset(ColumnOffset:=Columns + 1).Resize(ColumnSize:=Math.Abs(Columns))
        ElseIf Columns > 0 Then
            Set RTake = .Resize(ColumnSize:=Columns)
        End If
    End With
End Function

'
' Gets a range without going past any bounds
'
' Range - the starting range or worksheet
' RowIndex - the absolute index of the first row for the new range
' ColumnIndex - the absolute index of the first column for the new range
' Width - the number of columns in the new range
' Height - the number of rows in the new range
'
Function SafeRange(ByRef Range As Variant, ByVal RowIndex As Variant, ByVal ColumnIndex As Variant, Optional ByVal Width As Long = 1, Optional ByVal Height As Long = 1) As Range
    If IsNumeric(RowIndex) Then RowIndex = WorksheetFunction.Max(WorksheetFunction.Min(RowIndex, MaxRow), 1)
    If IsNumeric(ColumnIndex) Then ColumnIndex = WorksheetFunction.Max(WorksheetFunction.Min(ColumnIndex, MaxColumn), 1)

    With GetRange(Range).Cells(RowIndex, ColumnIndex)
        Set SafeRange = SafeRangeInternal(GetRange(.Worksheet).Resize(MaxRow, MaxColumn), .Row, .Column, Width, Height)
    End With
End Function

Function SafeRangeInternal(ByRef Range As Range, ByVal RowIndex As Long, ByVal ColumnIndex As Long, Optional ByVal Width As Long = 1, Optional ByVal Height As Long = 1) As Range
    RowIndex = WorksheetFunction.Min(Range.Rows.Count, WorksheetFunction.Max(RowIndex, 1))
    ColumnIndex = WorksheetFunction.Min(Range.Columns.Count, WorksheetFunction.Max(ColumnIndex, 1))
    
    With Range.Cells(RowIndex, ColumnIndex)
        Dim LastRow As Long, LastCol As Long
        LastRow = WorksheetFunction.Max(WorksheetFunction.Min(.Row + Height - 1, MaxRow), .Row)
        LastCol = WorksheetFunction.Max(WorksheetFunction.Min(.Column + Width - 1, MaxColumn), .Column)
        Set SafeRangeInternal = .Worksheet.Range(.Worksheet.Cells(.Row, .Column), .Worksheet.Cells(LastRow, LastCol))
    End With
End Function

Private Function SafeSetValue(ByVal Source As Range, ByVal Destination As Range, Optional ByVal Fit As Boolean = True) As Boolean
    On Error GoTo catch
    Destination.Value2 = Source.Value2
    On Error GoTo 0
    SafeSetValue = True
    Exit Function
catch:
    If Fit Then
        Set Source = ShrinkToFit(Source, exRow)
        Set Destination = Destination.Resize(RowSize:=Source.Rows.Count)
        Destination.Value2 = Source.Value2 ' if this errors out, it needs to be brought to your attention
    End If
    SafeSetValue = False
End Function

'
' Gets the special cells in a range. If there are no special cells, returns Nothing.
'
Function SafeSpecial(ByRef Range As Variant, ByVal CellType As XlCellType) As Range
    On Error GoTo catch
    Set SafeSpecial = GetRange(Range, IncludeHeader:=False).SpecialCells(CellType)
    Exit Function
catch:
    Set SafeSpecial = Nothing
End Function

'
' Sets all column widths in a range
'
Sub SetColumnWidths(ByRef Range As Variant, Optional ByRef Width As Variant)
    Dim NewWidth As Variant, Column As Range
    NewWidth = IIf(IsMissing(Width), 8.43, Width)
    With GetRange(Range)
        For Each Column In .Columns
            Column.EntireColumn.ColumnWidth = NewWidth
        Next Column
    End With
End Sub

'
' Sets all row heights in a range
'
Sub SetRowHeights(ByRef Range As Variant, Optional ByRef Height As Variant)
    Dim NewHeight As Variant, Row As Range
    NewHeight = IIf(IsMissing(Height), 15, Height)
    With GetRange(Range)
        For Each Row In .Rows
            Row.EntireRow.RowHeight = NewHeight
        Next Row
    End With
End Sub

'
' Shows a group of columns
'
Function ShowColumn(ByRef Column As Variant) As Boolean
    With GetRange(Column).EntireColumn
        If .Hidden Then
            If .OutlineLevel > 1 Then
                .ShowDetail = True
            Else
                .Hidden = False
            End If
            ShowColumn = True
        Else
            ShowColumn = False
        End If
    End With
End Function

'
' Gets a range to exclude empty columns and rows from the end
'
' Shrink - resize the row, column, or both
'
#If SHEETEX_VER > 1000 And SHEETEX_VER < 1005 Then
Function ShrinkToFit(ByRef Range As Variant, Optional ByVal Shrink As ExColumnRow = exRow) As Range
#Else
Function ShrinkToFit(ByRef Range As Variant, Optional ByVal Shrink As ExColumnRow = exBoth) As Range
#End If
    Dim firstRow As Long, firstCol As Long
    Dim LastRow As Long, LastCol As Long
    Dim StopAtTables As Boolean
    StopAtTables = (TypeName(Range) = "Worksheet")
    
#If SAFE_RANGE_ENABLED Then
    Dim rng As Range
    Set rng = GetRange(Range)
    With rng
#Else
    With GetRange(Range)
#End If
        If (Shrink And exRow) = exRow Then
            LastRow = FindLastRow(Range, , StopAtTables)
        Else
            LastRow = .Rows.Count + .Row - 1
        End If
        If (Shrink And exColumn) = exColumn Then
            LastCol = FindLastColumn(Range, , StopAtTables)
        Else
            LastCol = .Columns.Count + .Column - 1
        End If
#If SAFE_RANGE_ENABLED Then
        Set ShrinkToFit = SafeRange(rng.Worksheet, .Row, .Column, Height:=LastRow - .Row + 1, Width:=LastCol - .Column + 1)
#Else
        Set ShrinkToFit = .Resize(WorksheetFunction.Max(LastRow - .Row + 1, 1), WorksheetFunction.Max(LastCol - .Column + 1, 1))
#End If
    End With
End Function

'
' Sorts an array of values in ascending order
' This sorts the array in place. Make a copy if you want to keep the original.
'
Sub SortArray(ByRef UnsortedArray As Variant, Optional ByVal Sort As XlSortOrder = xlAscending)
    If Sort = xlAscending Then
        QuickSort UnsortedArray, LBound(UnsortedArray), UBound(UnsortedArray)
    ElseIf Sort = xlDescending Then
        QuickSortRev UnsortedArray, LBound(UnsortedArray), UBound(UnsortedArray)
    End If
End Sub

Private Sub SortRange(ByRef Range As Range)
    Dim AddrArray() As String, n As Long
    
    If Range.Areas.Count = 1 Then Exit Sub
    ReDim AddrArray(Range.Areas.Count - 1)
    
    For n = LBound(AddrArray) To UBound(AddrArray)
        AddrArray(n) = Range.Areas(n + 1).Address(External:=True)
    Next n
    
    SortArray AddrArray, xlAscending
    
    Set Range = Application.Range(AddrArray(LBound(AddrArray)))
    For n = LBound(AddrArray) + 1 To UBound(AddrArray)
        Set Range = Union(Range, Application.Range(AddrArray(n)))
    Next n
End Sub

'
' Sorts a table on a given column
' Range - the table to sort
' Order - the sort order
' Column - the column to sort on
'
Sub SortTable(ByRef Range As Variant, ByRef Key As Variant, Optional ByVal Order As XlSortOrder = XlSortOrder.xlAscending, Optional ByVal Header As XlYesNoGuess = xlYes)
    With GetRange(Range, IncludeHeader:=True)
        .Sort Key1:=GetColumnRange(Range, Key), Order1:=Order, DataOption1:=xlSortNormal, Header:=Header
    End With
End Sub

'
' Sums a column based on a key
'
' SourceLookup - the range that contains keys on the source table
' SourceValues - the range of values to copy to the destination
' DestLookup   - the range that contains keys on the destination table
' DestValues   - the range to place the copied values
'
Sub SumColumnByKey(ByRef SourceLookup As Variant, ByRef SourceValues As Variant, ByRef DestLookup As Variant, ByRef DestValues As Variant, Optional ByVal Calculate As Boolean = True)
    Dim rDestLookup As Range, rSrcLookup As Range, rDestValues As Range, rSrcValues As Range
    Dim Formula As String

    Set rDestLookup = GetRange(DestLookup)
    Set rSrcLookup = GetRange(SourceLookup)
    Set rDestValues = ROffset(DestValues, Height:=rDestLookup.Rows.Count)
    Set rSrcValues = ROffset(SourceValues, Height:=rSrcLookup.Rows.Count)
    
    'If IsMissing(DefaultValue) Then DefaultValue = """"""
    
    rDestValues.Formula = "=SUMIFS(" & rSrcValues.Address(External:=True) & ", " & rSrcLookup.Address(External:=True) & ", " & rDestLookup.Address(RowAbsolute:=False, External:=True) & ")"
    If Calculate Then rDestValues.Worksheet.Calculate
    
    rDestValues.Value2 = rDestValues.Value2 ' keep values
End Sub

'
' Swaps the values of the given variables
'
Sub Swap(ByRef A As Variant, ByRef B As Variant)
    Dim tmp As Variant
    LetOrSet tmp, A
    LetOrSet A, B
    LetOrSet B, tmp
End Sub

'
' Attempts to convert a variant to a range. If it cannot, Nothing is returned.
'
#If SHEETEX_VER > 1007 Then
Function TryGetRange(ByRef Range As Variant, Optional ByRef IncludeHeader As Boolean = False) As Range
#Else
Function TryGetRange(ByRef Range As Variant, Optional ByRef IncludeHeader As Boolean = True) As Range
#End If
    ' this is in the order that I think is most likely to happen
    Select Case TypeName(Range)
        Case Is = "Range"
            Set TryGetRange = Range
#If SHEETEX_VER > 1007 Then
            If IncludeHeader Then
                Dim loTable As ListObject: Set loTable = TryGetRange.ListObject
                Dim HeaderRange As Range
                If Not (loTable Is Nothing) Then
                    With GetBodyRange(loTable)
                        If .Worksheet Is TryGetRange.Worksheet And .Row = TryGetRange.Row And .Row + .Rows.Count = TryGetRange.Row + TryGetRange.Rows.Count Then
                            Set HeaderRange = loTable.HeaderRowRange.Resize(ColumnSize:=TryGetRange.Columns.Count).Offset(ColumnOffset:=TryGetRange.Column - .Column)
                            Set TryGetRange = Union(HeaderRange, TryGetRange)
                        End If
                    End With
                End If
            End If
#End If
        Case Is = "ListObject", "ListColumn"
            If IncludeHeader Then
                Set TryGetRange = Range.Range
            Else
                Set TryGetRange = GetBodyRange(Range)
            End If
        Case Is = "String"
            Set TryGetRange = Application.Range(Range)
        Case Is = "Worksheet"
            With Range
                Set TryGetRange = .Range("A1", .Cells(.UsedRange.Rows.Count + .UsedRange.Row - 1, .UsedRange.Columns.Count + .UsedRange.Column - 1))
            End With
        Case Is = "PivotTable"
            If Range.TableRange2 Is Nothing Then
                Set TryGetRange = Range.TableRange1
            Else
                Set TryGetRange = Union(Range.TableRange1, Range.TableRange2)
            End If
        Case Else
            Set TryGetRange = Nothing
    End Select
End Function

'
' Attempts to get a ListObject from a given range. If the range cannot be converted to a table, Nothing is returned
'
Function TryGetTable(ByRef Table As Variant) As ListObject
    Select Case TypeName(Table)
        Case "ListObject"
            Set TryGetTable = Table
        Case Else
            Dim rng As Range
            Set rng = TryGetRange(Table)
            If rng Is Nothing Then
                Set TryGetTable = Nothing
            Else
                Set TryGetTable = rng.ListObject
            End If
    End Select
End Function

'
' Attempts to get a sheet at a given index. If no sheet can be found, Nothing is returned
'
Function TryGetSheet(ByRef Workbook As Workbook, ByRef Index As Variant) As Worksheet
    On Error GoTo catch
    Set TryGetSheet = Workbook.Worksheets(Index)
    GoTo cleanup
catch:
    Set TryGetSheet = Nothing
cleanup:
    On Error GoTo 0
End Function

'
' Create a table in a worksheet
' (provided by Michael Cowan)
'
Sub CreateTable(ByRef Range As Variant, Optional Name As String, Optional ByVal Header As XlYesNoGuess = xlGuess, _
                Optional ByVal ShowHeader As Boolean = True, Optional ByVal ShowFilter As Boolean = True, _
                Optional ByVal BandedRows As Boolean = False)
    Dim InputRange As Range
    Set InputRange = GetRange(Range)
    
    With InputRange.Worksheet
        If Name = Empty Then Name = GenerateTableName(.Parent)
        .ListObjects.Add(xlSrcRange, InputRange, , Header).Name = Name
        With .ListObjects(Name)
            .ShowAutoFilterDropDown = ShowFilter
            .ShowHeaders = ShowHeader
            .ShowTableStyleRowStripes = BandedRows
        End With
    End With
End Sub

'
' Search for the index of a given value in a single dimensional array.
' If the value is not found, LBound(Arr)-1 is returned
' (provided by Michael Cowan)
'
Function IndexOf(ByRef arr As Variant, ByRef Value As Variant, Optional ByVal IgnoreCase As Boolean = False) As Long
    Dim i As Long, Match As Boolean
    IndexOf = LBound(arr) - 1
    For i = LBound(arr) To UBound(arr)
        If TypeName(arr(i)) = "String" Then
            Match = StrComp(arr(i), Value, IIf(IgnoreCase, vbTextCompare, vbBinaryCompare)) = 0
        Else
            If IsObject(arr(i)) Then
                Match = (arr(i) Is Value)
            Else
                Match = (arr(i) = Value)
            End If
        End If
        If Match Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

'
' Gets the last cell in a given range
'
Property Get LastCell(ByRef Range As Variant) As Range
    With GetRange(Range)
        Set LastCell = .Worksheet.Cells(.Row + .Rows.Count - 1, .Column + .Columns.Count - 1)
    End With
End Property

'
' Gets the absolute column number of the last column in this range
'
Property Get LastColumn(ByRef Range As Variant) As Long
    With GetRange(Range)
        LastColumn = .Column + .Columns.Count - 1
    End With
End Property

'
' Gets the absolute row number of the last row in this range
'
Property Get LastRow(ByRef Range As Variant) As Long
    With GetRange(Range)
        LastRow = .Row + .Rows.Count - 1
    End With
End Property

'
' Gets the last used cell in a given range
'
Property Get LastUsedCell(ByRef Range As Variant) As Range
    Dim rng As Range, nRow As Long, nCol As Long
    Set rng = GetRange(Range)
    nRow = FindLastRow(rng)
    nCol = FindLastColumn(rng)
    
    Set LastUsedCell = rng.Worksheet.Cells(nRow, nCol)
End Property

'
' Gets the last used Column in a given range
'
Property Get LastUsedColumn(ByRef Range As Variant) As Range
    Dim nLast As Long, rng As Range
    Set rng = GetRange(rng)
    nLast = FindLastColumn(rng)
    With rng
        Set LastUsedColumn = .Columns(nLast - .Column + 1)
    End With
End Property

'
' Gets the last used row in a given range
'
Property Get LastUsedRow(ByRef Range As Variant) As Range
    Dim nLast As Long, rng As Range
    Set rng = GetRange(Range)
    nLast = FindLastRow(rng)
    With rng
        Set LastUsedRow = .Rows(nLast - .Row + 1)
    End With
End Property

Private Sub LetOrSet(ByRef LValue As Variant, ByRef RValue As Variant)
    If IsObject(RValue) Then
        Set LValue = RValue
    Else
        LValue = RValue
    End If
End Sub

'
' Ignores any cell warnings in a given range
'
Sub IgnoreErrors(ByRef Range As Variant, ByRef ErrorIndex As Variant)
    Dim Cell As Range
    With GetRange(Range)
        For Each Cell In .Cells
            Cell.Errors(ErrorIndex).Ignore = True
        Next Cell
    End With
End Sub

'
' Copies the first row of a range down to match a specified size
' If no size is specified, adjacent cells are used to determine it
'
Sub DragDown(ByRef ColumnRange As Variant, Optional ByVal Size As Long, Optional ByVal Header As XlYesNoGuess = xlYes)
    Dim Col As Range, nLast As Long, nOffset As Long, nFirst As Long
    Set Col = GetRange(ColumnRange)
    Let nOffset = IIf(Header = xlYes, 2, 1)
    
    With Col
        Let nLast = FindLastRow(.EntireColumn)
        Let nFirst = FindFirstRow(.EntireColumn)
        With .Worksheet.Range(.Rows(nFirst), .Rows(nLast))
            .Resize(RowSize:=.Rows.Count - nOffset).Offset(RowOffset:=nOffset).Clear
        End With
        
        If Size = Empty Then
            Dim nLeft As Long, nRight As Long
            If Col.Column > 1 Then nLeft = FindLastRow(Col.Resize(ColumnSize:=1).Offset(ColumnOffset:=-Col.Columns.Count))
            If Col.Column < Col.Worksheet.Columns.Count Then nRight = FindLastRow(Col.Resize(ColumnSize:=1).Offset(ColumnOffset:=Col.Columns.Count))
            Size = WorksheetFunction.Max(nLeft, nRight)
        End If
        
        .Rows(nFirst + nOffset - 1).AutoFill .Rows(nFirst + nOffset - 1).Resize(RowSize:=Size - nFirst)
    End With
End Sub
