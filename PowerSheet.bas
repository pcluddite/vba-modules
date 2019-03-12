Attribute VB_Name = "PowerSheet"
'
' PowerSheet
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
' Last update: 7/5/17
' Requires: SheetEx.bas [Rev 1010]
' Description: Apply complex functions to transform data in ranges
'
' Change the POWERSHEET_VER constant to keep certain backward compatability
' Revisions stack on top of one another, i.e. If you roll back to Rev 1000, you will lose
' all features past that version
'
' Revisions:
'   1000 - use old parameter order for ApplyFormula, requiring `Destination` argument
'   1010 - use empty string default value for `TransformRange`, `JoinPreserve`
'   1012 - use old parameter order for TempSheet property
'   1013 - apply no restrictions on temp sheets
'   1015 - do not use clipboard by default
'   1016 - CURRENT VERSION
'
#Const POWERSHEET_VER = 1016

Public Const TEMP_SHEET As String = "PowerSheet.Temp"

Public ClearAfter As Boolean
Private CanUndo As Boolean

'
' Creates a temporary sheet that is xlSheetVeryHidden or returns the existing temp sheet
' Index - the temp sheet index. If no temp sheet exists at that index, it will be created. This is 0 by default.
' Create - whether or not to create the temp sheet if it is not found
'
#If POWERSHEET_VER > 1012 Then
Property Get TempSheet(Optional ByVal Index As Long = 0, Optional ByVal Create As Boolean = True) As Worksheet
#Else
Property Get TempSheet(Optional ByVal Create As Boolean = True, Optional ByVal Index As Long = 0) As Worksheet
#End If
#If POWERSHEET_VER > 1013 Then
    If Index < 0 Then Err.Raise 5, "Temporary sheet " & Index & " has been reserved by the PowerSheet"
#End If
    Set TempSheet = TempSheetInternal(Index, Create)
End Property

Private Property Get TempSheetInternal(Optional ByVal Index As Long = -1, Optional ByVal Create As Boolean = True)
    Dim SheetName As String
    If Index = -1 Then
        SheetName = TEMP_SHEET
        CanUndo = False
    Else
        SheetName = TEMP_SHEET & Index
    End If
    On Error GoTo catch
    Set TempSheetInternal = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0
    If Not ClearAfter Then TempSheetInternal.UsedRange.Delete
    Exit Property
catch:
    If Create Then
        Set TempSheetInternal = ThisWorkbook.Worksheets.Add
        TempSheetInternal.Visible = xlSheetVeryHidden
        TempSheetInternal.Name = SheetName
    End If
End Property

Private Function TryCreateSheet(ByVal SheetName As String) As Worksheet
    On Error GoTo catch
    Set TryCreateSheet = ThisWorkbook.Worksheets(SheetName)
    Set TryCreateSheet = Nothing
    Exit Function
catch:
    Set TryCreateSheet = Worksheets.Add
    TryCreateSheet.Name = SheetName
End Function

'
' Clears a temporary sheet. If the Unique flag has been set, the sheet is deleted. Otherwise, the used range is deleted.
'
Sub ClearTemp(Optional ByRef Sheet As Worksheet)
    If ClearAfter Then
        Dim tmp As Worksheet
        Set tmp = TempSheetInternal(Create:=False)
        If Not (tmp Is Nothing) Then tmp.UsedRange.Delete
    End If
End Sub

'
' Performs an index match on a range and replaces that range with those values
'
' DestValues - the destination keys that will be replaced
' SourceLookup - the source keys
' SourceValues - the source values that will replace the destination
'
#If POWERSHEET_VER < 1016 Then
Sub TransformRange(ByRef DestValues As Variant, ByRef SourceLookup As Variant, ByRef SourceValues As Variant, Optional ByVal DefaultValue As Variant, Optional ByVal Clipboard As Boolean = False)
#Else
Sub TransformRange(ByRef DestValues As Variant, ByRef SourceLookup As Variant, ByRef SourceValues As Variant, Optional ByVal DefaultValue As Variant, Optional ByVal Clipboard As Boolean = True)
#End If
    Dim rSrcLookup As Range, rDestValues As Range, rSrcValues As Range
    Dim tmp As Worksheet, rColumnB As Range
    Dim copied As Long, Formula As String

    Set tmp = TempSheetInternal
    
    Set rDestValues = SheetEx.GetRange(DestValues)
    Set rSrcLookup = SheetEx.GetRange(SourceLookup)
    Set rSrcValues = SheetEx.GetRange(SourceValues)
    
    copied = CopyOption(Array(rDestValues, tmp.Columns("A")), Clipboard)
    Set rColumnB = tmp.Columns("B").Resize(RowSize:=copied)
    
    Formula = SheetEx.CreateLookup(rSrcLookup, rSrcValues, tmp.Columns("A"))
    
    If IsMissing(DefaultValue) Then
#If POWERSHEET_VER > 1010 Then
        Formula = "=" & Formula
#Else
        Formula = "=IFNA(" & Formula & ", """")"
#End If
    Else
        If TypeName(DefaultValue) = "String" Then DefaultValue = """" & Replace(DefaultValue, """", """""") & """"
        Formula = "=IFNA(" & Formula & ", " & DefaultValue & ")"
    End If
    
    SheetEx.FillValues rColumnB, Formula, Clipboard:=Clipboard
    CopyOption Array(rColumnB, rDestValues), Clipboard
    
    ClearTemp tmp
End Sub

'
' Transforms a range into a set of numbers based on their values. Each value is assigned a unique number.
'
' Range - the range to transform into numbers
' StartCount - the number to start counting new items
'
Sub NumberValues(ByRef Range As Variant, Optional ByVal StartCount As Double = 1)
    Dim rng As Range, tmp As Worksheet
    Dim A As Range, B As Range, C As Range, D As Range
    Dim copied As Long
    
    Set rng = SheetEx.GetRange(Range)
    Set tmp = TempSheetInternal
    
    copied = SheetEx.FastCopy(rng, tmp.Columns("A"))
    Set A = tmp.Columns("A").Resize(RowSize:=copied)
    
    A.RemoveDuplicates 1
    Set A = ShrinkToFit(A, exRow)
    
    copied = SheetEx.FastCopy(rng, tmp.Columns("C"))
    Set B = tmp.Columns("B").Resize(RowSize:=A.Rows.Count)
    With B.Cells(1, 1)
        .Value2 = StartCount
        .AutoFill B, xlFillSeries
    End With
    
    SheetEx.SortTable tmp.Range("A:B"), 1, Order:=xlAscending, Header:=xlNo
    
    Set C = B.Resize(RowSize:=copied).Offset(ColumnOffset:=1)
    Set D = C.Resize(RowSize:=copied).Offset(ColumnOffset:=1)
    
    SheetEx.JoinReplace A, B, C, D
    SheetEx.FastCopy D, rng
    
    ClearTemp tmp
End Sub

'
' Removes duplicates in a range without taking into account the entire column
'
' Range - the range in which to remove duplicates
' Columns - a single column number or an array of column numbers to include in the remove duplicate range
' Header - whether or not this data set contains a header
' DeleteEmpty - delete any blank cells at the bottom of the range when duplicates are removed
'
Sub RRemoveDuplicates(ByRef Range As Variant, Optional ByRef Columns As Variant, Optional ByVal Header As XlYesNoGuess = xlNo, Optional ByVal DeleteEmpty As Boolean = True)
    Dim rng As Range, tmp As Worksheet
    Dim copied As Long
    
    Set rng = SheetEx.GetRange(Range)
    Set tmp = TempSheetInternal
    
    copied = SheetEx.CopySingle(rng, tmp, Paste:=xlPasteAll)
    If IsArray(Columns) Then
        SheetEx.GetRange(tmp).RemoveDuplicates (Columns), Header
    Else
        SheetEx.GetRange(tmp).RemoveDuplicates Columns, Header
    End If
    rng.Clear
    copied = SheetEx.CopySingle(tmp, rng, Paste:=xlPasteAll)
    
    If DeleteEmpty And copied < rng.Rows.Count Then
        Dim rEmpties As Range
        Set rEmpties = rng.Offset(RowOffset:=copied).Resize(RowSize:=Math.Abs(rng.Rows.Count - copied))
        rEmpties.Delete xlShiftUp
    End If
    
    ClearTemp tmp
    
End Sub

'
' Copies only the unique values from a range to a destination
'
' Source - the source range to copy
' Destination - the destination range
' Columns - the column or column array with unique values
' Paste - the paste type
' CellType - copy only special cells
' Append - append instead of overwrite the destination
' StackSources - if an array of ranges is specified as the source, indicates whether the sources should be
'                added to the bottom of the range rather than to the right of it
'
Function CopyUnique(ByVal Source As Variant, ByRef Destination As Variant, Optional ByRef Columns As Variant, Optional ByVal Header As XlYesNoGuess = xlNo, Optional ByVal Paste As XlPasteType = xlPasteValues, Optional ByVal CellType As XlCellType, Optional ByVal Append As Boolean = False, Optional ByVal StackSources As Boolean = False)
    Dim rSrc As Range, rDest As Range
    Dim tmp As Worksheet, copied, ColCount As Long
    Dim i As Long
    
    If Not IsArray(Source) Or TypeName(Source) = "Range" Then
        Source = Array(Source)
    End If
    
    Set tmp = TempSheetInternal
    Set rDest = SheetEx.GetRange(Destination)
    
    For i = LBound(Source) To UBound(Source)
        Dim tmpDest As Range
        
        Set rSrc = GetRange(Source(i))
        
        If tmpDest Is Nothing Then
            Set tmpDest = tmp.Columns("A").Resize(RowSize:=rSrc.Rows.Count)
            ColCount = CountColumns(rSrc)
        ElseIf Not StackSources Then
            Set tmpDest = tmpDest.Offset(ColumnOffset:=rSrc.Columns.Count)
            If Header = xlYes Then Set rSrc = SheetEx.RShrink(rSrc, Height:=-1)
            ColCount = ColCount + CountColumns(rSrc)
        End If
        
        copied = SheetEx.CopySingle(rSrc, tmpDest, CellType, Paste, Append:=StackSources And i <> LBound(Source))
    Next i
        
    If IsMissing(Columns) Then
        Dim NewCols() As Variant
        ReDim NewCols(ColCount - 1)
        For i = LBound(NewCols) To UBound(NewCols)
            NewCols(i) = i + 1
        Next i
        Columns = NewCols
    End If
        
    If IsArray(Columns) Then
        SheetEx.GetRange(tmp).RemoveDuplicates Columns:=(Columns), Header:=Header
    Else
        SheetEx.GetRange(tmp).RemoveDuplicates Columns, Header
    End If
    
    CopyUnique = CopySingle(tmp, rDest, Paste:=Paste, Append:=Append)
    
    ClearTemp tmp
    
End Function

'
' Join a column on a key without replacing any value that has already been filled
'
' SourceKeys   - the range that contains keys on the source table
' SourceValues - the range of values to copy to the destination
' DestKeys     - the range that contains keys on the destination table
' DestValues   - the range to place the copied values, only blank cells in this range will be overwritten
' ReplaceZero  - whether or not the value of zero should be replaced with a new value
'
#If POWERSHEET_VER < 1016 Then
Sub JoinPreserve(ByRef SourceKeys As Variant, ByRef SourceValues As Variant, ByRef DestKeys As Variant, ByRef DestValues As Variant, Optional ByVal DefaultValue As Variant, Optional ByVal Calculate As Boolean = True, Optional ByVal ReplaceZero As Boolean = True, Optional ByVal Clipboard As Boolean = False)
#Else
Sub JoinPreserve(ByRef SourceKeys As Variant, ByRef SourceValues As Variant, ByRef DestKeys As Variant, ByRef DestValues As Variant, Optional ByVal DefaultValue As Variant, Optional ByVal Calculate As Boolean = True, Optional ByVal ReplaceZero As Boolean = True, Optional ByVal Clipboard As Boolean = True)
#End If
    Dim rDestKeys As Range, rSrcKeys As Range, rDestValues As Range, rSrcValues As Range
    Dim tmp As Worksheet, copied As Long, Formula As String, ArrayFormula As Boolean
    Dim Offset As Long
    Set tmp = TempSheetInternal

    Set rDestKeys = SheetEx.CombineRange(DestKeys)
    Set rSrcKeys = SheetEx.CombineRange(SourceKeys)
    Set rDestValues = ROffset(DestValues, Height:=MaxRow(rDestKeys), Width:=1)
    Set rSrcValues = ROffset(SourceValues, Height:=MaxRow(rSrcKeys), Width:=1)
    
    If IsArray(DestKeys) And TypeName(DestKeys) <> "Range" Then
        Dim n As Long
        For n = LBound(DestKeys) To UBound(DestKeys)
            copied = WorksheetFunction.Max(CopyOption(Array(DestKeys(n), tmp.UsedRange.Offset(ColumnOffset:=Offset)), Clipboard), copied)
            Offset = Offset + CountColumns(DestKeys(n))
        Next n
    Else
        copied = CopyOption(Array(DestKeys, tmp), Clipboard)
        Offset = CountColumns(DestKeys)
    End If
    copied = WorksheetFunction.Max(CopyOption(Array(rDestValues, ROffset(tmp, ColumnOffset:=tmp.UsedRange.Columns.Count, Width:=1)), Clipboard), copied)
    
    With ROffset(tmp, Height:=copied)
        Dim MatchCol As Range, TestCol As Range, ValueCol As Range
        Dim AddrValue As String, AddrMatch As String
        
        Set ValueCol = .Offset(ColumnOffset:=Offset).Resize(ColumnSize:=1)
        Set MatchCol = ValueCol.Offset(ColumnOffset:=1)
        Set TestCol = MatchCol.Offset(ColumnOffset:=1)
        
        Formula = SheetEx.CreateLookup(SourceKeys, rSrcValues, SheetEx.RangeToArray(.Range(.Cells(1, 1), .Cells(.Rows.Count, WorksheetFunction.Max(.Columns.Count - 1, 1)))))
        ArrayFormula = IsArray(SourceKeys) And TypeName(SourceKeys) <> "Range" And IsArray(DestKeys) And TypeName(DestKeys) <> "Range"
        
        If IsMissing(DefaultValue) Then
#If POWERSHEET_VER > 1010 Then
            Formula = "=" & Formula
#Else
            Formula = "=IFNA(" & Formula & ", """")"
#End If
        Else
            If TypeName(DefaultValue) = "String" Then DefaultValue = """" & Replace(DefaultValue, """", """""") & """"
            Formula = "=IFNA(" & Formula & ", " & DefaultValue & ")"
        End If

        AddrValue = ValueCol.Cells(1, 1).Address(RowAbsolute:=False)
        AddrMatch = MatchCol.Cells(1, 1).Address(RowAbsolute:=False)
        
        SheetEx.FillValues MatchCol, Formula, Calculate:=Calculate, Clipboard:=Clipboard, ArrayFormula:=ArrayFormula
        If ReplaceZero Then
            SheetEx.FillValues TestCol, "=IF(OR(" & AddrValue & "="""", " & AddrValue & "=0), " & AddrMatch & ", " & AddrValue & ")", Calculate:=Calculate, Clipboard:=Clipboard
        Else
            SheetEx.FillValues TestCol, "=IF(" & AddrValue & "="""", " & AddrMatch & ", " & AddrValue & ")", Calculate:=Calculate, Clipboard:=Clipboard
        End If
        
        CopyOption Array(TestCol, rDestValues), Clipboard
    End With
    
    ClearTemp tmp
End Sub

'
' Applies a formula to a given range based on an excel formula
'
' Ranges - a range or array of ranges that will be used as an argument for this formula
' Formula - the formula that should be applied to the array. Arguments can be used in the formula string
'           in the format %1 is arg1, %2 is arg2, etc.
' Destination - where to place the final values. if this argument is omitted, the the first range in the Ranges array is used
' FillRight - whether or not the columns should be expanded and filled to the right
'
#If POWERSHEET_VER >= 1016 Then
Sub ApplyFormula(ByVal Ranges As Variant, ByVal Formula As String, Optional ByRef Destination As Variant, Optional ByVal Clipboard As Boolean = True, Optional ByVal FillRight As Boolean = False)
#ElseIf POWERSHEET_VER > 1000 Then
Sub ApplyFormula(ByVal Ranges As Variant, ByVal Formula As String, Optional ByRef Destination As Variant, Optional ByVal Clipboard As Boolean = False, Optional ByVal FillRight As Boolean = False)
#Else
Sub ApplyFormula(ByVal Ranges As Variant, ByRef Destination As Variant, ByVal Formula As String, Optional ByVal Clipboard As Boolean = False)
#End If
    Dim tmp As Worksheet, Last As Long
    Dim copied As Long, i As Long
    Dim ccount As Long
    
    If Not IsArray(Ranges) Or TypeName(Ranges) = "Range" Then Ranges = Array(Ranges)
    
    If IsMissing(Destination) Then Set Destination = Ranges(LBound(Ranges))
    
    Set tmp = TempSheetInternal
    Let Last = 1
    
    For i = LBound(Ranges) To UBound(Ranges)
        Dim arg As String, argRange As Range, ColCount As Long
        copied = WorksheetFunction.Max(CopyOption(Array(Ranges(i), tmp.Columns(Last)), Clipboard), copied)
        ColCount = CountColumns(Ranges(i))
        With GetRange(Ranges(i))
            Set argRange = tmp.Columns(Last).Resize(RowSize:=1, ColumnSize:=ColCount)
            If ColCount = 1 Then Set argRange = argRange.Range("A1")
            arg = argRange.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            Formula = Replace(Formula, "%" & (i + 1), arg)
            Last = Last + ColCount
            ccount = WorksheetFunction.Max(ColCount, ccount)
        End With
    Next i
    
    With ROffset(tmp, Height:=copied)
        Dim Final As Range
        Set Final = .Columns(tmp.UsedRange.Column + tmp.UsedRange.Columns.Count)
#If POWERSHEET_VER > 1000 Then
        If FillRight Then Set Final = .Columns(tmp.UsedRange.Column + tmp.UsedRange.Columns.Count).Resize(ColumnSize:=.Columns.Count)
#End If
        FillValues Final, Formula, Clipboard:=Clipboard
        CopyOption Array(Final, Destination), Clipboard
    End With
        
    ClearTemp tmp
End Sub

'
' Replaces entire cells whose contents match the Match argument with the Replace argument
'
#If POWERSHEET_VER < 1016 Then
Sub ReplaceValue(ByRef Range As Variant, ByVal Match As Variant, ByVal Replace As Variant, Optional ByVal Clipboard As Boolean = False)
#Else
Sub ReplaceValue(ByRef Range As Variant, ByVal Match As Variant, ByVal Replace As Variant, Optional ByVal Clipboard As Boolean = True)
#End If
    Dim rng As Range, tmp As Worksheet
    
    If TypeName(Match) = "String" Then Match = """" & Match & """"
    If TypeName(Replace) = "String" Then Replace = """" & Replace & """"
    
    ApplyFormula Ranges:=Range, Formula:="=IF(%1=" & Match & ", " & Replace & ", %1)", Destination:=Range, Clipboard:=Clipboard
End Sub

'
' Adds two or more columns together and pastes the result in a destination
'
' Columns - the columns that need to be added
' Destination - the range where the values will be pasted
' SkipBlank - whether or not to add blank columns together. If blanks are skipped, the result will be 0 instead of an empty cell.
'
#If POWERSHEET_VER < 1016 Then
Sub SumColumns(ByVal Columns As Variant, ByRef Destination As Variant, Optional ByVal SkipBlank As Boolean = True, Optional ByVal Clipboard As Boolean = False)
#Else
Sub SumColumns(ByVal Columns As Variant, ByRef Destination As Variant, Optional ByVal SkipBlank As Boolean = True, Optional ByVal Clipboard As Boolean = True)
#End If
    Dim varName As String, szFormula As String
    
    If Not IsArray(Columns) Or TypeName(Columns) = "Range" Then Columns = Array(Columns)
    varName = "%" & (UBound(Columns) + 1)
    szFormula = "SUM(%1:" & varName & ")"
    If SkipBlank Then szFormula = "IF(COUNTA(%1:" & varName & ")<>" & (UBound(Columns) + 1) & ", """", " & szFormula & ")"
    
    ApplyFormula Ranges:=Columns, Clipboard:=Clipboard, Destination:=Destination, Formula:="=" & szFormula
End Sub

'
' Adds a column to another column based on its key
'
' SourceKeys   - the range that contains keys on the source table
' SourceValues - the range of values to copy to the destination
' DestKeys     - the range that contains keys on the destination table
' DestValues   - the range to place the copied values, only blank cells in this range will be overwritten
'
#If POWERSHEET_VER < 1016 Then
Sub AddToColumn(ByRef SourceKeys As Variant, ByRef SourceValues As Variant, ByRef DestKeys As Variant, ByRef DestValues As Variant, Optional ByVal Clipboard As Boolean = False)
#Else
Sub AddToColumn(ByRef SourceKeys As Variant, ByRef SourceValues As Variant, ByRef DestKeys As Variant, ByRef DestValues As Variant, Optional ByVal Clipboard As Boolean = True)
#End If
    Dim rSrcKeys As Range, rSrcValues As Range
    Dim rDestKeys As Range, rDestValues As Range
    Dim tmp As Worksheet, copied As Long
    
    Set tmp = TempSheetInternal
    
    copied = CopyOption(Array(DestKeys, tmp.Columns("A"), DestValues, tmp.Columns("B")), Clipboard)
    
    Set rSrcKeys = GetRange(SourceKeys)
    Set rSrcValues = GetRange(SourceValues)
    
    With ROffset(tmp, Height:=copied)
        JoinReplace SourceKeys, SourceValues, .Columns("A"), .Columns("C"), DefaultValue:=0, Clipboard:=Clipboard
        FillValues .Columns("D"), "=$B1+$C1", Clipboard:=Clipboard
        CopyOption Array(.Columns("D"), DestValues), Clipboard
    End With
    
    ClearTemp tmp
End Sub

'
' Copies the values that are in Column1 that are not in Column2
' This function will use TempSheet -2, creating it if it doesn't exist.
'
' Column1 - a column of values
' Column2 - the second column with values
' Destination - where to place the distinct values
' Paste - the paste type
' CellType - copy only special cells
' Append - append instead of overwrite the destination
'
Function CopyDistinct(ByRef Column1 As Variant, ByRef Column2 As Variant, ByRef Destination As Variant, Optional ByVal Paste As XlPasteType = xlPasteValues, Optional ByVal CellType As XlCellType, Optional ByVal Append As Boolean = False) As Long
    Dim tmp As Worksheet, rDest As Range
    Dim copied As Long
    
    With TempSheetInternal(-2)
        CopyUnique Column1, .Columns("A"), CellType:=CellType, Paste:=Paste
        CopyUnique Column2, .Columns("B"), CellType:=CellType, Paste:=Paste
        
        Set tmp = TempSheetInternal
        copied = CopySingle(.Range(.Cells(1, 1), SheetEx.LastCell(.UsedRange)), tmp.Range("A2"))
    End With
    
    With ROffset(tmp, Height:=copied + 1)
        Dim nas As Range
        .Range("A1:C1").Value2 = "HEADER"
        .Range("A:C").AutoFilter
        FillValues RShrink(.Columns("C"), Height:=-1), "=MATCH($A2, " & RShrink(.Columns("B"), Height:=-1).Address & ", 0)"
        FilterTable .Range("A:C"), "C", "<>#N/A"
        
        Set nas = SafeSpecial(RShrink(.Range("A:C"), Height:=-1), xlCellTypeVisible)
        If Not (nas Is Nothing) Then nas.Delete xlUp
        
        ClearFilters .Range("A:C")
        
        CopyDistinct = FindLastRow(.Columns("A")) - 1
        If CopyDistinct > 0 Then CopySingle RShrink(.Columns("A"), Height:=-1), Destination, Append:=Append
    End With
    
    ClearTemp tmp
End Function

Private Function CopyOption(ByRef Ranges As Variant, ByVal Clipboard As Boolean, Optional ByVal StackSources As Boolean = True)
    If Clipboard Then
        CopyOption = CopyMult(Ranges)
    Else
        CopyOption = FastCopyEx(Ranges, StackSources:=StackSources)
    End If
End Function

'
' Formats a range and saves the values.
' Using the clipboard will ensure that the values are in the correct format, but it could be slower.
'
Sub FormatColumn(ByRef ColumnRange As Variant, ByVal Format As String, Optional ByVal Clipboard As Boolean = True, Optional ByVal FormatBlanks As Boolean = True)
    With GetRange(ColumnRange)
        Dim szFormula As String
        szFormula = "TEXT(%1, """ & Replace(Format, """", """""") & """)"
        
        If Not FormatBlanks Then szFormula = "IF(%1="""", """", " & szFormula & ")"
    
        ApplyFormula Ranges:=.Columns(1), Destination:=.Columns(1), Formula:="=" & szFormula, Clipboard:=Clipboard
    End With
End Sub
