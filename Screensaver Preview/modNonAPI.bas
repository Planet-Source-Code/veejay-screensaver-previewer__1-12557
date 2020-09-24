Attribute VB_Name = "modNonAPI"
Option Explicit

Public Function FixPathWithFile(inPath As String, inFile As String) As String
    Dim tPath As String
    If Right(inPath, 1) <> "\" Then
        tPath = inPath & "\" & inFile
    Else: tPath = inPath & inFile
    End If
    FixPathWithFile = tPath
End Function

Public Function GetFileName(inPath As String, inIncludeExtension As Boolean) As String
    Dim tname As String, i As Integer
    If InStr(1, inPath, "\") < 1 Then
        GetFileName = inPath
        Exit Function
    End If
    For i = Len(inPath) To 1 Step -1
        If Mid$(inPath, i, 1) = "\" Then
            GetFileName = Right$(inPath, Len(inPath) - i)
            If inIncludeExtension = False Then
                GetFileName = Left(GetFileName, Len(GetFileName) - 4)
            End If
            Exit Function
        End If
    Next i
End Function

Public Function StripNulls(inOriginalStr As String) As String
If (InStr(inOriginalStr, Chr(0)) > 0) Then
    inOriginalStr = Left(inOriginalStr, InStr(inOriginalStr, Chr(0)) - 1)
End If
StripNulls = inOriginalStr
End Function

Public Sub SelectionSort(List() As String, ByVal min As Integer, ByVal max As Integer)
Dim i As Integer
Dim j As Integer
Dim best_j As Integer
Dim best_str As String
Dim temp_str As String

    For i = min To max - 1
        best_j = i
        best_str = List(i)
        For j = i + 1 To max
            If StrComp(List(j), best_str, vbTextCompare) < 0 Then
                best_str = List(j)
                best_j = j
            End If
        Next j
        List(best_j) = List(i)
        List(i) = best_str
    Next i
End Sub

