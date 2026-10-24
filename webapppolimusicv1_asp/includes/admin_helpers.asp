<%
' ============================================================
'  PoliMusic - Utilidades del panel administrativo
' ============================================================
Function CleanIntList(csv)
    Dim parts, i, v, out
    out = ""
    If csv & "" <> "" Then
        parts = Split(csv & "", ",")
        For i = 0 To UBound(parts)
            v = Trim(parts(i))
            If v <> "" And IsNumeric(v) Then
                If out <> "" Then out = out & ","
                out = out & CLng(v)
            End If
        Next
    End If
    CleanIntList = out
End Function

Function NumOrZero(v)
    If Trim(v & "") <> "" And IsNumeric(v) Then
        NumOrZero = CLng(v)
    Else
        NumOrZero = 0
    End If
End Function

Function TrimText(v, maxLen)
    Dim s
    s = Trim(v & "")
    If Len(s) > maxLen Then s = Left(s, maxLen)
    TrimText = s
End Function

Function SelectedAttr(selectedCsv, value)
    If InStr("," & selectedCsv & ",", "," & value & ",") > 0 Then
        SelectedAttr = " checked"
    Else
        SelectedAttr = ""
    End If
End Function

' Convierte yyyy-mm-dd en Date. Devuelve Null si no es válida.
Function ParseIsoDate(value)
    Dim p
    ParseIsoDate = Null
    If Trim(value & "") = "" Then Exit Function
    p = Split(Trim(value & ""), "-")
    If UBound(p) = 2 Then
        If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
            On Error Resume Next
            ParseIsoDate = DateSerial(CInt(p(0)), CInt(p(1)), CInt(p(2)))
            If Err.Number <> 0 Then
                ParseIsoDate = Null
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
End Function
%>
