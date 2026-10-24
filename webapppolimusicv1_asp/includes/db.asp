<!--#include file="config.asp"-->
<%
' ============================================================
'  PoliMusic - Capa de acceso a datos (ADO / SQL Server)
' ============================================================
Response.Buffer = True
Response.CodePage = 65001
Response.Charset = "utf-8"
Session.CodePage = 65001

Dim gConn
Set gConn = Server.CreateObject("ADODB.Connection")

Sub OpenDbConnection()
    On Error Resume Next
    gConn.Open "Provider=MSOLEDBSQL19;Server=" & DB_SERVER & ";Database=" & DB_NAME & _
               ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";Encrypt=Optional;TrustServerCertificate=Yes;"
    If Err.Number <> 0 Then
        Err.Clear
        gConn.Open "Provider=MSOLEDBSQL;Server=" & DB_SERVER & ";Database=" & DB_NAME & _
                   ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";TrustServerCertificate=Yes;"
    End If
    If Err.Number <> 0 Then
        Err.Clear
        gConn.Open "Provider=SQLOLEDB;Server=" & DB_SERVER & ";Database=" & DB_NAME & _
                   ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";"
    End If
    If Err.Number <> 0 Then
        Response.Write "<!DOCTYPE html><html lang=""es""><head><meta charset=""utf-8""><title>PoliMusic</title></head><body style=""background:#0E0B16;color:#F2EFFA;font-family:Segoe UI,sans-serif;padding:48px"">"
        Response.Write "<h1>No se pudo conectar con la base de datos</h1>"
        Response.Write "<p style=""color:#B9A3F5"">Verifica que SQL Server este activo y que la base <strong>" & DB_NAME & "</strong> exista.</p>"
        Response.Write "<pre style=""color:#9C93B8;white-space:pre-wrap"">" & Server.HTMLEncode(Err.Description) & "</pre></body></html>"
        Response.End
    End If
    On Error GoTo 0
End Sub

OpenDbConnection

' ---------- Constructores de comandos parametrizados ----------
Function NewCmd(sqlText)
    Dim c
    Set c = Server.CreateObject("ADODB.Command")
    Set c.ActiveConnection = gConn
    c.CommandType = 1
    c.CommandText = sqlText
    Set NewCmd = c
End Function

Sub AddStr(c, pName, pValue)
    c.Parameters.Append c.CreateParameter(pName, 202, 1, 4000, CStr(pValue & ""))
End Sub

Sub AddInt(c, pName, pValue)
    c.Parameters.Append c.CreateParameter(pName, 3, 1, , CLng(pValue))
End Sub

Sub AddDate(c, pName, pValue)
    c.Parameters.Append c.CreateParameter(pName, 135, 1, , CDate(pValue))
End Sub

Sub AddIntOrNull(c, pName, pValue)
    If Trim(pValue & "") = "" Then
        c.Parameters.Append c.CreateParameter(pName, 3, 1, , Null)
    Else
        c.Parameters.Append c.CreateParameter(pName, 3, 1, , CLng(pValue))
    End If
End Sub

Sub AddDateOrNull(c, pName, pValue)
    If Trim(pValue & "") = "" Then
        c.Parameters.Append c.CreateParameter(pName, 135, 1, , Null)
    Else
        c.Parameters.Append c.CreateParameter(pName, 135, 1, , CDate(pValue))
    End If
End Sub

' ---------- Utilidades ----------
' Devuelve un recordset estático del lado del cliente: la conexión queda
' libre de inmediato y se pueden tener varios recordsets abiertos a la vez.
Function ExecRs(cmd)
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open cmd, , 3, 1
    Set ExecRs = rs
End Function

Function SqlRs(sqlText)
    Dim cmd
    Set cmd = NewCmd(sqlText)
    Set SqlRs = ExecRs(cmd)
End Function

Function ScalarVal(sqlText)
    Dim rs
    Set rs = gConn.Execute(sqlText)
    If Not rs.EOF Then
        ScalarVal = rs(0).Value
    Else
        ScalarVal = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function H(value)
    Dim s
    s = value & ""
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    H = s
End Function

Function JsStr(value)
    Dim s
    s = value & ""
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    JsStr = s
End Function

Function FileNameOf(pathValue)
    Dim p, pos
    p = Replace(pathValue & "", "\", "/")
    pos = InStrRev(p, "/")
    If pos > 0 Then
        FileNameOf = Mid(p, pos + 1)
    Else
        FileNameOf = p
    End If
End Function

Function AssetExists(relPath)
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    AssetExists = fso.FileExists(Server.MapPath(relPath))
    Set fso = Nothing
End Function

' Resuelve una ruta de imagen almacenada en la BD (~/playlistImages/x.png)
' hacia un archivo real dentro de /assets/img. Devuelve "" si no existe.
Function ImgUrl(dbPath, fallback)
    Dim file, candidate
    file = FileNameOf(dbPath)
    If file = "" Then
        ImgUrl = fallback
        Exit Function
    End If
    candidate = PageDepth & "assets/img/" & file
    If AssetExists(candidate) Then
        ImgUrl = candidate
    Else
        ImgUrl = fallback
    End If
End Function

Function AudioUrl(dbPath)
    Dim file, candidate
    file = FileNameOf(dbPath)
    If file = "" Then file = "bensound-adventure.mp3"
    candidate = PageDepth & "assets/audio/" & file
    If AssetExists(candidate) Then
        AudioUrl = candidate
    Else
        AudioUrl = PageDepth & "assets/audio/bensound-adventure.mp3"
    End If
End Function

Function IsoDate(value)
    If IsNull(value) Or value & "" = "" Then
        IsoDate = ""
    Else
        IsoDate = Year(value) & "-" & Right("0" & Month(value), 2) & "-" & Right("0" & Day(value), 2)
    End If
End Function

Function PrettyDate(value)
    If IsNull(value) Or value & "" = "" Then
        PrettyDate = "Sin fecha"
    Else
        PrettyDate = Right("0" & Day(value), 2) & "/" & Right("0" & Month(value), 2) & "/" & Year(value)
    End If
End Function

Function Initials(value)
    Dim parts, i, out
    parts = Split(Trim(value & ""), " ")
    out = ""
    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 And Len(out) < 2 Then
            out = out & UCase(Left(parts(i), 1))
        End If
    Next
    If out = "" Then out = "PM"
    Initials = out
End Function

Function Pluralize(count, singular, plural)
    If CLng(count) = 1 Then
        Pluralize = count & " " & singular
    Else
        Pluralize = count & " " & plural
    End If
End Function

' Cierra la conexion al final de cada pagina.
Sub CloseDbConnection()
    On Error Resume Next
    If IsObject(gConn) Then
        If gConn.State = 1 Then gConn.Close
    End If
    Set gConn = Nothing
    On Error GoTo 0
End Sub
%>
