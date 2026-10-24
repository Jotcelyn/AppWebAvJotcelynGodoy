<%
' ============================================================
'  PoliMusic - Render de una lista de canciones reutilizable
' ============================================================
Sub RenderTracklist(rs, listId)
    Dim idx, sid, sname, aname, albname, coverUrl, audioSrc, searchBlob
    idx = 0
    Response.Write "<div class='tracklist' id='" & H(listId) & "'>"

    Do While Not rs.EOF
        sid = rs("ID_SONG")
        sname = rs("SONG_NAME") & ""
        aname = rs("ARTISTS") & ""
        albname = rs("ALBUMS") & ""
        If aname = "" Then aname = "Artista desconocido"
        If albname = "" Then albname = "Sencillo"
        coverUrl = ImgUrl(rs("COVER") & "", PageDepth & "assets/img/default.jpg")
        audioSrc = AudioUrl(rs("SONG_PATH") & "")
        searchBlob = sname & " " & aname

        Response.Write "<div class='track' data-song-id='" & sid & "'" & _
            " data-title='" & H(sname) & "'" & _
            " data-artist='" & H(aname) & "'" & _
            " data-src='" & H(audioSrc) & "'" & _
            " data-cover='" & H(coverUrl) & "'" & _
            " data-search-item='" & H(searchBlob) & "'>"
        Response.Write "<span class='track__index'>" & (idx + 1) & "</span>"
        Response.Write "<div class='track__cover'><img src='" & H(coverUrl) & "' alt='Portada de " & H(sname) & "' loading='lazy'></div>"
        Response.Write "<div class='track__meta'><strong>" & H(sname) & "</strong><span>" & H(aname) & " · " & H(albname) & "</span></div>"
        Response.Write "<button type='button' class='track__play' data-track-index='" & idx & "' aria-label='Reproducir " & H(sname) & "'>" & Icon("play") & "</button>"
        Response.Write "</div>"

        idx = idx + 1
        rs.MoveNext
    Loop

    Response.Write "</div>"
End Sub
%>
