<%
' ============================================================
'  PoliMusic - Consultas reutilizables de canciones
' ============================================================
Function SongSelectBase()
    SongSelectBase = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
        "(SELECT TOP 1 al.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST t JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
        " WHERE t.ID_SONG = s.ID_SONG ORDER BY al.ID_ALBUM) AS COVER, " & _
        "(SELECT STRING_AGG(a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME, ''), ', ') FROM TBL_PLAYER p " & _
        " JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) AS ARTISTS, " & _
        "(SELECT STRING_AGG(al.TITLE_ALBUM, ', ') FROM TBL_TRACKLIST t JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
        " WHERE t.ID_SONG = s.ID_SONG) AS ALBUMS "
End Function

Function SongsByPlaylist(playlistId)
    Dim c
    Set c = NewCmd(SongSelectBase() & _
        "FROM TBL_PLAYLIST_SONG ps JOIN TBL_SONG s ON s.ID_SONG = ps.ID_SONG " & _
        "WHERE ps.ID_PLAYLIST = ? ORDER BY s.ID_SONG")
    AddInt c, "@playlist", playlistId
    Set SongsByPlaylist = ExecRs(c)
End Function

Function SongsByArtist(artistId)
    Dim c
    Set c = NewCmd(SongSelectBase() & _
        "FROM TBL_PLAYER p JOIN TBL_SONG s ON s.ID_SONG = p.ID_SONG " & _
        "WHERE p.ID_ARTIST = ? ORDER BY s.ID_SONG")
    AddInt c, "@artist", artistId
    Set SongsByArtist = ExecRs(c)
End Function

Function SongsByAlbum(albumId)
    Dim c
    Set c = NewCmd(SongSelectBase() & _
        "FROM TBL_TRACKLIST t JOIN TBL_SONG s ON s.ID_SONG = t.ID_SONG " & _
        "WHERE t.ID_ALBUM = ? ORDER BY s.ID_SONG")
    AddInt c, "@album", albumId
    Set SongsByAlbum = ExecRs(c)
End Function
%>
