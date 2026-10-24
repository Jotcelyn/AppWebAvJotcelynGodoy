<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/queries.asp"-->
<!--#include file="includes/tracklist.asp"-->
<%
RequireLogin
Dim PageTitle, ActiveNav
PageTitle = "Álbumes"
ActiveNav = "albums"

Dim alRs, alCount
Set alRs = SqlRs("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
    "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS SONG_COUNT " & _
    "FROM TBL_ALBUM al ORDER BY al.RELEASE_YEAR DESC, al.TITLE_ALBUM")
alCount = ScalarVal("SELECT COUNT(*) FROM TBL_ALBUM")
%>
<!--#include file="includes/layout_top.asp"-->

<div class="page-head">
    <div>
        <span class="page-head__eyebrow"><%=Icon("disc")%> Biblioteca</span>
        <h1>Álbumes</h1>
        <p>Colecciones completas organizadas por año de lanzamiento. Abre un álbum para recorrer su tracklist y reproducirlo.</p>
    </div>
    <div class="page-head__actions">
        <% If IsAdmin() Then %>
        <a class="btn btn--primary" href="admin-albums.asp?new=1"><%=Icon("plus")%> Nuevo álbum</a>
        <% End If %>
        <span class="chip chip--muted"><%=Pluralize(alCount, "álbum", "álbumes")%></span>
    </div>
</div>

<% If alCount = 0 Then %>
<div class="empty">
    <div class="empty__ico"><%=Icon("disc")%></div>
    <strong>Todavía no hay álbumes registrados</strong>
    <p>El administrador puede crear álbumes desde el panel de administración.</p>
</div>
<% Else %>
<div class="cards">
    <%
    Dim alId, alTitle, alYear, alPhoto, alSongs
    Do While Not alRs.EOF
        alId = alRs("ID_ALBUM")
        alTitle = alRs("TITLE_ALBUM") & ""
        alYear = alRs("RELEASE_YEAR") & ""
        alPhoto = ImgUrl(alRs("ALBUM_IMAGE_PATH") & "", PageDepth & "assets/img/default.jpg")
        alSongs = alRs("SONG_COUNT")
        If alYear = "" Then alYear = "s/f"
    %>
    <button type="button" class="card" data-open-modal="album-<%=alId%>"
            data-search-item="<%=H(alTitle & " " & alYear)%>">
        <div class="card__art">
            <img src="<%=alPhoto%>" alt="Portada de <%=H(alTitle)%>" loading="lazy">
            <span class="card__badge"><%=H(alYear)%></span>
            <span class="card__play" aria-hidden="true"><%=Icon("play")%></span>
        </div>
        <div class="card__title"><%=H(alTitle)%></div>
        <div class="card__meta"><%=Pluralize(alSongs, "canción", "canciones")%></div>
    </button>
    <%
        alRs.MoveNext
    Loop
    %>
</div>
<p class="no-results" data-no-results>No se encontraron álbumes con ese criterio.</p>
<% End If %>
<%
alRs.Close
Set alRs = Nothing

If alCount > 0 Then
    Dim mAlRs, mAlId, mAlTitle, mAlYear, mAlSongs, mAlPhoto, mAlSongsRs
    Set mAlRs = SqlRs("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
        "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS SONG_COUNT " & _
        "FROM TBL_ALBUM al ORDER BY al.RELEASE_YEAR DESC, al.TITLE_ALBUM")

    Do While Not mAlRs.EOF
        mAlId = mAlRs("ID_ALBUM")
        mAlTitle = mAlRs("TITLE_ALBUM") & ""
        mAlYear = mAlRs("RELEASE_YEAR") & ""
        mAlSongs = mAlRs("SONG_COUNT")
        mAlPhoto = ImgUrl(mAlRs("ALBUM_IMAGE_PATH") & "", PageDepth & "assets/img/default.jpg")
        If mAlYear = "" Then mAlYear = "s/f"
        Set mAlSongsRs = SongsByAlbum(mAlId)
%>
<div class="modal" id="album-<%=mAlId%>" role="dialog" aria-modal="true" aria-label="Álbum <%=H(mAlTitle)%>">
    <div class="modal__backdrop" data-close-modal></div>
    <div class="modal__card modal__card--wide">
        <div class="modal__head">
            <div style="display:flex;gap:14px;align-items:center">
                <div class="song-cover" style="width:54px;height:54px">
                    <img src="<%=mAlPhoto%>" alt="">
                </div>
                <div>
                    <h3><%=H(mAlTitle)%></h3>
                    <p><%=H(mAlYear)%> · <%=Pluralize(mAlSongs, "canción", "canciones")%></p>
                </div>
            </div>
            <div style="display:flex;gap:8px;align-items:center">
                <button type="button" class="btn btn--primary btn--sm" data-play-all="#album-tracks-<%=mAlId%>">
                    <%=Icon("play")%> Reproducir
                </button>
                <button type="button" class="icon-btn" data-close-modal aria-label="Cerrar"><%=Icon("close")%></button>
            </div>
        </div>
        <div class="modal__body">
            <% If mAlSongsRs.EOF Then %>
            <div class="empty" style="padding:30px 18px">
                <div class="empty__ico"><%=Icon("music")%></div>
                <strong>Álbum sin canciones</strong>
                <p>Este álbum todavía no tiene un tracklist asignado.</p>
            </div>
            <% Else %>
            <% RenderTracklist mAlSongsRs, "album-tracks-" & mAlId %>
            <% End If %>
        </div>
        <div class="modal__foot">
            <span class="modal__foot-note">Año de lanzamiento: <%=H(mAlYear)%></span>
            <button type="button" class="btn btn--ghost" data-close-modal>Cerrar</button>
        </div>
    </div>
</div>
<%
        mAlSongsRs.Close
        Set mAlSongsRs = Nothing
        mAlRs.MoveNext
    Loop
    mAlRs.Close
    Set mAlRs = Nothing
End If
%>

<!--#include file="includes/layout_bottom.asp"-->
