<%@ Language=VBScript CodePage=65001 %>
<% Dim PageDepth : PageDepth = "" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/admin_helpers.asp"-->
<%
RequireAdmin

Dim PageTitle, ActiveNav
PageTitle = "Administrar artistas"
ActiveNav = "admin-artists"

Dim action, postId, errMsg, formName, formLast, formPhoto, isEdit, openModal, editId
Dim c, rsEdit, newId, rsNew

isEdit = False
openModal = False
editId = ""
formName = ""
formLast = ""
formPhoto = ""

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    action = Request.Form("action") & ""
    postId = Trim(Request.Form("id") & "")

    If action = "delete" Then
        If postId <> "" And IsNumeric(postId) Then
            On Error Resume Next
            gConn.Execute "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = " & CLng(postId)
            gConn.Execute "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = " & CLng(postId)
            If Err.Number <> 0 Then
                SetFlash "No se pudo eliminar el artista: " & Err.Description, "error"
                Err.Clear
            Else
                SetFlash "Artista eliminado correctamente.", "success"
            End If
            On Error GoTo 0
        End If
        Response.Redirect "admin-artists.asp"
    End If

    If action = "save" Then
        formName = TrimText(Request.Form("artist_name"), 100)
        formLast = TrimText(Request.Form("artist_lastname"), 100)
        formPhoto = TrimText(Request.Form("artist_image"), 255)

        If formName = "" Then errMsg = "El nombre del artista es obligatorio."

        If errMsg <> "" Then
            SetFlash errMsg, "error"
            If postId <> "" Then
                Response.Redirect "admin-artists.asp?edit=" & postId
            Else
                Response.Redirect "admin-artists.asp?new=1"
            End If
        End If

        On Error Resume Next
        If postId = "" Then
            Set c = NewCmd("INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) OUTPUT INSERTED.ID_ARTIST AS NEW_ID VALUES (?, ?, ?)")
            AddStr c, "@name", formName
            AddStr c, "@lastname", formLast
            AddStr c, "@image", formPhoto
            Set rsNew = c.Execute
            If Err.Number = 0 Then
                newId = rsNew("NEW_ID")
                rsNew.Close
            End If
        Else
            Set c = NewCmd("UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?")
            AddStr c, "@name", formName
            AddStr c, "@lastname", formLast
            AddStr c, "@image", formPhoto
            AddInt c, "@id", CLng(postId)
            c.Execute
        End If

        If Err.Number <> 0 Then
            SetFlash "No se pudo guardar el artista: " & Err.Description, "error"
            Err.Clear
            On Error GoTo 0
            If postId <> "" Then
                Response.Redirect "admin-artists.asp?edit=" & postId
            Else
                Response.Redirect "admin-artists.asp?new=1"
            End If
        End If
        On Error GoTo 0

        If postId = "" Then
            SetFlash "Artista «" & Trim(formName & " " & formLast) & "» creado correctamente.", "success"
        Else
            SetFlash "Artista «" & Trim(formName & " " & formLast) & "» actualizado correctamente.", "success"
        End If
        Response.Redirect "admin-artists.asp"
    End If
End If

editId = Request.QueryString("edit") & ""
If editId <> "" And IsNumeric(editId) Then
    Set c = NewCmd("SELECT ID_ARTIST, ARTIST_NAME, ISNULL(ARTIST_LASTNAME,'') AS ARTIST_LASTNAME, ISNULL(ARTIST_IMAGE,'') AS ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?")
    AddInt c, "@id", CLng(editId)
    Set rsEdit = c.Execute
    If Not rsEdit.EOF Then
        isEdit = True
        openModal = True
        editId = rsEdit("ID_ARTIST")
        formName = rsEdit("ARTIST_NAME") & ""
        formLast = rsEdit("ARTIST_LASTNAME") & ""
        formPhoto = rsEdit("ARTIST_IMAGE") & ""
        rsEdit.Close
    Else
        rsEdit.Close
    End If
End If

If Request.QueryString("new") = "1" Then openModal = True

Dim rsArtists, totalArtists
Set rsArtists = SqlRs("SELECT a.ID_ARTIST, a.ARTIST_NAME, ISNULL(a.ARTIST_LASTNAME,'') AS ARTIST_LASTNAME, ISNULL(a.ARTIST_IMAGE,'') AS ARTIST_IMAGE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS SONG_COUNT " & _
    "FROM TBL_ARTIST a ORDER BY a.ARTIST_NAME")
totalArtists = ScalarVal("SELECT COUNT(*) FROM TBL_ARTIST")
%>
<!--#include file="includes/layout_top.asp"-->

<div class="page-head">
    <div>
        <span class="page-head__eyebrow"><%=Icon("shield")%> Panel administrativo</span>
        <h1>Artistas</h1>
        <p>Registra las voces del catálogo. Cada artista puede estar vinculado a varias canciones mediante la tabla TBL_PLAYER.</p>
    </div>
    <div class="page-head__actions">
        <a class="btn btn--ghost" href="home.asp"><%=Icon("arrow-left")%> Ir al inicio</a>
        <a class="btn btn--primary" href="admin-artists.asp?new=1"><%=Icon("plus")%> Nuevo artista</a>
    </div>
</div>

<div class="admin-bar">
    <span class="admin-bar__count"><strong><%=totalArtists%></strong> artistas registrados</span>
    <span class="admin-bar__spacer"></span>
    <label class="mini-search">
        <%=Icon("search")%>
        <input type="search" placeholder="Filtrar artistas..." data-mini-search="#artistsTable" aria-label="Filtrar artistas">
    </label>
</div>

<div class="table-wrap">
    <table class="song-table" id="artistsTable">
        <thead>
            <tr>
                <th style="width:54px">ID</th>
                <th>Artista</th>
                <th class="hide-sm">Imagen</th>
                <th>Canciones</th>
                <th class="cell-actions">Acciones</th>
            </tr>
        </thead>
        <tbody>
        <%
        Dim aId, aName, aLast, aFull, aPhoto, aSongs, aSearch
        Do While Not rsArtists.EOF
            aId = rsArtists("ID_ARTIST")
            aName = rsArtists("ARTIST_NAME") & ""
            aLast = rsArtists("ARTIST_LASTNAME") & ""
            aFull = Trim(aName & " " & aLast)
            aPhoto = ImgUrl(rsArtists("ARTIST_IMAGE") & "", "")
            aSongs = rsArtists("SONG_COUNT")
            aSearch = aId & " " & aFull
        %>
            <tr data-search-item="<%=H(aSearch)%>">
                <td class="muted" style="font-variant-numeric:tabular-nums">#<%=aId%></td>
                <td>
                    <div class="song-cell">
                        <div class="song-cover" style="border-radius:50%">
                            <% If aPhoto <> "" Then %>
                            <img src="<%=aPhoto%>" alt="">
                            <% Else %>
                            <div class="card__initials" style="font-size:14px"><%=H(Initials(aFull))%></div>
                            <% End If %>
                        </div>
                        <div class="song-name">
                            <strong><%=H(aFull)%></strong>
                            <span>Artista del catálogo</span>
                        </div>
                    </div>
                </td>
                <td class="hide-sm"><span class="tag tag--plain"><%=H(rsArtists("ARTIST_IMAGE") & "")%></span></td>
                <td><span class="plays"><%=Icon("music")%> <%=aSongs%></span></td>
                <td class="cell-actions">
                    <div class="cell-actions__inner">
                        <a class="icon-btn icon-btn--sm" href="admin-artists.asp?edit=<%=aId%>" title="Editar" aria-label="Editar <%=H(aFull)%>"><%=Icon("edit")%></a>
                        <form method="post" action="admin-artists.asp" style="display:inline"
                              data-confirm="¿Eliminar al artista «<%=H(aFull)%>»? Se eliminarán sus vínculos con las canciones.">
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=aId%>">
                            <button type="submit" class="icon-btn icon-btn--sm icon-btn--danger" title="Eliminar" aria-label="Eliminar <%=H(aFull)%>"><%=Icon("trash")%></button>
                        </form>
                    </div>
                </td>
            </tr>
        <%
            rsArtists.MoveNext
        Loop
        rsArtists.Close
        Set rsArtists = Nothing
        %>
        </tbody>
    </table>
</div>

<div class="modal" id="artistForm" <% If openModal Then %>data-auto-open="true"<% End If %> role="dialog" aria-modal="true" aria-label="Formulario de artista">
    <div class="modal__backdrop" data-close-modal></div>
    <div class="modal__card">
        <form method="post" action="admin-artists.asp">
            <div class="modal__head">
                <div>
                    <h3><% If isEdit Then %>Editar artista<% Else %>Nuevo artista<% End If %></h3>
                    <p>Completa los datos del artista. Solo el nombre es obligatorio.</p>
                </div>
                <button type="button" class="icon-btn" data-close-modal aria-label="Cerrar"><%=Icon("close")%></button>
            </div>
            <div class="modal__body">
                <input type="hidden" name="action" value="save">
                <input type="hidden" name="id" value="<% If isEdit Then %><%=editId%><% End If %>">
                <div class="form-grid">
                    <div class="field">
                        <label for="artist_name">Nombre *</label>
                        <input type="text" id="artist_name" name="artist_name" maxlength="100" required
                               value="<%=H(formName)%>" placeholder="Ej. Juan">
                    </div>
                    <div class="field">
                        <label for="artist_lastname">Apellido</label>
                        <input type="text" id="artist_lastname" name="artist_lastname" maxlength="100"
                               value="<%=H(formLast)%>" placeholder="Ej. Gabriel">
                    </div>
                    <div class="field field--full">
                        <label for="artist_image">Ruta de la imagen</label>
                        <input type="text" id="artist_image" name="artist_image" maxlength="255"
                               value="<%=H(formPhoto)%>" placeholder="~/ArtistImages/artist1.jpg">
                        <span class="field__hint">Si la imagen no existe en <code>/assets/img</code>, se muestra un avatar con las iniciales.</span>
                    </div>
                </div>
            </div>
            <div class="modal__foot">
                <span class="modal__foot-note">Tabla: TBL_ARTIST</span>
                <button type="button" class="btn btn--ghost" data-close-modal>Cancelar</button>
                <button type="submit" class="btn btn--primary"><%=Icon("check")%> Guardar artista</button>
            </div>
        </form>
    </div>
</div>

<!--#include file="includes/layout_bottom.asp"-->
