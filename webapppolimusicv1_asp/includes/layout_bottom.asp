        </main>

        <footer class="footer">
            <span><%=APP_NAME%> · Proyecto académico · ASP clásico + SQL Server</span>
            <span class="footer__sep"></span>
            <span>Sesión de <strong><%=H(currentUser)%></strong> · <%=roleText%></span>
        </footer>
    </div>
</div>

<!-- ===================== Reproductor global ===================== -->
<div class="player" id="playerBar">
    <audio id="audioEl" preload="metadata"></audio>

    <div class="player__now">
        <div class="player__cover" id="playerCover">
            <img id="playerCoverImg" src="<%=PageDepth%>assets/img/default.jpg" alt="">
            <span class="player__cover-fallback"><%=Icon("music")%></span>
        </div>
        <div class="player__meta">
            <strong id="playerTitle">Nada reproduciéndose</strong>
            <span id="playerArtist">Elige una canción del listado</span>
        </div>
        <button class="icon-btn player__like" id="likeBtn" aria-label="Me gusta" title="Me gusta">
            <%=Icon("heart")%>
        </button>
    </div>

    <div class="player__center">
        <div class="player__controls">
            <button class="icon-btn" id="shuffleBtn" aria-label="Aleatorio" title="Aleatorio"><%=Icon("shuffle")%></button>
            <button class="icon-btn" id="prevBtn" aria-label="Anterior" title="Anterior"><%=Icon("skipback")%></button>
            <button class="player__play" id="playBtn" aria-label="Reproducir" title="Reproducir">
                <span class="ic-play"><%=Icon("play")%></span>
                <span class="ic-pause"><%=Icon("pause")%></span>
            </button>
            <button class="icon-btn" id="nextBtn" aria-label="Siguiente" title="Siguiente"><%=Icon("skipforward")%></button>
            <button class="icon-btn" id="repeatBtn" aria-label="Repetir" title="Repetir"><%=Icon("repeat")%></button>
        </div>
        <div class="player__timeline">
            <span class="player__time" id="curTime">0:00</span>
            <div class="range-wrap">
                <input type="range" id="seek" min="0" max="1000" value="0" step="1" aria-label="Progreso de la canción">
            </div>
            <span class="player__time" id="durTime">0:00</span>
        </div>
    </div>

    <div class="player__right">
        <button class="icon-btn" id="muteBtn" aria-label="Silenciar" title="Silenciar"><%=Icon("volume")%></button>
        <div class="range-wrap range-wrap--sm">
            <input type="range" id="volume" min="0" max="1" step="0.01" value="0.85" aria-label="Volumen">
        </div>
        <div class="player__brand"><%=Icon("headphones")%><span>PoliMusic</span></div>
    </div>
</div>

<div class="toast-stack" id="toastStack">
    <% If flashMsg <> "" Then %>
    <div class="toast toast--<%=H(flashKind)%>" role="status">
        <span class="toast__ico"><% If flashKind = "error" Then %><%=Icon("alert")%><% Else %><%=Icon("check")%><% End If %></span>
        <p><%=H(flashMsg)%></p>
        <button class="toast__close" aria-label="Cerrar"><%=Icon("close")%></button>
    </div>
    <% End If %>
</div>

<script>
    window.APP_BASE = "<%=PageDepth%>";
    window.APP_ROLE = "<%=roleText%>";
</script>
<script src="<%=PageDepth%>assets/js/app.js"></script>
<%
CloseDbConnection
%>
</body>
</html>
