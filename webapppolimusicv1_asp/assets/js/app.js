/* ============================================================
   PoliMusic · Interacciones de interfaz y reproductor
   ============================================================ */
(function () {
    "use strict";

    var BASE = window.APP_BASE || "";

    /* ---------------- Utilidades ---------------- */
    function $(sel, ctx) { return (ctx || document).querySelector(sel); }
    function $$(sel, ctx) { return Array.prototype.slice.call((ctx || document).querySelectorAll(sel)); }

    function fmtTime(sec) {
        if (!isFinite(sec) || sec < 0) return "0:00";
        var m = Math.floor(sec / 60);
        var s = Math.floor(sec % 60);
        return m + ":" + (s < 10 ? "0" : "") + s;
    }

    function iconSvg(name) {
        var paths = {
            check: "<path d='m5 12.6 4.4 4.4L19 6.8'/>",
            alert: "<path d='M12 4.2 21 19.4H3z'/><path d='M12 9.6v4.2'/><path d='M12 16.8h.01'/>",
            close: "<path d='M6 6l12 12'/><path d='M18 6 6 18'/>"
        };
        return "<svg class='ico' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round' aria-hidden='true'>" +
            (paths[name] || paths.check) + "</svg>";
    }

    /* ---------------- Toasts ---------------- */
    var toastStack = $("#toastStack");

    function toast(message, kind) {
        if (!toastStack) return;
        kind = kind || "success";
        var el = document.createElement("div");
        el.className = "toast toast--" + kind;
        el.setAttribute("role", "status");
        el.innerHTML = "<span class='toast__ico'>" + iconSvg(kind === "error" ? "alert" : "check") +
            "</span><p></p><button class='toast__close' aria-label='Cerrar'>" + iconSvg("close") + "</button>";
        el.querySelector("p").textContent = message;
        toastStack.appendChild(el);

        var timer = setTimeout(remove, 5200);
        function remove() {
            clearTimeout(timer);
            if (!el.parentNode) return;
            el.classList.add("is-out");
            setTimeout(function () { if (el.parentNode) el.parentNode.removeChild(el); }, 300);
        }
        el.querySelector(".toast__close").addEventListener("click", remove);
    }
    window.PoliToast = toast;

    $$(".toast").forEach(function (t) {
        var timer = setTimeout(function () { dismiss(t); }, 5600);
        var close = t.querySelector(".toast__close");
        function dismiss(node) {
            clearTimeout(timer);
            if (!node.parentNode) return;
            node.classList.add("is-out");
            setTimeout(function () { if (node.parentNode) node.parentNode.removeChild(node); }, 300);
        }
        if (close) close.addEventListener("click", function () { dismiss(t); });
    });

    /* ---------------- Sidebar móvil ---------------- */
    var appEl = $(".app");
    var menuToggle = $("#menuToggle");
    var scrim = $("#sidebarScrim");

    function closeNav() { if (appEl) appEl.classList.remove("is-nav-open"); }

    if (menuToggle) {
        menuToggle.addEventListener("click", function () {
            appEl.classList.toggle("is-nav-open");
        });
    }
    if (scrim) scrim.addEventListener("click", closeNav);

    /* ---------------- Modales ---------------- */
    function openModal(id) {
        var modal = document.getElementById(id);
        if (!modal) return;
        modal.classList.add("is-open");
        document.body.style.overflow = "hidden";
        var focusable = modal.querySelector("input[type=text], input[type=password], input[type=email], input[type=date], input[type=number], select, textarea");
        if (!focusable) focusable = modal.querySelector("input:not([type=hidden]), button");
        if (focusable) setTimeout(function () { focusable.focus(); }, 80);
    }

    function closeModal(modal) {
        if (!modal) return;
        modal.classList.remove("is-open");
        if (!$(".modal.is-open")) document.body.style.overflow = "";
    }

    document.addEventListener("click", function (e) {
        var opener = e.target.closest("[data-open-modal]");
        if (opener) {
            e.preventDefault();
            openModal(opener.getAttribute("data-open-modal"));
            return;
        }
        var closer = e.target.closest("[data-close-modal]");
        if (closer) {
            e.preventDefault();
            closeModal(closer.closest(".modal"));
        }
    });

    $$(".modal[data-auto-open='true']").forEach(function (m) {
        openModal(m.id);
    });

    // Abrir un modal desde el hash (#playlist-3)
    if (location.hash && location.hash.length > 1) {
        try {
            var hashTarget = document.querySelector(location.hash);
            if (hashTarget && hashTarget.classList.contains("modal")) {
                openModal(hashTarget.id);
            }
        } catch (err) { /* selector inválido */ }
    }

    document.addEventListener("keydown", function (e) {
        if (e.key === "Escape") {
            var open = $(".modal.is-open");
            if (open) { closeModal(open); return; }
            closeNav();
        }
    });

    window.PoliModal = { open: openModal, close: function (id) { closeModal(document.getElementById(id)); } };

    /* ---------------- Confirmación de borrado ---------------- */
    var pendingForm = null;
    var confirmEl = null;

    function buildConfirm() {
        confirmEl = document.createElement("div");
        confirmEl.className = "modal";
        confirmEl.id = "confirmModal";
        confirmEl.innerHTML =
            "<div class='modal__backdrop' data-close-modal></div>" +
            "<div class='modal__card modal__card--sm' role='dialog' aria-modal='true'>" +
            "<div class='modal__head'><div><h3>Confirmar eliminación</h3><p id='confirmText'></p></div>" +
            "<button class='icon-btn' data-close-modal aria-label='Cerrar'>" + iconSvg("close") + "</button></div>" +
            "<div class='modal__foot'><button class='btn btn--ghost' data-close-modal>Cancelar</button>" +
            "<button class='btn btn--danger' id='confirmOk'>Sí, eliminar</button></div></div>";
        document.body.appendChild(confirmEl);

        confirmEl.querySelector("#confirmOk").addEventListener("click", function () {
            if (!pendingForm) return;
            pendingForm.dataset.confirmed = "1";
            var f = pendingForm;
            pendingForm = null;
            closeModal(confirmEl);
            f.submit();
        });
    }

    document.addEventListener("submit", function (e) {
        var form = e.target;
        if (form.dataset && form.dataset.confirm && !form.dataset.confirmed) {
            e.preventDefault();
            if (!confirmEl) buildConfirm();
            pendingForm = form;
            confirmEl.querySelector("#confirmText").textContent = form.dataset.confirm;
            openModal("confirmModal");
        }
    });

    /* ---------------- Búsqueda global ---------------- */
    var searchInput = $("#globalSearch");
    var searchItems = $$("[data-search-item]");
    var noResults = $("[data-no-results]");

    function normalize(s) {
        return (s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    }

    if (searchInput && searchItems.length) {
        searchInput.addEventListener("input", function () {
            var q = normalize(searchInput.value.trim());
            var visible = 0;
            searchItems.forEach(function (item) {
                var hit = q === "" || normalize(item.getAttribute("data-search-item")).indexOf(q) !== -1;
                item.classList.toggle("is-hidden", !hit);
                if (hit) visible++;
            });
            if (noResults) noResults.classList.toggle("is-visible", visible === 0);
        });
    }

    $$("[data-mini-search]").forEach(function (input) {
        var scopeSel = input.getAttribute("data-mini-search");
        var scope = scopeSel ? $(scopeSel) : document;
        input.addEventListener("input", function () {
            var q = normalize(input.value.trim());
            var rows = $$("[data-search-item]", scope);
            var visible = 0;
            rows.forEach(function (row) {
                var hit = q === "" || normalize(row.getAttribute("data-search-item")).indexOf(q) !== -1;
                row.classList.toggle("is-hidden", !hit);
                if (hit) visible++;
            });
            var counter = scope ? $("[data-visible-count]", scope) : null;
            if (counter) counter.textContent = visible;
        });
    });

    document.addEventListener("keydown", function (e) {
        var tag = (e.target.tagName || "").toLowerCase();
        var typing = tag === "input" || tag === "textarea" || tag === "select" || e.target.isContentEditable;
        if (e.key === "/" && !typing && searchInput) {
            e.preventDefault();
            searchInput.focus();
        }
    });

    /* ---------------- Reproductor ---------------- */
    var audio = $("#audioEl");
    var playBtn = $("#playBtn");
    var prevBtn = $("#prevBtn");
    var nextBtn = $("#nextBtn");
    var shuffleBtn = $("#shuffleBtn");
    var repeatBtn = $("#repeatBtn");
    var seek = $("#seek");
    var volume = $("#volume");
    var muteBtn = $("#muteBtn");
    var curTimeEl = $("#curTime");
    var durTimeEl = $("#durTime");
    var titleEl = $("#playerTitle");
    var artistEl = $("#playerArtist");
    var coverWrap = $("#playerCover");
    var coverImg = $("#playerCoverImg");
    var likeBtn = $("#likeBtn");

    var allTracks = $$(".song-row").map(function (row, i) {
        return {
            id: row.getAttribute("data-song-id"),
            title: row.getAttribute("data-title") || "Sin título",
            artist: row.getAttribute("data-artist") || "Artista desconocido",
            src: row.getAttribute("data-src"),
            cover: row.getAttribute("data-cover"),
            plays: parseInt(row.getAttribute("data-plays") || "0", 10),
            order: i
        };
    });

    var queue = allTracks.slice();
    var currentIndex = -1;
    var isShuffle = false;
    var repeatMode = 0; // 0 off · 1 all · 2 one
    var lastVolume = 0.85;
    var countedFor = null;

    function setFill(el, pct) {
        if (!el) return;
        el.style.setProperty("--fill", Math.max(0, Math.min(100, pct)) + "%");
    }

    function paintRows() {
        var activeId = currentIndex >= 0 && queue[currentIndex] ? queue[currentIndex].id : null;
        $$(".song-row").forEach(function (row) {
            var isActive = activeId !== null && row.getAttribute("data-song-id") === activeId;
            row.classList.toggle("is-current", isActive);
            row.classList.toggle("is-playing", isActive && audio && !audio.paused);
        });
        $$(".track").forEach(function (row) {
            var isActive = activeId !== null && row.getAttribute("data-song-id") === activeId;
            row.classList.toggle("is-current", isActive);
        });
    }

    function likes() {
        try { return JSON.parse(localStorage.getItem("polimusic.likes") || "[]"); }
        catch (err) { return []; }
    }

    function isLiked(id) {
        return likes().indexOf(String(id)) !== -1;
    }

    function toggleLike() {
        if (currentIndex < 0) { toast("Selecciona una canción para guardarla en tus favoritos.", "error"); return; }
        var id = String(queue[currentIndex].id);
        var list = likes();
        var idx = list.indexOf(id);
        if (idx === -1) { list.push(id); toast("Añadida a tus favoritos."); }
        else { list.splice(idx, 1); toast("Eliminada de tus favoritos."); }
        localStorage.setItem("polimusic.likes", JSON.stringify(list));
        paintLike();
        paintRowLikes();
    }

    function paintLike() {
        if (!likeBtn) return;
        var liked = currentIndex >= 0 && isLiked(queue[currentIndex].id);
        likeBtn.classList.toggle("is-liked", liked);
        likeBtn.setAttribute("aria-pressed", liked ? "true" : "false");
    }

    function paintRowLikes() {
        $$("[data-like-id]").forEach(function (btn) {
            var liked = isLiked(btn.getAttribute("data-like-id"));
            btn.classList.toggle("is-liked", liked);
            btn.setAttribute("aria-pressed", liked ? "true" : "false");
        });
    }

    document.addEventListener("click", function (e) {
        var btn = e.target.closest("[data-like-id]");
        if (!btn) return;
        e.preventDefault();
        var id = String(btn.getAttribute("data-like-id"));
        var list = likes();
        var idx = list.indexOf(id);
        if (idx === -1) { list.push(id); toast("Añadida a tus favoritos."); }
        else { list.splice(idx, 1); toast("Eliminada de tus favoritos."); }
        localStorage.setItem("polimusic.likes", JSON.stringify(list));
        paintRowLikes();
        paintLike();
    });

    function loadTrack(index, autoplay) {
        if (!audio || index < 0 || index >= queue.length) return;
        currentIndex = index;
        var t = queue[index];

        audio.src = t.src;
        if (titleEl) titleEl.textContent = t.title;
        if (artistEl) artistEl.textContent = t.artist;
        if (coverImg && t.cover) coverImg.src = t.cover;
        if (coverWrap) coverWrap.classList.remove("is-empty");
        if (likeBtn) likeBtn.disabled = false;
        paintLike();
        paintRows();

        if (autoplay !== false) {
            var p = audio.play();
            if (p && p.catch) {
                p.catch(function () {
                    toast("No se pudo iniciar la reproducción. Revisa el archivo de audio.", "error");
                });
            }
        }

        if (countedFor !== t.id) {
            countedFor = t.id;
            registerPlay(t);
        }
    }

    function registerPlay(track) {
        if (!track || !track.id) return;
        fetch(BASE + "api-play.asp", {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
            body: "id=" + encodeURIComponent(track.id)
        }).then(function (r) { return r.ok ? r.json() : null; })
          .then(function (data) {
              if (!data || !data.ok) return;
              $$("[data-plays-for='" + track.id + "']").forEach(function (el) {
                  el.textContent = data.plays;
              });
          }).catch(function () { /* silencioso */ });
    }

    function togglePlay() {
        if (!audio) return;
        if (!audio.src) {
            if (queue.length) loadTrack(0, true);
            else toast("No hay canciones disponibles en el catálogo.", "error");
            return;
        }
        if (audio.paused) {
            audio.play().catch(function () { toast("No se pudo reproducir el audio.", "error"); });
        } else {
            audio.pause();
        }
    }

    function nextTrack(auto) {
        if (!queue.length) return;
        if (currentIndex < 0) { loadTrack(0, true); return; }
        if (repeatMode === 2 && auto) { audio.currentTime = 0; audio.play(); return; }
        var next;
        if (isShuffle) {
            next = Math.floor(Math.random() * queue.length);
            if (queue.length > 1 && next === currentIndex) next = (next + 1) % queue.length;
        } else {
            next = currentIndex + 1;
            if (next >= queue.length) {
                if (repeatMode === 1 || !auto) next = 0;
                else { audio.pause(); return; }
            }
        }
        loadTrack(next, true);
    }

    function prevTrack() {
        if (!queue.length) return;
        if (audio && audio.currentTime > 4) { audio.currentTime = 0; return; }
        var prev = currentIndex <= 0 ? queue.length - 1 : currentIndex - 1;
        loadTrack(prev, true);
    }

    function playCollection(rows, startIndex) {
        if (!rows || !rows.length) return;
        queue = rows.slice();
        currentIndex = -1;
        loadTrack(startIndex || 0, true);
    }

    // Reproducir desde una fila de la tabla
    document.addEventListener("click", function (e) {
        var btn = e.target.closest("[data-play-id]");
        if (btn) {
            e.preventDefault();
            e.stopPropagation();
            var id = btn.getAttribute("data-play-id");
            var idx = -1;
            for (var i = 0; i < allTracks.length; i++) {
                if (String(allTracks[i].id) === String(id)) { idx = i; break; }
            }
            if (idx === -1) return;
            queue = allTracks.slice();
            if (currentIndex === idx) { togglePlay(); return; }
            loadTrack(idx, true);
            return;
        }

        var playAll = e.target.closest("[data-play-all]");
        if (playAll) {
            e.preventDefault();
            var selector = playAll.getAttribute("data-play-all");
            var scope = selector ? $(selector) : document;
            var rows = $$(".song-row[data-song-id], .track[data-song-id]", scope).map(function (row) {
                return {
                    id: row.getAttribute("data-song-id"),
                    title: row.getAttribute("data-title") || "Sin título",
                    artist: row.getAttribute("data-artist") || "Artista desconocido",
                    src: row.getAttribute("data-src"),
                    cover: row.getAttribute("data-cover")
                };
            });
            if (!rows.length) { toast("No hay canciones disponibles en esta lista.", "error"); return; }
            playCollection(rows, 0);
            return;
        }

        var trackPlay = e.target.closest("[data-track-index]");
        if (trackPlay) {
            e.preventDefault();
            var modal = trackPlay.closest(".modal");
            var tracks = $$(".track[data-song-id]", modal).map(function (row) {
                return {
                    id: row.getAttribute("data-song-id"),
                    title: row.getAttribute("data-title") || "Sin título",
                    artist: row.getAttribute("data-artist") || "Artista desconocido",
                    src: row.getAttribute("data-src"),
                    cover: row.getAttribute("data-cover")
                };
            });
            playCollection(tracks, parseInt(trackPlay.getAttribute("data-track-index"), 10) || 0);
        }
    });

    if (playBtn) playBtn.addEventListener("click", togglePlay);
    if (nextBtn) nextBtn.addEventListener("click", function () { nextTrack(false); });
    if (prevBtn) prevBtn.addEventListener("click", prevTrack);
    if (likeBtn) likeBtn.addEventListener("click", toggleLike);

    if (shuffleBtn) {
        shuffleBtn.addEventListener("click", function () {
            isShuffle = !isShuffle;
            shuffleBtn.classList.toggle("is-active", isShuffle);
            toast(isShuffle ? "Modo aleatorio activado." : "Modo aleatorio desactivado.");
        });
    }

    if (repeatBtn) {
        repeatBtn.addEventListener("click", function () {
            repeatMode = (repeatMode + 1) % 3;
            repeatBtn.classList.toggle("is-active", repeatMode !== 0);
            repeatBtn.setAttribute("data-mode", String(repeatMode));
            if (repeatMode === 2) toast("Repetir canción activado.");
            else if (repeatMode === 1) toast("Repetir toda la lista activado.");
            else toast("Repetición desactivada.");
        });
    }

    if (audio) {
        audio.addEventListener("play", function () {
            document.body.classList.add("is-playing");
            paintRows();
        });
        audio.addEventListener("pause", function () {
            document.body.classList.remove("is-playing");
            paintRows();
        });
        audio.addEventListener("loadedmetadata", function () {
            if (durTimeEl) durTimeEl.textContent = fmtTime(audio.duration);
            if (seek) seek.max = 1000;
        });
        audio.addEventListener("timeupdate", function () {
            if (!audio.duration) return;
            var pct = (audio.currentTime / audio.duration) * 100;
            if (seek) {
                seek.value = String(Math.round(pct * 10));
                setFill(seek, pct);
            }
            if (curTimeEl) curTimeEl.textContent = fmtTime(audio.currentTime);
        });
        audio.addEventListener("ended", function () { nextTrack(true); });
        audio.addEventListener("error", function () {
            if (audio.src) toast("El archivo de audio no está disponible.", "error");
        });
    }

    if (seek) {
        seek.addEventListener("input", function () {
            if (!audio || !audio.duration) return;
            var pct = parseFloat(seek.value) / 10;
            audio.currentTime = (pct / 100) * audio.duration;
            setFill(seek, pct);
        });
    }

    function applyVolume(v) {
        v = Math.max(0, Math.min(1, v));
        if (audio) audio.volume = v;
        if (volume) {
            volume.value = String(v);
            setFill(volume, v * 100);
        }
        if (muteBtn) muteBtn.classList.toggle("is-active", v === 0);
    }

    if (volume) {
        volume.addEventListener("input", function () {
            lastVolume = parseFloat(volume.value) || 0;
            applyVolume(lastVolume);
        });
        applyVolume(parseFloat(volume.value) || 0.85);
    }

    if (muteBtn) {
        muteBtn.addEventListener("click", function () {
            if (parseFloat(volume ? volume.value : 0) > 0) {
                lastVolume = parseFloat(volume.value);
                applyVolume(0);
            } else {
                applyVolume(lastVolume || 0.85);
            }
        });
    }

    // Barra espaciadora: play / pausa
    document.addEventListener("keydown", function (e) {
        if (e.code !== "Space" && e.key !== " ") return;
        var tag = (e.target.tagName || "").toLowerCase();
        if (tag === "input" || tag === "textarea" || tag === "select" || tag === "button" || e.target.isContentEditable) return;
        if ($(".modal.is-open")) return;
        e.preventDefault();
        togglePlay();
    });

    // Marca la primera fila como referencia y prepara el estado inicial
    if (allTracks.length && coverWrap) {
        coverWrap.classList.add("is-empty");
    }
    paintRows();
    paintRowLikes();

    /* ---------------- Cuentas demo (login) ---------------- */
    $$("[data-demo-user]").forEach(function (btn) {
        btn.addEventListener("click", function () {
            var u = $("#username");
            var p = $("#password");
            if (u) u.value = btn.getAttribute("data-demo-user");
            if (p) p.value = btn.getAttribute("data-demo-pass") || "12345678";
            if (u) u.focus();
        });
    });

    /* ---------------- Año dinámico ---------------- */
    $$("[data-year]").forEach(function (el) { el.textContent = String(new Date().getFullYear()); });
})();
