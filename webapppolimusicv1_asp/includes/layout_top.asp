<!--#include file="icons.asp"-->
<%
' ============================================================
'  PoliMusic - Cabecera y estructura comun
'  Requiere: PageTitle, ActiveNav, PageDepth definidos por la pagina
' ============================================================
If PageTitle & "" = "" Then PageTitle = "Inicio"
If ActiveNav & "" = "" Then ActiveNav = ""
If PageDepth & "" = "" Then PageDepth = ""

Dim themeClass, isAdminUser, currentUser, currentPhoto, avatarUrl, roleText, flashMsg, flashKind
isAdminUser = IsAdmin()
If isAdminUser Then themeClass = "theme-admin" Else themeClass = "theme-user"
currentUser = Session("username") & ""
currentPhoto = Session("user_photo") & ""
avatarUrl = ImgUrl(currentPhoto, PageDepth & "assets/img/default.jpg")
If isAdminUser Then roleText = "Administrador" Else roleText = "Usuario"

flashMsg = Session("flash_msg") & ""
flashKind = Session("flash_kind") & ""
If flashKind = "" Then flashKind = "success"
Session("flash_msg") = ""
Session("flash_kind") = ""

Function NavClass(key)
    If LCase(ActiveNav) = LCase(key) Then
        NavClass = "nav-item is-active"
    Else
        NavClass = "nav-item"
    End If
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="theme-color" content="<% If isAdminUser Then %>#12100A<% Else %>#0E0B16<% End If %>">
    <title><%=H(PageTitle)%> · <%=APP_NAME%></title>
    <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Crect width='32' height='32' rx='8' fill='%230E0B16'/%3E%3Cpath d='M5 20.5c3-7.5 5.5-7.5 8.5 0s5.5 7.5 8.5 0 4-6 5-4.5' stroke='%23B9A3F5' stroke-width='2.2' fill='none' stroke-linecap='round'/%3E%3C/svg%3E">
    <link rel="stylesheet" href="<%=PageDepth%>assets/css/app.css">
</head>
<body class="<%=themeClass%>" data-page="<%=H(ActiveNav)%>">
<div class="ambient" aria-hidden="true">
    <span class="ambient__orb ambient__orb--one"></span>
    <span class="ambient__orb ambient__orb--two"></span>
    <span class="ambient__grid"></span>
</div>

<div class="app">
    <aside class="sidebar" id="sidebar">
        <a class="brand" href="<%=PageDepth%>home.asp">
            <span class="brand__mark">
                <svg viewBox="0 0 32 32" fill="none" aria-hidden="true">
                    <path d="M5 20.5c3-7.5 5.5-7.5 8.5 0s5.5 7.5 8.5 0 4-6 5-4.5" stroke="currentColor" stroke-width="2.1" stroke-linecap="round"/>
                    <circle cx="5" cy="20.5" r="2" fill="currentColor"/>
                    <circle cx="27" cy="16" r="2" fill="currentColor"/>
                </svg>
            </span>
            <span class="brand__text">
                <strong><%=APP_NAME%></strong>
                <small>music streaming</small>
            </span>
        </a>

        <nav class="nav">
            <p class="nav__label">Explorar</p>
            <a class="<%=NavClass("home")%>" href="<%=PageDepth%>home.asp">
                <%=Icon("home")%><span>Inicio</span>
            </a>
            <a class="<%=NavClass("songs")%>" href="<%=PageDepth%>home.asp#song-list">
                <%=Icon("music")%><span>Canciones</span>
            </a>
            <a class="<%=NavClass("playlists")%>" href="<%=PageDepth%>playlists.asp">
                <%=Icon("playlist")%><span>Playlists</span>
            </a>
            <a class="<%=NavClass("artists")%>" href="<%=PageDepth%>artists.asp">
                <%=Icon("mic")%><span>Artistas</span>
            </a>
            <a class="<%=NavClass("albums")%>" href="<%=PageDepth%>albums.asp">
                <%=Icon("disc")%><span>Álbumes</span>
            </a>

            <% If isAdminUser Then %>
            <p class="nav__label nav__label--spaced">
                <span>Administración</span>
                <span class="nav__dot"></span>
            </p>
            <a class="<%=NavClass("admin-songs")%>" href="<%=PageDepth%>admin-songs.asp">
                <%=Icon("music")%><span>Canciones</span><em class="nav__tag">CRUD</em>
            </a>
            <a class="<%=NavClass("admin-users")%>" href="<%=PageDepth%>admin-users.asp">
                <%=Icon("users")%><span>Usuarios</span><em class="nav__tag">CRUD</em>
            </a>
            <a class="<%=NavClass("admin-artists")%>" href="<%=PageDepth%>admin-artists.asp">
                <%=Icon("mic")%><span>Artistas</span><em class="nav__tag">CRUD</em>
            </a>
            <a class="<%=NavClass("admin-albums")%>" href="<%=PageDepth%>admin-albums.asp">
                <%=Icon("disc")%><span>Álbumes</span><em class="nav__tag">CRUD</em>
            </a>
            <a class="<%=NavClass("admin-playlists")%>" href="<%=PageDepth%>admin-playlists.asp">
                <%=Icon("playlist")%><span>Playlists</span><em class="nav__tag">CRUD</em>
            </a>
            <% End If %>
        </nav>

        <div class="sidebar__foot">
            <div class="role-panel">
                <span class="role-panel__icon"><% If isAdminUser Then %><%=Icon("shield")%><% Else %><%=Icon("sparkles")%><% End If %></span>
                <div class="role-panel__text">
                    <strong><%=roleText%></strong>
                    <small><% If isAdminUser Then %>Acceso total al catálogo<% Else %>Modo escucha<% End If %></small>
                </div>
            </div>
            <div class="user-card">
                <img class="avatar" src="<%=avatarUrl%>" alt="Foto de <%=H(currentUser)%>">
                <div class="user-card__meta">
                    <strong><%=H(currentUser)%></strong>
                    <span><%=roleText%></span>
                </div>
                <a class="icon-btn icon-btn--ghost" href="<%=PageDepth%>logout.asp" title="Cerrar sesión" aria-label="Cerrar sesión">
                    <%=Icon("logout")%>
                </a>
            </div>
            <a class="logout-link" href="<%=PageDepth%>logout.asp">
                <%=Icon("logout")%><span>Cerrar sesión</span>
            </a>
        </div>
    </aside>

    <div class="sidebar-scrim" id="sidebarScrim"></div>

    <div class="main">
        <header class="topbar">
            <button class="icon-btn topbar__menu" id="menuToggle" aria-label="Abrir menú">
                <%=Icon("menu")%>
            </button>
            <div class="topbar__heading">
                <span class="topbar__crumb"><%=APP_NAME%> <span>/</span> <%=H(PageTitle)%></span>
                <span class="topbar__status"><i class="pulse"></i> Sesión activa</span>
            </div>

            <div class="topbar__search">
                <span class="topbar__search-ico"><%=Icon("search")%></span>
                <input type="search" id="globalSearch" placeholder="Buscar canciones, artistas, playlists..." autocomplete="off" aria-label="Buscar">
                <kbd>/</kbd>
            </div>

            <div class="topbar__right">
                <span class="chip chip--role">
                    <% If isAdminUser Then %><%=Icon("shield")%><% Else %><%=Icon("headphones")%><% End If %>
                    <%=roleText%>
                </span>
                <div class="topbar__user">
                    <img class="avatar avatar--sm" src="<%=avatarUrl%>" alt="">
                    <div class="topbar__user-meta">
                        <strong><%=H(currentUser)%></strong>
                        <span><%=Session("user_email")%></span>
                    </div>
                </div>
            </div>
        </header>

        <main class="content">
