<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard — Riconciliazione Contabile</title>
<link rel="stylesheet" href="/static/css/style.css">
</head>
<body>
<nav class="navbar">
  <div class="nav-brand">
    <svg width="24" height="24" viewBox="0 0 36 36" fill="none">
      <rect width="36" height="36" rx="8" fill="#1F3A5F"/>
      <path d="M10 24L14 12L18 20L22 15L26 24" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
    </svg>
    Riconciliazione Contabile
  </div>
  <div class="nav-right">
    <span class="badge-role badge-{{ session.role }}">{{ session.role }}</span>
    <span class="nav-user">{{ session.username }}</span>
    <a href="/logout" class="btn btn-ghost btn-sm">Esci</a>
  </div>
</nav>

<main class="container">
  <div class="page-header">
    <div>
      <h1>Report di riconciliazione</h1>
      <p class="text-muted">Storico di tutte le riconciliazioni elaborate</p>
    </div>
    {% if session.role == "operatore" %}
    <a href="/nuova" class="btn btn-primary">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
      Nuova riconciliazione
    </a>
    {% endif %}
  </div>

  {% if not reports %}
  <div class="empty-state">
    <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1" stroke-linecap="round" opacity=".3"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
    <p>Nessun report ancora. {% if session.role == "operatore" %}<a href="/nuova">Crea la prima riconciliazione →</a>{% endif %}</p>
  </div>
  {% else %}
  <div class="report-list">
    {% for r in reports %}
    <a href="/report/{{ r.id }}" class="report-card {% if r.stato == 'errore' %}report-card--error{% endif %}">
      <div class="report-card-main">
        <div class="report-periodo">{{ r.periodo }}</div>
        <div class="report-meta">
          <span>{{ r.data }}</span>
          <span>·</span>
          <span>{{ r.operatore }}</span>
          {% if r.note %}<span>·</span><span class="text-muted">{{ r.note[:40] }}</span>{% endif %}
        </div>
      </div>
      <div class="report-card-right">
        {% if r.stato == "completato" %}
        <span class="badge badge-success">Completato</span>
        {% else %}
        <span class="badge badge-error">Errore</span>
        {% endif %}
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" opacity=".4"><polyline points="9 18 15 12 9 6"/></svg>
      </div>
    </a>
    {% endfor %}
  </div>
  {% endif %}
</main>
</body>
</html>
