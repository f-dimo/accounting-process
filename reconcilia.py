<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ report.periodo }} — Riconciliazione</title>
<link rel="stylesheet" href="/static/css/style.css">
</head>
<body>
<nav class="navbar">
  <div class="nav-brand">
    <a href="/" class="nav-back">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><polyline points="15 18 9 12 15 6"/></svg>
    </a>
    <svg width="24" height="24" viewBox="0 0 36 36" fill="none">
      <rect width="36" height="36" rx="8" fill="#1F3A5F"/>
      <path d="M10 24L14 12L18 20L22 15L26 24" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
    </svg>
    {{ report.periodo }}
  </div>
  <div class="nav-right">
    <span class="badge-role badge-{{ session.role }}">{{ session.role }}</span>
    <a href="/logout" class="btn btn-ghost btn-sm">Esci</a>
  </div>
</nav>

<main class="container container-narrow">

  {% if report.stato == "errore" %}
  <div class="alert alert-error" style="margin-bottom:1.5rem">
    <strong>Errore durante l'elaborazione:</strong> {{ report.errore }}
  </div>
  {% endif %}

  <div class="detail-header">
    <div>
      <h1>{{ report.periodo }}</h1>
      <p class="text-muted">Elaborato il {{ report.data }} da {{ report.operatore }}</p>
      {% if report.note %}<p class="text-muted">{{ report.note }}</p>{% endif %}
    </div>
    {% if report.stato == "completato" %}
    <a href="/download/{{ report.id }}" class="btn btn-primary">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Scarica Excel
    </a>
    {% endif %}
  </div>

  {% if report.stato == "completato" %}
  <div class="info-card">
    <div class="info-row">
      <span class="info-label">Stato</span>
      <span class="badge badge-success">Completato</span>
    </div>
    <div class="info-row">
      <span class="info-label">Periodo</span>
      <span>{{ report.periodo }}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Elaborato da</span>
      <span>{{ report.operatore }}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Data elaborazione</span>
      <span>{{ report.data }}</span>
    </div>
    {% if report.note %}
    <div class="info-row">
      <span class="info-label">Note</span>
      <span>{{ report.note }}</span>
    </div>
    {% endif %}
  </div>

  <div class="action-box">
    <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1" stroke-linecap="round" opacity=".4"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="12" y1="12" x2="12" y2="18"/><line x1="9" y1="15" x2="15" y2="15"/></svg>
    <p>Il report Excel contiene 6 fogli: Riepilogo, Estratto Conto, Partitario Fornitori, Partitario Clienti, Contestazioni e Anomalie.</p>
    <a href="/download/{{ report.id }}" class="btn btn-primary">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Scarica il report Excel
    </a>
  </div>
  {% endif %}

</main>
</body>
</html>
