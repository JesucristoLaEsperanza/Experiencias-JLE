/*
 * ENTRADAS JLE — LÓGICA DE LA PORTADA
 * Este archivo coordina el catálogo público, las tarjetas y la ventana de reserva.
 * Los datos de eventos y precios se leen desde Google Sheets mediante Apps Script.
 */
(function () {
  'use strict';

  // ── 1. Elementos y configuración ──────────────────────────────────────────
  const appScriptUrl = window.JLE_APPS_SCRIPT_URL || '';
  const grid = document.getElementById('eventGrid');
  const catalogStatus = document.getElementById('catalogStatus');
  const catalogRetry = document.getElementById('catalogRetry');
  const dialog = document.getElementById('registrationModal');
  const frame = document.getElementById('registrationFrame');
  const loading = document.getElementById('formLoading');
  const openTab = document.getElementById('openFormTab');
  const title = document.getElementById('registrationTitle');
  const posters = {
    'activados-2027': { path: 'assets/afiche-campamento-2027.jpg', alt: 'Afiche del campamento Activados 2027' },
    'jle-kids-2027': { path: 'assets/anuncio-campamento-ninos.jpg', alt: 'Afiche del campamento JLE Kids' }
  };

  // ── 2. Formato de fechas, edades y precios ─────────────────────────────────
  function parseDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))) return null;
    const [year, month, day] = value.split('-').map(Number);
    return new Date(year, month - 1, day, 12);
  }

  function formatDate(value) {
    const date = parseDate(value);
    return date ? new Intl.DateTimeFormat('es-AR', { day: 'numeric', month: 'short', timeZone: 'America/Argentina/Buenos_Aires' }).format(date).replace('.', '') : '';
  }

  function dateRange(event) {
    const start = parseDate(event.inicio);
    const end = parseDate(event.fin);
    if (!start || !end) return 'Fecha a confirmar';
    const sameDay = event.inicio === event.fin;
    const sameMonth = start.getMonth() === end.getMonth() && start.getFullYear() === end.getFullYear();
    const first = sameDay ? formatDate(event.inicio) : (sameMonth ? String(start.getDate()) : formatDate(event.inicio));
    return first + (sameDay ? '' : ' – ' + formatDate(event.fin)) + ' ' + end.getFullYear();
  }

  function ageRange(event) {
    const minimum = Number(event.edadMin) || 0;
    const maximum = Number(event.edadMax) || 0;
    if (!minimum && !maximum) return 'Todas las edades';
    if (minimum && !maximum) return 'Desde ' + minimum + ' años';
    return 'De ' + (minimum || 0) + ' a ' + maximum + ' años';
  }

  function priceText(event) {
    const price = Number(event.precio) || 0;
    return price > 0
      ? new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS', maximumFractionDigits: 0 }).format(price)
      : (event.detallePrecio || 'A confirmar');
  }

  // ── 3. Creación segura de las tarjetas del catálogo ───────────────────────
  function textElement(tag, text, className) {
    const element = document.createElement(tag);
    if (className) element.className = className;
    element.textContent = text;
    return element;
  }

  function addFact(container, label, value) {
    const fact = document.createElement('span');
    fact.append(textElement('b', label), document.createTextNode(value));
    container.appendChild(fact);
  }

  function createArtwork(event) {
    const wrap = document.createElement('div');
    wrap.className = 'poster-wrap';
    const poster = posters[event.id];
    if (poster) {
      const image = document.createElement('img');
      image.className = 'poster';
      image.src = poster.path;
      image.alt = poster.alt;
      image.loading = 'lazy';
      wrap.appendChild(image);
      return wrap;
    }

    // Los nuevos eventos (como el Congreso) reciben una portada ilustrada sin imagen extra.
    wrap.classList.add('poster-wrap--ticket');
    const ticketArt = document.createElement('div');
    ticketArt.className = 'ticket-art';
    ticketArt.append(
      textElement('span', 'JLE · ' + (event.tipo || 'ENCUENTRO').toUpperCase(), 'ticket-art__eyebrow'),
      textElement('strong', 'NOS\nVEMOS', 'ticket-art__title'),
      textElement('span', 'FE · AMIGOS · PROPÓSITO', 'ticket-art__footer')
    );
    wrap.appendChild(ticketArt);
    return wrap;
  }

  function createEventCard(event, index) {
    const card = document.createElement('article');
    const isKids = event.id === 'jle-kids-2027';
    card.className = 'event-card' + (isKids ? ' card-kids' : (event.tipo === 'Congreso' ? ' card-congress' : ' card-youth'));
    card.appendChild(createArtwork(event));

    const info = document.createElement('div');
    info.className = 'event-info';
    const kicker = document.createElement('div');
    kicker.className = 'event-kicker';
    const audience = event.publico || event.tipo || 'JLE';
    const tagClass = isKids ? 'tag tag-orange' : (index % 2 ? 'tag tag-orange' : 'tag tag-cyan');
    kicker.append(textElement('span', audience, tagClass), textElement('span', (event.inicio || '').slice(5, 7) + ' / ' + (event.inicio || '').slice(0, 4), 'event-year'));
    info.appendChild(kicker);

    const heading = document.createElement('h3');
    const eventName = String(event.nombre || 'Experiencia JLE');
    const words = eventName.split(' ');
    heading.append(document.createTextNode(words.slice(0, -1).join(' ') + (words.length > 1 ? ' ' : '')),
      textElement('span', words.length > 1 ? words[words.length - 1] : 'JLE'));
    info.appendChild(heading);
    info.appendChild(textElement('p', event.descripcion || 'Un encuentro para compartir, crecer en la fe y vivir algo inolvidable.', 'event-description'));

    const facts = document.createElement('div');
    facts.className = 'event-facts';
    addFact(facts, 'FECHA', dateRange(event));
    addFact(facts, 'LUGAR', event.lugar || 'A confirmar');
    addFact(facts, 'EDAD', ageRange(event));
    info.appendChild(facts);

    const bottom = document.createElement('div');
    bottom.className = 'event-ticket-bottom';
    const price = document.createElement('div');
    price.className = 'event-ticket-price';
    price.append(textElement('strong', priceText(event)), textElement('small', 'por persona'));
    const button = textElement('button', 'Conseguir entrada →', 'button button-dark');
    button.type = 'button';
    button.dataset.eventId = String(event.id || '');
    button.dataset.eventName = eventName;
    bottom.append(price, button);
    info.append(bottom);
    card.appendChild(info);
    return card;
  }

  function renderCatalog(events) {
    grid.replaceChildren();
    if (!events.length) {
      catalogStatus.textContent = 'No encontramos experiencias próximas. Revisá en Google Sheets que los campamentos tengan fecha futura y estén abiertos.';
      catalogRetry.classList.remove('is-hidden');
      return;
    }
    const fragment = document.createDocumentFragment();
    events.forEach((event, index) => fragment.appendChild(createEventCard(event, index)));
    grid.appendChild(fragment);
    catalogStatus.classList.add('is-hidden');
  }

  // ── 4. Consulta del catálogo en Google Sheets ──────────────────────────────
  function loadCatalog() {
    if (window.location.protocol === 'file:') {
      catalogStatus.textContent = 'Abriste index.html como archivo local. Para conectar Google Sheets, usá GitHub Pages o iniciá un servidor local en http://localhost.';
      catalogRetry.classList.add('is-hidden');
      return;
    }
    if (!appScriptUrl.startsWith('https://script.google.com/macros/s/') || !appScriptUrl.endsWith('/exec')) {
      catalogStatus.textContent = 'Falta configurar la URL de Apps Script en site-config.js.';
      return;
    }

    const url = new URL(appScriptUrl);
    url.searchParams.set('action', 'eventos');
    url.searchParams.set('callback', 'JLECatalogResponse');
    url.searchParams.set('_', String(Date.now()));

    const script = document.createElement('script');
    script.src = url.toString();
    script.async = true;
    script.onerror = function () {
      catalogStatus.textContent = 'No pudimos cargar las experiencias. Actualizá la página o revisá la conexión con Google Sheets.';
      catalogRetry.classList.remove('is-hidden');
      delete window.JLECatalogResponse;
    };
    window.JLECatalogResponse = function (response) {
      script.remove();
      if (!response || response.ok !== true || !Array.isArray(response.eventos)) {
        catalogStatus.textContent = (response && response.error) || 'No se pudo leer el catálogo de Google Sheets.';
        catalogRetry.classList.remove('is-hidden');
        return;
      }
      catalogRetry.classList.add('is-hidden');
      renderCatalog(response.eventos);
    };
    document.head.appendChild(script);
  }

  // ── 5. Apertura del formulario y estados de la ventana ────────────────────
  function openRegistration(eventId, eventName) {
    const url = new URL(appScriptUrl);
    url.searchParams.set('campamento', eventId);
    title.textContent = 'Entrada · ' + eventName;
    loading.classList.remove('is-hidden');
    frame.src = url.toString();
    openTab.href = url.toString();
    dialog.showModal();
  }

  document.getElementById('closeForm').addEventListener('click', () => dialog.close());
  dialog.addEventListener('click', event => { if (event.target === dialog) dialog.close(); });
  frame.addEventListener('load', () => loading.classList.add('is-hidden'));
  grid.addEventListener('click', event => {
    const button = event.target.closest('[data-event-id]');
    if (button) openRegistration(button.dataset.eventId, button.dataset.eventName);
  });
  catalogRetry.addEventListener('click', loadCatalog);

  // ── 6. Inicio ─────────────────────────────────────────────────────────────
  loadCatalog();
})();
