'use strict';

  /* =========================================================
     Estado del documento
     ========================================================= */
  const editor = document.getElementById('editor');
  const toolbar = document.getElementById('toolbar');
  const pageEl = document.getElementById('page');
  const workspaceEl = document.querySelector('.workspace');
  const zoomRange = document.getElementById('zoomRange');
  const zoomLabel = document.getElementById('zoomLabel');
  const autosaveStatus = document.getElementById('autosaveStatus');
  const findBar = document.getElementById('findBar');
  const findText = document.getElementById('findText');
  const replaceText = document.getElementById('replaceText');
  const findCount = document.getElementById('findCount');

  let currentName = null;     // nombre del archivo actual
  let currentHandle = null;   // FileSystemFileHandle (si el navegador lo permite)
  let currentFormat = null;   // html | txt | docx | odt
  let dirty = false;          // cambios sin guardar
  let versionDocumento = 0;   // permite detectar cambios ocurridos durante un guardado
  let savedRange = null;      // última selección dentro del editor
  let zoomActual = 100;       // porcentaje de zoom
  let ajusteAutomatico = false; // true si el zoom sigue el ancho de la ventana

  let resolverGuardar = null;  // promesa pendiente del diálogo "Guardar como"
  let enlaceEnEdicion = null;  // <a> que se está editando
  let rangoEnlace = null;      // selección guardada al abrir el diálogo de enlace

  const AUTOSAVE_KEY = 'procesadorCopiaLocal';
  const AUTOSAVE_ESPERA = 2000;
  const AUTOSAVE_INTERVALO = 30000;
  const FOCUS_KEY = 'procesadorEnfoquePredeterminado';

  let temporizadorAutoguardado = null;
  let copiaPendiente = null;
  let modoEnfoque = false;
  let pegarSinFormato = false;

  let hallazgos = [];
  let hallazgoActual = -1;
  let mapaBusqueda = null;
  let textoBuscado = '';
  let temporizadorBusqueda = null;

  const BLOQUES_TEXTO = new Set([
    'ADDRESS', 'ARTICLE', 'BLOCKQUOTE', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6',
    'LI', 'P', 'PRE', 'SECTION', 'TABLE', 'TD', 'TH', 'TR'
  ]);

  const TIPOS_GUARDADO = [
    { description: 'Documento ODT', accept: { 'application/vnd.oasis.opendocument.text': ['.odt'] } },
    { description: 'Documento Word (DOCX)', accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] } },
    { description: 'Página HTML', accept: { 'text/html': ['.html'] } },
    { description: 'Texto plano', accept: { 'text/plain': ['.txt'] } }
  ];

  /* =========================================================
     Selección: guardar / restaurar
     ========================================================= */
  document.addEventListener('selectionchange', () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount) {
      const range = sel.getRangeAt(0);
      if (editor.contains(range.commonAncestorContainer)) {
        savedRange = range.cloneRange();
      }
    }
    updateToolbar();
  });

  function restoreSelection() {
    editor.focus();

    const sel = window.getSelection();

    // Si la selección actual ya está dentro del editor, se conserva.
    if (sel && sel.rangeCount) {
      const range = sel.getRangeAt(0);
      if (editor.contains(range.commonAncestorContainer)) return;
    }

    // Si se perdió (por ejemplo al pulsar un desplegable), se restaura.
    if (savedRange) {
      sel.removeAllRanges();
      sel.addRange(savedRange);
    }
  }

  /* Evita que los botones roben el foco/selección del editor */
  toolbar.querySelectorAll('button').forEach(btn => {
    btn.addEventListener('mousedown', e => e.preventDefault());
  });

  /* =========================================================
     Menús desplegables
     ========================================================= */
  const menusDesplegables = [];

  function crearMenuDesplegable(btnId, menuId) {
    const btn = document.getElementById(btnId);
    const menu = document.getElementById(menuId);

    const api = {
      btn,
      menu,
      estaAbierto() { return !menu.hidden; },
      abrir() {
        menusDesplegables.forEach(otro => { if (otro !== api) otro.cerrar(); });
        menu.hidden = false;
        btn.setAttribute('aria-expanded', 'true');
      },
      cerrar() {
        if (menu.hidden) return;
        menu.hidden = true;
        btn.setAttribute('aria-expanded', 'false');
      },
      alternar() {
        if (menu.hidden) api.abrir();
        else api.cerrar();
      }
    };

    document.addEventListener('click', e => {
      if (api.estaAbierto() && !btn.parentElement.contains(e.target)) api.cerrar();
    });

    btn.addEventListener('keydown', e => {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        api.abrir();
        const first = menu.querySelector('button');
        if (first) first.focus();
      }
    });

    menu.addEventListener('keydown', e => {
      const items = Array.from(menu.querySelectorAll('button'));
      const i = items.indexOf(document.activeElement);

      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        const next = e.key === 'ArrowDown'
          ? Math.min(i + 1, items.length - 1)
          : Math.max(i - 1, 0);
        items[next].focus();
      } else if (e.key === 'Home') {
        e.preventDefault();
        items[0].focus();
      } else if (e.key === 'End') {
        e.preventDefault();
        items[items.length - 1].focus();
      }
    });

    menusDesplegables.push(api);
    return api;
  }

  const menuArchivo = crearMenuDesplegable('fileMenuBtn', 'fileMenu');
  const menuInterlineado = crearMenuDesplegable('spacingMenuBtn', 'spacingMenu');

  function cerrarMenuArchivo() { menuArchivo.cerrar(); }
  function alternarMenuArchivo() { menuArchivo.alternar(); }

  function accionMenuArchivo(accion) {
    menuArchivo.cerrar();
    accion();
  }

  function alternarMenuInterlineado() { menuInterlineado.alternar(); }

  function elegirInterlineado(valor) {
    menuInterlineado.cerrar();
    aplicarInterlineado(valor);
  }

  function cerrarMenusDesplegables() {
    menusDesplegables.forEach(menu => menu.cerrar());
  }

  /* =========================================================
     Formato
     ========================================================= */
  function exec(command, value) {
    restoreSelection();
    document.execCommand(command, false, value ?? null);
    const sel = window.getSelection();
    if (sel && sel.rangeCount) savedRange = sel.getRangeAt(0).cloneRange();
    updateToolbar();
  }

  function aplicarFuente(font) {
    if (!font) return;
    restoreSelection();
    document.execCommand('fontName', false, font);
    updateToolbar();
  }

  function aplicarTamano(pt) {
    if (!pt) return;
    restoreSelection();

    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;

    const range = sel.getRangeAt(0);

    // Sin texto seleccionado: se prepara el tamaño para lo que se escriba a continuación.
    if (range.collapsed) {
      const span = document.createElement('span');
      span.style.fontSize = pt + 'pt';
      span.appendChild(document.createTextNode('\u200B')); // espacio de ancho cero
      range.insertNode(span);

      const caret = document.createRange();
      caret.setStart(span.firstChild, 1);
      caret.collapse(true);
      sel.removeAllRanges();
      sel.addRange(caret);
      savedRange = caret.cloneRange();

      updateToolbar();
      editor.dispatchEvent(new Event('input', { bubbles: true }));
      return;
    }

    // Con texto seleccionado: se aplica el tamaño exacto y se conserva la selección.
    // execCommand de tamaño sólo admite 1-7: se usa 7 y se sustituye por pt exactos.
    document.execCommand('fontSize', false, '7');

    const spans = [];
    editor.querySelectorAll('font[size="7"]').forEach(fontEl => {
      const span = document.createElement('span');
      span.style.fontSize = pt + 'pt';
      while (fontEl.firstChild) span.appendChild(fontEl.firstChild);
      fontEl.replaceWith(span);
      spans.push(span);
    });

    if (spans.length) {
      const newRange = document.createRange();
      newRange.setStartBefore(spans[0]);
      newRange.setEndAfter(spans[spans.length - 1]);
      sel.removeAllRanges();
      sel.addRange(newRange);
      savedRange = newRange.cloneRange();
    }

    updateToolbar();
  }

  /* Devuelve el elemento donde está el cursor/selección */
  function selectedElement() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;

    const range = sel.getRangeAt(0);
    let node = range.startContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;

    // Si la selección empieza en el propio editor, se usa el primer texto alcanzado.
    if (node === editor) {
      const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, null);
      let textNode;
      while ((textNode = walker.nextNode())) {
        if (range.intersectsNode(textNode)) {
          node = textNode.parentElement;
          break;
        }
      }
    }

    if (!node || !editor.contains(node)) return null;
    return node;
  }

  function getFontFamily() {
    const el = selectedElement();
    return el ? getComputedStyle(el).fontFamily : null;
  }

  function getFontSize() {
    let el = selectedElement();
    if (!el) return null;

    // Subíndice/superíndice reducen el tamaño visual: se muestra el del contenedor.
    const subSup = el.closest('sub, sup');
    if (subSup && subSup.parentElement) el = subSup.parentElement;

    const px = parseFloat(getComputedStyle(el).fontSize);
    if (!px) return null;
    return Math.round(px * 72 / 96); // px -> pt
  }

  /* =========================================================
     Interlineado
     ========================================================= */
  function getLineSpacing() {
    let el = selectedElement() || editor;

    const subSup = el.closest && el.closest('sub, sup');
    if (subSup && subSup.parentElement) el = subSup.parentElement;

    const cs = getComputedStyle(el);
    const lineHeight = parseFloat(cs.lineHeight);
    const fontSize = parseFloat(cs.fontSize);
    if (!lineHeight || !fontSize) return 1.5;

    const ratio = lineHeight / fontSize;
    return [1, 1.5, 2].reduce(
      (best, value) => Math.abs(value - ratio) < Math.abs(best - ratio) ? value : best
    );
  }

  function aplicarInterlineado(value) {
    if (!value) return;
    restoreSelection();

    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) {
      editor.style.lineHeight = value;
      updateToolbar();
      editor.dispatchEvent(new Event('input', { bubbles: true }));
      return;
    }

    const range = sel.getRangeAt(0);

    // Aplica a los bloques (párrafos/líneas) que toca la selección.
    const blocks = Array.from(editor.children).filter(child => range.intersectsNode(child));
    blocks.forEach(block => { block.style.lineHeight = value; });

    // Texto que está directamente dentro del editor, sin bloque propio.
    const directText = Array.from(editor.childNodes).some(node =>
      node.nodeType === Node.TEXT_NODE && range.intersectsNode(node)
    );

    if (directText || blocks.length === 0) {
      editor.style.lineHeight = value;
    }

    updateToolbar();
    editor.dispatchEvent(new Event('input', { bubbles: true }));
  }

  /* =========================================================
     Actualizar estado visual de la barra
     ========================================================= */
  function setActive(id, active) {
    const el = document.getElementById(id);
    if (el) el.classList.toggle('active', !!active);
  }

  function updateToolbar() {
    try {
      setActive('btnBold', document.queryCommandState('bold'));
      setActive('btnItalic', document.queryCommandState('italic'));
      setActive('btnUnderline', document.queryCommandState('underline'));
      setActive('btnSub', document.queryCommandState('subscript'));
      setActive('btnSup', document.queryCommandState('superscript'));

      const align =
        document.queryCommandState('justifyCenter') ? 'center' :
        document.queryCommandState('justifyRight')  ? 'right'  :
        document.queryCommandState('justifyFull')   ? 'justify' : 'left';

      setActive('btnLeft', align === 'left');
      setActive('btnCenter', align === 'center');
      setActive('btnRight', align === 'right');
      setActive('btnJustify', align === 'justify');
    } catch (e) { /* ignorar */ }

    // Fuente
    const familySel = document.getElementById('fontFamily');
    const family = getFontFamily();
    if (family) {
      const match = Array.from(familySel.options).find(
        o => o.value && family.toLowerCase().includes(o.value.toLowerCase())
      );
      familySel.value = match ? match.value : '';
    } else {
      familySel.value = '';
    }

    // Tamaño
    const sizeSel = document.getElementById('fontSize');
    const size = getFontSize();
    if (size) {
      const option = Array.from(sizeSel.options).find(o => o.value === String(size));
      sizeSel.value = option ? option.value : '';
    } else {
      sizeSel.value = '';
    }

    // Interlineado
    const spacing = String(getLineSpacing());
    const spacingBtn = document.getElementById('spacingMenuBtn');
    if (spacingBtn) spacingBtn.title = 'Interlineado ' + spacing.replace('.', ',');
    document.querySelectorAll('#spacingMenu [data-spacing]').forEach(item => {
      item.setAttribute('aria-checked', item.dataset.spacing === spacing ? 'true' : 'false');
    });
  }

  /* =========================================================
     Título / estado
     ========================================================= */
  function updateTitle() {
    const name = currentName || 'Sin título';
    document.title = (dirty ? '● ' : '') + name + ' — Koto';
  }

  function horaCorta(ts = Date.now()) {
    try {
      return new Date(ts).toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });
    } catch (e) {
      return '';
    }
  }

  /* =========================================================
     Copia local de recuperación
     ========================================================= */
  function programarAutoguardado() {
    clearTimeout(temporizadorAutoguardado);
    temporizadorAutoguardado = setTimeout(() => {
      if (dirty) guardarCopiaLocal();
    }, AUTOSAVE_ESPERA);
  }

  function guardarCopiaLocal() {
    if (!dirty) return;

    const html = editor.innerHTML;
    if (!html && !currentName) return;

    try {
      localStorage.setItem(AUTOSAVE_KEY, JSON.stringify({
        v: 1,
        html,
        name: currentName,
        format: currentFormat,
        at: Date.now()
      }));
      autosaveStatus.textContent = 'Copia local ' + horaCorta();
      autosaveStatus.title = 'Copia de recuperación guardada en este navegador';
    } catch (e) {
      autosaveStatus.textContent = 'No se pudo guardar la copia local';
      autosaveStatus.title = 'El navegador no permitió guardar la copia de recuperación.';
    }
  }

  function limpiarCopiaLocal(mostrarGuardado = true) {
    try { localStorage.removeItem(AUTOSAVE_KEY); } catch (e) { /* ignorar */ }
    copiaPendiente = null;

    if (mostrarGuardado) {
      autosaveStatus.textContent = 'Guardado ' + horaCorta();
      autosaveStatus.title = 'Documento guardado';
    } else {
      autosaveStatus.textContent = '';
      autosaveStatus.title = '';
    }
  }

  function marcarComoGuardado(version = versionDocumento) {
    // Si se escribió mientras se generaba el archivo, todavía hay cambios
    // pendientes: no se debe ocultar el punto del título ni borrar la copia.
    if (version !== versionDocumento) {
      updateTitle();
      programarAutoguardado();
      return;
    }

    dirty = false;
    updateTitle();
    limpiarCopiaLocal(true);
  }

  function ofrecerRecuperacion() {
    let data = null;

    try {
      const raw = localStorage.getItem(AUTOSAVE_KEY);
      if (!raw) return;
      data = JSON.parse(raw);
    } catch (e) {
      try { localStorage.removeItem(AUTOSAVE_KEY); } catch (e2) { /* ignorar */ }
      return;
    }

    if (!data || data.v !== 1 || typeof data.html !== 'string' || (!data.html && !data.name)) {
      try { localStorage.removeItem(AUTOSAVE_KEY); } catch (e) { /* ignorar */ }
      return;
    }

    copiaPendiente = data;
    const fecha = new Date(Number(data.at) || Date.now());
    const detalle = data.name ? ' de «' + data.name + '»' : '';

    document.getElementById('recoveryText').textContent =
      'Hay una copia sin guardar' + detalle + ' del ' +
      fecha.toLocaleString('es-AR', { dateStyle: 'short', timeStyle: 'short', hour12: false }) +
      '. ¿Quieres recuperarla?';
    document.getElementById('recoveryDialog').hidden = false;
    document.getElementById('recoveryAccept').focus();
  }

  function recuperarDocumento() {
    if (!copiaPendiente) return;

    cerrarBuscar(false);
    editor.innerHTML = copiaPendiente.html;
    currentName = copiaPendiente.name || null;
    currentFormat = copiaPendiente.format || null;
    currentHandle = null;
    versionDocumento++;
    dirty = true;
    updateTitle();

    autosaveStatus.textContent = 'Copia recuperada ' + horaCorta(copiaPendiente.at);
    autosaveStatus.title = 'Documento recuperado de la copia local';

    document.getElementById('recoveryDialog').hidden = true;
    copiaPendiente = null;

    editor.focus();
    const range = document.createRange();
    range.selectNodeContents(editor);
    range.collapse(false);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    savedRange = range.cloneRange();

    programarAutoguardado();
  }

  function descartarRecuperacion() {
    document.getElementById('recoveryDialog').hidden = true;
    copiaPendiente = null;
    limpiarCopiaLocal(false);
    editor.focus();
  }

  editor.addEventListener('input', () => {
    versionDocumento++;
    dirty = true;
    updateTitle();
    autosaveStatus.textContent = 'Cambios sin guardar';
    autosaveStatus.title = 'Se guardará una copia local en unos segundos';
    programarAutoguardado();

    // Los resultados de búsqueda dejan de ser válidos al cambiar el texto.
    if (!findBar.hidden) invalidarBusqueda();
  });

  /* =========================================================
     Tabulador: sangría (en lugar de cambiar de foco)
     ========================================================= */
  editor.addEventListener('keydown', e => {
    // Ctrl+Shift+V: se marca la intención y el evento "paste" inserta texto plano.
    // No se llama a preventDefault aquí para que el navegador emita "paste".
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'v') {
      pegarSinFormato = true;
      setTimeout(() => { pegarSinFormato = false; }, 1200);
      return;
    }

    if (e.key !== 'Tab') return;
    e.preventDefault();

    if (e.shiftKey) {
      desindentar();
    } else {
      document.execCommand('insertText', false, '\t');
    }
  });

  editor.addEventListener('paste', e => {
    if (!pegarSinFormato) return;

    e.preventDefault();
    pegarSinFormato = false;

    const texto = e.clipboardData ? e.clipboardData.getData('text/plain') : '';
    if (!texto) return;

    restoreSelection();
    document.execCommand('insertText', false, texto.replace(/\r\n?/g, '\n'));
  });

  function desindentar() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;

    const range = sel.getRangeAt(0);
    if (!range.collapsed) return;

    const node = range.startContainer;
    if (node.nodeType !== Node.TEXT_NODE) return;

    const offset = range.startOffset;
    const text = node.textContent;

    let remove = 0;
    if (text[offset - 1] === '\t') {
      remove = 1;
    } else {
      while (remove < 4 && text[offset - 1 - remove] === ' ') remove++;
    }
    if (!remove) return;

    node.textContent = text.slice(0, offset - remove) + text.slice(offset);

    const newRange = document.createRange();
    newRange.setStart(node, offset - remove);
    newRange.collapse(true);
    sel.removeAllRanges();
    sel.addRange(newRange);
    savedRange = newRange.cloneRange();

    editor.dispatchEvent(new Event('input', { bubbles: true }));
  }

  /* =========================================================
     Enlaces
     ========================================================= */
  function normalizarUrl(valor) {
    const url = (valor || '').trim();
    if (!url) return null;
    if (/^javascript:/i.test(url)) return null; // se evita ejecutar código
    if (/^(https?|mailto|tel|ftp):/i.test(url)) return url;
    if (url.startsWith('//')) return 'https:' + url;
    return 'https://' + url;
  }

  function obtenerEnlaceSeleccion() {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;

    let node = sel.getRangeAt(0).startContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    if (!node || !node.closest) return null;

    const enlace = node.closest('a');
    return enlace && editor.contains(enlace) ? enlace : null;
  }

  function insertarEnlace() {
    restoreSelection();

    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;

    rangoEnlace = sel.getRangeAt(0).cloneRange();
    enlaceEnEdicion = obtenerEnlaceSeleccion();

    const nombreInput = document.getElementById('linkName');
    const urlInput = document.getElementById('linkUrl');
    const titulo = document.getElementById('linkDialogTitle');
    const aceptar = document.getElementById('linkAccept');

    if (enlaceEnEdicion) {
      nombreInput.value = enlaceEnEdicion.textContent;
      urlInput.value = enlaceEnEdicion.getAttribute('href') || '';
      titulo.textContent = 'Editar enlace';
      aceptar.textContent = 'Actualizar';
    } else {
      nombreInput.value = sel.toString();
      urlInput.value = '';
      titulo.textContent = 'Insertar enlace';
      aceptar.textContent = 'Insertar';
    }

    document.getElementById('linkDialog').hidden = false;

    if (nombreInput.value) {
      urlInput.focus();
    } else {
      nombreInput.focus();
    }
    if (nombreInput.value) nombreInput.select();
    else urlInput.select();
  }

  function aceptarEnlace() {
    const nombre = document.getElementById('linkName').value.trim();
    const url = normalizarUrl(document.getElementById('linkUrl').value);

    if (!nombre) {
      alert('Ingresa el nombre del enlace.');
      return;
    }
    if (!url) {
      alert('Ingresa un enlace válido, por ejemplo: https://ejemplo.com');
      return;
    }

    document.getElementById('linkDialog').hidden = true;

    if (enlaceEnEdicion) {
      // Editar un enlace existente.
      if (enlaceEnEdicion.textContent !== nombre) {
        enlaceEnEdicion.textContent = nombre;
      }
      enlaceEnEdicion.href = url;

      const range = document.createRange();
      range.setStartAfter(enlaceEnEdicion);
      range.collapse(true);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      savedRange = range.cloneRange();

    } else {
      // Insertar un enlace nuevo con el formato heredado del punto de inserción.
      const sel = window.getSelection();
      sel.removeAllRanges();
      if (rangoEnlace) sel.addRange(rangoEnlace);

      const range = sel.getRangeAt(0);
      const a = document.createElement('a');
      a.href = url;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';

      if (!range.collapsed && range.toString() === nombre) {
        // Se conserva el formato interno del texto seleccionado.
        a.appendChild(range.extractContents());
      } else {
        range.deleteContents();
        a.textContent = nombre;
      }

      range.insertNode(a);

      const despues = document.createRange();
      despues.setStartAfter(a);
      despues.collapse(true);
      sel.removeAllRanges();
      sel.addRange(despues);
      savedRange = despues.cloneRange();
    }

    enlaceEnEdicion = null;
    rangoEnlace = null;

    editor.focus();
    versionDocumento++;
    dirty = true;
    updateTitle();
    updateToolbar();
    autosaveStatus.textContent = 'Cambios sin guardar';
    autosaveStatus.title = 'Se guardará una copia local en unos segundos';
    programarAutoguardado();
    if (!findBar.hidden) invalidarBusqueda();
  }

  function cancelarEnlace() {
    document.getElementById('linkDialog').hidden = true;
    enlaceEnEdicion = null;
    rangoEnlace = null;
    editor.focus();
  }

  /* En el editor, un clic normal no navega; Ctrl/Cmd + clic abre el enlace. */
  editor.addEventListener('click', e => {
    const enlace = e.target.closest && e.target.closest('a');
    if (!enlace || !editor.contains(enlace)) return;

    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      window.open(enlace.href, '_blank', 'noopener');
    } else {
      e.preventDefault();
    }
  });

  /* =========================================================
     Buscar y reemplazar
     ========================================================= */
  function soportaResaltado() {
    return !!(window.CSS && CSS.highlights && typeof Highlight === 'function');
  }

  function construirMapaTexto() {
    const nodos = [];
    let texto = '';

    function agregarTexto(node) {
      const valor = node.nodeValue || '';
      if (!valor) return;
      nodos.push({
        node,
        start: texto.length,
        end: texto.length + valor.length
      });
      texto += valor;
    }

    function caminar(node) {
      for (const hijo of Array.from(node.childNodes)) {
        if (hijo.nodeType === Node.TEXT_NODE) {
          agregarTexto(hijo);
          continue;
        }

        if (hijo.nodeType !== Node.ELEMENT_NODE) continue;
        if (hijo.tagName === 'SCRIPT' || hijo.tagName === 'STYLE' || hijo.tagName === 'NOSCRIPT') continue;

        if (hijo.tagName === 'BR') {
          texto += '\n';
          continue;
        }

        const esBloque = BLOQUES_TEXTO.has(hijo.tagName);
        if (esBloque) texto += '\n';
        caminar(hijo);
        if (esBloque) texto += '\n';
      }
    }

    caminar(editor);
    return { texto, nodos };
  }

  function posicionEnMapa(offset) {
    if (!mapaBusqueda || !mapaBusqueda.nodos.length) return null;

    let elegido = mapaBusqueda.nodos[0];
    for (const entrada of mapaBusqueda.nodos) {
      if (entrada.start <= offset) elegido = entrada;
      else break;
    }

    return {
      node: elegido.node,
      offset: Math.max(0, Math.min(offset - elegido.start, elegido.node.nodeValue.length))
    };
  }

  function rangoDeHallazgo(hallazgo) {
    if (!hallazgo) return null;

    const inicio = posicionEnMapa(hallazgo.start);
    const fin = posicionEnMapa(hallazgo.end);
    if (!inicio || !fin) return null;

    try {
      const range = document.createRange();
      range.setStart(inicio.node, inicio.offset);
      range.setEnd(fin.node, fin.offset);
      return range;
    } catch (e) {
      return null;
    }
  }

  function limpiarResaltados() {
    if (!window.CSS || !CSS.highlights) return;
    try {
      CSS.highlights.delete('buscar');
      CSS.highlights.delete('buscar-actual');
    } catch (e) { /* ignorar */ }
  }

  function resaltarHallazgos() {
    if (!soportaResaltado()) return;

    try {
      const ranges = hallazgos
        .slice(0, 1000)
        .map(rangoDeHallazgo)
        .filter(Boolean);

      if (!ranges.length) {
        CSS.highlights.delete('buscar');
        return;
      }

      CSS.highlights.set('buscar', new Highlight(...ranges));
    } catch (e) { /* ignorar */ }
  }

  function resaltarActual() {
    if (!soportaResaltado()) return;

    try {
      if (hallazgoActual >= 0 && hallazgos[hallazgoActual]) {
        const range = rangoDeHallazgo(hallazgos[hallazgoActual]);
        if (range) CSS.highlights.set('buscar-actual', new Highlight(range));
        else CSS.highlights.delete('buscar-actual');
      } else {
        CSS.highlights.delete('buscar-actual');
      }
    } catch (e) { /* ignorar */ }
  }

  function actualizarContadorBusqueda() {
    const sinResultados = !!findText.value && !hallazgos.length;
    findCount.classList.toggle('no-results', sinResultados);

    if (!findText.value) {
      findCount.textContent = '';
    } else if (!hallazgos.length) {
      findCount.textContent = 'Sin resultados';
    } else {
      const total = hallazgos.length + (hallazgos.length === 5000 ? '+' : '');
      findCount.textContent = (hallazgoActual >= 0 ? hallazgoActual + 1 : 0) + ' de ' + total;
    }
  }

  function calcularHallazgos() {
    const consulta = findText.value;

    limpiarResaltados();
    hallazgos = [];
    hallazgoActual = -1;
    mapaBusqueda = null;
    textoBuscado = consulta;

    if (!consulta) {
      actualizarContadorBusqueda();
      return;
    }

    mapaBusqueda = construirMapaTexto();
    const texto = mapaBusqueda.texto.toLocaleLowerCase('es');
    const buscado = consulta.toLocaleLowerCase('es');

    if (!buscado) {
      actualizarContadorBusqueda();
      return;
    }

    let pos = 0;
    while (hallazgos.length < 5000) {
      const encontrado = texto.indexOf(buscado, pos);
      if (encontrado === -1) break;

      hallazgos.push({ start: encontrado, end: encontrado + buscado.length });
      pos = encontrado + buscado.length;
    }

    resaltarHallazgos();
    actualizarContadorBusqueda();
  }

  function seleccionarHallazgo(indice) {
    if (!hallazgos.length) {
      hallazgoActual = -1;
      actualizarContadorBusqueda();
      return;
    }

    hallazgoActual = ((indice % hallazgos.length) + hallazgos.length) % hallazgos.length;
    const range = rangoDeHallazgo(hallazgos[hallazgoActual]);
    if (!range) {
      actualizarContadorBusqueda();
      return;
    }

    // La selección real se usa sobre todo como respaldo en navegadores sin
    // la API de resaltados; el foco permanece en el buscador.
    if (!soportaResaltado()) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      savedRange = range.cloneRange();
    }

    const rect = range.getBoundingClientRect();
    const zona = workspaceEl.getBoundingClientRect();
    if (rect.top < zona.top + 16 || rect.bottom > zona.bottom - 16) {
      const el = range.startContainer.nodeType === Node.ELEMENT_NODE
        ? range.startContainer
        : range.startContainer.parentElement;
      if (el) el.scrollIntoView({ block: 'center' });
    }

    resaltarActual();
    actualizarContadorBusqueda();
  }

  function abrirBuscar(enfocarReemplazo = false) {
    const recovery = document.getElementById('recoveryDialog');
    if (recovery && !recovery.hidden) return;

    findBar.hidden = false;

    const sel = window.getSelection();
    if (sel && sel.rangeCount && !sel.isCollapsed && editor.contains(sel.anchorNode)) {
      const texto = sel.toString().trim();
      if (texto && texto.length <= 200 && !/[\r\n]/.test(texto)) findText.value = texto;
    }

    calcularHallazgos();
    if (hallazgos.length) seleccionarHallazgo(0);

    if (enfocarReemplazo && findText.value) {
      replaceText.focus();
      replaceText.select();
    } else {
      findText.focus();
      findText.select();
    }
  }

  function cerrarBuscar(devolverFoco = true) {
    const rangeActual = hallazgoActual >= 0 && hallazgos[hallazgoActual]
      ? rangoDeHallazgo(hallazgos[hallazgoActual])
      : null;

    findBar.hidden = true;
    limpiarResaltados();
    hallazgos = [];
    hallazgoActual = -1;
    mapaBusqueda = null;
    textoBuscado = '';
    actualizarContadorBusqueda();

    if (!devolverFoco) return;

    editor.focus();
    if (rangeActual) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(rangeActual);
      savedRange = rangeActual.cloneRange();
    } else {
      restoreSelection();
    }
  }

  function buscarSiguiente() {
    if (findText.value !== textoBuscado || !mapaBusqueda) calcularHallazgos();
    if (!hallazgos.length) {
      actualizarContadorBusqueda();
      return;
    }
    seleccionarHallazgo(hallazgoActual < 0 ? 0 : hallazgoActual + 1);
  }

  function buscarAnterior() {
    if (findText.value !== textoBuscado || !mapaBusqueda) calcularHallazgos();
    if (!hallazgos.length) {
      actualizarContadorBusqueda();
      return;
    }
    seleccionarHallazgo(hallazgoActual < 0 ? hallazgos.length - 1 : hallazgoActual - 1);
  }

  function limpiarElementosVacios() {
    const candidatos = editor.querySelectorAll('b, i, u, s, sub, sup, span, font, a');

    candidatos.forEach(el => {
      if (el.querySelector('br, img, table, div, p')) return;

      const texto = el.textContent || '';
      if (texto.replace(/\u200B/g, '').trim()) return;
      if (texto.includes('\u200B')) return; // se conserva el marcador de formato
      el.remove();
    });
  }

  function insertarTextoEnRango(range, texto, moverCursor = true) {
    range.deleteContents();

    const limpio = String(texto).replace(/\r\n?/g, '\n');
    if (!limpio) {
      if (moverCursor) {
        range.collapse(true);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        savedRange = range.cloneRange();
      }
      return;
    }

    const fragmento = document.createDocumentFragment();
    limpio.split('\n').forEach((linea, i) => {
      if (i) fragmento.appendChild(document.createElement('br'));
      if (linea) fragmento.appendChild(document.createTextNode(linea));
    });

    const ultimo = fragmento.lastChild;
    range.insertNode(fragmento);

    if (moverCursor && ultimo) {
      const despues = document.createRange();
      despues.setStartAfter(ultimo);
      despues.collapse(true);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(despues);
      savedRange = despues.cloneRange();
    }
  }

  function reemplazarActual() {
    if (findText.value !== textoBuscado || !mapaBusqueda) calcularHallazgos();
    if (!hallazgos.length) {
      actualizarContadorBusqueda();
      return;
    }

    if (hallazgoActual < 0) seleccionarHallazgo(0);
    const hallazgo = hallazgos[hallazgoActual];
    if (!hallazgo) return;

    const range = rangoDeHallazgo(hallazgo);
    if (!range) return;

    const reemplazo = replaceText.value;
    const nuevaPos = hallazgo.start + reemplazo.replace(/\r\n?/g, '\n').length;

    insertarTextoEnRango(range, reemplazo);
    limpiarElementosVacios();
    editor.dispatchEvent(new Event('input', { bubbles: true }));

    calcularHallazgos();
    if (!hallazgos.length) {
      hallazgoActual = -1;
      actualizarContadorBusqueda();
      return;
    }

    let siguiente = hallazgos.findIndex(item => item.start >= nuevaPos);
    if (siguiente === -1) siguiente = 0;
    seleccionarHallazgo(siguiente);
  }

  function reemplazarTodo() {
    if (findText.value !== textoBuscado || !mapaBusqueda) calcularHallazgos();
    if (!hallazgos.length) {
      actualizarContadorBusqueda();
      return;
    }

    const reemplazo = replaceText.value;
    const total = hallazgos.length;

    // De atrás hacia adelante para que los desplazamientos no invaliden los rangos.
    for (let i = hallazgos.length - 1; i >= 0; i--) {
      const range = rangoDeHallazgo(hallazgos[i]);
      if (range) insertarTextoEnRango(range, reemplazo, false);
    }

    limpiarElementosVacios();
    editor.dispatchEvent(new Event('input', { bubbles: true }));
    calcularHallazgos();

    if (hallazgos.length) {
      seleccionarHallazgo(0);
    } else {
      findCount.textContent = total + (total === 1 ? ' reemplazo' : ' reemplazos');
      findCount.classList.remove('no-results');
    }
  }

  function invalidarBusqueda() {
    limpiarResaltados();
    hallazgos = [];
    hallazgoActual = -1;
    mapaBusqueda = null;
    textoBuscado = '';
    actualizarContadorBusqueda();
  }

  findText.addEventListener('input', () => {
    clearTimeout(temporizadorBusqueda);
    temporizadorBusqueda = setTimeout(() => {
      calcularHallazgos();
      if (hallazgos.length) seleccionarHallazgo(0);
    }, 120);
  });

  findText.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (e.shiftKey) buscarAnterior();
      else buscarSiguiente();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      cerrarBuscar();
    }
  });

  replaceText.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      reemplazarActual();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      cerrarBuscar();
    }
  });

  findBar.querySelectorAll('button').forEach(btn => {
    btn.addEventListener('mousedown', e => e.preventDefault());
  });

  /* =========================================================
     Modo enfoque
     ========================================================= */
  function cargarPreferenciaEnfoque() {
    let activo = false;

    try {
      activo = localStorage.getItem(FOCUS_KEY) === '1';
    } catch (e) {
      activo = false;
    }

    const checkbox = document.getElementById('focusDefault');
    if (checkbox) checkbox.checked = activo;
    return activo;
  }

  function cambiarPreferenciaEnfoque(activo) {
    try {
      if (activo) localStorage.setItem(FOCUS_KEY, '1');
      else localStorage.removeItem(FOCUS_KEY);
    } catch (e) { /* ignorar */ }
  }

  function alternarEnfoque() {
    if (modoEnfoque) salirEnfoque();
    else entrarEnfoque();
  }

  function entrarEnfoque() {
    if (modoEnfoque) return;

    const recovery = document.getElementById('recoveryDialog');
    if (recovery && !recovery.hidden) return;

    modoEnfoque = true;
    document.body.classList.add('modo-enfoque');
    cerrarMenusDesplegables();
    cerrarBuscar(false);
    editor.focus();
  }

  function salirEnfoque() {
    if (!modoEnfoque) return;

    modoEnfoque = false;
    document.body.classList.remove('modo-enfoque');
    editor.focus();
  }

  /* =========================================================
     Zoom
     ========================================================= */
  function setZoom(percent, persistir = true) {
    ajusteAutomatico = false;

    let value = Math.round(Number(percent));
    if (!Number.isFinite(value)) value = 100;
    value = Math.max(50, Math.min(400, value));

    zoomActual = value;
    pageEl.style.zoom = String(value / 100);
    zoomRange.value = String(value);
    zoomLabel.textContent = value + '%';

    if (persistir) {
      try { localStorage.setItem('procesadorZoom', String(value)); } catch (e) { /* ignorar */ }
    }
  }

  function cambiarZoom(delta) {
    setZoom(zoomActual + delta);
  }

  function ajustarAncho() {
    // Se mide la hoja sin zoom y se calcula el porcentaje que ocupa el ancho disponible.
    pageEl.style.zoom = '1';
    const base = pageEl.getBoundingClientRect().width;
    if (!base) return;

    const disponible = workspaceEl.clientWidth - 24; // padding lateral del área de trabajo
    const percent = Math.round((disponible / base) * 100);

    setZoom(percent);
    ajusteAutomatico = true;
  }

  window.addEventListener('resize', () => {
    if (ajusteAutomatico) ajustarAncho();
  });

  /* =========================================================
     Archivo: nuevo
     ========================================================= */
  function nuevoArchivo() {
    if (dirty && !confirm('Hay cambios sin guardar. ¿Descartarlos y crear un archivo nuevo?')) {
      return;
    }
    cerrarBuscar(false);
    limpiarCopiaLocal(false);
    editor.innerHTML = '';
    versionDocumento++;
    currentName = null;
    currentHandle = null;
    currentFormat = null;
    dirty = false;
    updateTitle();
    editor.focus();
  }

  /* =========================================================
     Archivo: abrir
     ========================================================= */
  function pickFileWithInput(accept) {
    return new Promise(resolve => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = accept;
      input.onchange = () => resolve(input.files[0] || null);
      input.click();
    });
  }

  async function abrirArchivo() {
    if (dirty && !confirm('Hay cambios sin guardar. ¿Descartarlos y abrir otro archivo?')) {
      return;
    }

    if (!window.Formatos) {
      alert('No se pudieron cargar los conversores de formato.');
      return;
    }

    try {
      let file = null;
      let handle = null;

      if (window.showOpenFilePicker) {
        const handles = await window.showOpenFilePicker({
          types: [{
            description: 'Documentos',
            accept: {
              'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
              'application/vnd.oasis.opendocument.text': ['.odt'],
              'application/msword': ['.doc'],
              'text/html': ['.html', '.htm'],
              'text/plain': ['.txt', '.md']
            }
          }],
          multiple: false
        });
        handle = handles[0];
        file = await handle.getFile();
      } else {
        file = await pickFileWithInput('.docx,.doc,.odt,.html,.htm,.txt,.md');
        if (!file) return;
      }

      const ext = (file.name.split('.').pop() || '').toLowerCase();
      const buffer = await file.arrayBuffer();
      const contenido = await window.Formatos.abrir(buffer, ext);
      if (contenido === null) return; // formato no soportado (ya se avisó)

      editor.innerHTML = contenido;
      versionDocumento++;

      currentName = ext === 'doc' ? file.name.replace(/\.doc$/i, '.docx') : file.name;
      currentHandle = handle;
      currentFormat = ext === 'doc' ? 'docx' : window.Formatos.formatoDeNombre(file.name);
      dirty = false;
      updateTitle();
      cerrarBuscar(false);
      limpiarCopiaLocal(false);
      editor.focus();
    } catch (err) {
      if (err && err.name !== 'AbortError') {
        alert('No se pudo abrir el archivo:\n' + err.message);
      }
    }
  }

  /* =========================================================
     Archivo: guardar / guardar como
     ========================================================= */
  async function generarContenido(formato) {
    if (formato === 'docx') return await window.Formatos.generarDocx(editor);
    if (formato === 'odt') return await window.Formatos.generarOdt(editor);
    if (formato === 'txt') return editor.innerText;
    return editor.innerHTML;
  }

  function descargar(contenido, nombre, formato) {
    const mime = window.Formatos.mimeDe(formato);
    const blob = contenido instanceof Blob
      ? contenido
      : new Blob([contenido], { type: mime + ';charset=utf-8' });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = nombre;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  async function escribirEnHandle(handle, contenido) {
    const writable = await handle.createWritable();
    await writable.write(contenido);
    await writable.close();
  }

  async function guardarArchivo() {
    // 1) Guardar sobre el archivo abierto, respetando su formato.
    if (currentHandle) {
      const formato = currentFormat || window.Formatos.formatoDeNombre(currentName) || 'odt';
      try {
        const version = versionDocumento;
        const contenido = await generarContenido(formato);
        await escribirEnHandle(currentHandle, contenido);
        marcarComoGuardado(version);
        return;
      } catch (err) {
        // Si falla, se ofrece "Guardar como".
      }
    }

    // 2) Si ya hay nombre (por ejemplo, abierto en Firefox), se guarda en ese formato.
    if (currentName) {
      const formato = currentFormat || window.Formatos.formatoDeNombre(currentName) || 'odt';
      const version = versionDocumento;
      const contenido = await generarContenido(formato);
      descargar(contenido, currentName, formato);
      marcarComoGuardado(version);
      return;
    }

    // 3) Documento nuevo: por defecto ODT.
    await guardarNuevoOdt();
  }

  async function guardarNuevoOdt() {
    if (window.showSaveFilePicker) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: 'documento.odt',
          types: TIPOS_GUARDADO
        });
        const formato = window.Formatos.formatoDeNombre(handle.name) || 'odt';
        const version = versionDocumento;
        const contenido = await generarContenido(formato);
        await escribirEnHandle(handle, contenido);

        currentHandle = handle;
        currentName = handle.name;
        currentFormat = formato;
        marcarComoGuardado(version);
        return;
      } catch (err) {
        if (err && err.name === 'AbortError') return;
        // Si el selector nativo falla, se descarga directamente.
      }
    }

    const version = versionDocumento;
    const contenido = await generarContenido('odt');
    descargar(contenido, 'documento.odt', 'odt');

    currentHandle = null;
    currentName = 'documento.odt';
    currentFormat = 'odt';
    marcarComoGuardado(version);
  }

  async function guardarComo() {
    const formatoSugerido = currentFormat || 'odt';
    const nombreSugerido = currentName || ('documento.' + window.Formatos.extensionDe(formatoSugerido));

    if (window.showSaveFilePicker) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: nombreSugerido,
          types: TIPOS_GUARDADO
        });

        const formato = window.Formatos.formatoDeNombre(handle.name);
        if (!formato) {
          alert('Extensión no soportada. Usa .odt, .docx, .html o .txt.');
          return;
        }

        const version = versionDocumento;
        const contenido = await generarContenido(formato);
        await escribirEnHandle(handle, contenido);

        currentHandle = handle;
        currentName = handle.name;
        currentFormat = formato;
        marcarComoGuardado(version);
        return;
      } catch (err) {
        if (err && err.name === 'AbortError') return;
        // Si el selector nativo falla, se usa el diálogo propio.
      }
    }

    const eleccion = await abrirDialogoGuardar(nombreSugerido, formatoSugerido);
    if (!eleccion) return;

    const version = versionDocumento;
    const contenido = await generarContenido(eleccion.formato);
    descargar(contenido, eleccion.nombre, eleccion.formato);

    currentHandle = null;
    currentName = eleccion.nombre;
    currentFormat = eleccion.formato;
    marcarComoGuardado(version);
  }

  /* ---------- Diálogo "Guardar como" (navegadores sin selector nativo) ---------- */
  function abrirDialogoGuardar(nombreSugerido, formatoSugerido) {
    const overlay = document.getElementById('saveDialog');
    const input = document.getElementById('saveFileName');
    const select = document.getElementById('saveFileFormat');

    input.value = (nombreSugerido || 'documento').replace(/\.[^.]+$/, '');
    select.value = formatoSugerido || 'odt';

    overlay.hidden = false;
    input.focus();
    input.select();

    return new Promise(resolve => { resolverGuardar = resolve; });
  }

  function cerrarDialogoGuardar(resultado) {
    document.getElementById('saveDialog').hidden = true;
    if (resolverGuardar) {
      const resolver = resolverGuardar;
      resolverGuardar = null;
      resolver(resultado);
    }
  }

  function aceptarDialogoGuardar() {
    const input = document.getElementById('saveFileName');
    const select = document.getElementById('saveFileFormat');
    const formato = select.value;

    let base = input.value.trim() || 'documento';
    base = base.replace(/\.[^.]+$/, '') || 'documento';

    cerrarDialogoGuardar({
      nombre: base + '.' + window.Formatos.extensionDe(formato),
      formato
    });
  }

  /* =========================================================
     Archivo: cerrar documento
     ========================================================= */
  function cerrarDocumento() {
    if (dirty && !confirm('Hay cambios sin guardar. ¿Cerrar el documento de todos modos?')) {
      return;
    }
    cerrarBuscar(false);
    limpiarCopiaLocal(false);
    editor.innerHTML = '';
    versionDocumento++;
    currentName = null;
    currentHandle = null;
    currentFormat = null;
    dirty = false;
    updateTitle();
    editor.focus();
  }

  /* =========================================================
     Exportar a PDF
     ========================================================= */
  async function exportarPDF() {
    if (!window.html2canvas || !window.jspdf || !window.jspdf.jsPDF) {
      alert('No se pudo cargar el generador de PDF (parece que no hay conexión).\n' +
            'Se abrirá el diálogo de impresión: elige "Guardar como PDF".');
      imprimir();
      return;
    }

    const previousZoom = pageEl.style.zoom;

    try {
      // El PDF se genera siempre sin el zoom de pantalla.
      if (previousZoom) {
        pageEl.style.zoom = '1';
        await new Promise(resolve => requestAnimationFrame(resolve));
      }

      const canvas = await window.html2canvas(editor, {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: true,
        logging: false
      });

      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

      const pageW = pdf.internal.pageSize.getWidth();   // 210
      const pageH = pdf.internal.pageSize.getHeight();  // 297
      const margin = 20;
      const contentW = pageW - margin * 2;
      const imgH = canvas.height * contentW / canvas.width;
      const usableH = pageH - margin * 2;

      pdf.addImage(canvas, 'PNG', margin, margin, contentW, imgH);

      let remaining = imgH - usableH;
      let offset = 0;
      while (remaining > 0) {
        offset += usableH;
        pdf.addPage();
        pdf.addImage(canvas, 'PNG', margin, margin - offset, contentW, imgH);
        remaining -= usableH;
      }

      const base = (currentName || 'documento').replace(/\.[^.]+$/, '');
      pdf.save(base + '.pdf');
    } catch (err) {
      alert('No se pudo exportar a PDF:\n' + err.message);
    } finally {
      if (previousZoom) pageEl.style.zoom = previousZoom;
    }
  }

  /* =========================================================
     Imprimir
     ========================================================= */
  function imprimir() {
    editor.focus();
    window.print();
  }

  /* =========================================================
     Atajos de teclado
     ========================================================= */
  document.addEventListener('keydown', e => {
    // Diálogo de recuperación: no se permite descartar por accidente con Escape.
    if (!document.getElementById('recoveryDialog').hidden) {
      if (e.key === 'Escape') {
        e.preventDefault();
        document.getElementById('recoveryAccept').focus();
      }
      return;
    }

    // Diálogos abiertos
    if (!document.getElementById('saveDialog').hidden) {
      if (e.key === 'Escape') { e.preventDefault(); cerrarDialogoGuardar(null); return; }
      if (e.key === 'Enter' && e.target.id === 'saveFileName') {
        e.preventDefault();
        aceptarDialogoGuardar();
        return;
      }
      return;
    }

    if (!document.getElementById('linkDialog').hidden) {
      if (e.key === 'Escape') { e.preventDefault(); cancelarEnlace(); return; }
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.target.id === 'linkName') document.getElementById('linkUrl').focus();
        else aceptarEnlace();
        return;
      }
      return;
    }

    if (e.key === 'Escape') {
      const menuAbierto = menusDesplegables.find(menu => menu.estaAbierto());
      if (menuAbierto) {
        e.preventDefault();
        menuAbierto.cerrar();
        menuAbierto.btn.focus();
        return;
      }
      if (!findBar.hidden) { e.preventDefault(); cerrarBuscar(); return; }
      if (modoEnfoque) { e.preventDefault(); salirEnfoque(); return; }
    }

    const ctrl = e.ctrlKey || e.metaKey;
    const key = e.key.toLowerCase();

    // Buscar / reemplazar y modo enfoque funcionan también desde el buscador.
    if (ctrl && !e.shiftKey && (key === 'f' || key === 'h')) {
      e.preventDefault();
      abrirBuscar(key === 'h');
      return;
    }

    if (ctrl && e.shiftKey && key === 'f') {
      e.preventDefault();
      alternarEnfoque();
      return;
    }

    // No interferir con campos de formulario.
    if (e.target && e.target.matches && e.target.matches('input, select, textarea')) return;

    if (!ctrl) return;

    if (key === 'b' && !e.shiftKey) { e.preventDefault(); exec('bold'); }
    else if (key === 'i' && !e.shiftKey) { e.preventDefault(); exec('italic'); }
    else if (key === 'u' && !e.shiftKey) { e.preventDefault(); exec('underline'); }
    else if (key === 'k') { e.preventDefault(); insertarEnlace(); }
    else if (key === 's' && e.shiftKey) { e.preventDefault(); guardarComo(); }
    else if (key === 's') { e.preventDefault(); guardarArchivo(); }
    else if (key === 'o') { e.preventDefault(); abrirArchivo(); }
    else if (key === 'n') { e.preventDefault(); nuevoArchivo(); }
    else if (key === 'p') { e.preventDefault(); imprimir(); }
    else if (e.shiftKey && key === 'l') { e.preventDefault(); exec('justifyLeft'); }
    else if (e.shiftKey && key === 'e') { e.preventDefault(); exec('justifyCenter'); }
    else if (e.shiftKey && key === 'r') { e.preventDefault(); exec('justifyRight'); }
    else if (e.shiftKey && key === 'j') { e.preventDefault(); exec('justifyFull'); }
  });

  /* Cerrar diálogos al hacer clic fuera */
  const linkDialogEl = document.getElementById('linkDialog');
  linkDialogEl.addEventListener('mousedown', e => {
    if (e.target === linkDialogEl) cancelarEnlace();
  });

  const saveDialogEl = document.getElementById('saveDialog');
  saveDialogEl.addEventListener('mousedown', e => {
    if (e.target === saveDialogEl) cerrarDialogoGuardar(null);
  });

  /* Aviso al cerrar la pestaña con cambios sin guardar */
  window.addEventListener('beforeunload', e => {
    if (!dirty) return;
    guardarCopiaLocal();
    e.preventDefault();
    e.returnValue = '';
  });

  document.addEventListener('visibilitychange', () => {
    if (document.hidden && dirty) guardarCopiaLocal();
  });

  setInterval(() => {
    if (dirty) guardarCopiaLocal();
  }, AUTOSAVE_INTERVALO);

  /* =========================================================
     Inicio
     ========================================================= */
  // Se recuerda el último zoom utilizado.
  try {
    const savedZoom = parseInt(localStorage.getItem('procesadorZoom'), 10);
    setZoom(savedZoom >= 50 && savedZoom <= 400 ? savedZoom : 100, false);
  } catch (e) {
    setZoom(100, false);
  }

  const abrirEnEnfoque = cargarPreferenciaEnfoque();

  updateTitle();
  updateToolbar();
  editor.focus();
  if (abrirEnEnfoque) entrarEnfoque();
  ofrecerRecuperacion();
