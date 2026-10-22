'use strict';

/* =========================================================
   Conversión entre el editor HTML y los formatos DOCX / ODT.
   - DOCX: se lee con Mammoth y se escribe con JSZip.
   - ODT: se lee y se escribe con JSZip + XML propio.
   Todo funciona en el navegador y sin conexión.
   ========================================================= */

(function () {

  const MIME = {
    odt: 'application/vnd.oasis.opendocument.text',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    html: 'text/html',
    txt: 'text/plain'
  };

  const NS = {
    office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    xlink: 'http://www.w3.org/1999/xlink',
    table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'
  };

  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  /* ---------------------------------------------------------
     Utilidades
     --------------------------------------------------------- */
  function escaparXml(texto) {
    return String(texto)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function escaparAtributo(texto) {
    return escaparXml(texto).replace(/"/g, '&quot;');
  }

  function primeraFuente(valor) {
    if (!valor) return '';
    return valor.split(',')[0].replace(/^["']|["']$/g, '').trim();
  }

  function colorAHex(valor) {
    if (!valor) return null;

    const texto = String(valor).trim().toLowerCase();
    if (texto.startsWith('#')) {
      const hex = texto.slice(1);
      if (hex.length === 3) {
        return hex.split('').map(c => c + c).join('').toUpperCase();
      }
      return hex.length === 6 ? hex.toUpperCase() : null;
    }

    const rgb = texto.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
    if (!rgb) return null;

    return [rgb[1], rgb[2], rgb[3]]
      .map(n => Math.max(0, Math.min(255, parseInt(n, 10))).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  }

  function ptDesdeCss(valor) {
    const px = parseFloat(valor);
    if (isNaN(px) || px === 0) return 0;
    return Math.round(px * 72 / 96 * 10) / 10;
  }

  function formatoDeNombre(nombre) {
    const ext = (String(nombre).split('.').pop() || '').toLowerCase();
    if (ext === 'docx') return 'docx';
    if (ext === 'odt') return 'odt';
    if (ext === 'html' || ext === 'htm') return 'html';
    if (ext === 'txt' || ext === 'md') return 'txt';
    return null;
  }

  function extensionDe(formato) {
    return { odt: 'odt', docx: 'docx', html: 'html', txt: 'txt' }[formato] || 'odt';
  }

  function mimeDe(formato) {
    return MIME[formato] || MIME.odt;
  }

  /* =========================================================
     LECTURA
     ========================================================= */
  async function abrir(buffer, ext) {
    if (ext === 'docx') return await docxAHtml(buffer);

    if (ext === 'doc') {
      // Algunos archivos .doc son en realidad .docx renombrados.
      const bytes = new Uint8Array(buffer);
      const esZip = bytes.length > 3 && bytes[0] === 0x50 && bytes[1] === 0x4B;
      if (esZip) return await docxAHtml(buffer);

      alert('El formato .doc (Word 97-2003) es binario y no puede abrirse directamente en el navegador.\n\n' +
            'Abre el archivo en Word o LibreOffice y guárdalo como .docx o .odt.');
      return null;
    }

    if (ext === 'odt') return await odtAHtml(buffer);

    const texto = new TextDecoder('utf-8').decode(buffer);

    if (ext === 'txt' || ext === 'md') {
      const div = document.createElement('div');
      div.textContent = texto;
      return div.innerHTML;
    }

    // html / htm
    return texto;
  }

  async function docxAHtml(buffer) {
    // Se intenta primero el lector propio, que conserva fuente, tamaño,
    // alineación, interlineado y sangría. Mammoth queda como respaldo.
    try {
      return await docxXmlAHtml(buffer);
    } catch (err) {
      console.warn('No se pudo leer el DOCX con el lector propio; se usa Mammoth.', err);
    }

    if (!window.mammoth) throw new Error('No se pudo cargar el lector de DOCX.');
    const resultado = await window.mammoth.convertToHtml({ arrayBuffer: buffer });
    return resultado.value || '';
  }

  /* ---------- DOCX: lectura con formato ---------- */
  function hijosW(el, local) {
    return Array.from(el.children).filter(c => c.namespaceURI === NS_W && c.localName === local);
  }

  function primerW(el, local) {
    return hijosW(el, local)[0] || null;
  }

  function attrW(el, local) {
    return el ? el.getAttributeNS(NS_W, local) : null;
  }

  function onOffW(el) {
    if (!el) return false;
    const valor = attrW(el, 'val');
    if (valor == null) return true;
    return !/^(0|false|off|none)$/i.test(valor);
  }

  function leerPPrDocx(pPr) {
    const props = {};
    if (!pPr) return props;

    const jc = primerW(pPr, 'jc');
    if (jc) {
      const mapa = {
        left: 'left', start: 'left', center: 'center',
        right: 'right', end: 'right', both: 'justify',
        justify: 'justify', distribute: 'justify'
      };
      const alineacion = mapa[attrW(jc, 'val')];
      if (alineacion) props.align = alineacion;
    }

    const spacing = primerW(pPr, 'spacing');
    if (spacing) {
      const linea = parseFloat(attrW(spacing, 'line'));
      const regla = attrW(spacing, 'lineRule') || 'auto';
      if (!isNaN(linea)) {
        if (regla === 'auto') props.line = Math.round((linea / 240) * 100) / 100;
        else props.lineHeightCss = Math.round((linea / 20) * 10) / 10 + 'pt';
      }
    }

    const ind = primerW(pPr, 'ind');
    if (ind) {
      const izquierda = parseFloat(attrW(ind, 'left') ?? attrW(ind, 'start'));
      const derecha = parseFloat(attrW(ind, 'right') ?? attrW(ind, 'end'));
      const primera = parseFloat(attrW(ind, 'firstLine'));
      const francesa = parseFloat(attrW(ind, 'hanging'));

      if (!isNaN(izquierda)) props.marginLeft = (izquierda / 20) + 'pt';
      if (!isNaN(derecha)) props.marginRight = (derecha / 20) + 'pt';
      if (!isNaN(francesa)) props.textIndent = (-francesa / 20) + 'pt';
      else if (!isNaN(primera)) props.textIndent = (primera / 20) + 'pt';
    }

    return props;
  }

  function leerRPrDocx(rPr) {
    const props = {};
    if (!rPr) return props;

    const fuentes = primerW(rPr, 'rFonts');
    if (fuentes) {
      const fuente = attrW(fuentes, 'ascii') || attrW(fuentes, 'hAnsi') || attrW(fuentes, 'cs');
      if (fuente) props.fontFamily = fuente;
    }

    const sz = primerW(rPr, 'sz');
    if (sz) {
      const valor = parseFloat(attrW(sz, 'val'));
      if (!isNaN(valor)) props.fontSizePt = valor / 2;
    }

    const negrita = primerW(rPr, 'b');
    if (negrita) props.bold = onOffW(negrita);

    const cursiva = primerW(rPr, 'i');
    if (cursiva) props.italic = onOffW(cursiva);

    const subrayado = primerW(rPr, 'u');
    if (subrayado) {
      const valor = (attrW(subrayado, 'val') || 'single').toLowerCase();
      props.underline = valor !== 'none';
    }

    const color = primerW(rPr, 'color');
    if (color) {
      const valor = attrW(color, 'val');
      if (valor && valor.toLowerCase() !== 'auto') {
        props.color = '#' + valor.replace(/^#/, '').toLowerCase();
      }
    }

    const vertical = primerW(rPr, 'vertAlign');
    if (vertical) {
      const valor = attrW(vertical, 'val');
      if (valor === 'superscript' || valor === 'subscript') props.vertAlign = valor;
    }

    const tachado = primerW(rPr, 'strike');
    if (tachado) props.strike = onOffW(tachado);

    const mayusculas = primerW(rPr, 'caps');
    if (mayusculas) props.allCaps = onOffW(mayusculas);

    const estilo = primerW(rPr, 'rStyle');
    if (estilo) props.rStyle = attrW(estilo, 'val');

    return props;
  }

  function leerEstilosDocx(docEstilos) {
    const estilos = Object.create(null);
    const defaults = { pPr: {}, rPr: {} };
    let defaultPara = null;
    let defaultChar = null;

    if (!docEstilos) return { estilos, defaults, defaultPara, defaultChar };

    const docDefaults = docEstilos.getElementsByTagNameNS(NS_W, 'docDefaults')[0];
    if (docDefaults) {
      const rPrDefault = docDefaults.getElementsByTagNameNS(NS_W, 'rPrDefault')[0];
      if (rPrDefault) {
        defaults.rPr = leerRPrDocx(rPrDefault.getElementsByTagNameNS(NS_W, 'rPr')[0]);
      }
      const pPrDefault = docDefaults.getElementsByTagNameNS(NS_W, 'pPrDefault')[0];
      if (pPrDefault) {
        defaults.pPr = leerPPrDocx(pPrDefault.getElementsByTagNameNS(NS_W, 'pPr')[0]);
      }
    }

    const lista = docEstilos.getElementsByTagNameNS(NS_W, 'style');
    for (let i = 0; i < lista.length; i++) {
      const el = lista[i];
      const id = attrW(el, 'styleId');
      if (!id) continue;

      const tipo = attrW(el, 'type') || 'paragraph';
      const nombre = primerW(el, 'name');
      const basedOn = primerW(el, 'basedOn');

      estilos[id] = {
        tipo,
        name: nombre ? attrW(nombre, 'val') : null,
        basedOn: basedOn ? attrW(basedOn, 'val') : null,
        pPr: leerPPrDocx(primerW(el, 'pPr')),
        rPr: leerRPrDocx(primerW(el, 'rPr'))
      };

      if (attrW(el, 'default') === '1' || attrW(el, 'default') === 'true') {
        if (tipo === 'paragraph') defaultPara = id;
        else if (tipo === 'character') defaultChar = id;
      }
    }

    if (!defaultPara && estilos.Normal) defaultPara = 'Normal';

    return { estilos, defaults, defaultPara, defaultChar };
  }

  function resolverEstiloDocx(datos, id, cache, visitando) {
    if (!id || !datos.estilos[id]) return { pPr: {}, rPr: {} };
    if (cache[id]) return cache[id];
    if (visitando.has(id)) return { pPr: {}, rPr: {} };

    visitando.add(id);
    const estilo = datos.estilos[id];
    const base = estilo.basedOn
      ? resolverEstiloDocx(datos, estilo.basedOn, cache, visitando)
      : { pPr: {}, rPr: {} };

    const resultado = {
      pPr: Object.assign({}, base.pPr, estilo.pPr),
      rPr: Object.assign({}, base.rPr, estilo.rPr)
    };

    visitando.delete(id);
    cache[id] = resultado;
    return resultado;
  }

  async function leerRelacionesDocx(zip) {
    const relaciones = Object.create(null);
    const archivo = zip.file('word/_rels/document.xml.rels');
    if (!archivo) return relaciones;

    const doc = new DOMParser().parseFromString(await archivo.async('string'), 'application/xml');
    if (doc.querySelector('parsererror')) return relaciones;

    const lista = doc.getElementsByTagName('Relationship');
    for (let i = 0; i < lista.length; i++) {
      const el = lista[i];
      const id = el.getAttribute('Id');
      const target = el.getAttribute('Target');
      if (id && target) relaciones[id] = target;
    }
    return relaciones;
  }

  function crearRunDocx(run, propsBase, ctx) {
    const rPr = primerW(run, 'rPr');
    const directo = leerRPrDocx(rPr);

    let estiloChar = { pPr: {}, rPr: {} };
    if (directo.rStyle) {
      estiloChar = resolverEstiloDocx(ctx.datos, directo.rStyle, ctx.cache, new Set());
    }

    const props = Object.assign({}, propsBase, estiloChar.rPr, directo);

    const partes = [];
    for (const hijo of Array.from(run.childNodes)) {
      if (hijo.nodeType !== Node.ELEMENT_NODE || hijo.namespaceURI !== NS_W) continue;

      const local = hijo.localName;
      if (local === 't') {
        if (hijo.textContent) partes.push(document.createTextNode(hijo.textContent));
      } else if (local === 'tab') {
        partes.push(document.createTextNode('\t'));
      } else if (local === 'br' || local === 'cr') {
        partes.push(document.createElement('br'));
      } else if (local === 'noBreakHyphen') {
        partes.push(document.createTextNode('-'));
      } else if (local === 'softHyphen') {
        partes.push(document.createTextNode('\u00AD'));
      }
    }

    if (!partes.length) return null;

    let nodo = document.createDocumentFragment();
    partes.forEach(parte => nodo.appendChild(parte));

    const estilos = [];
    if (props.fontFamily && props.fontFamily !== propsBase.fontFamily) {
      estilos.push('font-family:' + props.fontFamily);
    }
    if (props.fontSizePt && props.fontSizePt !== propsBase.fontSizePt) {
      estilos.push('font-size:' + props.fontSizePt + 'pt');
    }
    if (props.color && props.color !== propsBase.color) {
      estilos.push('color:' + props.color);
    }
    if (props.strike) estilos.push('text-decoration:line-through');
    if (props.allCaps) estilos.push('text-transform:uppercase');

    if (estilos.length) {
      const span = document.createElement('span');
      span.style.cssText = estilos.join(';');
      span.appendChild(nodo);
      nodo = span;
    }

    if (props.bold) {
      const b = document.createElement('b');
      b.appendChild(nodo);
      nodo = b;
    }
    if (props.italic) {
      const i = document.createElement('i');
      i.appendChild(nodo);
      nodo = i;
    }
    if (props.underline) {
      const u = document.createElement('u');
      u.appendChild(nodo);
      nodo = u;
    }
    if (props.vertAlign === 'superscript') {
      const sup = document.createElement('sup');
      sup.appendChild(nodo);
      nodo = sup;
    } else if (props.vertAlign === 'subscript') {
      const sub = document.createElement('sub');
      sub.appendChild(nodo);
      nodo = sub;
    }

    return nodo;
  }

  function procesarContenidoDocx(nodo, destino, propsBase, ctx) {
    for (const hijo of Array.from(nodo.childNodes)) {
      if (hijo.nodeType !== Node.ELEMENT_NODE || hijo.namespaceURI !== NS_W) continue;

      const local = hijo.localName;

      if (local === 'r') {
        const elRun = crearRunDocx(hijo, propsBase, ctx);
        if (elRun) destino.appendChild(elRun);

      } else if (local === 'hyperlink') {
        const a = document.createElement('a');
        const rid = hijo.getAttributeNS(NS_R, 'id');
        const ancla = attrW(hijo, 'anchor');
        let href = null;

        if (rid && ctx.relaciones[rid]) href = ctx.relaciones[rid];
        else if (ancla) href = '#' + ancla;

        if (href) {
          a.href = href;
          a.target = '_blank';
          a.rel = 'noopener noreferrer';
        }

        procesarContenidoDocx(hijo, a, propsBase, ctx);
        if (a.childNodes.length) destino.appendChild(a);

      } else if (local === 'ins' || local === 'smartTag' || local === 'fldSimple' ||
                 local === 'dir' || local === 'bdo' || local === 'customXml') {
        procesarContenidoDocx(hijo, destino, propsBase, ctx);

      } else if (local === 'sdt') {
        const contenido = primerW(hijo, 'sdtContent');
        if (contenido) procesarContenidoDocx(contenido, destino, propsBase, ctx);

      } else if (local === 'p') {
        destino.appendChild(crearParrafoDocx(hijo, ctx));
      }
    }
  }

  function crearParrafoDocx(p, ctx) {
    const pPr = primerW(p, 'pPr');

    const estiloEl = pPr ? primerW(pPr, 'pStyle') : null;
    const estiloId = estiloEl ? attrW(estiloEl, 'val') : ctx.datos.defaultPara;
    const estilo = resolverEstiloDocx(ctx.datos, estiloId, ctx.cache, new Set());
    const directo = leerPPrDocx(pPr);

    const propsParrafo = Object.assign({}, ctx.datos.defaults.pPr, estilo.pPr, directo);
    const propsTexto = Object.assign(
      {},
      ctx.datos.defaults.rPr,
      estilo.rPr,
      pPr ? leerRPrDocx(primerW(pPr, 'rPr')) : {}
    );

    const el = document.createElement('p');

    if (propsParrafo.align) el.style.textAlign = propsParrafo.align;
    if (propsParrafo.lineHeightCss) el.style.lineHeight = propsParrafo.lineHeightCss;
    else el.style.lineHeight = String(propsParrafo.line != null ? propsParrafo.line : 1);
    if (propsParrafo.marginLeft) el.style.marginLeft = propsParrafo.marginLeft;
    if (propsParrafo.marginRight) el.style.marginRight = propsParrafo.marginRight;
    if (propsParrafo.textIndent) el.style.textIndent = propsParrafo.textIndent;

    if (propsTexto.fontFamily) el.style.fontFamily = propsTexto.fontFamily;
    if (propsTexto.fontSizePt) el.style.fontSize = propsTexto.fontSizePt + 'pt';
    if (propsTexto.color) el.style.color = propsTexto.color;

    procesarContenidoDocx(p, el, propsTexto, ctx);

    if (!el.childNodes.length) el.appendChild(document.createElement('br'));
    return el;
  }

  function crearTablaDocx(tabla, ctx) {
    const table = document.createElement('table');

    for (const fila of hijosW(tabla, 'tr')) {
      const tr = document.createElement('tr');

      for (const celda of hijosW(fila, 'tc')) {
        const td = document.createElement('td');
        procesarBloquesDocx(celda, td, ctx);
        if (!td.childNodes.length) td.appendChild(document.createElement('br'));
        tr.appendChild(td);
      }

      table.appendChild(tr);
    }

    return table;
  }

  function procesarBloquesDocx(nodo, contenedor, ctx) {
    let lista = null;

    for (const hijo of Array.from(nodo.childNodes)) {
      if (hijo.nodeType !== Node.ELEMENT_NODE || hijo.namespaceURI !== NS_W) continue;

      const local = hijo.localName;

      if (local === 'p') {
        const pPr = primerW(hijo, 'pPr');
        const numPr = pPr ? primerW(pPr, 'numPr') : null;
        const parrafo = crearParrafoDocx(hijo, ctx);

        if (numPr) {
          if (!lista) {
            lista = document.createElement('ul');
            contenedor.appendChild(lista);
          }
          const li = document.createElement('li');
          li.appendChild(parrafo);
          lista.appendChild(li);
        } else {
          lista = null;
          contenedor.appendChild(parrafo);
        }

      } else if (local === 'tbl') {
        lista = null;
        contenedor.appendChild(crearTablaDocx(hijo, ctx));

      } else if (local === 'sdt') {
        const contenido = primerW(hijo, 'sdtContent');
        if (contenido) procesarBloquesDocx(contenido, contenedor, ctx);

      } else if (local !== 'sectPr' && local !== 'bookmarkStart' && local !== 'bookmarkEnd') {
        procesarBloquesDocx(hijo, contenedor, ctx);
      }
    }
  }

  async function docxXmlAHtml(buffer) {
    if (!window.JSZip) throw new Error('No se pudo cargar el lector de DOCX.');

    const zip = await window.JSZip.loadAsync(buffer);
    const archivo = zip.file('word/document.xml');
    if (!archivo) throw new Error('El archivo DOCX no contiene word/document.xml.');

    const parser = new DOMParser();
    const doc = parser.parseFromString(await archivo.async('string'), 'application/xml');
    if (doc.querySelector('parsererror')) throw new Error('El contenido del DOCX no es válido.');

    let docEstilos = null;
    const archivoEstilos = zip.file('word/styles.xml');
    if (archivoEstilos) {
      const estilos = parser.parseFromString(await archivoEstilos.async('string'), 'application/xml');
      if (!estilos.querySelector('parsererror')) docEstilos = estilos;
    }

    const datos = leerEstilosDocx(docEstilos);
    const relaciones = await leerRelacionesDocx(zip);
    const ctx = { datos, relaciones, cache: Object.create(null) };

    const cuerpo = doc.getElementsByTagNameNS(NS_W, 'body')[0];
    if (!cuerpo) return '';

    const avisos = [];
    if (Object.keys(zip.files).some(n => n.startsWith('word/media/'))) avisos.push('imágenes');
    if (zip.file('word/footnotes.xml')) avisos.push('notas al pie');
    if (zip.file('word/endnotes.xml')) avisos.push('notas finales');
    if (zip.file('word/comments.xml')) avisos.push('comentarios');
    if (Object.keys(zip.files).some(n => /^word\/(header|footer)\d*\.xml$/.test(n))) {
      avisos.push('encabezados o pies de página');
    }
    if (cuerpo.getElementsByTagNameNS(NS_W, 'tbl').length) avisos.push('tablas');
    if (cuerpo.getElementsByTagNameNS(NS_W, 'txbxContent').length) avisos.push('cuadros de texto');
    if (cuerpo.getElementsByTagNameNS(NS_W, 'ins').length ||
        cuerpo.getElementsByTagNameNS(NS_W, 'del').length) {
      avisos.push('control de cambios');
    }

    if (avisos.length) {
      alert('El documento contiene ' + avisos.join(', ') + '.\n\n' +
            'Se abrirá el texto, pero esos elementos no se conservarán si guardas desde este procesador.');
    }

    const contenedor = document.createElement('div');
    procesarBloquesDocx(cuerpo, contenedor, ctx);
    return contenedor.innerHTML;
  }

  /* ---------- ODT: lectura ---------- */
  async function odtAHtml(buffer) {
    if (!window.JSZip) throw new Error('No se pudo cargar el lector de ODT.');

    const zip = await window.JSZip.loadAsync(buffer);
    const archivoContenido = zip.file('content.xml');
    if (!archivoContenido) throw new Error('El archivo ODT no contiene content.xml.');

    const parser = new DOMParser();
    const doc = parser.parseFromString(await archivoContenido.async('string'), 'application/xml');
    if (doc.querySelector('parsererror')) throw new Error('El contenido del ODT no es válido.');

    let docEstilos = null;
    const archivoEstilos = zip.file('styles.xml');
    if (archivoEstilos) {
      const estilos = parser.parseFromString(await archivoEstilos.async('string'), 'application/xml');
      if (!estilos.querySelector('parsererror')) docEstilos = estilos;
    }

    return odtDocumentoAHtml(doc, docEstilos);
  }

  function recolectarEstilos(doc) {
    const estilos = Object.create(null);
    if (!doc) return estilos;

    const lista = doc.getElementsByTagNameNS(NS.style, 'style');
    for (let i = 0; i < lista.length; i++) {
      const el = lista[i];
      const nombre = el.getAttributeNS(NS.style, 'name');
      if (!nombre) continue;

      const pp = el.getElementsByTagNameNS(NS.style, 'paragraph-properties')[0] || null;
      const tp = el.getElementsByTagNameNS(NS.style, 'text-properties')[0] || null;

      estilos[nombre] = {
        padre: el.getAttributeNS(NS.style, 'parent-style-name') || null,
        align: pp ? pp.getAttributeNS(NS.fo, 'text-align') : null,
        lineHeight: pp ? pp.getAttributeNS(NS.fo, 'line-height') : null,
        textIndent: pp ? pp.getAttributeNS(NS.fo, 'text-indent') : null,
        fontWeight: tp ? tp.getAttributeNS(NS.fo, 'font-weight') : null,
        fontStyle: tp ? tp.getAttributeNS(NS.fo, 'font-style') : null,
        fontSize: tp ? tp.getAttributeNS(NS.fo, 'font-size') : null,
        fontFamily: tp ? tp.getAttributeNS(NS.fo, 'font-family') : null,
        color: tp ? tp.getAttributeNS(NS.fo, 'color') : null,
        underline: tp ? tp.getAttributeNS(NS.style, 'text-underline-style') : null,
        textPosition: tp ? tp.getAttributeNS(NS.style, 'text-position') : null
      };
    }
    return estilos;
  }

  function resolverEstilo(estilos, nombre, cache, visitando) {
    if (!nombre) return {};
    if (cache[nombre]) return cache[nombre];
    if (!estilos[nombre]) return {};
    if (visitando.has(nombre)) return {};

    visitando.add(nombre);

    const propio = estilos[nombre];
    const base = propio.padre ? resolverEstilo(estilos, propio.padre, cache, visitando) : {};
    const resultado = Object.assign({}, base);

    Object.keys(propio).forEach(clave => {
      if (clave !== 'padre' && propio[clave] != null) resultado[clave] = propio[clave];
    });

    visitando.delete(nombre);
    cache[nombre] = resultado;
    return resultado;
  }

  function normalizarInterlineado(valor) {
    if (/^\d+(\.\d+)?%$/.test(valor)) return String(parseFloat(valor) / 100);
    return valor;
  }

  function aplicarPropiedadesTexto(el, props) {
    if (props.fontFamily) el.style.fontFamily = props.fontFamily.replace(/^['"]|['"]$/g, '');
    if (props.fontSize) el.style.fontSize = props.fontSize;
    if (props.fontWeight) el.style.fontWeight = props.fontWeight;
    if (props.fontStyle) el.style.fontStyle = props.fontStyle;
    if (props.color) el.style.color = props.color;
    if (props.underline && props.underline !== 'none') el.style.textDecoration = 'underline';
  }

  function etiquetaSuperSub(props) {
    if (!props.textPosition) return null;
    const valor = String(props.textPosition).trim().toLowerCase();
    if (valor.indexOf('super') === 0) return 'sup';
    if (valor.indexOf('sub') === 0) return 'sub';
    return null;
  }

  function odtDocumentoAHtml(doc, docEstilos) {
    const estilos = Object.assign(
      Object.create(null),
      recolectarEstilos(docEstilos),
      recolectarEstilos(doc)
    );
    const cache = Object.create(null);

    const cuerpo = doc.getElementsByTagNameNS(NS.office, 'text')[0];
    if (!cuerpo) return '';

    const contenedor = document.createElement('div');

    function crearParrafo(nodo, estiloPadre) {
      const props = resolverEstilo(estilos, nodo.getAttributeNS(NS.text, 'style-name'), cache, new Set());

      let el;
      if (nodo.localName === 'h') {
        const nivel = parseInt(nodo.getAttributeNS(NS.text, 'outline-level') || '1', 10);
        el = document.createElement('h' + Math.min(6, Math.max(1, nivel)));
      } else {
        el = document.createElement('p');
      }

      if (props.align) {
        const mapa = {
          start: 'left', end: 'right',
          center: 'center', justify: 'justify',
          left: 'left', right: 'right'
        };
        el.style.textAlign = mapa[props.align] || props.align;
      }
      if (props.lineHeight) el.style.lineHeight = normalizarInterlineado(props.lineHeight);
      if (props.textIndent) el.style.textIndent = props.textIndent;

      aplicarPropiedadesTexto(el, props);
      if (estiloPadre) aplicarPropiedadesTexto(el, estiloPadre);

      procesarHijos(nodo, el, props);
      return el;
    }

    function procesarHijos(nodo, destino, estiloTextoPadre) {
      const hijos = Array.from(nodo.childNodes);

      for (const hijo of hijos) {
        if (hijo.nodeType === Node.TEXT_NODE) {
          if (hijo.nodeValue) destino.appendChild(document.createTextNode(hijo.nodeValue));
          continue;
        }
        if (hijo.nodeType !== Node.ELEMENT_NODE) continue;

        const ns = hijo.namespaceURI;
        const local = hijo.localName;

        if (ns !== NS.text) {
          // Tablas y otros contenedores: se recorre su contenido para no perder texto.
          procesarHijos(hijo, destino, estiloTextoPadre);
          continue;
        }

        if (local === 'p' || local === 'h') {
          destino.appendChild(crearParrafo(hijo, estiloTextoPadre));

        } else if (local === 'span') {
          const props = resolverEstilo(estilos, hijo.getAttributeNS(NS.text, 'style-name'), cache, new Set());
          const tag = etiquetaSuperSub(props);
          const el = document.createElement(tag || 'span');
          aplicarPropiedadesTexto(el, props);
          if (tag) el.style.fontSize = ''; // el navegador ya reduce sup/sub
          procesarHijos(hijo, el, props);
          destino.appendChild(el);

        } else if (local === 's') {
          const cantidad = parseInt(hijo.getAttributeNS(NS.text, 'c') || '1', 10) || 1;
          destino.appendChild(document.createTextNode(' '.repeat(cantidad)));

        } else if (local === 'tab') {
          const span = document.createElement('span');
          span.style.whiteSpace = 'pre';
          span.textContent = '\t';
          destino.appendChild(span);

        } else if (local === 'line-break') {
          destino.appendChild(document.createElement('br'));

        } else if (local === 'a') {
          const a = document.createElement('a');
          const href = hijo.getAttributeNS(NS.xlink, 'href');
          if (href) {
            a.href = href;
            a.target = '_blank';
            a.rel = 'noopener noreferrer';
          }
          procesarHijos(hijo, a, estiloTextoPadre);
          destino.appendChild(a);

        } else if (local === 'list') {
          const ul = document.createElement('ul');
          procesarHijos(hijo, ul, estiloTextoPadre);
          destino.appendChild(ul);

        } else if (local === 'list-item') {
          const li = document.createElement('li');
          procesarHijos(hijo, li, estiloTextoPadre);
          destino.appendChild(li);

        } else {
          procesarHijos(hijo, destino, estiloTextoPadre);
        }
      }
    }

    procesarHijos(cuerpo, contenedor, null);
    return contenedor.innerHTML;
  }

  /* =========================================================
     EXTRACCIÓN DE PÁRRAFOS DESDE EL EDITOR
     ========================================================= */
  const TAGS_BLOQUE = new Set([
    'DIV', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6',
    'LI', 'BLOCKQUOTE', 'PRE', 'SECTION', 'ARTICLE'
  ]);

  function extraerParrafos(editor) {
    const parrafos = [];
    let actual = null;

    function interlineadoDe(cs) {
      const lh = parseFloat(cs.lineHeight);
      const fs = parseFloat(cs.fontSize);
      if (!lh || !fs) return 1.5;
      return Math.round((lh / fs) * 100) / 100;
    }

    function alineacionDe(valor) {
      if (valor === 'center') return 'center';
      if (valor === 'right' || valor === 'end') return 'right';
      if (valor === 'justify') return 'justify';
      return 'left';
    }

    function propsBloque(el) {
      const cs = getComputedStyle(el);
      return {
        align: alineacionDe(cs.textAlign),
        line: interlineadoDe(cs),
        indent: cs.textIndent
      };
    }

    function propsEditor() {
      const cs = getComputedStyle(editor);
      return {
        align: 'left',
        line: interlineadoDe(cs),
        indent: cs.textIndent
      };
    }

    function formatoBase() {
      const cs = getComputedStyle(editor);
      return {
        bold: parseInt(cs.fontWeight, 10) >= 600,
        italic: cs.fontStyle === 'italic' || cs.fontStyle === 'oblique',
        underline: (cs.textDecorationLine || cs.textDecoration || '').indexOf('underline') !== -1,
        fontFamily: primeraFuente(cs.fontFamily),
        fontSizePt: Math.round(parseFloat(cs.fontSize) * 72 / 96 * 10) / 10,
        color: cs.color,
        vertAlign: null,
        href: null
      };
    }

    function formatoDe(el, fmtPadre) {
      const cs = getComputedStyle(el);
      const fmt = Object.assign({}, fmtPadre);

      fmt.bold = parseInt(cs.fontWeight, 10) >= 600;
      fmt.italic = cs.fontStyle === 'italic' || cs.fontStyle === 'oblique';
      fmt.underline = (cs.textDecorationLine || cs.textDecoration || '').indexOf('underline') !== -1;
      fmt.fontFamily = primeraFuente(cs.fontFamily) || fmtPadre.fontFamily;
      fmt.color = cs.color || fmtPadre.color;

      let tamanio = parseFloat(cs.fontSize) * 72 / 96;
      if (el.tagName === 'SUB' || el.tagName === 'SUP') tamanio = fmtPadre.fontSizePt;
      fmt.fontSizePt = Math.round(tamanio * 10) / 10;

      if (el.tagName === 'SUB' || cs.verticalAlign === 'sub') fmt.vertAlign = 'subscript';
      else if (el.tagName === 'SUP' || cs.verticalAlign === 'super') fmt.vertAlign = 'superscript';
      else fmt.vertAlign = null;

      if (el.tagName === 'A' && el.getAttribute('href')) fmt.href = el.getAttribute('href');
      else fmt.href = fmtPadre.href;

      return fmt;
    }

    function asegurarParrafo() {
      if (!actual) {
        const props = propsEditor();
        actual = { align: props.align, line: props.line, indent: props.indent, runs: [] };
      }
    }

    function cerrarParrafo() {
      if (!actual) return;
      parrafos.push(actual);
      actual = null;
    }

    function caminar(nodo, fmt) {
      if (nodo.nodeType === Node.TEXT_NODE) {
        if (!nodo.nodeValue) return;
        asegurarParrafo();
        actual.runs.push({ text: nodo.nodeValue, fmt });
        return;
      }
      if (nodo.nodeType !== Node.ELEMENT_NODE) return;

      const tag = nodo.tagName;

      if (TAGS_BLOQUE.has(tag) && nodo !== editor) {
        cerrarParrafo();
        const props = propsBloque(nodo);
        actual = { align: props.align, line: props.line, indent: props.indent, runs: [] };
        const fmtHijo = formatoDe(nodo, fmt);
        for (const hijo of Array.from(nodo.childNodes)) caminar(hijo, fmtHijo);
        cerrarParrafo();
        return;
      }

      if (tag === 'BR') {
        asegurarParrafo();
        actual.runs.push({ br: true });
        return;
      }

      const fmtHijo = formatoDe(nodo, fmt);
      for (const hijo of Array.from(nodo.childNodes)) caminar(hijo, fmtHijo);
    }

    const base = formatoBase();
    for (const hijo of Array.from(editor.childNodes)) caminar(hijo, base);
    cerrarParrafo();

    if (!parrafos.length) parrafos.push({ align: 'left', line: 1.5, indent: '0px', runs: [] });
    return parrafos;
  }

  /* =========================================================
     ESCRITURA: DOCX
     ========================================================= */
  function docxTextoXml(texto) {
    return String(texto).split(/(\t|\n)/).map(parte => {
      if (!parte) return '';
      if (parte === '\t') return '<w:tab/>';
      if (parte === '\n') return '<w:br/>';
      return `<w:t xml:space="preserve">${escaparXml(parte)}</w:t>`;
    }).join('');
  }

  function docxParrafoXml(p, relaciones) {
    let pPr = '';

    if (p.line) pPr += `<w:spacing w:line="${Math.round(p.line * 240)}" w:lineRule="auto"/>`;

    const indentPt = ptDesdeCss(p.indent);
    if (indentPt) {
      const twips = Math.round(Math.abs(indentPt) * 20);
      pPr += indentPt > 0
        ? `<w:ind w:firstLine="${twips}"/>`
        : `<w:ind w:hanging="${twips}"/>`;
    }

    if (p.align && p.align !== 'left') {
      pPr += `<w:jc w:val="${p.align === 'justify' ? 'both' : p.align}"/>`;
    }

    let contenido = '';

    for (const run of p.runs) {
      if (run.br) {
        contenido += '<w:r><w:br/></w:r>';
        continue;
      }
      if (!run.text) continue;

      const f = run.fmt;
      let rPr = '';

      if (f.href) rPr += '<w:rStyle w:val="Hyperlink"/>';
      if (f.fontFamily) {
        const fuente = escaparAtributo(f.fontFamily);
        rPr += `<w:rFonts w:ascii="${fuente}" w:hAnsi="${fuente}"/>`;
      }
      if (f.bold) rPr += '<w:b/>';
      if (f.italic) rPr += '<w:i/>';
      if (f.href) {
        rPr += '<w:color w:val="0563C1"/>';
      } else {
        const color = colorAHex(f.color);
        if (color && color !== '000000') rPr += `<w:color w:val="${color}"/>`;
      }
      if (f.fontSizePt) {
        const mitad = Math.round(f.fontSizePt * 2);
        rPr += `<w:sz w:val="${mitad}"/><w:szCs w:val="${mitad}"/>`;
      }
      if (f.underline || f.href) rPr += '<w:u w:val="single"/>';
      if (f.vertAlign === 'superscript') rPr += '<w:vertAlign w:val="superscript"/>';
      else if (f.vertAlign === 'subscript') rPr += '<w:vertAlign w:val="subscript"/>';

      const runXml = `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}` +
                     `${docxTextoXml(run.text)}</w:r>`;

      if (f.href) {
        if (!relaciones.has(f.href)) {
          relaciones.set(f.href, 'rId' + (relaciones.size + 2)); // rId1 = styles
        }
        contenido += `<w:hyperlink r:id="${relaciones.get(f.href)}">${runXml}</w:hyperlink>`;
      } else {
        contenido += runXml;
      }
    }

    return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${contenido}</w:p>`;
  }

  function docxEstilosXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
        <w:sz w:val="24"/><w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>
  </w:style>
</w:styles>`;
  }

  async function generarDocx(editor) {
    if (!window.JSZip) throw new Error('No se pudo cargar el generador de DOCX.');

    const parrafos = extraerParrafos(editor);
    const relaciones = new Map();
    const cuerpo = parrafos.map(p => docxParrafoXml(p, relaciones)).join('');

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${cuerpo}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"
               w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

    const relacionesXml = Array.from(relaciones.entries()).map(([href, id]) =>
      `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ` +
      `Target="${escaparAtributo(href)}" TargetMode="External"/>`
    ).join('');

    const documentRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  ${relacionesXml}
</Relationships>`;

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

    const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const zip = new window.JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.file('_rels/.rels', rootRels);
    zip.file('word/document.xml', documentXml);
    zip.file('word/styles.xml', docxEstilosXml());
    zip.file('word/_rels/document.xml.rels', documentRels);

    return await zip.generateAsync({ type: 'blob', mimeType: MIME.docx });
  }

  /* =========================================================
     ESCRITURA: ODT
     ========================================================= */
  function odtTexto(texto) {
    let salida = '';
    let i = 0;

    while (i < texto.length) {
      const ch = texto[i];

      if (ch === ' ') {
        let j = i;
        while (j < texto.length && texto[j] === ' ') j++;
        const cantidad = j - i;
        salida += cantidad === 1 ? ' ' : ` <text:s text:c="${cantidad - 1}"/>`;
        i = j;
      } else if (ch === '\t') {
        salida += '<text:tab/>';
        i++;
      } else if (ch === '\n') {
        salida += '<text:line-break/>';
        i++;
      } else {
        salida += escaparXml(ch);
        i++;
      }
    }

    return salida;
  }

  async function generarOdt(editor) {
    if (!window.JSZip) throw new Error('No se pudo cargar el generador de ODT.');

    const parrafos = extraerParrafos(editor);

    const estilosParrafo = [];
    const estilosTexto = [];
    const indiceParrafo = Object.create(null);
    const indiceTexto = Object.create(null);

    function nombreEstiloParrafo(p) {
      const clave = p.align + '|' + p.line + '|' + (p.indent || '');
      if (!indiceParrafo[clave]) {
        const nombre = 'P' + (estilosParrafo.length + 1);
        indiceParrafo[clave] = nombre;
        estilosParrafo.push({ nombre, align: p.align, line: p.line, indent: p.indent });
      }
      return indiceParrafo[clave];
    }

    function nombreEstiloTexto(f) {
      const clave = [
        f.bold, f.italic, f.underline, f.fontFamily,
        f.fontSizePt, f.color, f.vertAlign, f.href ? '1' : '0', f.href || ''
      ].join('|');

      if (!indiceTexto[clave]) {
        const nombre = 'T' + (estilosTexto.length + 1);
        indiceTexto[clave] = nombre;
        estilosTexto.push(Object.assign({ nombre }, f));
      }
      return indiceTexto[clave];
    }

    function parrafoXml(p) {
      const estiloP = nombreEstiloParrafo(p);
      let contenido = '';

      for (const run of p.runs) {
        if (run.br) {
          contenido += '<text:line-break/>';
          continue;
        }
        if (!run.text) continue;

        const estiloT = nombreEstiloTexto(run.fmt);
        const span = `<text:span text:style-name="${estiloT}">${odtTexto(run.text)}</text:span>`;

        if (run.fmt.href) {
          contenido += `<text:a xlink:href="${escaparAtributo(run.fmt.href)}" office:type="simple">${span}</text:a>`;
        } else {
          contenido += span;
        }
      }

      return `<text:p text:style-name="${estiloP}">${contenido}</text:p>`;
    }

    const parrafosXml = parrafos.map(parrafoXml).join('');

    const estilosParrafoXml = estilosParrafo.map(s => {
      let props = '';
      if (s.align && s.align !== 'left') {
        const align = s.align === 'right' ? 'end' : s.align === 'left' ? 'start' : s.align;
        props += ` fo:text-align="${align}"`;
      }
      if (s.line) props += ` fo:line-height="${Math.round(s.line * 100)}%"`;
      const indentPt = ptDesdeCss(s.indent);
      if (indentPt) props += ` fo:text-indent="${indentPt}pt"`;

      return `<style:style style:name="${s.nombre}" style:family="paragraph" style:parent-style-name="Standard">` +
             (props ? `<style:paragraph-properties${props}/>` : '') +
             `</style:style>`;
    }).join('');

    const estilosTextoXml = estilosTexto.map(s => {
      let props = '';
      if (s.fontFamily) props += ` fo:font-family="${escaparAtributo(s.fontFamily)}"`;
      if (s.fontSizePt) props += ` fo:font-size="${s.fontSizePt}pt"`;
      if (s.bold) props += ' fo:font-weight="bold"';
      if (s.italic) props += ' fo:font-style="italic"';
      if (s.underline || s.href) {
        props += ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"';
      }
      if (s.vertAlign === 'superscript') props += ' style:text-position="super 58%"';
      else if (s.vertAlign === 'subscript') props += ' style:text-position="sub 58%"';
      if (s.href) {
        props += ' fo:color="#0563C1"';
      } else {
        const color = colorAHex(s.color);
        if (color && color !== '000000') props += ` fo:color="#${color}"`;
      }

      return `<style:style style:name="${s.nombre}" style:family="text">` +
             `<style:text-properties${props}/>` +
             `</style:style>`;
    }).join('');

    const contentXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    office:version="1.2">
  <office:automatic-styles>
    ${estilosParrafoXml}
    ${estilosTextoXml}
  </office:automatic-styles>
  <office:body>
    <office:text>
      ${parrafosXml}
    </office:text>
  </office:body>
</office:document-content>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    office:version="1.2">
  <office:styles>
    <style:default-style style:family="paragraph">
      <style:paragraph-properties fo:line-height="150%"/>
      <style:text-properties fo:font-family="Arial" fo:font-size="12pt"/>
    </style:default-style>
    <style:style style:name="Standard" style:family="paragraph" style:class="text"/>
  </office:styles>
</office:document-styles>`;

    const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`;

    const metaXml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
    office:version="1.2">
  <office:meta>
    <meta:generator>Procesador de Textos Minimalista</meta:generator>
  </office:meta>
</office:document-meta>`;

    const zip = new window.JSZip();
    zip.file('mimetype', MIME.odt, { compression: 'STORE' });
    zip.file('META-INF/manifest.xml', manifestXml);
    zip.file('content.xml', contentXml);
    zip.file('styles.xml', stylesXml);
    zip.file('meta.xml', metaXml);

    return await zip.generateAsync({ type: 'blob', mimeType: MIME.odt });
  }

  /* =========================================================
     API pública
     ========================================================= */
  window.Formatos = {
    abrir,
    generarDocx,
    generarOdt,
    formatoDeNombre,
    extensionDe,
    mimeDe
  };

})();
