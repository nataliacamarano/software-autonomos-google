// ============================================================
// MaquinariaEnlace.gs — Rellenar una máquina a partir del enlace de una tienda
// ============================================================
//
// SIN IA: solo reglas fijas sobre los datos que casi todas las tiendas publican
// dentro de la propia página:
//   1. JSON-LD de schema.org (<script type="application/ld+json"> con un "Product")
//   2. Etiquetas Open Graph / meta (og:title, og:image, og:description...)
//   3. El <title> de la página
// Del título se deducen la marca (lista de marcas conocidas) y el modelo, y el tipo
// de máquina por palabras clave. Todo se muestra para que Álvaro lo revise.
//
// Límites reales (probados con tiendas de verdad): algunas tiendas bloquean a los
// programas que no son un navegador (error 403) o cargan la ficha con JavaScript;
// en esos casos se avisa y se rellena a mano.
//
// Seguridad: solo http/https a webs públicas con nombre de dominio (nunca IPs,
// localhost ni puertos raros) y solo Álvaro puede usarlo.
// ============================================================

var MAQE_ = {
  MAX_HTML: 1500000,
  UA: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36',
  MARCAS: ['Stihl', 'Husqvarna', 'Honda', 'Bosch', 'Makita', 'DeWalt', 'Ryobi', 'Black+Decker', 'Black & Decker', 'Einhell',
           'Kärcher', 'Karcher', 'Echo', 'Stiga', 'Greenworks', 'Worx', 'Milwaukee', 'Hilti', 'Metabo', 'Ego', 'Oleo-Mac',
           'Oleo Mac', 'Efco', 'Mitox', 'Al-Ko', 'McCulloch', 'Dolmar', 'Cub Cadet', 'Toro', 'Kawasaki', 'Yamaha', 'Victus',
           'Matabi', 'Gardena', 'Fiskars', 'Hyundai', 'Villager', 'Ibea', 'Grizzly', 'Ferm', 'Fuxtec', 'Stanley', 'Hikoki',
           'Hitachi', 'Cramer', 'Shindaiwa', 'Kubota', 'Briggs & Stratton', 'MTD', 'Wolf-Garten', 'Mountfield', 'Sabo',
           'Snapper', 'Weibang', 'Ariens', 'Dremel', 'Festool', 'Texas'],
  // Orden importa: lo más específico primero (un "hedge trimmer" no es una desbrozadora).
  TIPOS: [
    ['Cortasetos',      /cortasetos|hedge ?trimmer|hedge ?cutter|tijera.{0,12}seto|taille-haie/i],
    ['Motosierra',      /motosierra|chain ?saw|tron[cç]onneuse/i],
    ['Soplador',        /soplador|blower|aspirador de hojas|leaf ?blower/i],
    ['Hidrolimpiadora', /hidrolimpiadora|pressure ?washer|limpiadora a presi[oó]n|hidrolavadora/i],
    ['Taladro',         /taladro|drill|atornillador|martillo perforador/i],
    ['Cortacésped',     /cortac[eé]sped|lawn ?mower|tondeuse|\bmower\b/i],
    ['Desbrozadora',    /desbrozadora|brush ?cutter|clearing ?saw|strimmer|\btrimmer\b|decespugliatore|d[eé]broussailleuse/i]
  ],
  PALABRAS_TIPO: /\b(desbrozadora|cortac[eé]sped|motosierra|soplador|cortasetos|hidrolimpiadora|taladro|brush ?cutter|lawn ?mower|chain ?saw|leaf ?blower|blower|hedge ?trimmer|pressure ?washer|strimmer|trimmer|clearing ?saw|drill)\b/gi,
  PALABRAS_GENERICAS: /\b(petrol|gasolina|el[eé]ctric[oa]|electric|cordless|inal[aá]mbric[oa]|profesional|professional|garden|jard[ií]n)\b/gi
};

// ---- Dirección segura ----

/** Valida un enlace: '' si está vacío; si no, devuelve la dirección con https:// o lanza un aviso. */
function _maq_urlSegura_(u) {
  var s = String(u == null ? '' : u).trim();
  if (!s) return '';
  if (s.length > 800 || /\s/.test(s)) _wa_fail_('El enlace no es válido.');
  if (!/^https?:\/\//i.test(s)) s = 'https://' + s;
  var m = s.match(/^https?:\/\/([^\/?#:@]+)(?::(\d+))?([\/?#].*)?$/i);
  if (!m) _wa_fail_('El enlace no es válido.');
  var host = m[1].toLowerCase();
  if (m[2] && m[2] !== '80' && m[2] !== '443') _wa_fail_('Ese enlace no se puede usar.');
  if (!/^[a-z0-9][a-z0-9.-]*\.[a-z]{2,}$/.test(host) || /^\d+\.\d+\.\d+\.\d+$/.test(host) ||
      host === 'localhost' || /\.(local|internal|localhost|lan|home|corp)$/.test(host) || host.indexOf('metadata') !== -1) {
    _wa_fail_('Ese enlace no se puede usar.');
  }
  return s;
}

// ---- Descarga ----

function _maqe_cabecera_(r, nombre) {
  try {
    var h = r.getHeaders() || {}, k;
    for (k in h) { if (String(k).toLowerCase() === nombre) return String(h[k] || ''); }
  } catch (e) {}
  return '';
}

/** Descarga una página web (HTML). Lanza avisos claros cuando la tienda no deja leerla. */
function _maqe_descargaHtml_(url) {
  var r;
  try {
    r = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': MAQE_.UA, 'Accept': 'text/html,application/xhtml+xml', 'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8' }
    });
  } catch (e) {
    logEvent('WARN', '_maqe_descargaHtml_', e.message, url);
    _wa_fail_('No he podido abrir ese enlace. Revisa que esté bien escrito.');
  }
  var code = r.getResponseCode();
  if (code === 401 || code === 403 || code === 429 || code === 503) {
    _wa_fail_('Esa tienda no deja leer sus páginas de forma automática. Rellena los datos a mano (puedes dejar el enlace guardado).');
  }
  if (code === 404 || code === 410) _wa_fail_('Esa página no existe (error ' + code + '). Revisa el enlace.');
  if (code >= 400) _wa_fail_('La tienda ha respondido con un error (' + code + '). Inténtalo más tarde o rellena a mano.');
  var tipo = _maqe_cabecera_(r, 'content-type').toLowerCase();
  if (tipo && !/html|xml/.test(tipo)) _wa_fail_('Ese enlace no es una página web (es ' + tipo.split(';')[0] + ').');
  var charset = (tipo.match(/charset=([\w-]+)/) || [])[1];
  var html = '';
  try { html = charset && !/utf-?8/i.test(charset) ? r.getContentText(charset) : r.getContentText(); }
  catch (e2) { try { html = r.getContentText(); } catch (e3) { html = ''; } }
  if (!charset) {   // algunas webs antiguas solo lo indican dentro de la página
    var mc = html.slice(0, 3000).match(/<meta[^>]+charset=["']?([\w-]+)/i);
    if (mc && !/utf-?8/i.test(mc[1])) { try { html = r.getContentText(mc[1]); } catch (e4) {} }
  }
  return html.length > MAQE_.MAX_HTML ? html.slice(0, MAQE_.MAX_HTML) : html;
}

/** Descarga la foto de la tienda. Devuelve {mime, b64} o null si no se puede usar (nunca lanza). */
function _maqe_bajaFoto_(url) {
  try {
    var u = _maq_urlSegura_(url);
    if (!u) return null;
    var r = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': MAQE_.UA, 'Accept': 'image/jpeg,image/png,image/webp,*/*' } });
    if (r.getResponseCode() !== 200) return null;
    var mime = _maqe_cabecera_(r, 'content-type').toLowerCase().split(';')[0].trim();
    if (!WA_MIME_EXT_[mime]) return null;                 // avif, gif, svg...: no se admiten
    var bytes = r.getBlob().getBytes();
    if (!bytes.length || bytes.length > WA_.MAX_FOTO_BYTES) return null;
    return { mime: mime, b64: Utilities.base64Encode(bytes) };
  } catch (e) {
    logEvent('WARN', '_maqe_bajaFoto_', e.message, String(url).slice(0, 200));
    return null;
  }
}

// ---- Lectura del HTML (sin DOM: solo expresiones regulares) ----

function _maqe_dec_(s) {
  return String(s == null ? '' : s)
    .replace(/&#x([0-9a-f]+);/gi, function(m, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function(m, d) { return String.fromCharCode(parseInt(d, 10)); })
    .replace(/&(quot|apos|lt|gt|nbsp);/gi, function(m, n) { return { quot: '"', apos: "'", lt: '<', gt: '>', nbsp: ' ' }[n.toLowerCase()]; })
    .replace(/&amp;/gi, '&');
}

function _maqe_metas_(html) {
  var out = {};
  (html.match(/<meta\b[^>]*>/gi) || []).forEach(function(tag) {
    var at = {};
    tag.replace(/([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g, function(m, k, v1, v2) { at[k.toLowerCase()] = v1 != null ? v1 : v2; return m; });
    var clave = String(at.property || at.name || at.itemprop || '').toLowerCase();
    if (clave && at.content != null && out[clave] === undefined) out[clave] = _maqe_dec_(at.content).trim();
  });
  return out;
}

function _maqe_buscaProductos_(o, out, prof) {
  if (o == null || prof > 8) return;
  if (Array.isArray(o)) { o.forEach(function(x) { _maqe_buscaProductos_(x, out, prof + 1); }); return; }
  if (typeof o !== 'object') return;
  var t = [].concat(o['@type'] || []);
  if (t.indexOf('Product') !== -1 || t.indexOf('IndividualProduct') !== -1 || t.indexOf('ProductModel') !== -1) out.push(o);
  ['@graph', 'mainEntity', 'mainEntityOfPage', 'itemListElement', 'item', 'hasVariant', 'isVariantOf'].forEach(function(k) {
    if (o[k] && typeof o[k] === 'object') _maqe_buscaProductos_(o[k], out, prof + 1);
  });
}

function _maqe_tiposJsonLd_(html) {
  var tipos = [], re = /<script\b[^>]*type\s*=\s*["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi, m;
  while ((m = re.exec(html)) !== null) {
    var txt = m[1].trim(); if (!txt) continue;
    var j; try { j = JSON.parse(txt); } catch (e) { try { j = JSON.parse(txt.replace(/[\u0000-\u001f]+/g, ' ')); } catch (e2) { continue; } }
    tipos.push(j);
  }
  return tipos;
}

function _maqe_producto_(bloques) {
  var prods = [];
  bloques.forEach(function(j) { _maqe_buscaProductos_(j, prods, 0); });
  prods = prods.filter(function(p) { return _maqe_txt_(p.name); });
  prods.sort(function(a, b) { return _maqe_puntos_(b) - _maqe_puntos_(a); });
  return prods[0] || null;
}
function _maqe_puntos_(p) { return (p.offers ? 2 : 0) + (p.brand ? 1 : 0) + (p.sku || p.mpn ? 1 : 0) + (p.image ? 1 : 0); }

/** Texto de un valor schema.org: cadena, {name}, {@value} o lista (el primero). */
function _maqe_txt_(v) {
  if (v == null) return '';
  if (Array.isArray(v)) { for (var i = 0; i < v.length; i++) { var t = _maqe_txt_(v[i]); if (t) return t; } return ''; }
  if (typeof v === 'object') return _maqe_txt_(v.name != null ? v.name : (v['@value'] != null ? v['@value'] : ''));
  return _maqe_dec_(String(v)).replace(/\s+/g, ' ').trim();
}

function _maqe_absoluta_(u, base) {
  u = _maqe_dec_(u).trim();
  if (!u) return '';
  if (/^\/\//.test(u)) u = 'https:' + u;
  else if (/^\//.test(u)) { var m = String(base).match(/^(https?:\/\/[^\/?#]+)/i); u = m ? m[1] + u : ''; }
  else if (!/^https?:\/\//i.test(u)) return '';
  return u && u.length <= 800 && !/\s/.test(u) ? u : '';
}

function _maqe_imagen_(prod, metas, base) {
  var c = [];
  if (prod && prod.image) {
    [].concat(prod.image).forEach(function(i) { var u = typeof i === 'string' ? i : (i && (i.url || i.contentUrl)); if (u) c.push(u); });
  }
  ['og:image:secure_url', 'og:image', 'twitter:image', 'twitter:image:src'].forEach(function(k) { if (metas[k]) c.push(metas[k]); });
  for (var i = 0; i < c.length; i++) { var u = _maqe_absoluta_(c[i], base); if (u) return u; }
  return '';
}

function _maqe_norm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, ''); }

function _maqe_host_(url) {
  var m = String(url).match(/^https?:\/\/([^\/?#:]+)/i);
  return m ? m[1].toLowerCase().replace(/^www\./, '') : '';
}

/** Quita del título el nombre de la tienda y otros añadidos ("... | Tienda", "Tienda - ..."). */
function _maqe_limpiaTitulo_(titulo, sitio, host) {
  var t = _maqe_dec_(titulo).replace(/\s+/g, ' ').trim();
  var i = t.indexOf(' | ');
  if (i > 3) t = t.slice(0, i).trim();                       // lo que va tras " | " suele ser la tienda
  var etiqueta = (host || '').split('.').slice(-2, -1)[0] || '';
  var tienda = [sitio, etiqueta].filter(function(x) { return x && _maqe_norm_(x).length >= 3; });
  tienda.forEach(function(s) {                               // "Tienda 123 — Producto"
    if (t.toLowerCase().indexOf(String(s).toLowerCase()) === 0) t = t.slice(String(s).length).replace(/^[\s\-–—:|·]+/, '');
  });
  var partes = t.split(/\s[-–—:·]\s/);                       // "Producto - Tienda"
  if (partes.length > 1) {
    var cola = _maqe_norm_(partes[partes.length - 1]);
    if (tienda.some(function(s) { var n = _maqe_norm_(s); return cola === n || (cola.length > 4 && n.indexOf(cola) === 0) || (n.length > 4 && cola.indexOf(n) === 0); })) {
      partes.pop(); t = partes.join(' - ');
    }
  }
  var mc = t.match(/^([A-Za-z0-9._\/]{4,})\s[-–—]\s(.{6,})$/);   // código suelto delante: "41491290900 — Stihl ..."
  if (mc && /\d/.test(mc[1])) t = mc[2];
  return t.slice(0, 80).trim();
}

function _maqe_regexMarca_(m) {
  return new RegExp('(^|[^A-Za-zÀ-ÿ0-9])(' + m.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')(?![A-Za-zÀ-ÿ0-9])', 'i');
}

function _maqe_marcaConocida_(texto) {
  for (var i = 0; i < MAQE_.MARCAS.length; i++) {
    var m = String(texto || '').match(_maqe_regexMarca_(MAQE_.MARCAS[i]));
    if (m) return MAQE_.MARCAS[i];
  }
  return '';
}

function _maqe_marca_(prod, metas, titulo, sitio, host) {
  var candidatas = [prod && _maqe_txt_(prod.brand), metas['product:brand'], metas['og:brand'], metas['brand']];
  var etiqueta = (host || '').split('.').slice(-2, -1)[0] || '';
  for (var i = 0; i < candidatas.length; i++) {
    var b = (candidatas[i] || '').trim();
    if (!b || b.length > 30) continue;
    var n = _maqe_norm_(b);
    if (n === _maqe_norm_(sitio) || n === _maqe_norm_(etiqueta)) continue;   // la "marca" era la propia tienda
    var conocida = _maqe_marcaConocida_(b);
    return conocida && _maqe_norm_(conocida) === n ? conocida : b;
  }
  return _maqe_marcaConocida_(titulo);
}

function _maqe_tipo_(texto) {
  for (var i = 0; i < MAQE_.TIPOS.length; i++) if (MAQE_.TIPOS[i][1].test(texto)) return MAQE_.TIPOS[i][0];
  return '';
}

function _maqe_modelo_(titulo, marca, modeloLd) {
  var ld = String(modeloLd || '').trim();
  if (ld && ld.length <= 40) return ld;
  if (!marca) return '';
  var m = String(titulo).match(_maqe_regexMarca_(marca));
  if (!m) return '';
  var r = String(titulo).slice(m.index + m[1].length + marca.length);
  r = r.split(/,| - | – | — | \| /)[0];
  r = r.replace(/\s+(a|de)\s+(gasolina|bater[ií]a|explosi[oó]n)\b.*$/i, '')
       .replace(MAQE_.PALABRAS_TIPO, ' ').replace(MAQE_.PALABRAS_GENERICAS, ' ')
       .replace(/\s+/g, ' ').replace(/\s+(a|de|con|para|y)$/i, '').replace(/^[\s\-–—:]+|[\s\-–—:,]+$/g, '');
  return r.slice(0, 40).trim();
}

function _maqe_precio_(prod) {
  if (!prod || !prod.offers) return '';
  var cand = [];
  (function rec(o, prof) {
    if (o == null || prof > 4) return;
    if (Array.isArray(o)) { o.forEach(function(x) { rec(x, prof + 1); }); return; }
    if (typeof o !== 'object') return;
    var p = o.price != null ? o.price : (o.lowPrice != null ? o.lowPrice : null);
    if (p !== null && String(p).trim() !== '') cand.push({ p: String(p).trim(), c: String(o.priceCurrency || '').trim() });
    if (o.priceSpecification) rec(o.priceSpecification, prof + 1);
  })(prod.offers, 0);
  if (!cand.length) return '';
  var c = cand[0], n = parseFloat(String(c.p).replace(',', '.'));
  return (isNaN(n) ? c.p : (Math.round(n * 100) / 100).toString().replace('.', ',')) + (c.c ? ' ' + c.c : '');
}

function _maqe_descripcion_(prod, metas) {
  var d = (prod && _maqe_txt_(prod.description)) || metas['og:description'] || metas['description'] || '';
  return _maqe_dec_(d).replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 240);
}

/** Lee la ficha de producto de una página. Devuelve los datos sugeridos para el formulario. */
function _maqe_parse_(html, url) {
  var metas = _maqe_metas_(html);
  var bloques = _maqe_tiposJsonLd_(html);
  var prod = _maqe_producto_(bloques);
  var host = _maqe_host_(url);
  var sitio = metas['og:site_name'] || '';

  var crudo = (prod && _maqe_txt_(prod.name)) || metas['og:title'] || metas['twitter:title'] || '';
  if (!crudo) { var mt = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i); crudo = mt ? mt[1] : ''; }
  var nombre = _maqe_limpiaTitulo_(crudo, sitio, host);
  if (!nombre) _wa_fail_('No he encontrado datos de producto en esa página (puede que la tienda los cargue con JavaScript). Rellena los datos a mano.');

  var marca = _maqe_marca_(prod, metas, nombre, sitio, host);
  var modelo = _maqe_modelo_(nombre, marca, prod && _maqe_txt_(prod.model));
  var descripcion = _maqe_descripcion_(prod, metas);
  // El tipo se deduce SOLO del título y la categoría: la descripción suele listar otras máquinas compatibles y confunde.
  var tipo = _maqe_tipo_(nombre + ' ' + (prod && _maqe_txt_(prod.category) || ''));
  var referencia = prod ? (_maqe_txt_(prod.sku) || _maqe_txt_(prod.mpn) || _maqe_txt_(prod.gtin13) || _maqe_txt_(prod.gtin)) : '';
  var precio = _maqe_precio_(prod);
  var tienda = (sitio ? sitio.split(/\s[|–—-]\s/)[0].trim() : '') || host;

  var ogTipo = String(metas['og:type'] || '').toLowerCase();
  var esLista = ogTipo === 'product.group' || bloques.some(function(j) {
    return [].concat(j && j['@type'] || []).some(function(t) { return t === 'ItemList' || t === 'CollectionPage'; });
  });
  var esProducto = !!prod || /^product(\.item)?$/.test(ogTipo);
  var aviso = '';
  if (esLista && !prod) aviso = 'Parece una página con varios productos, no la ficha de uno solo: revisa bien los datos.';
  else if (!esProducto) aviso = 'No parece la ficha de un producto: revisa bien los datos.';

  var notas = [tienda ? 'Tienda: ' + tienda : '', referencia ? 'Ref.: ' + referencia : '', precio ? 'Precio en la web: ' + precio : ''].filter(Boolean).join(' · ');
  return {
    nombre: nombre, marca: marca, modelo: modelo, marcaModelo: [marca, modelo].filter(Boolean).join(' '), tipo: tipo,
    imagen: _maqe_imagen_(prod, metas, url), descripcion: descripcion, referencia: referencia, precio: precio,
    tienda: tienda, notas: notas, esProducto: esProducto, aviso: aviso, url: url
  };
}

// ---- Acción de la web app (solo Álvaro) ----

function wa_maqLeerEnlace_(user, p) {
  var url = _maq_urlSegura_(p.url);
  if (!url) _wa_fail_('Pega el enlace de la tienda o de una máquina igual.');
  var datos = _maqe_parse_(_maqe_descargaHtml_(url), url);
  logEvent('INFO', 'wa_maqLeerEnlace', datos.tienda + ' → ' + datos.nombre, url.slice(0, 200));
  return datos;
}
