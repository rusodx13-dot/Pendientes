/**
 * Control de Mercadería — backend en Google Apps Script.
 * Se pega dentro de la planilla (Extensiones > Apps Script) y se publica como "Aplicación web":
 *   Ejecutar como: Yo   |   Quién tiene acceso: Cualquier usuario
 * Las solapas se crean solas la primera vez que se usa la app.
 */
var TZ = 'America/Argentina/Buenos_Aires';

var SCHEMA = {
  Movimientos:  {cols: ['c','a','k','pf','pe','f','p','o','y','v','u'], num: ['pf','pe'],
                 head: ['Cuenta','Artículo','Comprobante','Pend. facturar','Pend. entregar','Fecha','Pedido','Observaciones','Proveedor','Vendedor','Últ. actualización']},
  Pedidos:      {cols: ['d','cr','ref','sc','sn','pr','q','um','bl','bd','fe','qe','pe','st'], num: ['q','qe','pe'],
                 head: ['Doc. de ventas','Creado el','Referencia (factura)','Solicitante','Nombre solicitante','Producto','Cantidad pedida','Un.','Bloqueo','Denominación bloqueo','Fecha de entrega','Cantidad entregada','Pend. entrega','Estado de entrega']},
  Estado:       {cols: ['d','rec','recFecha','desbl','desblFecha','desblNota'], num: [], bool: ['rec','desbl'],
                 head: ['Doc. de ventas','Recibido','Fecha recibido','Desbloqueado','Fecha desbloqueo','Nota']},
  Stock:        {cols: ['id','producto','tipo','cantidad','fecha','nota','ts'], num: ['cantidad','ts'],
                 head: ['Id','Producto','Tipo','Cantidad','Fecha','Nota','Marca de tiempo']},
  FacturasMano: {cols: ['id','key','c','a','fac','q','f','sb','ts','env','envF','bq'], num: ['q','sb','ts'],
                 head: ['Id','Clave','Cliente','Artículo','Factura','Cantidad','Fecha','Saldo al cargar','Marca de tiempo','Estado del envío','Fecha del estado','Tipo de bloqueo']},
  DatosFact:    {cols: ['key','valor','fp','vto','cj','cereal','ts'], num: ['valor','ts'],
                 head: ['Clave (cliente|artículo)','Valor a facturar','Forma de pago','Vencimiento','Tipo de canje','Cereal a entregar','Marca de tiempo']},
  Usuarios:     {cols: ['usuario','sal','hash','rol','activo'], num: [], bool: ['activo'],
                 head: ['Usuario','Sal','Clave (hash)','Rol','Activo']},
  Sesiones:     {cols: ['th','usuario','exp'], num: ['exp'], head: ['Token (hash)','Usuario','Vence']},
  Archivo:      {cols: ['c','a','k','pf','pe','f','p','o','y','v','u'], num: ['pf','pe'],
                 head: ['Cuenta','Artículo','Comprobante','Pend. facturar','Pend. entregar','Fecha','Pedido','Observaciones','Proveedor','Vendedor','Últ. actualización']},
  ArchivoClaves: {cols: ['t','x','y'], num: [], head: ['Tipo','Valor','Extra']}
};
var MAX_FILAS_POR_PEDIDO = 2000;

/* ---------- entrada ---------- */
function doGet() {
  return json_({ok: true, app: 'Control de Mercadería'});
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({error: 'Pedido inválido.'}); }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (err) { return json_({error: 'El servidor está ocupado. Probá de nuevo.', code: 'ocupado'}); }
  try {
    asegurarHojas_();
    return json_(despachar_(body));
  } catch (err) {
    if (err && err.pub) return json_({error: err.pub, code: err.code || 'error'});
    return json_({error: 'Error del servidor: ' + (err && err.message ? err.message : err), code: 'servidor'});
  } finally {
    lock.releaseLock();
  }
}

function despachar_(b) {
  switch (b.a) {
    case 'ping':   return {ok: true, setup: !hayUsuarios_()};
    case 'setup':  return setup_(b);
    case 'login':  return login_(b);
  }
  var me = sesion_(b.token);
  switch (b.a) {
    case 'ver':      return {ok: true, v: version_(), me: me};
    case 'cargar':   return cargar_(me, b);
    case 'append':   exigir_(me, 'append'); return append_(b);
    case 'estadoSet': exigir_(me, 'estadoSet'); return upsert_('Estado', b.d, b.o, {d: b.d});
    case 'stkSet':   exigir_(me, 'stkSet'); return upsert_('Stock', b.m.id, b.m);
    case 'stkDel':   exigir_(me, 'stkDel'); return borrar_('Stock', b.id);
    case 'datSet':   exigir_(me, 'datSet'); return upsert_('DatosFact', b.m.key, b.m);
    case 'datLote':  exigir_(me, 'datSet'); return datLote_(b.ms);
    case 'archivar': exigirAdmin_(me); return archivar_(b);
    case 'facSet':   exigir_(me, 'facSet'); return facSet_(me, b.m);
    case 'facEnv':   exigir_(me, 'facEnv'); return upsert_('FacturasMano', b.m.id, b.m);
    case 'facDel':   exigir_(me, 'facDel'); return borrar_('FacturasMano', b.id);
    case 'limpiar':  exigir_(me, 'limpiar'); return limpiar_();
    case 'usuarios': exigirAdmin_(me); return listarUsuarios_();
    case 'usuarioGuardar': exigirAdmin_(me); return guardarUsuario_(b);
    case 'claveCambiar': return cambiarClave_(me, b);
    case 'salir':    return salir_(b.token);
  }
  throw pub_('Acción desconocida.');
}

/* ---------- hojas ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

var SCHEMA_V = '2';
function asegurarHojas_() {
  var ss = ss_(), props = PropertiesService.getScriptProperties(), migrar = props.getProperty('sv') !== SCHEMA_V;
  Object.keys(SCHEMA).forEach(function (nombre) {
    var sh = ss.getSheetByName(nombre), def = SCHEMA[nombre];
    if (sh) {
      if (!migrar) return;
      // planillas creadas con una versión anterior: completa los encabezados de las columnas nuevas
      if (sh.getRange(1, 1, 1, def.head.length).getValues()[0].join('|') !== def.head.join('|')) {
        sh.getRange(1, 1, 1, def.head.length).setValues([def.head]).setFontWeight('bold').setBackground('#123524').setFontColor('#ffffff');
        def.cols.forEach(function (k, i) {
          if ((def.num || []).indexOf(k) < 0 && (def.bool || []).indexOf(k) < 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
        });
      }
      return;
    }
    sh = ss.insertSheet(nombre);
    // columnas de texto en formato "texto plano" para que Sheets no convierta fechas ni números de comprobante
    var textCols = [];
    def.cols.forEach(function (k, i) {
      if ((def.num || []).indexOf(k) < 0 && (def.bool || []).indexOf(k) < 0) textCols.push(i + 1);
    });
    textCols.forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
    sh.getRange(1, 1, 1, def.head.length).setValues([def.head]).setFontWeight('bold').setBackground('#123524').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (nombre === 'Usuarios' || nombre === 'Sesiones') sh.hideSheet();
  });
  if (migrar) props.setProperty('sv', SCHEMA_V);
  var vacia = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja1');
  if (vacia && ss.getSheets().length > 1 && vacia.getLastRow() === 0) ss.deleteSheet(vacia);
}

/* La «época» cambia cuando se reescriben Movimientos o Pedidos (limpiar, archivar): ahí la app vuelve a bajar todo. */
function epoca_() { return Number(PropertiesService.getScriptProperties().getProperty('ep') || 0); }
function subirEpoca_() {
  var p = PropertiesService.getScriptProperties(), ep = epoca_() + 1;
  p.setProperty('ep', String(ep));
  return ep;
}
function filas_(nombre) { return Math.max(0, ss_().getSheetByName(nombre).getLastRow() - 1); }

function version_() { return Number(PropertiesService.getScriptProperties().getProperty('v') || 0); }
function subirVersion_() {
  var p = PropertiesService.getScriptProperties(), v = Number(p.getProperty('v') || 0) + 1;
  p.setProperty('v', String(v));
  return v;
}

function aValor_(def, k, v) {
  if (v === null || v === undefined) return '';
  if ((def.bool || []).indexOf(k) >= 0) return v === true || v === 'TRUE' || v === 'true';
  if ((def.num || []).indexOf(k) >= 0) { var n = Number(v); return isFinite(n) ? n : 0; }
  return String(v);
}

function deValor_(def, k, v) {
  if ((def.bool || []).indexOf(k) >= 0) return v === '' || v === null || v === undefined ? undefined : (v === true || String(v).toUpperCase() === 'TRUE');
  if ((def.num || []).indexOf(k) >= 0) { var n = Number(v); return isFinite(n) ? n : 0; }
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

/* Lee una hoja; con «desde» lee solo las filas que vienen después de las primeras «desde». */
function leer_(nombre, desde) {
  var def = SCHEMA[nombre], sh = ss_().getSheetByName(nombre), d = Math.max(0, Number(desde) || 0), n = sh.getLastRow() - 1 - d;
  if (n < 1) return [];
  var vals = sh.getRange(2 + d, 1, n, def.cols.length).getValues();
  if (d === 0) guardarIndice_(sh, vals.map(function (r) { return r[0]; }));   // de paso, rehace el índice de búsqueda
  return vals.map(function (r) {
    var o = {};
    def.cols.forEach(function (k, i) { o[k] = deValor_(def, k, r[i]); });
    return o;
  }).filter(function (o) { return o[def.cols[0]] !== ''; });
}

function cargar_(me, b) {
  var est = {};
  leer_('Estado').forEach(function (o) {
    var d = o.d; delete o.d;
    est[d] = o;
  });
  var ep = epoca_(), nm = filas_('Movimientos'), np = filas_('Pedidos');
  /* Carga incremental: si la app ya tiene las primeras filas (misma época), se mandan solo las nuevas. */
  var parcial = b && b.ep !== undefined && Number(b.ep) === ep && Number(b.nm) >= 0 && Number(b.np) >= 0 && Number(b.nm) <= nm && Number(b.np) <= np;
  var arch = {fc: [], h: [], v: []};
  leer_('ArchivoClaves').forEach(function (o) {
    if (o.t === 'h') arch.h.push(o.x); else if (o.t === 'fc') arch.fc.push(o.x); else if (o.t === 'v') arch.v.push([o.x, o.y]);
  });
  return {
    ok: true, v: version_(), ep: ep, nm: nm, np: np, parcial: !!parcial, me: me,
    mov: parcial ? leer_('Movimientos', b.nm) : leer_('Movimientos'),
    ped: parcial ? leer_('Pedidos', b.np) : leer_('Pedidos'),
    estado: est, stk: leer_('Stock'), fman: leer_('FacturasMano'), dat: leer_('DatosFact'), arch: arch
  };
}

/* ---------- escritura ---------- */
function append_(b) {
  var nombre = b.kind === 'mov' ? 'Movimientos' : b.kind === 'ped' ? 'Pedidos' : null;
  if (!nombre) throw pub_('Tipo de base inválido.');
  var rows = b.rows || [];
  if (!rows.length) return {ok: true, v: version_()};
  if (rows.length > MAX_FILAS_POR_PEDIDO) throw pub_('Demasiadas filas en un solo envío.');
  var def = SCHEMA[nombre], sh = ss_().getSheetByName(nombre), actuales = Math.max(0, sh.getLastRow() - 1);
  if (Number(b.oldLen) !== actuales) throw pub_('Otra persona actualizó la base mientras tanto. Se recarga para que no se dupliquen datos.', 'conflicto');
  var data = rows.map(function (r) { return def.cols.map(function (k) { return aValor_(def, k, r[k]); }); });
  var primera = actuales + 2, necesarias = primera + data.length - 1;
  if (necesarias > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), necesarias - sh.getMaxRows());
  sh.getRange(primera, 1, data.length, def.cols.length).setValues(data);
  return {ok: true, v: subirVersion_()};
}

/* Quita de Movimientos los renglones de subtotal: sin artículo ni comprobante. */
function limpiar_() {
  var def = SCHEMA.Movimientos, sh = ss_().getSheetByName('Movimientos'), n = sh.getLastRow() - 1;
  if (n < 1) return {ok: true, borradas: 0, v: version_()};
  var ia = def.cols.indexOf('a'), ik = def.cols.indexOf('k');
  var vals = sh.getRange(2, 1, n, def.cols.length).getValues();
  var keep = vals.filter(function (r) { return String(r[ia]).trim() !== '' || String(r[ik]).trim() !== ''; });
  var borradas = vals.length - keep.length;
  if (!borradas) return {ok: true, borradas: 0, v: version_()};
  sh.getRange(2, 1, n, def.cols.length).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, def.cols.length).setValues(keep);
  return {ok: true, borradas: borradas, v: subirVersion_(), ep: subirEpoca_(), nm: keep.length};
}

/* Pasa a «Archivo» los movimientos de los cliente + artículo indicados (cerrados) y guarda sus claves. */
function archivar_(b) {
  if (Number(b.ep) !== epoca_()) throw pub_('La base cambió mientras tanto. Se recarga; volvé a intentar.', 'conflicto');
  var keys = {}; (b.keys || []).forEach(function (k) { keys[String(k)] = 1; });
  var def = SCHEMA.Movimientos, ss = ss_(), sh = ss.getSheetByName('Movimientos'), n = sh.getLastRow() - 1;
  if (n < 1) return {ok: true, archivadas: 0, v: version_(), ep: epoca_()};
  var ic = def.cols.indexOf('c'), ia = def.cols.indexOf('a');
  var vals = sh.getRange(2, 1, n, def.cols.length).getValues(), keep = [], arch = [];
  vals.forEach(function (r) { (keys[String(r[ic]) + '|' + String(r[ia])] ? arch : keep).push(r); });
  if (arch.length !== Number(b.n)) throw pub_('La base cambió mientras tanto. Se recarga; volvé a intentar.', 'conflicto');
  if (!arch.length) return {ok: true, archivadas: 0, v: version_(), ep: epoca_()};
  agregarFilas_(ss.getSheetByName('Archivo'), arch, def.cols.length);
  var claves = [];
  (b.h || []).forEach(function (x) { claves.push(['h', String(x), '']); });
  (b.fc || []).forEach(function (x) { claves.push(['fc', String(x), '']); });
  (b.v || []).forEach(function (x) { claves.push(['v', String(x[0]), String(x[1])]); });
  if (claves.length) agregarFilas_(ss.getSheetByName('ArchivoClaves'), claves, 3);
  sh.getRange(2, 1, n, def.cols.length).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, def.cols.length).setValues(keep);
  return {ok: true, archivadas: arch.length, v: subirVersion_(), ep: subirEpoca_()};
}

function agregarFilas_(sh, data, ancho) {
  var primera = sh.getLastRow() + 1, necesarias = primera + data.length - 1;
  if (necesarias > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), necesarias - sh.getMaxRows());
  sh.getRange(primera, 1, data.length, ancho).setValues(data);
}

/* Índice clave → fila guardado en la caché del script, para no leer toda la columna en cada guardado.
   Solo para las hojas que la app escribe por clave; se rehace al cargar y ante cualquier duda. */
var INDEXADAS = {Estado: 1, Stock: 1, FacturasMano: 1, DatosFact: 1};
function guardarIndice_(sh, col) {
  if (!INDEXADAS[sh.getName()]) return null;
  var k = {};
  col.forEach(function (v, i) { var s = String(v); if (s !== '' && !k.hasOwnProperty(s)) k[s] = i + 2; });
  var txt = JSON.stringify({n: col.length, k: k});
  try { if (txt.length < 95000) CacheService.getScriptCache().put('ix_' + sh.getName(), txt, 21600); } catch (e) {}
  return k;
}
function leerIndice_(sh, n) {
  if (!INDEXADAS[sh.getName()]) return null;
  try {
    var raw = CacheService.getScriptCache().get('ix_' + sh.getName());
    if (!raw) return null;
    var m = JSON.parse(raw);
    return m && m.n === n ? m.k : null;
  } catch (e) { return null; }
}
function olvidarIndice_(sh) { try { CacheService.getScriptCache().remove('ix_' + sh.getName()); } catch (e) {} }

function buscarFila_(sh, clave) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  clave = String(clave);
  var k = leerIndice_(sh, n);
  if (k) {
    if (!k.hasOwnProperty(clave)) return -1;
    if (String(sh.getRange(k[clave], 1).getValue()) === clave) return k[clave];
  }
  var col = sh.getRange(2, 1, n, 1).getValues();
  k = guardarIndice_(sh, col.map(function (r) { return r[0]; }));
  if (k) return k.hasOwnProperty(clave) ? k[clave] : -1;
  for (var i = 0; i < col.length; i++) if (String(col[i][0]) === clave) return i + 2;
  return -1;
}

/* Escribe (o agrega) una fila por clave, sin subir la versión. */
function escribirFila_(nombre, clave, obj, extra) {
  if (clave === undefined || clave === null || clave === '') throw pub_('Falta la clave del registro.');
  var def = SCHEMA[nombre], sh = ss_().getSheetByName(nombre), o = {};
  Object.keys(obj || {}).forEach(function (k) { o[k] = obj[k]; });
  Object.keys(extra || {}).forEach(function (k) { o[k] = extra[k]; });
  var fila = def.cols.map(function (k) { return aValor_(def, k, o[k]); });
  var pos = buscarFila_(sh, clave);
  if (pos < 0) {
    var n = sh.getLastRow() - 1, k = leerIndice_(sh, Math.max(0, n));
    pos = sh.getLastRow() + 1; if (pos > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
    sh.getRange(pos, 1, 1, def.cols.length).setValues([fila]);
    if (k) { k[String(clave)] = pos; try { CacheService.getScriptCache().put('ix_' + nombre, JSON.stringify({n: n + 1, k: k}), 21600); } catch (e) {} }
    return;
  }
  sh.getRange(pos, 1, 1, def.cols.length).setValues([fila]);
}

function upsert_(nombre, clave, obj, extra) {
  escribirFila_(nombre, clave, obj, extra);
  return {ok: true, v: subirVersion_()};
}

/* Varios valores / formas de pago en un solo envío: una sola subida de versión. */
function datLote_(ms) {
  ms = ms || [];
  if (ms.length > 500) throw pub_('Demasiados cambios en un solo envío.');
  ms.forEach(function (m) { escribirFila_('DatosFact', m.key, m); });
  return {ok: true, v: ms.length ? subirVersion_() : version_()};
}

/* El vendedor puede cargar facturas a mano, pero no cambiar el estado del envío de una ya cargada. */
function facSet_(me, m) {
  if (me.rol !== 'admin') {
    var def = SCHEMA.FacturasMano, sh = ss_().getSheetByName('FacturasMano'), pos = buscarFila_(sh, m.id);
    if (pos > 0) {
      var fila = sh.getRange(pos, 1, 1, def.cols.length).getValues()[0];
      ['env', 'envF', 'bq'].forEach(function (k) { m[k] = deValor_(def, k, fila[def.cols.indexOf(k)]); });
    }
  }
  return upsert_('FacturasMano', m.id, m);
}

function borrar_(nombre, clave) {
  var sh = ss_().getSheetByName(nombre), pos = buscarFila_(sh, clave);
  if (pos > 0) { sh.deleteRow(pos); olvidarIndice_(sh); }
  return {ok: true, v: subirVersion_()};
}

/* ---------- usuarios y sesiones ---------- */
function pub_(msg, code) { return {pub: msg, code: code}; }
function hex_(bytes) { return bytes.map(function (x) { return ('0' + (x & 0xff).toString(16)).slice(-2); }).join(''); }
function sha_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }

function hayUsuarios_() { return ss_().getSheetByName('Usuarios').getLastRow() > 1; }

function usuarios_() { return leer_('Usuarios'); }

function setup_(b) {
  if (hayUsuarios_()) throw pub_('El administrador ya está creado.');
  var u = String(b.usuario || '').trim().toLowerCase(), c = String(b.clave || '');
  if (u.length < 3) throw pub_('El usuario necesita al menos 3 letras.');
  if (c.length < 6) throw pub_('La clave necesita al menos 6 caracteres.');
  guardarUsuarioFila_(u, c, 'admin', true);
  return abrirSesion_(u, 'admin', false);
}

function guardarUsuarioFila_(u, clave, rol, activo) {
  var def = SCHEMA.Usuarios, sh = ss_().getSheetByName('Usuarios'), pos = buscarFila_(sh, u);
  var actual = pos > 0 ? sh.getRange(pos, 1, 1, def.cols.length).getValues()[0] : null;
  var sal = actual ? String(actual[1]) : Utilities.getUuid();
  var hash = clave ? sha_(sal + clave) : (actual ? String(actual[2]) : '');
  var fila = [u, sal, hash, rol, activo === true];
  if (pos < 0) pos = sh.getLastRow() + 1;
  sh.getRange(pos, 1, 1, def.cols.length).setValues([fila]);
}

function login_(b) {
  var u = String(b.usuario || '').trim().toLowerCase(), c = String(b.clave || '');
  var us = usuarios_().filter(function (x) { return x.usuario === u; })[0];
  if (!us || !us.activo || sha_(us.sal + c) !== us.hash) throw pub_('Usuario o clave incorrectos.');
  return abrirSesion_(us.usuario, us.rol, !!b.recordar);
}

function abrirSesion_(usuario, rol, recordar) {
  var sh = ss_().getSheetByName('Sesiones'), ahora = Date.now();
  limpiarSesiones_(sh, ahora);
  var token = Utilities.getUuid() + Utilities.getUuid();
  var dur = (recordar ? 60 : 1) * 86400000;
  sh.appendRow([sha_(token), usuario, ahora + dur]);
  return {ok: true, token: token, me: {usuario: usuario, rol: rol}, v: version_()};
}

function limpiarSesiones_(sh, ahora) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  var vals = sh.getRange(2, 1, n, 3).getValues();
  for (var i = vals.length - 1; i >= 0; i--) if (Number(vals[i][2]) < ahora) sh.deleteRow(i + 2);
}

function sesion_(token) {
  if (!token) throw pub_('Iniciá sesión.', 'auth');
  var h = sha_(String(token)), sh = ss_().getSheetByName('Sesiones'), n = sh.getLastRow() - 1;
  if (n < 1) throw pub_('Tu sesión venció. Volvé a ingresar.', 'auth');
  var vals = sh.getRange(2, 1, n, 3).getValues(), ahora = Date.now();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][0] === h && Number(vals[i][2]) >= ahora) {
      var us = usuarios_().filter(function (x) { return x.usuario === vals[i][1]; })[0];
      if (!us || !us.activo) throw pub_('Tu usuario está desactivado.', 'auth');
      return {usuario: us.usuario, rol: us.rol};
    }
  }
  throw pub_('Tu sesión venció. Volvé a ingresar.', 'auth');
}

function salir_(token) {
  var h = sha_(String(token || '')), sh = ss_().getSheetByName('Sesiones'), pos = buscarFila_(sh, h);
  if (pos > 0) sh.deleteRow(pos);
  return {ok: true};
}

/* Roles: admin edita todo; vendedor edita Pendientes de facturar y Pedidos; deposito edita Mercadería normal. */
var ROLES = ['admin', 'vendedor', 'deposito'];
var NOMBRES_ROL = {admin: 'Administrador', vendedor: 'Vendedor', deposito: 'Depósito', consulta: 'Solo consulta'};
var PERMISOS = {
  append: ['admin'], estadoSet: ['admin'], limpiar: ['admin'], facEnv: ['admin'],
  stkSet: ['admin', 'deposito'], stkDel: ['admin', 'deposito'],
  datSet: ['admin', 'vendedor'], facSet: ['admin', 'vendedor'], facDel: ['admin', 'vendedor']
};
function exigir_(me, accion) {
  if ((PERMISOS[accion] || []).indexOf(me.rol) < 0) throw pub_('Tu rol (' + (NOMBRES_ROL[me.rol] || me.rol) + ') no puede editar esto.', 'permiso');
}
function exigirAdmin_(me) { if (me.rol !== 'admin') throw pub_('Solo el administrador puede hacer esto.', 'permiso'); }

function listarUsuarios_() {
  return {ok: true, usuarios: usuarios_().map(function (u) { return {usuario: u.usuario, rol: u.rol, activo: u.activo}; })};
}

function guardarUsuario_(b) {
  var u = String(b.usuario || '').trim().toLowerCase(), c = String(b.clave || ''), rol = String(b.rol || '');
  if (u.length < 3) throw pub_('El usuario necesita al menos 3 letras.');
  if (ROLES.indexOf(rol) < 0) throw pub_('Elegí un rol: administrador, vendedor o depósito.');
  var existente = usuarios_().filter(function (x) { return x.usuario === u; })[0];
  if (!existente && c.length < 6) throw pub_('La clave necesita al menos 6 caracteres.');
  if (c && c.length < 6) throw pub_('La clave necesita al menos 6 caracteres.');
  var activo = b.activo !== false;
  var admins = usuarios_().filter(function (x) { return x.rol === 'admin' && x.activo && x.usuario !== u; }).length;
  if ((rol !== 'admin' || !activo) && admins === 0) throw pub_('Tiene que quedar al menos un administrador activo.');
  guardarUsuarioFila_(u, c, rol, activo);
  return listarUsuarios_();
}

function cambiarClave_(me, b) {
  var c = String(b.clave || '');
  if (c.length < 6) throw pub_('La clave necesita al menos 6 caracteres.');
  var us = usuarios_().filter(function (x) { return x.usuario === me.usuario; })[0];
  if (sha_(us.sal + String(b.actual || '')) !== us.hash) throw pub_('La clave actual no coincide.');
  guardarUsuarioFila_(me.usuario, c, us.rol, true);
  return {ok: true};
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Mantener el script activo ----------
   Apps Script tarda más en responder la primera vez después de un rato sin uso.
   Ejecutá «instalarActivador» UNA vez desde el editor (botón Ejecutar) para que se consulte solo cada 10 minutos.
   Para sacarlo, ejecutá «quitarActivador». */
function mantenerActivo() {
  asegurarHojas_();
  ['Estado', 'Stock', 'FacturasMano', 'DatosFact'].forEach(function (n) {
    var sh = ss_().getSheetByName(n), filas = sh.getLastRow() - 1;
    if (filas > 0 && !leerIndice_(sh, filas)) guardarIndice_(sh, sh.getRange(2, 1, filas, 1).getValues().map(function (r) { return r[0]; }));
  });
}
function instalarActivador() {
  quitarActivador();
  ScriptApp.newTrigger('mantenerActivo').timeBased().everyMinutes(10).create();
}
function quitarActivador() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'mantenerActivo') ScriptApp.deleteTrigger(t); });
}
