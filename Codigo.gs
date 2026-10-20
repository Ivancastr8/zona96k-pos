// ============================================================
// SISTEMA DE INVENTARIO - LICORERÍA
// Desarrollado con Google Apps Script
// ============================================================
//
// NOTA DE ESTA VERSIÓN:
// Se separó el precio de venta en dos canales:
//   - Precio Local      (columna H / índice 7)
//   - Precio Domicilio  (columna I / índice 8, NUEVA)
// Esto desplazó "Margen" (idx 9) y "Ganancia Potencial" (idx 10)
// una posición hacia la derecha respecto a la versión anterior.
//
// CORRECCIÓN: al eliminar/restar un ítem de una comanda que es una
// PRESENTACIÓN (six-pack, caja, etc.), ahora se devuelve al inventario
// la cantidad correcta de unidades base (paquetes × unidades por paquete).
// Antes no se devolvía nada porque el nombre de la presentación no existe
// como producto en Inventario. Ver resolverItemInventario().
// También se corrigió cerrarMesa(): los movimientos de presentaciones
// ahora se registran contra el producto base con su costo real.
//
// Layout de la hoja Inventario (0-indexado):
// 0 ID | 1 Producto | 2 Categoría | 3 Unidad | 4 Stock Actual
// 5 Stock Mínimo | 6 Precio Compra | 7 Precio Local
// 8 Precio Domicilio | 9 Margen | 10 Ganancia Potencial
// ============================================================

// ─────────────────────────────────────────
// CONSTANTES DE ESTADO
// ─────────────────────────────────────────
const ESTADOS_MESA = {
  LIBRE:   "LIBRE",
  OCUPADA: "OCUPADA"
};

const ESTADOS_COMANDA = {
  ABIERTA: "ABIERTA",
  CERRADA: "CERRADA"
};

const TIPOS_MOVIMIENTO = {
  ENTRADA:         "ENTRADA",
  SALIDA:          "SALIDA",
  ENTRADA_INICIAL: "ENTRADA INICIAL"
};

const TIPOS_ENTREGA = {
  LOCAL:     "local",
  DOMICILIO: "domicilio"
};

// ─────────────────────────────────────────
// CONFIGURACIÓN GLOBAL
// ─────────────────────────────────────────
const CONFIG = {
  CORREOS_ADMIN:        ["ivancastrouju20@gmail.com"],
  CORREOS_OPERARIO:     ["ivancastrouju20@gmail.com"],
  STOCK_MINIMO_DEFAULT: 10,
  CORREO_ALERTAS:       ["zona97k@gmail.com"],
  NOMBRE_NEGOCIO:       "Zona 96 K",
  NUMERO_MESAS:         10,
  ALERTA_MESA_MINUTOS:  120
};

const HOJAS = {
  INVENTARIO:     "Inventario",
  MOVIMIENTOS:    "Movimientos",
  REPORTES:       "Reportes",
  ALERTAS:        "Alertas",
  DASHBOARD:      "Dashboard",
  LOGS:           "Logs",
  CONFIG:         "Configuracion",
  MESAS:          "Mesas",
  COMANDAS:       "Comandas",
  PRESENTACIONES: "Presentaciones"
};
// ─────────────────────────────────────────
// PUNTO DE ENTRADA
// ─────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Interfaz")
    .setTitle("Inventario — " + CONFIG.NOMBRE_NEGOCIO)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function getLogoBase64() {
  var fileId = '18UYIBH_ljG8s5ZjzrYa9jy5yy5E-4Rwm';
  var blob = DriveApp.getFileById(fileId).getBlob();
  var base64Data = Utilities.base64Encode(blob.getBytes());
  var mimeType = blob.getContentType();
  return 'data:' + mimeType + ';base64,' + base64Data;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🍾 Inventario Licorería")
    .addItem("📦 Registrar Entrada",     "mostrarFormularioEntrada")
    .addItem("📤 Registrar Salida",      "mostrarFormularioSalida")
    .addSeparator()
    .addItem("📊 Generar Reporte",       "generarReporte")
    .addItem("📈 Actualizar Dashboard",  "actualizarDashboard")
    .addItem("🔔 Ver Alertas de Stock",  "verificarAlertas")
    .addSeparator()
    .addItem("➕ Agregar Producto",      "mostrarFormularioProducto")
    .addItem("✏️ Editar Precios",        "mostrarFormularioPrecios")
    .addSeparator()
    .addItem("🛠️ Crear hoja Presentaciones (six-pack, etc.)", "crearHojaPresentacionesDesdeMenu")
    .addSeparator()
    //.addItem("⚙️ Inicializar Sistema (⚠️ BORRA TODO — solo spreadsheets nuevos)", "inicializarSistema")
    .addToUi();

  verificarAcceso();
  verificarAlertas();
}

// ─────────────────────────────────────────
// SEGURIDAD
// ─────────────────────────────────────────
function obtenerUsuarioActual() {
  return Session.getActiveUser().getEmail();
}

function obtenerRolUsuario() {
  const correo = obtenerUsuarioActual();
  if (CONFIG.CORREOS_ADMIN.includes(correo))    return "ADMIN";
  if (CONFIG.CORREOS_OPERARIO.includes(correo)) return "OPERARIO";
  return null;
}

function verificarAcceso() {
  const rol = obtenerRolUsuario();
  if (!rol) throw new Error("Acceso denegado: " + obtenerUsuarioActual());
  return rol;
}

function verificarRolAdmin() {
  const rol = verificarAcceso();
  if (rol !== "ADMIN") throw new Error("Se requiere rol ADMIN");
}

// Wrapper seguro para el menú: crea la hoja Presentaciones SOLO si no existe
// todavía, para no borrar presentaciones ya creadas si se ejecuta 2 veces.
function crearHojaPresentacionesDesdeMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (ss.getSheetByName(HOJAS.PRESENTACIONES)) {
    ui.alert("✅ La hoja 'Presentaciones' ya existe. No se modificó nada.");
    return;
  }
  crearHojaPresentaciones(ss);
  ui.alert("✅ Hoja 'Presentaciones' creada. Ya puedes agregar six-packs desde la interfaz web.");
}

// ─────────────────────────────────────────
// DIAGNÓSTICO (SOLO LECTURA) — nunca modifica nada.
// Úsalo SIEMPRE antes de decidir si migrar o no.
// ─────────────────────────────────────────
function obtenerEncabezadosInventario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJAS.INVENTARIO);
  if (!hoja) return [];
  const ultimaCol = hoja.getLastColumn();
  if (ultimaCol < 1) return [];
  return hoja.getRange(1, 1, 1, ultimaCol).getValues()[0];
}

// Esquema NUEVO esperado por el código actual, en orden exacto.
const ESQUEMA_INVENTARIO_NUEVO = [
  "ID","Producto","Categoría","Unidad","Stock Actual","Stock Mínimo",
  "Precio Compra","Precio Local","Precio Domicilio","Margen","Ganancia Potencial"
];

// Comprueba las columnas H-K (índices 7-10) específicamente, que son las
// que la migración necesita insertar/renombrar. Si ya coinciden, no hay
// nada que hacer — sin importar si alguien ya las agregó a mano.
function esquemaInventarioYaMigrado() {
  const enc = obtenerEncabezadosInventario();
  return enc[7] === "Precio Local" && enc[8] === "Precio Domicilio" &&
         enc[9] === "Margen" && enc[10] === "Ganancia Potencial";
}

// ─────────────────────────────────────────
// INICIALIZACIÓN
// ─────────────────────────────────────────
function inicializarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    "⚙️ Inicializar Sistema (⚠️ BORRA TODO)",
    "Esto BORRARÁ el contenido de TODAS las hojas (Inventario, Movimientos, etc.) y las dejará en blanco.\n\n" +
    "Úsalo solo la primera vez que configuras el sistema. Si ya tienes datos y solo necesitas " +
    "agregar una hoja nueva (ej. Presentaciones), usa en cambio '🔧 Actualizar hojas faltantes'.\n\n" +
    "¿Deseas continuar y borrar todo?",
    ui.ButtonSet.YES_NO
  );
  if (respuesta !== ui.Button.YES) return;
  crearHojaInventario(ss);
  crearHojaMovimientos(ss);
  crearHojaReportes(ss);
  crearHojaAlertas(ss);
  crearHojaDashboard(ss);
  crearHojaLogs(ss);
  crearHojaMesas(ss);
  crearHojaComandas(ss);
  crearHojaPresentaciones(ss);
  ui.alert("✅ Sistema inicializado correctamente.");
}

function crearHojaInventario(ss) {
  let hoja = ss.getSheetByName(HOJAS.INVENTARIO);
  if (!hoja) hoja = ss.insertSheet(HOJAS.INVENTARIO);
  hoja.clearContents();
  const enc = ["ID","Producto","Categoría","Unidad","Stock Actual","Stock Mínimo","Precio Compra","Precio Local","Precio Domicilio","Margen","Ganancia Potencial"];
  hoja.getRange(1,1,1,enc.length).setValues([enc]);
  estilizarEncabezado(hoja, enc.length);
  const regla = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=E2<=F2")
    .setBackground("#FFCCCC")
    .setFontColor("#CC0000")
    .setRanges([hoja.getRange("A2:K1000")])
    .build();
  hoja.setConditionalFormatRules([regla]);
  hoja.setFrozenRows(1);
  hoja.setColumnWidth(2, 200);
}

function crearHojaMovimientos(ss) {
  let hoja = ss.getSheetByName(HOJAS.MOVIMIENTOS);
  if (!hoja) hoja = ss.insertSheet(HOJAS.MOVIMIENTOS);
  hoja.clearContents();
  const enc = ["Fecha y Hora","Tipo","Producto","Cantidad","Stock Anterior","Stock Nuevo","Precio Compra","Precio Venta","Valor Movimiento","Usuario","Notas"];
  hoja.getRange(1,1,1,enc.length).setValues([enc]);
  estilizarEncabezado(hoja, enc.length);
  hoja.setFrozenRows(1);
  hoja.setColumnWidth(1, 160);
  hoja.setColumnWidth(3, 200);
}

function crearHojaReportes(ss) {
  let hoja = ss.getSheetByName(HOJAS.REPORTES);
  if (!hoja) hoja = ss.insertSheet(HOJAS.REPORTES);
  hoja.clearContents();
  hoja.getRange("A1").setValue("Los reportes generados aparecerán aquí.");
}

function crearHojaAlertas(ss) {
  let hoja = ss.getSheetByName(HOJAS.ALERTAS);
  if (!hoja) hoja = ss.insertSheet(HOJAS.ALERTAS);
  hoja.clearContents();
  const enc = ["Producto","Stock Actual","Stock Mínimo","Unidades Faltantes","Estado"];
  hoja.getRange(1,1,1,enc.length).setValues([enc]);
  estilizarEncabezado(hoja, enc.length);
  hoja.setFrozenRows(1);
}

function crearHojaLogs(ss) {
  let hoja = ss.getSheetByName(HOJAS.LOGS);
  if (!hoja) hoja = ss.insertSheet(HOJAS.LOGS);
  hoja.clearContents();
  const enc = ["Fecha y Hora","Usuario","Correo","Acción","Detalle"];
  hoja.getRange(1,1,1,enc.length).setValues([enc]);
  estilizarEncabezado(hoja, enc.length);
  hoja.setFrozenRows(1);
  hoja.hideSheet();
}

function crearHojaMesas(ss) {
  let hoja = ss.getSheetByName(HOJAS.MESAS);
  if (!hoja) hoja = ss.insertSheet(HOJAS.MESAS);
  hoja.clearContents();
  const enc = ["Mesa","Estado","Hora Apertura","Total Acumulado"];
  hoja.getRange(1,1,1,enc.length).setValues([enc]);
  estilizarEncabezado(hoja, enc.length);
  const mesas = Array.from({ length: CONFIG.NUMERO_MESAS }, (_, i) => [`Mesa ${i+1}`, ESTADOS_MESA.LIBRE, "", 0]);
  hoja.getRange(2,1,mesas.length,4).setValues(mesas);
  hoja.setFrozenRows(1);
}

function crearHojaComandas(ss) {
  let hoja = ss.getSheetByName(HOJAS.COMANDAS);
  if (!hoja) hoja = ss.insertSheet(HOJAS.COMANDAS);
  hoja.clearContents();
  const enc = ["ID Comanda","Mesa","Fecha Apertura","Fecha Cierre","Producto","Cantidad","Precio Unitario","Subtotal","Estado","Método Pago","Total Mesa"];
  hoja.getRange(1,1,1,enc.length).setValues([enc]);
  estilizarEncabezado(hoja, enc.length);
  hoja.setFrozenRows(1);
  hoja.setColumnWidth(1, 140);
  hoja.setColumnWidth(5, 200);
}

// Presentaciones (six-pack, caja x24, etc.): NO tienen stock propio.
// Solo referencian un producto base por ID y definen cuántas unidades
// de ese producto consume cada paquete vendido.
function crearHojaPresentaciones(ss) {
  let hoja = ss.getSheetByName(HOJAS.PRESENTACIONES);
  if (!hoja) hoja = ss.insertSheet(HOJAS.PRESENTACIONES);
  hoja.clearContents();
  const enc = ["ID","Producto Base ID","Nombre Presentación","Unidades por Paquete","Precio Local","Precio Domicilio","Activo"];
  hoja.getRange(1,1,1,enc.length).setValues([enc]);
  estilizarEncabezado(hoja, enc.length);
  hoja.setFrozenRows(1);
  hoja.setColumnWidth(3, 200);
}

// ─────────────────────────────────────────
// UTILIDADES DE INVENTARIO
// ─────────────────────────────────────────
function construirMapaInventario() {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.INVENTARIO);
  const datos = hoja.getDataRange().getValues();
  const mapa  = {};
  for (let i = 1; i < datos.length; i++) {
    if (!datos[i][0]) continue;
    mapa[datos[i][1]] = {
      id:              datos[i][0],
      fila:            i + 1,
      stock:           datos[i][4],
      stockMinimo:     datos[i][5],
      precioCompra:    datos[i][6],
      precioLocal:     datos[i][7],
      precioDomicilio: datos[i][8]
    };
    mapa[`id_${datos[i][0]}`] = mapa[datos[i][1]];
  }
  return { hoja, mapa };
}

function precioSegunEntrega(precioLocal, precioDomicilio, tipoEntrega) {
  return tipoEntrega === TIPOS_ENTREGA.DOMICILIO ? precioDomicilio : precioLocal;
}

// Dado el nombre guardado en Comandas, devuelve la fila (0-indexada dentro de
// datosInv) del producto base en Inventario y cuántas unidades base consume
// cada unidad del ítem (1 para un producto suelto, N para una presentación x N).
//
// Los nombres de presentación se guardan en Comandas con el formato:
//   "<Producto base> - <Nombre presentación> (x<unidades>)"
// y ese nombre NO existe como producto en Inventario, por eso antes no se
// encontraba al eliminar/restar y el stock nunca se devolvía.
//
// Incluye presentaciones desactivadas a propósito: una mesa abierta todavía
// puede tener una presentación que luego se desactivó.
function resolverItemInventario(nombreItem, datosInv, datosPres) {
  // 1) Producto suelto
  for (let j = 1; j < datosInv.length; j++) {
    if (datosInv[j][1] === nombreItem) {
      return { filaInv: j, nombreBase: datosInv[j][1], unidadesPorItem: 1 };
    }
  }
  // 2) Presentación
  for (let p = 1; p < (datosPres || []).length; p++) {
    if (!datosPres[p][0]) continue;
    for (let j = 1; j < datosInv.length; j++) {
      if (datosInv[j][0] != datosPres[p][1]) continue;
      const esperado = `${datosInv[j][1]} - ${datosPres[p][2]} (x${datosPres[p][3]})`;
      if (esperado === nombreItem) {
        return { filaInv: j, nombreBase: datosInv[j][1], unidadesPorItem: Number(datosPres[p][3]) || 1 };
      }
    }
  }
  return null;
}

// Un número solo es "válido" para operar si es un number real y finito
// (descarta NaN, undefined, Infinity). Usado para blindar las celdas combinadas
// de Comandas contra líneas dañadas que de otro modo contaminarían las sumas.
function numeroValido(v) {
  return typeof v === "number" && isFinite(v);
}

// Evita que un nombre de producto/presentación con coma rompa el parseo de las
// celdas combinadas de Comandas (que usan ", " como separador entre ítems).
function sanearNombre(nombre) {
  return String(nombre || "").replace(/,/g, " ").replace(/\s+/g, " ").trim();
}

// ─────────────────────────────────────────
// MESAS  (siempre precio LOCAL — consumo en el negocio)
// ─────────────────────────────────────────
function obtenerEstadoMesas() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(HOJAS.MESAS);
  const datos = hoja.getDataRange().getValues();

  // El total mostrado se calcula directamente de la comanda ABIERTA (fuente de
  // verdad), no de la columna "Total Acumulado" de Mesas, que puede desincronizarse
  // si alguna operación no la actualiza correctamente (evita el caso "la tarjeta
  // dice $33.900 pero la comanda está vacía").
  const hojaComandas  = ss.getSheetByName(HOJAS.COMANDAS);
  const datosComandas = hojaComandas ? hojaComandas.getDataRange().getValues() : [];
  const totalesReales = {};
  for (let i = 1; i < datosComandas.length; i++) {
    if (datosComandas[i][8] !== ESTADOS_COMANDA.ABIERTA) continue;
    const mesa = datosComandas[i][1];
    const subtotal = Number(datosComandas[i][7]) || 0;
    totalesReales[mesa] = (totalesReales[mesa] || 0) + subtotal;
  }

  const ahora = new Date();
  const mesas = [];
  for (let i = 1; i < datos.length; i++) {
    const horaApertura   = datos[i][2] ? new Date(datos[i][2]) : null;
    const minutosAbierta = horaApertura ? Math.floor((ahora - horaApertura) / 60000) : 0;
    const nombre = datos[i][0];
    const totalReal = totalesReales.hasOwnProperty(nombre) ? totalesReales[nombre] : 0;
    mesas.push({
      nombre,
      estado:        datos[i][1],
      horaApertura:  horaApertura ? horaApertura.toLocaleTimeString("es-CO") : "",
      total:         datos[i][1] === ESTADOS_MESA.OCUPADA ? totalReal : 0,
      minutosAbierta,
      alerta:        datos[i][1] === ESTADOS_MESA.OCUPADA && minutosAbierta >= CONFIG.ALERTA_MESA_MINUTOS,
      fila:          i + 1
    });
  }
  return mesas;
}

function abrirMesa(nombreMesa) {
  verificarAcceso();
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.MESAS);
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === nombreMesa) {
      if (datos[i][1] === ESTADOS_MESA.OCUPADA)
        return { exito: false, mensaje: `${nombreMesa} ya está ocupada.` };
      hoja.getRange(i+1, 2).setValue(ESTADOS_MESA.OCUPADA);
      hoja.getRange(i+1, 3).setValue(new Date());
      hoja.getRange(i+1, 4).setValue(0);
      registrarLog("MESA ABIERTA", nombreMesa);
      return { exito: true };
    }
  }
  return { exito: false, mensaje: "Mesa no encontrada." };
}

// Agrega un ítem a la comanda ABIERTA de una mesa. Cada producto vive en SU
// PROPIA fila (nunca se combinan varios productos en una misma celda — ese
// diseño causaba "NaN" y mezclas entre productos distintos). Si el mismo
// producto ya tiene una fila abierta en esa mesa, se suma cantidad/subtotal
// en esa fila; si es nuevo, se crea una fila con el mismo ID Comanda.
function agregarItemAComandaMesa(hojaComandas, nombreMesa, nombreItem, cantidad, precioUnitario, subtotal) {
  if (!Number.isFinite(cantidad) || !Number.isFinite(precioUnitario) || !Number.isFinite(subtotal) || cantidad <= 0) {
    throw new Error(`Datos inválidos al agregar "${nombreItem}" a la comanda (cantidad, precio o subtotal no numéricos).`);
  }

  const datos = hojaComandas.getDataRange().getValues();
  let idComandaExistente = null;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] !== nombreMesa || datos[i][8] !== ESTADOS_COMANDA.ABIERTA) continue;
    idComandaExistente = datos[i][0];
    if (datos[i][4] === nombreItem) {
      const cantidadNueva = (Number(datos[i][5]) || 0) + cantidad;
      const subtotalNuevo = (Number(datos[i][7]) || 0) + subtotal;
      hojaComandas.getRange(i+1, 6).setValue(cantidadNueva);
      hojaComandas.getRange(i+1, 8).setValue(subtotalNuevo);
      return;
    }
  }
  const idComanda = idComandaExistente || `${nombreMesa.replace(" ","")}-${new Date().toISOString().slice(0,10)}-${new Date().getTime()}`;
  hojaComandas.appendRow([idComanda, nombreMesa, new Date(), "", nombreItem, cantidad, precioUnitario, subtotal, ESTADOS_COMANDA.ABIERTA, "", ""]);
}

function agregarProductoMesa(nombreMesa, idProducto, cantidad) {
  verificarAcceso();
  cantidad = Number(cantidad);
  if (!Number.isFinite(cantidad) || cantidad <= 0) return { exito: false, mensaje: "Cantidad inválida." };

  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInv     = ss.getSheetByName(HOJAS.INVENTARIO);
  const hojaComandas= ss.getSheetByName(HOJAS.COMANDAS);
  const hojaMesas   = ss.getSheetByName(HOJAS.MESAS);
  const datosInv    = hojaInv.getDataRange().getValues();

  let producto = null, filaProducto = -1;
  for (let i = 1; i < datosInv.length; i++) {
    if (datosInv[i][0] == idProducto) {
      producto = {
        nombre:          datosInv[i][1],
        stock:           datosInv[i][4],
        stockMinimo:     datosInv[i][5],
        precioCompra:    datosInv[i][6],
        precioLocal:     datosInv[i][7]   // Mesas siempre venden a precio LOCAL
      };
      filaProducto = i + 1;
      break;
    }
  }
  if (!producto)             return { exito: false, mensaje: "Producto no encontrado." };
  if (producto.stock < cantidad) return { exito: false, mensaje: `Stock insuficiente. Solo hay ${producto.stock} unidades.` };

  const stockNuevo = producto.stock - cantidad;
  hojaInv.getRange(filaProducto, 5).setValue(stockNuevo);
  hojaInv.getRange(filaProducto, 11).setValue((producto.precioLocal - producto.precioCompra) * stockNuevo);

  if (stockNuevo <= producto.stockMinimo) enviarAlertaEmail(producto.nombre, stockNuevo, producto.stockMinimo);

  const subtotal = cantidad * producto.precioLocal;
  agregarItemAComandaMesa(hojaComandas, nombreMesa, producto.nombre, cantidad, producto.precioLocal, subtotal);

  const datosMesas = hojaMesas.getDataRange().getValues();
  for (let i = 1; i < datosMesas.length; i++) {
    if (datosMesas[i][0] === nombreMesa) {
      hojaMesas.getRange(i+1, 4).setValue((datosMesas[i][3] || 0) + subtotal);
      break;
    }
  }

  registrarLog("PRODUCTO AGREGADO A MESA", `${nombreMesa} | ${producto.nombre} x${cantidad}`);
  return { exito: true, mensaje: `${producto.nombre} agregado a ${nombreMesa}.`, subtotal };
}

function obtenerComandaMesa(nombreMesa) {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.COMANDAS);
  const datos = hoja.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === nombreMesa && datos[i][8] === ESTADOS_COMANDA.ABIERTA) {
      items.push({
        fila:           i + 1,
        producto:       datos[i][4],
        cantidad:       Number(datos[i][5]) || 0,
        precioUnitario: Number(datos[i][6]) || 0,
        subtotal:       Number(datos[i][7]) || 0
      });
    }
  }
  return items;
}

// Libera una mesa que quedó OCUPADA sin ítems reales en su comanda (por ejemplo,
// si se eliminaron todos los productos uno por uno, o por una inconsistencia de
// datos). Verifica DE NUEVO contra la hoja antes de liberar, para que esto nunca
// se use como atajo para cerrar una cuenta con productos sin cobrar.
function liberarMesaSinComanda(nombreMesa) {
  verificarAcceso();
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const hojaComandas = ss.getSheetByName(HOJAS.COMANDAS);
  const hojaMesas    = ss.getSheetByName(HOJAS.MESAS);
  const datos = hojaComandas.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === nombreMesa && datos[i][8] === ESTADOS_COMANDA.ABIERTA) {
      const productos  = String(datos[i][4]).split(", ");
      const cantidades = String(datos[i][5]).split(", ").map(Number);
      const subtotales = String(datos[i][7]).split(", ").map(Number);
      // Solo bloquea la liberación si hay al menos UN ítem realmente cobrable
      // (nombre no vacío + cantidad y subtotal numéricos). Filas que solo tienen
      // datos dañados (ej. "NaN" de una corrupción anterior) no cuentan como
      // "tiene productos" — de lo contrario la mesa quedaría atrapada para siempre.
      const hayItemValido = productos.some((p, k) =>
        p.trim() !== "" && Number.isFinite(cantidades[k]) && cantidades[k] > 0 && Number.isFinite(subtotales[k])
      );
      if (hayItemValido) {
        return { exito: false, mensaje: "La mesa todavía tiene productos válidos en su comanda; ciérrala normalmente con un método de pago." };
      }
    }
  }

  const datosMesas = hojaMesas.getDataRange().getValues();
  for (let i = 1; i < datosMesas.length; i++) {
    if (datosMesas[i][0] === nombreMesa) {
      hojaMesas.getRange(i+1, 2).setValue(ESTADOS_MESA.LIBRE);
      hojaMesas.getRange(i+1, 3).setValue("");
      hojaMesas.getRange(i+1, 4).setValue(0);
      registrarLog("MESA LIBERADA SIN COMANDA", nombreMesa);
      return { exito: true, mensaje: `${nombreMesa} liberada.` };
    }
  }
  return { exito: false, mensaje: "Mesa no encontrada." };
}

function cerrarMesa(nombreMesa, pagos) {
  verificarAcceso();
  const ss              = SpreadsheetApp.getActiveSpreadsheet();
  const hojaComandas    = ss.getSheetByName(HOJAS.COMANDAS);
  const hojaMesas       = ss.getSheetByName(HOJAS.MESAS);
  const hojaInv         = ss.getSheetByName(HOJAS.INVENTARIO);
  const hojaPres        = ss.getSheetByName(HOJAS.PRESENTACIONES);

  const datos     = hojaComandas.getDataRange().getValues();
  const datosInv  = hojaInv.getDataRange().getValues();
  const datosPres = hojaPres ? hojaPres.getDataRange().getValues() : [];

  const itemsCerrar = [];
  let total = 0;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === nombreMesa && datos[i][8] === ESTADOS_COMANDA.ABIERTA) {
      const subtotal = Number(datos[i][7]) || 0;
      total += subtotal;
      itemsCerrar.push({
        fila:           i + 1,
        nombreProducto: datos[i][4],
        cantidad:       Number(datos[i][5]) || 0,
        precioVenta:    Number(datos[i][6]) || 0,
        subtotal
      });
    }
  }
  if (itemsCerrar.length === 0) return { exito: false, mensaje: "La comanda está vacía." };

  const norm = normalizarPagos(pagos, total);
  if (!norm.ok) return { exito: false, mensaje: norm.mensaje };

  itemsCerrar.forEach(item => {
    hojaComandas.getRange(item.fila, 4).setValue(new Date());
    hojaComandas.getRange(item.fila, 9).setValue(ESTADOS_COMANDA.CERRADA);
    hojaComandas.getRange(item.fila, 10).setValue(norm.texto);
    hojaComandas.getRange(item.fila, 11).setValue(total);
  });

  // Los movimientos se registran SIEMPRE contra el producto base y en unidades
  // base. Para una presentación (six-pack, etc.) se convierte: paquetes ×
  // unidades por paquete, y el precio de venta se expresa por unidad base, así
  // Dashboard y Reportes calculan ingresos/ganancia reales.
  itemsCerrar.forEach(item => {
    const res = resolverItemInventario(item.nombreProducto, datosInv, datosPres);
    let nombreMov = item.nombreProducto, precioCompra = 0, stockActual = 0;
    let unidades = item.cantidad, precioUnitarioMov = item.precioVenta, extra = "";

    if (res) {
      nombreMov         = res.nombreBase;
      precioCompra      = datosInv[res.filaInv][6];
      stockActual       = datosInv[res.filaInv][4];
      unidades          = item.cantidad * res.unidadesPorItem;
      precioUnitarioMov = item.precioVenta / res.unidadesPorItem;
      if (res.unidadesPorItem > 1) extra = ` | Presentación: ${item.nombreProducto} (${item.cantidad} paq.)`;
    }

    registrarMovimiento(
      TIPOS_MOVIMIENTO.SALIDA, nombreMov, unidades,
      stockActual + unidades, stockActual, precioCompra, precioUnitarioMov,
      `Mesa: ${nombreMesa} | Pago: ${norm.texto}${extra}`
    );
  });

  const datosMesas = hojaMesas.getDataRange().getValues();
  for (let i = 1; i < datosMesas.length; i++) {
    if (datosMesas[i][0] === nombreMesa) {
      hojaMesas.getRange(i+1, 2).setValue(ESTADOS_MESA.LIBRE);
      hojaMesas.getRange(i+1, 3).setValue("");
      hojaMesas.getRange(i+1, 4).setValue(0);
      break;
    }
  }

  actualizarHojaAlertas();
  registrarLog("MESA CERRADA", `${nombreMesa} | Total: $${total} | Pago: ${norm.texto}`);
  return { exito: true, total };
}

function eliminarItemComanda(nombreMesa, nombreProducto, cantidad, subtotal) {
  verificarAcceso();
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const hojaComandas = ss.getSheetByName(HOJAS.COMANDAS);
  const hojaInv      = ss.getSheetByName(HOJAS.INVENTARIO);
  const hojaMesas    = ss.getSheetByName(HOJAS.MESAS);
  const hojaPres     = ss.getSheetByName(HOJAS.PRESENTACIONES);

  const datos     = hojaComandas.getDataRange().getValues();
  const datosInv  = hojaInv.getDataRange().getValues();
  const datosPres = hojaPres ? hojaPres.getDataRange().getValues() : [];

  for (let i = 1; i < datos.length; i++) {
    if (
      datos[i][1] === nombreMesa        &&
      datos[i][4] === nombreProducto    &&
      datos[i][8] === ESTADOS_COMANDA.ABIERTA
    ) {
      // Resuelve tanto productos sueltos como presentaciones (six-pack, etc.)
      const res = resolverItemInventario(nombreProducto, datosInv, datosPres);

      const cantidadFila = Number(datos[i][5]) || 0;
      hojaComandas.deleteRow(i + 1);

      if (res) {
        const filaInv         = res.filaInv;
        const stockRestaurado = datosInv[filaInv][4] + cantidadFila * res.unidadesPorItem;
        const gananciaNueva   = (datosInv[filaInv][7] - datosInv[filaInv][6]) * stockRestaurado;
        hojaInv.getRange(filaInv + 1, 5).setValue(stockRestaurado);
        hojaInv.getRange(filaInv + 1, 11).setValue(gananciaNueva);
      }

      const datosMesas = hojaMesas.getDataRange().getValues();
      for (let k = 1; k < datosMesas.length; k++) {
        if (datosMesas[k][0] === nombreMesa) {
          hojaMesas.getRange(k+1, 4).setValue(Math.max(0, (datosMesas[k][3] || 0) - subtotal));
          break;
        }
      }

      registrarLog("ITEM ELIMINADO DE COMANDA", `${nombreMesa} | ${nombreProducto} x${cantidadFila}`);
      return { exito: true, mensaje: `${nombreProducto} eliminado de la comanda.` };
    }
  }
  return { exito: false, mensaje: "Item no encontrado en la comanda." };
}

// Resta una cantidad PARCIAL de un producto ya agregado a la comanda ABIERTA de una
// mesa (ej. el cliente pidió 7 Poker pero solo quiere pagar/consumir 3). Si la
// cantidad a restar iguala o supera lo que hay, el ítem se elimina por completo,
// igual que eliminarItemComanda. Soporta presentaciones (six-pack, etc.).
function restarCantidadComanda(nombreMesa, nombreProducto, cantidadARestar) {
  verificarAcceso();
  cantidadARestar = parseInt(cantidadARestar);
  if (!cantidadARestar || cantidadARestar <= 0) return { exito: false, mensaje: "Cantidad inválida." };

  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const hojaComandas = ss.getSheetByName(HOJAS.COMANDAS);
  const hojaInv      = ss.getSheetByName(HOJAS.INVENTARIO);
  const hojaMesas    = ss.getSheetByName(HOJAS.MESAS);
  const hojaPres     = ss.getSheetByName(HOJAS.PRESENTACIONES);

  const datos     = hojaComandas.getDataRange().getValues();
  const datosPres = hojaPres ? hojaPres.getDataRange().getValues() : [];

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] !== nombreMesa || datos[i][4] !== nombreProducto || datos[i][8] !== ESTADOS_COMANDA.ABIERTA) continue;

    const cantidadActual = Number(datos[i][5]) || 0;
    const subtotalActual = Number(datos[i][7]) || 0;
    if (!cantidadActual) return { exito: false, mensaje: "Item no encontrado en la comanda." };

    if (cantidadARestar > cantidadActual)
      return { exito: false, mensaje: `Solo hay ${cantidadActual} unidades de "${nombreProducto}" en la comanda.` };

    // Reparte el subtotal proporcionalmente (soporta el caso poco común de que el
    // precio haya cambiado a mitad de servicio y el subtotal acumulado no sea un
    // múltiplo exacto del precio unitario original).
    const precioPromedio       = subtotalActual / cantidadActual;
    const subtotalARestar      = Math.round(precioPromedio * cantidadARestar);
    const cantidadRestante     = cantidadActual - cantidadARestar;
    const subtotalRestante     = Math.max(0, subtotalActual - subtotalARestar);
    const eliminadoPorCompleto = cantidadRestante <= 0;

    const datosInv = hojaInv.getDataRange().getValues();
    const res = resolverItemInventario(nombreProducto, datosInv, datosPres);
    if (res) {
      const stockRestaurado = datosInv[res.filaInv][4] + cantidadARestar * res.unidadesPorItem;
      const gananciaNueva   = (datosInv[res.filaInv][7] - datosInv[res.filaInv][6]) * stockRestaurado;
      hojaInv.getRange(res.filaInv + 1, 5).setValue(stockRestaurado);
      hojaInv.getRange(res.filaInv + 1, 11).setValue(gananciaNueva);
    }

    if (eliminadoPorCompleto) {
      hojaComandas.deleteRow(i + 1);
    } else {
      hojaComandas.getRange(i+1, 6).setValue(cantidadRestante);
      hojaComandas.getRange(i+1, 8).setValue(subtotalRestante);
    }

    const datosMesas = hojaMesas.getDataRange().getValues();
    for (let k2 = 1; k2 < datosMesas.length; k2++) {
      if (datosMesas[k2][0] === nombreMesa) {
        hojaMesas.getRange(k2+1, 4).setValue(Math.max(0, (datosMesas[k2][3] || 0) - subtotalARestar));
        break;
      }
    }

    registrarLog("CANTIDAD RESTADA DE COMANDA", `${nombreMesa} | ${nombreProducto} -${cantidadARestar} (quedan ${eliminadoPorCompleto ? 0 : cantidadRestante})`);
    return {
      exito: true,
      mensaje: eliminadoPorCompleto
        ? `${nombreProducto} eliminado de la comanda.`
        : `Se restaron ${cantidadARestar} de ${nombreProducto}. Quedan ${cantidadRestante}.`,
      eliminado:        eliminadoPorCompleto,
      cantidadRestante: eliminadoPorCompleto ? 0 : cantidadRestante,
      subtotalRestante: eliminadoPorCompleto ? 0 : subtotalRestante
    };
  }
  return { exito: false, mensaje: "Item no encontrado en la comanda." };
}

function actualizarNumeroMesas(nuevoNumero) {
  verificarRolAdmin();
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJAS.MESAS);
  const datos = hoja.getDataRange().getValues();
  const mesasActuales = datos.length - 1;
  if (nuevoNumero === mesasActuales) return { exito: false, mensaje: "Ya tienes ese número de mesas." };
  if (nuevoNumero > mesasActuales) {
    for (let i = mesasActuales+1; i <= nuevoNumero; i++) hoja.appendRow([`Mesa ${i}`, ESTADOS_MESA.LIBRE, "", 0]);
  } else {
    for (let i = nuevoNumero+1; i <= mesasActuales; i++) {
      if (datos[i][1] === ESTADOS_MESA.OCUPADA)
        return { exito: false, mensaje: `Mesa ${i} está ocupada. Cierra la cuenta antes.` };
    }
    for (let i = mesasActuales; i > nuevoNumero; i--) hoja.deleteRow(i+1);
  }
  registrarLog("MESAS ACTUALIZADAS", `Total: ${nuevoNumero} mesas`);
  return { exito: true, mensaje: `Mesas actualizadas a ${nuevoNumero}.` };
}

// ─────────────────────────────────────────
// HISTORIAL Y CAJA
// ─────────────────────────────────────────
function obtenerHistorialComandas(filtroMesa, filtroFecha) {
  verificarAcceso();
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.COMANDAS);
  const datos = hoja.getDataRange().getValues();
  const hoy   = new Date();
  let fechaInicio;
  switch (filtroFecha) {
    case "hoy":    fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()); break;
    case "semana": fechaInicio = new Date(hoy.getTime() - 7*24*60*60*1000); break;
    case "mes":    fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth(), 1); break;
    default:       fechaInicio = new Date(2000, 0, 1);
  }

  const comandasMap = {};
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[0] || fila[8] !== ESTADOS_COMANDA.CERRADA) continue;
    const fechaCierre = fila[3] ? new Date(fila[3]) : null;
    if (!fechaCierre || fechaCierre < fechaInicio) continue;
    if (filtroMesa && filtroMesa !== "todas" && fila[1] !== filtroMesa) continue;

    const idComanda = fila[0];
    if (!comandasMap[idComanda]) {
      comandasMap[idComanda] = {
        id:           idComanda,
        mesa:         fila[1],
        fechaApertura:fila[2] ? new Date(fila[2]).toLocaleString("es-CO") : "--",
        fechaCierre:  fechaCierre.toLocaleString("es-CO"),
        metodoPago:   fila[9] || "--",
        total:        0,
        items:        []
      };
    }
    const subtotal = Number(fila[7]) || 0;
    comandasMap[idComanda].total += subtotal;
    comandasMap[idComanda].items.push({
      producto:      fila[4],
      cantidad:      Number(fila[5]) || 0,
      precioUnitario:Number(fila[6]) || 0,
      subtotal
    });
  }
  return Object.values(comandasMap).sort((a,b) => new Date(b.fechaCierre) - new Date(a.fechaCierre));
}

function obtenerResumenCaja(filtroFecha) {
  verificarAcceso();
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.COMANDAS);
  const datos = hoja.getDataRange().getValues();
  const hoy   = new Date();
  let fechaInicio;
  switch (filtroFecha) {
    case "hoy":    fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()); break;
    case "semana": fechaInicio = new Date(hoy.getTime() - 7*24*60*60*1000); break;
    case "mes":    fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth(), 1); break;
    default:       fechaInicio = new Date(2000, 0, 1);
  }

  const metodos = { Efectivo:0, Tarjeta:0, Transferencia:0, Nequi:0, Daviplata:0, "Bre-B":0 };
  let totalGeneral = 0, totalComandas = 0;

  const comandasMap = {};
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[0] || fila[8] !== ESTADOS_COMANDA.CERRADA) continue;
    const fechaCierre = fila[3] ? new Date(fila[3]) : null;
    if (!fechaCierre || fechaCierre < fechaInicio) continue;

    const idComanda = fila[0];
    if (!comandasMap[idComanda]) {
      comandasMap[idComanda] = { metodoPago: fila[9] || "Otro", total: 0 };
    }
    comandasMap[idComanda].total += Number(fila[7]) || 0;
  }

  Object.values(comandasMap).forEach(function(c) {
    totalGeneral += c.total;
    totalComandas++;
    parsearPagos(c.metodoPago, c.total).forEach(function(pago){
      if (metodos.hasOwnProperty(pago.metodo)) metodos[pago.metodo] += pago.monto;
      else metodos["Otro"] = (metodos["Otro"] || 0) + pago.monto;
    });
  });

  return {
    totalGeneral,
    totalComandas,
    porMetodo: Object.entries(metodos).map(([metodo, total]) => ({ metodo, total }))
  };
}

// ─────────────────────────────────────────
// PAGOS (uno o dos métodos)
// ─────────────────────────────────────────
function normalizarPagos(pagos, total) {
  total = Math.round(total || 0);
  var arr = [];
  if (typeof pagos === "string") {
    arr = [{ metodo: pagos || "Otro", monto: total }];
  } else if (Array.isArray(pagos)) {
    arr = pagos
      .map(function(p){ return { metodo: String(p.metodo || "Otro"), monto: Math.round(Number(p.monto) || 0) }; })
      .filter(function(p){ return p.monto > 0; });
  }
  if (!arr.length) return { ok:false, mensaje:"No se recibió información de pago." };
  var suma = arr.reduce(function(s,p){ return s + p.monto; }, 0);
  if (Math.abs(suma - total) > 1)
    return { ok:false, mensaje:"Los montos ($"+suma+") no coinciden con el total ($"+total+")." };
  return { ok:true, pagos: arr, texto: formatearPagos(arr) };
}

function formatearPagos(arr) {
  if (arr.length === 1) return arr[0].metodo;               // compatible con datos actuales
  return arr.map(function(p){ return p.metodo + ": " + p.monto; }).join(" | ");
}

// Reparte el total de una comanda entre sus métodos de pago (para la caja).
function parsearPagos(metodoPagoStr, totalComanda) {
  metodoPagoStr = String(metodoPagoStr || "Otro").trim();
  if (metodoPagoStr.indexOf(" | ") !== -1 || /: \d+$/.test(metodoPagoStr)) {
    return metodoPagoStr.split(" | ").map(function(par){
      var i = par.lastIndexOf(": ");
      if (i === -1) return { metodo: par.trim(), monto: 0 };
      return { metodo: par.slice(0, i).trim(), monto: Number(par.slice(i + 2)) || 0 };
    });
  }
  return [{ metodo: metodoPagoStr, monto: totalComanda }];
}

// ─────────────────────────────────────────
// PRODUCTOS
// ─────────────────────────────────────────
function mostrarFormularioProducto() {
  verificarRolAdmin();
  const html = HtmlService.createHtmlOutputFromFile("FormularioProducto")
    .setWidth(450).setHeight(560).setTitle("Agregar Producto");
  SpreadsheetApp.getUi().showModalDialog(html, "➕ Agregar Nuevo Producto");
}

function agregarProducto(datos) {
  verificarRolAdmin();
  if (!datos.nombre || !datos.nombre.trim())
    return { exito: false, mensaje: "El nombre del producto es requerido." };

  // Sanea el nombre ANTES de todo: una coma en el nombre rompería el parseo de
  // las celdas combinadas de Comandas ("Producto A, Producto B") más adelante.
  const nombreProducto = sanearNombre(datos.nombre);
  if (!nombreProducto) return { exito: false, mensaje: "El nombre del producto es requerido." };

  const precioCompra = parseFloat(datos.precioCompra);
  const precioLocal  = parseFloat(datos.precioLocal !== undefined ? datos.precioLocal : datos.precioVenta);
  let   precioDomicilio = parseFloat(datos.precioDomicilio);
  const stockInicial = parseInt(datos.stockInicial) || 0;
  const stockMinimo  = parseInt(datos.stockMinimo)  || CONFIG.STOCK_MINIMO_DEFAULT;

  if (isNaN(precioCompra) || precioCompra < 0)
    return { exito: false, mensaje: "Precio de compra inválido." };
  if (isNaN(precioLocal) || precioLocal < precioCompra)
    return { exito: false, mensaje: "El precio local debe ser mayor o igual al de compra." };

  // Si no se especifica precio de domicilio, se asume igual al local
  if (isNaN(precioDomicilio) || precioDomicilio <= 0) precioDomicilio = precioLocal;
  if (precioDomicilio < precioCompra)
    return { exito: false, mensaje: "El precio de domicilio no puede ser menor al precio de compra." };

  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.INVENTARIO);
  const filas = hoja.getDataRange().getValues();

  const nombreNormalizado = nombreProducto.toLowerCase();
  for (let i = 1; i < filas.length; i++) {
    if (filas[i][1] && filas[i][1].toString().trim().toLowerCase() === nombreNormalizado) {
      return {
        exito: false,
        mensaje: 'Ya existe "' + filas[i][1] + '" en el inventario. Para agregar stock usa "Registrar Entrada".'
      };
    }
  }

  // ID = máximo existente + 1. No se usa getLastRow() porque ahora se pueden
  // eliminar productos y eso repetiría IDs ya usados.
  let maxId = 0;
  for (let i = 1; i < filas.length; i++) {
    const idN = Number(filas[i][0]);
    if (idN > maxId) maxId = idN;
  }
  const nuevoId = maxId + 1;
  const margen  = precioLocal - precioCompra;

  hoja.appendRow([
    nuevoId, nombreProducto, datos.categoria || "",
    datos.unidad || "Unidad", stockInicial, stockMinimo,
    precioCompra, precioLocal, precioDomicilio, margen, margen * stockInicial
  ]);

  registrarLog("PRODUCTO AGREGADO", `${nombreProducto} | Stock inicial: ${stockInicial} | Local: $${precioLocal} | Domicilio: $${precioDomicilio}`);
  registrarMovimiento(
    TIPOS_MOVIMIENTO.ENTRADA_INICIAL, nombreProducto,
    stockInicial, 0, stockInicial, precioCompra, precioLocal,
    "Stock inicial al crear producto"
  );
  return { exito: true, mensaje: `Producto "${nombreProducto}" agregado correctamente.` };
}

// Elimina un producto del Inventario (solo ADMIN).
// - Bloquea si el producto (o alguna de sus presentaciones) está en una comanda
//   ABIERTA de alguna mesa, para no dejar cuentas con ítems huérfanos.
// - Desactiva las presentaciones (six-pack, etc.) que dependían de ese producto.
// - El historial (Movimientos, Comandas cerradas, Reportes) se conserva: ahí el
//   producto queda guardado por nombre.
function eliminarProducto(idProducto) {
  verificarRolAdmin();
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInv  = ss.getSheetByName(HOJAS.INVENTARIO);
  const datosInv = hojaInv.getDataRange().getValues();

  let fila = -1, nombre = "", stock = 0;
  for (let i = 1; i < datosInv.length; i++) {
    if (datosInv[i][0] == idProducto) {
      fila   = i + 1;
      nombre = datosInv[i][1];
      stock  = datosInv[i][4];
      break;
    }
  }
  if (fila === -1) return { exito: false, mensaje: "Producto no encontrado." };

  // Nombres con los que este producto puede aparecer en una comanda
  const nombresEnComanda = {};
  nombresEnComanda[nombre] = true;

  const hojaPres  = ss.getSheetByName(HOJAS.PRESENTACIONES);
  const datosPres = hojaPres ? hojaPres.getDataRange().getValues() : [];
  const filasPres = [];
  for (let p = 1; p < datosPres.length; p++) {
    if (!datosPres[p][0] || datosPres[p][1] != idProducto) continue;
    filasPres.push(p + 1);
    nombresEnComanda[`${nombre} - ${datosPres[p][2]} (x${datosPres[p][3]})`] = true;
  }

  const hojaComandas = ss.getSheetByName(HOJAS.COMANDAS);
  const datosCom     = hojaComandas ? hojaComandas.getDataRange().getValues() : [];
  for (let i = 1; i < datosCom.length; i++) {
    if (datosCom[i][8] === ESTADOS_COMANDA.ABIERTA && nombresEnComanda[datosCom[i][4]]) {
      return {
        exito: false,
        mensaje: `No se puede eliminar "${nombre}": está en la cuenta abierta de ${datosCom[i][1]}. Cierra la cuenta o quita el producto primero.`
      };
    }
  }

  // Todo validado: desactivar presentaciones y eliminar la fila del inventario
  filasPres.forEach(f => hojaPres.getRange(f, 7).setValue(false));
  hojaInv.deleteRow(fila);
  actualizarHojaAlertas();

  registrarLog("PRODUCTO ELIMINADO", `${nombre} | ID ${idProducto} | Stock al eliminar: ${stock} | Presentaciones desactivadas: ${filasPres.length}`);
  return {
    exito: true,
    mensaje: `Producto "${nombre}" eliminado.` +
      (filasPres.length ? ` Se desactivaron ${filasPres.length} presentación(es).` : "")
  };
}

function mostrarFormularioPrecios() {
  verificarRolAdmin();
  const html = HtmlService.createHtmlOutput(generarHtmlPrecios(obtenerListaProductos()))
    .setWidth(500).setHeight(560).setTitle("Editar Precios");
  SpreadsheetApp.getUi().showModalDialog(html, "✏️ Editar Precios");
}

function actualizarPrecio(idProducto, precioCompra, precioLocal, precioDomicilio) {
  verificarRolAdmin();

  precioCompra    = parseFloat(precioCompra);
  precioLocal     = parseFloat(precioLocal);
  precioDomicilio = parseFloat(precioDomicilio);

  if (isNaN(precioCompra) || precioCompra < 0)
    return { exito: false, mensaje: "Precio de compra inválido." };
  if (isNaN(precioLocal) || precioLocal < precioCompra)
    return { exito: false, mensaje: "El precio local debe ser mayor o igual al de compra." };
  if (isNaN(precioDomicilio) || precioDomicilio <= 0) precioDomicilio = precioLocal;
  if (precioDomicilio < precioCompra)
    return { exito: false, mensaje: "El precio de domicilio no puede ser menor al precio de compra." };

  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.INVENTARIO);
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == idProducto) {
      const margen = precioLocal - precioCompra;
      hoja.getRange(i+1, 7).setValue(precioCompra);
      hoja.getRange(i+1, 8).setValue(precioLocal);
      hoja.getRange(i+1, 9).setValue(precioDomicilio);
      hoja.getRange(i+1, 10).setValue(margen);
      hoja.getRange(i+1, 11).setValue(margen * datos[i][4]);
      registrarLog("PRECIO ACTUALIZADO", `Producto ID ${idProducto} | Compra: $${precioCompra} | Local: $${precioLocal} | Domicilio: $${precioDomicilio}`);
      return { exito: true };
    }
  }
  return { exito: false, mensaje: "Producto no encontrado." };
}

// ─────────────────────────────────────────
// PRESENTACIONES (six-pack, caja x24, etc.)
// No tienen stock propio: siempre descuentan del producto base.
// ─────────────────────────────────────────
function agregarPresentacion(datos) {
  verificarRolAdmin();

  const productoId = datos.productoId;
  const nombre = sanearNombre(datos.nombre) || "Presentación";
  const unidadesPorPaquete = parseInt(datos.unidadesPorPaquete);
  let precioLocal = parseFloat(datos.precioLocal);
  let precioDomicilio = parseFloat(datos.precioDomicilio);

  if (!productoId) return { exito: false, mensaje: "Selecciona el producto base." };
  if (!unidadesPorPaquete || unidadesPorPaquete < 2)
    return { exito: false, mensaje: "Las unidades por paquete deben ser 2 o más." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInv = ss.getSheetByName(HOJAS.INVENTARIO);
  const datosInv = hojaInv.getDataRange().getValues();

  let productoBase = null;
  for (let i = 1; i < datosInv.length; i++) {
    if (datosInv[i][0] == productoId) { productoBase = { nombre: datosInv[i][1], precioCompra: datosInv[i][6] }; break; }
  }
  if (!productoBase) return { exito: false, mensaje: "Producto base no encontrado." };

  const costoPaquete = productoBase.precioCompra * unidadesPorPaquete;
  if (isNaN(precioLocal) || precioLocal <= 0) return { exito: false, mensaje: "Precio local del paquete inválido." };
  if (isNaN(precioDomicilio) || precioDomicilio <= 0) precioDomicilio = precioLocal;
  if (precioLocal < costoPaquete)
    return { exito: false, mensaje: `El precio local ($${precioLocal}) es menor al costo del paquete ($${costoPaquete}).` };
  if (precioDomicilio < costoPaquete)
    return { exito: false, mensaje: `El precio de domicilio ($${precioDomicilio}) es menor al costo del paquete ($${costoPaquete}).` };

  const hojaPres = ss.getSheetByName(HOJAS.PRESENTACIONES);
  const nuevoId = hojaPres.getLastRow();
  hojaPres.appendRow([nuevoId, productoId, nombre, unidadesPorPaquete, precioLocal, precioDomicilio, true]);

  registrarLog("PRESENTACIÓN AGREGADA", `${nombre} x${unidadesPorPaquete} de "${productoBase.nombre}" | Local: $${precioLocal} | Domicilio: $${precioDomicilio}`);
  return { exito: true, mensaje: `Presentación "${nombre}" agregada para ${productoBase.nombre}.` };
}

function obtenerPresentaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPres = ss.getSheetByName(HOJAS.PRESENTACIONES);
  const hojaInv  = ss.getSheetByName(HOJAS.INVENTARIO);
  if (!hojaPres || !hojaInv) return [];

  const datosPres = hojaPres.getDataRange().getValues();
  const datosInv  = hojaInv.getDataRange().getValues();

  const mapaInv = {};
  for (let i = 1; i < datosInv.length; i++) {
    if (!datosInv[i][0]) continue;
    mapaInv[datosInv[i][0]] = {
      nombre: datosInv[i][1], stock: datosInv[i][4],
      precioCompra: datosInv[i][6], precioLocal: datosInv[i][7], precioDomicilio: datosInv[i][8]
    };
  }

  const out = [];
  for (let i = 1; i < datosPres.length; i++) {
    if (!datosPres[i][0]) continue;
    if (datosPres[i][6] === false) continue; // presentación desactivada
    const base = mapaInv[datosPres[i][1]];
    if (!base) continue;
    const unidadesPorPaquete = datosPres[i][3];
    const precioLocalPaquete     = datosPres[i][4];
    const precioDomicilioPaquete = datosPres[i][5];
    const costoPaquete = base.precioCompra * unidadesPorPaquete;

    out.push({
      id:                      datosPres[i][0],
      productoId:              datosPres[i][1],
      nombreProductoBase:      base.nombre,
      nombre:                  datosPres[i][2],
      unidadesPorPaquete,
      precioLocal:             precioLocalPaquete,
      precioDomicilio:         precioDomicilioPaquete,
      stockBaseActual:         base.stock,
      stockDisponiblePaquetes: Math.floor(base.stock / unidadesPorPaquete),
      costoPaquete,
      margenLocal:             precioLocalPaquete - costoPaquete,
      margenDomicilio:         precioDomicilioPaquete - costoPaquete,
      // Cuánto se ahorra el cliente comprando el paquete vs. las unidades sueltas al mismo canal
      ahorroLocal:             (base.precioLocal * unidadesPorPaquete) - precioLocalPaquete,
      ahorroDomicilio:         (base.precioDomicilio * unidadesPorPaquete) - precioDomicilioPaquete
    });
  }
  return out;
}

function eliminarPresentacion(idPresentacion) {
  verificarRolAdmin();
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.PRESENTACIONES);
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == idPresentacion) {
      hoja.getRange(i+1, 7).setValue(false);
      registrarLog("PRESENTACIÓN DESACTIVADA", `ID ${idPresentacion}`);
      return { exito: true, mensaje: "Presentación desactivada." };
    }
  }
  return { exito: false, mensaje: "Presentación no encontrada." };
}

// Actualiza unidades/precios de una presentación existente (six-pack, caja, etc.),
// con la misma validación de "no vender por debajo del costo" que al crearla.
function actualizarPresentacion(idPresentacion, unidadesPorPaquete, precioLocal, precioDomicilio) {
  verificarRolAdmin();
  unidadesPorPaquete = parseInt(unidadesPorPaquete);
  precioLocal = parseFloat(precioLocal);
  precioDomicilio = parseFloat(precioDomicilio);

  if (!unidadesPorPaquete || unidadesPorPaquete < 2)
    return { exito: false, mensaje: "Las unidades por paquete deben ser 2 o más." };
  if (isNaN(precioLocal) || precioLocal <= 0)
    return { exito: false, mensaje: "Precio local del paquete inválido." };

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPres = ss.getSheetByName(HOJAS.PRESENTACIONES);
  const hojaInv  = ss.getSheetByName(HOJAS.INVENTARIO);
  const datosPres = hojaPres.getDataRange().getValues();

  for (let i = 1; i < datosPres.length; i++) {
    if (datosPres[i][0] != idPresentacion) continue;

    const productoId = datosPres[i][1];
    const datosInv = hojaInv.getDataRange().getValues();
    let productoBase = null;
    for (let j = 1; j < datosInv.length; j++) {
      if (datosInv[j][0] == productoId) { productoBase = { nombre: datosInv[j][1], precioCompra: datosInv[j][6] }; break; }
    }
    if (!productoBase) return { exito: false, mensaje: "Producto base no encontrado." };

    if (isNaN(precioDomicilio) || precioDomicilio <= 0) precioDomicilio = precioLocal;
    const costoPaquete = productoBase.precioCompra * unidadesPorPaquete;
    if (precioLocal < costoPaquete)
      return { exito: false, mensaje: `El precio local ($${precioLocal}) es menor al costo del paquete ($${costoPaquete}).` };
    if (precioDomicilio < costoPaquete)
      return { exito: false, mensaje: `El precio de domicilio ($${precioDomicilio}) es menor al costo del paquete ($${costoPaquete}).` };

    hojaPres.getRange(i+1, 4).setValue(unidadesPorPaquete);
    hojaPres.getRange(i+1, 5).setValue(precioLocal);
    hojaPres.getRange(i+1, 6).setValue(precioDomicilio);

    registrarLog("PRESENTACIÓN ACTUALIZADA", `ID ${idPresentacion} | ${productoBase.nombre} | Unidades: ${unidadesPorPaquete} | Local: $${precioLocal} | Domicilio: $${precioDomicilio}`);
    return { exito: true, mensaje: "Precios de la presentación actualizados." };
  }
  return { exito: false, mensaje: "Presentación no encontrada." };
}

// Venta directa de una presentación: descuenta unidades base = paquetes * unidadesPorPaquete
function registrarSalidaPresentacion(idPresentacion, cantidadPaquetes, metodoPago, tipoEntrega, notas) {
  verificarAcceso();
  if (!metodoPago) return { exito: false, mensaje: "Selecciona un método de pago." };
  cantidadPaquetes = parseInt(cantidadPaquetes);
  if (!cantidadPaquetes || cantidadPaquetes <= 0) return { exito: false, mensaje: "Cantidad inválida." };
  tipoEntrega = (tipoEntrega === TIPOS_ENTREGA.DOMICILIO) ? TIPOS_ENTREGA.DOMICILIO : TIPOS_ENTREGA.LOCAL;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPres = ss.getSheetByName(HOJAS.PRESENTACIONES);
  const hojaInv  = ss.getSheetByName(HOJAS.INVENTARIO);
  const datosPres = hojaPres.getDataRange().getValues();

  let presentacion = null;
  for (let i = 1; i < datosPres.length; i++) {
    if (datosPres[i][0] == idPresentacion) {
      presentacion = {
        productoId: datosPres[i][1], nombre: datosPres[i][2],
        unidadesPorPaquete: datosPres[i][3], precioLocal: datosPres[i][4], precioDomicilio: datosPres[i][5]
      };
      break;
    }
  }
  if (!presentacion) return { exito: false, mensaje: "Presentación no encontrada." };

  const datosInv = hojaInv.getDataRange().getValues();
  for (let i = 1; i < datosInv.length; i++) {
    if (datosInv[i][0] == presentacion.productoId) {
      const nombreProducto = datosInv[i][1];
      const stockAnterior  = datosInv[i][4];
      const stockMinimo    = datosInv[i][5];
      const pC     = datosInv[i][6];
      const pLocal = datosInv[i][7];
      const unidadesADescontar = cantidadPaquetes * presentacion.unidadesPorPaquete;

      if (unidadesADescontar > stockAnterior) {
        const paquetesDisponibles = Math.floor(stockAnterior / presentacion.unidadesPorPaquete);
        return { exito: false, mensaje: `Stock insuficiente. Solo hay ${stockAnterior} unidades (${paquetesDisponibles} paquete(s) de ${presentacion.unidadesPorPaquete}).` };
      }

      const stockNuevo      = stockAnterior - unidadesADescontar;
      const precioAplicado  = precioSegunEntrega(presentacion.precioLocal, presentacion.precioDomicilio, tipoEntrega);
      const subtotal         = cantidadPaquetes * precioAplicado;
      const etiquetaEntrega  = tipoEntrega === TIPOS_ENTREGA.DOMICILIO ? "Domicilio" : "Local";
      const nombreItem       = `${nombreProducto} - ${presentacion.nombre} (x${presentacion.unidadesPorPaquete})`;
      // Precio "por unidad base" equivalente, para que el Dashboard (que agrupa
      // por Movimientos) calcule ingresos/ganancia reales sin lógica especial.
      const precioUnitarioEquivalente = precioAplicado / presentacion.unidadesPorPaquete;

      hojaInv.getRange(i+1, 5).setValue(stockNuevo);
      hojaInv.getRange(i+1, 11).setValue((pLocal - pC) * stockNuevo);

      registrarMovimiento(
        TIPOS_MOVIMIENTO.SALIDA, nombreProducto, unidadesADescontar, stockAnterior, stockNuevo, pC, precioUnitarioEquivalente,
        `Pago: ${metodoPago} | Entrega: ${etiquetaEntrega} | Presentación: ${presentacion.nombre} x${presentacion.unidadesPorPaquete} (${cantidadPaquetes} paquete(s)) | ${notas || ""}`
      );

      const hojaComandas = ss.getSheetByName(HOJAS.COMANDAS);
      const ahora     = new Date();
      const idComanda = `DIR-${ahora.toISOString().slice(0,10)}-${ahora.getTime()}`;
      hojaComandas.appendRow([
        idComanda,
        etiquetaEntrega === "Domicilio" ? "Salida Directa (Domicilio)" : "Salida Directa",
        ahora, ahora, nombreItem, cantidadPaquetes, precioAplicado, subtotal,
        ESTADOS_COMANDA.CERRADA, metodoPago, subtotal
      ]);

      registrarLog(
        TIPOS_MOVIMIENTO.SALIDA,
        `${nombreItem} | Paquetes: ${cantidadPaquetes} | Unidades: ${unidadesADescontar} | Stock: ${stockAnterior} → ${stockNuevo} | Pago: ${metodoPago} | Entrega: ${etiquetaEntrega}`
      );
      if (stockNuevo <= stockMinimo) enviarAlertaEmail(nombreProducto, stockNuevo, stockMinimo);
      actualizarHojaAlertas();

      return {
        exito: true,
        mensaje: `Salida registrada. Nuevo stock de ${nombreProducto}: ${stockNuevo} unidades.`,
        subtotal, unidadesADescontar, precioAplicado, tipoEntrega, producto: nombreProducto
      };
    }
  }
  return { exito: false, mensaje: "Producto base no encontrado." };
}

// Venta directa con carrito: acepta unidades y presentaciones mezcladas,
// uno o dos métodos de pago, y un solo canal (local/domicilio) por venta.
// Valida TODO (stock acumulado + montos) antes de escribir nada, así una
// venta nunca queda a medias.
function registrarVentaDirecta(items, pagos, tipoEntrega, notas) {
  verificarAcceso();
  if (!Array.isArray(items) || !items.length) return { exito:false, mensaje:"El carrito está vacío." };
  tipoEntrega = (tipoEntrega === TIPOS_ENTREGA.DOMICILIO) ? TIPOS_ENTREGA.DOMICILIO : TIPOS_ENTREGA.LOCAL;

  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInv      = ss.getSheetByName(HOJAS.INVENTARIO);
  const hojaPres     = ss.getSheetByName(HOJAS.PRESENTACIONES);
  const hojaComandas = ss.getSheetByName(HOJAS.COMANDAS);
  const datosInv     = hojaInv.getDataRange().getValues();
  const datosPres    = hojaPres ? hojaPres.getDataRange().getValues() : [];

  const invPorId = {};
  for (let i = 1; i < datosInv.length; i++) {
    if (datosInv[i][0] === "" || datosInv[i][0] === null) continue;
    invPorId[datosInv[i][0]] = {
      fila:i+1, nombre:datosInv[i][1], stock:datosInv[i][4], stockMinimo:datosInv[i][5],
      precioCompra:datosInv[i][6], precioLocal:datosInv[i][7], precioDomicilio:datosInv[i][8]
    };
  }
  const presPorId = {};
  for (let i = 1; i < datosPres.length; i++) {
    if (datosPres[i][0] === "" || datosPres[i][0] === null) continue;
    presPorId[datosPres[i][0]] = {
      productoId:datosPres[i][1], nombre:datosPres[i][2], unidadesPorPaquete:datosPres[i][3],
      precioLocal:datosPres[i][4], precioDomicilio:datosPres[i][5]
    };
  }

  // 1) Resolver líneas y acumular consumo por producto base
  const lineas = [];
  const consumoBase = {};
  for (const it of items) {
    const cantidad = parseInt(it.cantidad);
    if (!cantidad || cantidad <= 0) return { exito:false, mensaje:"Cantidad inválida en el carrito." };

    if (it.tipo === "presentacion") {
      const pres = presPorId[it.id];
      if (!pres) return { exito:false, mensaje:"Presentación no encontrada." };
      const base = invPorId[pres.productoId];
      if (!base) return { exito:false, mensaje:"Producto base no encontrado." };
      const unidades = cantidad * pres.unidadesPorPaquete;
      const precioPaquete = precioSegunEntrega(pres.precioLocal, pres.precioDomicilio, tipoEntrega);
      consumoBase[pres.productoId] = (consumoBase[pres.productoId] || 0) + unidades;
      lineas.push({
        tipo:"presentacion", baseId:pres.productoId, nombreBase:base.nombre,
        nombreItem: base.nombre + " - " + pres.nombre + " (x" + pres.unidadesPorPaquete + ")",
        cantidadPaquetes:cantidad, unidadesBase:unidades, precioCompra:base.precioCompra,
        precioPaquete:precioPaquete, precioUnitarioEquivalente: precioPaquete / pres.unidadesPorPaquete,
        presNombre: pres.nombre, unidadesPorPaquete: pres.unidadesPorPaquete
      });
    } else {
      const base = invPorId[it.id];
      if (!base) return { exito:false, mensaje:"Producto no encontrado." };
      const precio = precioSegunEntrega(base.precioLocal, base.precioDomicilio, tipoEntrega);
      consumoBase[it.id] = (consumoBase[it.id] || 0) + cantidad;
      lineas.push({
        tipo:"unidad", baseId:it.id, nombreBase:base.nombre, nombreItem:base.nombre,
        cantidad:cantidad, unidadesBase:cantidad, precioCompra:base.precioCompra, precioUnitario:precio
      });
    }
  }

  // 2) Validar stock ANTES de aplicar cambios
  for (const id in consumoBase) {
    const base = invPorId[id];
    if (consumoBase[id] > base.stock)
      return { exito:false, mensaje:`Stock insuficiente de "${base.nombre}". Necesitas ${consumoBase[id]} y hay ${base.stock}.` };
  }

  // 3) Total y validación de pagos
  let total = 0;
  lineas.forEach(l => { total += (l.tipo==="presentacion") ? l.cantidadPaquetes*l.precioPaquete : l.cantidad*l.precioUnitario; });
  total = Math.round(total);
  const norm = normalizarPagos(pagos, total);
  if (!norm.ok) return { exito:false, mensaje: norm.mensaje };

  // 4) Aplicar: descontar stock, registrar movimientos y comandas
  const ahora     = new Date();
  const idComanda = `DIR-${ahora.toISOString().slice(0,10)}-${ahora.getTime()}`;
  const etiqueta  = tipoEntrega === TIPOS_ENTREGA.DOMICILIO ? "Domicilio" : "Local";
  const stockTrabajo = {};
  for (const id in consumoBase) stockTrabajo[id] = invPorId[id].stock;

  lineas.forEach(l => {
    const base = invPorId[l.baseId];
    const stockAnterior = stockTrabajo[l.baseId];
    const stockNuevo    = stockAnterior - l.unidadesBase;
    stockTrabajo[l.baseId] = stockNuevo;

    hojaInv.getRange(base.fila, 5).setValue(stockNuevo);
    hojaInv.getRange(base.fila, 11).setValue((base.precioLocal - base.precioCompra) * stockNuevo);

    if (l.tipo === "presentacion") {
      registrarMovimiento(TIPOS_MOVIMIENTO.SALIDA, l.nombreBase, l.unidadesBase, stockAnterior, stockNuevo,
        l.precioCompra, l.precioUnitarioEquivalente,
        `Pago: ${norm.texto} | Entrega: ${etiqueta} | Presentación: ${l.presNombre} x${l.unidadesPorPaquete} (${l.cantidadPaquetes} paq.) | ${notas||""}`);
      hojaComandas.appendRow([idComanda, etiqueta==="Domicilio"?"Salida Directa (Domicilio)":"Salida Directa",
        ahora, ahora, l.nombreItem, l.cantidadPaquetes, l.precioPaquete, l.cantidadPaquetes*l.precioPaquete,
        ESTADOS_COMANDA.CERRADA, norm.texto, total]);
    } else {
      registrarMovimiento(TIPOS_MOVIMIENTO.SALIDA, l.nombreBase, l.cantidad, stockAnterior, stockNuevo,
        l.precioCompra, l.precioUnitario,
        `Pago: ${norm.texto} | Entrega: ${etiqueta} | ${notas||""}`);
      hojaComandas.appendRow([idComanda, etiqueta==="Domicilio"?"Salida Directa (Domicilio)":"Salida Directa",
        ahora, ahora, l.nombreItem, l.cantidad, l.precioUnitario, l.cantidad*l.precioUnitario,
        ESTADOS_COMANDA.CERRADA, norm.texto, total]);
    }
  });

  for (const id in stockTrabajo) {
    const base = invPorId[id];
    if (stockTrabajo[id] <= base.stockMinimo) enviarAlertaEmail(base.nombre, stockTrabajo[id], base.stockMinimo);
  }
  actualizarHojaAlertas();

  registrarLog("VENTA DIRECTA", `Ítems: ${lineas.length} | Total: $${total} | Pago: ${norm.texto} | Entrega: ${etiqueta}`);
  return { exito:true, total, mensaje:`Venta registrada por $${total}.` };
}

// Venta de una presentación en Mesas: siempre a precio LOCAL (dine-in)
function agregarPresentacionAMesa(nombreMesa, idPresentacion, cantidadPaquetes) {
  verificarAcceso();
  cantidadPaquetes = parseInt(cantidadPaquetes);
  if (!cantidadPaquetes || cantidadPaquetes < 1) return { exito: false, mensaje: "Cantidad inválida." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPres     = ss.getSheetByName(HOJAS.PRESENTACIONES);
  const hojaInv      = ss.getSheetByName(HOJAS.INVENTARIO);
  const hojaComandas = ss.getSheetByName(HOJAS.COMANDAS);
  const hojaMesas    = ss.getSheetByName(HOJAS.MESAS);

  const datosPres = hojaPres.getDataRange().getValues();
  let presentacion = null;
  for (let i = 1; i < datosPres.length; i++) {
    if (datosPres[i][0] == idPresentacion) {
      presentacion = { productoId: datosPres[i][1], nombre: datosPres[i][2], unidadesPorPaquete: datosPres[i][3], precioLocal: datosPres[i][4] };
      break;
    }
  }
  if (!presentacion) return { exito: false, mensaje: "Presentación no encontrada." };

  const datosInv = hojaInv.getDataRange().getValues();
  for (let i = 1; i < datosInv.length; i++) {
    if (datosInv[i][0] == presentacion.productoId) {
      const nombreProducto = datosInv[i][1];
      const stockAnterior  = datosInv[i][4];
      const stockMinimo    = datosInv[i][5];
      const pC = datosInv[i][6];
      const unidadesADescontar = cantidadPaquetes * presentacion.unidadesPorPaquete;

      if (unidadesADescontar > stockAnterior) {
        const paquetesDisponibles = Math.floor(stockAnterior / presentacion.unidadesPorPaquete);
        return { exito: false, mensaje: `Stock insuficiente. Solo hay ${stockAnterior} unidades (${paquetesDisponibles} paquete(s)).` };
      }

      const stockNuevo = stockAnterior - unidadesADescontar;
      const precioUnitarioEquivalente = presentacion.precioLocal / presentacion.unidadesPorPaquete;
      hojaInv.getRange(i+1, 5).setValue(stockNuevo);
      hojaInv.getRange(i+1, 11).setValue((precioUnitarioEquivalente - pC) * stockNuevo);

      if (stockNuevo <= stockMinimo) enviarAlertaEmail(nombreProducto, stockNuevo, stockMinimo);

      const nombreItem = `${nombreProducto} - ${presentacion.nombre} (x${presentacion.unidadesPorPaquete})`;
      const subtotal    = cantidadPaquetes * presentacion.precioLocal;
      agregarItemAComandaMesa(hojaComandas, nombreMesa, nombreItem, cantidadPaquetes, presentacion.precioLocal, subtotal);

      const datosMesas = hojaMesas.getDataRange().getValues();
      for (let k = 1; k < datosMesas.length; k++) {
        if (datosMesas[k][0] === nombreMesa) { hojaMesas.getRange(k+1, 4).setValue((datosMesas[k][3] || 0) + subtotal); break; }
      }

      registrarLog("PRESENTACIÓN AGREGADA A MESA", `${nombreMesa} | ${nombreItem} x${cantidadPaquetes}`);
      return { exito: true, mensaje: `${nombreItem} agregado a ${nombreMesa}.`, subtotal };
    }
  }
  return { exito: false, mensaje: "Producto base no encontrado." };
}

// ─────────────────────────────────────────
// MOVIMIENTOS
// ─────────────────────────────────────────
function mostrarFormularioEntrada() {
  verificarAcceso();
  const template = HtmlService.createTemplateFromFile("FormularioMovimiento");
  template.tipoMovimiento = "ENTRADA";
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(430).setHeight(430), "📦 Registrar Entrada");
}

function mostrarFormularioSalida() {
  verificarAcceso();
  const template = HtmlService.createTemplateFromFile("FormularioMovimiento");
  template.tipoMovimiento = "SALIDA";
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(430).setHeight(430), "📤 Registrar Salida");
}

function registrarEntrada(idProducto, cantidad, notas) {
  verificarAcceso();
  return procesarEntrada(idProducto, cantidad, notas);
}

// tipoEntrega: "local" (default) o "domicilio"
function registrarSalida(idProducto, cantidad, metodoPago, tipoEntrega, notas) {
  verificarAcceso();
  if (!metodoPago) return { exito: false, mensaje: "Selecciona un método de pago." };
  tipoEntrega = (tipoEntrega === TIPOS_ENTREGA.DOMICILIO) ? TIPOS_ENTREGA.DOMICILIO : TIPOS_ENTREGA.LOCAL;
  return procesarSalida(idProducto, cantidad, metodoPago, tipoEntrega, notas);
}

function procesarEntrada(idProducto, cantidad, notas) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(HOJAS.INVENTARIO);
  const datos = hoja.getDataRange().getValues();

  if (!cantidad || cantidad <= 0) return { exito: false, mensaje: "Cantidad inválida." };

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == idProducto) {
      const stockAnterior = datos[i][4];
      const stockNuevo    = stockAnterior + cantidad;
      const nombreProducto= datos[i][1];
      const pC = datos[i][6], pLocal = datos[i][7];

      hoja.getRange(i+1, 5).setValue(stockNuevo);
      hoja.getRange(i+1, 11).setValue((pLocal - pC) * stockNuevo);

      registrarMovimiento(TIPOS_MOVIMIENTO.ENTRADA, nombreProducto, cantidad, stockAnterior, stockNuevo, pC, pLocal, notas || "");
      registrarLog(TIPOS_MOVIMIENTO.ENTRADA, `${nombreProducto} | Cantidad: ${cantidad} | Stock: ${stockAnterior} → ${stockNuevo}`);

      return { exito: true, mensaje: `Entrada registrada. Nuevo stock: ${stockNuevo} unidades.`, producto: nombreProducto };
    }
  }
  return { exito: false, mensaje: "Producto no encontrado." };
}

function procesarSalida(idProducto, cantidad, metodoPago, tipoEntrega, notas) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hoja  = ss.getSheetByName(HOJAS.INVENTARIO);
  const datos = hoja.getDataRange().getValues();

  if (!cantidad || cantidad <= 0) return { exito: false, mensaje: "Cantidad inválida." };

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == idProducto) {
      const stockAnterior = datos[i][4];
      if (cantidad > stockAnterior)
        return { exito: false, mensaje: `Stock insuficiente. Solo hay ${stockAnterior} unidades.` };

      const stockNuevo     = stockAnterior - cantidad;
      const nombreProducto = datos[i][1];
      const pC          = datos[i][6];
      const pLocal      = datos[i][7];
      const pDomicilio  = datos[i][8];
      const precioAplicado = precioSegunEntrega(pLocal, pDomicilio, tipoEntrega);
      const subtotal        = cantidad * precioAplicado;
      const etiquetaEntrega  = tipoEntrega === TIPOS_ENTREGA.DOMICILIO ? "Domicilio" : "Local";

      hoja.getRange(i+1, 5).setValue(stockNuevo);
      // Ganancia potencial del inventario restante se valora siempre a precio local
      hoja.getRange(i+1, 11).setValue((pLocal - pC) * stockNuevo);

      registrarMovimiento(
        TIPOS_MOVIMIENTO.SALIDA, nombreProducto, cantidad, stockAnterior, stockNuevo, pC, precioAplicado,
        `Pago: ${metodoPago} | Entrega: ${etiquetaEntrega} | ${notas || ""}`
      );

      // Se registra también en Comandas (mismo mecanismo que las mesas) para
      // que alimente el resumen de caja y el historial.
      const hojaComandas = ss.getSheetByName(HOJAS.COMANDAS);
      const ahora        = new Date();
      const idComanda    = `DIR-${ahora.toISOString().slice(0,10)}-${ahora.getTime()}`;
      hojaComandas.appendRow([
        idComanda,
        etiquetaEntrega === "Domicilio" ? "Salida Directa (Domicilio)" : "Salida Directa",
        ahora, ahora, nombreProducto, cantidad, precioAplicado, subtotal,
        ESTADOS_COMANDA.CERRADA, metodoPago, subtotal
      ]);

      registrarLog(
        TIPOS_MOVIMIENTO.SALIDA,
        `${nombreProducto} | Cantidad: ${cantidad} | Stock: ${stockAnterior} → ${stockNuevo} | Pago: ${metodoPago} | Entrega: ${etiquetaEntrega}`
      );
      if (stockNuevo <= datos[i][5]) enviarAlertaEmail(nombreProducto, stockNuevo, datos[i][5]);
      actualizarHojaAlertas();

      return {
        exito:    true,
        mensaje:  `Salida registrada. Nuevo stock: ${stockNuevo} unidades.`,
        subtotal,
        producto: nombreProducto,
        precioAplicado,
        tipoEntrega
      };
    }
  }
  return { exito: false, mensaje: "Producto no encontrado." };
}

function registrarMovimiento(tipo, producto, cantidad, stockAnterior, stockNuevo, pC, pV, notas) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.MOVIMIENTOS);
  hoja.appendRow([
    new Date(), tipo, producto, cantidad, stockAnterior, stockNuevo, pC, pV,
    tipo === TIPOS_MOVIMIENTO.SALIDA ? cantidad * pV : cantidad * pC,
    obtenerUsuarioActual(), notas || ""
  ]);
}

// ─────────────────────────────────────────
// ALERTAS DE STOCK
// ─────────────────────────────────────────
function verificarAlertas() {
  actualizarHojaAlertas();
  const alertas = obtenerProductosBajoStock();
  if (alertas.length === 0) {
    SpreadsheetApp.getUi().alert("✅ Stock en Orden", "Todos los productos están sobre el stock mínimo.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  let mensaje = "⚠️ Los siguientes productos necesitan reposición:\n\n";
  alertas.forEach(p => { mensaje += `• ${p.nombre}: ${p.stock} unidades (mínimo: ${p.minimo})\n`; });
  SpreadsheetApp.getUi().alert("🔔 Alerta de Stock Bajo", mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
}

function actualizarHojaAlertas() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInv = ss.getSheetByName(HOJAS.INVENTARIO);
  const hojaAl  = ss.getSheetByName(HOJAS.ALERTAS);
  if (!hojaAl) return;
  const uf = hojaAl.getLastRow();
  if (uf > 1) hojaAl.getRange(2, 1, uf-1, 5).clearContent();
  const datos   = hojaInv.getDataRange().getValues();
  const alertas = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][4] <= datos[i][5])
      alertas.push([datos[i][1], datos[i][4], datos[i][5], datos[i][5]-datos[i][4],
        datos[i][4] === 0 ? "🔴 AGOTADO" : "🟡 STOCK BAJO"]);
  }
  if (alertas.length > 0) hojaAl.getRange(2, 1, alertas.length, 5).setValues(alertas);
}

function obtenerProductosBajoStock() {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.INVENTARIO);
  const datos = hoja.getDataRange().getValues();
  const alertas = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][4] <= datos[i][5])
      alertas.push({ nombre: datos[i][1], stock: datos[i][4], minimo: datos[i][5] });
  }
  return alertas;
}

function enviarAlertaEmail(producto, stockActual, stockMinimo) {
  try {
    MailApp.sendEmail({
      to:      CONFIG.CORREO_ALERTAS,
      subject: `⚠️ Stock Bajo - ${CONFIG.NOMBRE_NEGOCIO}`,
      body:    `Alerta de inventario:\n\nProducto: ${producto}\nStock actual: ${stockActual} unidades\nStock mínimo: ${stockMinimo} unidades\n\n${CONFIG.NOMBRE_NEGOCIO}`
    });
  } catch(e) { console.log("No se pudo enviar el correo: " + e.message); }
}

// ─────────────────────────────────────────
// REPORTES
// ─────────────────────────────────────────
function generarReporte() {
  verificarAcceso();
  const ui        = SpreadsheetApp.getUi();
  const respuesta = ui.prompt(
    "📊 Generar Reporte",
    "Ingresa el período:\n1 - Hoy\n2 - Esta semana\n3 - Este mes\n4 - Todo el historial",
    ui.ButtonSet.OK_CANCEL
  );
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;
  const opcion = respuesta.getResponseText().trim();
  const ahora  = new Date();
  let fechaInicio;
  switch(opcion) {
    case "1": fechaInicio = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate()); break;
    case "2": fechaInicio = new Date(ahora.getTime() - 7*24*60*60*1000); break;
    case "3": fechaInicio = new Date(ahora.getFullYear(), ahora.getMonth(), 1); break;
    case "4": fechaInicio = new Date(2000, 0, 1); break;
    default:  ui.alert("Opción no válida."); return;
  }
  construirReporte(fechaInicio, ahora);
}

function construirReporte(fechaInicio, fechaFin) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const hojaM = ss.getSheetByName(HOJAS.MOVIMIENTOS);
  const hojaR = ss.getSheetByName(HOJAS.REPORTES);
  const hojaI = ss.getSheetByName(HOJAS.INVENTARIO);
  hojaR.clearContents();

  const movimientos = hojaM.getDataRange().getValues();
  let totalEntradas=0, totalSalidas=0, valorEntradas=0, valorSalidas=0, gananciaReal=0;
  const detalle = [];

  for (let i = 1; i < movimientos.length; i++) {
    const fecha = new Date(movimientos[i][0]);
    if (fecha >= fechaInicio && fecha <= fechaFin) {
      const tipo     = movimientos[i][1];
      const cantidad = movimientos[i][3];
      const valorMov = movimientos[i][8];
      const pC       = movimientos[i][6];
      const pV       = movimientos[i][7];
      if (tipo === TIPOS_MOVIMIENTO.ENTRADA || tipo === TIPOS_MOVIMIENTO.ENTRADA_INICIAL) {
        totalEntradas += cantidad; valorEntradas += valorMov;
      } else if (tipo === TIPOS_MOVIMIENTO.SALIDA) {
        totalSalidas += cantidad; valorSalidas += valorMov; gananciaReal += (pV - pC) * cantidad;
      }
      detalle.push(movimientos[i]);
    }
  }

  const inventario = hojaI.getDataRange().getValues();
  let valorInventario = 0, gananciaPotencial = 0;
  for (let i = 1; i < inventario.length; i++) {
    valorInventario   += inventario[i][4] * inventario[i][6];
    gananciaPotencial += inventario[i][10];
  }

  const titulo = [
    [`📊 REPORTE DE INVENTARIO - ${CONFIG.NOMBRE_NEGOCIO}`],
    [`Período: ${fechaInicio.toLocaleDateString()} al ${fechaFin.toLocaleDateString()}`],
    [`Generado: ${new Date().toLocaleString()} | Usuario: ${obtenerUsuarioActual()}`],
    [""],
    ["═══════ RESUMEN FINANCIERO ═══════"],
    ["Total unidades ingresadas",         totalEntradas],
    ["Total unidades vendidas/salidas",   totalSalidas],
    ["Valor total en compras",            `$${valorEntradas.toLocaleString()}`],
    ["Valor total en ventas",             `$${valorSalidas.toLocaleString()}`],
    ["💰 Ganancia real del período",      `$${gananciaReal.toLocaleString()}`],
    [""],
    ["═══════ ESTADO ACTUAL INVENTARIO ═══════"],
    ["Valor total del inventario (a precio compra)", `$${valorInventario.toLocaleString()}`],
    ["Ganancia potencial si se vende todo (a precio local)", `$${gananciaPotencial.toLocaleString()}`],
    [""],
    ["═══════ DETALLE DE MOVIMIENTOS ═══════"],
    ["Fecha","Tipo","Producto","Cantidad","Stock Ant.","Stock Nuevo","P. Compra","P. Venta Aplicado","Valor","Usuario","Notas"],
    ...detalle
  ];

  hojaR.getRange(1,1,titulo.length,11).setValues(titulo.map(f => f.length < 11 ? [...f, ...Array(11-f.length).fill("")] : f));
  hojaR.getRange("A1").setFontSize(14).setFontWeight("bold").setFontColor("#1a1a2e");
  hojaR.getRange("A5").setFontWeight("bold").setBackground("#e8f4f8");
  hojaR.getRange("A12").setFontWeight("bold").setBackground("#e8f4f8");
  hojaR.getRange("A16").setFontWeight("bold").setBackground("#e8f4f8");
  ss.setActiveSheet(hojaR);
  registrarLog("REPORTE GENERADO", `Período: ${fechaInicio.toLocaleDateString()} - ${fechaFin.toLocaleDateString()}`);
}

// ─────────────────────────────────────────
// LOGS
// ─────────────────────────────────────────
function registrarLog(accion, detalle) {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.LOGS);
    if (!hoja) return;
    const correo = obtenerUsuarioActual();
    hoja.appendRow([new Date(), correo.split("@")[0], correo, accion, detalle]);
  } catch(e) { console.log("Error al registrar log: " + e.message); }
}

// ─────────────────────────────────────────
// MIGRACIÓN ÚNICA: separa filas antiguas que quedaron con varios productos
// combinados en una sola celda (formato usado brevemente antes de volver a
// "una fila por producto"). Se puede ejecutar más de una vez sin problema:
// las filas que ya están limpias (sin coma en Producto) se dejan intactas.
//
// CÓMO USARLA: en el editor de Apps Script, selecciona
// "migrarDatosAntiguosCombinados" en el menú desplegable de arriba (junto al
// botón ▶ Ejecutar) y presiona Ejecutar. Revisa "Ver > Registros" para ver
// cuántas filas se repararon. No borra ni modifica el stock del Inventario.
// ─────────────────────────────────────────
function migrarComandasAntiguasCombinadas() {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.COMANDAS);
  const datos = hoja.getDataRange().getValues();

  const filasNuevas  = [];
  const filasABorrar = [];
  let reparadas = 0;

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const producto = String(fila[4] || "");
    if (producto.indexOf(",") === -1) continue; // fila ya limpia, no tocar

    const productos  = producto.split(", ");
    const cantidades = String(fila[5]).split(", ").map(Number);
    const precios    = String(fila[6]).split(", ").map(Number);
    const subtotales = String(fila[7]).split(", ").map(Number);
    const n = Math.min(productos.length, cantidades.length, precios.length, subtotales.length);
    if (n === 0) continue;

    for (let k = 0; k < n; k++) {
      // Si algún valor no es numérico (dato ya corrupto), se descarta esa línea
      // en vez de arrastrar el problema a la fila nueva.
      if (!Number.isFinite(cantidades[k]) || !Number.isFinite(precios[k]) || !Number.isFinite(subtotales[k])) continue;
      filasNuevas.push([
        fila[0], fila[1], fila[2], fila[3], productos[k], cantidades[k], precios[k], subtotales[k],
        fila[8], fila[9], fila[10]
      ]);
    }
    filasABorrar.push(i + 1); // fila real en la hoja (1-based, +1 por encabezado)
    reparadas++;
  }

  filasABorrar.sort((a,b) => b - a).forEach(fila => hoja.deleteRow(fila));
  if (filasNuevas.length > 0) {
    hoja.getRange(hoja.getLastRow() + 1, 1, filasNuevas.length, 11).setValues(filasNuevas);
  }

  return { reparadas, filasNuevas: filasNuevas.length };
}

function migrarMovimientosAntiguosCombinados() {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.MOVIMIENTOS);
  const datos = hoja.getDataRange().getValues();

  const filasNuevas  = [];
  const filasABorrar = [];
  let reparadas = 0;

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const producto = String(fila[2] || "");
    if (producto.indexOf(",") === -1) continue;

    const tipo = fila[1];
    const productos         = producto.split(", ");
    const cantidades        = String(fila[3]).split(", ").map(Number);
    const stocksAnteriores  = String(fila[4]).split(", ").map(Number);
    const stocksNuevos      = String(fila[5]).split(", ").map(Number);
    const preciosCompra     = String(fila[6]).split(", ").map(Number);
    const preciosVenta      = String(fila[7]).split(", ").map(Number);
    const n = Math.min(productos.length, cantidades.length, stocksAnteriores.length, stocksNuevos.length, preciosCompra.length, preciosVenta.length);
    if (n === 0) continue;

    for (let k = 0; k < n; k++) {
      if (![cantidades[k], stocksAnteriores[k], stocksNuevos[k], preciosCompra[k], preciosVenta[k]].every(Number.isFinite)) continue;
      const valor = tipo === TIPOS_MOVIMIENTO.SALIDA ? cantidades[k]*preciosVenta[k] : cantidades[k]*preciosCompra[k];
      filasNuevas.push([
        fila[0], tipo, productos[k], cantidades[k], stocksAnteriores[k], stocksNuevos[k],
        preciosCompra[k], preciosVenta[k], valor, fila[9], fila[10]
      ]);
    }
    filasABorrar.push(i + 1);
    reparadas++;
  }

  filasABorrar.sort((a,b) => b - a).forEach(fila => hoja.deleteRow(fila));
  if (filasNuevas.length > 0) {
    hoja.getRange(hoja.getLastRow() + 1, 1, filasNuevas.length, 11).setValues(filasNuevas);
  }

  return { reparadas, filasNuevas: filasNuevas.length };
}

// Ejecuta AMBAS migraciones (Comandas y Movimientos) y muestra un resumen.
// Esta es la función que debes seleccionar y ejecutar desde el editor.
function migrarDatosAntiguosCombinados() {
  const resComandas    = migrarComandasAntiguasCombinadas();
  const resMovimientos = migrarMovimientosAntiguosCombinados();
  const mensaje =
    `Comandas: ${resComandas.reparadas} fila(s) combinada(s) → ${resComandas.filasNuevas} fila(s) nueva(s) separadas.\n` +
    `Movimientos: ${resMovimientos.reparadas} fila(s) combinada(s) → ${resMovimientos.filasNuevas} fila(s) nueva(s) separadas.`;
  Logger.log(mensaje);
  try { SpreadsheetApp.getUi().alert("✅ Migración completa", mensaje, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e) {}
  return mensaje;
}

// ─────────────────────────────────────────
// UTILIDADES
// ─────────────────────────────────────────
function obtenerListaProductos() {
  const hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.INVENTARIO);
  const datos = hoja.getDataRange().getValues();
  const productos = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0]) productos.push({
      id:              datos[i][0],
      nombre:          datos[i][1],
      stock:           datos[i][4],
      precioCompra:    datos[i][6],
      precioLocal:     datos[i][7],
      precioDomicilio: datos[i][8]
    });
  }
  return productos;
}

function obtenerProductosParaFormulario() { return obtenerListaProductos(); }

function estilizarEncabezado(hoja, numColumnas) {
  hoja.getRange(1,1,1,numColumnas)
    .setBackground("#1a1a2e")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  hoja.setRowHeight(1, 35);
}

function generarHtmlPrecios(productos) {
  let opciones = productos.map(p =>
    `<option value="${p.id}" data-compra="${p.precioCompra}" data-local="${p.precioLocal}" data-domicilio="${p.precioDomicilio}">${p.nombre} (Local: $${p.precioLocal} | Domicilio: $${p.precioDomicilio})</option>`
  ).join("");
  return `<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;padding:20px;background:#f5f5f5;}select,input{width:100%;padding:8px;margin:6px 0 12px;border:1px solid #ccc;border-radius:6px;font-size:14px;}button{background:#1a1a2e;color:white;padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-size:15px;width:100%;}h3{color:#1a1a2e;}</style></head><body><h3>✏️ Editar Precios</h3><label>Producto:</label><select id="producto">${opciones}</select><label>Nuevo precio de compra c/u ($):</label><input type="number" id="compra" min="0" step="100"><label>Nuevo precio de venta LOCAL c/u ($):</label><input type="number" id="local" min="0" step="100"><label>Nuevo precio de venta DOMICILIO c/u ($):</label><input type="number" id="domicilio" min="0" step="100"><br><button onclick="guardar()">💾 Guardar Cambios</button><script>function guardar(){var id=document.getElementById('producto').value;var compra=parseFloat(document.getElementById('compra').value);var local=parseFloat(document.getElementById('local').value);var domicilio=parseFloat(document.getElementById('domicilio').value);if(!compra||!local){alert('Ingresa al menos precio de compra y precio local.');return;}if(local<compra){alert('El precio local debe ser mayor o igual al de compra.');return;}google.script.run.withSuccessHandler(function(r){alert(r.exito?'✅ Precios actualizados.':r.mensaje);if(r.exito)google.script.host.close();}).actualizarPrecio(id,compra,local,domicilio||local);}<\/script></body></html>`;
}

function reporteDiarioAutomatico() {
  const ahora = new Date();
  construirReporte(new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate()), ahora);
  const alertas = obtenerProductosBajoStock();
  if (alertas.length > 0) {
    let cuerpo = `Buenos días,\n\nResumen de inventario - ${CONFIG.NOMBRE_NEGOCIO}\n\n⚠️ PRODUCTOS QUE NECESITAN REPOSICIÓN:\n`;
    alertas.forEach(a => { cuerpo += `• ${a.nombre}: ${a.stock} unidades\n`; });
    MailApp.sendEmail({ to: CONFIG.CORREO_ALERTAS, subject: `📊 Reporte Diario - ${CONFIG.NOMBRE_NEGOCIO}`, body: cuerpo });
  }
}

// ─────────────────────────────────────────
// DASHBOARD EN SHEETS
// ─────────────────────────────────────────
function crearHojaDashboard(ss) {
  let hoja = ss.getSheetByName(HOJAS.DASHBOARD);
  if (!hoja) hoja = ss.insertSheet(HOJAS.DASHBOARD, 0);
  hoja.clearContents(); hoja.clearFormats();
  hoja.getRange("A1:H1").merge().setValue("📈 DASHBOARD DE VENTAS - " + CONFIG.NOMBRE_NEGOCIO);
  hoja.getRange("A1").setFontSize(16).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff").setHorizontalAlignment("center");
  hoja.setRowHeight(1, 45);
  hoja.getRange("A2:H2").merge().setValue("Última actualización: pendiente");
  hoja.getRange("A2").setFontColor("#888888").setFontSize(10).setHorizontalAlignment("center");
  hoja.setRowHeight(2, 22);
}

function actualizarDashboard() {
  verificarAcceso();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(HOJAS.DASHBOARD);
  if (!hoja) { crearHojaDashboard(ss); hoja = ss.getSheetByName(HOJAS.DASHBOARD); }
  hoja.clearContents(); hoja.clearFormats(); hoja.clearNotes();
  hoja.getCharts().forEach(g => hoja.removeChart(g));

  const movimientos = ss.getSheetByName(HOJAS.MOVIMIENTOS).getDataRange().getValues();
  const inventario  = ss.getSheetByName(HOJAS.INVENTARIO).getDataRange().getValues();

  const metricas = calcularMetricasDashboard(movimientos, inventario, "mes");

  escribirEncabezadoDashboard(hoja);
  escribirKPIs(hoja, metricas);
  escribirTablaTopProductos(hoja, metricas.topProductos);
  escribirTablaTendencia(hoja, metricas.tendenciaMensual);
  escribirTablaRentabilidad(hoja, metricas.rentabilidad);
  Utilities.sleep(800);
  crearGraficaTopProductos(hoja, metricas.topProductos);
  crearGraficaTendencia(hoja, metricas.tendenciaMensual);
  crearGraficaRentabilidad(hoja, metricas.rentabilidad);
  ss.setActiveSheet(hoja);
  registrarLog("DASHBOARD ACTUALIZADO", "Métricas del mes generadas");
  SpreadsheetApp.getUi().alert("✅ Dashboard actualizado correctamente.");
}

// Nota: las métricas se calculan a partir de "Movimientos", donde cada
// SALIDA ya quedó registrada con el precio realmente aplicado (local o
// domicilio), así que ingresos/ganancia reflejan el canal real de venta
// sin necesitar lógica adicional aquí.
function calcularMetricasDashboard(movimientos, inventario, periodo) {
  const ahora = new Date();
  let fechaInicio;
  switch (periodo) {
    case "hoy":    fechaInicio = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate()); break;
    case "semana": fechaInicio = new Date(ahora.getTime() - 7*24*60*60*1000); break;
    case "mes":    fechaInicio = new Date(ahora.getFullYear(), ahora.getMonth(), 1); break;
    default:       fechaInicio = new Date(2000, 0, 1);
  }

  const vProd = {}, vMes = {};
  let totalIngresos = 0, totalGanancia = 0, totalUnidades = 0;

  for (let i = 1; i < movimientos.length; i++) {
    const f = movimientos[i];
    if (!f[0] || f[1] !== TIPOS_MOVIMIENTO.SALIDA) continue;
    const fecha = new Date(f[0]);
    if (fecha < fechaInicio) continue;

    const prod = f[2], cant = Number(f[3]) || 0;
    const pC = Number(f[6]) || 0, pV = Number(f[7]) || 0;
    const ing = cant * pV, gan = cant * (pV - pC);
    const mk = `${fecha.getFullYear()}-${String(fecha.getMonth()+1).padStart(2,"0")}`;
    const ml = fecha.toLocaleDateString("es-CO", { month:"short", year:"numeric" });

    totalIngresos += ing; totalGanancia += gan; totalUnidades += cant;

    if (!vProd[prod]) vProd[prod] = { unidades:0, ingresos:0, ganancia:0 };
    vProd[prod].unidades += cant; vProd[prod].ingresos += ing; vProd[prod].ganancia += gan;

    if (!vMes[mk]) vMes[mk] = { label:ml, ingresos:0, ganancia:0, unidades:0 };
    vMes[mk].ingresos += ing; vMes[mk].ganancia += gan; vMes[mk].unidades += cant;
  }

  const topProductos = Object.entries(vProd)
    .map(([n,d]) => ({ nombre:n, ...d, margenPct: d.ingresos>0 ? Math.round(d.ganancia/d.ingresos*100) : 0 }))
    .sort((a,b) => b.unidades - a.unidades).slice(0,8);

  const tendenciaMensual = Object.entries(vMes)
    .sort((a,b) => a[0].localeCompare(b[0])).slice(-6).map(([,d]) => d);

  const rentabilidad = Object.entries(vProd)
    .map(([n,d]) => ({ nombre:n, ganancia:d.ganancia, margenPct: d.ingresos>0 ? Math.round(d.ganancia/d.ingresos*100) : 0 }))
    .sort((a,b) => b.ganancia - a.ganancia).slice(0,6);

  let valorInventario = 0;
  for (let i = 1; i < inventario.length; i++)
    valorInventario += (Number(inventario[i][4])||0) * (Number(inventario[i][6])||0);

  return { totalIngresos, totalGanancia, totalUnidades, valorInventario, topProductos, tendenciaMensual, rentabilidad };
}

function escribirEncabezadoDashboard(hoja) {
  hoja.getRange("A1:J1").merge().setValue("📈 DASHBOARD DE VENTAS — " + CONFIG.NOMBRE_NEGOCIO);
  hoja.getRange("A1").setFontSize(16).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff").setHorizontalAlignment("center");
  hoja.setRowHeight(1, 48);
  hoja.getRange("A2:J2").merge().setValue("Última actualización: " + new Date().toLocaleString("es-CO")).setFontSize(10).setFontColor("#888888").setHorizontalAlignment("center");
  hoja.setRowHeight(2, 20);
}

function escribirKPIs(hoja, m) {
  hoja.getRange("A4:H4").setValues([["💰 Ingresos Totales","","📦 Unidades Vendidas","","🏆 Ganancia Total","","🗄️ Valor Inventario",""]]);
  ["A4","C4","E4","G4"].forEach(c => hoja.getRange(c).setBackground("#16213e").setFontColor("#aaaaaa").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center"));
  hoja.setRowHeight(4, 28);
  hoja.getRange("A5").setValue("$" + Math.round(m.totalIngresos).toLocaleString("es-CO"));
  hoja.getRange("C5").setValue(m.totalUnidades.toLocaleString("es-CO") + " uds");
  hoja.getRange("E5").setValue("$" + Math.round(m.totalGanancia).toLocaleString("es-CO"));
  hoja.getRange("G5").setValue("$" + Math.round(m.valorInventario).toLocaleString("es-CO"));
  ["A5","C5","E5","G5"].forEach(c => hoja.getRange(c).setFontSize(18).setFontWeight("bold").setBackground("#0f3460").setFontColor("#e2b714").setHorizontalAlignment("center"));
  ["B5","D5","F5","H5"].forEach(c => hoja.getRange(c).setBackground("#0f3460"));
  ["A4:B5","C4:D5","E4:F5","G4:H5"].forEach(r => hoja.getRange(r).merge());
  hoja.setRowHeight(5, 52); hoja.setRowHeight(6, 14);
}

function escribirTablaTopProductos(hoja, top) {
  const fi = 7;
  hoja.getRange(`A${fi}:D${fi}`).merge().setValue("🏅 TOP PRODUCTOS MÁS VENDIDOS").setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  hoja.setRowHeight(fi, 30);
  hoja.getRange(fi+1,1,1,4).setValues([["Producto","Unidades","Ingresos ($)","Ganancia ($)"]]).setBackground("#16213e").setFontColor("#cccccc").setFontWeight("bold");
  if (!top.length) { hoja.getRange(fi+2,1,1,4).merge().setValue("Sin datos de ventas aún"); return; }
  const datos = top.map(p => [p.nombre, p.unidades, "$"+Math.round(p.ingresos).toLocaleString("es-CO"), "$"+Math.round(p.ganancia).toLocaleString("es-CO")]);
  hoja.getRange(fi+2,1,datos.length,4).setValues(datos);
  for (let i = 0; i < datos.length; i++) hoja.getRange(fi+2+i,1,1,4).setBackground(i%2===0?"#f8f9fa":"#ffffff");
  hoja.setColumnWidth(1,200); hoja.setColumnWidth(2,90); hoja.setColumnWidth(3,120); hoja.setColumnWidth(4,120);
}

function escribirTablaTendencia(hoja, tendencia) {
  const fi = 7;
  hoja.getRange(`F${fi}:H${fi}`).merge().setValue("📅 TENDENCIA MENSUAL (6 meses)").setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  hoja.getRange(fi+1,6,1,3).setValues([["Mes","Ingresos ($)","Ganancia ($)"]]).setBackground("#16213e").setFontColor("#cccccc").setFontWeight("bold");
  if (!tendencia.length) { hoja.getRange(fi+2,6,1,3).merge().setValue("Sin datos aún"); return; }
  const datos = tendencia.map(m => [m.label, "$"+Math.round(m.ingresos).toLocaleString("es-CO"), "$"+Math.round(m.ganancia).toLocaleString("es-CO")]);
  hoja.getRange(fi+2,6,datos.length,3).setValues(datos);
  for (let i = 0; i < datos.length; i++) hoja.getRange(fi+2+i,6,1,3).setBackground(i%2===0?"#f8f9fa":"#ffffff");
  hoja.setColumnWidth(6,100); hoja.setColumnWidth(7,120); hoja.setColumnWidth(8,120);
}

function escribirTablaRentabilidad(hoja, rent) {
  const fi = 7;
  hoja.getRange(`J${fi}:K${fi}`).merge().setValue("💎 RENTABILIDAD").setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  hoja.getRange(fi+1,10,1,2).setValues([["Producto","Margen %"]]).setBackground("#16213e").setFontColor("#cccccc").setFontWeight("bold");
  if (!rent.length) { hoja.getRange(fi+2,10,1,2).merge().setValue("Sin datos aún"); return; }
  const datos = rent.map(p => [p.nombre, p.margenPct+"%"]);
  hoja.getRange(fi+2,10,datos.length,2).setValues(datos);
  for (let i = 0; i < datos.length; i++) hoja.getRange(fi+2+i,10,1,2).setBackground(i%2===0?"#f8f9fa":"#ffffff");
  hoja.setColumnWidth(10,200); hoja.setColumnWidth(11,80);
}

function crearGraficaTopProductos(hoja, top) {
  if (!top||!top.length) return;
  const rango = hoja.getRange(8,1,top.length,2);
  hoja.insertChart(hoja.newChart().setChartType(Charts.ChartType.BAR).addRange(rango).setPosition(20,1,0,0)
    .setOption("title","Unidades vendidas por producto").setOption("titleTextStyle",{fontSize:13,bold:true,color:"#1a1a2e"})
    .setOption("hAxis.title","Unidades").setOption("legend",{position:"none"})
    .setOption("colors",["#e2b714"]).setOption("backgroundColor","#f8f9fa")
    .setOption("width",480).setOption("height",300).build());
}

function crearGraficaTendencia(hoja, tendencia) {
  if (!tendencia||!tendencia.length) return;
  const dg = tendencia.map(m => [m.label, m.ingresos, m.ganancia]);
  hoja.getRange(8,14,dg.length,3).setValues(dg);
  const rango = hoja.getRange(8,14,dg.length,3);
  hoja.insertChart(hoja.newChart().setChartType(Charts.ChartType.LINE).addRange(rango).setPosition(20,6,0,0)
    .setOption("title","Tendencia de ingresos y ganancias (6 meses)").setOption("vAxis.title","Pesos ($)")
    .setOption("legend",{position:"bottom"}).setOption("series",{0:{color:"#0f3460",lineWidth:3},1:{color:"#e2b714",lineWidth:3}})
    .setOption("backgroundColor","#f8f9fa").setOption("width",480).setOption("height",300).build());
  hoja.hideColumns(14,3);
}

function crearGraficaRentabilidad(hoja, rent) {
  if (!rent||!rent.length) return;
  const dg = rent.map(p => [p.nombre, p.margenPct]);
  hoja.getRange(8,17,dg.length,2).setValues(dg);
  const rango = hoja.getRange(8,17,dg.length,2);
  hoja.insertChart(hoja.newChart().setChartType(Charts.ChartType.COLUMN).addRange(rango).setPosition(20,10,0,0)
    .setOption("title","Margen de ganancia por producto (%)").setOption("vAxis.title","Margen %")
    .setOption("legend",{position:"none"}).setOption("colors",["#1b5e20"])
    .setOption("backgroundColor","#f8f9fa").setOption("width",380).setOption("height",300).build());
  hoja.hideColumns(17,2);
}

// ─────────────────────────────────────────
// INTERFAZ WEB
// ─────────────────────────────────────────
function abrirInterfaz() {
  const html = HtmlService.createHtmlOutputFromFile("Interfaz").setWidth(1200).setHeight(700).setTitle("Sistema de Inventario");
  SpreadsheetApp.getUi().showModalDialog(html, "🍾 " + CONFIG.NOMBRE_NEGOCIO);
}

function obtenerDatosInterfaz() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const correo = obtenerUsuarioActual();
  const rol    = obtenerRolUsuario() || "OPERARIO";

  const hojaInv  = ss.getSheetByName(HOJAS.INVENTARIO);
  const datosInv = hojaInv ? hojaInv.getDataRange().getValues() : [];
  const productos = [];
  for (let i = 1; i < datosInv.length; i++) {
    if (!datosInv[i][0]) continue;
    productos.push({
      id:              datosInv[i][0],
      nombre:          datosInv[i][1],
      categoria:       datosInv[i][2],
      unidad:          datosInv[i][3],
      stock:           datosInv[i][4],
      stockMinimo:     datosInv[i][5],
      precioCompra:    datosInv[i][6],
      precioLocal:     datosInv[i][7],
      precioDomicilio: datosInv[i][8],
      margen:          datosInv[i][9]
    });
  }

  const alertas     = productos.filter(p => p.stock <= p.stockMinimo).map(p => ({ nombre:p.nombre, stock:p.stock, minimo:p.stockMinimo }));
  const hojaMov     = ss.getSheetByName(HOJAS.MOVIMIENTOS);
  const movimientos = hojaMov ? hojaMov.getDataRange().getValues() : [];
  const metricas    = calcularMetricasDashboard(movimientos, datosInv, "mes");

  const hojaLogs = ss.getSheetByName(HOJAS.LOGS);
  const logs     = [];
  if (hojaLogs) {
    const dl    = hojaLogs.getDataRange().getValues();
    const inicio= Math.max(1, dl.length-8);
    for (let i = dl.length-1; i >= inicio; i--) {
      if (!dl[i][0]) continue;
      logs.push({ fecha:new Date(dl[i][0]).toLocaleString("es-CO"), usuario:dl[i][1], accion:dl[i][3], detalle:dl[i][4] });
    }
  }

  const usuarios = [
    ...CONFIG.CORREOS_ADMIN.map(c    => ({ correo:c, nombre:c.split("@")[0], rol:"ADMIN"    })),
    ...CONFIG.CORREOS_OPERARIO.map(c => ({ correo:c, nombre:c.split("@")[0], rol:"OPERARIO" }))
  ];

  const hojaMesasData = ss.getSheetByName(HOJAS.MESAS);
  const ahora         = new Date();
  const nombresMesas  = [];
  const mesasResumen  = { total:0, ocupadas:0, totalAcumulado:0, alertas:[] };
  if (hojaMesasData) {
    const dm = hojaMesasData.getDataRange().getValues();
    for (let i = 1; i < dm.length; i++) {
      nombresMesas.push(dm[i][0]);
      mesasResumen.total++;
      if (dm[i][1] === ESTADOS_MESA.OCUPADA) {
        mesasResumen.ocupadas++;
        mesasResumen.totalAcumulado += (dm[i][3] || 0);
        const horaAp = dm[i][2] ? new Date(dm[i][2]) : null;
        const mins   = horaAp ? Math.floor((ahora - horaAp) / 60000) : 0;
        if (mins >= CONFIG.ALERTA_MESA_MINUTOS) mesasResumen.alertas.push({ mesa:dm[i][0], minutos:mins });
      }
    }
  }

  metricas.logs    = logs;
  metricas.usuarios= usuarios;
  metricas.config  = {
    nombreNegocio:     CONFIG.NOMBRE_NEGOCIO,
    correoAlertas:     CONFIG.CORREO_ALERTAS,
    stockMinDefault:   CONFIG.STOCK_MINIMO_DEFAULT,
    numeroMesas:       CONFIG.NUMERO_MESAS,
    alertaMesaMinutos: CONFIG.ALERTA_MESA_MINUTOS
  };
  metricas.nombresMesas = nombresMesas;
  metricas.mesasResumen = mesasResumen;

  const presentaciones = obtenerPresentaciones();

  return { productos, alertas, metricas, presentaciones, usuario:{ correo, nombre:correo.split("@")[0], rol } };
}

// Resuelve el rango [fechaInicio, fechaFin] para las opciones del selector de
// Reportes (1=Hoy, 2=Semana, 3=Mes, 4=Historial, 5=Fecha específica). Centralizado
// aquí porque tanto generarReporteDesdeInterfaz como obtenerMetricasReporte
// necesitan exactamente el mismo rango — así nunca se desincronizan entre sí.
function resolverRangoReporte(opcion, fechaPersonalizada) {
  const ahora = new Date();
  let fechaInicio, fechaFin = ahora;
  switch(String(opcion)) {
    case "1": fechaInicio = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate()); break;
    case "2": fechaInicio = new Date(ahora.getTime() - 7*24*60*60*1000); break;
    case "3": fechaInicio = new Date(ahora.getFullYear(), ahora.getMonth(), 1); break;
    case "5":
      // fechaPersonalizada llega como "2026-07-15" desde <input type="date">
      if (!fechaPersonalizada) throw new Error("Falta la fecha seleccionada.");
      var partes = fechaPersonalizada.split("-").map(Number);
      fechaInicio = new Date(partes[0], partes[1]-1, partes[2], 0, 0, 0);
      fechaFin    = new Date(partes[0], partes[1]-1, partes[2], 23, 59, 59);
      break;
    default:  fechaInicio = new Date(2000, 0, 1);
  }
  return { fechaInicio, fechaFin };
}

function generarReporteDesdeInterfaz(opcion, fechaPersonalizada) {
  const r = resolverRangoReporte(opcion, fechaPersonalizada);
  construirReporte(r.fechaInicio, r.fechaFin);
}

// Calcula ingresos/ganancia/unidades del período SIN escribir en Sheets —
// para alimentar los KPIs de la página Reportes en tiempo real, antes de
// que el usuario decida generar el reporte completo.
function obtenerMetricasReporte(opcion, fechaPersonalizada) {
  verificarAcceso();
  const r = resolverRangoReporte(opcion, fechaPersonalizada);
  const hojaM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS.MOVIMIENTOS);
  const movimientos = hojaM.getDataRange().getValues();

  let ingresos = 0, ganancia = 0, unidades = 0;
  for (let i = 1; i < movimientos.length; i++) {
    const fila = movimientos[i];
    if (!fila[0] || fila[1] !== TIPOS_MOVIMIENTO.SALIDA) continue;
    const fecha = new Date(fila[0]);
    if (fecha < r.fechaInicio || fecha > r.fechaFin) continue;

    const cantidad = Number(fila[3]) || 0;
    const pC = Number(fila[6]) || 0;
    const pV = Number(fila[7]) || 0;
    ingresos += cantidad * pV;
    ganancia += cantidad * (pV - pC);
    unidades += cantidad;
  }
  return { ingresos, ganancia, unidades };
}

// Exportar para pruebas unitarias en Node.js (no afecta a Apps Script,
// que ignora "module" al no existir ese objeto global).
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    ESTADOS_MESA, ESTADOS_COMANDA, TIPOS_MOVIMIENTO, TIPOS_ENTREGA, CONFIG, HOJAS,
    crearHojaInventario, crearHojaMovimientos, crearHojaAlertas, crearHojaLogs,
    crearHojaMesas, crearHojaComandas, obtenerHistorialComandas,
    construirMapaInventario, precioSegunEntrega, resolverItemInventario,
    agregarProducto, actualizarPrecio, eliminarProducto,
    registrarEntrada, registrarSalida, procesarEntrada, procesarSalida,
    registrarMovimiento, actualizarHojaAlertas, obtenerProductosBajoStock,
    obtenerListaProductos, obtenerDatosInterfaz, calcularMetricasDashboard,
    construirReporte, resolverRangoReporte, obtenerMetricasReporte,
    agregarProductoMesa, abrirMesa, cerrarMesa,
    eliminarItemComanda, obtenerComandaMesa, obtenerEstadoMesas, restarCantidadComanda, liberarMesaSinComanda,
    crearHojaPresentaciones, agregarPresentacion, obtenerPresentaciones,
    eliminarPresentacion, registrarSalidaPresentacion, agregarPresentacionAMesa, actualizarPresentacion,
    normalizarPagos, formatearPagos, parsearPagos, registrarVentaDirecta,
    obtenerResumenCaja, agregarItemAComandaMesa,
    migrarComandasAntiguasCombinadas, migrarMovimientosAntiguosCombinados, migrarDatosAntiguosCombinados,
  };
}