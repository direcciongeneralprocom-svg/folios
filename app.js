/* Generador de Formato de Operaciones a partir de factura CFDI (XML) + PDF
   Todo corre en el navegador, no se sube nada a ningun servidor. */

// ---------- Catalogos SAT (igual que en las plantillas) ----------
const REGIMEN = {
  "601": "601 - General de Ley Personas Morales",
  "603": "603 - Personas Morales con Fines No Lucrativos",
  "605": "605 - Sueldos y Salarios e Ingresos Asimilados a Salarios",
  "606": "606 - Arrendamiento",
  "608": "608 - Demás ingresos",
  "610": "610 - Residentes en el Extranjero sin Establecimiento Permanente en México",
  "611": "611 - Ingreso por Dividendos (socios y accionistas)",
  "612": "612 - Personas Físicas con Actividades Empresariales y Profesionales",
  "614": "614 - Ingresos por intereses",
  "615": "615 - Régimen de los ingresos por obtención de premios",
  "616": "616 - Sin obligaciones fiscales",
  "620": "620 - Sociedades Cooperativas de Producción que optan por diferir sus ingresos",
  "621": "621 - Incorporación Fiscal",
  "622": "622 - Actividades Agrícolas, Ganaderas, Silvícolas y Pesqueras",
  "623": "623 - Opcional para Grupos de Sociedades",
  "624": "624 - Régimen de los Coordinados",
  "625": "625 - Régimen de las Actividades Empresariales con ingresos a través de Plataformas Técnologicas",
  "626": "626 - Régimen Simplicifado de Confianza",
};

const USOCFDI = {
  "G01": "G01 - Adquisición de mercancías.",
  "G02": "G02 - Devoluciones, descuentos o bonificaciones.",
  "G03": "G03 - Gastos en general.",
  "I01": "I01 - Construcciones.",
  "I02": "I02 - Mobiliario y equipo de oficina por inversiones.",
  "I03": "I03 - Equipo de transporte.",
  "I04": "I04 - Equipo de computo y accesorios.",
  "I05": "I05 - Dados, troqueles, moldes, matrices y herramental.",
  "I06": "I06 - Comunicaciones telefónicas.",
  "I07": "I07 - Comunicaciones satelitales.",
  "I08": "I08 - Otra maquinaria y equipo.",
  "D01": "D01 - Honorarios médicos, dentales y gastos hospitalarios.",
  "D02": "D02 - Gastos médicos por incapacidad o discapacidad.",
  "D03": "D03 - Gastos funerales.",
  "D04": "D04 - Donativos.",
  "D05": "D05 - Intereses reales efectivamente pagados por créditos hipotecarios (casa habitación).",
  "D06": "D06 - Aportaciones voluntarias al SAR.",
  "D07": "D07 - Primas por seguros de gastos médicos.",
  "D08": "D08 - Gastos de transportación escolar obligatoria.",
  "D09": "D09 - Depósitos en cuentas para el ahorro, primas que tengan como base planes de pensiones.",
  "D10": "D10 - Pagos por servicios educativos (colegiaturas).",
  "S01": "S01 - Sin efectos fiscales.",
  "CP01": "CP01 - Pagos",
  "CN01": "CN01 - Nómina",
};

const CONDICIONES = {
  "PUE": "PUE - Pago en una sola exhibición",
  "PPD": "PPD - Pago en parcialidades o diferido",
};

const FORMAPAGO = {
  "01": "1 - Efectivo", "02": "2 - Cheque nominativo", "03": "3 - Transferencia electrónica de fondos",
  "04": "4 - Tarjeta de crédito", "05": "5 - Monedero electrónico", "06": "6 - Dinero electrónico",
  "08": "8 - Vales de despensa", "12": "12 - Dación en pago", "13": "13 - Pago por subrogación",
  "14": "14 - Pago por consignación", "15": "15 - Condonación", "17": "17 - Compensación",
  "23": "23 - Novación", "24": "24 - Confusión", "25": "25 - Remisión de deuda",
  "26": "26 - Prescripción o caducidad", "27": "27 - A satisfacción del acreedor",
  "28": "28 - Tarjeta de débito", "29": "29 - Tarjeta de servicios", "30": "30 - Aplicación de anticipos",
  "31": "31 - Intermediario pagos", "99": "99 - Por definir",
};

// FormaPago corto (usado en plantilla Berna: "3 - Transferencia")
const FORMAPAGO_CORTO = {
  "01": "1 - Efectivo", "02": "2 - Cheque", "03": "3 - Transferencia",
  "04": "4 - Tarjeta de crédito", "28": "28 - Tarjeta de débito", "99": "99 - Por definir",
};

// ---------- Configuración por cliente ----------
const CLIENTES = {
  pape:   { label: "Luis Pape",  template: "templates/template_pape.xlsx",   sheet: "FACTURA", tipo: "A", pisoRow: 77 },
  estela: { label: "Estela H",   template: "templates/template_estela.xlsx", sheet: "FACTURA", tipo: "A", pisoRow: 71 },
  avila:  { label: "Luis Avila", template: "templates/template_avila.xlsx",  sheet: "FACTURA", tipo: "A", pisoRow: 61 },
  berna:  { label: "Berna",      template: "templates/template_berna.xlsx", sheet: "Factura", tipo: "B" },
};

// PISO: 1% del subtotal si la empresa emisora es BESTEN, 2.3% para cualquier otra empresa.
function calcularPisoPct(nombreEmisor) {
  return /besten/i.test(nombreEmisor || "") ? 0.01 : 0.023;
}

function regimenTipo(codigo) {
  const personaMoral = ["601", "603", "606", "608", "610", "611", "620", "622", "623", "624", "625", "626"];
  return personaMoral.includes(codigo) ? "PERSONA MORAL" : "PERSONA FISICA";
}

// ---------- Parseo del XML CFDI 4.0 ----------
function parseCFDI(xmlText) {
  const dom = new DOMParser().parseFromString(xmlText, "application/xml");
  if (dom.querySelector("parsererror")) {
    throw new Error("El XML no se pudo leer (formato inválido).");
  }
  const ns = "http://www.sat.gob.mx/cfd/4";

  const comprobante = dom.getElementsByTagNameNS(ns, "Comprobante")[0] || dom.querySelector("Comprobante");
  const emisorEl = dom.getElementsByTagNameNS(ns, "Emisor")[0] || dom.querySelector("Emisor");
  const receptorEl = dom.getElementsByTagNameNS(ns, "Receptor")[0] || dom.querySelector("Receptor");
  const conceptoEls = Array.from(dom.getElementsByTagNameNS(ns, "Concepto")).length
    ? Array.from(dom.getElementsByTagNameNS(ns, "Concepto"))
    : Array.from(dom.querySelectorAll("Concepto"));

  if (!comprobante || !emisorEl || !receptorEl) {
    throw new Error("El XML no parece un CFDI 4.0 válido (faltan nodos Comprobante/Emisor/Receptor).");
  }

  const attr = (el, name) => (el && el.getAttribute(name)) || "";

  const conceptos = conceptoEls.map(c => ({
    claveProdServ: attr(c, "ClaveProdServ"),
    claveUnidad: attr(c, "ClaveUnidad"),
    unidad: attr(c, "Unidad"),
    cantidad: parseFloat(attr(c, "Cantidad") || "0"),
    descripcion: attr(c, "Descripcion"),
    valorUnitario: parseFloat(attr(c, "ValorUnitario") || "0"),
    importe: parseFloat(attr(c, "Importe") || "0"),
  }));

  return {
    folio: attr(comprobante, "Folio") || attr(comprobante, "Serie"),
    fecha: attr(comprobante, "Fecha"),
    subTotal: parseFloat(attr(comprobante, "SubTotal") || "0"),
    total: parseFloat(attr(comprobante, "Total") || "0"),
    formaPago: attr(comprobante, "FormaPago"),
    metodoPago: attr(comprobante, "MetodoPago"),
    moneda: attr(comprobante, "Moneda"),
    emisor: {
      rfc: attr(emisorEl, "Rfc"),
      nombre: attr(emisorEl, "Nombre"),
      regimenFiscal: attr(emisorEl, "RegimenFiscal"),
    },
    receptor: {
      rfc: attr(receptorEl, "Rfc"),
      nombre: attr(receptorEl, "Nombre"),
      domicilioFiscal: attr(receptorEl, "DomicilioFiscalReceptor"),
      regimenFiscal: attr(receptorEl, "RegimenFiscalReceptor"),
      usoCFDI: attr(receptorEl, "UsoCFDI"),
    },
    conceptos,
  };
}

function formatFecha(fechaISO) {
  if (!fechaISO) return "";
  const d = fechaISO.split("T")[0];
  const [y, m, day] = d.split("-");
  if (!y || !m || !day) return fechaISO;
  return `${day}/${m}/${y}`;
}

function descripcionCombinada(conceptos) {
  if (conceptos.length === 1) return conceptos[0].descripcion;
  return conceptos.map((c, i) => `${i + 1}) ${c.descripcion}`).join("\n");
}

// ---------- Construcción del Excel ----------
async function generarExcel(clienteKey, datos, extras) {
  const cfg = CLIENTES[clienteKey];
  const resp = await fetch(cfg.template);
  if (!resp.ok) throw new Error("No se pudo cargar la plantilla " + cfg.template);
  const buffer = await resp.arrayBuffer();

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.getWorksheet(cfg.sheet);

  const conceptos = datos.conceptos.length ? datos.conceptos : [{
    claveProdServ: "", claveUnidad: "", unidad: "", cantidad: 1, descripcion: "", valorUnitario: 0, importe: 0,
  }];
  const cantidadTotal = conceptos.reduce((s, c) => s + (c.cantidad || 1), 0) || 1;
  const importeTotal = conceptos.reduce((s, c) => s + (c.importe || 0), 0) || datos.subTotal;
  const descripcion = descripcionCombinada(conceptos);
  const claveProdServ = conceptos[0].claveProdServ;
  const claveUnidad = conceptos[0].claveUnidad;
  const unidadLabel = extras.unidadLabel || conceptos[0].unidad || "SERVICIO";

  if (cfg.tipo === "A") {
    ws.getCell("D13").value = datos.emisor.nombre;

    ws.getCell("D15").value = datos.receptor.nombre;
    ws.getCell("D16").value = datos.receptor.rfc;
    ws.getCell("D17").value = extras.calle || "";
    ws.getCell("D18").value = extras.colonia || "";
    ws.getCell("D19").value = extras.ciudad || "";
    ws.getCell("D20").value = extras.cp || datos.receptor.domicilioFiscal || "";

    ws.getCell("D22").value = REGIMEN[datos.receptor.regimenFiscal] || datos.receptor.regimenFiscal;
    ws.getCell("D23").value = USOCFDI[datos.receptor.usoCFDI] || datos.receptor.usoCFDI;
    ws.getCell("D24").value = CONDICIONES[datos.metodoPago] || datos.metodoPago;
    ws.getCell("D25").value = FORMAPAGO[datos.formaPago] || datos.formaPago;

    if (extras.folioInterno) ws.getCell("H8").value = extras.folioInterno;
    ws.getCell("H9").value = datos.folio;
    ws.getCell("H10").value = formatFecha(datos.fecha);
    if (extras.fechaPago) ws.getCell("H11").value = extras.fechaPago;
    if (extras.fechaDevolucion) ws.getCell("H12").value = extras.fechaDevolucion;

    ws.getCell("B30").value = 1;
    ws.getCell("D30").value = descripcion;
    ws.getCell("E30").value = claveUnidad;
    ws.getCell("F30").value = unidadLabel;
    ws.getCell("G30").value = claveProdServ ? Number(claveProdServ) : claveProdServ;
    ws.getCell("H30").value = cantidadTotal;
    ws.getCell("I30").value = cantidadTotal ? importeTotal / cantidadTotal : importeTotal;

    // PISO: 1% del subtotal si la empresa emisora es BESTEN, 2.3% para las demás empresas.
    const pisoPct = calcularPisoPct(datos.emisor.nombre);
    ws.getCell(`G${cfg.pisoRow}`).value = { formula: `J31*${pisoPct}` };
    ws.getCell(`C${cfg.pisoRow}`).value = `PISO ${(pisoPct * 100).toFixed(1)}%`;
  } else {
    // Tipo B - plantilla Berna
    ws.getCell("C4").value = datos.emisor.nombre;
    ws.getCell("C6").value = datos.receptor.nombre;
    ws.getCell("C7").value = datos.receptor.rfc;
    ws.getCell("F6").value = regimenTipo(datos.receptor.regimenFiscal);
    ws.getCell("H6").value = extras.numExterior || "";
    ws.getCell("H7").value = extras.numInterior || "";
    ws.getCell("H8").value = extras.cp || datos.receptor.domicilioFiscal || "";
    ws.getCell("C8").value = extras.calle || "";
    ws.getCell("C9").value = extras.colonia || "";
    ws.getCell("C10").value = extras.estado || "";

    ws.getCell("B13").value = cantidadTotal;
    ws.getCell("C13").value = unidadLabel;
    ws.getCell("D13").value = claveUnidad;
    ws.getCell("E13").value = claveProdServ ? Number(claveProdServ) : claveProdServ;
    ws.getCell("F13").value = descripcion;
    ws.getCell("G13").value = cantidadTotal ? importeTotal / cantidadTotal : importeTotal;

    ws.getCell("C14").value = FORMAPAGO_CORTO[datos.formaPago] || datos.formaPago;
    ws.getCell("C15").value = datos.metodoPago;
    ws.getCell("C16").value = datos.receptor.usoCFDI;
  }

  const out = await wb.xlsx.writeBuffer();
  return new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

// ---------- Nombre del archivo descargado ----------
function formatMonto(valor) {
  return Number(valor || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function limpiarNombre(texto) {
  // Quita caracteres no válidos para nombres de archivo en Windows/Mac y espacios duplicados
  return String(texto || "")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function nombreArchivo(clienteKey, datos) {
  const broker = CLIENTES[clienteKey].label;
  const folio = datos.folio || "SIN-FOLIO";
  const clienteReceptor = datos.receptor.nombre || "";
  const empresaEmisora = datos.emisor.nombre || "";
  const monto = "$" + formatMonto(datos.total);
  const partes = [`Folio ${folio}`, broker, clienteReceptor, empresaEmisora, monto];
  return limpiarNombre(partes.join(" ")) + ".xlsx";
}

function descargarBlob(blob, nombre) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
