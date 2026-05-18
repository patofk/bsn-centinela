import fetch from 'node-fetch';
import * as XLSX from 'xlsx';

const SHEET_ID = '1tpgQbCpXav_P6_d28TVKLezXmSeFu7BBbc120FtGdFY';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;

function safeFloat(val, def = 0) {
  if (val === null || val === undefined) return def;
  const n = parseFloat(String(val).replace(',', '.'));
  return isNaN(n) ? def : n;
}
function safeStr(val, def = '') {
  if (val === null || val === undefined) return def;
  const s = String(val).trim();
  return (s === 'undefined' || s === 'null' || s === '') ? def : s;
}
function fechaStr(val) {
  if (!val) return '';
  if (typeof val === 'number') {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(epoch.getTime() + val * 86400000);
    if (d.getFullYear() < 2000) return '';
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }
  return safeStr(val);
}
function getCol(row, ...partials) {
  for (const partial of partials)
    for (const k of Object.keys(row))
      if (k.toLowerCase().includes(partial.toLowerCase())) return row[k];
  return null;
}
function leerHoja(wb, nombre) {
  if (!wb.SheetNames.includes(nombre)) { console.log('Hoja no encontrada:', nombre); return []; }
  const ws = wb.Sheets[nombre];
  const ref = ws['!ref'];
  if (!ref) return [];
  const range = XLSX.utils.decode_range(ref);
  range.s.r = 3;
  ws['!ref'] = XLSX.utils.encode_range(range);
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
  return rows;
}

function procesarExcel(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false });

  const factProd   = leerHoja(wb, 'FACT_PRODUCCION');
  const factCargas = leerHoja(wb, 'FACT_CARGAS');
  const factInv    = leerHoja(wb, 'FACT_INVENTARIO');
  const dimProd    = leerHoja(wb, 'DIM_PRODUCTO');
  const dimZona    = leerHoja(wb, 'DIM_ZONA');

  // ── PRODUCCIÓN ──
  const lotes = [];
  let totalLitrosProd = 0, totalKgProd = 0;
  for (const r of factProd) {
    const idLote   = safeStr(getCol(r,'ID Lote','id_lote','clave'));
    const producto = safeStr(getCol(r,'Producto'));
    const superv   = safeStr(getCol(r,'Supervisor','responsable'));
    const tk       = safeStr(getCol(r,'TK','Estiba','Estanque'));
    const litros   = safeFloat(getCol(r,'Terminado','Litros','Volumen','(L)'));
    const kg       = safeFloat(getCol(r,'usado','kg','Kg','KG'));
    const fechaProd = fechaStr(getCol(r,'Fecha','fecha'));
    if (!idLote || idLote.length < 3) continue;
    totalLitrosProd += litros;
    totalKgProd += kg;
    const tieneCarga = factCargas.some(c =>
      Object.entries(c).some(([k,v]) => k.toLowerCase().includes('lote') && safeStr(v) === idLote)
    );
    lotes.push({ id:idLote, producto, supervisor:superv, tk, litros, kg,
      duracion:0, fecha:fechaProd, estado: tieneCarga ? 'Despachado' : 'En TK' });
  }
  const lotesPend = lotes.filter(l=>l.estado==='En TK').length;
  const lotesDesp = lotes.length - lotesPend;
  const rendProm = totalLitrosProd > 0 ? Math.round(totalKgProd/totalLitrosProd*10)/10 : 0;

  // ── INVENTARIO ──
  const minimos = {};
  for (const d of dimProd) {
    const prod = safeStr(getCol(d,'Producto','nombre','ID'));
    if (prod) minimos[prod] = safeFloat(getCol(d,'Mínimo','Minimo','min','mínimo'));
  }
  const movimientos = [];
  for (const r of factInv) {
    const prod = safeStr(getCol(r,'Producto'));
    if (!prod || prod.length < 2) continue;
    const kgPorEnvase = safeFloat(getCol(r,'Cantidad por envase','envase','(ke)'), 750);
    const salidas  = safeFloat(getCol(r,'Salidas','salida','(-)')) * kgPorEnvase;
    const entradas = safeFloat(getCol(r,'Entradas','entrada','(+)')) * kgPorEnvase;
    const stockFinal = safeFloat(getCol(r,'Stock Final (kg)','Stock final (kg)')) ||
                       safeFloat(getCol(r,'Stock final','stock final','final')) * kgPorEnvase;
    const fecha = fechaStr(getCol(r,'Fecha'));
    movimientos.push({
      fecha, producto: prod,
      tipo: salidas > 0 ? 'Salida' : 'Entrada',
      kg: salidas > 0 ? salidas : entradas,
      responsable: safeStr(getCol(r,'Responsable','responsable','Bodega')),
      stock_final: stockFinal
    });
  }
  const stockPorProd = {};
  for (const m of movimientos) if (m.producto) stockPorProd[m.producto] = m.stock_final;
  const materiaPrima = Object.entries(stockPorProd).map(([prod, stock]) => ({
    producto: prod, stock_kg: stock, minimo_kg: minimos[prod]||0,
    alerta: stock < (minimos[prod]||0),
    autonomia_dias: Math.round(stock/50*10)/10
  }));
  const prodTerminado = lotes.filter(l=>l.estado==='En TK').map(l=>({
    tk:l.tk, producto:l.producto, litros:l.litros, lote:l.id,
    supervisor:l.supervisor, dias_en_tk:1, estado:'Disponible'
  }));

  // ── CARGAS ──
  const cargas = [];
  let totalLitrosCarga=0, totalKgCarga=0, aguaTotal=0;
  const camionStats={}, operStats={};
  for (const r of factCargas) {
    const idC    = safeStr(getCol(r,'ID Carga','único'));
    if (!idC || idC.length < 3) continue;
    const litros  = safeFloat(getCol(r,'cargada','Litros','Cantidad'));
    const agua    = safeFloat(getCol(r,'agua','Agua','Volumen agua'));
    const camion  = safeStr(getCol(r,'Camión','Camion','patente'));
    const oper    = safeStr(getCol(r,'Operador','operador'));
    const destino = safeStr(getCol(r,'Destino','destino','Rajo'));
    const lote    = safeStr(getCol(r,'Lote Producción','Lote','lote'));
    const tk      = safeStr(getCol(r,'TK','Estiba'));
    const fecha   = fechaStr(getCol(r,'Fecha'));
    const prod    = safeStr(getCol(r,'Producto'));
    // Buscar kg del lote asociado
    const loteData = lotes.find(l => l.id === lote);
    const kgCarga  = loteData ? loteData.kg : 0;
    totalLitrosCarga += litros;
    totalKgCarga += kgCarga;
    aguaTotal += agua;
    if (camion) { if(!camionStats[camion]) camionStats[camion]={cargas:0,litros:0,kg:0}; camionStats[camion].cargas++; camionStats[camion].litros+=litros; camionStats[camion].kg+=kgCarga; }
    if (oper)   { if(!operStats[oper])   operStats[oper]={cargas:0,litros:0,kg:0};   operStats[oper].cargas++;   operStats[oper].litros+=litros; operStats[oper].kg+=kgCarga; }
    cargas.push({ id:idC, fecha, hora:'', producto:prod, tk, lote, camion, operador:oper,
      litros, kg:kgCarga, agua, duracion:0, destino });
  }
  const porCamion = Object.entries(camionStats).map(([k,v])=>({
    camion:k, ...v, pct:Math.round(totalLitrosCarga?v.litros/totalLitrosCarga*100:0)
  }));
  const porOperador = Object.entries(operStats).map(([k,v])=>({operador:k,...v}));

  // ── ZONAS con superficie desde DIM_ZONA ──
  // DIM_ZONA: col A = "Destino (Rajo + Zona)", col F = "Superficie (m2)"
  const superficies = {};
  const zonaDestinos = {}; // mapa nombre_dim → nombre_cargas
  for (const d of dimZona) {
    const nombre = safeStr(getCol(d,'Destino','Rajo','destino'));
    const sup    = safeFloat(getCol(d,'Superficie','m2','superficie'));
    if (nombre && sup > 0) superficies[nombre] = sup;
  }

  // Función para buscar superficie haciendo match flexible entre nombres
  function buscarSuperficie(destino) {
    // Match exacto
    if (superficies[destino]) return superficies[destino];
    // Match parcial — buscar zona que contenga palabras clave del destino
    for (const [k, v] of Object.entries(superficies)) {
      const kLow = k.toLowerCase();
      const dLow = destino.toLowerCase();
      // Extraer número de zona para comparar
      const zonaNumK = kLow.match(/zona\s*(\d+)/)?.[1];
      const zonaNumD = dLow.match(/zona\s*(\d+)/)?.[1];
      if (zonaNumK && zonaNumD && zonaNumK === zonaNumD) return v;
    }
    return 0;
  }

  const zonaStats = {};
  for (const c of cargas) {
    if (!c.destino) continue;
    if (!zonaStats[c.destino]) zonaStats[c.destino] = { litros:0, kg:0, agua:0 };
    zonaStats[c.destino].litros += c.litros;
    zonaStats[c.destino].kg    += c.kg;
    zonaStats[c.destino].agua  += c.agua;
  }
  const totalZonasL = Object.values(zonaStats).reduce((s,v)=>s+v.litros, 0);
  const porZona = Object.entries(zonaStats)
    .map(([k,v]) => {
      const sup = buscarSuperficie(k);
      return {
        zona: k,
        litros: v.litros,
        kg: v.kg,
        agua: v.agua,
        total: v.litros,
        pct: Math.round(totalZonasL ? v.litros/totalZonasL*100 : 0),
        superficie_m2: sup,
        litros_m2: sup > 0 ? Math.round(v.litros/sup*100)/100 : null,
        kg_m2: sup > 0 ? Math.round(v.kg/sup*1000)/1000 : null,
        agua_m2: sup > 0 ? Math.round(v.agua/sup*100)/100 : null
      };
    })
    .sort((a,b)=>b.litros-a.litros);

  return {
    meta:{ ultima_actualizacion:new Date().toISOString().slice(0,19), periodo:'Mayo 2026', operacion:'BSN Centinela' },
    produccion:{ total_litros:totalLitrosProd, total_kg:totalKgProd, rendimiento_promedio:rendProm,
      duracion_promedio_min:0, lotes_total:lotes.length, lotes_despachados:lotesDesp,
      lotes_pendientes:lotesPend, lotes },
    inventario:{ materia_prima:materiaPrima, producto_terminado_tk:prodTerminado, movimientos },
    cargas:{ total_litros:totalLitrosCarga, total_kg:totalKgCarga, agua_total:aguaTotal,
      duracion_promedio_min:0, cargas, por_camion:porCamion, por_operador:porOperador },
    trazabilidad:{ lotes_completos:lotesDesp, lotes_total:lotes.length,
      lotes_pendientes:lotesPend, tiempo_promedio_dias:1.8,
      responsables:[...new Set(lotes.map(l=>l.supervisor).filter(Boolean))] },
    zonas:{ por_zona:porZona }
  };
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-cache');
  try {
    const response = await fetch(SHEET_URL, { headers:{'User-Agent':'Mozilla/5.0'}, redirect:'follow' });
    if (!response.ok) throw new Error(`Google Sheets respondió ${response.status}`);
    const buffer = Buffer.from(await response.arrayBuffer());
    res.status(200).json(procesarExcel(buffer));
  } catch (err) {
    console.error(err.message);
    res.status(500).json({ error: err.message });
  }
}
