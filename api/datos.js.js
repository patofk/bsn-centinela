import fetch from 'node-fetch';
import * as XLSX from 'xlsx';

const SHEET_ID = '1tpgQbCpXav_P6_d28TVKLezXmSeFu7BBbc120FtGdFY';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;

function safeFloat(val, def = 0) {
  const n = parseFloat(val);
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
    const date = XLSX.SSF.parse_date_code(val);
    if (!date || date.y < 2000) return '';
    return `${date.y}-${String(date.m).padStart(2,'0')}-${String(date.d).padStart(2,'0')}`;
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
  if (!wb.SheetNames.includes(nombre)) return [];
  return XLSX.utils.sheet_to_json(wb.Sheets[nombre], { defval: null, raw: true });
}

function procesarExcel(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const factProd   = leerHoja(wb, 'FACT_PRODUCCION');
  const factCargas = leerHoja(wb, 'FACT_CARGAS');
  const factInv    = leerHoja(wb, 'FACT_INVENTARIO');
  const dimProd    = leerHoja(wb, 'DIM_PRODUCTO');

  const lotes = [];
  let totalLitrosProd = 0, totalKgProd = 0;
  const duraciones = [];
  for (const r of factProd) {
    const idLote = safeStr(getCol(r,'ID Lote','id_lote','ID'));
    const producto = safeStr(getCol(r,'Producto'));
    const superv = safeStr(getCol(r,'Supervisor'));
    const tk = safeStr(getCol(r,'TK','Estanque'));
    const litros = safeFloat(getCol(r,'Litros','Volumen'));
    const kg = safeFloat(getCol(r,'kg','Kg','KG'));
    const duracion = safeFloat(getCol(r,'Duración','Duracion','Min'));
    const fecha = fechaStr(getCol(r,'Fecha'));
    if (!idLote) continue;
    totalLitrosProd += litros; totalKgProd += kg;
    if (duracion > 0) duraciones.push(duracion);
    const tieneCarga = factCargas.some(c =>
      Object.entries(c).some(([k,v]) => k.toLowerCase().includes('lote') && safeStr(v) === idLote));
    lotes.push({ id:idLote, producto, supervisor:superv, tk, litros, kg, duracion, fecha,
      estado: tieneCarga ? 'Despachado' : 'En TK' });
  }
  const rendProm = totalLitrosProd > 0 ? Math.round(totalKgProd/totalLitrosProd*10)/10 : 0;
  const durProm = duraciones.length > 0 ? Math.round(duraciones.reduce((a,b)=>a+b,0)/duraciones.length) : 0;
  const lotesPend = lotes.filter(l=>l.estado==='En TK').length;
  const lotesDesp = lotes.length - lotesPend;

  const minimos = {};
  for (const d of dimProd) {
    const prod = safeStr(getCol(d,'Producto','nombre','ID'));
    if (prod) minimos[prod] = safeFloat(getCol(d,'Mínimo','Minimo','min'));
  }
  const movimientos = [];
  for (const r of factInv) {
    const prod = safeStr(getCol(r,'Producto'));
    if (!prod) continue;
    movimientos.push({ fecha:fechaStr(getCol(r,'Fecha')), producto:prod,
      tipo:safeStr(getCol(r,'Tipo','Movimiento')),
      kg:safeFloat(getCol(r,'kg','Kg','KG','Cantidad')),
      responsable:safeStr(getCol(r,'Responsable','Operador')),
      stock_final:safeFloat(getCol(r,'Stock','stock')) });
  }
  const stockPorProd = {};
  for (const m of movimientos) if (m.producto) stockPorProd[m.producto] = m.stock_final;
  const materiaPrima = Object.entries(stockPorProd).map(([prod,stock]) => ({
    producto:prod, stock_kg:stock, minimo_kg:minimos[prod]||0,
    alerta:stock<(minimos[prod]||0), autonomia_dias:Math.round(stock/50*10)/10 }));
  const prodTerminado = lotes.filter(l=>l.estado==='En TK').map(l=>({
    tk:l.tk, producto:l.producto, litros:l.litros, lote:l.id,
    supervisor:l.supervisor, dias_en_tk:1, estado:'Disponible' }));

  const cargas = [];
  let totalLitrosCarga=0, aguaTotal=0;
  const durCargas=[], camionStats={}, operStats={};
  for (const r of factCargas) {
    const idC = safeStr(getCol(r,'ID Carga','ID'));
    if (!idC) continue;
    const litros = safeFloat(getCol(r,'Litros','Volumen'));
    const agua = safeFloat(getCol(r,'Agua'));
    const dur = safeFloat(getCol(r,'Duración','Duracion'));
    const camion = safeStr(getCol(r,'Camión','Camion'));
    const oper = safeStr(getCol(r,'Operador'));
    totalLitrosCarga += litros; aguaTotal += agua;
    if (dur>0) durCargas.push(dur);
    if (camion) { if(!camionStats[camion]) camionStats[camion]={cargas:0,litros:0}; camionStats[camion].cargas++; camionStats[camion].litros+=litros; }
    if (oper) { if(!operStats[oper]) operStats[oper]={cargas:0,litros:0}; operStats[oper].cargas++; operStats[oper].litros+=litros; }
    cargas.push({ id:idC, fecha:fechaStr(getCol(r,'Fecha')), hora:safeStr(getCol(r,'Hora')),
      producto:safeStr(getCol(r,'Producto')), tk:safeStr(getCol(r,'TK','Estanque')),
      lote:safeStr(getCol(r,'Lote','ID Lote')), camion, operador:oper,
      litros, agua, duracion:dur, destino:safeStr(getCol(r,'Destino')) });
  }
  const durPromCarga = durCargas.length>0 ? Math.round(durCargas.reduce((a,b)=>a+b,0)/durCargas.length) : 0;
  const porCamion = Object.entries(camionStats).map(([k,v])=>({camion:k,...v,pct:Math.round(totalLitrosCarga?v.litros/totalLitrosCarga*100:0)}));
  const porOperador = Object.entries(operStats).map(([k,v])=>({operador:k,...v}));

  const zonaStats={};
  for (const c of cargas) {
    if (!c.destino) continue;
    if (!zonaStats[c.destino]) zonaStats[c.destino]={s19:0,s20:0};
    zonaStats[c.destino].s19 += c.litros;
  }
  const totalZonas = Object.values(zonaStats).reduce((s,v)=>s+v.s19+v.s20,0);
  const porZona = Object.entries(zonaStats)
    .map(([k,v])=>({zona:k,s19:v.s19,s20:v.s20,total:v.s19+v.s20,pct:Math.round(totalZonas?(v.s19+v.s20)/totalZonas*100:0)}))
    .sort((a,b)=>b.total-a.total);

  return {
    meta:{ ultima_actualizacion:new Date().toISOString().slice(0,19), periodo:'Mayo 2026', operacion:'BSN Centinela' },
    produccion:{ total_litros:totalLitrosProd, total_kg:totalKgProd, rendimiento_promedio:rendProm,
      duracion_promedio_min:durProm, lotes_total:lotes.length, lotes_despachados:lotesDesp,
      lotes_pendientes:lotesPend, lotes },
    inventario:{ materia_prima:materiaPrima, producto_terminado_tk:prodTerminado, movimientos },
    cargas:{ total_litros:totalLitrosCarga, agua_total:aguaTotal, duracion_promedio_min:durPromCarga,
      cargas, por_camion:porCamion, por_operador:porOperador },
    trazabilidad:{ lotes_completos:lotesDesp, lotes_total:lotes.length, lotes_pendientes:lotesPend,
      tiempo_promedio_dias:1.8, responsables:[...new Set(lotes.map(l=>l.supervisor).filter(Boolean))] },
    zonas:{ semanas:['S19','S20'], por_zona:porZona }
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
