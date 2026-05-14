import { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { leerSheet, escribirSheet } from "./sheets.js";

function useXLSX() { return XLSX; }

function useStore(key, init, sheetName) {
  const [val, setVal] = useState(() => {
    try { const s = localStorage.getItem(key); return s ? JSON.parse(s) : init; } catch { return init; }
  });

  const save = async (v) => {
    setVal(v);
    try { localStorage.setItem(key, JSON.stringify(v)); } catch {}
    if (sheetName) {
      try { await escribirSheet(sheetName, v); } catch(e) { console.error("Sheets error:", e); }
    }
  };

  const syncFromSheet = async () => {
    if (!sheetName) return;
    try {
      const data = await leerSheet(sheetName);
      if (data && data.length > 0) {
        setVal(data);
        try { localStorage.setItem(key, JSON.stringify(data)); } catch {}
      }
    } catch(e) { console.error("Sync error:", e); }
  };

  return [val, save, syncFromSheet];
}

const PRODUCTOS_INIT = [
  { id: "AA-001", nombre: "Termo Skinny" },
  { id: "AA-002", nombre: "Playera cuello redondo" },
  { id: "AA-003", nombre: "Playera tipo polo" },
  { id: "AA-004", nombre: "Sudadera con capucha" },
  { id: "AA-005", nombre: "Llaveros" },
];

const TEMPERATURAS_INIT = [
  { productoId: "AA-001", temp: 320, tiempo: 120, indicaciones: "Darle al termo 60 segundos primero y voltear para darle otros 60 segundos" },
  { productoId: "AA-002", temp: 320, tiempo: 20, indicaciones: "Planchar antes para quitar humedad y después 5 segundos para corregir imperfecciones" },
  { productoId: "AA-003", temp: 320, tiempo: 20, indicaciones: "Planchar antes para quitar humedad y después 5 segundos para corregir imperfecciones" },
  { productoId: "AA-004", temp: 320, tiempo: 20, indicaciones: "Planchar antes para quitar humedad y después 5 segundos para corregir imperfecciones" },
  { productoId: "AA-005", temp: 320, tiempo: 90, indicaciones: "Planchar con presión alta" },
];

const PRECIOS_INIT = [
  { productoId: "AA-001", precioNormal: 250, precioMayoreo: 235, cantMayoreo: 8, precioSuperMayoreo: 220, cantSuperMayoreo: 16, observacion: "Termo blanco comprado en ML o Amazon" },
  { productoId: "AA-002", precioNormal: 250, precioMayoreo: 225, cantMayoreo: 12, precioSuperMayoreo: 200, cantSuperMayoreo: 25, observacion: "Playeras Yazbek o M&O" },
  { productoId: "AA-003", precioNormal: 280, precioMayoreo: 260, cantMayoreo: 12, precioSuperMayoreo: 250, cantSuperMayoreo: 25, observacion: "Playeras Yazbek o M&O" },
  { productoId: "AA-004", precioNormal: 350, precioMayoreo: 340, cantMayoreo: 12, precioSuperMayoreo: 320, cantSuperMayoreo: 25, observacion: "Playeras Yazbek o M&O" },
  { productoId: "AA-005", precioNormal: 40, precioMayoreo: 35, cantMayoreo: 20, precioSuperMayoreo: null, cantSuperMayoreo: null, observacion: "Llavero plateado" },
];

const COSTOS_INIT = [
  { productoId: "AA-001", costoProducto: 98, dtf: 0, tinta: 6, papel: 5, otros: 0 },
  { productoId: "AA-002", costoProducto: 78, dtf: 50, tinta: 0, papel: 0, otros: 0 },
  { productoId: "AA-003", costoProducto: 120, dtf: 50, tinta: 0, papel: 0, otros: 0 },
  { productoId: "AA-004", costoProducto: 213, dtf: 50, tinta: 0, papel: 0, otros: 0 },
  { productoId: "AA-005", costoProducto: 20, dtf: 0, tinta: 2, papel: 5, otros: 0 },
];

const USUARIOS = [
  { email: "adrianl.zamora26@gmail.com", password: "Zamora007", nombre: "Adrian Lopez", rol: "admin" },
  { email: "ln.anahiespinozaloera@gmail.com", password: "Anahi1992", nombre: "Anahí Espinoza", rol: "usuario" },
];

const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const fmt = (n) => n != null ? `$${Number(n).toLocaleString("es-MX", { minimumFractionDigits: 2 })}` : "N/A";
const pct = (n) => n != null ? `${(n * 100).toFixed(1)}%` : "N/A";
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2);
const today = () => new Date().toISOString().split("T")[0];

function calcCosto(c) {
  if (!c) return 0;
  return (c.costoProducto||0)+(c.dtf||0)+(c.tinta||0)+(c.papel||0)+(c.otros||0);
}
function calcMargen(precio, costo) {
  if (!precio||!costo) return null;
  return (precio-costo)/precio;
}
function diffPct(actual, anterior) {
  if (!anterior || anterior === 0) return null;
  return ((actual - anterior) / anterior) * 100;
}

function Badge({ children, color = "gray" }) {
  const colors = {
    green: "bg-emerald-100 text-emerald-700 border-emerald-200",
    red: "bg-red-100 text-red-700 border-red-200",
    blue: "bg-blue-100 text-blue-700 border-blue-200",
    amber: "bg-amber-100 text-amber-700 border-amber-200",
    gray: "bg-gray-100 text-gray-600 border-gray-200",
    purple: "bg-purple-100 text-purple-700 border-purple-200",
    rose: "bg-rose-100 text-rose-700 border-rose-200",
  };
  return <span className={`text-xs font-semibold px-2 py-0.5 rounded-full border ${colors[color]||colors.gray}`}>{children}</span>;
}

function Card({ children, className = "" }) {
  return <div className={`bg-white rounded-2xl shadow-sm border border-gray-100 ${className}`}>{children}</div>;
}

function Btn({ children, onClick, variant = "primary", size = "md", className = "", disabled = false }) {
  const base = "font-semibold rounded-xl transition-all duration-150 focus:outline-none focus:ring-2 focus:ring-offset-1";
  const sizes = { sm: "px-3 py-1.5 text-xs", md: "px-4 py-2 text-sm", lg: "px-6 py-3 text-base" };
  const variants = {
    primary: "bg-rose-500 hover:bg-rose-600 text-white focus:ring-rose-400 shadow-sm",
    secondary: "bg-gray-100 hover:bg-gray-200 text-gray-700 focus:ring-gray-300",
    ghost: "bg-transparent hover:bg-gray-100 text-gray-600 focus:ring-gray-200",
    success: "bg-emerald-500 hover:bg-emerald-600 text-white focus:ring-emerald-400",
  };
  return (
    <button onClick={onClick} disabled={disabled}
      className={`${base} ${sizes[size]} ${variants[variant]||variants.primary} ${disabled?"opacity-50 cursor-not-allowed":""} ${className}`}>
      {children}
    </button>
  );
}

function Input({ label, value, onChange, type="text", placeholder="", required=false, className="" }) {
  return (
    <div className={`flex flex-col gap-1 ${className}`}>
      {label && <label className="text-xs font-semibold text-gray-500 uppercase tracking-wide">{label}{required&&<span className="text-rose-400 ml-0.5">*</span>}</label>}
      <input type={type} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
        className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-rose-300 bg-gray-50 transition"/>
    </div>
  );
}

function Select({ label, value, onChange, options, placeholder="Seleccionar...", required=false, className="" }) {
  return (
    <div className={`flex flex-col gap-1 ${className}`}>
      {label && <label className="text-xs font-semibold text-gray-500 uppercase tracking-wide">{label}{required&&<span className="text-rose-400 ml-0.5">*</span>}</label>}
      <select value={value} onChange={e=>onChange(e.target.value)}
        className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-rose-300 bg-gray-50 transition">
        <option value="">{placeholder}</option>
        {options.map(o=><option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    </div>
  );
}

function Modal({ title, children, onClose }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between p-5 border-b border-gray-100">
          <h3 className="font-bold text-gray-800 text-lg">{title}</h3>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-2xl leading-none">&times;</button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Login({ onLogin }) {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [err, setErr] = useState("");
  const handle = () => {
    const u = USUARIOS.find(u=>u.email===email&&u.password===pass);
    if (u) { onLogin(u); setErr(""); } else setErr("Correo o contraseña incorrectos");
  };
  return (
    <div className="min-h-screen flex items-center justify-center" style={{background:"linear-gradient(135deg,#fff5f5 0%,#fce7e7 50%,#fff 100%)"}}>
      <div className="w-full max-w-sm">
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-rose-500 shadow-lg mb-4">
            <span className="text-white text-2xl font-black">A</span>
          </div>
          <h1 className="text-3xl font-black text-gray-800 tracking-tight">AdriAnis</h1>
          <p className="text-gray-400 text-sm mt-1">Sistema de Gestión</p>
        </div>
        <Card className="p-6 shadow-xl">
          <div className="flex flex-col gap-4">
            <Input label="Correo" value={email} onChange={setEmail} type="email" placeholder="tu@correo.com"/>
            <Input label="Contraseña" value={pass} onChange={setPass} type="password" placeholder="••••••••"/>
            {err && <p className="text-red-500 text-xs text-center">{err}</p>}
            <Btn onClick={handle} size="lg" className="w-full mt-1">Iniciar Sesión</Btn>
          </div>
        </Card>
      </div>
    </div>
  );
}

function Dashboard({ ingresos, gastos, costos, precios }) {
  const now = new Date();
  const [modo, setModo] = useState("mes");
  const [anio, setAnio] = useState(now.getFullYear());
  const [mes, setMes] = useState(now.getMonth());

  const filtrar = (arr, a, m, esMes) => {
    if (esMes === "todo") return arr;
    return arr.filter(r => {
      if (!r.fecha) return false;
      const [ry, rm] = r.fecha.split("-").map(Number);
      if (esMes) return ry === a && (rm - 1) === m;
      return ry === a;
    });
  };

  const ingActual   = useMemo(() => filtrar(ingresos, anio, mes, modo==="mes" ? true : modo==="año" ? false : "todo"), [ingresos, anio, mes, modo]);
  const gasActual   = useMemo(() => filtrar(gastos,   anio, mes, modo==="mes" ? true : modo==="año" ? false : "todo"), [gastos,   anio, mes, modo]);
  const ingAnterior = useMemo(() => {
    if (modo==="todo") return [];
    if (modo==="mes") { const pm=mes===0?11:mes-1, pa=mes===0?anio-1:anio; return filtrar(ingresos,pa,pm,true); }
    return filtrar(ingresos,anio-1,mes,false);
  }, [ingresos,anio,mes,modo]);
  const gasAnterior = useMemo(() => {
    if (modo==="todo") return [];
    if (modo==="mes") { const pm=mes===0?11:mes-1, pa=mes===0?anio-1:anio; return filtrar(gastos,pa,pm,true); }
    return filtrar(gastos,anio-1,mes,false);
  }, [gastos,anio,mes,modo]);

  const suma = arr => arr.reduce((s,r)=>s+(r.precioFinal||0),0);
  const totalIng = suma(ingActual), totalGas = suma(gasActual);
  const ganancia = totalIng - totalGas;
  const prevIng  = suma(ingAnterior), prevGas = suma(gasAnterior);
  const prevGan  = prevIng - prevGas;

  // Utilidades desde gastos
  const utilidades = useMemo(() => {
    const filtrar = (arr, texto) => arr.filter(r => {
      const campos = [r.nombreProducto, r.concepto, r.proveedor, r.observacion].map(v=>(v||"").toLowerCase());
      return campos.some(c => c.includes(texto.toLowerCase()));
    }).reduce((s,r)=>s+(r.precioFinal||0),0);
    return {
      adrian: filtrar(gasActual, "utilidad adrián") || filtrar(gasActual, "utilidad adrian"),
      anahi:  filtrar(gasActual, "utilidad anahí")  || filtrar(gasActual, "utilidad anahi"),
      total:  filtrar(gasActual, "utilidad"),
    };
  },[gasActual]);
  const utilPrev = useMemo(() => {
    const filtrar = (arr, texto) => arr.filter(r => {
      const campos = [r.nombreProducto, r.concepto, r.proveedor, r.observacion].map(v=>(v||"").toLowerCase());
      return campos.some(c => c.includes(texto.toLowerCase()));
    }).reduce((s,r)=>s+(r.precioFinal||0),0);
    return {
      adrian: filtrar(gasAnterior, "utilidad adrián") || filtrar(gasAnterior, "utilidad adrian"),
      anahi:  filtrar(gasAnterior, "utilidad anahí")  || filtrar(gasAnterior, "utilidad anahi"),
      total:  filtrar(gasAnterior, "utilidad"),
    };
  },[gasAnterior]);

  const aniosDisp = useMemo(() => {
    const ys = new Set([...[...ingresos,...gastos].map(r=>r.fecha?.split("-")[0]).filter(Boolean).map(Number), now.getFullYear()]);
    return [...ys].sort((a,b)=>b-a);
  }, [ingresos,gastos]);

  const statusIng = useMemo(() => {
    const g = {"En proceso":0,"Entregado":0,"Cancelado":0};
    ingActual.forEach(r=>{ if(g[r.estatus]!==undefined) g[r.estatus]++; });
    return g;
  },[ingActual]);

  const pagosIng = useMemo(() => {
    const cobrado   = ingActual.filter(r=>r.pago==="Pagado").reduce((s,r)=>s+(r.precioFinal||0),0);
    const pendiente = ingActual.filter(r=>r.pago==="Pendiente").reduce((s,r)=>s+(r.precioFinal||0),0);
    const anticipo  = ingActual.filter(r=>r.pago==="Anticipo").reduce((s,r)=>s+(r.precioFinal||0),0);
    return {cobrado,pendiente,anticipo};
  },[ingActual]);

  const topClientes = useMemo(() => {
    const m={};
    ingActual.forEach(r=>{ if(r.cliente) m[r.cliente]=(m[r.cliente]||0)+(r.precioFinal||0); });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,5);
  },[ingActual]);

  const topProveedores = useMemo(() => {
    const m={};
    gasActual.forEach(r=>{ if(r.proveedor) m[r.proveedor]=(m[r.proveedor]||0)+(r.precioFinal||0); });
    return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,5);
  },[gasActual]);

  const topProductos = useMemo(() => {
    const m={};
    ingActual.forEach(r=>{ const k=r.nombreProducto||r.codigoProducto||"Sin nombre"; if(!m[k]) m[k]={qty:0,monto:0}; m[k].qty+=(r.cantidad||1); m[k].monto+=(r.precioFinal||0); });
    return Object.entries(m).sort((a,b)=>b[1].qty-a[1].qty).slice(0,5);
  },[ingActual]);

  const analisisPrecios = useMemo(() => {
    const tipos={normal:{qty:0,monto:0,costo:0},mayoreo:{qty:0,monto:0,costo:0},super:{qty:0,monto:0,costo:0},manual:{qty:0,monto:0,costo:0}};
    ingActual.forEach(r=>{
      const costoItem=costos.find(c=>c.productoId===r.codigoProducto);
      const costoU=calcCosto(costoItem);
      const pr=precios.find(p=>p.productoId===r.codigoProducto);
      const qty=r.cantidad||1, monto=r.precioFinal||0, costoTotal=costoU*qty;
      let tipo="manual";
      if(pr){ const pu=Number(r.precioUnitario); if(pu===pr.precioSuperMayoreo) tipo="super"; else if(pu===pr.precioMayoreo) tipo="mayoreo"; else if(pu===pr.precioNormal) tipo="normal"; }
      tipos[tipo].qty+=qty; tipos[tipo].monto+=monto; tipos[tipo].costo+=costoTotal;
    });
    return tipos;
  },[ingActual,costos,precios]);

  const totalAn = Object.values(analisisPrecios).reduce((s,t)=>({qty:s.qty+t.qty,monto:s.monto+t.monto,costo:s.costo+t.costo}),{qty:0,monto:0,costo:0});

  const periodoLabel = modo==="mes" ? `${MESES[mes]} ${anio}` : modo==="año" ? `Año ${anio}` : "Todo el historial";

  const Chip = ({active,onClick,children}) => (
    <button onClick={onClick} className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-all ${active?"bg-rose-500 text-white shadow":"bg-gray-100 text-gray-500 hover:bg-gray-200"}`}>{children}</button>
  );

  const KPI = ({label,actual,prev,color="rose"}) => {
    const diff = modo==="todo" ? null : diffPct(actual,prev);
    const sube = diff>0;
    const bgs = {rose:"border-rose-200 bg-rose-50",emerald:"border-emerald-200 bg-emerald-50",blue:"border-blue-200 bg-blue-50",amber:"border-amber-200 bg-amber-50",purple:"border-purple-200 bg-purple-50"};
    const txs = {rose:"text-rose-600",emerald:"text-emerald-600",blue:"text-blue-600",amber:"text-amber-600",purple:"text-purple-600"};
    return (
      <Card className={`p-4 border-l-4 ${bgs[color]}`}>
        <p className="text-xs font-semibold text-gray-400 uppercase tracking-wider mb-1">{label}</p>
        <p className={`text-xl font-black ${txs[color]}`}>{fmt(actual)}</p>
        <div className="flex items-center gap-1 mt-1 flex-wrap">
          {diff===null
            ? <span className="text-xs text-gray-300">{modo==="todo"?"Historial completo":"Sin período anterior"}</span>
            : <><span className={`text-xs font-bold ${sube?"text-emerald-600":"text-red-500"}`}>{sube?"+":""}{diff.toFixed(1)}%</span><span className="text-xs text-gray-400">vs anterior ({fmt(prev)})</span></>
          }
        </div>
      </Card>
    );
  };

  const TopList = ({items,colorClass,montoKey}) => (
    items.length===0
      ? <p className="text-xs text-gray-300 text-center py-3">Sin datos en este período</p>
      : items.map(([nombre,val],i)=>{
          const monto = montoKey ? val[montoKey] : val;
          const qty   = montoKey ? val.qty : null;
          return (
            <div key={nombre} className="flex items-center gap-3 py-2.5 border-b border-gray-50 last:border-0">
              <span className={`w-6 h-6 rounded-full ${colorClass} text-xs font-black flex items-center justify-center flex-shrink-0`}>{i+1}</span>
              <span className="text-sm text-gray-700 flex-1 truncate">{nombre}</span>
              <div className="text-right">
                <p className="text-sm font-bold text-gray-800">{fmt(monto)}</p>
                {qty!==null && <p className="text-xs text-gray-400">{qty} pzs</p>}
              </div>
            </div>
          );
        })
  );

  return (
    <div className="space-y-6">
      {/* Filtros */}
      <div className="flex flex-col gap-3">
        <div>
          <h2 className="text-2xl font-black text-gray-800">Dashboard</h2>
          <p className="text-gray-400 text-sm">Período: <strong className="text-gray-700">{periodoLabel}</strong></p>
        </div>
        <div className="flex flex-wrap gap-2 items-center">
          <div className="flex gap-1 bg-gray-100 p-1 rounded-xl">
            <Chip active={modo==="mes"} onClick={()=>setModo("mes")}>Por Mes</Chip>
            <Chip active={modo==="año"} onClick={()=>setModo("año")}>Por Año</Chip>
            <Chip active={modo==="todo"} onClick={()=>setModo("todo")}>Historial</Chip>
          </div>
          {modo==="mes" && (
            <select value={mes} onChange={e=>setMes(Number(e.target.value))}
              className="border border-gray-200 rounded-xl px-3 py-1.5 text-xs font-semibold bg-white focus:outline-none focus:ring-2 focus:ring-rose-300">
              {MESES.map((m,i)=><option key={i} value={i}>{m}</option>)}
            </select>
          )}
          {modo!=="todo" && (
            <select value={anio} onChange={e=>setAnio(Number(e.target.value))}
              className="border border-gray-200 rounded-xl px-3 py-1.5 text-xs font-semibold bg-white focus:outline-none focus:ring-2 focus:ring-rose-300">
              {aniosDisp.map(a=><option key={a} value={a}>{a}</option>)}
            </select>
          )}
        </div>
      </div>

      {/* ① Resumen financiero */}
      <section>
        <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-3">① Resumen Financiero</p>
        <div className="grid grid-cols-2 gap-3">
          <KPI label="Ingresos" actual={totalIng} prev={prevIng} color="emerald"/>
          <KPI label="Gastos" actual={totalGas} prev={prevGas} color="rose"/>
          <KPI label="Ganancia" actual={ganancia} prev={prevGan} color="blue"/>
          <KPI label="Utilidades Totales" actual={utilidades.total} prev={utilPrev.total} color="amber"/>
        </div>
        <div className="grid grid-cols-2 gap-3 mt-3">
          <KPI label="Utilidad Adrián" actual={utilidades.adrian} prev={utilPrev.adrian} color="purple"/>
          <KPI label="Utilidad Anahí" actual={utilidades.anahi} prev={utilPrev.anahi} color="rose"/>
        </div>
      </section>

      {/* ② Status de Ingresos */}
      <section>
        <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-3">② Status de Ingresos</p>
        <div className="grid grid-cols-3 gap-3 mb-3">
          {[{label:"En Proceso",val:statusIng["En proceso"],color:"text-amber-500"},{label:"Entregados",val:statusIng["Entregado"],color:"text-emerald-500"},{label:"Cancelados",val:statusIng["Cancelado"],color:"text-red-500"}]
            .map(({label,val,color})=>(
              <Card key={label} className="p-3 text-center">
                <p className="text-xs text-gray-400 font-semibold mb-1">{label}</p>
                <p className={`text-2xl font-black ${color}`}>{val}</p>
                <p className="text-xs text-gray-300">pedidos</p>
              </Card>
            ))}
        </div>
        <Card className="p-4">
          <p className="text-xs font-bold text-gray-500 mb-3">Estado de Cobro</p>
          <div className="space-y-3">
            {[{label:"Cobrado",val:pagosIng.cobrado,hex:"#10b981"},{label:"Pendiente de Cobro",val:pagosIng.pendiente,hex:"#f59e0b"},{label:"Anticipo",val:pagosIng.anticipo,hex:"#3b82f6"}]
              .map(({label,val,hex})=>{
                const bar = totalIng>0?(val/totalIng)*100:0;
                return (
                  <div key={label}>
                    <div className="flex justify-between text-xs mb-1">
                      <span className="text-gray-500">{label}</span>
                      <span className="font-bold text-gray-700">{fmt(val)}</span>
                    </div>
                    <div className="h-2 bg-gray-100 rounded-full overflow-hidden">
                      <div className="h-full rounded-full" style={{width:`${bar}%`,background:hex}}/>
                    </div>
                  </div>
                );
              })}
          </div>
        </Card>
      </section>

      {/* ③ Rankings */}
      <section>
        <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-3">③ Rankings</p>
        <div className="space-y-3">
          <Card className="p-4">
            <p className="text-xs font-bold text-gray-500 mb-2">Top Clientes</p>
            <TopList items={topClientes} colorClass="bg-rose-100 text-rose-600"/>
          </Card>
          <Card className="p-4">
            <p className="text-xs font-bold text-gray-500 mb-2">Top Proveedores</p>
            <TopList items={topProveedores} colorClass="bg-blue-100 text-blue-600"/>
          </Card>
          <Card className="p-4">
            <p className="text-xs font-bold text-gray-500 mb-2">Top Productos Vendidos</p>
            <TopList items={topProductos} colorClass="bg-emerald-100 text-emerald-600" montoKey="monto"/>
          </Card>
        </div>
      </section>

      {/* ④ Análisis Financiero */}
      <section>
        <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-3">④ Análisis por Tipo de Venta</p>
        <Card className="overflow-hidden">
          <div className="p-4">
            <div className="grid grid-cols-4 gap-1 text-xs font-bold text-gray-400 uppercase mb-3 pb-2 border-b border-gray-100">
              <span>Tipo</span><span className="text-right">Piezas</span><span className="text-right">Monto</span><span className="text-right">% Costo</span>
            </div>
            {[{key:"normal",label:"Normal",color:"gray"},{key:"mayoreo",label:"Mayoreo",color:"blue"},{key:"super",label:"S. Mayoreo",color:"rose"},{key:"manual",label:"Manual/Otro",color:"amber"}]
              .map(({key,label,color})=>{
                const t=analisisPrecios[key];
                const costoP=t.monto>0?(t.costo/t.monto)*100:0;
                return (
                  <div key={key} className="grid grid-cols-4 gap-1 items-center py-2.5 border-b border-gray-50 last:border-0">
                    <span><Badge color={color}>{label}</Badge></span>
                    <span className="text-right text-sm font-bold text-gray-700">{t.qty}</span>
                    <span className="text-right text-sm font-bold text-gray-700">{fmt(t.monto)}</span>
                    <span className={`text-right text-xs font-bold ${costoP>70?"text-red-500":costoP>50?"text-amber-500":"text-emerald-600"}`}>
                      {t.monto>0?`${costoP.toFixed(1)}%`:"—"}
                    </span>
                  </div>
                );
              })}
          </div>
          <div className="px-4 py-3 bg-gray-800 rounded-b-2xl">
            <div className="grid grid-cols-4 gap-1">
              <span className="text-xs font-black text-white">TOTAL</span>
              <span className="text-right text-xs font-black text-white">{totalAn.qty} pzs</span>
              <span className="text-right text-xs font-black text-white">{fmt(totalAn.monto)}</span>
              <span className={`text-right text-xs font-black ${totalAn.monto>0&&(totalAn.costo/totalAn.monto)*100>70?"text-red-400":totalAn.monto>0&&(totalAn.costo/totalAn.monto)*100>50?"text-amber-400":"text-emerald-400"}`}>
                {totalAn.monto>0?`${((totalAn.costo/totalAn.monto)*100).toFixed(1)}%`:"—"}
              </span>
            </div>
            <p className="text-xs text-gray-500 mt-1">% Costo = costo total ÷ monto de venta</p>
          </div>
        </Card>
      </section>
    </div>
  );
}

function Productos({ productos, setProductos }) {
  const [modal, setModal] = useState(false);
  const [form, setForm] = useState({id:"",nombre:""});
  const [editIdx, setEditIdx] = useState(null);
  const openNew = () => { setForm({id:"",nombre:""}); setEditIdx(null); setModal(true); };
  const openEdit = (i) => { setForm({...productos[i]}); setEditIdx(i); setModal(true); };
  const save = () => {
    if (!form.id||!form.nombre) return;
    const arr=[...productos]; if(editIdx!==null) arr[editIdx]=form; else arr.push(form);
    setProductos(arr); setModal(false);
  };
  const del = (i) => { if(confirm("¿Eliminar producto?")) setProductos(productos.filter((_,j)=>j!==i)); };
  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between">
        <div><h2 className="text-2xl font-black text-gray-800">Productos</h2><p className="text-gray-400 text-sm">Catálogo base</p></div>
        <Btn onClick={openNew} size="sm">+ Nuevo</Btn>
      </div>
      <Card>
        {productos.length===0 ? <p className="text-center text-gray-400 py-8 text-sm">Sin productos</p>
          : productos.map((p,i)=>(
            <div key={i} className="flex items-center justify-between p-4 border-b border-gray-50 last:border-0 hover:bg-gray-50 transition">
              <div className="flex items-center gap-3"><Badge color="blue">{p.id}</Badge><span className="font-semibold text-gray-700">{p.nombre}</span></div>
              <div className="flex gap-2"><Btn onClick={()=>openEdit(i)} variant="ghost" size="sm">✏️</Btn><Btn onClick={()=>del(i)} variant="ghost" size="sm">🗑️</Btn></div>
            </div>
          ))}
      </Card>
      {modal&&(
        <Modal title={editIdx!==null?"Editar Producto":"Nuevo Producto"} onClose={()=>setModal(false)}>
          <div className="flex flex-col gap-4">
            <Input label="Código" value={form.id} onChange={v=>setForm({...form,id:v})} placeholder="AA-001" required/>
            <Input label="Nombre" value={form.nombre} onChange={v=>setForm({...form,nombre:v})} placeholder="Taza cerámica" required/>
            <div className="flex gap-3 justify-end pt-2"><Btn variant="secondary" onClick={()=>setModal(false)}>Cancelar</Btn><Btn onClick={save}>Guardar</Btn></div>
          </div>
        </Modal>
      )}
    </div>
  );
}

function Temperaturas({ temperaturas, setTemperaturas, productos }) {
  const [modal, setModal] = useState(false);
  const [form, setForm] = useState({productoId:"",temp:"",tiempo:"",indicaciones:""});
  const [editIdx, setEditIdx] = useState(null);
  const pNombre = (id) => productos.find(p=>p.id===id)?.nombre||id;
  const openNew = () => { setForm({productoId:"",temp:"",tiempo:"",indicaciones:""}); setEditIdx(null); setModal(true); };
  const openEdit = (i) => { setForm({...temperaturas[i]}); setEditIdx(i); setModal(true); };
  const save = () => {
    if (!form.productoId) return;
    const arr=[...temperaturas];
    const entry={...form,temp:Number(form.temp),tiempo:Number(form.tiempo)};
    if(editIdx!==null) arr[editIdx]=entry; else arr.push(entry);
    setTemperaturas(arr); setModal(false);
  };
  const del = (i) => { if(confirm("¿Eliminar?")) setTemperaturas(temperaturas.filter((_,j)=>j!==i)); };
  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between">
        <div><h2 className="text-2xl font-black text-gray-800">Temperaturas</h2><p className="text-gray-400 text-sm">Parámetros de planchado por producto</p></div>
        <Btn onClick={openNew} size="sm">+ Nuevo</Btn>
      </div>
      <div className="space-y-3">
        {temperaturas.length===0 ? <Card><p className="text-center text-gray-400 py-8 text-sm">Sin registros</p></Card>
          : temperaturas.map((t,i)=>(
            <Card key={i} className="p-4">
              <div className="flex items-start justify-between mb-3">
                <div><Badge color="purple">{t.productoId}</Badge><span className="font-bold text-gray-800 ml-2">{pNombre(t.productoId)}</span></div>
                <div className="flex gap-1"><Btn onClick={()=>openEdit(i)} variant="ghost" size="sm">✏️</Btn><Btn onClick={()=>del(i)} variant="ghost" size="sm">🗑️</Btn></div>
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div className="bg-orange-50 rounded-xl p-3 text-center">
                  <p className="text-xs text-orange-400 font-semibold">Temperatura</p>
                  <p className="text-2xl font-black text-orange-600">{t.temp}°C</p>
                </div>
                <div className="bg-blue-50 rounded-xl p-3 text-center">
                  <p className="text-xs text-blue-400 font-semibold">Tiempo</p>
                  <p className="text-2xl font-black text-blue-600">{t.tiempo}s</p>
                </div>
              </div>
              <p className="text-xs text-gray-500 mt-3 bg-gray-50 rounded-lg p-2">📋 {t.indicaciones}</p>
            </Card>
          ))}
      </div>
      {modal&&(
        <Modal title={editIdx!==null?"Editar Temperatura":"Nueva Temperatura"} onClose={()=>setModal(false)}>
          <div className="flex flex-col gap-4">
            <Select label="Producto" value={form.productoId} onChange={v=>setForm({...form,productoId:v})} options={productos.map(p=>({value:p.id,label:`${p.id} - ${p.nombre}`}))} required/>
            <div className="grid grid-cols-2 gap-3">
              <Input label="Temperatura °C" type="number" value={form.temp} onChange={v=>setForm({...form,temp:v})} placeholder="320" required/>
              <Input label="Tiempo (segundos)" type="number" value={form.tiempo} onChange={v=>setForm({...form,tiempo:v})} placeholder="120" required/>
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-xs font-semibold text-gray-500 uppercase tracking-wide">Indicaciones</label>
              <textarea value={form.indicaciones} onChange={e=>setForm({...form,indicaciones:e.target.value})}
                className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-rose-300 bg-gray-50 h-24 resize-none"/>
            </div>
            <div className="flex gap-3 justify-end pt-2"><Btn variant="secondary" onClick={()=>setModal(false)}>Cancelar</Btn><Btn onClick={save}>Guardar</Btn></div>
          </div>
        </Modal>
      )}
    </div>
  );
}

function Precios({ precios, setPrecios, productos }) {
  const [modal, setModal] = useState(false);
  const [form, setForm] = useState({productoId:"",precioNormal:"",precioMayoreo:"",cantMayoreo:"",precioSuperMayoreo:"",cantSuperMayoreo:"",observacion:""});
  const [editIdx, setEditIdx] = useState(null);
  const pNombre = (id) => productos.find(p=>p.id===id)?.nombre||id;
  const openNew = () => { setForm({productoId:"",precioNormal:"",precioMayoreo:"",cantMayoreo:"",precioSuperMayoreo:"",cantSuperMayoreo:"",observacion:""}); setEditIdx(null); setModal(true); };
  const openEdit = (i) => { setForm({...precios[i],precioSuperMayoreo:precios[i].precioSuperMayoreo??'',cantSuperMayoreo:precios[i].cantSuperMayoreo??''}); setEditIdx(i); setModal(true); };
  const save = () => {
    if (!form.productoId) return;
    const arr=[...precios];
    const entry={...form,precioNormal:Number(form.precioNormal),precioMayoreo:Number(form.precioMayoreo),cantMayoreo:Number(form.cantMayoreo),precioSuperMayoreo:form.precioSuperMayoreo?Number(form.precioSuperMayoreo):null,cantSuperMayoreo:form.cantSuperMayoreo?Number(form.cantSuperMayoreo):null};
    if(editIdx!==null) arr[editIdx]=entry; else arr.push(entry);
    setPrecios(arr); setModal(false);
  };
  const del = (i) => { if(confirm("¿Eliminar?")) setPrecios(precios.filter((_,j)=>j!==i)); };
  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between">
        <div><h2 className="text-2xl font-black text-gray-800">Costos y Precios</h2><p className="text-gray-400 text-sm">Niveles por volumen</p></div>
        <Btn onClick={openNew} size="sm">+ Nuevo</Btn>
      </div>
      <div className="space-y-3">
        {precios.map((p,i)=>(
          <Card key={i} className="p-4">
            <div className="flex items-center justify-between mb-3">
              <div><Badge color="blue">{p.productoId}</Badge><span className="font-bold text-gray-800 ml-2">{pNombre(p.productoId)}</span></div>
              <div className="flex gap-1"><Btn onClick={()=>openEdit(i)} variant="ghost" size="sm">✏️</Btn><Btn onClick={()=>del(i)} variant="ghost" size="sm">🗑️</Btn></div>
            </div>
            <div className="grid grid-cols-3 gap-2">
              <div className="bg-gray-50 rounded-xl p-3 text-center"><p className="text-xs text-gray-400 font-semibold mb-1">Normal</p><p className="text-lg font-black text-gray-700">{fmt(p.precioNormal)}</p><p className="text-xs text-gray-400">1 pieza</p></div>
              <div className="bg-blue-50 rounded-xl p-3 text-center"><p className="text-xs text-blue-400 font-semibold mb-1">Mayoreo</p><p className="text-lg font-black text-blue-600">{fmt(p.precioMayoreo)}</p><p className="text-xs text-blue-300">{p.cantMayoreo}+ pzs</p></div>
              <div className="bg-rose-50 rounded-xl p-3 text-center"><p className="text-xs text-rose-400 font-semibold mb-1">S. Mayoreo</p><p className="text-lg font-black text-rose-600">{p.precioSuperMayoreo?fmt(p.precioSuperMayoreo):"N/A"}</p><p className="text-xs text-rose-300">{p.cantSuperMayoreo?`${p.cantSuperMayoreo}+ pzs`:"—"}</p></div>
            </div>
            {p.observacion&&<p className="text-xs text-gray-400 mt-2 bg-gray-50 rounded-lg p-2">📝 {p.observacion}</p>}
          </Card>
        ))}
      </div>
      {modal&&(
        <Modal title={editIdx!==null?"Editar Precios":"Nuevos Precios"} onClose={()=>setModal(false)}>
          <div className="flex flex-col gap-3">
            <Select label="Producto" value={form.productoId} onChange={v=>setForm({...form,productoId:v})} options={productos.map(p=>({value:p.id,label:`${p.id} - ${p.nombre}`}))} required/>
            <div className="grid grid-cols-2 gap-3">
              <Input label="Precio Normal" type="number" value={form.precioNormal} onChange={v=>setForm({...form,precioNormal:v})}/>
              <div/>
              <Input label="Precio Mayoreo" type="number" value={form.precioMayoreo} onChange={v=>setForm({...form,precioMayoreo:v})}/>
              <Input label="Cantidad Mayoreo" type="number" value={form.cantMayoreo} onChange={v=>setForm({...form,cantMayoreo:v})}/>
              <Input label="Precio Super Mayoreo" type="number" value={form.precioSuperMayoreo} onChange={v=>setForm({...form,precioSuperMayoreo:v})}/>
              <Input label="Cantidad S. Mayoreo" type="number" value={form.cantSuperMayoreo} onChange={v=>setForm({...form,cantSuperMayoreo:v})}/>
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-xs font-semibold text-gray-500 uppercase tracking-wide">Observación</label>
              <textarea value={form.observacion} onChange={e=>setForm({...form,observacion:e.target.value})} className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-rose-300 bg-gray-50 h-16 resize-none"/>
            </div>
            <div className="flex gap-3 justify-end pt-2"><Btn variant="secondary" onClick={()=>setModal(false)}>Cancelar</Btn><Btn onClick={save}>Guardar</Btn></div>
          </div>
        </Modal>
      )}
    </div>
  );
}

function Margen({ costos, setCostos, precios, productos }) {
  const [modal, setModal] = useState(false);
  const [form, setForm] = useState({productoId:"",costoProducto:"",dtf:"",tinta:"",papel:"",otros:""});
  const [editIdx, setEditIdx] = useState(null);
  const pNombre = (id) => productos.find(p=>p.id===id)?.nombre||id;
  const getPrecio = (id) => precios.find(p=>p.productoId===id);
  const openNew = () => { setForm({productoId:"",costoProducto:"",dtf:"",tinta:"",papel:"",otros:""}); setEditIdx(null); setModal(true); };
  const openEdit = (i) => { setForm({...costos[i]}); setEditIdx(i); setModal(true); };
  const save = () => {
    if (!form.productoId) return;
    const arr=[...costos];
    const entry={productoId:form.productoId,costoProducto:Number(form.costoProducto)||0,dtf:Number(form.dtf)||0,tinta:Number(form.tinta)||0,papel:Number(form.papel)||0,otros:Number(form.otros)||0};
    if(editIdx!==null) arr[editIdx]=entry; else arr.push(entry);
    setCostos(arr); setModal(false);
  };
  const del = (i) => { if(confirm("¿Eliminar?")) setCostos(costos.filter((_,j)=>j!==i)); };
  const Bar = ({label,precio,costo}) => {
    if (!precio) return null;
    const mg=calcMargen(precio,costo), p2=Math.max(0,Math.min(100,(mg||0)*100));
    return (
      <div className="mb-2">
        <div className="flex justify-between text-xs mb-1">
          <span className="text-gray-500">{label}</span>
          <span className={`font-bold ${mg>0.3?"text-emerald-600":mg>0.15?"text-amber-600":"text-red-500"}`}>{pct(mg)} · {fmt(precio-costo)}</span>
        </div>
        <div className="h-2 bg-gray-100 rounded-full overflow-hidden">
          <div className="h-full rounded-full" style={{width:`${p2}%`,background:mg>0.3?"#10b981":mg>0.15?"#f59e0b":"#ef4444"}}/>
        </div>
      </div>
    );
  };
  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between">
        <div><h2 className="text-2xl font-black text-gray-800">Margen</h2><p className="text-gray-400 text-sm">Rentabilidad por producto</p></div>
        <Btn onClick={openNew} size="sm">+ Nuevo</Btn>
      </div>
      <div className="space-y-3">
        {costos.map((c,i)=>{
          const total=calcCosto(c), pr=getPrecio(c.productoId);
          return (
            <Card key={i} className="p-4">
              <div className="flex items-center justify-between mb-3">
                <div><Badge color="green">{c.productoId}</Badge><span className="font-bold text-gray-800 ml-2">{pNombre(c.productoId)}</span></div>
                <div className="flex gap-1"><Btn onClick={()=>openEdit(i)} variant="ghost" size="sm">✏️</Btn><Btn onClick={()=>del(i)} variant="ghost" size="sm">🗑️</Btn></div>
              </div>
              <div className="grid grid-cols-5 gap-1 mb-3 text-center">
                {[["Producto",c.costoProducto],["DTF",c.dtf],["Tinta",c.tinta],["Papel",c.papel],["Otros",c.otros]].map(([lbl,val])=>(
                  <div key={lbl} className="bg-gray-50 rounded-lg p-2"><p className="text-xs text-gray-400">{lbl}</p><p className="text-sm font-bold text-gray-700">{fmt(val)}</p></div>
                ))}
              </div>
              <div className="bg-gray-800 rounded-xl p-3 mb-3 text-center"><p className="text-xs text-gray-400">Costo Total</p><p className="text-2xl font-black text-white">{fmt(total)}</p></div>
              {pr&&<div><p className="text-xs font-semibold text-gray-400 uppercase mb-2">Margen por tipo</p><Bar label="Normal" precio={pr.precioNormal} costo={total}/><Bar label="Mayoreo" precio={pr.precioMayoreo} costo={total}/>{pr.precioSuperMayoreo&&<Bar label="Super Mayoreo" precio={pr.precioSuperMayoreo} costo={total}/>}</div>}
            </Card>
          );
        })}
      </div>
      {modal&&(
        <Modal title={editIdx!==null?"Editar Costos":"Nuevos Costos"} onClose={()=>setModal(false)}>
          <div className="flex flex-col gap-3">
            <Select label="Producto" value={form.productoId} onChange={v=>setForm({...form,productoId:v})} options={productos.map(p=>({value:p.id,label:`${p.id} - ${p.nombre}`}))} required/>
            <div className="grid grid-cols-2 gap-3">
              <Input label="Costo Producto" type="number" value={form.costoProducto} onChange={v=>setForm({...form,costoProducto:v})}/>
              <Input label="DTF" type="number" value={form.dtf} onChange={v=>setForm({...form,dtf:v})}/>
              <Input label="Tinta" type="number" value={form.tinta} onChange={v=>setForm({...form,tinta:v})}/>
              <Input label="Papel" type="number" value={form.papel} onChange={v=>setForm({...form,papel:v})}/>
              <Input label="Otros" type="number" value={form.otros} onChange={v=>setForm({...form,otros:v})}/>
            </div>
            <div className="flex gap-3 justify-end pt-2"><Btn variant="secondary" onClick={()=>setModal(false)}>Cancelar</Btn><Btn onClick={save}>Guardar</Btn></div>
          </div>
        </Modal>
      )}
    </div>
  );
}

function Movimientos({ tipo, registros, setRegistros, productos, precios }) {
  const [modal, setModal] = useState(false);
  const [manual, setManual] = useState(false);
  const [form, setForm] = useState({});
  const [editIdx, setEditIdx] = useState(null);
  const [tipoPrecio, setTipoPrecio] = useState("normal");
  const [busqueda, setBusqueda] = useState("");
  const [orden, setOrden] = useState("desc");
  const esIngreso = tipo==="ingreso";
  const getPU = (prodId,tp) => { const pr=precios.find(p=>p.productoId===prodId); if(!pr) return ""; if(tp==="mayoreo") return pr.precioMayoreo||""; if(tp==="super") return pr.precioSuperMayoreo||""; return pr.precioNormal||""; };
  const openNew = () => { setForm({fecha:today(),nombreProducto:"",codigoProducto:"",cantidad:1,precioUnitario:"",precioFinal:"",cliente:"",proveedor:"",estatus:"En proceso",pago:"Pendiente",observacion:""}); setTipoPrecio("normal"); setManual(false); setEditIdx(null); setModal(true); };
  const openEdit = (i) => { setForm({...registros[i]}); setEditIdx(i); setManual(true); setModal(true); };
  const onProd = (id) => { const p=productos.find(x=>x.id===id), pu=getPU(id,tipoPrecio); setForm(f=>({...f,codigoProducto:id,nombreProducto:p?.nombre||"",precioUnitario:pu,precioFinal:pu?pu*(f.cantidad||1):""})); };
  const onTipo = (tp) => { setTipoPrecio(tp); if(form.codigoProducto){ const pu=getPU(form.codigoProducto,tp); setForm(f=>({...f,precioUnitario:pu,precioFinal:pu?pu*(f.cantidad||1):""})); } };
  const onQty = (v) => setForm(f=>({...f,cantidad:v,precioFinal:f.precioUnitario?f.precioUnitario*v:""}));
  const save = () => { const arr=[...registros]; const e={...form,id:form.id||uid(),cantidad:Number(form.cantidad)||1,precioUnitario:Number(form.precioUnitario)||0,precioFinal:Number(form.precioFinal)||0}; if(editIdx!==null) arr[editIdx]=e; else arr.push(e); setRegistros(arr); setModal(false); };
  const del = (idx) => { if(confirm("¿Eliminar?")) setRegistros(registros.filter((_,j)=>j!==idx)); };
  const sc = {"En proceso":"amber","Entregado":"green","Cancelado":"red"};
  const pc = {"Pagado":"green","Pendiente":"amber","Anticipo":"blue"};

  const registrosFiltrados = useMemo(() => {
    let arr = [...registros];
    if (busqueda.trim()) {
      const b = busqueda.toLowerCase();
      arr = arr.filter(r => {
        const cp = (esIngreso?r.cliente:r.proveedor)||"";
        return cp.toLowerCase().includes(b) || (r.nombreProducto||"").toLowerCase().includes(b);
      });
    }
    arr.sort((a,b) => orden==="desc"?(b.fecha||"").localeCompare(a.fecha||""):(a.fecha||"").localeCompare(b.fecha||""));
    return arr;
  }, [registros, busqueda, orden, esIngreso]);

  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between">
        <div><h2 className="text-2xl font-black text-gray-800">{esIngreso?"Ingresos":"Gastos"}</h2><p className="text-gray-400 text-sm">{registrosFiltrados.length} de {registros.length} registros</p></div>
        <Btn onClick={openNew} variant={esIngreso?"success":"primary"} size="sm">+ Nuevo</Btn>
      </div>
      <div className="flex gap-2">
        <div className="flex-1 relative">
          <input value={busqueda} onChange={e=>setBusqueda(e.target.value)}
            placeholder={`Buscar ${esIngreso?"cliente":"proveedor"} o producto...`}
            className="w-full border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-rose-300 bg-gray-50"/>
          {busqueda&&<button onClick={()=>setBusqueda("")} className="absolute right-3 top-2.5 text-gray-400 hover:text-gray-600 text-xs">✕</button>}
        </div>
        <button onClick={()=>setOrden(o=>o==="desc"?"asc":"desc")}
          className="border border-gray-200 rounded-xl px-3 py-2 text-xs font-semibold bg-gray-50 hover:bg-gray-100 transition">
          Fecha {orden==="desc"?"↓":"↑"}
        </button>
      </div>
      <Card>
        {registrosFiltrados.length===0
          ? <p className="text-center text-gray-400 py-8 text-sm">{busqueda?"Sin resultados":"Sin registros"}</p>
          : registrosFiltrados.map((r,i)=>{
            const origIdx = registros.indexOf(r);
            return (
              <div key={r.id||i} className="p-4 border-b border-gray-50 last:border-0 hover:bg-gray-50 transition">
                <div className="flex items-start justify-between mb-2">
                  <div className="flex-1 min-w-0">
                    <p className="font-black text-gray-900 text-base truncate">{(esIngreso?r.cliente:r.proveedor)||"—"}</p>
                    <div className="flex items-center gap-2 mt-0.5 flex-wrap">
                      <span className="text-sm text-gray-500">{r.nombreProducto||r.concepto||"—"}</span>
                      {r.codigoProducto&&<Badge color="blue">{r.codigoProducto}</Badge>}
                    </div>
                    <p className="text-xs text-gray-400 mt-0.5">{r.fecha}</p>
                  </div>
                  <div className="text-right ml-3 flex-shrink-0">
                    <p className={`font-black text-lg ${esIngreso?"text-emerald-600":"text-rose-500"}`}>{fmt(r.precioFinal)}</p>
                    <p className="text-xs text-gray-400">{r.cantidad} × {fmt(r.precioUnitario)}</p>
                  </div>
                </div>
                <div className="flex items-center justify-between">
                  <div className="flex gap-2"><Badge color={sc[r.estatus]||"gray"}>{r.estatus}</Badge><Badge color={pc[r.pago]||"gray"}>{r.pago}</Badge></div>
                  <div className="flex gap-1"><Btn onClick={()=>openEdit(origIdx)} variant="ghost" size="sm">✏️</Btn><Btn onClick={()=>del(origIdx)} variant="ghost" size="sm">🗑️</Btn></div>
                </div>
                {r.observacion&&<p className="text-xs text-gray-400 mt-2 bg-gray-50 rounded-lg px-2 py-1">📝 {r.observacion}</p>}
              </div>
            );
          })}
      </Card>
      {modal&&(
        <Modal title={editIdx!==null?`Editar ${esIngreso?"Ingreso":"Gasto"}`:`Nuevo ${esIngreso?"Ingreso":"Gasto"}`} onClose={()=>setModal(false)}>
          <div className="flex flex-col gap-3">
            <div className="flex gap-2 p-1 bg-gray-100 rounded-xl">
              <button onClick={()=>setManual(false)} className={`flex-1 py-1.5 text-xs font-semibold rounded-lg transition ${!manual?"bg-white shadow text-gray-800":"text-gray-400"}`}>📦 Desde catálogo</button>
              <button onClick={()=>setManual(true)} className={`flex-1 py-1.5 text-xs font-semibold rounded-lg transition ${manual?"bg-white shadow text-gray-800":"text-gray-400"}`}>✏️ Manual</button>
            </div>
            {!manual
              ? <><Select label="Producto" value={form.codigoProducto} onChange={onProd} options={productos.map(p=>({value:p.id,label:`${p.id} - ${p.nombre}`}))}/><Select label="Tipo de Precio" value={tipoPrecio} onChange={onTipo} options={[{value:"normal",label:"Normal"},{value:"mayoreo",label:"Mayoreo"},{value:"super",label:"Super Mayoreo"}]}/></>
              : <div className="grid grid-cols-2 gap-3"><Input label="Código" value={form.codigoProducto} onChange={v=>setForm({...form,codigoProducto:v})}/><Input label="Nombre" value={form.nombreProducto} onChange={v=>setForm({...form,nombreProducto:v})}/></div>
            }
            <Input label="Fecha" type="date" value={form.fecha} onChange={v=>setForm({...form,fecha:v})} required/>
            <div className="grid grid-cols-3 gap-3">
              <Input label="Cantidad" type="number" value={form.cantidad} onChange={onQty}/>
              <Input label="Precio Unitario" type="number" value={form.precioUnitario} onChange={v=>setForm({...form,precioUnitario:v,precioFinal:v*(form.cantidad||1)})}/>
              <Input label="Precio Final" type="number" value={form.precioFinal} onChange={v=>setForm({...form,precioFinal:v})}/>
            </div>
            <Input label={esIngreso?"Cliente":"Proveedor"} value={esIngreso?form.cliente:form.proveedor} onChange={v=>setForm({...form,[esIngreso?"cliente":"proveedor"]:v})}/>
            <div className="grid grid-cols-2 gap-3">
              <Select label="Estatus" value={form.estatus} onChange={v=>setForm({...form,estatus:v})} options={[{value:"En proceso",label:"En proceso"},{value:"Entregado",label:"Entregado"},{value:"Cancelado",label:"Cancelado"}]}/>
              <Select label="Pago" value={form.pago} onChange={v=>setForm({...form,pago:v})} options={[{value:"Pagado",label:"Pagado"},{value:"Pendiente",label:"Pendiente"},{value:"Anticipo",label:"Anticipo"}]}/>
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-xs font-semibold text-gray-500 uppercase tracking-wide">Observación</label>
              <textarea value={form.observacion} onChange={e=>setForm({...form,observacion:e.target.value})} className="border border-gray-200 rounded-xl px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-rose-300 bg-gray-50 h-16 resize-none"/>
            </div>
            <div className="flex gap-3 justify-end pt-2"><Btn variant="secondary" onClick={()=>setModal(false)}>Cancelar</Btn><Btn variant={esIngreso?"success":"primary"} onClick={save}>Guardar</Btn></div>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════════════════
// IMPORTAR (módulo de importación masiva desde Excel)
// ══════════════════════════════════════════════════════════════════════════════
function Importar({ setProductos, setTemperaturas, setPrecios, setCostos, setIngresos, setGastos }) {
  const XLSX = useXLSX();
  const fileRef = useRef();
  const [estado, setEstado] = useState("idle"); // idle | leyendo | preview | importado | error
  const [preview, setPreview] = useState(null);
  const [errMsg, setErrMsg] = useState("");
  const [modoImport, setModoImport] = useState("reemplazar"); // reemplazar | agregar
  const [progreso, setProgreso] = useState({});

  // ── Columnas esperadas por hoja ───────────────────────────────────────────
  const ESQUEMAS = {
    Productos:    ["Codigo","Nombre"],
    Temperaturas: ["Codigo","Temperatura","Tiempo","Indicaciones"],
    Precios:      ["Codigo","Nombre","PrecioNormal","PrecioMayoreo","CantidadMayoreo","PrecioSuperMayoreo","CantidadSuperMayoreo","Observacion"],
    Margen:       ["Codigo","CostoProducto","DTF","Tinta","Papel","Otros"],
    Ingresos:     ["Fecha","Codigo","NombreProducto","Cliente","Cantidad","PrecioUnitario","PrecioFinal","Estatus","Pago","Observacion"],
    Gastos:       ["Fecha","Codigo","NombreProducto","Proveedor","Cantidad","PrecioUnitario","PrecioFinal","Estatus","Pago","Observacion"],
  };

  // ── Parsear fecha de Excel (número o string) ──────────────────────────────
  const parseFecha = (v) => {
    if (!v) return today();
    if (typeof v === "number") {
      const d = XLSX.SSF.parse_date_code(v);
      return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
    }
    const s = String(v).trim();
    // dd/mm/yyyy → yyyy-mm-dd
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
      const [d,m,y] = s.split("/"); return `${y}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;
    }
    return s;
  };

  // ── Parsear fecha de Excel
  const leerArchivo = (file) => {
    if (!XLSX) { setErrMsg("La librería de Excel aún está cargando, espera un momento."); setEstado("error"); return; }
    setEstado("leyendo");
    setErrMsg("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: false });
        const hojas = {};
        wb.SheetNames.forEach(nombre => {
          const datos = XLSX.utils.sheet_to_json(wb.Sheets[nombre], { defval: "" });
          hojas[nombre] = datos;
        });
        const reconocidas = Object.keys(hojas).filter(h => ESQUEMAS[h]);
        if (reconocidas.length === 0) {
          setErrMsg(`No se encontraron hojas reconocidas. Nombres esperados: ${Object.keys(ESQUEMAS).join(", ")}`);
          setEstado("error");
          return;
        }
        const prev = {};
        reconocidas.forEach(h => { prev[h] = hojas[h]; });
        setPreview(prev);
        setEstado("preview");
      } catch (err) {
        setErrMsg("Error al leer el archivo: " + err.message);
        setEstado("error");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // ── Transformar filas por tipo ────────────────────────────────────────────
  const transformar = (hoja, filas) => {
    if (hoja === "Productos") return filas.filter(r=>r.Codigo).map(r=>({ id: String(r.Codigo).trim(), nombre: String(r.Nombre||"").trim() }));
    if (hoja === "Temperaturas") return filas.filter(r=>r.Codigo).map(r=>({ productoId: String(r.Codigo).trim(), temp: Number(r.Temperatura)||0, tiempo: Number(r.Tiempo)||0, indicaciones: String(r.Indicaciones||"").trim() }));
    if (hoja === "Precios") return filas.filter(r=>r.Codigo).map(r=>({ productoId: String(r.Codigo).trim(), precioNormal: Number(r.PrecioNormal)||0, precioMayoreo: Number(r.PrecioMayoreo)||0, cantMayoreo: Number(r.CantidadMayoreo)||0, precioSuperMayoreo: r.PrecioSuperMayoreo?Number(r.PrecioSuperMayoreo):null, cantSuperMayoreo: r.CantidadSuperMayoreo?Number(r.CantidadSuperMayoreo):null, observacion: String(r.Observacion||"").trim() }));
    if (hoja === "Margen") return filas.filter(r=>r.Codigo).map(r=>({ productoId: String(r.Codigo).trim(), costoProducto: Number(r.CostoProducto)||0, dtf: Number(r.DTF)||0, tinta: Number(r.Tinta)||0, papel: Number(r.Papel)||0, otros: Number(r.Otros)||0 }));
    if (hoja === "Ingresos") return filas.filter(r=>r.Fecha).map(r=>({ id: uid(), fecha: parseFecha(r.Fecha), codigoProducto: String(r.Codigo||"").trim(), nombreProducto: String(r.NombreProducto||"").trim(), cliente: String(r.Cliente||"").trim(), cantidad: Number(r.Cantidad)||1, precioUnitario: Number(r.PrecioUnitario)||0, precioFinal: Number(r.PrecioFinal)||0, estatus: r.Estatus||"Entregado", pago: r.Pago||"Pagado", observacion: String(r.Observacion||"").trim() }));
    if (hoja === "Gastos") return filas.filter(r=>r.Fecha).map(r=>({ id: uid(), fecha: parseFecha(r.Fecha), codigoProducto: String(r.Codigo||"").trim(), nombreProducto: String(r.NombreProducto||"").trim(), proveedor: String(r.Proveedor||"").trim(), cantidad: Number(r.Cantidad)||1, precioUnitario: Number(r.PrecioUnitario)||0, precioFinal: Number(r.PrecioFinal)||0, estatus: r.Estatus||"Entregado", pago: r.Pago||"Pagado", observacion: String(r.Observacion||"").trim() }));
    return [];
  };

  // ── Ejecutar importación ──────────────────────────────────────────────────
  const importar = async () => {
    const prog = {};
    for (const [hoja, filas] of Object.entries(preview)) {
      const datos = transformar(hoja, filas);
      prog[hoja] = datos.length;
      if (hoja === "Productos")    await setProductos(modoImport==="reemplazar" ? datos : (prev => [...prev, ...datos]));
      if (hoja === "Temperaturas") await setTemperaturas(modoImport==="reemplazar" ? datos : (prev => [...prev, ...datos]));
      if (hoja === "Precios")      await setPrecios(modoImport==="reemplazar" ? datos : (prev => [...prev, ...datos]));
      if (hoja === "Margen")       await setCostos(modoImport==="reemplazar" ? datos : (prev => [...prev, ...datos]));
      if (hoja === "Ingresos")     await setIngresos(modoImport==="reemplazar" ? datos : (prev => [...prev, ...datos]));
      if (hoja === "Gastos")       await setGastos(modoImport==="reemplazar" ? datos : (prev => [...prev, ...datos]));
    }
    setProgreso(prog);
    setEstado("importado");
  };

  // ── Descargar plantilla ───────────────────────────────────────────────────
  const descargarPlantilla = () => {
    if (!XLSX) { alert("La librería de Excel aún está cargando, espera un momento."); return; }
    const wb = XLSX.utils.book_new();

    const hojas = {
      Productos:    [["Codigo","Nombre"],["AA-001","Termo Skinny"],["AA-002","Playera cuello redondo"]],
      Temperaturas: [["Codigo","Temperatura","Tiempo","Indicaciones"],["AA-001",320,120,"Darle 60s por cada lado"]],
      Precios:      [["Codigo","Nombre","PrecioNormal","PrecioMayoreo","CantidadMayoreo","PrecioSuperMayoreo","CantidadSuperMayoreo","Observacion"],["AA-001","Termo Skinny",250,235,8,220,16,"Termo ML"]],
      Margen:       [["Codigo","CostoProducto","DTF","Tinta","Papel","Otros"],["AA-001",98,0,6,5,0]],
      Ingresos:     [["Fecha","Codigo","NombreProducto","Cliente","Cantidad","PrecioUnitario","PrecioFinal","Estatus","Pago","Observacion"],["2026-01-15","AA-001","Termo Skinny","Juan Pérez",2,250,500,"Entregado","Pagado",""]],
      Gastos:       [["Fecha","Codigo","NombreProducto","Proveedor","Cantidad","PrecioUnitario","PrecioFinal","Estatus","Pago","Observacion"],["2026-01-10","AA-001","Termo Skinny","Proveedor SA",10,98,980,"Entregado","Pagado",""]],
    };

    Object.entries(hojas).forEach(([nombre, datos]) => {
      const ws = XLSX.utils.aoa_to_sheet(datos);
      // Estilo encabezados (ancho de columnas)
      ws["!cols"] = datos[0].map(() => ({ wch: 20 }));
      XLSX.utils.book_append_sheet(wb, ws, nombre);
    });

    XLSX.writeFile(wb, "Plantilla_AdriAnis.xlsx");
  };

  const resetear = () => { setEstado("idle"); setPreview(null); setErrMsg(""); setProgreso({}); if(fileRef.current) fileRef.current.value=""; };

  // ── Colores por hoja ──────────────────────────────────────────────────────
  const hojaColor = { Productos:"blue", Temperaturas:"purple", Precios:"amber", Margen:"green", Ingresos:"emerald", Gastos:"rose" };

  return (
    <div className="space-y-5">
      <div>
        <h2 className="text-2xl font-black text-gray-800">Importar Datos</h2>
        <p className="text-gray-400 text-sm">Carga masiva desde Excel (.xlsx)</p>
      </div>

      {/* Paso 1: Descargar plantilla */}
      <Card className="p-5">
        <div className="flex items-start gap-4">
          <div className="w-10 h-10 rounded-xl bg-emerald-100 flex items-center justify-center flex-shrink-0">
            <span className="text-emerald-600 font-black text-lg">1</span>
          </div>
          <div className="flex-1">
            <p className="font-bold text-gray-800 mb-1">Descarga la plantilla</p>
            <p className="text-sm text-gray-500 mb-3">El Excel incluye 6 hojas: <strong>Productos, Temperaturas, Precios, Margen, Ingresos y Gastos</strong>. Llena solo las que necesites, deja vacías las demás.</p>
            <Btn onClick={descargarPlantilla} variant="success" size="sm">⬇ Descargar Plantilla Excel</Btn>
          </div>
        </div>
      </Card>

      {/* Paso 2: Modo de importación */}
      <Card className="p-5">
        <div className="flex items-start gap-4">
          <div className="w-10 h-10 rounded-xl bg-blue-100 flex items-center justify-center flex-shrink-0">
            <span className="text-blue-600 font-black text-lg">2</span>
          </div>
          <div className="flex-1">
            <p className="font-bold text-gray-800 mb-3">¿Cómo quieres importar?</p>
            <div className="flex gap-3">
              {[{val:"reemplazar",label:"Reemplazar todo",desc:"Borra los datos actuales y los sustituye"},{val:"agregar",label:"Agregar al final",desc:"Mantiene los datos actuales y agrega los nuevos"}].map(op=>(
                <div key={op.val} onClick={()=>setModoImport(op.val)}
                  className={`flex-1 p-3 rounded-xl border-2 cursor-pointer transition-all ${modoImport===op.val?"border-rose-400 bg-rose-50":"border-gray-200 hover:border-gray-300"}`}>
                  <p className={`text-sm font-bold ${modoImport===op.val?"text-rose-600":"text-gray-700"}`}>{op.label}</p>
                  <p className="text-xs text-gray-400 mt-0.5">{op.desc}</p>
                </div>
              ))}
            </div>
          </div>
        </div>
      </Card>

      {/* Paso 3: Subir archivo */}
      <Card className="p-5">
        <div className="flex items-start gap-4">
          <div className="w-10 h-10 rounded-xl bg-rose-100 flex items-center justify-center flex-shrink-0">
            <span className="text-rose-600 font-black text-lg">3</span>
          </div>
          <div className="flex-1">
            <p className="font-bold text-gray-800 mb-1">Sube tu archivo Excel</p>
            <p className="text-sm text-gray-500 mb-3">Puede ser la plantilla llena o tu propio Excel siempre que los nombres de hojas y columnas coincidan.</p>
            <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed border-gray-300 rounded-xl cursor-pointer hover:border-rose-400 hover:bg-rose-50 transition-all">
              <span className="text-3xl mb-1">📂</span>
              <span className="text-sm font-semibold text-gray-500">Haz clic para seleccionar</span>
              <span className="text-xs text-gray-400">Solo archivos .xlsx</span>
              <input ref={fileRef} type="file" accept=".xlsx" className="hidden"
                onChange={e=>{ if(e.target.files[0]) leerArchivo(e.target.files[0]); }}/>
            </label>
          </div>
        </div>
      </Card>

      {/* Estado: leyendo */}
      {estado==="leyendo" && (
        <Card className="p-6 text-center">
          <p className="text-2xl mb-2">⏳</p>
          <p className="font-semibold text-gray-600">Leyendo archivo...</p>
        </Card>
      )}

      {/* Estado: error */}
      {estado==="error" && (
        <Card className="p-5 border-l-4 border-red-400 bg-red-50">
          <p className="font-bold text-red-600 mb-1">Error al leer el archivo</p>
          <p className="text-sm text-red-500">{errMsg}</p>
          <Btn onClick={resetear} variant="secondary" size="sm" className="mt-3">Intentar de nuevo</Btn>
        </Card>
      )}

      {/* Estado: preview */}
      {estado==="preview" && preview && (
        <div className="space-y-4">
          <Card className="p-5 border-l-4 border-emerald-400 bg-emerald-50">
            <p className="font-bold text-emerald-700 mb-1">✅ Archivo leído correctamente</p>
            <p className="text-sm text-emerald-600">Se encontraron <strong>{Object.keys(preview).length}</strong> hojas reconocidas. Revisa el resumen y confirma la importación.</p>
          </Card>

          {/* Resumen por hoja */}
          <div className="space-y-3">
            {Object.entries(preview).map(([hoja, filas])=>(
              <Card key={hoja} className="p-4">
                <div className="flex items-center justify-between mb-3">
                  <div className="flex items-center gap-2">
                    <Badge color={hojaColor[hoja]||"gray"}>{hoja}</Badge>
                    <span className="text-sm font-bold text-gray-700">{filas.length} filas detectadas</span>
                  </div>
                  {modoImport==="reemplazar" && <Badge color="amber">Reemplazará datos actuales</Badge>}
                </div>
                {/* Preview primeras 3 filas */}
                {filas.length > 0 && (
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="bg-gray-50 rounded-lg">
                          {Object.keys(filas[0]).slice(0,6).map(col=>(
                            <th key={col} className="text-left px-2 py-1.5 text-gray-400 font-semibold truncate max-w-20">{col}</th>
                          ))}
                          {Object.keys(filas[0]).length>6 && <th className="text-gray-300 px-2">+{Object.keys(filas[0]).length-6} más</th>}
                        </tr>
                      </thead>
                      <tbody>
                        {filas.slice(0,3).map((fila,i)=>(
                          <tr key={i} className="border-t border-gray-50">
                            {Object.values(fila).slice(0,6).map((val,j)=>(
                              <td key={j} className="px-2 py-1.5 text-gray-600 truncate max-w-20">{String(val)}</td>
                            ))}
                            {Object.values(fila).length>6 && <td className="px-2 text-gray-300">...</td>}
                          </tr>
                        ))}
                        {filas.length>3 && (
                          <tr className="border-t border-gray-50">
                            <td colSpan={7} className="px-2 py-1.5 text-xs text-gray-400 text-center">... y {filas.length-3} filas más</td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                )}
              </Card>
            ))}
          </div>

          {/* Advertencia reemplazar */}
          {modoImport==="reemplazar" && (
            <Card className="p-4 border-l-4 border-amber-400 bg-amber-50">
              <p className="text-sm font-bold text-amber-700">⚠️ Atención</p>
              <p className="text-xs text-amber-600 mt-0.5">El modo "Reemplazar" eliminará todos los datos actuales de las hojas importadas. Esta acción no se puede deshacer.</p>
            </Card>
          )}

          <div className="flex gap-3">
            <Btn onClick={resetear} variant="secondary" className="flex-1">Cancelar</Btn>
            <Btn onClick={importar} variant="primary" className="flex-1">Importar {Object.values(preview).reduce((s,f)=>s+f.length,0)} registros</Btn>
          </div>
        </div>
      )}

      {/* Estado: importado */}
      {estado==="importado" && (
        <Card className="p-6">
          <div className="text-center mb-5">
            <p className="text-4xl mb-2">🎉</p>
            <p className="text-xl font-black text-gray-800">¡Importación completada!</p>
          </div>
          <div className="space-y-2 mb-5">
            {Object.entries(progreso).map(([hoja,cant])=>(
              <div key={hoja} className="flex items-center justify-between py-2 border-b border-gray-50 last:border-0">
                <div className="flex items-center gap-2">
                  <span className="text-emerald-500">✓</span>
                  <Badge color={hojaColor[hoja]||"gray"}>{hoja}</Badge>
                </div>
                <span className="text-sm font-bold text-gray-700">{cant} registros importados</span>
              </div>
            ))}
          </div>
          <Btn onClick={resetear} variant="secondary" className="w-full">Importar otro archivo</Btn>
        </Card>
      )}

      {/* Guía de columnas */}
      <Card className="p-5">
        <p className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-3">Referencia de columnas por hoja</p>
        <div className="space-y-2">
          {Object.entries(ESQUEMAS).map(([hoja,cols])=>(
            <div key={hoja} className="flex flex-wrap gap-1 items-center">
              <Badge color={hojaColor[hoja]||"gray"}>{hoja}</Badge>
              <span className="text-xs text-gray-400">→</span>
              {cols.map(c=><span key={c} className="text-xs bg-gray-100 text-gray-600 px-1.5 py-0.5 rounded font-mono">{c}</span>)}
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}

export default function App() {
  const [usuario, setUsuario] = useState(null);
  const [seccion, setSeccion] = useState("dashboard");
  const [syncMsg, setSyncMsg] = useState("");

  const [productos,    setProductos,    syncProductos]    = useStore("aa_productos",    PRODUCTOS_INIT,    "Productos");
  const [temperaturas, setTemperaturas, syncTemperaturas] = useStore("aa_temperaturas", TEMPERATURAS_INIT, "Temperaturas");
  const [precios,      setPrecios,      syncPrecios]      = useStore("aa_precios",      PRECIOS_INIT,      "Precios");
  const [costos,       setCostos,       syncCostos]       = useStore("aa_costos",       COSTOS_INIT,       "Margen");
  const [ingresos,     setIngresos,     syncIngresos]     = useStore("aa_ingresos",     [],                "Ingresos");
  const [gastos,       setGastos,       syncGastos]       = useStore("aa_gastos",       [],                "Gastos");

  const sincronizarTodo = async () => {
    setSyncMsg("Sincronizando...");
    await Promise.all([syncProductos(), syncTemperaturas(), syncPrecios(), syncCostos(), syncIngresos(), syncGastos()]);
    setSyncMsg("✅ Sincronizado");
    setTimeout(() => setSyncMsg(""), 3000);
  };

  if (!usuario) return <Login onLogin={setUsuario}/>;

  const NAV = [
    {key:"dashboard",    label:"Dashboard"},
    {key:"ingresos",     label:"Ingresos"},
    {key:"gastos",       label:"Gastos"},
    {key:"productos",    label:"Productos"},
    {key:"temperaturas", label:"Temperaturas"},
    {key:"precios",      label:"Precios"},
    {key:"margen",       label:"Margen"},
    {key:"importar",     label:"Importar"},
  ];

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      <header className="bg-white border-b border-gray-100 px-4 py-3 flex items-center justify-between sticky top-0 z-40 shadow-sm">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 bg-rose-500 rounded-lg flex items-center justify-center"><span className="text-white font-black text-sm">A</span></div>
          <span className="font-black text-gray-800 text-lg tracking-tight">AdriAnis</span>
        </div>
        <div className="flex items-center gap-3">
          {syncMsg && <span className="text-xs text-emerald-600 font-semibold">{syncMsg}</span>}
          <Btn onClick={sincronizarTodo} variant="secondary" size="sm">⟳ Sincronizar</Btn>
          <span className="text-xs text-gray-400 hidden sm:block">{usuario.nombre}</span>
          <Btn onClick={()=>setUsuario(null)} variant="ghost" size="sm">Salir</Btn>
        </div>
      </header>
      <div className="flex flex-1">
        <nav className="w-36 sm:w-48 bg-white border-r border-gray-100 flex flex-col py-4 gap-1 sticky top-14 h-[calc(100vh-56px)] overflow-y-auto flex-shrink-0">
          {NAV.map(n=>(
            <button key={n.key} onClick={()=>setSeccion(n.key)}
              className={`flex items-center px-4 py-2.5 mx-2 rounded-xl text-sm font-semibold transition-all text-left ${seccion===n.key?"bg-rose-500 text-white shadow-sm":"text-gray-500 hover:bg-gray-100 hover:text-gray-800"}`}>
              {n.label}
            </button>
          ))}
        </nav>
        <main className="flex-1 p-4 sm:p-6 max-w-2xl w-full overflow-y-auto">
          {seccion==="dashboard"    && <Dashboard ingresos={ingresos} gastos={gastos} costos={costos} precios={precios}/>}
          {seccion==="ingresos"     && <Movimientos tipo="ingreso" registros={ingresos} setRegistros={setIngresos} productos={productos} precios={precios}/>}
          {seccion==="gastos"       && <Movimientos tipo="gasto"   registros={gastos}   setRegistros={setGastos}   productos={productos} precios={precios}/>}
          {seccion==="productos"    && <Productos    productos={productos} setProductos={setProductos}/>}
          {seccion==="temperaturas" && <Temperaturas temperaturas={temperaturas} setTemperaturas={setTemperaturas} productos={productos}/>}
          {seccion==="precios"      && <Precios      precios={precios} setPrecios={setPrecios} productos={productos}/>}
          {seccion==="margen"       && <Margen       costos={costos}   setCostos={setCostos}   precios={precios} productos={productos}/>}
          {seccion==="importar"     && <Importar setProductos={setProductos} setTemperaturas={setTemperaturas} setPrecios={setPrecios} setCostos={setCostos} setIngresos={setIngresos} setGastos={setGastos}/>}
        </main>
      </div>
    </div>
  );
}
