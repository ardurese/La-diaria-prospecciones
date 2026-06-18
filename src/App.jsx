import { useState, useRef } from "react";
import * as XLSX from "xlsx";

const ASESORES = ["Matias", "Alejandro", "Lucas", "Juan Pablo", "Agustin"];
const ZONAS = ["CENTRO", "NORTE", "SUR", "FACUNDO", "WALTER", "SANTI"];
const COMPETENCIA = ["La Red", "Pronto", "Muebleria Express", "Global", "Otro"];
const INTERES = [
  { label: "Alto",  emoji: "🔥", color: "#16a34a", bg: "#f0fdf4", border: "#bbf7d0" },
  { label: "Medio", emoji: "⚡", color: "#d97706", bg: "#fffbeb", border: "#fde68a" },
  { label: "Bajo",  emoji: "❄️", color: "#dc2626", bg: "#fef2f2", border: "#fecaca" },
];

const T = {
  bg: "#eeedf2", surface: "#ffffff", surface2: "#e4e3ec",
  border: "rgba(0,0,0,0.08)", accent: "#7c3aed", accentLight: "#ede9fe",
  text: "#16131f", muted: "#6b7280", divider: "rgba(0,0,0,0.06)",
};
const AVATAR_COLORS = ["#7c3aed","#0284c7","#d97706","#16a34a","#dc2626"];
const FF = "'DM Sans', system-ui, sans-serif";
const init = { asesor:"", nombre:"", direccion:"", zona:"", interes:"", credito:[], comentarios:"", foto:null };

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyBpi-12oH6CzMT8eWEOHqBOytaoPecb2A3bjshXGB0OnOK7cvGaJ0yxxiH2qrnGk5Xwg/exec";

const SEED_RECORDS = 
[{"ts": "2026/06/09 10:25:08 a.m. GMT-3", "asesor": "Agustin", "nombre": "Lisandro", "direccion": "Santa Fe 7522", "zona": "CENTRO", "interes": "Medio", "credito": ["Otro"], "comentarios": "3415995706", "foto": null}, {"ts": "2026/06/09 10:30:55 a.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Julieta", "direccion": "José M. Rosa", "zona": "CENTRO", "interes": "Medio", "credito": ["Global"], "comentarios": "3413536787", "foto": null}, {"ts": "2026/06/09 10:40:22 a.m. GMT-3", "asesor": "Agustin", "nombre": "Jorge", "direccion": "Derqui 7548", "zona": "CENTRO", "interes": "Bajo", "credito": ["Muebleria Express"], "comentarios": "3415473539\nPásate más adelante y que le escapa a eso", "foto": null}, {"ts": "2026/06/09 10:57:10 a.m. GMT-3", "asesor": "Matias", "nombre": "Ignacio", "direccion": "17 de agosto", "zona": "NORTE", "interes": "Medio", "credito": [], "comentarios": "Baigorria", "foto": null}, {"ts": "2026/06/09 11:19:54 a.m. GMT-3", "asesor": "Matias", "nombre": "Alexis", "direccion": "Dante Aliguieri 3410", "zona": "NORTE", "interes": "Bajo", "credito": [], "comentarios": "Baigorria", "foto": null}, {"ts": "2026/06/09 12:40:06 p.m. GMT-3", "asesor": "Matias", "nombre": "Miguelina", "direccion": "El rosedal 3200", "zona": "NORTE", "interes": "Medio", "credito": ["Otro"], "comentarios": "Baigorria", "foto": null}, {"ts": "2026/06/09 1:51:21 p.m. GMT-3", "asesor": "Lucas", "nombre": "Beatriz", "direccion": "Alem y saveedra", "zona": "NORTE", "interes": "Medio", "credito": ["Pronto"], "comentarios": "Baigorria", "foto": null}, {"ts": "2026/06/09 1:52:07 p.m. GMT-3", "asesor": "Lucas", "nombre": "Romina", "direccion": "Los olivos 3434", "zona": "NORTE", "interes": "Medio", "credito": ["La Red"], "comentarios": "Baigorria", "foto": null}, {"ts": "2026/06/09 1:52:56 p.m. GMT-3", "asesor": "Lucas", "nombre": "Emilce", "direccion": "Alem y los naranjos", "zona": "NORTE", "interes": "Bajo", "credito": [], "comentarios": "Baigorria", "foto": null}, {"ts": "2026/06/09 2:48:18 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Mara", "direccion": "White 7515", "zona": "CENTRO", "interes": "Medio", "credito": ["Otro"], "comentarios": "3413757081", "foto": null}, {"ts": "2026/06/10 8:48:16 a.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Florencia", "direccion": "José María rosa", "zona": "CENTRO", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/10 8:48:55 a.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Verónica", "direccion": "José María rosa 1256", "zona": "CENTRO", "interes": "Alto", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/10 8:51:59 a.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Carlos", "direccion": "José María rosa 1280", "zona": "CENTRO", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/10 8:52:40 a.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Adrian", "direccion": "Chubut 7637", "zona": "CENTRO", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/10 8:54:17 a.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Sandra", "direccion": "Chubut 7596", "zona": "CENTRO", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/10 8:54:50 a.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Rosa", "direccion": "Chubut 7596", "zona": "CENTRO", "interes": "Medio", "credito": [], "comentarios": "", "foto": null}, {"ts": "2026/06/10 8:55:16 a.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Yanina", "direccion": "White 7640", "zona": "CENTRO", "interes": "Medio", "credito": [], "comentarios": "", "foto": null}, {"ts": "2026/06/10 10:52:26 a.m. GMT-3", "asesor": "Matias", "nombre": "Noelia", "direccion": "C.1882", "zona": "SUR", "interes": "Alto", "credito": ["La Red", "Otro"], "comentarios": "Barrio Toba", "foto": null}, {"ts": "2026/06/10 11:05:15 a.m. GMT-3", "asesor": "Matias", "nombre": "Micaela", "direccion": "Aborigenes argentinos", "zona": "SUR", "interes": "Bajo", "credito": ["La Red"], "comentarios": "Negocio lindo, barrio toba", "foto": null}, {"ts": "2026/06/10 11:53:49 a.m. GMT-3", "asesor": "Matias", "nombre": "Myriam", "direccion": "Cisneros 5400", "zona": "SUR", "interes": "Bajo", "credito": ["Muebleria Express", "Global"], "comentarios": "Barrio Cisneros", "foto": null}, {"ts": "2026/06/11 2:13:38 p.m. GMT-3", "asesor": "Lucas", "nombre": "Giuliano", "direccion": "Cereseto 5540", "zona": "SUR", "interes": "Bajo", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/11 2:14:23 p.m. GMT-3", "asesor": "Lucas", "nombre": "Ivana", "direccion": "Sánchez Bustamante 2450", "zona": "SUR", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/11 2:16:44 p.m. GMT-3", "asesor": "Lucas", "nombre": "Marisa", "direccion": "Santiago 5549", "zona": "SUR", "interes": "Alto", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/11 2:17:35 p.m. GMT-3", "asesor": "Lucas", "nombre": "Ruth", "direccion": "Arijón 2400", "zona": "SUR", "interes": "Medio", "credito": [], "comentarios": "Trabajaba con pago por día hace mucho tiempo", "foto": null}, {"ts": "2026/06/11 2:18:25 p.m. GMT-3", "asesor": "Lucas", "nombre": "Elías", "direccion": "Av rosario 2514", "zona": "SUR", "interes": "Medio", "credito": [], "comentarios": "Me comentó que le habían rechazado un crédito de otra empresa", "foto": null}, {"ts": "2026/06/12 10:55:35 a.m. GMT-3", "asesor": "Matias", "nombre": "Rosario", "direccion": "3 de febrero 2300", "zona": "CENTRO", "interes": "Bajo", "credito": ["La Red"], "comentarios": "Vineria", "foto": null}, {"ts": "2026/06/12 12:57:27 p.m. GMT-3", "asesor": "Matias", "nombre": "Martin", "direccion": "Balcarce 1437", "zona": "CENTRO", "interes": "Bajo", "credito": [], "comentarios": "Barberia", "foto": null}, {"ts": "2026/06/16 11:22:49 a.m. GMT-3", "asesor": "Matias", "nombre": "Mariela", "direccion": "Lamadrid y oroño", "zona": "WALTER", "interes": "Medio", "credito": ["La Red"], "comentarios": "Panaderia familiar", "foto": null}, {"ts": "2026/06/16 11:45:24 a.m. GMT-3", "asesor": "Matias", "nombre": "Jonatan", "direccion": "Lamadrid y moreno", "zona": "WALTER", "interes": "Medio", "credito": ["Pronto", "Otro"], "comentarios": "Supermercado", "foto": null}, {"ts": "2026/06/16 11:55:09 a.m. GMT-3", "asesor": "Alejandro", "nombre": "Alicia", "direccion": "La Madrid 2091", "zona": "WALTER", "interes": "Medio", "credito": ["La Red", "Pronto"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 12:16:22 p.m. GMT-3", "asesor": "Alejandro", "nombre": "Mario almada", "direccion": "Ilarion de la quinta 1793", "zona": "WALTER", "interes": "Medio", "credito": [], "comentarios": "Quería un cartel luminoso", "foto": null}, {"ts": "2026/06/16 12:30:24 p.m. GMT-3", "asesor": "Matias", "nombre": "Dario", "direccion": "Patria 1800", "zona": "WALTER", "interes": "Medio", "credito": ["Otro"], "comentarios": "Almacen en crecimiento", "foto": null}, {"ts": "2026/06/16 1:26:33 p.m. GMT-3", "asesor": "Agustin", "nombre": "Liliana", "direccion": "Ituzaingo 5115", "zona": "SANTI", "interes": "Medio", "credito": ["Pronto"], "comentarios": "3416575217", "foto": null}, {"ts": "2026/06/16 1:32:15 p.m. GMT-3", "asesor": "Agustin", "nombre": "Alicia", "direccion": "Felipe more 2010", "zona": "SANTI", "interes": "Bajo", "credito": ["La Red"], "comentarios": "3415472134", "foto": null}, {"ts": "2026/06/16 1:33:02 p.m. GMT-3", "asesor": "Agustin", "nombre": "Micaela", "direccion": "Felipe more 2055", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "3412804788", "foto": null}, {"ts": "2026/06/16 1:34:27 p.m. GMT-3", "asesor": "Agustin", "nombre": "Juan", "direccion": "Matienzo 1926", "zona": "SANTI", "interes": "Bajo", "credito": [], "comentarios": "Matienzo 1926", "foto": null}, {"ts": "2026/06/16 1:34:58 p.m. GMT-3", "asesor": "Alejandro", "nombre": "Luciana", "direccion": "Ayala gauna 7886", "zona": "FACUNDO", "interes": "Medio", "credito": ["La Red", "Pronto"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 1:37:59 p.m. GMT-3", "asesor": "Agustin", "nombre": "Basilia", "direccion": "Riobamba 5054", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "3416603412", "foto": null}, {"ts": "2026/06/16 1:39:01 p.m. GMT-3", "asesor": "Agustin", "nombre": "Beatriz", "direccion": "Gutemberg 2292", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "3413647684", "foto": null}, {"ts": "2026/06/16 1:40:04 p.m. GMT-3", "asesor": "Agustin", "nombre": "Ivana", "direccion": "Viamonte 4682", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "3417459925", "foto": null}, {"ts": "2026/06/16 1:41:02 p.m. GMT-3", "asesor": "Agustin", "nombre": "Ernesto", "direccion": "Pascual rosas 2308", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "3412836527", "foto": null}, {"ts": "2026/06/16 1:41:47 p.m. GMT-3", "asesor": "Agustin", "nombre": "Nadia", "direccion": "Riobamba 4628", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "3413777159", "foto": null}, {"ts": "2026/06/16 1:53:00 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Rocio", "direccion": "Felipe more", "zona": "SANTI", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 1:53:32 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Marcela", "direccion": "Felipe more", "zona": "SANTI", "interes": "Alto", "credito": ["Global"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 1:54:40 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Federico", "direccion": "Matienzo 1725", "zona": "SANTI", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 1:55:37 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Romina", "direccion": "Matienzo 2085", "zona": "SANTI", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 1:56:11 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Debora", "direccion": "Riobamba 4700", "zona": "SANTI", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 1:56:39 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Melina", "direccion": "Gutemberg 2226", "zona": "SANTI", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 1:57:48 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Rossy", "direccion": "La paz y Gutemberg", "zona": "SANTI", "interes": "Alto", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 1:58:20 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Paola", "direccion": "Pascual rosas 2235", "zona": "SANTI", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/16 1:58:51 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Silvia", "direccion": "Riobamba 4613", "zona": "SANTI", "interes": "Alto", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 10:45:34 a.m. GMT-3", "asesor": "Alejandro", "nombre": "Mauro", "direccion": "Batlle y Ordóñez 1565", "zona": "SANTI", "interes": "Medio", "credito": ["La Red"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 10:56:55 a.m. GMT-3", "asesor": "Alejandro", "nombre": "Milagros", "direccion": "Batlle y ordoñes 1563", "zona": "SANTI", "interes": "Medio", "credito": ["La Red"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 11:27:32 a.m. GMT-3", "asesor": "Matias", "nombre": "Graciela", "direccion": "Ipe 6400", "zona": "WALTER", "interes": "Medio", "credito": [], "comentarios": "Polleria/Heladeria. No trabaja con pago por dia", "foto": null}, {"ts": "2026/06/17 11:28:05 a.m. GMT-3", "asesor": "Matias", "nombre": "Giuliana", "direccion": "España 6400", "zona": "WALTER", "interes": "Medio", "credito": [], "comentarios": "Forraje. No trabaja con pago por dia.", "foto": null}, {"ts": "2026/06/17 12:50:16 p.m. GMT-3", "asesor": "Matias", "nombre": "Arce", "direccion": "Melian 6300", "zona": "WALTER", "interes": "Bajo", "credito": ["Global"], "comentarios": "Panaderia, trabaja hace 14 años con Global", "foto": null}, {"ts": "2026/06/17 2:14:48 p.m. GMT-3", "asesor": "Lucas", "nombre": "Gabriela", "direccion": "Orquídea 6600", "zona": "WALTER", "interes": "Bajo", "credito": ["La Red"], "comentarios": "Trabajaba hace mucho con pago por dia", "foto": null}, {"ts": "2026/06/17 2:15:28 p.m. GMT-3", "asesor": "Lucas", "nombre": "Sandra", "direccion": "Ceibo 1300", "zona": "WALTER", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:15:46 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Mauro", "direccion": "Roullon 2100", "zona": "SANTI", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:16:06 p.m. GMT-3", "asesor": "Matias", "nombre": "Daiana", "direccion": "Caña de Ambar", "zona": "WALTER", "interes": "Alto", "credito": ["Global"], "comentarios": "Polleria/Almacen consulto producto", "foto": null}, {"ts": "2026/06/17 2:16:53 p.m. GMT-3", "asesor": "Lucas", "nombre": "Natalia", "direccion": "Melian y khantuta", "zona": "WALTER", "interes": "Bajo", "credito": ["Pronto"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:17:13 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Marisa", "direccion": "Cerrito 5699", "zona": "SANTI", "interes": "Medio", "credito": ["Global"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:18:01 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Vanina", "direccion": "teniente agneta 2150", "zona": "SANTI", "interes": "Medio", "credito": ["La Red"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:18:43 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Carlos", "direccion": "Riobamba 5618", "zona": "SANTI", "interes": "Alto", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:19:47 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Melina", "direccion": "Riobamba 5658", "zona": "SANTI", "interes": "Bajo", "credito": ["La Red", "Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:20:26 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Nahuel", "direccion": "Riobamba 5684", "zona": "SANTI", "interes": "Alto", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:21:01 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Gustavo", "direccion": "PL Funes 2284", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:21:38 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Romina", "direccion": "teniente agneta 2270", "zona": "SANTI", "interes": "Medio", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:22:08 p.m. GMT-3", "asesor": "Juan Pablo", "nombre": "Marsu", "direccion": "teniente agneta 2410", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "", "foto": null}, {"ts": "2026/06/17 2:27:09 p.m. GMT-3", "asesor": "Agustin", "nombre": "Eva", "direccion": "Avenido roullon 5200", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "3415007037", "foto": null}, {"ts": "2026/06/17 2:27:47 p.m. GMT-3", "asesor": "Agustin", "nombre": "Lucas", "direccion": "Cerrito 5751", "zona": "SANTI", "interes": "Medio", "credito": ["Otro"], "comentarios": "3412596111", "foto": null}, {"ts": "2026/06/17 2:28:57 p.m. GMT-3", "asesor": "Agustin", "nombre": "Romina", "direccion": "Cerrito 5632", "zona": "SANTI", "interes": "Bajo", "credito": ["Global"], "comentarios": "3413319520", "foto": null}, {"ts": "2026/06/17 2:29:49 p.m. GMT-3", "asesor": "Agustin", "nombre": "Karen", "direccion": "Pedro lino funes 2230", "zona": "SANTI", "interes": "Bajo", "credito": ["Global"], "comentarios": "3417489712", "foto": null}, {"ts": "2026/06/17 2:30:45 p.m. GMT-3", "asesor": "Agustin", "nombre": "Mauricio", "direccion": "Pedro lino funes 2267", "zona": "SANTI", "interes": "Bajo", "credito": ["Pronto"], "comentarios": "3412749308", "foto": null}, {"ts": "2026/06/17 2:31:31 p.m. GMT-3", "asesor": "Agustin", "nombre": "Laura", "direccion": "La Paz 5584", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "3416017634", "foto": null}, {"ts": "2026/06/17 2:32:19 p.m. GMT-3", "asesor": "Agustin", "nombre": "Cecilia", "direccion": "Ocampo 5754", "zona": "SANTI", "interes": "Bajo", "credito": ["Otro"], "comentarios": "3413476964", "foto": null}];

function Chip({ label, selected, onClick, color }) {
  const ac = color || T.accent;
  return (
    <button onClick={onClick} style={{
      padding:"8px 15px", borderRadius:99, fontSize:13, fontWeight:600, cursor:"pointer",
      border:`1.5px solid ${selected ? ac : T.border}`,
      background: selected ? ac : T.surface, color: selected ? "#fff" : T.muted,
      transition:"all 0.12s", outline:"none", boxShadow: selected ? `0 2px 8px ${ac}44` : "none", fontFamily:FF,
    }}>{label}</button>
  );
}

function FieldLabel({ children, required, error }) {
  return (
    <label style={{ display:"block", fontSize:10, fontWeight:700, letterSpacing:1.3,
      textTransform:"uppercase", color: error ? "#dc2626" : T.muted, marginBottom:8, fontFamily:FF }}>
      {children}{required && <span style={{ color:T.accent, marginLeft:2 }}>*</span>}
    </label>
  );
}

function TextInput({ value, onChange, placeholder, error }) {
  const [focus, setFocus] = useState(false);
  return (
    <input value={value} onChange={onChange} placeholder={placeholder}
      onFocus={() => setFocus(true)} onBlur={() => setFocus(false)}
      style={{
        width:"100%", padding:"12px 14px", borderRadius:10, fontSize:15, outline:"none",
        boxSizing:"border-box", background:T.surface2,
        border:`1.5px solid ${error ? "#dc2626" : focus ? T.accent : T.border}`,
        color:T.text, transition:"border-color 0.15s", fontFamily:FF,
      }} />
  );
}

function Divider() {
  return <div style={{ height:1, background:T.divider, margin:"4px 0 20px" }} />;
}

function PhotoUpload({ value, onChange }) {
  const ref = useRef();
  const handleFile = (e) => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => onChange(ev.target.result);
    reader.readAsDataURL(file);
  };
  return (
    <div>
      <input ref={ref} type="file" accept="image/*" capture="environment" onChange={handleFile} style={{ display:"none" }} />
      {!value ? (
        <button onClick={() => ref.current.click()} style={{
          width:"100%", padding:"18px 0", borderRadius:12, cursor:"pointer", outline:"none",
          border:`2px dashed ${T.border}`, background:T.surface2,
          display:"flex", flexDirection:"column", alignItems:"center", gap:6,
        }}>
          <span style={{ fontSize:26 }}>📷</span>
          <span style={{ fontSize:13, fontWeight:700, color:T.accent, fontFamily:FF }}>Sacar o elegir foto</span>
          <span style={{ fontSize:11, color:T.muted, fontFamily:FF }}>Fachada, cartel, local…</span>
        </button>
      ) : (
        <div style={{ position:"relative", borderRadius:12, overflow:"hidden" }}>
          <img src={value} alt="Foto" style={{ width:"100%", maxHeight:200, objectFit:"cover", display:"block" }} />
          <button onClick={() => onChange(null)} style={{
            position:"absolute", top:8, right:8, background:"rgba(0,0,0,0.55)",
            color:"#fff", border:"none", borderRadius:99, width:28, height:28,
            fontSize:14, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center",
          }}>✕</button>
        </div>
      )}
    </div>
  );
}

function InterestBadge({ label, count }) {
  if (!count) return null;
  const opt = INTERES.find(i => i.label === label);
  return (
    <span style={{
      background: opt.bg, color: opt.color, border: `1px solid ${opt.border}`,
      borderRadius: 99, padding: "2px 8px", fontSize: 11, fontWeight: 700,
      display: "inline-flex", alignItems: "center", gap: 3,
    }}>{opt.emoji} {count}</span>
  );
}

function downloadExcel(records) {
  const rows = records.map(r => ({
    Fecha: r.ts,
    Asesor: r.asesor,
    Nombre: r.nombre,
    Dirección: r.direccion,
    Zona: r.zona,
    Interés: r.interes,
    Competencia: Array.isArray(r.credito) ? r.credito.join(", ") : r.credito,
    Comentarios: r.comentarios,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [
    { wch: 22 }, { wch: 14 }, { wch: 16 }, { wch: 28 },
    { wch: 10 }, { wch: 9 }, { wch: 20 }, { wch: 32 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Relevamientos");
  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Relevamientos_La_Diaria_${fecha}.xlsx`);
}

function Dashboard({ records, onBack }) {
  const total = records.length;
  const globalAlto  = records.filter(r => r.interes === "Alto").length;
  const globalMedio = records.filter(r => r.interes === "Medio").length;
  const globalBajo  = records.filter(r => r.interes === "Bajo").length;

  const counts = ASESORES.map((a, i) => ({
    name: a,
    color: AVATAR_COLORS[i],
    total: records.filter(r => r.asesor === a).length,
    alto:  records.filter(r => r.asesor === a && r.interes === "Alto").length,
    medio: records.filter(r => r.asesor === a && r.interes === "Medio").length,
    bajo:  records.filter(r => r.asesor === a && r.interes === "Bajo").length,
  })).sort((a, b) => b.total - a.total);

  const max = Math.max(...counts.map(c => c.total), 1);

  return (
    <div style={{ minHeight:"100vh", background:T.bg, padding:"24px 16px 40px", fontFamily:FF }}>
      <div style={{ maxWidth:440, margin:"0 auto" }}>

        <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:24 }}>
          <button onClick={onBack} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:10, padding:"8px 12px", cursor:"pointer", fontSize:18, lineHeight:1 }}>←</button>
          <div style={{ flex:1 }}>
            <p style={{ margin:0, fontSize:10, fontWeight:700, letterSpacing:1.3, textTransform:"uppercase", color:T.muted }}>La Diaria Equipamientos</p>
            <h1 style={{ margin:0, fontSize:22, fontWeight:900, color:T.text, letterSpacing:-0.5 }}>Dashboard</h1>
          </div>
          <button onClick={() => downloadExcel(records)} style={{
            background:T.surface, border:`1px solid ${T.border}`, borderRadius:10, padding:"9px 14px",
            cursor:"pointer", fontSize:13, fontWeight:700, color:T.accent, display:"flex", alignItems:"center", gap:6,
            fontFamily:FF, whiteSpace:"nowrap",
          }}>⬇ Excel</button>
        </div>

        <div style={{ background:T.accent, borderRadius:18, padding:"22px 22px", marginBottom:12, position:"relative", overflow:"hidden" }}>
          <div style={{ position:"absolute", right:-20, top:-20, width:100, height:100, borderRadius:"50%", background:"rgba(255,255,255,0.07)" }} />
          <div style={{ position:"absolute", right:20, bottom:-30, width:80, height:80, borderRadius:"50%", background:"rgba(255,255,255,0.05)" }} />
          <p style={{ margin:"0 0 4px", fontSize:10, fontWeight:700, letterSpacing:1.3, textTransform:"uppercase", color:"rgba(255,255,255,0.6)" }}>Total cargado</p>
          <p style={{ margin:0, fontSize:52, fontWeight:900, color:"#fff", lineHeight:1, letterSpacing:-2 }}>{total}</p>
          <p style={{ margin:"6px 0 0", fontSize:13, color:"rgba(255,255,255,0.6)" }}>relevamiento{total !== 1 ? "s" : ""} en total</p>
        </div>

        <div style={{ background:T.surface, borderRadius:16, padding:"16px 18px", marginBottom:12, display:"flex", gap:8, alignItems:"center", boxShadow:"0 2px 12px rgba(0,0,0,0.05)" }}>
          <span style={{ fontSize:12, fontWeight:700, color:T.muted, flex:1 }}>General</span>
          <InterestBadge label="Alto"  count={globalAlto} />
          <InterestBadge label="Medio" count={globalMedio} />
          <InterestBadge label="Bajo"  count={globalBajo} />
          {total === 0 && <span style={{ fontSize:12, color:T.muted }}>Sin datos aún</span>}
        </div>

        <div style={{ background:T.surface, borderRadius:18, padding:"20px 18px", boxShadow:"0 2px 16px rgba(0,0,0,0.06)" }}>
          <p style={{ margin:"0 0 18px", fontSize:10, fontWeight:700, letterSpacing:1.3, textTransform:"uppercase", color:T.muted }}>Por asesor</p>

          {counts.map(({ name, total: cnt, color, alto, medio, bajo }) => (
            <div key={name} style={{ marginBottom:20 }}>
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:7 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ width:34, height:34, borderRadius:"50%", background:color, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:800, fontSize:13, flexShrink:0 }}>
                    {name[0]}
                  </div>
                  <div>
                    <p style={{ margin:0, fontSize:14, fontWeight:700, color:T.text }}>{name}</p>
                    <div style={{ display:"flex", gap:5, marginTop:3 }}>
                      <InterestBadge label="Alto"  count={alto} />
                      <InterestBadge label="Medio" count={medio} />
                      <InterestBadge label="Bajo"  count={bajo} />
                      {cnt === 0 && <span style={{ fontSize:11, color:T.muted }}>Sin relevamientos</span>}
                    </div>
                  </div>
                </div>
                <span style={{ fontSize:24, fontWeight:900, color: cnt > 0 ? color : T.muted, letterSpacing:-1, minWidth:28, textAlign:"right" }}>{cnt}</span>
              </div>
              <div style={{ height:6, background:T.surface2, borderRadius:99, overflow:"hidden" }}>
                <div style={{
                  height:"100%", borderRadius:99,
                  width:`${(cnt / max) * 100}%`,
                  background: cnt > 0 ? color : T.border,
                  transition:"width 0.5s ease",
                }} />
              </div>
            </div>
          ))}

          {total === 0 && (
            <p style={{ textAlign:"center", color:T.muted, fontSize:13, padding:"16px 0" }}>Todavía no hay relevamientos</p>
          )}
        </div>

        <p style={{ textAlign:"center", color:T.muted, fontSize:11, marginTop:16 }}>
          {new Date().toLocaleDateString("es-AR", { dateStyle:"long" })}
        </p>
      </div>
    </div>
  );
}

export default function App() {
  const [form, setForm] = useState(init);
  const [errors, setErrors] = useState({});
  const [screen, setScreen] = useState("form");
  const [records, setRecords] = useState(SEED_RECORDS);
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState(false);

  const set = (k, v) => { setForm(f => ({ ...f, [k]: v })); setErrors(e => ({ ...e, [k]: "" })); };
  const toggleCredito = v => set("credito", form.credito.includes(v) ? form.credito.filter(x => x !== v) : [...form.credito, v]);

  const submit = async () => {
    const e = {};
    if (!form.asesor) e.asesor = true;
    if (!form.nombre.trim()) e.nombre = true;
    if (!form.direccion.trim()) e.direccion = true;
    if (!form.zona) e.zona = true;
    if (!form.interes) e.interes = true;
    if (Object.keys(e).length) { setErrors(e); return; }

    const newRecord = { ...form, ts: new Date().toISOString() };
    setSaving(true);
    setSaveError(false);
    try {
      await fetch(SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify(newRecord),
      });
      setRecords(r => [...r, newRecord]);
      setScreen("confirm");
    } catch (err) {
      setSaveError(true);
    } finally {
      setSaving(false);
    }
  };

  const intOpt = INTERES.find(i => i.label === form.interes);

  if (screen === "dashboard") return <Dashboard records={records} onBack={() => setScreen("form")} />;

  if (screen === "confirm") return (
    <div style={{ minHeight:"100vh", background:T.bg, display:"flex", alignItems:"center", justifyContent:"center", padding:20, fontFamily:FF }}>
      <div style={{ background:T.surface, borderRadius:20, padding:28, maxWidth:400, width:"100%", boxShadow:"0 4px 24px rgba(0,0,0,0.08)" }}>
        <div style={{ textAlign:"center", marginBottom:20 }}>
          <div style={{ fontSize:44, marginBottom:8 }}>✅</div>
          <h2 style={{ margin:0, color:T.text, fontSize:20, fontWeight:800 }}>¡Listo!</h2>
          <p style={{ color:T.muted, fontSize:13, marginTop:4 }}>{new Date().toLocaleString("es-AR", { dateStyle:"medium", timeStyle:"short" })}</p>
        </div>
        {form.foto && <img src={form.foto} alt="Foto" style={{ width:"100%", maxHeight:150, objectFit:"cover", borderRadius:12, marginBottom:16 }} />}
        <div style={{ background:T.bg, borderRadius:12, padding:"14px 16px", fontSize:13, color:T.muted, lineHeight:2.1, marginBottom:20 }}>
          {[["Asesor", form.asesor],["Cliente", form.nombre],["Dirección", form.direccion],["Zona", form.zona]].map(([k,v]) => (
            <div key={k}><span>{k}: </span><span style={{ color:T.text, fontWeight:700 }}>{v}</span></div>
          ))}
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span>Interés: </span>
            {intOpt && <span style={{ background:intOpt.bg, color:intOpt.color, border:`1px solid ${intOpt.border}`, borderRadius:99, padding:"1px 10px", fontWeight:700, fontSize:12 }}>{intOpt.emoji} {intOpt.label}</span>}
          </div>
          {form.credito.length > 0 && <div><span>Competencia: </span><span style={{ color:T.text, fontWeight:700 }}>{form.credito.join(", ")}</span></div>}
          {form.comentarios && <div><span>Nota: </span><span style={{ color:T.text }}>{form.comentarios}</span></div>}
        </div>
        <div style={{ display:"flex", gap:10 }}>
          <button onClick={() => setScreen("dashboard")} style={{ flex:1, padding:"13px 0", background:T.accentLight, color:T.accent, border:"none", borderRadius:10, fontWeight:700, fontSize:14, cursor:"pointer", fontFamily:FF }}>Ver dashboard</button>
          <button onClick={() => { setForm(init); setScreen("form"); }} style={{ flex:1, padding:"13px 0", background:T.accent, color:"#fff", border:"none", borderRadius:10, fontWeight:700, fontSize:14, cursor:"pointer", fontFamily:FF }}>+ Nuevo</button>
        </div>
      </div>
    </div>
  );

  return (
    <div style={{ minHeight:"100vh", background:T.bg, display:"flex", justifyContent:"center", padding:"24px 16px 40px", fontFamily:FF }}>
      <div style={{ maxWidth:440, width:"100%" }}>

        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:20 }}>
          <div>
            <span style={{ fontSize:10, fontWeight:700, letterSpacing:1.4, textTransform:"uppercase", color:T.accent, background:T.accentLight, padding:"3px 10px", borderRadius:6 }}>La Diaria Equipamientos</span>
            <h1 style={{ margin:"8px 0 2px", fontSize:26, fontWeight:900, color:T.text, letterSpacing:-0.5 }}>Relevamiento</h1>
            <p style={{ margin:0, fontSize:13, color:T.muted }}>Completá los datos del prospecto</p>
          </div>
          <div style={{ display:"flex", gap:8 }}>
            <button onClick={() => downloadExcel(records)} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:12, padding:"8px 12px", cursor:"pointer", fontSize:18 }}>⬇</button>
            <button onClick={() => setScreen("dashboard")} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:12, padding:"8px 12px", cursor:"pointer", display:"flex", flexDirection:"column", alignItems:"center", gap:1 }}>
              <span style={{ fontSize:18 }}>📊</span>
              {records.length > 0 && <span style={{ fontSize:10, fontWeight:800, color:T.accent }}>{records.length}</span>}
            </button>
          </div>
        </div>

        <div style={{ background:T.surface, borderRadius:20, padding:"22px 18px", boxShadow:"0 2px 16px rgba(0,0,0,0.06)" }}>

          <div style={{ marginBottom:20 }}>
            <FieldLabel required error={errors.asesor}>Asesor</FieldLabel>
            <div style={{ display:"flex", flexWrap:"wrap", gap:7 }}>
              {ASESORES.map(a => <Chip key={a} label={a} selected={form.asesor === a} onClick={() => set("asesor", a)} />)}
            </div>
            {errors.asesor && <p style={{ color:"#dc2626", fontSize:11, margin:"5px 0 0" }}>Seleccioná un asesor</p>}
          </div>

          <Divider />

          <div style={{ marginBottom:16 }}>
            <FieldLabel required error={errors.nombre}>Nombre del cliente</FieldLabel>
            <TextInput value={form.nombre} onChange={e => set("nombre", e.target.value)} placeholder="Ej: María López" error={errors.nombre} />
            {errors.nombre && <p style={{ color:"#dc2626", fontSize:11, margin:"5px 0 0" }}>Campo requerido</p>}
          </div>

          <div style={{ marginBottom:20 }}>
            <FieldLabel required error={errors.direccion}>Dirección</FieldLabel>
            <TextInput value={form.direccion} onChange={e => set("direccion", e.target.value)} placeholder="Ej: Rivadavia 542" error={errors.direccion} />
            {errors.direccion && <p style={{ color:"#dc2626", fontSize:11, margin:"5px 0 0" }}>Campo requerido</p>}
          </div>

          <Divider />

          <div style={{ marginBottom:20 }}>
            <FieldLabel required error={errors.zona}>Zona</FieldLabel>
            <div style={{ display:"flex", flexWrap:"wrap", gap:7 }}>
              {ZONAS.map(z => <Chip key={z} label={z} selected={form.zona === z} onClick={() => set("zona", z)} color="#7c3aed" />)}
            </div>
            {errors.zona && <p style={{ color:"#dc2626", fontSize:11, margin:"5px 0 0" }}>Seleccioná una zona</p>}
          </div>

          <Divider />

          <div style={{ marginBottom:20 }}>
            <FieldLabel required error={errors.interes}>Nivel de interés</FieldLabel>
            <div style={{ display:"flex", gap:8 }}>
              {INTERES.map(({ label, emoji, color, bg, border }) => {
                const sel = form.interes === label;
                return (
                  <button key={label} onClick={() => set("interes", label)} style={{
                    flex:1, padding:"13px 8px", borderRadius:12, cursor:"pointer", outline:"none", textAlign:"center",
                    border:`1.5px solid ${sel ? color : T.border}`,
                    background: sel ? bg : T.surface2, transition:"all 0.12s",
                    boxShadow: sel ? `0 2px 8px ${color}33` : "none",
                  }}>
                    <div style={{ fontSize:22 }}>{emoji}</div>
                    <div style={{ fontSize:13, fontWeight:700, color: sel ? color : T.muted, marginTop:4 }}>{label}</div>
                  </button>
                );
              })}
            </div>
            {errors.interes && <p style={{ color:"#dc2626", fontSize:11, margin:"5px 0 0" }}>Indicá el interés</p>}
          </div>

          <Divider />

          <div style={{ marginBottom:20 }}>
            <FieldLabel>Crédito en competencia</FieldLabel>
            <div style={{ display:"flex", flexWrap:"wrap", gap:7 }}>
              {COMPETENCIA.map(c => <Chip key={c} label={c} selected={form.credito.includes(c)} onClick={() => toggleCredito(c)} color="#0284c7" />)}
            </div>
          </div>

          <Divider />

          <div style={{ marginBottom:20 }}>
            <FieldLabel>Foto del local</FieldLabel>
            <PhotoUpload value={form.foto} onChange={v => set("foto", v)} />
          </div>

          <Divider />

          <div style={{ marginBottom:22 }}>
            <FieldLabel>Comentarios</FieldLabel>
            <textarea value={form.comentarios} onChange={e => set("comentarios", e.target.value)}
              placeholder="Observaciones del encuentro..." rows={3}
              style={{ width:"100%", padding:"12px 14px", borderRadius:10, fontSize:14, outline:"none", resize:"vertical", boxSizing:"border-box", background:T.surface2, border:`1.5px solid ${T.border}`, color:T.text, fontFamily:FF }} />
          </div>

          <button onClick={submit} disabled={saving} style={{ width:"100%", padding:"15px 0", background: saving ? T.muted : T.accent, color:"#fff", border:"none", borderRadius:12, fontWeight:800, fontSize:16, cursor: saving ? "default" : "pointer", boxShadow: saving ? "none" : `0 4px 14px ${T.accent}44`, fontFamily:FF }}>
            {saving ? "Guardando..." : "Guardar relevamiento"}
          </button>
          {saveError && (
            <p style={{ color:"#dc2626", fontSize:12, textAlign:"center", marginTop:10 }}>
              No se pudo guardar online. Revisá tu conexión e intentá de nuevo.
            </p>
          )}
        </div>

        <p style={{ textAlign:"center", color:T.muted, fontSize:11, marginTop:16 }}>La Diaria Equipamientos · {new Date().getFullYear()}</p>
      </div>
    </div>
  );
}

