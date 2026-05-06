const express = require('express');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'misionero2025';
const DATA_FILE = path.join(__dirname, 'data', 'ventas.json');

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

if (!fs.existsSync(path.join(__dirname, 'data'))) {
  fs.mkdirSync(path.join(__dirname, 'data'));
}
if (!fs.existsSync(DATA_FILE)) {
  fs.writeFileSync(DATA_FILE, JSON.stringify([]));
}

function leerVentas() {
  try { return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')); }
  catch { return []; }
}

function guardarVentas(ventas) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(ventas, null, 2));
}

// Vendedor envía sus datos
app.post('/api/ventas', (req, res) => {
  const { vendedor, compradores } = req.body;
  if (!vendedor || !compradores || !Array.isArray(compradores)) {
    return res.status(400).json({ error: 'Datos incompletos' });
  }
  const ventas = leerVentas();
  const idx = ventas.findIndex(v => v.vendedor.trim().toLowerCase() === vendedor.trim().toLowerCase());
  const entrada = { vendedor: vendedor.trim(), compradores, fechaEnvio: new Date().toISOString() };
  if (idx >= 0) ventas[idx] = entrada;
  else ventas.push(entrada);
  guardarVentas(ventas);
  res.json({ ok: true });
});

// Admin: ver datos
app.get('/api/admin/ventas', (req, res) => {
  if (req.query.password !== ADMIN_PASSWORD) return res.status(401).json({ error: 'No autorizado' });
  res.json(leerVentas());
});

// Admin: exportar Excel
app.get('/api/admin/exportar', (req, res) => {
  if (req.query.password !== ADMIN_PASSWORD) return res.status(401).json({ error: 'No autorizado' });

  const ventas = leerVentas();
  const PRODUCTOS = ['Combo 1','Combo 2','Combo 3','Muzzarella','Especial','Napolitana','Calabresa','Fugazzeta'];
  const wb = XLSX.utils.book_new();

  // Hoja detalle
  const rows = [['Vendedor','Comprador',...PRODUCTOS,'Fecha envío']];
  ventas.forEach(v => {
    v.compradores.forEach(c => {
      if (!PRODUCTOS.some(p => (c.cantidades[p]||0) > 0)) return;
      rows.push([v.vendedor, c.nombre||'Sin nombre', ...PRODUCTOS.map(p=>c.cantidades[p]||0), new Date(v.fechaEnvio).toLocaleString('es-AR')]);
    });
  });
  const totales = PRODUCTOS.map(p => ventas.reduce((s,v) => s + v.compradores.reduce((s2,c) => s2+(c.cantidades[p]||0),0),0));
  rows.push(['','TOTAL',...totales,'']);
  const ws1 = XLSX.utils.aoa_to_sheet(rows);
  ws1['!cols'] = [{wch:22},{wch:22},...PRODUCTOS.map(()=>({wch:14})),{wch:22}];
  XLSX.utils.book_append_sheet(wb, ws1, 'Detalle');

  // Hoja por vendedor
  const rows2 = [['Vendedor',...PRODUCTOS]];
  ventas.forEach(v => rows2.push([v.vendedor,...PRODUCTOS.map(p=>v.compradores.reduce((s,c)=>s+(c.cantidades[p]||0),0))]));
  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  ws2['!cols'] = [{wch:22},...PRODUCTOS.map(()=>({wch:14}))];
  XLSX.utils.book_append_sheet(wb, ws2, 'Por Vendedor');

  const buffer = XLSX.write(wb, { type:'buffer', bookType:'xlsx' });
  const fecha = new Date().toLocaleDateString('es-AR').replace(/\//g,'-');
  res.setHeader('Content-Disposition',`attachment; filename="ventas_${fecha}.xlsx"`);
  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buffer);
});

app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));
