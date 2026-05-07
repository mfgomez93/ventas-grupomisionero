const express = require('express');
const path = require('path');
const XLSX = require('xlsx');
const { MongoClient } = require('mongodb');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'misionero2025';
const MONGODB_URI = process.env.MONGODB_URI;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

let db;

async function conectarDB() {
  const client = new MongoClient(MONGODB_URI);
  await client.connect();
  db = client.db('ventas_misionero');
  console.log('MongoDB conectado');
}

function coleccion() {
  return db.collection('ventas');
}

// Vendedor envía sus datos
app.post('/api/ventas', async (req, res) => {
  const { vendedor, compradores } = req.body;
  if (!vendedor || !compradores || !Array.isArray(compradores)) {
    return res.status(400).json({ error: 'Datos incompletos' });
  }
  const entrada = { vendedor: vendedor.trim(), compradores, fechaEnvio: new Date().toISOString() };
  await coleccion().updateOne(
    { vendedor: { $regex: new RegExp(`^${vendedor.trim()}$`, 'i') } },
    { $set: entrada },
    { upsert: true }
  );
  res.json({ ok: true });
});

// Admin: ver datos
app.get('/api/admin/ventas', async (req, res) => {
  if (req.query.password !== ADMIN_PASSWORD) return res.status(401).json({ error: 'No autorizado' });
  const ventas = await coleccion().find({}).toArray();
  res.json(ventas);
});

// Admin: exportar Excel
app.get('/api/admin/exportar', async (req, res) => {
  if (req.query.password !== ADMIN_PASSWORD) return res.status(401).json({ error: 'No autorizado' });

  const ventas = await coleccion().find({}).toArray();
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

conectarDB().then(() => {
  app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));
}).catch(err => {
  console.error('Error conectando MongoDB:', err);
  process.exit(1);
});

  
