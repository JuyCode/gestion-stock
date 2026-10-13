/* ============================================================
   GESTIÓN STOCK — Lógica de la aplicación
   Persistencia: localStorage
   ============================================================ */

const STORAGE_KEY = 'gestionStock_v1';

/* ---------- Estado ---------- */
const DEFAULT_CONFIG = {
  nombre: 'Mi Negocio',
  moneda: '$',
  stockMin: 10,   // <= rojo
  stockMedio: 20, // <= amarillo
  bgColor: '#f1f5f4',
  sidebarColor: '#14352a',
  logo: null,
  tutorialVisto: false
};

let state = {
  config: { ...DEFAULT_CONFIG },
  marcas: [],
  productos: [],
  ventas: [],
  gastos: [],
  clientes: []
};

let editandoProductoId = null;
let editandoClienteId = null;
let listaEntrada = []; // items del panel de entrada de mercadería

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

/* ---------- Utilidades ---------- */
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function fmt(n) {
  const v = Number(n) || 0;
  return state.config.moneda + v.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtFecha(iso) {
  const d = new Date(iso);
  return d.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' }) +
    ' ' + d.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });
}

function hoyMes() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function toast(msg, tipo = '') {
  const t = document.createElement('div');
  t.className = 'toast ' + tipo;
  t.textContent = msg;
  $('#toastContainer').appendChild(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transition = 'opacity .3s'; }, 2600);
  setTimeout(() => t.remove(), 3000);
}

function estadoStock(p) {
  if (p.stock <= state.config.stockMin) return 'bajo';
  if (p.stock <= state.config.stockMedio) return 'medio';
  return 'ok';
}

/* ---------- Persistencia ---------- */
function guardar() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function cargar() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const data = JSON.parse(raw);
      state = { ...state, ...data };
      // Aplicar los valores por defecto para las claves nuevas/ausentes
      state.config = { ...DEFAULT_CONFIG, ...(data.config || {}) };
    }
  } catch (e) {
    console.error('Error cargando datos', e);
  }
}

/* ---------- Exportar a Excel ---------- */
function fechaArchivo() {
  return new Date().toISOString().slice(0, 10);
}

function exportarDatos() {
  if (typeof XLSX === 'undefined') {
    toast('No se pudo cargar la librería de Excel', 'error');
    return;
  }

  const wb = XLSX.utils.book_new();

  /* ---- Hoja: Resumen ---- */
  const mes = hoyMes();
  const valorCosto = state.productos.reduce((s, p) => s + p.stock * p.compra, 0);
  const valorVenta = state.productos.reduce((s, p) => s + p.stock * p.venta, 0);
  const ventasMes = state.ventas.filter(v => v.fecha.startsWith(mes)).reduce((s, v) => s + v.total, 0);
  const resumen = [
    ['RESUMEN DEL NEGOCIO', ''],
    ['Negocio', state.config.nombre],
    ['Moneda', state.config.moneda],
    ['Generado', new Date()],
    ['', ''],
    ['Productos', state.productos.length],
    ['Unidades en stock', state.productos.reduce((s, p) => s + p.stock, 0)],
    ['Valor del stock (a costo)', valorCosto],
    ['Valor del stock (a venta)', valorVenta],
    ['Ganancia potencial', valorVenta - valorCosto],
    ['Productos con stock bajo', state.productos.filter(p => estadoStock(p) === 'bajo').length],
    ['Ventas del mes', ventasMes],
    ['Ventas totales', state.ventas.reduce((s, v) => s + v.total, 0)],
    ['Gastos totales', state.gastos.reduce((s, g) => s + g.monto, 0)],
    ['Clientes', state.clientes.length]
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumen), 'Resumen');

  /* ---- Hoja: Productos ---- */
  const prod = [['Producto', 'Marca', 'Stock', 'P. Compra', 'P. Venta', 'Ganancia unitaria', 'Valor en stock']];
  state.productos.forEach(p => prod.push([
    p.nombre, nombreMarca(p.marca), p.stock, p.compra, p.venta, p.venta - p.compra, p.stock * p.compra
  ]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(prod), 'Productos');

  /* ---- Hoja: Ventas ---- */
  const ven = [['Fecha', 'Cliente', 'Producto', 'Cantidad', 'Precio unitario', 'Subtotal']];
  state.ventas.forEach(v => {
    v.items.forEach(it => ven.push([
      new Date(v.fecha),
      v.clienteId ? nombreCliente(v.clienteId) : '',
      it.productoNombre, it.cantidad, it.precio, it.cantidad * it.precio
    ]));
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ven), 'Ventas');

  /* ---- Hoja: Gastos ---- */
  const gas = [['Fecha', 'Descripción', 'Categoría', 'Monto']];
  state.gastos.forEach(g => gas.push([new Date(g.fecha), g.descripcion, g.categoria || 'General', g.monto]));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(gas), 'Gastos');

  /* ---- Hoja: Clientes ---- */
  const cli = [['Nombre', 'Teléfono', 'Email', 'Notas', 'Total comprado', 'Última compra']];
  state.clientes.forEach(c => {
    const compras = state.ventas.filter(v => v.clienteId === c.id);
    const total = compras.reduce((s, v) => s + v.total, 0);
    const ultima = compras.length ? new Date(compras.map(v => v.fecha).sort().reverse()[0]) : '';
    cli.push([c.nombre, c.telefono || '', c.email || '', c.notas || '', total, ultima]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cli), 'Clientes');

  /* ---- Hoja: Marcas ---- */
  const mar = [['Marca'], ...state.marcas.map(m => [m.nombre])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mar), 'Marcas');

  /* ---- Hoja: Configuración ---- */
  const cfg = [
    ['Parámetro', 'Valor'],
    ['Nombre del negocio', state.config.nombre],
    ['Moneda', state.config.moneda],
    ['Stock mínimo (alerta roja)', state.config.stockMin],
    ['Stock medio (alerta amarilla)', state.config.stockMedio],
    ['Color de fondo', state.config.bgColor || '#f1f5f4'],
    ['Color barra lateral', state.config.sidebarColor || '#14352a']
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cfg), 'Configuración');

  XLSX.writeFile(wb, 'gestion-stock-' + fechaArchivo() + '.xlsx', { cellDates: true });
  toast('Respaldo exportado a Excel');
}

/* ---------- Importar respaldo ---------- */
function importarDatos(file) {
  const nombre = (file.name || '').toLowerCase();
  if (nombre.endsWith('.xlsx') || nombre.endsWith('.xls')) {
    importarExcel(file);
  } else {
    importarJSON(file);
  }
}

function importarJSON(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target.result);
      state = { ...state, ...data };
      state.config = { ...DEFAULT_CONFIG, ...(data.config || {}) };
      guardar();
      renderTodo();
      toast('Respaldo importado correctamente');
    } catch (err) {
      toast('El archivo no es válido', 'error');
    }
  };
  reader.readAsText(file);
}

function importarExcel(file) {
  if (typeof XLSX === 'undefined') {
    toast('No se pudo cargar la librería de Excel', 'error');
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
      const sheetRows = (name) => {
        const key = wb.SheetNames.find(n => n.toLowerCase().trim() === name.toLowerCase());
        if (!key) return null;
        return XLSX.utils.sheet_to_json(wb.Sheets[key], { header: 1, raw: true, defval: null });
      };

      const prodRows = sheetRows('Productos');
      if (!prodRows) { toast('El archivo no parece un respaldo válido', 'error'); return; }
      if (!confirm('Se reemplazarán los datos actuales por los del respaldo. ¿Continuar?')) return;

      const nuevasMarcas = [];
      const marcaIdPorNombre = new Map();
      const getMarcaId = (nombre) => {
        const key = String(nombre == null ? '' : nombre).trim();
        if (!key) return '';
        if (marcaIdPorNombre.has(key.toLowerCase())) return marcaIdPorNombre.get(key.toLowerCase());
        const id = uid();
        marcaIdPorNombre.set(key.toLowerCase(), id);
        nuevasMarcas.push({ id, nombre: key });
        return id;
      };

      const marRows = sheetRows('Marcas');
      if (marRows) marRows.slice(1).forEach(r => { if (r[0]) getMarcaId(r[0]); });

      const productos = [];
      const prodIdPorNombre = new Map();
      prodRows.slice(1).forEach(r => {
        const nombre = String(r[0] == null ? '' : r[0]).trim();
        if (!nombre) return;
        const id = uid();
        prodIdPorNombre.set(nombre.toLowerCase(), id);
        productos.push({
          id, nombre,
          marca: getMarcaId(r[1]),
          stock: Number(r[2]) || 0,
          compra: Number(r[3]) || 0,
          venta: Number(r[4]) || 0
        });
      });

      const clientes = [];
      const clienteIdPorNombre = new Map();
      const cliRows = sheetRows('Clientes');
      if (cliRows) {
        cliRows.slice(1).forEach(r => {
          const nombre = String(r[0] == null ? '' : r[0]).trim();
          if (!nombre) return;
          const id = uid();
          clienteIdPorNombre.set(nombre.toLowerCase(), id);
          clientes.push({
            id, nombre,
            telefono: String(r[1] == null ? '' : r[1]),
            email: String(r[2] == null ? '' : r[2]),
            notas: String(r[3] == null ? '' : r[3])
          });
        });
      }

      const ventas = [];
      const venRows = sheetRows('Ventas');
      if (venRows) {
        const grupos = new Map();
        venRows.slice(1).forEach(r => {
          const prodNombre = String(r[2] == null ? '' : r[2]).trim();
          if (!prodNombre) return;
          const fechaVal = r[0];
          const fechaISO = fechaVal instanceof Date ? fechaVal.toISOString() : new Date().toISOString();
          const clienteNombre = String(r[1] == null ? '' : r[1]).trim();
          const key = fechaISO + '||' + clienteNombre.toLowerCase();
          if (!grupos.has(key)) {
            grupos.set(key, {
              fecha: fechaISO,
              clienteId: clienteNombre ? (clienteIdPorNombre.get(clienteNombre.toLowerCase()) || null) : null,
              items: []
            });
          }
          grupos.get(key).items.push({
            productoId: prodIdPorNombre.get(prodNombre.toLowerCase()) || '',
            productoNombre: prodNombre,
            cantidad: Number(r[3]) || 0,
            precio: Number(r[4]) || 0
          });
        });
        grupos.forEach(g => {
          ventas.push({
            id: uid(), fecha: g.fecha, clienteId: g.clienteId, items: g.items,
            total: g.items.reduce((s, i) => s + i.cantidad * i.precio, 0)
          });
        });
      }

      const gastos = [];
      const gasRows = sheetRows('Gastos');
      if (gasRows) {
        gasRows.slice(1).forEach(r => {
          const desc = String(r[1] == null ? '' : r[1]).trim();
          if (!desc) return;
          const fechaVal = r[0];
          gastos.push({
            id: uid(),
            fecha: fechaVal instanceof Date ? fechaVal.toISOString() : new Date().toISOString(),
            descripcion: desc,
            categoria: String(r[2] == null ? 'General' : r[2]) || 'General',
            monto: Number(r[3]) || 0
          });
        });
      }

      const nuevaConfig = { ...state.config };
      const cfgRows = sheetRows('Configuración');
      if (cfgRows) {
        cfgRows.slice(1).forEach(r => {
          const k = String(r[0] == null ? '' : r[0]).toLowerCase();
          const v = r[1];
          if (k.includes('nombre')) nuevaConfig.nombre = String(v == null ? nuevaConfig.nombre : v);
          else if (k.includes('moneda')) nuevaConfig.moneda = String(v == null ? nuevaConfig.moneda : v);
          else if (k.includes('stock mínimo') || k.includes('stock minimo')) nuevaConfig.stockMin = Number(v) || nuevaConfig.stockMin;
          else if (k.includes('stock medio')) nuevaConfig.stockMedio = Number(v) || nuevaConfig.stockMedio;
        });
      }

      state = { ...state, config: nuevaConfig, marcas: nuevasMarcas, productos, ventas, gastos, clientes };
      guardar();
      renderTodo();
      toast('Respaldo importado desde Excel');
    } catch (err) {
      console.error(err);
      toast('No se pudo leer el archivo de Excel', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ---------- Datos de ejemplo (solo primera vez) ---------- */
function seed() {
  if (localStorage.getItem(STORAGE_KEY)) return;
  state.marcas = [
    { id: uid(), nombre: "Levi's" },
    { id: uid(), nombre: 'Zara' },
    { id: uid(), nombre: 'Nike' },
    { id: uid(), nombre: 'Mango' },
    { id: uid(), nombre: 'Adidas' },
    { id: uid(), nombre: 'H&M' },
    { id: uid(), nombre: 'Under Armour' },
    { id: uid(), nombre: 'Tommy Hilfiger' },
    { id: uid(), nombre: 'The North Face' },
    { id: uid(), nombre: 'Pull&Bear' }
  ];
  const [levis, zara, nike, mango, adidas, hm, under, tommy, tnf, pull] = state.marcas.map(m => m.id);
  state.productos = [
    { id: uid(), nombre: 'Camiseta Básica Algodón', marca: levis, stock: 50, compra: 15000, venta: 45000 },
    { id: uid(), nombre: 'Pantalón Vaquero Slim', marca: zara, stock: 30, compra: 40000, venta: 45000 },
    { id: uid(), nombre: 'Sudadera con Capucha', marca: nike, stock: 25, compra: 100000, venta: 145600 },
    { id: uid(), nombre: 'Chaqueta de Cuero', marca: mango, stock: 12, compra: 85600, venta: 124000 },
    { id: uid(), nombre: 'Zapatillas Casuales', marca: adidas, stock: 18, compra: 5000, venta: 100000 },
    { id: uid(), nombre: 'Vestido Corto Verano', marca: hm, stock: 20, compra: 100000, venta: 200000 },
    { id: uid(), nombre: 'Pantalón Deportivo', marca: under, stock: 35, compra: 155000, venta: 200000 },
    { id: uid(), nombre: 'Camisa Formal Manga Larga', marca: tommy, stock: 15, compra: 200000, venta: 220000 },
    { id: uid(), nombre: 'Chaqueta Impermeable', marca: tnf, stock: 10, compra: 65000, venta: 89500 },
    { id: uid(), nombre: 'Falda Plisada', marca: pull, stock: 22, compra: 90500, venta: 100500 }
  ];
  guardar();
}

function nombreMarca(id) {
  const m = state.marcas.find(m => m.id === id);
  return m ? m.nombre : '—';
}

/* ============================================================
   RENDERIZADO
   ============================================================ */
function renderTodo() {
  $('#brandTitle').textContent = state.config.nombre;
  // Logo
  const logoImg = $('#brandLogo');
  const iconDefault = $('#brandIconDefault');
  if (state.config.logo) {
    logoImg.src = state.config.logo;
    logoImg.hidden = false;
    iconDefault.hidden = true;
  } else {
    logoImg.hidden = true;
    iconDefault.hidden = false;
  }
  // Colores personalizados
  document.documentElement.style.setProperty('--bg', state.config.bgColor);
  document.documentElement.style.setProperty('--sidebar-bg', state.config.sidebarColor);
  renderDashboard();
  renderStock();
  renderVentas();
  renderGastos();
  renderClientes();
  renderConfig();
  renderSelectsMarca();
  updateStockBadge();
}

/* ---------- Dashboard ---------- */
function renderDashboard() {
  $('#statProductos').textContent = state.productos.length;
  $('#statUnidades').textContent = state.productos.reduce((s, p) => s + p.stock, 0);
  $('#statValor').textContent = fmt(state.productos.reduce((s, p) => s + p.stock * p.compra, 0));
  $('#statGanancia').textContent = fmt(state.productos.reduce((s, p) => s + p.stock * (p.venta - p.compra), 0));

  const bajos = state.productos.filter(p => estadoStock(p) === 'bajo');
  $('#statStockBajo').textContent = bajos.length;

  const mes = hoyMes();
  const ventasMes = state.ventas.filter(v => v.fecha.startsWith(mes));
  $('#statVentasMes').textContent = fmt(ventasMes.reduce((s, v) => s + v.total, 0));

  // Alertas de stock bajo
  const alertasEl = $('#alertasStock');
  if (bajos.length === 0) {
    alertasEl.innerHTML = '<p class="empty-state">No hay alertas. ¡Tu stock está saludable!</p>';
  } else {
    alertasEl.innerHTML = bajos.map(p => `
      <div class="alerta-item">
        <span class="stock-pill stock-bajo">${p.stock}</span>
        <div class="alerta-info">
          <div class="alerta-nombre">${esc(p.nombre)}</div>
          <div class="alerta-detalle">${esc(nombreMarca(p.marca))} · Reposición urgente</div>
        </div>
        <button class="btn btn-primary btn-sm" onclick="abrirEntradaPara('${p.id}')">Reponer</button>
      </div>
    `).join('');
  }

  // Últimas ventas
  const uvEl = $('#ultimasVentas');
  const ultimas = [...state.ventas].sort((a, b) => b.fecha.localeCompare(a.fecha)).slice(0, 6);
  if (ultimas.length === 0) {
    uvEl.innerHTML = '<p class="empty-state">Todavía no registraste ventas.</p>';
  } else {
    uvEl.innerHTML = ultimas.map(v => `
      <div class="venta-item">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="color:var(--text-muted)"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
        <div class="alerta-info">
          <div class="alerta-nombre">${v.items.length} producto${v.items.length !== 1 ? 's' : ''}</div>
          <div class="alerta-detalle">${fmtFecha(v.fecha)}${v.clienteId ? ' · ' + esc(nombreCliente(v.clienteId)) : ''}</div>
        </div>
        <strong>${fmt(v.total)}</strong>
      </div>
    `).join('');
  }
}

function nombreCliente(id) {
  const c = state.clientes.find(c => c.id === id);
  return c ? c.nombre : '—';
}

function updateStockBadge() {
  const n = state.productos.filter(p => estadoStock(p) === 'bajo').length;
  const badge = $('#stockAlertBadge');
  badge.hidden = n === 0;
  badge.textContent = n;
}

/* ---------- Stock ---------- */
function renderStock() {
  const busqueda = ($('#buscarProducto').value || '').toLowerCase();
  const marcaFiltro = $('#filtroMarcaStock').value;
  const estadoFiltro = $('#filtroEstado').value;

  // Filtro de marcas
  $('#filtroMarcaStock').innerHTML = '<option value="">Todas las marcas</option>' +
    state.marcas.map(m => `<option value="${m.id}" ${m.id === marcaFiltro ? 'selected' : ''}>${esc(m.nombre)}</option>`).join('');

  let lista = state.productos.filter(p => {
    if (busqueda && !(p.nombre.toLowerCase().includes(busqueda) || nombreMarca(p.marca).toLowerCase().includes(busqueda))) return false;
    if (marcaFiltro && p.marca !== marcaFiltro) return false;
    if (estadoFiltro && estadoStock(p) !== estadoFiltro) return false;
    return true;
  });

  lista.sort((a, b) => {
    const ea = estadoStock(a) === 'bajo' ? 0 : estadoStock(a) === 'medio' ? 1 : 2;
    const eb = estadoStock(b) === 'bajo' ? 0 : estadoStock(b) === 'medio' ? 1 : 2;
    if (ea !== eb) return ea - eb;
    return a.nombre.localeCompare(b.nombre);
  });

  const tbody = $('#tbodyStock');
  if (lista.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8"><p class="empty-state">No hay productos para mostrar.</p></td></tr>';
    return;
  }

  tbody.innerHTML = lista.map(p => {
    const est = estadoStock(p);
    const pillClass = est === 'bajo' ? 'stock-bajo' : est === 'medio' ? 'stock-medio' : 'stock-ok';
    const ganancia = p.venta - p.compra;
    return `
      <tr>
        <td class="producto-nombre">${esc(p.nombre)}</td>
        <td><span class="marca-tag">${esc(nombreMarca(p.marca))}</span></td>
        <td class="col-stock"><span class="stock-pill ${pillClass}">${p.stock}</span></td>
        <td class="num">${fmt(p.compra)}</td>
        <td class="num">${fmt(p.venta)}</td>
        <td class="num" style="color:var(--success);font-weight:600">${fmt(ganancia)}</td>
        <td class="num">${fmt(p.stock * p.compra)}</td>
        <td>
          <div class="acciones">
            <button class="icon-btn" title="Entrada de stock" onclick="abrirEntradaPara('${p.id}')">📥</button>
            <button class="icon-btn" title="Editar" onclick="abrirModalProducto('${p.id}')">✏️</button>
            <button class="icon-btn" title="Eliminar" onclick="eliminarProducto('${p.id}')">🗑️</button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

/* ---------- Ventas ---------- */
function renderVentas() {
  const busqueda = ($('#buscarVenta').value || '').toLowerCase();
  let lista = [...state.ventas].sort((a, b) => b.fecha.localeCompare(a.fecha));
  if (busqueda) {
    lista = lista.filter(v => {
      const txt = v.items.map(i => i.productoNombre).join(' ').toLowerCase();
      return txt.includes(busqueda) || (v.clienteId && nombreCliente(v.clienteId).toLowerCase().includes(busqueda));
    });
  }

  const tbody = $('#tbodyVentas');
  if (lista.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4"><p class="empty-state">No hay ventas registradas.</p></td></tr>';
    return;
  }

  tbody.innerHTML = lista.map(v => `
    <tr>
      <td>${fmtFecha(v.fecha)}</td>
      <td>${v.items.map(i => esc(i.productoNombre) + ' ×' + i.cantidad).join(', ')}</td>
      <td class="num" style="font-weight:700">${fmt(v.total)}</td>
      <td>
        <div class="acciones">
          <button class="icon-btn" title="Eliminar venta (devuelve stock)" onclick="eliminarVenta('${v.id}')">🗑️</button>
        </div>
      </td>
    </tr>
  `).join('');
}

/* ---------- Gastos ---------- */
function renderGastos() {
  const cats = [...new Set(state.gastos.map(g => g.categoria).filter(Boolean))];
  const filtro = $('#filtroCategoriaGasto').value;
  $('#filtroCategoriaGasto').innerHTML = '<option value="">Todas las categorías</option>' +
    cats.map(c => `<option value="${esc(c)}" ${c === filtro ? 'selected' : ''}>${esc(c)}</option>`).join('');
  $('#categoriasGasto').innerHTML = cats.map(c => `<option value="${esc(c)}">`).join('');

  let lista = [...state.gastos].sort((a, b) => b.fecha.localeCompare(a.fecha));
  if (filtro) lista = lista.filter(g => g.categoria === filtro);

  const tbody = $('#tbodyGastos');
  if (lista.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5"><p class="empty-state">No hay gastos registrados.</p></td></tr>';
    return;
  }

  tbody.innerHTML = lista.map(g => `
    <tr>
      <td>${fmtFecha(g.fecha)}</td>
      <td class="producto-nombre">${esc(g.descripcion)}</td>
      <td><span class="marca-tag">${esc(g.categoria || 'General')}</span></td>
      <td class="num" style="color:var(--danger);font-weight:600">${fmt(g.monto)}</td>
      <td>
        <div class="acciones">
          <button class="icon-btn" title="Eliminar" onclick="eliminarGasto('${g.id}')">🗑️</button>
        </div>
      </td>
    </tr>
  `).join('');
}

/* ---------- Clientes ---------- */
function renderClientes() {
  const busqueda = ($('#buscarCliente').value || '').toLowerCase();
  let lista = [...state.clientes];
  if (busqueda) {
    lista = lista.filter(c =>
      c.nombre.toLowerCase().includes(busqueda) ||
      (c.email || '').toLowerCase().includes(busqueda) ||
      (c.telefono || '').toLowerCase().includes(busqueda)
    );
  }

  const tbody = $('#tbodyClientes');
  if (lista.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5"><p class="empty-state">No hay clientes registrados.</p></td></tr>';
    return;
  }

  tbody.innerHTML = lista.map(c => {
    const compras = state.ventas.filter(v => v.clienteId === c.id);
    const total = compras.reduce((s, v) => s + v.total, 0);
    const ultima = compras.length ? compras.map(v => v.fecha).sort().reverse()[0] : null;
    return `
      <tr>
        <td class="producto-nombre">${esc(c.nombre)}</td>
        <td>
          ${c.telefono ? '📱 ' + esc(c.telefono) : ''}
          ${c.email ? '<br>✉️ ' + esc(c.email) : ''}
          ${!c.telefono && !c.email ? '<span class="muted">—</span>' : ''}
        </td>
        <td class="num" style="font-weight:600">${fmt(total)}</td>
        <td>${ultima ? fmtFecha(ultima) : '<span class="muted">—</span>'}</td>
        <td>
          <div class="acciones">
            <button class="icon-btn" title="Editar" onclick="abrirModalCliente('${c.id}')">✏️</button>
            <button class="icon-btn" title="Eliminar" onclick="eliminarCliente('${c.id}')">🗑️</button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

/* ---------- Config ---------- */
function renderConfig() {
  $('#cfgNombre').value = state.config.nombre;
  $('#cfgMoneda').value = state.config.moneda;
  $('#cfgStockMin').value = state.config.stockMin;
  $('#cfgStockMedio').value = state.config.stockMedio;
  $('#cfgBgColor').value = state.config.bgColor || '#f1f5f4';
  $('#cfgBgColorText').value = state.config.bgColor || '#f1f5f4';
  $('#cfgSidebarColor').value = state.config.sidebarColor || '#14352a';
  $('#cfgSidebarColorText').value = state.config.sidebarColor || '#14352a';

  $('#listaMarcas').innerHTML = state.marcas.map(m => `
    <li>
      <span>${esc(m.nombre)}</span>
      <button class="icon-btn" title="Eliminar marca" onclick="eliminarMarca('${m.id}')">✕</button>
    </li>
  `).join('');
}

function renderSelectsMarca() {
  const opts = state.marcas.map(m => `<option value="${m.id}">${esc(m.nombre)}</option>`).join('');
  ['#prodMarca', '#selMarcaNuevo'].forEach(s => {
    const el = $(s);
    if (el) el.innerHTML = '<option value="">Selecciona Marca</option>' + opts;
  });

  // Select de marca en entrada (con opción "nueva")
  const selEntrada = $('#selMarcaEntrada');
  if (selEntrada) {
    selEntrada.innerHTML = '<option value="">Selecciona Marca</option>' + opts;
  }
}

/* ============================================================
   NAVEGACIÓN
   ============================================================ */
const TITULOS = {
  dashboard: 'Panel de Control',
  stock: 'Stock de Productos',
  ventas: 'Ventas',
  gastos: 'Gastos',
  crm: 'Mini CRM — Clientes',
  config: 'Configuración'
};

function switchView(view) {
  $$('.nav-item').forEach(b => b.classList.toggle('active', b.dataset.view === view));
  $$('.view').forEach(v => v.classList.toggle('active', v.id === 'view-' + view));
  $('#viewTitle').textContent = TITULOS[view] || '';
}

/* ============================================================
   PANEL: ENTRADA DE MERCADERÍA
   ============================================================ */
function abrirPanelEntrada() {
  listaEntrada = [];
  renderListaEntrada();
  $('#overlayEntrada').classList.add('open');
  $('#panelEntrada').classList.add('open');
  cargarProductosParaEntrada();
  switchTabEntrada('existente');
  $('#chkMarcaNueva').checked = false;
  $('#inputMarcaNuevaWrap').hidden = true;
  $('#cantidadEntrada').value = 0;
}

function cerrarPanelEntrada() {
  $('#overlayEntrada').classList.remove('open');
  $('#panelEntrada').classList.remove('open');
}

function switchTabEntrada(tab) {
  $('#tabExistente').classList.toggle('active', tab === 'existente');
  $('#tabNuevo').classList.toggle('active', tab === 'nuevo');
  $('#panelExistente').classList.toggle('active', tab === 'existente');
  $('#panelNuevo').classList.toggle('active', tab === 'nuevo');
}

function cargarProductosParaEntrada() {
  const sel = $('#selProductoEntrada');
  sel.innerHTML = '<option value="">Selecciona Producto</option>' +
    state.productos.map(p => `<option value="${p.id}">${esc(p.nombre)} (${nombreMarca(p.marca)})</option>`).join('');
}

function renderListaEntrada() {
  const ul = $('#listaEntrada');
  $('#listaCount').textContent = listaEntrada.length;
  $('#listaEmpty').style.display = listaEntrada.length === 0 ? 'block' : 'none';

  let totalUnits = 0;
  let totalCosto = 0;

  ul.innerHTML = listaEntrada.map((item, i) => {
    totalUnits += item.cantidad;
    totalCosto += item.cantidad * item.compra;
    return `
      <li class="sp-list-item">
        <div class="sp-list-info">
          <div class="sp-list-nombre">${esc(item.nombre)}</div>
          <div class="sp-list-detalle">${esc(item.marcaNombre)} · ${fmt(item.compra)} c/u</div>
        </div>
        <span class="sp-list-cant">+${item.cantidad}</span>
        <button class="icon-btn" title="Quitar" onclick="quitarItemLista(${i})">✕</button>
      </li>
    `;
  }).join('');

  $('#totalUnidadesEntrada').textContent = totalUnits;
  $('#costoTotalEntrada').textContent = fmt(totalCosto);
  $('#btnConfirmarEntrada').disabled = listaEntrada.length === 0;
}

function quitarItemLista(i) {
  listaEntrada.splice(i, 1);
  renderListaEntrada();
}

/* Abrir panel precargado con un producto (desde alertas o acciones) */
function abrirEntradaPara(productoId) {
  abrirPanelEntrada();
  switchTabEntrada('existente');
  $('#selProductoEntrada').value = productoId;
  $('#cantidadEntrada').focus();
}

/* Agregar item desde tab "Existente" */
function agregarDesdeExistente() {
  let marcaId = $('#selMarcaEntrada').value;
  if ($('#chkMarcaNueva').checked) {
    const nombre = $('#inputMarcaNueva').value.trim();
    if (!nombre) { toast('Escribí el nombre de la nueva marca', 'error'); return; }
    const existente = state.marcas.find(m => m.nombre.toLowerCase() === nombre.toLowerCase());
    if (existente) {
      marcaId = existente.id;
    } else {
      const nueva = { id: uid(), nombre };
      state.marcas.push(nueva);
      marcaId = nueva.id;
      renderSelectsMarca();
      $('#chkMarcaNueva').checked = false;
      $('#inputMarcaNuevaWrap').hidden = true;
      $('#selMarcaEntrada').value = marcaId;
    }
  }
  if (!marcaId) { toast('Seleccioná una marca', 'error'); return; }

  const productoId = $('#selProductoEntrada').value;
  if (!productoId) { toast('Seleccioná un producto', 'error'); return; }

  const cantidad = parseInt($('#cantidadEntrada').value, 10);
  if (!cantidad || cantidad < 1) { toast('Ingresá una cantidad válida', 'error'); return; }

  const prod = state.productos.find(p => p.id === productoId);
  if (!prod) { toast('Producto no encontrado', 'error'); return; }
  if (prod.marca !== marcaId) {
    toast('La marca seleccionada no coincide con el producto', 'error');
    return;
  }

  const existente = listaEntrada.find(i => i.productoId === productoId);
  if (existente) {
    existente.cantidad += cantidad;
  } else {
    listaEntrada.push({
      productoId,
      nombre: prod.nombre,
      marcaNombre: nombreMarca(marcaId),
      compra: prod.compra,
      cantidad
    });
  }

  $('#cantidadEntrada').value = 0;
  renderListaEntrada();
  toast('Agregado a la lista', 'success');
}

/* Crear producto nuevo y agregar a la lista */
function crearProductoYAgregar() {
  const marcaId = $('#selMarcaNuevo').value;
  if (!marcaId) { toast('Seleccioná una marca', 'error'); return; }

  const nombre = $('#nuevoProductoNombre').value.trim();
  if (!nombre) { toast('Escribí el nombre del producto', 'error'); return; }

  const compra = parseFloat($('#nuevoProductoCompra').value) || 0;
  const venta = parseFloat($('#nuevoProductoVenta').value) || 0;
  if (compra <= 0 || venta <= 0) { toast('Ingresá precios válidos (mayores a 0)', 'error'); return; }

  const cantidad = parseInt($('#nuevoProductoCantidad').value, 10);
  if (!cantidad || cantidad < 1) { toast('Ingresá una cantidad válida', 'error'); return; }

  const prod = { id: uid(), nombre, marca: marcaId, stock: 0, compra, venta };
  state.productos.push(prod);

  listaEntrada.push({
    productoId: prod.id,
    nombre: prod.nombre,
    marcaNombre: nombreMarca(marcaId),
    compra,
    cantidad
  });

  $('#nuevoProductoNombre').value = '';
  $('#nuevoProductoCompra').value = '';
  $('#nuevoProductoVenta').value = '';
  $('#nuevoProductoCantidad').value = 1;

  cargarProductosParaEntrada();
  renderListaEntrada();
  toast('Producto creado y agregado a la lista', 'success');
}

function confirmarEntrada() {
  if (listaEntrada.length === 0) return;
  let total = 0;
  listaEntrada.forEach(item => {
    const prod = state.productos.find(p => p.id === item.productoId);
    if (prod) {
      prod.stock += item.cantidad;
      total += item.cantidad;
    }
  });
  guardar();
  renderTodo();
  cerrarPanelEntrada();
  toast(`Entrada confirmada: ${total} unidad${total !== 1 ? 'es' : ''} sumada${total !== 1 ? 's' : ''} al stock`, 'success');
}

/* ============================================================
   MODAL: PRODUCTO
   ============================================================ */
function abrirModalProducto(id = null) {
  editandoProductoId = id;
  $('#modalProductoTitle').textContent = id ? 'Editar producto' : 'Nuevo producto';
  const prod = id ? state.productos.find(p => p.id === id) : null;
  $('#prodNombre').value = prod ? prod.nombre : '';
  $('#prodMarca').value = prod ? prod.marca : (state.marcas[0] ? state.marcas[0].id : '');
  $('#prodStock').value = prod ? prod.stock : 0;
  $('#prodCompra').value = prod ? prod.compra : '';
  $('#prodVenta').value = prod ? prod.venta : '';
  $('#overlayProducto').classList.add('open');
  $('#modalProducto').classList.add('open');
}

function guardarProducto() {
  const nombre = $('#prodNombre').value.trim();
  if (!nombre) { toast('Escribí el nombre del producto', 'error'); return; }
  const marca = $('#prodMarca').value;
  if (!marca) { toast('Seleccioná una marca', 'error'); return; }
  const stock = parseInt($('#prodStock').value, 10) || 0;
  const compra = parseFloat($('#prodCompra').value) || 0;
  const venta = parseFloat($('#prodVenta').value) || 0;

  if (editandoProductoId) {
    const prod = state.productos.find(p => p.id === editandoProductoId);
    if (prod) Object.assign(prod, { nombre, marca, stock, compra, venta });
    toast('Producto actualizado', 'success');
  } else {
    state.productos.push({ id: uid(), nombre, marca, stock, compra, venta });
    toast('Producto creado', 'success');
  }
  guardar();
  cerrarModales();
  renderTodo();
}

function eliminarProducto(id) {
  const prod = state.productos.find(p => p.id === id);
  if (!prod) return;
  if (!confirm(`¿Eliminar "${prod.nombre}" del stock? Esta acción no se puede deshacer.`)) return;
  state.productos = state.productos.filter(p => p.id !== id);
  guardar();
  renderTodo();
  toast('Producto eliminado', 'success');
}

/* ============================================================
   MODAL: VENTA
   ============================================================ */
let ventaItems = [];

function abrirModalVenta() {
  ventaItems = [];
  $('#ventaCliente').innerHTML = '<option value="">— Venta sin cliente —</option>' +
    state.clientes.map(c => `<option value="${c.id}">${esc(c.nombre)}</option>`).join('');
  renderVentaItems();
  $('#overlayVenta').classList.add('open');
  $('#modalVenta').classList.add('open');
  agregarItemVenta();
}

function renderVentaItems() {
  const wrap = $('#ventaItems');
  if (ventaItems.length === 0) {
    wrap.innerHTML = '<p class="muted">Agregá al menos un producto.</p>';
    return;
  }
  wrap.innerHTML = ventaItems.map((item, i) => {
    const options = state.productos.map(p =>
      `<option value="${p.id}" ${p.id === item.productoId ? 'selected' : ''}>${esc(p.nombre)} (stock: ${p.stock})</option>`
    ).join('');
    return `
      <div class="venta-item-row">
        <select class="input" onchange="cambiarItemVenta(${i}, 'productoId', this.value)">
          <option value="">— Seleccionar —</option>${options}
        </select>
        <input type="number" class="input" min="1" value="${item.cantidad}" onchange="cambiarItemVenta(${i}, 'cantidad', this.value)">
        <span class="venta-item-subtotal" id="subtotal-${i}">${fmt(item.subtotal)}</span>
        <button class="icon-btn" onclick="quitarItemVenta(${i})">✕</button>
      </div>
    `;
  }).join('');
}

function agregarItemVenta() {
  ventaItems.push({ productoId: '', cantidad: 1, precio: 0, subtotal: 0 });
  renderVentaItems();
}

function cambiarItemVenta(i, campo, valor) {
  const item = ventaItems[i];
  if (campo === 'productoId') {
    item.productoId = valor;
    const prod = state.productos.find(p => p.id === valor);
    item.precio = prod ? prod.venta : 0;
  } else {
    item.cantidad = parseInt(valor, 10) || 0;
  }
  item.subtotal = (item.cantidad || 0) * (item.precio || 0);
  renderVentaItems();
}

function quitarItemVenta(i) {
  ventaItems.splice(i, 1);
  renderVentaItems();
}

function guardarVenta() {
  const items = ventaItems.filter(i => i.productoId && i.cantidad > 0);
  if (items.length === 0) { toast('Agregá al menos un producto con cantidad', 'error'); return; }

  // Validar stock disponible
  for (const item of items) {
    const prod = state.productos.find(p => p.id === item.productoId);
    if (!prod) { toast('Producto no encontrado', 'error'); return; }
    if (item.cantidad > prod.stock) {
      toast(`Stock insuficiente de "${prod.nombre}" (disponible: ${prod.stock})`, 'error');
      return;
    }
  }

  const total = items.reduce((s, i) => s + i.subtotal, 0);
  const venta = {
    id: uid(),
    fecha: new Date().toISOString(),
    clienteId: $('#ventaCliente').value || null,
    items: items.map(i => {
      const prod = state.productos.find(p => p.id === i.productoId);
      return { productoId: i.productoId, productoNombre: prod.nombre, cantidad: i.cantidad, precio: i.precio };
    }),
    total
  };

  // Descontar stock
  items.forEach(i => {
    const prod = state.productos.find(p => p.id === i.productoId);
    prod.stock -= i.cantidad;
  });

  state.ventas.push(venta);
  guardar();
  cerrarModales();
  renderTodo();
  toast(`Venta registrada: ${fmt(total)}`, 'success');
}

function eliminarVenta(id) {
  const venta = state.ventas.find(v => v.id === id);
  if (!venta) return;
  if (!confirm('¿Eliminar esta venta? Las unidades volverán al stock.')) return;
  venta.items.forEach(i => {
    const prod = state.productos.find(p => p.id === i.productoId);
    if (prod) prod.stock += i.cantidad;
  });
  state.ventas = state.ventas.filter(v => v.id !== id);
  guardar();
  renderTodo();
  toast('Venta eliminada, stock restaurado', 'success');
}

/* ============================================================
   MODAL: GASTO
   ============================================================ */
function abrirModalGasto() {
  $('#gastoDescripcion').value = '';
  $('#gastoCategoria').value = '';
  $('#gastoMonto').value = '';
  $('#overlayGasto').classList.add('open');
  $('#modalGasto').classList.add('open');
}

function guardarGasto() {
  const descripcion = $('#gastoDescripcion').value.trim();
  if (!descripcion) { toast('Escribí una descripción', 'error'); return; }
  const monto = parseFloat($('#gastoMonto').value) || 0;
  if (monto <= 0) { toast('Ingresá un monto válido', 'error'); return; }

  state.gastos.push({
    id: uid(),
    fecha: new Date().toISOString(),
    descripcion,
    categoria: $('#gastoCategoria').value.trim() || 'General',
    monto
  });
  guardar();
  cerrarModales();
  renderTodo();
  toast('Gasto registrado', 'success');
}

function eliminarGasto(id) {
  if (!confirm('¿Eliminar este gasto?')) return;
  state.gastos = state.gastos.filter(g => g.id !== id);
  guardar();
  renderTodo();
  toast('Gasto eliminado', 'success');
}

/* ============================================================
   MODAL: CLIENTE
   ============================================================ */
function abrirModalCliente(id = null) {
  editandoClienteId = id;
  $('#modalClienteTitle').textContent = id ? 'Editar cliente' : 'Nuevo cliente';
  const c = id ? state.clientes.find(c => c.id === id) : null;
  $('#cliNombre').value = c ? c.nombre : '';
  $('#cliTelefono').value = c ? (c.telefono || '') : '';
  $('#cliEmail').value = c ? (c.email || '') : '';
  $('#cliNotas').value = c ? (c.notas || '') : '';
  $('#overlayCliente').classList.add('open');
  $('#modalCliente').classList.add('open');
}

function guardarCliente() {
  const nombre = $('#cliNombre').value.trim();
  if (!nombre) { toast('Escribí el nombre del cliente', 'error'); return; }

  const datos = {
    nombre,
    telefono: $('#cliTelefono').value.trim(),
    email: $('#cliEmail').value.trim(),
    notas: $('#cliNotas').value.trim()
  };

  if (editandoClienteId) {
    const c = state.clientes.find(c => c.id === editandoClienteId);
    if (c) Object.assign(c, datos);
    toast('Cliente actualizado', 'success');
  } else {
    state.clientes.push({ id: uid(), ...datos });
    toast('Cliente creado', 'success');
  }
  guardar();
  cerrarModales();
  renderTodo();
}

function eliminarCliente(id) {
  const c = state.clientes.find(c => c.id === id);
  if (!c) return;
  if (!confirm(`¿Eliminar a "${c.nombre}"? Sus ventas se conservarán como "sin cliente".`)) return;
  state.clientes = state.clientes.filter(c => c.id !== id);
  guardar();
  renderTodo();
  toast('Cliente eliminado', 'success');
}

/* ============================================================
   MARCAS
   ============================================================ */
function agregarMarca() {
  const input = $('#nuevaMarca');
  const nombre = input.value.trim();
  if (!nombre) return;
  if (state.marcas.some(m => m.nombre.toLowerCase() === nombre.toLowerCase())) {
    toast('Esa marca ya existe', 'error');
    return;
  }
  state.marcas.push({ id: uid(), nombre });
  input.value = '';
  guardar();
  renderTodo();
  toast('Marca agregada', 'success');
}

function eliminarMarca(id) {
  const prodCount = state.productos.filter(p => p.marca === id).length;
  const msg = prodCount > 0
    ? `Esta marca tiene ${prodCount} producto${prodCount !== 1 ? 's' : ''} asociado${prodCount !== 1 ? 's' : ''}. ¿Eliminar de todas formas?`
    : '¿Eliminar esta marca?';
  if (!confirm(msg)) return;
  state.marcas = state.marcas.filter(m => m.id !== id);
  state.productos.forEach(p => { if (p.marca === id) p.marca = ''; });
  guardar();
  renderTodo();
  toast('Marca eliminada', 'success');
}

/* ============================================================
   CONFIGURACIÓN
   ============================================================ */
function guardarConfig() {
  state.config.nombre = $('#cfgNombre').value.trim() || 'Mi Negocio';
  state.config.moneda = $('#cfgMoneda').value.trim() || '$';
  state.config.stockMin = parseInt($('#cfgStockMin').value, 10) || 1;
  state.config.stockMedio = parseInt($('#cfgStockMedio').value, 10) || state.config.stockMin;
  state.config.bgColor = $('#cfgBgColor').value || '#f1f5f4';
  $('#cfgBgColorText').value = state.config.bgColor;
  state.config.sidebarColor = $('#cfgSidebarColor').value || '#14352a';
  $('#cfgSidebarColorText').value = state.config.sidebarColor;
  guardar();
  renderTodo();
  toast('Configuración guardada', 'success');
}

function cargarLogo(file) {
  if (!file || !file.type.startsWith('image/')) {
    toast('Seleccioná una imagen válida', 'error');
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    // Redimensionar para que no pese tanto en localStorage
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const maxSize = 200;
      let { width, height } = img;
      if (width > height) {
        if (width > maxSize) { height *= maxSize / width; width = maxSize; }
      } else {
        if (height > maxSize) { width *= maxSize / height; height = maxSize; }
      }
      canvas.width = width;
      canvas.height = height;
      canvas.getContext('2d').drawImage(img, 0, 0, width, height);
      state.config.logo = canvas.toDataURL('image/jpeg', 0.85);
      guardar();
      renderTodo();
      toast('Logo actualizado', 'success');
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function resetearTodo() {
  if (!confirm('⚠️ ¿Restablecer TODOS los datos? Se borrarán productos, ventas, gastos y clientes. Esta acción no se puede deshacer.')) return;
  if (!confirm('¿Estás seguro/a? Es la última oportunidad para cancelar.')) return;
  localStorage.removeItem(STORAGE_KEY);
  location.reload();
}

/* ============================================================
   HELPERS GENERALES
   ============================================================ */
function esc(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function cerrarModales() {
  $$('.modal.open').forEach(m => m.classList.remove('open'));
  $$('.overlay').forEach(o => {
    if (o.id !== 'overlayEntrada' && o.id !== 'overlayBienvenida') o.classList.remove('open');
  });
}

/* ============================================================
   BIENVENIDA + TUTORIAL
   ============================================================ */
function abrirBienvenida(esPrimeraVez) {
  const checkWrap = document.querySelector('.bienvenida-check');
  if (checkWrap) checkWrap.style.display = esPrimeraVez ? '' : 'none';
  $('#chkCargarEjemplos').checked = !!esPrimeraVez;
  $('#modalBienvenida').dataset.primeraVez = esPrimeraVez ? '1' : '0';
  $('#overlayBienvenida').classList.add('open');
  $('#modalBienvenida').classList.add('open');
}

function cerrarBienvenida() {
  $('#overlayBienvenida').classList.remove('open');
  $('#modalBienvenida').classList.remove('open');
}

function vaciarDatos() {
  state.marcas = [];
  state.productos = [];
  state.ventas = [];
  state.gastos = [];
  state.clientes = [];
}

function confirmarBienvenida(verTutorial) {
  const esPrimeraVez = $('#modalBienvenida').dataset.primeraVez === '1';
  if (esPrimeraVez && !$('#chkCargarEjemplos').checked) {
    vaciarDatos();
  }
  state.config.tutorialVisto = true;
  guardar();
  renderTodo();
  cerrarBienvenida();
  if (verTutorial) {
    setTimeout(iniciarTour, 350);
  }
}

const TOUR_PASOS = [
  { sel: '.sidebar-brand', view: 'dashboard', titulo: 'Tu espacio de trabajo', texto: '¡Te damos la bienvenida! Desde el menú de la izquierda accedés a todas las secciones de tu negocio.' },
  { sel: '[data-view="dashboard"]', view: 'dashboard', titulo: 'Inicio', texto: 'El panel principal. De un vistazo ves tus métricas, las alertas de stock bajo y tus últimas ventas.' },
  { sel: '[data-view="stock"]', view: 'stock', titulo: 'Stock', texto: 'Tu inventario completo. Cargá productos, editá precios y mirá el semáforo de stock: verde, amarillo y rojo.' },
  { sel: '[data-view="ventas"]', view: 'ventas', titulo: 'Ventas', texto: 'Registrá cada venta. El stock se descuenta solo y podés asociarla a un cliente.' },
  { sel: '[data-view="gastos"]', view: 'gastos', titulo: 'Gastos', texto: 'Anotá tus gastos por categoría para conocer tu ganancia real.' },
  { sel: '[data-view="crm"]', view: 'crm', titulo: 'Mini CRM', texto: 'Tu lista de clientes con contacto, notas y su historial de compras.' },
  { sel: '[data-view="config"]', view: 'config', titulo: 'Configuración', texto: 'Personalizá el nombre, el logo y los colores, y ajustá los niveles de alerta de stock.' },
  { sel: '#btnExport', view: 'config', titulo: 'Respaldá tus datos', texto: 'Descargá todo en Excel cuando quieras. Si cambiás de dispositivo, lo volvés a cargar con "Importar respaldo".' }
];

let tourIndex = 0;

function iniciarTour() {
  tourIndex = 0;
  $('#tourOverlay').hidden = false;
  mostrarPasoTour();
}

function cerrarTour() {
  $('#tourOverlay').hidden = true;
  state.config.tutorialVisto = true;
  guardar();
}

function mostrarPasoTour() {
  const paso = TOUR_PASOS[tourIndex];
  if (paso.view) switchView(paso.view);
  const el = document.querySelector(paso.sel);
  if (!el) { siguientePaso(); return; }

  const mostrar = () => {
    const rect = el.getBoundingClientRect();
    destacarElemento(rect);
    $('#tourStepCount').textContent = `Paso ${tourIndex + 1} de ${TOUR_PASOS.length}`;
    $('#tourTitulo').textContent = paso.titulo;
    $('#tourTexto').textContent = paso.texto;
    $('#tourAnterior').disabled = tourIndex === 0;
    $('#tourSiguiente').textContent = tourIndex === TOUR_PASOS.length - 1 ? 'Finalizar' : 'Siguiente';
    posicionarTooltip(rect);
  };
  // Pequeña espera si cambiamos de vista (por la animación)
  setTimeout(mostrar, paso.view ? 150 : 20);
}

function destacarElemento(rect) {
  const pad = 6;
  const h = $('#tourHighlight');
  h.style.top = (rect.top - pad) + 'px';
  h.style.left = (rect.left - pad) + 'px';
  h.style.width = (rect.width + pad * 2) + 'px';
  h.style.height = (rect.height + pad * 2) + 'px';
}

function posicionarTooltip(rect) {
  const tooltip = $('#tourTooltip');
  const tw = tooltip.offsetWidth || 320;
  const th = tooltip.offsetHeight || 180;
  const margin = 16;
  let left, top;

  if (rect.right + margin + tw <= window.innerWidth) {
    left = rect.right + margin;
    top = rect.top + rect.height / 2 - th / 2;
  } else if (rect.left - margin - tw >= 0) {
    left = rect.left - margin - tw;
    top = rect.top + rect.height / 2 - th / 2;
  } else {
    left = rect.left + rect.width / 2 - tw / 2;
    top = rect.bottom + margin;
  }

  top = Math.max(12, Math.min(top, window.innerHeight - th - 12));
  left = Math.max(12, Math.min(left, window.innerWidth - tw - 12));
  tooltip.style.left = left + 'px';
  tooltip.style.top = top + 'px';
}

function anteriorPaso() {
  if (tourIndex > 0) { tourIndex--; mostrarPasoTour(); }
}

function siguientePaso() {
  if (tourIndex < TOUR_PASOS.length - 1) {
    tourIndex++;
    mostrarPasoTour();
  } else {
    cerrarTour();
    toast('¡Listo! Ya podés empezar a usar tu gestión de stock');
  }
}

/* ============================================================
   EVENTOS
   ============================================================ */
function bindEvents() {
  // Navegación
  $$('.nav-item').forEach(btn => {
    btn.addEventListener('click', () => switchView(btn.dataset.view));
  });

  // Panel entrada
  $('#btnCerrarEntrada').addEventListener('click', cerrarPanelEntrada);
  $('#overlayEntrada').addEventListener('click', cerrarPanelEntrada);
  $('#tabExistente').addEventListener('click', () => switchTabEntrada('existente'));
  $('#tabNuevo').addEventListener('click', () => switchTabEntrada('nuevo'));
  $('#chkMarcaNueva').addEventListener('change', (e) => {
    $('#inputMarcaNuevaWrap').hidden = !e.target.checked;
  });
  $('#btnAgregarLista').addEventListener('click', agregarDesdeExistente);
  $('#btnCrearProducto').addEventListener('click', crearProductoYAgregar);
  $('#btnConfirmarEntrada').addEventListener('click', confirmarEntrada);

  // Stock
  $('#buscarProducto').addEventListener('input', renderStock);
  $('#filtroMarcaStock').addEventListener('change', renderStock);
  $('#filtroEstado').addEventListener('change', renderStock);
  $('#btnAgregarProducto').addEventListener('click', () => abrirModalProducto());

  // Producto modal
  $('#btnGuardarProducto').addEventListener('click', guardarProducto);

  // Ventas
  $('#buscarVenta').addEventListener('input', renderVentas);
  $('#btnNuevaVenta').addEventListener('click', abrirModalVenta);
  $('#btnAgregarItemVenta').addEventListener('click', agregarItemVenta);
  $('#btnGuardarVenta').addEventListener('click', guardarVenta);

  // Gastos
  $('#filtroCategoriaGasto').addEventListener('change', renderGastos);
  $('#btnNuevoGasto').addEventListener('click', abrirModalGasto);
  $('#btnGuardarGasto').addEventListener('click', guardarGasto);

  // Clientes
  $('#buscarCliente').addEventListener('input', renderClientes);
  $('#btnNuevoCliente').addEventListener('click', () => abrirModalCliente());
  $('#btnGuardarCliente').addEventListener('click', guardarCliente);

  // Config
  $('#btnGuardarConfig').addEventListener('click', guardarConfig);
  $('#cfgBgColor').addEventListener('input', (e) => {
    $('#cfgBgColorText').value = e.target.value;
    document.documentElement.style.setProperty('--bg', e.target.value);
  });
  $('#cfgBgColorText').addEventListener('change', (e) => {
    let v = e.target.value.trim();
    if (!v.startsWith('#')) v = '#' + v;
    if (/^#[0-9a-fA-F]{6}$/.test(v)) {
      $('#cfgBgColor').value = v;
      document.documentElement.style.setProperty('--bg', v);
    }
  });
  $('#cfgSidebarColor').addEventListener('input', (e) => {
    $('#cfgSidebarColorText').value = e.target.value;
    document.documentElement.style.setProperty('--sidebar-bg', e.target.value);
  });
  $('#cfgSidebarColorText').addEventListener('change', (e) => {
    let v = e.target.value.trim();
    if (!v.startsWith('#')) v = '#' + v;
    if (/^#[0-9a-fA-F]{6}$/.test(v)) {
      $('#cfgSidebarColor').value = v;
      document.documentElement.style.setProperty('--sidebar-bg', v);
    }
  });
  $('#cfgLogo').addEventListener('change', (e) => {
    if (e.target.files[0]) cargarLogo(e.target.files[0]);
    e.target.value = '';
  });
  $('#btnAgregarMarca').addEventListener('click', agregarMarca);
  $('#nuevaMarca').addEventListener('keydown', (e) => { if (e.key === 'Enter') agregarMarca(); });
  $('#btnReset').addEventListener('click', resetearTodo);

  // Exportar / Importar
  $('#btnExport').addEventListener('click', exportarDatos);
  $('#btnImport').addEventListener('click', () => $('#importFile').click());
  $('#importFile').addEventListener('change', (e) => {
    if (e.target.files[0]) importarDatos(e.target.files[0]);
    e.target.value = '';
  });

  // Bienvenida
  $('#btnBienvenidaSaltar').addEventListener('click', () => confirmarBienvenida(false));
  $('#btnBienvenidaTutorial').addEventListener('click', () => confirmarBienvenida(true));

  // Tutorial
  $('#btnTutorial').addEventListener('click', iniciarTour);
  $('#tourSaltar').addEventListener('click', cerrarTour);
  $('#tourAnterior').addEventListener('click', anteriorPaso);
  $('#tourSiguiente').addEventListener('click', siguientePaso);
  window.addEventListener('resize', () => {
    if (!$('#tourOverlay').hidden) {
      const paso = TOUR_PASOS[tourIndex];
      const el = document.querySelector(paso.sel);
      if (el) {
        const r = el.getBoundingClientRect();
        destacarElemento(r);
        posicionarTooltip(r);
      }
    }
  });

  // Cerrar modales
  $$('[data-close]').forEach(btn => {
    btn.addEventListener('click', () => cerrarModales());
  });
  $$('.overlay').forEach(o => {
    if (o.id !== 'overlayEntrada') {
      o.addEventListener('click', cerrarModales);
    }
  });

  // Atajos de teclado
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (!$('#tourOverlay').hidden) { cerrarTour(); return; }
      cerrarModales();
      cerrarPanelEntrada();
    }
    if (!$('#tourOverlay').hidden) {
      if (e.key === 'ArrowRight') { e.preventDefault(); siguientePaso(); }
      if (e.key === 'ArrowLeft') { e.preventDefault(); anteriorPaso(); }
    }
  });
}

/* ============================================================
   INIT
   ============================================================ */
const esPrimeraVez = !localStorage.getItem(STORAGE_KEY);
seed();
cargar();
renderTodo();
bindEvents();

if (esPrimeraVez || !state.config.tutorialVisto) {
  abrirBienvenida(esPrimeraVez);
}
