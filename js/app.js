
        window.onerror = function(msg, url, line, col, error) {
            alert("🔥 ERROR DE EJECUCIÓN detectado:\n" + msg + "\nFila: " + line);
            console.error(error);
            return false;
        };

        let matrizOriginal =[]; 
        let nombresMapping =[];
        window.dataIA =[]; 
        let excelData =[];
        let headers =[];
        let chartDiario = null;
        let chartVentas = null;
        let horasEjeX =[8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21];
        let rolesEspecialesData = {};
        let turnosDisponibles = [];
        let mesFiltrando = null; // null = todos | número YYYYMM para filtrar un mes
        let festivosCache = {};
        let hayColumnaNomina = false; // true si la Matriz Turnos trae la columna "Código Nómina"

        // ================== FESTIVOS COLOMBIANOS ==================
        function calcEaster(y) {
            let a=y%19, b=Math.floor(y/100), c=y%100, d=Math.floor(b/4), e=b%4;
            let f=Math.floor((b+8)/25), g=Math.floor((b-f+1)/3);
            let h=(19*a+b-d-g+15)%30, ii=Math.floor(c/4), k=c%4;
            let l=(32+2*e+2*ii-h-k)%7, m=Math.floor((a+11*h+22*l)/451);
            let month=Math.floor((h+l-7*m+114)/31), day=((h+l-7*m+114)%31)+1;
            return new Date(y, month-1, day);
        }

        function calcularFestivosColombiaYear(y) {
            let set = new Set();
            function fmt(d) { return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
            function fixed(m, d)    { set.add(fmt(new Date(y, m-1, d))); }
            function emiliani(m, d) {
                let dt = new Date(y, m-1, d);
                let dow = dt.getDay();
                if (dow !== 1) dt.setDate(dt.getDate() + (8 - dow) % 7);
                set.add(fmt(dt));
            }
            function easterOffset(days, useEmiliani) {
                let dt = new Date(calcEaster(y));
                dt.setDate(dt.getDate() + days);
                if (useEmiliani) { let dow = dt.getDay(); if (dow !== 1) dt.setDate(dt.getDate() + (8-dow)%7); }
                set.add(fmt(dt));
            }
            // Fijos
            fixed(1,1); fixed(5,1); fixed(7,20); fixed(8,7); fixed(12,8); fixed(12,25);
            // Ley Emiliani
            emiliani(1,6); emiliani(3,19); emiliani(6,29); emiliani(8,15);
            emiliani(10,12); emiliani(11,1); emiliani(11,11);
            // Semana Santa y pascua
            easterOffset(-3, false); easterOffset(-2, false);
            easterOffset(39, true); easterOffset(60, true); easterOffset(68, true);
            return set;
        }

        function esFestivoColombia(date) {
            let y = date.getFullYear();
            if (!festivosCache[y]) festivosCache[y] = calcularFestivosColombiaYear(y);
            let key = `${y}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
            return festivosCache[y].has(key);
        }

        function getTipoDia(date) {
            if (esFestivoColombia(date)) return 'festivo';
            let dow = date.getDay();
            if (dow === 0) return 'domingo';
            if (dow === 6) return 'sabado';
            return null;
        }
        
        let relevosGlobales = [];
        let personalGlobal =[];
        let rolesHoyGlobal = {};

        window.onload = function() {
            let hoy = new Date();
            document.getElementById('fechaInicio').value = hoy.toISOString().split('T')[0];
            
            if (typeof XLSX === 'undefined') {
                document.getElementById('redAlert').style.display = 'block';
                mostrarMensaje("❌ El sistema está bloqueado por tu red. No podrá procesar archivos.", true);
            }
        };

        function switchTab(tabId, btn) {
            document.querySelectorAll('.view-section').forEach(el => el.classList.remove('active'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
            document.getElementById(tabId).classList.add('active');
            if(btn) btn.classList.add('active');
        }

        // Sistema para mostrar mensajes de forma SEGURA sin destruir el DOM
        function mostrarMensaje(htmlText, isSuccess = false) {
            let msgEl = document.getElementById('msgCentral');
            if (msgEl) {
                msgEl.innerHTML = htmlText;
                if (isSuccess) {
                    msgEl.style.border = '1px solid #27ae60';
                    msgEl.style.background = '#e8f5e9';
                    msgEl.style.color = '#2c3e50';
                } else {
                    msgEl.style.border = '1px dashed #bdc3c7';
                    msgEl.style.background = '#f8f9fa';
                    msgEl.style.color = '#7f8c8d';
                }
            }
        }

        function getRoleInfo(nombre, rolRealOriginal) {
            let match = excelData.find(e => e.data[0] === nombre);
            let n = (nombre + ' ' + (match ? match.rolGenerico : '')).toUpperCase();
            
            if (n.includes('CAMBISTA')) return { class: 'rol-tag-CAMBISTA', icon: '💵 CAM', color: '#d35400', label: 'CAMBISTA' };
            if (n.includes('VISADO')) return { class: 'rol-tag-VISADO', icon: '🔍 VIS', color: '#e65100', label: 'VISADO' };
            if (n.includes('EMERGENTE')) return { class: 'rol-tag-EMERGENTE', icon: '🚨 EME', color: '#c0392b', label: 'EMERGENTE' };
            
            if (rolRealOriginal === 'CAJERO') return { class: 'rol-tag-CAJERO', icon: '💳 CAJ', color: '#2e7d32', label: 'CAJERO' };
            if (rolRealOriginal === 'EMPACADOR') return { class: 'rol-tag-EMPACADOR', icon: '🛍️ EMP', color: '#1565c0', label: 'EMPACADOR' };
            
            return { class: `rol-tag-${rolRealOriginal}`, icon: `👤 ${rolRealOriginal.substring(0,3)}`, color: '', label: rolRealOriginal };
        }

        // Lee la fecha real de inicio desde la celda B1 del Excel de carga (Matriz Turnos)
        function leerFechaInicioDesdeB1(worksheet) {
            let cell = worksheet ? worksheet['B1'] : null;
            if (!cell) return null;
            let d = null;
            if (cell.t === 'd' && cell.v instanceof Date) {
                d = cell.v;
            } else if (typeof cell.v === 'number') {
                // Fecha serial de Excel -> fecha JS
                d = new Date(Math.round((cell.v - 25569) * 86400 * 1000));
                d = new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
            } else {
                let s = (cell.w || cell.v || '').toString().trim();
                let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
                if (m) {
                    d = new Date(+m[1], +m[2] - 1, +m[3]);
                } else {
                    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
                    if (m) {
                        let dd = +m[1], mm = +m[2], yy = +m[3];
                        if (yy < 100) yy += 2000;
                        d = new Date(yy, mm - 1, dd);
                    }
                }
            }
            if (!d || isNaN(d.getTime())) return null;
            return d;
        }

        document.getElementById('fileInput').addEventListener('change', function(e) {
            if (typeof XLSX === 'undefined') return;
            let file = e.target.files[0];
            if (!file) return;

            mostrarMensaje("⌛ Cargando y procesando matriz...");
            let reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    let data = new Uint8Array(evt.target.result);
                    let workbook = XLSX.read(data, {type: 'array'});
                    let worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    let json = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false, defval: ""});

                    matrizOriginal = json.filter(row => row && row.length > 0 && row.some(c => c && c.toString().trim() !== '') && row[0] !== 'PARAMETROS_SISTEMA');
                    if (matrizOriginal.length < 2) { alert("❌ Excel vacío o inválido."); return; }

                    let fechaB1 = leerFechaInicioDesdeB1(worksheet);
                    if (fechaB1) {
                        let iso = fechaB1.getFullYear() + '-' + String(fechaB1.getMonth() + 1).padStart(2, '0') + '-' + String(fechaB1.getDate()).padStart(2, '0');
                        document.getElementById('fechaInicio').value = iso;
                    }

                    combinarDatosSeguro();
                    mostrarMensaje("✅ Matriz de Turnos procesada. Puedes ir a la pestaña 'Operación Diaria'.", true);
                } catch (err) { alert("❌ Error analizando matriz: " + err.message); }
            };
            reader.readAsArrayBuffer(file);
        });

        document.getElementById('fileMapping').addEventListener('change', function(e) {
            if (typeof XLSX === 'undefined') return;
            let file = e.target.files[0];
            if (!file) return;
            let reader = new FileReader();
            reader.onload = function(evt) {
                let data = new Uint8Array(evt.target.result);
                let workbook = XLSX.read(data, {type: 'array'});
                nombresMapping = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {raw: false});
                if (matrizOriginal.length > 0) combinarDatosSeguro();
            };
            reader.readAsArrayBuffer(file);
        });

        // ================== CARGA DE IA CON DIAGNÓSTICO PROFUNDO ==================
        document.getElementById('fileAnalisis').addEventListener('change', function(e) {
            if (typeof XLSX === 'undefined') return;
            let file = e.target.files[0];
            if (!file) return;

            mostrarMensaje("⌛ Analizando Archivo de Inteligencia Artificial...");

            let reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    let data = new Uint8Array(evt.target.result);
                    let workbook = XLSX.read(data, {type: 'array'});
                    
                    let sheetName = workbook.SheetNames.find(s => {
                        let cleanName = s.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                        return cleanName.includes('analisis detallado') || cleanName.includes('analisis');
                    }) || workbook.SheetNames[0];

                    window.dataIA = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {raw: false, defval: ""});
                    
                    if (window.dataIA.length > 0) {
                        let firstRow = window.dataIA[0];
                        let kF = Object.keys(firstRow).find(k => k.toLowerCase().includes('fecha'));
                        let kH = Object.keys(firstRow).find(k => k.toLowerCase().includes('hora'));
                        let k7 = Object.keys(firstRow).find(k => {
                            let clean = k.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                            return clean.includes('7. accion') || clean.includes('emergentes');
                        });

                        // 🔍 Mensaje de Diagnóstico de Éxito en Pantalla
                        let info = `<span style="color:#27ae60; font-size: 16px;">✅ Inteligencia Artificial Cargada y Enlazada.</span><br><br>`;
                        info += `<span style="font-size: 13px; font-weight: normal; color: #34495e; display: inline-block; text-align: left;">`;
                        info += `<b>🔎 Diagnóstico Interno de Cruce:</b><br>`;
                        info += kF ? `✔️ Etiqueta Fecha: <b>OK</b> ("${kF}" | Ej: ${firstRow[kF]})<br>` : `❌ Columna Fecha NO ENCONTRADA.<br>`;
                        info += kH ? `✔️ Etiqueta Hora: <b>OK</b> ("${kH}" | Ej: ${firstRow[kH]})<br>` : `❌ Columna Hora NO ENCONTRADA.<br>`;
                        info += k7 ? `✔️ Etiqueta Acción Predictiva: <b>OK</b> ("${k7}")<br>` : `❌ Columna Emergentes NO ENCONTRADA.<br>`;
                        info += `</span>`;

                        mostrarMensaje(info, true);

                        if (matrizOriginal.length > 0 && excelData.length > 0) procesarDiaEspecificoSeguro();
                        else if (excelData.length === 0) renderVentasChart();
                    } else {
                        mostrarMensaje("❌ El archivo IA parece estar vacío o en un formato que no se puede leer.");
                    }
                } catch (err) {
                    alert("❌ Error leyendo el archivo de Análisis IA: " + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        });

        function generarFiltroMeses() {
            let container = document.getElementById('filtroMesesContainer');
            if (!container) return;
            if (!headers || headers.length === 0) { container.style.display = 'none'; return; }

            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());

            // Detectar meses únicos presentes en la programación (en orden)
            let mesesOrden = [];
            let mesesVistos = new Set();
            headers.forEach((h, idx) => {
                if (h === 'Total Horas Periodo') return;
                let d = new Date(startDate);
                d.setDate(d.getDate() + idx);
                let key = d.getFullYear() * 100 + d.getMonth();
                if (!mesesVistos.has(key)) {
                    mesesVistos.add(key);
                    let label = d.toLocaleDateString('es-CO', { month: 'short' });
                    label = label.charAt(0).toUpperCase() + label.slice(1).replace('.', '');
                    // Si hay más de un año en la programación, añadir el año
                    mesesOrden.push({ key, label, anio: d.getFullYear() });
                }
            });

            if (mesesOrden.length === 0) { container.style.display = 'none'; return; }
            let aniosDistintos = new Set(mesesOrden.map(m => m.anio)).size > 1;

            let html = '<div style="display:flex; gap:7px; flex-wrap:wrap; align-items:center;">';
            if (mesesOrden.length > 1) {   // filtrar por mes solo si hay más de uno
                html += `<span style="font-size:12px; font-weight:700; color:#555; margin-right:4px;">📅 Mes:</span>`;
                html += `<button class="mes-btn${mesFiltrando === null ? ' activo' : ''}" onclick="filtrarPorMes(null)">Todos</button>`;
                mesesOrden.forEach(m => {
                    let etiqueta = aniosDistintos ? `${m.label} ${m.anio}` : m.label;
                    html += `<button class="mes-btn${mesFiltrando === m.key ? ' activo' : ''}" onclick="filtrarPorMes(${m.key})">${etiqueta}</button>`;
                });
            }
            // Botón IMPRIMIR: imprime la matriz del mes filtrado (o todo si 'Todos')
            let etqImp = mesFiltrando === null
                ? (mesesOrden.length > 1 ? 'Imprimir todo' : 'Imprimir')
                : 'Imprimir mes';
            html += `<button class="mes-btn" style="background:#2980b9; color:#fff; font-weight:700; margin-left:8px;" onclick="imprimirMatrizMes()">🖨️ ${etqImp}</button>`;
            html += '</div>';

            container.innerHTML = html;
            container.style.display = 'block';
        }

        function filtrarPorMes(key) {
            mesFiltrando = key;
            generarFiltroMeses();   // actualiza estado activo de botones
            aplicarFiltrosGlobales();
        }

        // IMPRIMIR la Matriz de Turnos del mes filtrado (o todo) con diseño moderno:
        // una tabla por SEMANA, turnos en cápsulas del mismo color de la matriz.
        window.imprimirMatrizMes = function () {
            if (!excelData || !excelData.length) { alert('Carga primero la Matriz de Turnos.'); return; }
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            let cols = [];
            headers.forEach((h, idx) => {
                if (h === 'Total Horas Periodo') return;
                let d = new Date(startDate); d.setDate(d.getDate() + idx);
                let key = d.getFullYear() * 100 + d.getMonth();
                if (mesFiltrando === null || key === mesFiltrando) cols.push({ colIdx: idx + 1, date: d });
            });
            if (!cols.length) { alert('No hay días para imprimir.'); return; }
            let rolF = (document.getElementById('rolSelector') || {}).value || 'TODOS';
            let personas = excelData.filter(it => rolF === 'TODOS' ? true : it.rolReal === rolF);

            // UNA sola tabla con TODAS las columnas del período (mes o año).
            // Ancho 100% + table-layout:fixed => siempre cabe en la página; las
            // columnas quedan UNIFORMES y la fuente se reduce según cuántas haya.
            const INI = ['D', 'L', 'M', 'X', 'J', 'V', 'S'];   // iniciales de día
            const nCols = cols.length;
            let personPct = nCols <= 31 ? 12 : (nCols <= 90 ? 8 : 6);
            let dayPct = (100 - personPct) / nCols;
            // Tamaños de letra en 'vw' (relativos al ancho de la página): el TURNO
            // entra SIEMPRE en UNA sola línea sin recortarse, sin importar el
            // tamaño del papel; al imprimir grande la letra crece proporcional y
            // se mantiene nítida (texto vectorial).
            let pillVw = (dayPct * 0.118).toFixed(3);   // cabe ~13 caracteres en 1 línea
            let letVw = (dayPct * 0.125).toFixed(3);    // inicial del día
            let numVw = (dayPct * 0.175).toFixed(3);    // número del día
            let nomVw = (personPct * 0.052).toFixed(3); // nombre (puede envolver)

            let mesTxt = cols[Math.floor(cols.length / 2)].date.toLocaleDateString('es-CO', { month: 'long', year: 'numeric' });
            mesTxt = mesTxt.charAt(0).toUpperCase() + mesTxt.slice(1);
            if (nCols > 45) {   // varios meses -> mostrar rango
                let dIni = cols[0].date, dFin = cols[cols.length - 1].date;
                mesTxt = dIni.toLocaleDateString('es-CO', { month: 'short', year: 'numeric' }) + ' — ' + dFin.toLocaleDateString('es-CO', { month: 'short', year: 'numeric' });
            }
            let tienda = (window.APP_TIENDA || '').toUpperCase();
            const esc = s => (s || '').toString().replace(/&/g, '&amp;').replace(/</g, '&lt;');

            let colgroup = `<colgroup><col style="width:${personPct}%">` + `<col style="width:${dayPct}%">`.repeat(nCols) + '</colgroup>';
            let thDias = '';
            let mesAnt = -1;
            cols.forEach(c => {
                let dow = c.date.getDay();
                let fin = dow === 0 || dow === 6;
                // Separador sutil al cambiar de mes (línea izquierda)
                let sep = (c.date.getMonth() !== mesAnt) ? ' sepmes' : '';
                mesAnt = c.date.getMonth();
                thDias += `<th class="${fin ? 'finde' : ''}${sep}">${INI[dow]}<span class="dnum">${c.date.getDate()}</span></th>`;
            });
            let cuerpo = `<table>${colgroup}<thead><tr><th class="pcol">Rol / Persona</th>${thDias}</tr></thead><tbody>`;
            personas.forEach(it => {
                let colNom = (typeof getRoleInfo === 'function' ? (getRoleInfo(it.data[0], it.rolReal) || {}).color : '') || '#2c3e50';
                cuerpo += `<tr><td class="pcol"><span class="pnom" style="color:${colNom};">${esc(it.data[0])}</span></td>`;
                let mA = -1;
                cols.forEach(c => {
                    let sep = (c.date.getMonth() !== mA) ? ' sepmes' : ''; mA = c.date.getMonth();
                    let v = (it.data[c.colIdx] || '').toString().trim();
                    let st = v ? estiloTurno(v.toUpperCase()) : null;
                    if (v && st) cuerpo += `<td class="${sep}"><span class="pill" style="background:${st.bg}; color:${st.color};">${esc(v)}</span></td>`;
                    else if (v && v.toUpperCase() !== 'SIN TURNO') cuerpo += `<td class="${sep}"><span class="pill" style="background:#eceff1; color:#607d8b;">${esc(v)}</span></td>`;
                    else cuerpo += `<td class="vacio${sep}">·</td>`;
                });
                cuerpo += '</tr>';
            });
            cuerpo += '</tbody></table>';

            let leyenda = `<div class="leyenda">
                <span><i style="background:#ffe0b2"></i>Apertura</span>
                <span><i style="background:#dcedc8"></i>Intermedio</span>
                <span><i style="background:#d1c4e9"></i>Cierre</span>
                <span><i style="background:#ffcdd2"></i>Libre</span>
                <span><i style="background:#dcdcdc"></i>Comp</span>
                <span><i style="background:#f8bbd0"></i>Vacaciones</span></div>`;

            let doc = `<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Matriz ${esc(tienda)} ${esc(mesTxt)}</title>
            <style>
              @page { size: landscape; margin: 6mm; }
              * { box-sizing:border-box; }
              body { font-family:'Segoe UI',Arial,sans-serif; color:#2c3e50; margin:0; padding:10px; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
              .cab { display:flex; justify-content:space-between; align-items:flex-end; border-bottom:2px solid #2c3e50; padding-bottom:5px; margin-bottom:8px; }
              .cab h1 { margin:0; font-size:16px; letter-spacing:.3px; }
              .cab .mes { font-size:12px; color:#7f8c8d; font-weight:700; }
              table { width:100%; border-collapse:separate; border-spacing:1px 2px; table-layout:fixed; }
              th { font-size:clamp(3px, ${letVw}vw, 12px); color:#fff; background:#34495e; padding:2px 1px; text-align:center; border-radius:3px; line-height:1.05; font-weight:600; }
              th.finde { background:#5d6d7e; }
              th.pcol { text-align:left; background:#2c3e50; }
              th.sepmes, td.sepmes { border-left:2px solid #b0bec5; }
              .dnum { display:block; font-size:clamp(3.5px, ${numVw}vw, 15px); font-weight:600; }
              td { padding:0; text-align:center; }
              td.pcol { text-align:left; background:#f4f6f8; border-radius:4px; padding:1px 5px; }
              .pnom { display:block; font-size:clamp(3.5px, ${nomVw}vw, 12px); font-weight:400; line-height:1.05; }
              .pill { display:block; padding:1px 1px; border-radius:4px; font-size:clamp(2.5px, ${pillVw}vw, 12px); font-weight:400; line-height:1.15; white-space:nowrap; overflow:hidden; text-overflow:clip; }
              td.vacio { color:#cfd8dc; font-size:clamp(2.5px, ${pillVw}vw, 12px); }
              .leyenda { display:flex; gap:14px; flex-wrap:wrap; margin-top:10px; font-size:10px; color:#555; }
              .leyenda span { display:inline-flex; align-items:center; gap:5px; }
              .leyenda i { width:12px; height:12px; border-radius:3px; display:inline-block; }
              .barra { text-align:center; margin-top:16px; }
              @media print { .barra { display:none; } }
            </style></head><body>
            <div class="cab"><h1>${tienda ? esc(tienda) + ' · ' : ''}Programación de Turnos</h1><div class="mes">${esc(mesTxt)}${rolF !== 'TODOS' ? ' · ' + esc(rolF) : ''}</div></div>
            ${cuerpo}${leyenda}
            <div class="barra"><button onclick="window.print()" style="padding:10px 26px; font-size:15px; background:#2980b9; color:#fff; border:none; border-radius:6px; cursor:pointer;">🖨️ Imprimir</button></div>
            <scr` + `ipt>window.onload=function(){setTimeout(function(){try{window.print();}catch(e){}},350);};</scr` + `ipt>
            </body></html>`;

            let w = window.open('', '_blank');
            if (!w) { alert('Permite las ventanas emergentes para poder imprimir.'); return; }
            w.document.open(); w.document.write(doc); w.document.close();
        };

        // Quita el consecutivo final del rol, ej. "Cajero 2" -> "Cajero"
        function limpiarNumeroRol(s) {
            return (s || '').toString().replace(/\s*\d+\s*$/, '').trim();
        }

        function normalizarTexto(s) {
            return (s || '').toString().trim().toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
        }

        function combinarDatosSeguro() {
            if (!matrizOriginal || matrizOriginal.length < 2) return;
            excelData =[];

            // Detecta las columnas "Rol", "Nombre" y "Código Nómina" (formato V26.PY actual).
            // Si no existen por separado, cae al formato antiguo: columna 0 = "Rol - Nombre" combinado.
            let headerRow = matrizOriginal[0];
            let idxRol    = headerRow.findIndex(h => normalizarTexto(h) === 'ROL');
            let idxNombre = headerRow.findIndex(h => normalizarTexto(h) === 'NOMBRE');
            let idxNomina = headerRow.findIndex((h, i) => i > 0 && normalizarTexto(h).includes('NOMINA'));
            let idxRestric = headerRow.findIndex((h, i) => i > 0 && normalizarTexto(h).includes('RESTRIC'));
            hayColumnaNomina = idxNomina >= 0;
            let hayColumnasSeparadas = idxRol >= 0 && idxNombre >= 0;

            let idxExcluir = new Set([idxNomina, idxRestric]);
            idxExcluir.delete(-1);
            if (hayColumnasSeparadas) { idxExcluir.add(idxRol); idxExcluir.add(idxNombre); }
            else { idxExcluir.add(0); } // formato antiguo: columna 0 = "Rol y Empleado" combinado
            headers = headerRow.filter((h, i) => !idxExcluir.has(i));

            for (let r = 1; r < matrizOriginal.length; r++) {
                let rowFull = Array.from(matrizOriginal[r]);
                let codigoNomina = hayColumnaNomina ? (rowFull[idxNomina] || '').toString().trim() : '';
                let restriccion = idxRestric >= 0 ? (rowFull[idxRestric] || '').toString().trim() : '';
                let excluida = restriccion.trim().toUpperCase() === 'X';   // 'X' = no suma en su rol

                let rolTexto, nombreTexto;
                if (hayColumnasSeparadas) {
                    rolTexto = (rowFull[idxRol] || '').toString().trim();
                    nombreTexto = (rowFull[idxNombre] || '').toString().trim();
                } else {
                    let combinado = (rowFull[0] || '').toString();
                    let partes = combinado.split(' - ');
                    rolTexto = partes.length > 1 ? partes[0].trim() : '';
                    nombreTexto = partes.length > 1 ? partes.slice(1).join(' - ').trim() : combinado.trim();
                }
                rolTexto = limpiarNumeroRol(rolTexto);

                // Reconstruye la fila con la misma forma de siempre: data[0] = nombre, data[1..] = turnos
                let row = rowFull.filter((_, i) => !idxExcluir.has(i));
                row.unshift(nombreTexto);

                let rolOriginalText = rolTexto.toUpperCase();
                let isEmergente = rolOriginalText.includes('EMERGENTE');
                let rolReal = (rolOriginalText.includes('EMPACADOR') || rolOriginalText.includes('SADOFE') || rolOriginalText.includes('FUNDACIÓN') || rolOriginalText.includes('FUNDACION') || isEmergente) ? 'EMPACADOR' : 'CAJERO';

                if (nombresMapping && nombresMapping.length > 0) {
                    let m = nombresMapping.find(map => (map.ROL || '').toString().trim().toUpperCase() === rolOriginalText);
                    if (m && m.NOMBRE) row[0] = m.NOMBRE;
                }
                excelData.push({ data: row, rolGenerico: rolOriginalText, rolReal: rolReal, rol: rolTexto,
                                 codigoNomina: codigoNomina, restriccion: restriccion, excluida: excluida });
            }

            // Recolectar todos los turnos únicos para los desplegables
            let turnosSet = new Set(['', 'VC', 'LIC', 'INC']);
            excelData.forEach(item => {
                item.data.slice(1).forEach(val => {
                    let v = (val || '').toString().trim().toUpperCase();
                    turnosSet.add(v);
                });
            });
            turnosDisponibles = Array.from(turnosSet).sort((a, b) => {
                if (a === '') return -1;
                if (b === '') return 1;
                let rA = a.match(/^(\d+)/), rB = b.match(/^(\d+)/);
                if (rA && rB) return parseInt(rA[1]) - parseInt(rB[1]);
                if (rA) return -1;
                if (rB) return 1;
                return a.localeCompare(b);
            });

            mesFiltrando = null; // resetear filtro al cargar nueva programación
            document.getElementById('tabsContainer').style.display = 'flex';
            generarFiltroMeses();
            llenarSelectorDias();
            llenarSelectorDiasCobertura();
            aplicarFiltrosGlobales();
            renderLineaTiempoCobertura();
            if (typeof renderCumplimiento === 'function') renderCumplimiento();
            poblarFiltroMesEquidadAux();
            renderEquidadAuxiliar();
            poblarFiltroMesCinta();
            poblarSelectorPersonaCinta();
            renderCintaTurnos();
        }

        function aplicarFiltrosGlobales() {
            let f = document.getElementById('rolSelector').value;
            let filtrados = excelData.filter(item => f === 'TODOS' ? true : item.rolReal === f);
            generarTablaMensual(filtrados);
        }

        // ============= COLOR POR GRUPO DE TURNO (A=Apertura naranja · I=Intermedio verde · C=Cierre azul/morado) =============
        const PALETA_APERTURA   = { '6.5': '#fff3e0', '7': '#ffe0b2', '8': '#ffcc80' }; // 3 tonos naranja pastel (claros)
        const PALETA_INTERMEDIO = { '6.5': '#f1f8e9', '7': '#dcedc8', '8': '#c5e1a5' }; // 3 tonos verde (claros)
        const PALETA_CIERRE     = { '6.5': '#e3f2fd', '7': '#d1c4e9', '8': '#e1bee7' }; // azul -> morado (claros)

        function claveHorasTurno(horasStr) {
            let n = parseFloat((horasStr || '').toString().replace(',', '.'));
            if (isNaN(n)) return '7';
            if (n <= 6.75) return '6.5';
            if (n <= 7.5) return '7';
            return '8';
        }

        function estiloTurno(val) {
            if (!val) return null;
            if (val === 'COMP') return { bg: '#dcdcdc', color: '#424242' };
            if (val === 'LIBRE' || val === 'LBRE') return { bg: '#ffcdd2', color: '#b71c1c' };
            if (val === 'VC') return { bg: '#f8bbd0', color: '#ad1457', bold: true };
            let m = val.match(/^(\d+(?:[.,]\d+)?)([AIC])/);
            if (!m) return null;
            let paleta = m[2] === 'A' ? PALETA_APERTURA : (m[2] === 'I' ? PALETA_INTERMEDIO : PALETA_CIERRE);
            return { bg: paleta[claveHorasTurno(m[1])], color: '#2c3e50' };
        }

        // Extrae las horas (con decimales tipo 6,5) que representa un código de turno para la Σ semanal
        function horasDeTurno(val) {
            if (!val) return 0;
            let m = val.match(/^(\d+(?:[.,]\d+)?)/);
            if (!m) return 0;
            return parseFloat(m[1].replace(',', '.')) || 0;
        }

        function formatHorasSemanal(n) {
            let r = Math.round(n * 10) / 10;
            let txt = Number.isInteger(r) ? r.toString() : r.toFixed(1).replace('.', ',');
            return txt + 'h';
        }

        function generarTablaMensual(datos) {
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());

            // Precomputar visibilidad y tipo de día por columna
            let colMeta = headers.map((h, index) => {
                if (h === 'Total Horas Periodo') return { visible: false, dayType: null };
                let d = new Date(startDate);
                d.setDate(d.getDate() + index);
                let key = d.getFullYear() * 100 + d.getMonth();
                let visible = mesFiltrando === null || key === mesFiltrando;
                return { visible, dayType: getTipoDia(d), date: d };
            });

            // Agrupar columnas por semana (Lunes-Domingo) y ubicar la última columna VISIBLE de
            // cada semana: ahí se pinta la Σ SEMANAL, así nunca desaparece al filtrar por mes o rol.
            let colWeekKey = new Array(headers.length + 1);
            let weekAnchor = {};
            colMeta.forEach((meta, idx) => {
                if (!meta.date) return;
                let dow = meta.date.getDay();
                let wEnd = new Date(meta.date);
                wEnd.setDate(wEnd.getDate() + (7 - dow) % 7);
                let wk = wEnd.toISOString().slice(0, 10);
                let colIdx = idx + 1;
                colWeekKey[colIdx] = wk;
                if (meta.visible) weekAnchor[wk] = colIdx;
            });

            // Colores de fines de semana/festivo en encabezados (tonos claros, no oscuros)
            let thColors = { sabado: '#5dade2', domingo: '#2e6da4', festivo: '#e67e22' };

            // Columnas fijas de identificación (sticky), en este orden: Rol / Nombre / Código Nómina
            let ROL_COL_W = 130, NOMBRE_COL_W = 190, NOMINA_COL_W = 90;
            let nombreLeft = ROL_COL_W;
            let nominaLeft = ROL_COL_W + NOMBRE_COL_W;

            let leyendaEl = document.getElementById('leyendaDiasContainer');
            if (leyendaEl) leyendaEl.innerHTML = `<div class="leyenda-dias">
                <strong style="color:#555;">Referencia:</strong>
                <span style="background:#5dade2;">Sábado</span>
                <span style="background:#2e6da4;">Domingo</span>
                <span style="background:#e67e22;">Festivo Colombia</span>
            </div>`;

            let html = `<table><tr>
                <th style="position: sticky; left: 0; background: #2c3e50; z-index: 3; min-width:${ROL_COL_W}px; max-width:${ROL_COL_W}px;">Rol</th>
                <th style="position: sticky; left: ${nombreLeft}px; background: #2c3e50; z-index: 3; min-width:${NOMBRE_COL_W}px; max-width:${NOMBRE_COL_W}px;">Nombre</th>`;
            if (hayColumnaNomina) html += `<th style="position: sticky; left: ${nominaLeft}px; background: #2c3e50; z-index: 3; min-width:${NOMINA_COL_W}px; max-width:${NOMINA_COL_W}px;">Código Nómina</th>`;

            headers.forEach((h, index) => {
                let meta = colMeta[index];
                if (!meta.visible) return;
                let tipo = meta.dayType;
                let d = meta.date;
                let thStyle = tipo ? `background:${thColors[tipo]};` : '';
                html += `<th style="${thStyle}">${h}<br><small>${d.getDate()}/${d.getMonth()+1}</small></th>`;
                if (weekAnchor[colWeekKey[index + 1]] === index + 1) html += `<th class="col-semanal">Σ SEMANAL</th>`;
            });
            html += '</tr>';

            datos.forEach(item => {
                let row = item.data;
                let sumasPorSemana = {};
                let info = getRoleInfo(row[0], item.rolReal);
                let rowIdxGlobal = excelData.indexOf(item);

                html += `<tr>
                    <td style="position: sticky; left: 0; background: #f8f9fa; z-index: 1; text-align: left; padding-left: 10px; min-width:${ROL_COL_W}px; max-width:${ROL_COL_W}px;">
                        <span class="${info.class}" style="margin-left:0;">${info.icon}</span> <span style="color:${info.color}; font-weight:bold; font-size:11px;">${item.rol || ''}</span>
                    </td>
                    <td style="position: sticky; left: ${nombreLeft}px; background: #f8f9fa; z-index: 1; text-align: left; padding-left: 10px; min-width:${NOMBRE_COL_W}px; max-width:${NOMBRE_COL_W}px;">
                        <strong style="color: ${info.color};">${row[0]}</strong>
                    </td>`;
                if (hayColumnaNomina) html += `<td style="position: sticky; left: ${nominaLeft}px; background: #eef2f7; z-index: 1; font-weight:bold; color:#34495e; min-width:${NOMINA_COL_W}px; max-width:${NOMINA_COL_W}px;">${item.codigoNomina || '—'}</td>`;
                for(let i = 1; i <= headers.length; i++) {
                    let meta = colMeta[i-1];
                    let val = (row[i] || '').toString().trim().toUpperCase();
                    let wk = colWeekKey[i];
                    if (wk) sumasPorSemana[wk] = (sumasPorSemana[wk] || 0) + horasDeTurno(val);

                    if (!meta.visible) continue;

                    let tdDayClass   = meta.dayType ? `col-${meta.dayType}` : '';
                    let tdExtraClass = (val === 'COMP' && meta.dayType === 'festivo') ? 'comp-festivo' : '';
                    let estilo = tdExtraClass ? null : estiloTurno(val);
                    let esNegrita = !!(meta.dayType || (estilo && estilo.bold));
                    let styleParts = [];
                    if (estilo) { styleParts.push(`background:${estilo.bg} !important`); styleParts.push(`color:${estilo.color}`); }
                    if (esNegrita) styleParts.push('font-weight:bold');
                    let styleAttr = styleParts.length ? ` style="${styleParts.join(';')}"` : '';

                    let optsHtml = turnosDisponibles.map(t =>
                        `<option value="${t}"${t === val ? ' selected' : ''}>${t === '' ? '—' : t}</option>`
                    ).join('');
                    html += `<td class="cell-turno ${tdDayClass} ${tdExtraClass}"${styleAttr}><select class="shift-select" onchange="cambiarTurnoEnMatriz(${rowIdxGlobal},${i},this.value)">${optsHtml}</select></td>`;

                    if (wk && weekAnchor[wk] === i) {
                        html += `<td class="cell-semanal">${formatHorasSemanal(sumasPorSemana[wk])}</td>`;
                    }
                }
                html += '</tr>';
            });
            document.getElementById('tableContainer').innerHTML = html + '</table>';
            generarContadores(datos, colMeta);
        }

        // ============= CONTADORES DIARIOS POR ROL Y TURNO =============
        // Modo resumen: agrupa los turnos horarios en Apertura/Intermedio/Cierre (A/I/C)
        // en vez de mostrar cada código de turno por separado.
        let modoResumenContadores = false;
        let ultimosDatosContadores = null, ultimosColMetaContadores = null;

        function toggleModoResumenContadores() {
            modoResumenContadores = !modoResumenContadores;
            if (ultimosDatosContadores) generarContadores(ultimosDatosContadores, ultimosColMetaContadores);
        }

        function categoriaTurnoResumen(v) {
            let val = (v || '').toString().trim().toUpperCase();
            if (val === '' || val === '(VACÍO)') return '(VACÍO)';
            if (val === 'LIBRE' || val === 'LBRE') return 'LIBRE';
            if (val === 'COMP') return 'COMP';
            if (val === 'VC') return 'VC';
            if (val === 'LIC') return 'LIC';
            if (val === 'INC') return 'INC';
            let m = val.match(/^\d+(?:[.,]\d+)?([AIC])/);
            if (m) return m[1] === 'A' ? 'APERTURA' : (m[1] === 'I' ? 'INTERMEDIO' : 'CIERRE');
            return 'OTROS'; // SIN TURNO, etc.
        }

        // Clasificación A/I/C por HORA DE INICIO (para Equidad por Auxiliar):
        //   A = entra hasta las 8:00 (apertura) · I = entra de 9:00 a 12:30 ·
        //   C = entra desde la 1:00 pm en adelante. No mira la letra del turno.
        function categoriaPorInicio(v) {
            let val = (v || '').toString().trim().toUpperCase();
            if (val === '' || val === '(VACÍO)') return '(VACÍO)';
            if (val === 'LIBRE' || val === 'LBRE') return 'LIBRE';
            if (val === 'COMP') return 'COMP';
            if (val === 'VC') return 'VC';
            if (val === 'LIC') return 'LIC';
            if (val === 'INC') return 'INC';
            let r = _rangoTurnoDec(val);
            if (!r) return 'OTROS';
            if (r.ini < 9) return 'APERTURA';        // hasta las 8:00
            if (r.ini > 12.5) return 'CIERRE';       // desde la 1:00 pm
            return 'INTERMEDIO';                      // de 9:00 a 12:30
        }

        function generarContadores(datos, colMeta) {
            ultimosDatosContadores = datos;
            ultimosColMetaContadores = colMeta;
            let container = document.getElementById('contadoresContainer');
            if (!container || !datos || datos.length === 0) { if(container) container.innerHTML=''; return; }

            // Columnas visibles con metadatos de fecha
            let colsVis = []; // { colIdx (1-based), label, dayType }
            const DIAS_CORTOS = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
            colMeta.forEach((meta, idx) => {
                if (!meta.visible || !meta.date) return;
                let d = meta.date;
                colsVis.push({
                    colIdx : idx + 1,
                    label  : `${DIAS_CORTOS[d.getDay()]} ${d.getDate()}/${d.getMonth()+1}`,
                    dayType: meta.dayType
                });
            });
            if (colsVis.length === 0) { container.innerHTML = ''; return; }

            // Categoría de rol
            function getRolCateg(item) {
                let g = item.rolGenerico.toUpperCase();
                if (g.includes('CAMBISTA'))  return 'Cambista';
                if (g.includes('VISADO'))    return 'Visado';
                if (g.includes('EMERGENTE')) return 'Emergente';
                if (g.includes('SATELITE') || g.includes('SATÉLITE')) return 'Satélite';
                if (g.includes('FUNDACION') || g.includes('FUNDACIÓN')) return 'Fundación';
                if (g.includes('SADOFE'))    return 'SADOFE';
                return item.rolReal === 'CAJERO' ? 'Cajero' : 'Empacador';
            }

            // Acumular: porRol[rol][turno][colIdx] = count  /  general[turno][colIdx] = count
            let porRol = {}, general = {}, allTurnosSet = new Set(), allRoles = [], rolesVistos = new Set();

            datos.forEach(item => {
                let rol = getRolCateg(item);
                if (!rolesVistos.has(rol)) { rolesVistos.add(rol); allRoles.push(rol); }
                if (!porRol[rol]) porRol[rol] = {};

                colsVis.forEach(col => {
                    let vRaw = (item.data[col.colIdx] || '').toString().trim().toUpperCase() || '(VACÍO)';
                    let v = modoResumenContadores ? categoriaTurnoResumen(vRaw) : vRaw;
                    allTurnosSet.add(v);
                    if (!porRol[rol][v])  porRol[rol][v]  = {};
                    if (!general[v])      general[v]      = {};
                    porRol[rol][v][col.colIdx]  = (porRol[rol][v][col.colIdx]  || 0) + 1;
                    general[v][col.colIdx]       = (general[v][col.colIdx]       || 0) + 1;
                });
            });

            let turnosOrdenados;
            if (modoResumenContadores) {
                const ORDEN_RESUMEN = ['APERTURA', 'INTERMEDIO', 'CIERRE', 'LIBRE', 'COMP', 'VC', 'LIC', 'INC', 'OTROS', '(VACÍO)'];
                turnosOrdenados = Array.from(allTurnosSet).sort((a, b) => ORDEN_RESUMEN.indexOf(a) - ORDEN_RESUMEN.indexOf(b));
            } else {
                // Ordenar turnos: orden fijo definido → otros horarios → especiales → vacío
                const ESPECIALES = new Set(['LIBRE','COMP','VC','LIC','INC','(VACÍO)']);
                const ORDEN_FIJO = ['7I7.3-15.3','7I8-16','7I9-17','7I11-19','7C13-20','7C14-21','7C14.3-21.3'];
                turnosOrdenados = Array.from(allTurnosSet).sort((a, b) => {
                    let eA = ESPECIALES.has(a), eB = ESPECIALES.has(b);
                    let iA = ORDEN_FIJO.indexOf(a),  iB = ORDEN_FIJO.indexOf(b);
                    // Especiales siempre al final, (VACÍO) lo último
                    if (eA && eB) { if (a==='(VACÍO)') return 1; if (b==='(VACÍO)') return -1; return a.localeCompare(b); }
                    if (eA) return 1; if (eB) return -1;
                    // Turnos del orden fijo primero
                    if (iA >= 0 && iB >= 0) return iA - iB;
                    if (iA >= 0) return -1;
                    if (iB >= 0) return 1;
                    // Resto de turnos horarios: ordenar numéricamente
                    let rA = a.match(/(\d+)/), rB = b.match(/(\d+)/);
                    if (rA && rB) return parseInt(rA[1]) - parseInt(rB[1]);
                    return a.localeCompare(b);
                });
            }

            // Ordenar roles
            const ORDEN_ROL = ['Cajero','Empacador','Cambista','Visado','Emergente','Satélite','Fundación','SADOFE'];
            allRoles.sort((a,b) => { let iA=ORDEN_ROL.indexOf(a),iB=ORDEN_ROL.indexOf(b); return (iA<0?99:iA)-(iB<0?99:iB); });

            const ROL_COLOR = { Cajero:'#2e7d32', Empacador:'#1565c0', Cambista:'#e65100', Visado:'#d35400', Emergente:'#c0392b', 'Satélite':'#7d3c98', 'Fundación':'#6a1b9a', SADOFE:'#00695c' };
            const TURNO_BG  = { LIBRE:'#fff3f3', COMP:'#f3fff3', VC:'#f5f0ff', LIC:'#fffde7', INC:'#fffde7', '(VACÍO)':'#f5f5f5' };
            const TURNO_BG_RESUMEN = { APERTURA:'#fff3e0', INTERMEDIO:'#f1f8e9', CIERRE:'#e3f2fd', LIBRE:'#fff3f3', COMP:'#f3fff3', VC:'#f5f0ff', LIC:'#fffde7', INC:'#fffde7', OTROS:'#fffde7', '(VACÍO)':'#f5f5f5' };
            const DIA_COL   = { sabado:'#5dade2', domingo:'#2e6da4', festivo:'#e67e22' };

            // ── Construye una tabla individual ──
            function buildTable(rolLabel, dataMap, color, cantPersonas) {
                // Filtrar filas que tienen al menos un valor > 0
                let turnosConDatos = turnosOrdenados.filter(t => colsVis.some(col => (dataMap[t]?.[col.colIdx] || 0) > 0));
                if (turnosConDatos.length === 0) return '';

                let tituloIcono = { Cajero:'💳', Empacador:'🛍️', Cambista:'💵', Visado:'🔍', Emergente:'🚨', 'Fundación':'🤝', SADOFE:'🏷️', General:'🔢' }[rolLabel] || '👤';

                let html = `
                <div style="margin-bottom:22px;">
                  <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                    <h4 style="margin:0; font-size:13px; color:${color}; border-left:4px solid ${color}; padding-left:8px;">
                      ${tituloIcono} ${rolLabel === 'General' ? 'TOTAL GENERAL' : rolLabel.toUpperCase()}
                    </h4>
                    <span style="font-size:11px; color:#95a5a6;">${cantPersonas} persona${cantPersonas!==1?'s':''}</span>
                  </div>
                  <div style="overflow-x:auto; border:1px solid #dee2e6; border-radius:6px; box-shadow:0 1px 4px rgba(0,0,0,0.06);">
                  <table style="border-collapse:collapse; font-size:11px; text-align:center; white-space:nowrap;">
                  <thead><tr>
                    <th style="background:${color}; color:white; padding:7px 14px; text-align:left; position:sticky; left:0; z-index:2; min-width:110px;">Turno</th>`;

                colsVis.forEach(col => {
                    let thBg = col.dayType ? DIA_COL[col.dayType] : color;
                    html += `<th style="background:${thBg}; color:white; padding:6px 7px; min-width:50px; font-size:10px;">${col.label}</th>`;
                });
                html += `<th style="background:#1a252f; color:#f1c40f; padding:6px 10px; min-width:50px;">Total</th>
                  </tr></thead><tbody>`;

                turnosConDatos.forEach((turno, ri) => {
                    let rowBg = (modoResumenContadores ? TURNO_BG_RESUMEN[turno] : TURNO_BG[turno]) || (ri % 2 === 0 ? '#fff' : '#f8f9fa');
                    let etiq  = turno === '(VACÍO)' ? '<em style="color:#aaa">— vacío</em>' : `<b>${turno}</b>`;
                    let rowTotal = 0;
                    html += `<tr style="background:${rowBg};">
                      <td style="text-align:left; padding:5px 14px; position:sticky; left:0; background:${rowBg}; border-right:2px solid #dee2e6;">${etiq}</td>`;
                    colsVis.forEach(col => {
                        let n = dataMap[turno]?.[col.colIdx] || 0;
                        rowTotal += n;
                        html += `<td style="padding:5px 7px; font-weight:${n>0?'700':'400'}; color:${n>0?color:'#ccc'}; border:1px solid #eee;">${n>0?n:'—'}</td>`;
                    });
                    html += `<td style="padding:5px 10px; font-weight:700; color:#2c3e50; background:#eaf0fb; border:1px solid #d5d8dc;">${rowTotal}</td>
                      </tr>`;
                });

                // Fila TOTAL personas por día
                html += `<tr style="background:${color}; color:white; font-weight:700;">
                  <td style="padding:7px 14px; text-align:left; position:sticky; left:0; background:${color}; border-right:2px solid rgba(255,255,255,0.3);">TOTAL / día</td>`;
                let grandTotal = 0;
                colsVis.forEach(col => {
                    let dTotal = turnosConDatos.reduce((s,t) => s + (dataMap[t]?.[col.colIdx]||0), 0);
                    grandTotal += dTotal;
                    html += `<td style="padding:7px 7px;">${dTotal||'—'}</td>`;
                });
                html += `<td style="padding:7px 10px; background:#1a252f; color:#f1c40f;">${grandTotal}</td>
                  </tr></tbody></table></div></div>`;
                return html;
            }

            let periodLabel = `${colsVis.length} día${colsVis.length!==1?'s':''} visibles · ${datos.length} personas`;
            let html = `<div style="margin-top:24px; border-top:2px dashed #ddd; padding-top:20px;">
              <div style="display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; margin-bottom:18px;">
                <h3 style="color:#2c3e50; margin:0; font-size:15px;">
                  📊 Distribución Diaria de Turnos por Rol
                  <span style="font-weight:400; font-size:12px; color:#95a5a6; margin-left:10px;">${periodLabel}</span>
                </h3>
                <button onclick="toggleModoResumenContadores()" style="padding:7px 14px; background:${modoResumenContadores ? '#8e44ad' : '#34495e'}; color:#fff; border:none; border-radius:6px; cursor:pointer; font-size:12px; font-weight:bold; font-family:inherit;">
                  ${modoResumenContadores ? '📋 Ver Detallado (por turno)' : '🗂️ Ver Resumen (Apertura/Intermedio/Cierre)'}
                </button>
              </div>`;

            // Una tabla por rol
            allRoles.forEach(rol => {
                let cantP = datos.filter(item => getRolCateg(item) === rol).length;
                html += buildTable(rol, porRol[rol], ROL_COLOR[rol] || '#555', cantP);
            });

            // Tabla general
            html += buildTable('General', general, '#2c3e50', datos.length);

            html += '</div>';
            container.innerHTML = html;
        }
        // ================================================================

        // ============= INFORME DE EQUIDAD POR AUXILIAR (A/I/C/Libre/Comp por persona) =============
        function poblarFiltroMesEquidadAux() {
            let sel = document.getElementById('eqAuxMesSelector');
            if (!sel) return;
            sel.innerHTML = '<option value="TODOS">Todos los meses</option>';
            if (!headers || headers.length === 0) return;
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            let vistos = new Set();
            headers.forEach((h, idx) => {
                if (h === 'Total Horas Periodo') return;
                let d = new Date(startDate);
                d.setDate(d.getDate() + idx);
                let key = d.getFullYear() * 100 + d.getMonth();
                if (!vistos.has(key)) {
                    vistos.add(key);
                    let label = d.toLocaleDateString('es-CO', { month: 'long', year: 'numeric' });
                    label = label.charAt(0).toUpperCase() + label.slice(1);
                    sel.innerHTML += `<option value="${key}">${label}</option>`;
                }
            });
        }

        function toggleDetalleFila(trigger, id) {
            let el = document.getElementById(id);
            if (!el) return;
            let abierto = el.style.display !== 'none';
            el.style.display = abierto ? 'none' : 'table-row';
            let arrow = trigger.querySelector('.eq-arrow');
            if (arrow) arrow.textContent = abierto ? '▸' : '▾';
        }

        function renderDetalleTiposTurno(detalle) {
            function bloque(cat, color, label) {
                let entradas = Object.entries(detalle[cat] || {}).sort((a, b) => b[1] - a[1]);
                if (entradas.length === 0) return '';
                let chips = entradas.map(([turno, n]) =>
                    `<span style="background:${color}22; color:${color}; border:1px solid ${color}55; border-radius:12px; padding:3px 9px; margin:2px; display:inline-block; font-size:11px;"><b>${turno}</b> × ${n}</span>`
                ).join('');
                return `<div style="margin-bottom:6px;"><strong style="color:${color};">${label}:</strong> ${chips}</div>`;
            }
            let out = bloque('APERTURA', '#e67e22', '🟠 Apertura') +
                      bloque('INTERMEDIO', '#27ae60', '🟢 Intermedio') +
                      bloque('CIERRE', '#2980b9', '🔵 Cierre');
            return out || '<em style="color:#95a5a6;">Sin turnos horarios (A/I/C) en el período seleccionado.</em>';
        }

        function renderEquidadAuxiliar() {
            let tabla = document.getElementById('eqAuxTabla');
            if (!tabla) return;
            if (!excelData || excelData.length === 0) {
                tabla.innerHTML = '<tr><td style="padding:12px; color:#7f8c8d;">Carga primero la Matriz de Turnos.</td></tr>';
                return;
            }

            let mesSelEl = document.getElementById('eqAuxMesSelector');
            let rolSelEl = document.getElementById('eqAuxRolSelector');
            let mesSel = mesSelEl ? mesSelEl.value : 'TODOS';
            let rolSel = rolSelEl ? rolSelEl.value : 'TODOS';

            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());

            // Columnas de día a considerar según el filtro de mes
            let colIndices = [];
            headers.forEach((h, idx) => {
                if (h === 'Total Horas Periodo') return;
                let d = new Date(startDate);
                d.setDate(d.getDate() + idx);
                let key = d.getFullYear() * 100 + d.getMonth();
                if (mesSel === 'TODOS' || String(key) === String(mesSel)) colIndices.push(idx + 1);
            });

            let datosFiltrados = excelData.filter(item => rolSel === 'TODOS' ? true : item.rolReal === rolSel);

            if (colIndices.length === 0 || datosFiltrados.length === 0) {
                tabla.innerHTML = '<tr><td style="padding:12px; color:#7f8c8d;">Sin datos para el mes/rol seleccionado.</td></tr>';
                return;
            }

            let filas = datosFiltrados.map(item => {
                let cuentas = { APERTURA: 0, INTERMEDIO: 0, CIERRE: 0, LIBRE: 0, COMP: 0, VC: 0, LIC: 0, INC: 0, OTROS: 0 };
                let detalle = { APERTURA: {}, INTERMEDIO: {}, CIERRE: {} };
                // Horas trabajadas (para el promedio por día) y COMP por día de semana.
                let horasTrab = 0, diasTrab = 0;
                let compDia = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };  // getDay: 1=Lun … 6=Sáb
                colIndices.forEach(ci => {
                    let val = (item.data[ci] || '').toString().trim().toUpperCase();
                    let cat = categoriaPorInicio(val);   // A/I/C según la hora de inicio
                    if (cuentas[cat] === undefined) cat = 'OTROS';
                    cuentas[cat]++;
                    if (detalle[cat]) detalle[cat][val] = (detalle[cat][val] || 0) + 1;
                    // Fecha real de esta columna (día de la semana)
                    let d = new Date(startDate);
                    d.setDate(d.getDate() + (ci - 1));
                    let dow = d.getDay();
                    let hs = horasDeTurno(val);          // horas si es turno de trabajo; 0 si LIBRE/COMP/VC…
                    if (hs > 0) { horasTrab += hs; diasTrab++; }
                    if (val === 'COMP' && dow >= 1 && dow <= 6) compDia[dow]++;
                });
                let total = Object.values(cuentas).reduce((a, b) => a + b, 0);
                let pct = k => total > 0 ? Math.round((cuentas[k] / total) * 1000) / 10 : 0;
                let horasProm = diasTrab > 0 ? (horasTrab / diasTrab) : 0;
                return { item, cuentas, detalle, pct, horasProm, diasTrab, compDia };
            });

            // Columnas ordenables: clic en el encabezado ordena de MAYOR a MENOR
            // (segundo clic invierte). 'str' ordena alfabético.
            const EQAUX_COLS = [
                { k: 'nombre', label: 'Nombre', tipo: 'str', get: f => (f.item.data[0] || '') },
                { k: 'rol', label: 'Rol', tipo: 'str', get: f => (f.item.rol || '') },
                { k: 'hrs', label: 'Hrs prom/día', tipo: 'num', get: f => f.horasProm || 0 },
                { k: 'A', label: 'A', tipo: 'num', get: f => f.cuentas.APERTURA },
                { k: 'I', label: 'I', tipo: 'num', get: f => f.cuentas.INTERMEDIO },
                { k: 'C', label: 'C', tipo: 'num', get: f => f.cuentas.CIERRE },
                { k: 'Libre', label: 'Libre', tipo: 'num', get: f => f.cuentas.LIBRE },
                { k: 'Comp', label: 'Comp', tipo: 'num', get: f => f.cuentas.COMP },
                { k: 'VC', label: 'VC', tipo: 'num', get: f => f.cuentas.VC },
                { k: 'LIC', label: 'LIC', tipo: 'num', get: f => f.cuentas.LIC },
                { k: 'INC', label: 'INC', tipo: 'num', get: f => f.cuentas.INC },
                { k: 'cLun', label: 'C·Lun', tipo: 'num', gris: true, get: f => f.compDia[1] },
                { k: 'cMar', label: 'C·Mar', tipo: 'num', gris: true, get: f => f.compDia[2] },
                { k: 'cMie', label: 'C·Mié', tipo: 'num', gris: true, get: f => f.compDia[3] },
                { k: 'cJue', label: 'C·Jue', tipo: 'num', gris: true, get: f => f.compDia[4] },
                { k: 'cVie', label: 'C·Vie', tipo: 'num', gris: true, get: f => f.compDia[5] },
                { k: 'cSab', label: 'C·Sáb', tipo: 'num', gris: true, get: f => f.compDia[6] },
                { k: 'pctA', label: '% A', tipo: 'num', get: f => f.pct('APERTURA') },
                { k: 'pctI', label: '% I', tipo: 'num', get: f => f.pct('INTERMEDIO') },
                { k: 'pctC', label: '% C', tipo: 'num', get: f => f.pct('CIERRE') }
            ];
            let sort = window.__eqAuxSort || { k: 'nombre', dir: 1 };
            let colOrd = EQAUX_COLS.find(c => c.k === sort.k) || EQAUX_COLS[0];
            filas.sort((a, b) => {
                let va = colOrd.get(a), vb = colOrd.get(b);
                if (colOrd.tipo === 'str') return sort.dir * String(va).localeCompare(String(vb));
                return sort.dir * ((va || 0) - (vb || 0));
            });

            let html = '<thead><tr style="background:#2980b9; color:white; position:sticky; top:0;">';
            EQAUX_COLS.forEach(c => {
                let flecha = sort.k === c.k ? (sort.dir < 0 ? ' ▼' : ' ▲') : '';
                let bg = c.gris ? ' background:#7f8c8d;' : '';
                let al = c.k === 'nombre' ? ' text-align:left; padding:6px;' : '';
                html += `<th onclick="ordenarEqAux('${c.k}')" title="Clic para ordenar de mayor a menor" style="cursor:pointer; user-select:none;${bg}${al}">${c.label}${flecha}</th>`;
            });
            html += '</tr></thead><tbody>';

            filas.forEach((f, i) => {
                let info = getRoleInfo(f.item.data[0], f.item.rolReal);
                let rowId = 'eqAuxDetalle_' + i;
                let cd = f.compDia;
                let celComp = dw => `<td style="color:#7f8c8d; font-weight:${cd[dw] > 0 ? 'bold' : 'normal'};">${cd[dw] || '·'}</td>`;
                // ⏰ restricción de horario "SOLO HASTA LAS HH" · ✕ si tiene X (no suma)
                let restr = (f.item.restriccion || '').toString();
                let icoRestr = restr.toUpperCase().includes('HASTA')
                    ? ` <span style="font-size:10px; color:#e67e22;" title="${restr}">⏰</span>` : '';
                let icoX = f.item.excluida ? ` <span style="font-size:9px; color:#c0392b;" title="Restricción X: no suma en su rol">✕</span>` : '';
                html += `<tr style="cursor:pointer; border-bottom:1px solid #eee;" onclick="toggleDetalleFila(this, '${rowId}')">
                    <td style="padding:6px; text-align:left;"><span class="eq-arrow" style="display:inline-block; width:14px; color:#7f8c8d;">▸</span><strong style="color:${info.color};">${f.item.data[0]}</strong>${icoRestr}${icoX}</td>
                    <td>${f.item.rol || ''}</td>
                    <td style="font-weight:bold; color:#16a085;">${f.horasProm ? f.horasProm.toFixed(1).replace('.', ',') + 'h' : '—'}</td>
                    <td style="font-weight:bold; color:#e67e22;">${f.cuentas.APERTURA}</td>
                    <td style="font-weight:bold; color:#27ae60;">${f.cuentas.INTERMEDIO}</td>
                    <td style="font-weight:bold; color:#2980b9;">${f.cuentas.CIERRE}</td>
                    <td style="font-weight:bold; color:#c0392b;">${f.cuentas.LIBRE}</td>
                    <td style="font-weight:bold; color:#7f8c8d;">${f.cuentas.COMP}</td>
                    <td style="font-weight:bold; color:#ad1457;">${f.cuentas.VC}</td>
                    <td style="font-weight:bold; color:#f39c12;">${f.cuentas.LIC}</td>
                    <td style="font-weight:bold; color:#d35400;">${f.cuentas.INC}</td>
                    ${celComp(1)}${celComp(2)}${celComp(3)}${celComp(4)}${celComp(5)}${celComp(6)}
                    <td>${f.pct('APERTURA')}%</td><td>${f.pct('INTERMEDIO')}%</td><td>${f.pct('CIERRE')}%</td>
                </tr>`;
                html += `<tr id="${rowId}" style="display:none; background:#f8f9fa;">
                    <td colspan="21" style="padding:10px 20px;">${renderDetalleTiposTurno(f.detalle)}
                        <div style="margin-top:8px; font-size:12px; color:#555;">
                          <b>Resumen de la persona:</b> trabaja en promedio <b>${f.horasProm ? f.horasProm.toFixed(2).replace('.', ',') : 0} horas por día</b>
                          (${f.diasTrab} día${f.diasTrab !== 1 ? 's' : ''} trabajado${f.diasTrab !== 1 ? 's' : ''} en el período).
                          COMP por día: Lun ${cd[1]} · Mar ${cd[2]} · Mié ${cd[3]} · Jue ${cd[4]} · Vie ${cd[5]} · Sáb ${cd[6]}.
                        </div>
                    </td>
                </tr>`;
            });

            tabla.innerHTML = html + '</tbody>';
        }

        // Ordena el informe de Equidad por Auxiliar por la columna elegida.
        // Primer clic en una columna -> de MAYOR a MENOR (números) o A→Z (texto);
        // segundo clic en la misma -> invierte el orden.
        window.ordenarEqAux = function (k) {
            let s = window.__eqAuxSort || { k: 'nombre', dir: 1 };
            if (s.k === k) {
                s.dir = -s.dir;
            } else {
                // por defecto: números de mayor a menor; texto A→Z
                s = { k: k, dir: (k === 'nombre' || k === 'rol') ? 1 : -1 };
            }
            window.__eqAuxSort = s;
            renderEquidadAuxiliar();
        };
        // ================================================================

        // ============= CINTA DE TURNOS (secuencia semanal, solo lectura) =============
        // Se calcula 100% a partir de la Matriz ya cargada (excelData): no toca el
        // motor de generación ni el 27.PY, así que la operación actual no cambia.

        function agruparColumnasPorSemana() {
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            let semanas = [];
            let porKey = {};
            headers.forEach((h, idx) => {
                if (h === 'Total Horas Periodo') return;
                let d = new Date(startDate);
                d.setDate(d.getDate() + idx);
                let dow = d.getDay();
                let wEnd = new Date(d);
                wEnd.setDate(wEnd.getDate() + (7 - dow) % 7);
                let wStart = new Date(wEnd);
                wStart.setDate(wStart.getDate() - 6);
                let key = wEnd.toISOString().slice(0, 10);
                if (!porKey[key]) {
                    porKey[key] = { key, cols: [], inicio: wStart, fin: wEnd };
                    semanas.push(porKey[key]);
                }
                porKey[key].cols.push({ colIdx: idx + 1, date: d, dow });
            });
            semanas.sort((a, b) => a.inicio - b.inicio);
            return semanas;
        }

        // Turno principal de la semana para una persona, SIN contar el COMP.
        // Si el domingo de esa semana fue LIBRE, lo marca aparte (libre=true).
        function turnoPrincipalSemana(item, semana) {
            let conteo = {};
            let libre = false;
            semana.cols.forEach(col => {
                let val = (item.data[col.colIdx] || '').toString().trim().toUpperCase();
                if (val === 'LIBRE' || val === 'LBRE') { libre = true; return; }
                if (val === 'COMP' || val === '') return;
                conteo[val] = (conteo[val] || 0) + 1;
            });
            let ordenado = Object.entries(conteo).sort((a, b) => b[1] - a[1]);
            return { turno: ordenado.length ? ordenado[0][0] : '', libre };
        }

        function poblarFiltroMesCinta() {
            let sel = document.getElementById('cintaMesSelector');
            if (!sel) return;
            sel.innerHTML = '<option value="TODOS">Todos los meses</option>';
            if (!headers || headers.length === 0) return;
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            let vistos = new Set();
            headers.forEach((h, idx) => {
                if (h === 'Total Horas Periodo') return;
                let d = new Date(startDate);
                d.setDate(d.getDate() + idx);
                let key = d.getFullYear() * 100 + d.getMonth();
                if (!vistos.has(key)) {
                    vistos.add(key);
                    let label = d.toLocaleDateString('es-CO', { month: 'long', year: 'numeric' });
                    label = label.charAt(0).toUpperCase() + label.slice(1);
                    sel.innerHTML += `<option value="${key}">${label}</option>`;
                }
            });
        }

        function poblarSelectorPersonaCinta() {
            let sel = document.getElementById('cintaPersonaSelector');
            let rolSelEl = document.getElementById('cintaRolSelector');
            if (!sel) return;
            let rolSel = rolSelEl ? rolSelEl.value : 'TODOS';
            let datos = excelData
                .map((item, idxGlobal) => ({ item, idxGlobal }))
                .filter(x => rolSel === 'TODOS' ? true : x.item.rolReal === rolSel)
                .sort((a, b) => (a.item.data[0] || '').localeCompare(b.item.data[0] || ''));
            sel.innerHTML = datos.map(x => `<option value="${x.idxGlobal}">${x.item.data[0]}</option>`).join('');
        }

        function toggleModoTodosCinta() {
            let modoTodos = document.getElementById('cintaModoTodos').checked;
            let personaWrap = document.getElementById('cintaPersonaWrap');
            if (personaWrap) personaWrap.style.display = modoTodos ? 'none' : 'flex';
            renderCintaTurnos();
        }

        function renderResumenCintaPersona(item, mesSel) {
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            let cuentas = { APERTURA: 0, INTERMEDIO: 0, CIERRE: 0, LIBRE: 0, COMP: 0, VC: 0, LIC: 0, INC: 0, OTROS: 0 };
            headers.forEach((h, idx) => {
                if (h === 'Total Horas Periodo') return;
                let d = new Date(startDate);
                d.setDate(d.getDate() + idx);
                let key = d.getFullYear() * 100 + d.getMonth();
                if (mesSel !== 'TODOS' && String(key) !== String(mesSel)) return;
                let val = (item.data[idx + 1] || '').toString().trim().toUpperCase();
                let cat = categoriaTurnoResumen(val);
                if (cuentas[cat] === undefined) cat = 'OTROS';
                cuentas[cat]++;
            });
            let chip = (label, n, color) =>
                `<div style="background:${color}22; border:1px solid ${color}55; color:${color}; border-radius:8px; padding:6px 12px; font-size:12px; font-weight:bold;">${label}: ${n}</div>`;
            return `<div style="display:flex; gap:8px; flex-wrap:wrap;">
                ${chip('Apertura', cuentas.APERTURA, '#e67e22')}
                ${chip('Intermedio', cuentas.INTERMEDIO, '#27ae60')}
                ${chip('Cierre', cuentas.CIERRE, '#2980b9')}
                ${chip('Libre', cuentas.LIBRE, '#c0392b')}
                ${chip('Comp', cuentas.COMP, '#7f8c8d')}
                ${chip('VC', cuentas.VC, '#ad1457')}
                ${chip('LIC', cuentas.LIC, '#f39c12')}
                ${chip('INC', cuentas.INC, '#d35400')}
            </div>`;
        }

        function renderCintaPersona(item, semanas) {
            let beads = semanas.map(semana => {
                let { turno, libre } = turnoPrincipalSemana(item, semana);
                let estilo = estiloTurno(turno);
                let bg = estilo ? estilo.bg : '#f5f5f5';
                let color = estilo ? estilo.color : '#aaa';
                let labelFecha = `${semana.inicio.getDate()}/${semana.inicio.getMonth() + 1}–${semana.fin.getDate()}/${semana.fin.getMonth() + 1}`;
                return `<div style="display:flex; flex-direction:column; align-items:center; min-width:120px;">
                    <div style="font-size:10px; color:#95a5a6; margin-bottom:4px;">${labelFecha}</div>
                    <div style="background:${bg}; color:${color}; border-radius:8px; padding:8px 12px; font-weight:bold; font-size:12px; text-align:center; min-width:110px;">${turno || '—'}</div>
                    ${libre ? '<div style="margin-top:4px; background:#ffcdd2; color:#b71c1c; border-radius:8px; padding:3px 10px; font-weight:bold; font-size:10px; text-align:center;">LIBRE Dom</div>' : ''}
                </div>`;
            });
            return `<div style="display:flex; align-items:flex-start; gap:6px; padding:10px 4px;">
                ${beads.join('<div style="align-self:center; font-size:18px; color:#bdc3c7; margin-top:26px;">→</div>')}
            </div>`;
        }

        function renderCintaTodos(datos, semanas) {
            let html = '<table style="border-collapse:collapse; font-size:11px; text-align:center; white-space:nowrap;">' +
                '<thead><tr style="background:#16a085; color:white; position:sticky; top:0;">' +
                '<th style="padding:6px 14px; text-align:left; position:sticky; left:0; background:#16a085; z-index:2;">Nombre</th>';
            semanas.forEach(s => {
                html += `<th style="padding:5px 8px; font-size:10px;">${s.inicio.getDate()}/${s.inicio.getMonth() + 1}</th>`;
            });
            html += '</tr></thead><tbody>';

            datos.forEach(item => {
                html += `<tr style="border-bottom:1px solid #eee;">
                    <td style="padding:5px 14px; text-align:left; position:sticky; left:0; background:#f8f9fa; font-weight:bold; border-right:2px solid #dee2e6;">${item.data[0]}</td>`;
                semanas.forEach(semana => {
                    let { turno, libre } = turnoPrincipalSemana(item, semana);
                    let estilo = estiloTurno(turno);
                    let bg = estilo ? estilo.bg : '#f5f5f5';
                    let color = estilo ? estilo.color : '#aaa';
                    let texto = turno || '—';
                    if (libre) texto += '<br><span style="color:#b71c1c; font-size:9px; font-weight:bold;">LIBRE Dom</span>';
                    html += `<td style="padding:5px 6px; background:${bg}; color:${color}; font-weight:bold; border:1px solid #eee;">${texto}</td>`;
                });
                html += '</tr>';
            });
            return html + '</tbody></table>';
        }

        function renderCintaTurnos() {
            let container = document.getElementById('cintaContainer');
            let resumenEl = document.getElementById('cintaResumenPersona');
            if (!container) return;
            if (!excelData || excelData.length === 0) {
                container.innerHTML = '<div style="padding:12px; color:#7f8c8d;">Carga primero la Matriz de Turnos.</div>';
                if (resumenEl) resumenEl.innerHTML = '';
                return;
            }

            let mesSel = document.getElementById('cintaMesSelector').value || 'TODOS';
            let rolSel = document.getElementById('cintaRolSelector').value || 'TODOS';
            let modoTodos = document.getElementById('cintaModoTodos').checked;

            let semanas = agruparColumnasPorSemana().filter(s => mesSel === 'TODOS' ||
                String(s.inicio.getFullYear() * 100 + s.inicio.getMonth()) === String(mesSel));

            let datosFiltrados = excelData.filter(item => rolSel === 'TODOS' ? true : item.rolReal === rolSel);

            if (semanas.length === 0 || datosFiltrados.length === 0) {
                container.innerHTML = '<div style="padding:12px; color:#7f8c8d;">Sin datos para el mes/rol seleccionado.</div>';
                if (resumenEl) resumenEl.innerHTML = '';
                return;
            }

            if (modoTodos) {
                if (resumenEl) resumenEl.innerHTML = '';
                container.innerHTML = renderCintaTodos(datosFiltrados, semanas);
                return;
            }

            let personaSelEl = document.getElementById('cintaPersonaSelector');
            let idx = personaSelEl && personaSelEl.value !== '' ? parseInt(personaSelEl.value) : -1;
            let item = idx >= 0 ? excelData[idx] : datosFiltrados[0];
            if (!item) {
                container.innerHTML = '<div style="padding:12px; color:#7f8c8d;">Selecciona una persona.</div>';
                if (resumenEl) resumenEl.innerHTML = '';
                return;
            }

            container.innerHTML = renderCintaPersona(item, semanas);
            if (resumenEl) resumenEl.innerHTML = renderResumenCintaPersona(item, mesSel);
        }
        // ================================================================

        function cambiarTurnoEnMatriz(rowIdx, colIdx, nuevoValor) {
            excelData[rowIdx].data[colIdx] = nuevoValor;

            // Guardar posición de scroll antes de re-renderizar
            let tc = document.getElementById('tableContainer');
            let sl = tc ? tc.scrollLeft : 0;
            let st = tc ? tc.scrollTop : 0;

            aplicarFiltrosGlobales();

            // Restaurar posición de scroll
            let tc2 = document.getElementById('tableContainer');
            if (tc2) { tc2.scrollLeft = sl; tc2.scrollTop = st; }

            // Si el día modificado es el que está abierto en la vista diaria, actualizar todo
            let selectElem = document.getElementById('daySelector');
            if (selectElem && parseInt(selectElem.value) === colIdx) {
                procesarDiaEspecificoSeguro();
            }
        }

        // Cambia el turno de una persona DESDE la línea de tiempo (día colIdx) y
        // propaga el cambio a la Matriz de Turnos, la propia línea de tiempo, el
        // Cumplimiento, la Equidad por Auxiliar y la vista diaria si está abierta.
        window.cambiarTurnoDesdeTimeline = function (idx, colIdx, valor) {
            if (idx == null || !excelData || !excelData[idx]) return;
            excelData[idx].data[colIdx] = valor;
            if (typeof aplicarFiltrosGlobales === 'function') aplicarFiltrosGlobales();      // Matriz
            if (typeof renderLineaTiempoCobertura === 'function') renderLineaTiempoCobertura(); // esta vista
            if (typeof renderCumplimiento === 'function') renderCumplimiento();
            if (typeof renderEquidadAuxiliar === 'function') renderEquidadAuxiliar();
            if (typeof renderChartDia === 'function') { try { renderChartDia(); } catch (e) {} }
            let daySel = document.getElementById('daySelector');
            if (daySel && parseInt(daySel.value) === colIdx && typeof procesarDiaEspecificoSeguro === 'function') {
                procesarDiaEspecificoSeguro();
            }
        };

        // Clic en el encabezado de una hora: muestra/oculta la velocidad de caja
        // (segundos por cliente y si la atención es rápida/promedio/lenta).
        window.verVelocidadHora = function (h) {
            window.__VEL_HORA = (window.__VEL_HORA === h) ? null : h;
            if (typeof renderLineaTiempoCobertura === 'function') renderLineaTiempoCobertura();
        };

        function llenarSelectorDias() {
            let select = document.getElementById('daySelector');
            select.innerHTML = '';
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            headers.forEach((h, i) => {
                if(h !== 'Total Horas Periodo') {
                    let d = new Date(startDate);
                    d.setDate(d.getDate() + i);
                    select.innerHTML += `<option value="${i+1}">${h} (${d.toLocaleDateString()})</option>`;
                }
            });
        }

        // ============= CRONOGRAMA DE COBERTURA POR PERSONA (línea de tiempo) =============
        // Convierte un token de hora de la nomenclatura V26 a horas decimales.
        // El ".3" significa ":30" (media hora). Ej: "14.3" -> 14.5 ; "8" -> 8.0
        function _horaDec(tok) {
            let p = String(tok).split('.');
            let hh = parseInt(p[0], 10);
            if (isNaN(hh)) return null;
            if (p.length < 2 || p[1] === '') return hh;
            let f = p[1];
            let min = (f.length === 1) ? parseInt(f, 10) * 10 : parseInt(f, 10); // "3"->30, "15"->15
            return hh + (min / 60);
        }
        // Rango [inicio, fin] en horas decimales de un turno; null si no es turno horario.
        function _rangoTurnoDec(t) {
            let s = (t || '').toString().trim().toUpperCase();
            let m = s.match(/(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)/);
            if (!m) return null;
            let ini = _horaDec(m[1]), fin = _horaDec(m[2]);
            if (ini === null || fin === null) return null;
            if (fin <= ini) fin += 24; // por si cruzara medianoche (turnos noche)
            return { ini, fin };
        }
        // Letra del turno (A/I/C/N) para el color de la barra.
        function _letraTurnoCober(t) {
            let m = (t || '').toString().trim().toUpperCase().match(/^\d+(?:[.,]\d+)?([AIC N])/);
            return m ? m[1].trim() : '';
        }
        // Horas laborales (el número del prefijo). Ej: "6,5A8-14.3" -> 6.5
        function _horasLaboralesTurno(t) {
            let m = (t || '').toString().trim().match(/^(\d+(?:[.,]\d+)?)/);
            return m ? parseFloat(m[1].replace(',', '.')) : null;
        }
        // ¿El turno incluye hora de almuerzo? Solo si el TRAMO físico (fin-inicio)
        // supera las horas laborales (≈1h). Los turnos donde tramo == laborales
        // (ej. 6,5A8-14.3 de apertura, o 7C15-22 de cierre) NO tienen almuerzo.
        function _tieneAlmuerzo(t) {
            let r = _rangoTurnoDec(t);
            let lab = _horasLaboralesTurno(t);
            if (!r || lab === null) return false;
            return (r.fin - r.ini) - lab >= 0.5;
        }

        // Categoría de rol para la cobertura (mismo criterio que los contadores).
        function _rolCategoria(item) {
            let g = (item.rolGenerico || '').toUpperCase();
            if (g.includes('CAMBISTA')) return 'Cambista';
            if (g.includes('VISADO')) return 'Visado';
            if (g.includes('EMERGENTE')) return 'Emergente';
            if (g.includes('SATELITE') || g.includes('SATÉLITE')) return 'Satélite';
            if (g.includes('FUNDACION') || g.includes('FUNDACIÓN')) return 'Fundación';
            if (g.includes('SADOFE')) return 'SADOFE';
            return item.rolReal === 'CAJERO' ? 'Cajero' : 'Empacador';
        }

        // Almuerzos reales del día (nombre -> [ini,fin] decimal), reutilizando el
        // motor de relevos. TODOS los días los almuerzos se reparten entre las
        // 12:00 y las 16:00 (salidas 12-13, 13-14, 14-15 y 15-16) minimizando
        // el impacto: cada almuerzo va al slot con MÁS holgura de cobertura en
        // ese momento (personas en caja menos almuerzos ya asignados menos la
        // meta de la hora), así no salen todos en un mismo núcleo y los grupos
        // NO tienen que ser iguales. Solo se reubica si cabe dentro del turno
        // (1h después de entrar y 1h antes de salir); si no, conserva su hora.
        function _almuerzosDelDia(colIdx, esDomingoDia) {
            let almuerzoDe = {};
            try {
                ['CAJERO', 'EMPACADOR'].forEach(rr => {
                    let grupo = excelData
                        .filter(it => it.rolReal === rr && _rangoTurnoDec(it.data[colIdx]))
                        .map(it => ({ nombre: it.data[0], turno: (it.data[colIdx] || '').toString().trim(), rol: rr }));
                    if (!grupo.length) return;
                    let relevos = generarRelevosAvanzados(grupo) || [];
                    relevos.forEach(rl => {
                        // Solo se marca ALM si el turno de esa persona realmente
                        // incluye almuerzo (los de 6,5h de apertura/cierre no).
                        if (rl.tipo && rl.tipo.toLowerCase().includes('almuerzo') && _tieneAlmuerzo(rl.turno)) {
                            almuerzoDe[rl.nombre] = [rl.minSalida / 60, rl.minRegreso / 60];
                        }
                    });
                });
            } catch (e) { /* si algo falla, simplemente no se marca ALM */ }
            // --- Reparto óptimo 12:00-15:00 (todos los días) ---
            const SLOTS_ALM = [12, 13, 14, 15];   // salidas 12-13, 13-14, 14-15 y 15-16
            let rangoDe = {};
            let cobBase = { 12: 0, 13: 0, 14: 0, 15: 0 };
            excelData.forEach(it => {
                let r = _rangoTurnoDec((it.data[colIdx] || '').toString().trim());
                if (!r) return;
                rangoDe[it.data[0]] = r;
                if (it.excluida) return;
                SLOTS_ALM.forEach(h => { if (r.ini < h + 1 && r.fin > h) cobBase[h]++; });
            });
            let bateriaAlm = Number(window.APP_BATERIA) || 0;
            let metasAlm = bateriaAlm > 0 ? _metasHoraDia(bateriaAlm, esDomingoDia) : { 12: 0, 13: 0, 14: 0, 15: 0 };
            let almEnSlot = { 12: 0, 13: 0, 14: 0, 15: 0 };
            let nombresAlm = Object.keys(almuerzoDe).sort((a, b) => {
                let ia = rangoDe[a] ? rangoDe[a].ini : 0, ib = rangoDe[b] ? rangoDe[b].ini : 0;
                return ia - ib || a.localeCompare(b);
            });
            nombresAlm.forEach(n => {
                let r = rangoDe[n];
                if (!r) return;
                let dur = Math.max(0.5, Math.min(1, almuerzoDe[n][1] - almuerzoDe[n][0]));
                // slots factibles: dentro del turno (1h tras entrar, 1h antes de
                // salir) y con el almuerzo terminando a más tardar a las 16:00
                let candidatos = SLOTS_ALM.filter(h =>
                    h >= r.ini + 1 - 1e-9 && h + dur <= Math.min(r.fin - 1, 16) + 1e-9);
                if (!candidatos.length) return;   // no cabe en la ventana: conserva su hora
                // el slot con MÁS holgura (cobertura restante - meta) recibe el almuerzo
                let mejor = candidatos[0];
                let holgura = h => (cobBase[h] - almEnSlot[h]) - (metasAlm[h] || 0);
                candidatos.forEach(h => { if (holgura(h) > holgura(mejor)) mejor = h; });
                almuerzoDe[n] = [mejor, mejor + dur];
                almEnSlot[mejor]++;
            });
            return almuerzoDe;
        }

        // ===== ANCLAS PERSONALIZADAS POR TIENDA =====
        // Cajas ABIERTAS exigidas por hora, definidas por cada tienda. Deben ir
        // a la par de la cinta del PY de esa tienda (ej. 28_AKB30.PY). Cuando el
        // administrador entra con una tienda que está aquí, TODAS las alertas
        // (banner, Cumplimiento, justificación, simulador) usan estos números y
        // se ignoran los % genéricos y la regla dominical del 80%.
        // Las horas no listadas se interpolan en rampa entre anclas.
        const ANCLAS_TIENDA = {
            // Mínimo de CAJAS ABIERTAS por hora (cajeros + emergentes) para AKB30.
            'AKB30': { 8: 7, 9: 10, 10: 16, 11: 19, 12: 19, 13: 19, 14: 18, 15: 19,
                       16: 19, 17: 19, 18: 19, 19: 19, 20: 12, 21: 8 }
        };
        // Ajuste de mínimos SEGÚN EL FILTRO de días del informe de Cumplimiento:
        //   Todos -> curva base · L-V -> 10:00=15 y 19:00=18 · Sáb/Dom -> 10:00=14.
        const MIN_HORA_OVERRIDE = {
            'AKB30': { LV: { 10: 15, 19: 18 }, SAB: { 10: 14 }, DOM: { 10: 14 } }
        };
        let _minOverrideActivo = null;   // se activa solo mientras corre renderCumplimiento
        function _tablaAnclasTienda() {
            const t = (window.APP_TIENDA || '').toUpperCase();
            const base = ANCLAS_TIENDA[t];
            if (!base) return null;
            return _minOverrideActivo ? Object.assign({}, base, _minOverrideActivo) : base;
        }
        // Descripción del perfil vigente (para banners e informes)
        function _descPerfil() {
            const tabla = _tablaAnclasTienda();
            if (tabla) {
                const hs = Object.keys(tabla).map(Number).sort((a, b) => a - b);
                return `anclas propias de ${(window.APP_TIENDA || 'la tienda')}: ` +
                       hs.map(h => `${h}:00→${tabla[h]}`).join(' · ') + ' cajas abiertas';
            }
            return '1º ≥25% GARANTIZADO a las 8:00 (meta 30%) · 2º ≥60% GARANTIZADO al cierre (21:30) · 3º 100% de 11:00 a 13:00 (mín. 85%)';
        }

        // Anclas del PERFIL DE COBERTURA que rige TODOS los días (% de la batería),
        // en ORDEN DE PRIORIDAD según el histórico real:
        //   1) 8:00 -> 30% · 2) cierre del último turno (21:30) -> 60%
        //   3) 11:00-13:00 -> meta 100%, aceptable desde el 85%.
        //   Domingo además: >= 80% entre 13:00 y 20:00.
        // La hora 21 la cubren SOLO los turnos que cierran a las 21:30.
        const H_ULTIMO_TURNO = 21;
        const PICO_MIN_ACEPTABLE = 0.85;
        // Devuelve las anclas con su meta ABSOLUTA (objN, en cajeros). Si la
        // tienda definió objetivos absolutos (window.APP_OBJETIVOS, guardados
        // desde el Python en _CONFIG), esos números mandan; si no, se usa el
        // % del perfil sobre la batería. Incluye el pico de la tarde 16-19
        // cuando la tienda definió ese objetivo.
        function _anclasPerfil(esDomingoDia, bateria) {
            // Tienda con anclas propias: esos números mandan (todos los días)
            const tabla = _tablaAnclasTienda();
            if (tabla) {
                return Object.keys(tabla).map(Number).sort((a, b) => a - b).map(h => ({
                    h: h, objN: tabla[h],
                    txt: `Ancla ${(window.APP_TIENDA || '').toUpperCase()} ${h}:00 — ${tabla[h]} cajas abiertas`,
                    ultimo: h === H_ULTIMO_TURNO
                }));
            }
            const O = window.APP_OBJETIVOS || {};
            const nApe = (O.apertura > 0) ? O.apertura : Math.ceil(bateria * 0.25);
            const nCie = (O.cierre > 0) ? O.cierre : Math.ceil(bateria * 0.60);
            const nPam = (O.pico_am > 0) ? O.pico_am : Math.ceil(bateria * PICO_MIN_ACEPTABLE);
            let anclas = [
                { h: 8, objN: nApe, txt: 'PRIORIDAD 5: apertura 8:00 — ' + (O.apertura > 0 ? `objetivo de la tienda: ${nApe} cajeros` : 'meta 30%, mínimo GARANTIZADO 25%') },
                { h: H_ULTIMO_TURNO, objN: nCie, txt: 'PRIORIDAD 4: cierre 21:00-21:30 — ' + (O.cierre > 0 ? `objetivo de la tienda: ${nCie} cajeros` : 'mínimo GARANTIZADO 60%'), ultimo: true },
                { h: 11, objN: nPam, txt: 'PRIORIDAD 5: pico 11:00-12:00 — ' + (O.pico_am > 0 ? `objetivo: ${nPam} cajeros` : 'meta 100%, mínimo 85%') },
                { h: 12, objN: nPam, txt: 'PRIORIDAD 5: pico 12:00-13:00 — ' + (O.pico_am > 0 ? `objetivo: ${nPam} cajeros` : 'meta 100%, mínimo 85%') }
            ];
            if (esDomingoDia) {
                for (let h = 13; h < 20; h++) {
                    let objN = Math.ceil(bateria * 0.80);
                    if (O.pico_pm > 0 && h >= 15) objN = Math.max(objN, O.pico_pm);
                    anclas.push({ h: h, objN: objN, txt: '80% de la tarde dominical' + (O.pico_pm > 0 && h >= 16 ? ` / objetivo 16-19: ${O.pico_pm}` : '') });
                }
            } else if (O.pico_pm > 0) {
                for (let h = 15; h < 20; h++) anclas.push({ h: h, objN: O.pico_pm, txt: `Pico de la tarde ${h}:00 — objetivo de la tienda: ${O.pico_pm} cajeros (prioridad 5)` });
            }
            return anclas;
        }

        // Cajas cubiertas por hora en un día (Cajero + Emergente, sin personas
        // con X, descontando la hora de almuerzo). Devuelve {hora: cajas}.
        function _cobCajasDia(colIdx, esDomingoDia) {
            let almuerzoDe = _almuerzosDelDia(colIdx, esDomingoDia);
            let cob = {};
            for (let h = 5; h < 24; h++) cob[h] = 0;
            excelData.forEach(item => {
                if (item.excluida) return;
                let cat = _rolCategoria(item);
                if (cat !== 'Cajero' && cat !== 'Emergente') return;
                let r = _rangoTurnoDec((item.data[colIdx] || '').toString().trim());
                if (!r) return;
                let alm = almuerzoDe[item.data[0]];
                for (let h = 5; h < 24; h++) {
                    // La hora h cuenta SOLO si la persona está TRABAJANDO durante la
                    // primera media hora (h:00-h:30): quien entra a las h:30 cuenta
                    // desde la franja siguiente (no se marca personal que aún no está).
                    if (!(r.ini <= h + 1e-9 && r.fin >= h + 0.5 - 1e-9)) continue;
                    if (alm && alm[0] < h + 0.5 - 1e-9 && alm[1] > h + 1e-9) continue;
                    cob[h]++;
                }
            });
            return cob;
        }

        // Cobertura de la SEGUNDA media hora (h:30-h+1:00): sirve para saber si a
        // las :30 llega más personal (entonces el faltante de esa hora dura solo
        // los primeros 30 min).
        function _cobCajasSeg(colIdx, esDomingoDia) {
            let almuerzoDe = _almuerzosDelDia(colIdx, esDomingoDia);
            let cob = {};
            for (let h = 5; h < 24; h++) cob[h] = 0;
            excelData.forEach(item => {
                if (item.excluida) return;
                let cat = _rolCategoria(item);
                if (cat !== 'Cajero' && cat !== 'Emergente') return;
                let r = _rangoTurnoDec((item.data[colIdx] || '').toString().trim());
                if (!r) return;
                let alm = almuerzoDe[item.data[0]];
                for (let h = 5; h < 24; h++) {
                    if (!(r.ini <= h + 0.5 + 1e-9 && r.fin >= h + 1 - 1e-9)) continue;   // trabaja h:30-h+1:00
                    if (alm && alm[0] < h + 1 - 1e-9 && alm[1] > h + 0.5 + 1e-9) continue;
                    cob[h]++;
                }
            });
            return cob;
        }

        // Evalúa el perfil de un día: lista de horas incumplidas (con faltantes)
        function _evaluarPerfilDia(colIdx, esDomingoDia, bateria) {
            let cob1 = _cobCajasDia(colIdx, esDomingoDia);   // primera media hora (h:00)
            let cobSeg = _cobCajasSeg(colIdx, esDomingoDia); // segunda media hora (h:30)
            // Cobertura "asentada" de cada hora = la MAYOR de las dos medias horas.
            // Así la transición de 30 min (p.ej. gente que entra a las 10:30) no
            // dispara la alerta, y los cierres (que salen a las :30) no se penalizan.
            let cob = {};
            for (let h = 5; h < 24; h++) cob[h] = Math.max(cob1[h] || 0, cobSeg[h] || 0);
            let fallos = [], falloUltimo = false;
            _anclasPerfil(esDomingoDia, bateria).forEach(a => {
                let n = cob[a.h] || 0;
                if (a.objN > 0 && n < a.objN) {
                    fallos.push({ h: a.h, n: n, obj: a.objN / bateria, pct: n / bateria,
                                  faltan: a.objN - n, txt: a.txt, ultimo: !!a.ultimo });
                    if (a.ultimo) falloUltimo = true;
                }
            });
            return { fallos: fallos, falloUltimo: falloUltimo, cob: cob };
        }

        // ===== JUSTIFICACIÓN MATEMÁTICA DE LA FALTA DE PERSONAL =====
        // Curva de metas por hora (espejo de metas_diarias del Python):
        // objetivos absolutos de la tienda + rampas; % de la batería de respaldo.
        function _metasHoraDia(bateria, esDomingoDia) {
            // Tienda con anclas propias: curva construida desde su tabla, con
            // rampas lineales entre anclas (misma curva para todos los días).
            const tabla = _tablaAnclasTienda();
            if (tabla) {
                const hs = Object.keys(tabla).map(Number).sort((a, b) => a - b);
                let metas = {};
                for (let h = 8; h <= 21; h++) {
                    if (tabla[h] !== undefined) { metas[h] = tabla[h]; continue; }
                    let prev = null, next = null;
                    hs.forEach(a => { if (a < h) prev = a; if (a > h && next === null) next = a; });
                    if (prev === null) metas[h] = tabla[hs[0]];
                    else if (next === null) metas[h] = tabla[prev];
                    else metas[h] = Math.ceil(tabla[prev] + (tabla[next] - tabla[prev]) * (h - prev) / (next - prev));
                }
                return metas;
            }
            const O = window.APP_OBJETIVOS || {};
            const pct = h => {
                let p;
                if (h >= 11 && h <= 13) p = 1.0;
                else if (h < 11) p = 0.30 + Math.max(0, h - 8) * (0.70 / 3);
                else if (h >= 21) p = 0.60;
                else p = 1.0 - (h - 13) * (0.40 / 8);
                return p;
            };
            const val = (k, hRef) => (O[k] > 0) ? O[k] : Math.ceil(bateria * pct(hRef));
            const a = val('apertura', 8), pm = val('pico_am', 11), pt = val('pico_pm', 16), c = val('cierre', 21);
            let metas = {};
            for (let h = 8; h <= 21; h++) {
                let v;
                if (h === 8) v = a;
                else if (h === 9 || h === 10) v = a + (pm - a) * (h - 8) / 3;
                else if (h >= 11 && h <= 13) v = pm;
                else if (h === 14) v = pm + (pt - pm) * 0.5;
                else if (h >= 15 && h <= 19) v = pt;
                else if (h === 20) v = pt + (c - pt) * 0.5;
                else v = c;
                if (esDomingoDia && h >= 13 && h < 20) v = Math.max(v, bateria * 0.80);
                metas[h] = Math.min(bateria, Math.ceil(v));
            }
            return metas;
        }

        // Dos COTAS matemáticas de personal mínimo para un día:
        //  1) HORAS-CAJA: la meta suma D horas-caja; la plantilla del día aporta
        //     S = Σ horas laborales. Si S < D, NINGÚN acomodo de turnos alcanza:
        //     faltan ceil((D-S)/6.5) personas como mínimo.
        //  2) VENTANAS INCOMPATIBLES: un turno cubre a lo sumo un tramo de 9h
        //     (8h + 1h de almuerzo). Dos horas separadas 10h o más (ej. 11:00 y
        //     21:00) NUNCA las cubre la misma persona: se necesitan al menos
        //     meta(h1) + meta(h2) personas ese día.
        // El faltante matemático del día es el MAYOR de ambas cotas.
        function _justifPersonalDia(colIdx, esDomingoDia, bateria) {
            const metas = _metasHoraDia(bateria, esDomingoDia);
            const cob = _cobCajasDia(colIdx, esDomingoDia);
            let P = 0, S = 0;
            excelData.forEach(item => {
                if (item.excluida) return;
                const cat = _rolCategoria(item);
                if (cat !== 'Cajero' && cat !== 'Emergente') return;
                const t = (item.data[colIdx] || '').toString().trim();
                if (!_rangoTurnoDec(t)) return;
                P++;
                S += _horasLaboralesTurno(t) || 0;
            });
            const D = Object.values(metas).reduce((x, y) => x + y, 0);
            const deficitHoras = Math.max(0, D - S);
            const porHoras = Math.ceil(deficitHoras / 6.5);
            let reqVentanas = 0, mejorPar = null;
            for (let h1 = 8; h1 <= 11; h1++) {
                for (let h2 = h1 + 10; h2 <= 21; h2++) {
                    const req = (metas[h1] || 0) + (metas[h2] || 0);
                    if (req > reqVentanas) { reqVentanas = req; mejorPar = [h1, h2]; }
                }
            }
            const porVentanas = Math.max(0, reqVentanas - P);
            const faltan = Math.max(porHoras, porVentanas);
            return { P: P, S: S, D: D, deficitHoras: deficitHoras, porHoras: porHoras,
                     reqVentanas: reqVentanas, mejorPar: mejorPar, porVentanas: porVentanas,
                     faltan: faltan, metas: metas, cob: cob };
        }

        // ===== SIMULADOR: ¿qué pasa si añado N cajeros extra de X horas? =====
        // Coloca cada persona virtual en el inicio de turno que MÁS huecos tapa
        // (greedy), con su hora de almuerzo en la hora de menor déficit del
        // tramo, y devuelve la cobertura simulada del día.
        function _simulaDia(cob, metas, n, dur) {
            let cobS = {};
            for (let h = 5; h < 24; h++) cobS[h] = cob[h] || 0;
            let colocados = [];
            for (let k = 0; k < n; k++) {
                let best = null;
                for (let s = 7.5; s <= 21.5 - dur; s += 0.5) {
                    const span = (s + dur + 1 <= 21.5) ? dur + 1 : dur;   // con almuerzo si cabe
                    const fin = s + span;
                    let horas = [];
                    for (let h = Math.floor(s); h < Math.ceil(fin); h++) if (h >= 8 && h <= 21) horas.push(h);
                    let almH = -1;
                    if (span > dur && horas.length > 2) {
                        let minDef = Infinity;
                        horas.slice(1, -1).forEach(h => {
                            const def = (metas[h] || 0) - cobS[h];
                            if (def < minDef) { minDef = def; almH = h; }
                        });
                    }
                    let gain = 0;
                    horas.forEach(h => { if (h !== almH && (metas[h] || 0) - cobS[h] > 0) gain++; });
                    if (!best || gain > best.gain) best = { s: s, horas: horas, almH: almH, gain: gain };
                }
                if (!best || best.gain <= 0) break;
                best.horas.forEach(h => { if (h !== best.almH) cobS[h]++; });
                const hIni = Math.floor(best.s) + (best.s % 1 ? ':30' : ':00');
                colocados.push(hIni);
            }
            return { cobS: cobS, colocados: colocados };
        }

        window.simularPersonalExtra = function () {
            const n = parseInt((document.getElementById('simExtraN') || {}).value) || 0;
            const dur = parseFloat((document.getElementById('simExtraDur') || {}).value) || 6.5;
            window.__SIM_EXTRA = n > 0 ? { n: n, dur: dur } : null;
            renderCumplimiento();
        };
        window.limpiarSimulacion = function () {
            window.__SIM_EXTRA = null;
            renderCumplimiento();
        };

        function llenarSelectorDiasCobertura() {
            let select = document.getElementById('covDaySelector');
            if (!select) return;
            select.innerHTML = '';
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            headers.forEach((h, i) => {
                if (h !== 'Total Horas Periodo') {
                    let d = new Date(startDate);
                    d.setDate(d.getDate() + i);
                    select.innerHTML += `<option value="${i + 1}">${h} (${d.toLocaleDateString()})</option>`;
                }
            });
        }

        function renderLineaTiempoCobertura() {
            let cont = document.getElementById('covTimelineContainer');
            if (!cont) return;
            if (!excelData || excelData.length === 0) {
                cont.innerHTML = '<div style="color:#7f8c8d; padding:10px;">Carga la Matriz de Turnos para ver la cobertura por persona.</div>';
                return;
            }
            let sel = document.getElementById('covDaySelector');
            let colIdx = sel && sel.value ? parseInt(sel.value) : 1;

            // Categoría de rol (igual criterio que los contadores)
            let rolCateg = _rolCategoria;
            const ORDEN_ROL = ['Cajero', 'Emergente', 'Satélite', 'Empacador', 'Cambista', 'Visado', 'Fundación', 'SADOFE'];
            const ROL_ICONO = { Cajero: '💳', Empacador: '🛍️', Emergente: '🚨', 'Satélite': '🛰️', Cambista: '💵', Visado: '🔍', 'Fundación': '🤝', SADOFE: '🏷️' };
            const ROL_COLOR = { Cajero: '#2e7d32', Empacador: '#1565c0', Emergente: '#c0392b', 'Satélite': '#7d3c98', Cambista: '#e65100', Visado: '#d35400', 'Fundación': '#6a1b9a', SADOFE: '#00695c' };
            // Colores de barra por letra de turno
            const BARRA = { A: '#f5a623', I: '#57b894', C: '#4a90d9', N: '#8e44ad', '': '#95a5a6' };

            // Batería real de la tienda (total de cajas físicas), cargada de _CONFIG
            let bateria = Number(window.APP_BATERIA) || 0;

            // Reunir personas con turno horario ese día, por rol
            let porRol = {};
            let minH = 24, maxH = 0;
            excelData.forEach((item, idx) => {
                let t = (item.data[colIdx] || '').toString().trim();
                let r = _rangoTurnoDec(t);
                if (!r) return;                       // LIBRE/COMP/VC/etc. no cubren
                minH = Math.min(minH, Math.floor(r.ini));
                maxH = Math.max(maxH, Math.ceil(r.fin));
                let cat = rolCateg(item);
                (porRol[cat] = porRol[cat] || []).push({
                    idx: idx, nombre: item.data[0], turno: t, rango: r, letra: _letraTurnoCober(t),
                    rol: item.rolReal, excluida: !!item.excluida, restriccion: item.restriccion || ''
                });
            });

            if (maxH <= minH) {
                cont.innerHTML = '<div style="color:#7f8c8d; padding:10px;">Nadie tiene turno horario este día (todo LIBRE/COMP/VC).</div>';
                return;
            }
            minH = Math.max(5, minH); maxH = Math.min(24, maxH);
            let horas = [];
            for (let h = minH; h < maxH; h++) horas.push(h);

            // Fecha real del día mostrado (para reglas de sábado/domingo)
            let startVal2 = document.getElementById('fechaInicio').value;
            let startDate2 = new Date(startVal2 ? (startVal2 + 'T12:00:00') : new Date());
            let fechaDia = new Date(startDate2);
            fechaDia.setDate(fechaDia.getDate() + (colIdx - 1));
            let esDomingoDia = fechaDia.getDay() === 0;
            let esSabadoDia = fechaDia.getDay() === 6;

            // Almuerzos (ALM) reales del día — repartidos entre 12:00 y 16:00
            // minimizando el impacto en la cobertura (ver _almuerzosDelDia).
            let almuerzoDe = _almuerzosDelDia(colIdx, esDomingoDia);

            // Estado de una MEDIA HORA (t0 = h o h+0.5): 'work' | 'alm' | 'off'.
            // Así quien entra a las 9:30 NO se marca a las 9:00 (antes contaba
            // la hora completa aunque la persona aún no estuviera).
            function _estadoMedia(p, t0) {
                if (!(p.rango.ini < t0 + 0.5 - 1e-9 && p.rango.fin > t0 + 1e-9)) return 'off';
                let a = almuerzoDe[p.nombre];
                if (a && a[0] < t0 + 0.5 - 1e-9 && a[1] > t0 + 1e-9) return 'alm';
                return 'work';
            }

            let anchoCol = 34;
            let html = '';

            // Datos de la IA para el día (unidades y facturas por hora + tiempo real
            // de transacción): sirven para la carga por cajero y la velocidad de caja.
            let iaUndsHora = {}, iaFactHora = {}, iaMetaHora = {}, iaCapHora = {}, promSegDia = 0;
            if (window.IA_PREDICCIONES && window.IA_PREDICCIONES.length) {
                let iso = fechaDia.getFullYear() + '-' + String(fechaDia.getMonth() + 1).padStart(2, '0') + '-' + String(fechaDia.getDate()).padStart(2, '0');
                window.IA_PREDICCIONES.forEach(r => {
                    if (String(r['Fecha']) === iso) {
                        iaUndsHora[Number(r['Hora'])] = Number(r['Pred_Unds']) || 0;
                        iaFactHora[Number(r['Hora'])] = Number(r['Facturas_Pred']) || 0;
                        iaMetaHora[Number(r['Hora'])] = Number(r['Cajas_Meta']) || 0;
                        iaCapHora[Number(r['Hora'])] = Number(r['Capacidad_Hora']) || 0;  // unds/caja-h reales (pico/valle)
                        if (!promSegDia) promSegDia = Number(r['Prom_Seg_Actual']) || Number(r['Prom_Seg_Transaccion']) || 0;
                    }
                });
            }
            let _cu1 = _cobCajasDia(colIdx, esDomingoDia), _cu2 = _cobCajasSeg(colIdx, esDomingoDia);
            let undsPorCajero = {};
            horas.forEach(h => {
                let u = iaUndsHora[h], c = Math.max(_cu1[h] || 0, _cu2[h] || 0);   // cobertura asentada
                undsPorCajero[h] = (u && c) ? Math.round(u / c) : null;
            });

            // ---- VELOCIDAD DE CAJA de la hora señalada (clic en el encabezado) ----
            let velBox = '';
            if (window.__VEL_HORA != null) {
                let h = window.__VEL_HORA;
                let fact = iaFactHora[h], cajas = Math.max(_cu1[h] || 0, _cu2[h] || 0);
                let cerrar = `<span onclick="verVelocidadHora(${h})" title="Cerrar" style="cursor:pointer; color:#2980b9; float:right; font-weight:bold;">✕</span>`;
                let meta = iaMetaHora[h] || 0;
                if (!fact || !cajas || !promSegDia) {
                    velBox = `<div style="margin:0 0 10px 0; padding:8px 14px; background:#fff3cd; border-left:5px solid #f1c40f; border-radius:6px; font-size:12px; color:#7d6608;">
                        ${cerrar}⏱️ <b>Velocidad ${h}:00</b> — necesito el archivo <b>Analisis_Prioridad_IA.xlsx</b> cargado (facturas y tiempo de transacción) y que haya cajas abiertas a esa hora.</div>`;
                } else {
                    let unds = iaUndsHora[h] || 0;
                    // CAPACIDAD REAL por caja-hora del informe IA (según pico/valle y
                    // entre semana / fin de semana). Es la misma que el Panel muestra
                    // como "Capacidad Real por Segmento". Respaldo: derivarla del tiempo.
                    let capUndsCaja = iaCapHora[h] || 0;
                    if (!capUndsCaja && promSegDia) {
                        let sumU = 0, sumF = 0;
                        horas.forEach(hh => { sumU += iaUndsHora[hh] || 0; sumF += iaFactHora[hh] || 0; });
                        let undsPorFact = sumF > 0 ? sumU / sumF : 0;
                        capUndsCaja = undsPorFact > 0 ? 3600 / (promSegDia / undsPorFact) : 0;
                    }
                    let undsPorCaja = unds / cajas;                                     // unidades/h que atiende
                    let ocup = capUndsCaja > 0 ? undsPorCaja / capUndsCaja : 0;         // carga vs capacidad real
                    let ocupPct = Math.round(ocup * 100);
                    // ¿la demanda exige MÁS cajas de las que el almacén tiene físicamente?
                    let topeFisico = bateria > 0 && meta > bateria;
                    let cls, nota = '';
                    if (ocup <= 0.80) {
                        cls = { t: 'RÁPIDA', s: 'el cliente se atiende sin cola', c: '#27ae60', bg: '#e8f8f0', b: '#27ae60', i: '🟢' };
                        if (ocup <= 0.55) nota = ` <b>Sobran cajas</b> a esta hora: podrías abrir menos.`;
                    } else if (ocup <= 0.92) {
                        cls = { t: 'PROMEDIO', s: 'al límite: cola leve en los picos', c: '#b9770e', bg: '#fff8e1', b: '#f1c40f', i: '🟡' };
                    } else {
                        cls = { t: 'LENTA', s: 'la caja no da abasto: se forma cola', c: '#c0392b', bg: '#fdecea', b: '#e74c3c', i: '🔴' };
                        if (ocup > 1) nota = topeFisico
                            ? ` <b>Tope físico del almacén:</b> la demanda exige ~${meta} cajas, más de las <b>${bateria}</b> que el almacén puede tener; no se cubre abriendo más cajas.`
                            : ` <b>Faltan cajas</b>: la demanda supera la capacidad de esta franja.`;
                    }
                    let metaTxt = topeFisico
                        ? ` · se necesitarían ~<b>${meta}</b> cajas pero el almacén solo tiene <b>${bateria}</b>`
                        : (meta ? ` · la IA recomienda ~<b>${meta}</b>` : '');
                    velBox = `<div style="margin:0 0 10px 0; padding:9px 14px; background:${cls.bg}; border-left:5px solid ${cls.b}; border-radius:6px; font-size:12px; color:#2c3e50;">
                        ${cerrar}${cls.i} <b style="color:${cls.c}; font-size:13px;">Atención ${cls.t} a las ${h}:00</b> — ${cls.s}.${nota}<br>
                        <span style="color:#555;">Cada caja recibe <b>${Math.round(undsPorCaja)}</b> unidades/h y puede con <b>${Math.round(capUndsCaja)}</b> (capacidad real de la franja) → está al <b>${ocupPct}%</b> de su capacidad. Esa hora: <b>${cajas}</b> cajas abiertas${metaTxt} · <b>${Math.round(unds)}</b> unidades/h.</span></div>`;
                }
            }

            ORDEN_ROL.forEach(cat => {
                let gente = porRol[cat];
                if (!gente || !gente.length) return;
                // Ordenar por hora de inicio
                gente.sort((a, b) => a.rango.ini - b.rango.ini || a.nombre.localeCompare(b.nombre));
                let esCajeroBox = (cat === 'Cajero');   // el nº de unidades/caja solo en CAJERO (aligera el render)
                let color = ROL_COLOR[cat] || '#555';
                let cob1 = horas.map(() => 0);   // trabajando en la 1ª media hora (h:00-h:30)
                let cob2 = horas.map(() => 0);   // trabajando en la 2ª media hora (h:30-h+1:00)
                let nExcluidas = gente.filter(p => p.excluida).length;

                html += `<div style="margin-bottom:22px;">
                  <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                    <h4 style="margin:0; font-size:13px; color:${color}; border-left:4px solid ${color}; padding-left:8px;">
                      ${ROL_ICONO[cat] || '👤'} ${cat.toUpperCase()}</h4>
                    <span style="font-size:11px; color:#95a5a6;">${gente.length} persona${gente.length !== 1 ? 's' : ''}${nExcluidas ? ` · <span style="color:#c0392b; font-weight:bold;">${nExcluidas} con restricción (X): no suman</span>` : ''}</span>
                  </div>
                  <div style="overflow-x:auto; border:1px solid #dee2e6; border-radius:6px;">
                  <table style="border-collapse:collapse; font-size:10px; text-align:center; white-space:nowrap;">
                    <thead><tr>
                      <th style="background:${color}; color:#fff; padding:6px 10px; text-align:left; position:sticky; left:0; z-index:2; min-width:180px;">Persona</th>
                      <th style="background:${color}; color:#fff; padding:6px 6px; min-width:100px;">Turno ✏️</th>`;
                horas.forEach(h => {
                    let sel = (window.__VEL_HORA === h);
                    html += `<th onclick="verVelocidadHora(${h})" title="Clic: ver velocidad de atención a las ${h}:00"
                        style="background:${sel ? '#2980b9' : '#34495e'}; color:#fff; padding:5px 0; min-width:${anchoCol}px; cursor:pointer; ${sel ? 'box-shadow:inset 0 -3px 0 #f1c40f;' : ''}">${h}${sel ? ' ⏱️' : ''}</th>`;
                });
                html += `</tr></thead><tbody>`;

                gente.forEach((p, ri) => {
                    const nAttr = (p.nombre || '').replace(/"/g, '&quot;');
                    // Toggle COMPACTO (solo icono) para inhabilitar/habilitar la X
                    const btnInhab = `<button data-n="${nAttr}" onclick="inhabilitarPersona(this.getAttribute('data-n'))"
                                   title="Inhabilitar (poner X): deja de entrar a caja y de sumar"
                                   style="margin-left:6px; padding:0 5px; font-size:11px; line-height:16px; color:#b0b0b0; background:transparent; border:1px solid #e0e0e0; border-radius:4px; cursor:pointer;">⊘</button>`;
                    const btnHab = `<button data-n="${nAttr}" onclick="habilitarPersona(this.getAttribute('data-n'))"
                                   title="Habilitar: quitar la X para que vuelva a sumar"
                                   style="margin-left:6px; padding:0 5px; font-size:11px; line-height:16px; color:#fff; background:#27ae60; border:none; border-radius:4px; cursor:pointer;">↺</button>`;
                    // Restricción de horario "SOLO HASTA LAS HH": incumple si su turno
                    // termina DESPUÉS del límite -> se resalta la fila y las celdas en rojo.
                    let lim = _limiteHasta(p.restriccion);
                    let violaHorario = !p.excluida && lim != null && p.rango.fin > lim + 1e-6;
                    // 'X' en Restricción: fila en ROJO y NO suma; violación de horario: fila resaltada
                    let bg = (p.excluida || violaHorario) ? '#fdecea' : (ri % 2 === 0 ? '#fff' : '#f8f9fa');
                    let restrIcono = (p.restriccion && p.restriccion.toUpperCase().includes('HASTA'))
                        ? ` <span style="font-size:9px; color:${violaHorario ? '#c0392b' : '#e67e22'};" title="${p.restriccion}${violaHorario ? ' — ⚠ INCUMPLE: su turno termina después del límite' : ''}">⏰${violaHorario ? '⚠' : ''}</span>` : '';
                    let nombreHtml = p.excluida
                        ? `<span style="color:#c0392b; font-weight:600;">${p.nombre}</span><span style="font-size:8px; color:#c0392b; margin-left:4px;">✕ no suma</span>${btnHab}`
                        : `<span style="color:${violaHorario ? '#c0392b' : '#2c3e50'}; font-weight:600;">${p.nombre}</span>${restrIcono}${btnInhab}`;
                    // Turno EDITABLE: cambia la línea de tiempo, la matriz y todo lo que dependa
                    let optsTurno = turnosDisponibles.map(t =>
                        `<option value="${t}"${t === p.turno ? ' selected' : ''}>${t || '—'}</option>`).join('');
                    if (turnosDisponibles.indexOf(p.turno) < 0) optsTurno = `<option value="${p.turno}" selected>${p.turno}</option>` + optsTurno;
                    let selTurno = `<select onchange="cambiarTurnoDesdeTimeline(${p.idx},${colIdx},this.value)"
                        style="font-size:11px; font-weight:bold; padding:2px 4px; border:1px solid #dfe3e8; border-radius:4px; background:#fff; color:${violaHorario ? '#c0392b' : '#2c3e50'}; max-width:92px;">${optsTurno}</select>`;
                    html += `<tr style="background:${bg};">
                      <td style="text-align:left; padding:4px 10px; position:sticky; left:0; background:${bg}; border-right:2px solid #dee2e6; min-width:180px; font-size:11px;">${nombreHtml}</td>
                      <td style="padding:3px 6px; border-right:1px solid #eee;">${selTurno}</td>`;
                    horas.forEach((h, hi) => {
                        let e1 = _estadoMedia(p, h), e2 = _estadoMedia(p, h + 0.5);
                        if (!p.excluida) {
                            if (e1 === 'work') cob1[hi]++;
                            if (e2 === 'work') cob2[hi]++;
                        }
                        let colorBarra = p.excluida ? '#e6b0aa' : (violaHorario ? '#e74c3c' : (BARRA[p.letra] || BARRA['']));
                        const colDe = e => e === 'work' ? colorBarra : (e === 'alm' ? '#fdf0d5' : 'transparent');
                        let cel;
                        if (e1 === 'off' && e2 === 'off') {
                            cel = `<td style="padding:0; border:1px solid #f0f0f0;"></td>`;
                        } else if (e1 === 'alm' && e2 === 'alm') {
                            cel = `<td style="padding:0; border:1px solid #eee; background:#fdf0d5; color:#b9770e; font-weight:bold; font-size:8px;">ALM</td>`;
                        } else if (e1 === e2) {   // trabaja toda la hora
                            let nU = (esCajeroBox && undsPorCajero[h] != null)
                                ? `<span style="font-size:8px; color:rgba(255,255,255,0.92); font-weight:600;" title="≈ ${undsPorCajero[h]} unidades por cajero a las ${h}:00 (previsión IA)">${undsPorCajero[h]}</span>` : '';
                            cel = `<td style="padding:0; border:1px solid #eee; background:${colorBarra}; text-align:center; line-height:1;">${nU}</td>`;
                        } else {
                            // media hora distinta: celda partida (izquierda = h:00-h:30)
                            let txt = (e1 === 'alm' || e2 === 'alm') ? 'ALM' : '';
                            let titulo = `${h}:00-${h}:30 ${e1 === 'work' ? 'trabaja' : (e1 === 'alm' ? 'almuerzo' : 'no está')} · ${h}:30-${h + 1}:00 ${e2 === 'work' ? 'trabaja' : (e2 === 'alm' ? 'almuerzo' : 'no está')}`;
                            cel = `<td title="${titulo}" style="padding:0; border:1px solid #eee; background:linear-gradient(90deg, ${colDe(e1)} 0 50%, ${colDe(e2)} 50% 100%); color:#b9770e; font-weight:bold; font-size:8px;">${txt}</td>`;
                        }
                        html += cel;
                    });
                    html += `</tr>`;
                });

                // Fila de cobertura total por MEDIA hora + % (sobre la BATERÍA real).
                // Si las dos medias horas difieren se muestran ambas (h:00-30 / h:30-00):
                // así quien entra a las 9:30 ya no infla la cobertura de las 9:00.
                let pico = Math.max(1, ...cob1, ...cob2);
                let usaBateria = bateria > 0 && (cat === 'Cajero' || cat === 'Emergente');
                let denom = usaBateria ? bateria : pico;
                let etiquetaPct = usaBateria ? `% de la BATERÍA (${bateria} cajas)` : `% sobre el pico (${pico})`;
                html += `<tr style="background:${color}; color:#fff; font-weight:bold;">
                    <td style="padding:5px 10px; text-align:left; position:sticky; left:0; background:${color};">Cajas cubiertas (:00 / :30)</td>
                    <td style="background:${color};"></td>`;
                horas.forEach((h, hi) => {
                    let n1 = cob1[hi], n2 = cob2[hi];
                    let celTxt = (n1 === n2) ? (n1 || '') : `${n1}<span style="opacity:.65;">/</span>${n2}`;
                    html += `<td title="${h}:00-${h}:30 → ${n1} · ${h}:30-${h + 1}:00 → ${n2}" style="padding:4px 0; font-size:${n1 === n2 ? 10 : 9}px;">${celTxt}</td>`;
                });
                html += `</tr><tr style="background:#eaf0fb; color:#2c3e50; font-weight:bold;">
                    <td style="padding:4px 10px; text-align:left; position:sticky; left:0; background:#eaf0fb;">${etiquetaPct}</td>
                    <td style="background:#eaf0fb;"></td>`;
                horas.forEach((h, hi) => {
                    let nMin = Math.min(cob1[hi], cob2[hi]);
                    let pct = Math.round(nMin / denom * 100);
                    let colPct = usaBateria ? (pct >= 80 ? '#27ae60' : (pct >= 60 ? '#e67e22' : '#c0392b')) : '#2c3e50';
                    html += `<td title="mínimo de las dos medias horas" style="padding:3px 0; font-size:9px; color:${colPct};">${pct}%</td>`;
                });
                html += `</tr></tbody></table></div></div>`;
            });

            // ===== CHEQUEO DIARIO DEL PERFIL DE COBERTURA (rige TODOS los días) =====
            //   1) 8:00 -> 30% · 2) cierre del último turno (21:30) -> 60% · 3) 11-13 -> 100% (mín. 85%)
            //   Domingo además: mínimo 80% entre 13:00 y 20:00.
            let bannerDomingo = '';
            const NOMBRES_DIA = ['DOMINGO', 'LUNES', 'MARTES', 'MIÉRCOLES', 'JUEVES', 'VIERNES', 'SÁBADO'];
            let nombreDia = NOMBRES_DIA[fechaDia.getDay()];
            if (bateria > 0) {
                let ev = _evaluarPerfilDia(colIdx, esDomingoDia, bateria);
                if (ev.fallos.length) {
                    let detalle = ev.fallos.map(f => {
                        let req = f.n + f.faltan;   // cajas requeridas por el ancla de esa hora
                        return `<b>${f.h}:00</b> → ${f.n}/${bateria} cajas (${Math.round(f.pct * 100)}%): faltan ${f.faltan} para el ${Math.round(f.obj * 100)}% [Ancla AKB30 ${f.h}:00 — ${req} cajas abiertas]`;
                    });
                    bannerDomingo = `<div style="margin:0 0 14px 0; padding:11px 15px; background:#fff8e1; border:1px solid #f0d264; border-left:5px solid #f1c40f; border-radius:8px;">
                        <div style="font-size:12px; color:#7d6608;">Sumando Cajeros + Emergentes (sin personas con X):<br>${detalle.join('<br>')}</div>
                    </div>`;
                } else {
                    bannerDomingo = `<div style="margin:0 0 14px 0; padding:10px 16px; background:#e8f8f0; border:1px solid #27ae60; border-left:6px solid #27ae60; border-radius:8px; font-size:12px; color:#1e8449;">
                        ✅ ${nombreDia}: se cumple el perfil de cobertura de la batería (${bateria} cajas): ${_descPerfil()}${!_tablaAnclasTienda() && esDomingoDia ? ' · ≥80% en la tarde' : ''}.</div>`;
                }
            } else {
                bannerDomingo = `<div style="margin:0 0 14px 0; padding:10px 16px; background:#fff3cd; border-left:5px solid #f1c40f; border-radius:6px; font-size:12px; color:#7d6608;">
                    ⚠️ No hay BATERÍA configurada para esta tienda: genera la malla desde el Python indicando el total de cajas para activar el chequeo diario del perfil (1º 30% @8:00 · 2º 60% @cierre 21:30 · 3º 100% @11-13, mín. 85%).</div>`;
            }

            // ---- Radar del día: VACACIONES y personas que NO pueden entrar a caja ----
            let infoDia = '';
            let deVacaciones = [];
            (excelData || []).forEach(it => {
                const v = (it.data[colIdx] || '').toString().trim().toUpperCase();
                if (v === 'VC') deVacaciones.push((it.data[0] || '').toString().trim());
            });
            const piPl = _plantaInfo();
            if (piPl) {
                Object.keys(piPl.vacacionesDe).forEach(nom => {
                    const [d1, d2] = piPl.vacacionesDe[nom];
                    if (fechaDia >= d1 && fechaDia <= d2 && deVacaciones.indexOf(nom) < 0) deVacaciones.push(nom);
                });
            }
            if (deVacaciones.length) {
                infoDia += `<div style="margin:0 0 10px 0; padding:8px 14px; background:#eaf2f8; border-left:5px solid #2980b9; border-radius:6px; font-size:12px; color:#1a5276;">
                    🏖️ <b>EN VACACIONES este día (${deVacaciones.length}):</b> ${deVacaciones.join(' · ')} — no cuentan en la cobertura; tenlas en el radar para la planta del día.</div>`;
            }
            const bloqueadas = (excelData || []).filter(it => it.excluida).map(it => (it.data[0] || '').toString().trim());
            if (bloqueadas.length) {
                infoDia += `<div style="margin:0 0 10px 0; padding:8px 14px; background:#fdecea; border-left:5px solid #c0392b; border-radius:6px; font-size:12px; color:#7b241c;">
                    ⛔ <b>NO pueden entrar a caja (X en la columna Restricción del Sheets, ${bloqueadas.length}):</b> ${bloqueadas.join(' · ')} — aparecen en rojo y no suman; usa "✅ Habilitar" cuando termine su restricción.</div>`;
            }

            cont.innerHTML = velBox + infoDia + bannerDomingo + (html || '<div style="color:#7f8c8d; padding:10px;">Sin personal con turno horario este día.</div>');
        }

        // HABILITAR a una persona con 'X': quita la restricción para que vuelva
        // a sumar en la cobertura de su rol (cuando su restricción ya terminó).
        // El cambio queda en la matriz en memoria; para que persista hay que
        // pulsar "Guardar cambios en Sheets" (SCHED). Para que las PRÓXIMAS
        // mallas también la incluyan hay que borrar la X en PLANTA_<tienda>.
        window.habilitarPersona = function (nombre) {
            if (!nombre) return;
            if (!confirm(`¿Habilitar a ${nombre}?\n\nSe quita su restricción (X) y vuelve a sumar en la cobertura de su rol.`)) return;
            let tocado = false;
            (excelData || []).forEach(it => {
                if ((it.data[0] || '').toString().trim() === nombre.trim()) {
                    it.excluida = false;
                    it.restriccion = '';
                    tocado = true;
                }
            });
            // Reflejarlo también en la matriz original (columna Restricción),
            // que es la que se reconstruye al guardar / recombinar.
            if (matrizOriginal && matrizOriginal.length > 1) {
                let head = matrizOriginal[0];
                let iN = head.findIndex(h => normalizarTexto(h) === 'NOMBRE');
                let iR = head.findIndex((h, i) => i > 0 && normalizarTexto(h).includes('RESTRIC'));
                if (iN >= 0 && iR >= 0) {
                    for (let r = 1; r < matrizOriginal.length; r++) {
                        if ((matrizOriginal[r][iN] || '').toString().trim() === nombre.trim()) {
                            matrizOriginal[r][iR] = '';
                        }
                    }
                }
            }
            if (!tocado) { alert('No se encontró a "' + nombre + '" en la matriz cargada.'); return; }
            // quitar la X también de la copia en memoria de la pestaña PLANTA
            // (para que la sincronización no vuelva a marcarla en esta sesión)
            if (window.PLANTA_DATA && window.PLANTA_DATA.length > 1) {
                const headP = window.PLANTA_DATA[0].map(x => (x || '').toString().toUpperCase());
                const iN = headP.findIndex(x => x.includes('NOMBRE'));
                const iR = headP.findIndex(x => x.includes('RESTRIC'));
                if (iN >= 0 && iR >= 0) {
                    for (let r = 1; r < window.PLANTA_DATA.length; r++) {
                        if ((window.PLANTA_DATA[r][iN] || '').toString().trim() === nombre.trim() &&
                            (window.PLANTA_DATA[r][iR] || '').toString().trim().toUpperCase() === 'X') {
                            window.PLANTA_DATA[r][iR] = '';
                        }
                    }
                }
            }
            renderLineaTiempoCobertura();
            aplicarFiltrosGlobales();
            alert(`✅ ${nombre} quedó HABILITADA/O y ya suma en la cobertura de su rol.\n\n` +
                  `· Pulsa "Guardar cambios en Sheets" para que quede guardado en la malla (SCHED).\n` +
                  `· Para que las PRÓXIMAS mallas también la incluyan, borra la X en la pestaña PLANTA de la tienda en Google Sheets.`);
            if (typeof renderCumplimiento === 'function') renderCumplimiento();
        };

        // INHABILITAR a una persona: le pone 'X' en Restricción para que NO entre
        // a caja y no sume en su rol (misma columna del Sheets que la 'X').
        window.inhabilitarPersona = function (nombre) {
            if (!nombre) return;
            if (!confirm(`¿Inhabilitar a ${nombre}?\n\nSe le pone la restricción (X): deja de entrar a caja y no suma en la cobertura de su rol.`)) return;
            let tocado = false;
            (excelData || []).forEach(it => {
                if ((it.data[0] || '').toString().trim() === nombre.trim()) {
                    it.excluida = true;
                    it.restriccion = 'X';
                    tocado = true;
                }
            });
            if (matrizOriginal && matrizOriginal.length > 1) {
                let head = matrizOriginal[0];
                let iN = head.findIndex(h => normalizarTexto(h) === 'NOMBRE');
                let iR = head.findIndex((h, i) => i > 0 && normalizarTexto(h).includes('RESTRIC'));
                if (iN >= 0 && iR >= 0) {
                    for (let r = 1; r < matrizOriginal.length; r++) {
                        if ((matrizOriginal[r][iN] || '').toString().trim() === nombre.trim()) matrizOriginal[r][iR] = 'X';
                    }
                }
            }
            if (window.PLANTA_DATA && window.PLANTA_DATA.length > 1) {
                const headP = window.PLANTA_DATA[0].map(x => (x || '').toString().toUpperCase());
                const iN = headP.findIndex(x => x.includes('NOMBRE'));
                const iR = headP.findIndex(x => x.includes('RESTRIC'));
                if (iN >= 0 && iR >= 0) {
                    for (let r = 1; r < window.PLANTA_DATA.length; r++) {
                        if ((window.PLANTA_DATA[r][iN] || '').toString().trim() === nombre.trim()) window.PLANTA_DATA[r][iR] = 'X';
                    }
                }
            }
            if (!tocado) { alert('No se encontró a "' + nombre + '" en la matriz cargada.'); return; }
            renderLineaTiempoCobertura();
            aplicarFiltrosGlobales();
            alert(`⛔ ${nombre} quedó INHABILITADA/O (X): ya no entra a caja ni suma.\n\n` +
                  `· Pulsa "Guardar cambios en Sheets" para dejarlo guardado en la malla (SCHED).\n` +
                  `· Para que aplique a las PRÓXIMAS mallas, pon la X en la pestaña PLANTA de la tienda.`);
            if (typeof renderCumplimiento === 'function') renderCumplimiento();
        };

        // Límite horario de una restricción "SOLO HASTA LAS HH(:MM)" -> hora decimal.
        function _limiteHasta(restr) {
            let s = (restr || '').toString().toUpperCase();
            if (!s.includes('HASTA')) return null;
            let m = s.match(/HASTA\s+(?:LAS?\s+)?(\d{1,2})(?:[:.,](\d{1,2}))?/);
            if (!m) return null;
            let mm = m[2] ? (m[2].length === 1 ? parseInt(m[2]) * 10 : parseInt(m[2])) : 0;
            return parseInt(m[1]) + mm / 60;
        }

        // ===== INFORME: RESTRICCIONES DE HORARIO (incumplimientos + sugerencia) =====
        // Recorre TODA la malla y lista, por persona con restricción "HASTA las X",
        // los días en que su turno TERMINA después de su límite (incumple). Para
        // cada día propone un turno que SÍ cumple y que, según las ventas (unds) de
        // ese día, cubre las horas de mayor demanda que la persona puede atender.
        function _decAHM(x) {
            let h = Math.floor(x + 1e-9), m = Math.round((x - h) * 60);
            if (m === 60) { h++; m = 0; }
            return h + ':' + String(m).padStart(2, '0');
        }

        function _poblarMesesRestr() {
            let sel = document.getElementById('restrMes');
            if (!sel || sel.dataset.lleno === '1' || !headers.length) return;
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            let meses = new Set();
            headers.forEach((hd, i) => {
                if (hd === 'Total Horas Periodo') return;
                let d = new Date(startDate); d.setDate(d.getDate() + i);
                meses.add(d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'));
            });
            const NOM = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
            let opts = '<option value="TODOS">Todos los meses</option>';
            [...meses].sort().forEach(mk => {
                let [a, m] = mk.split('-');
                opts += `<option value="${mk}">${NOM[parseInt(m) - 1]} ${a}</option>`;
            });
            sel.innerHTML = opts;
            sel.dataset.lleno = '1';
        }

        function renderRestriccionesHorario() {
            const cont = document.getElementById('restrHorarioCont');
            if (!cont) return;
            if (!excelData || !excelData.length || !headers.length) {
                cont.innerHTML = '<div style="color:#7f8c8d; padding:10px;">Carga la Matriz de Turnos para ver este informe.</div>';
                return;
            }
            _poblarMesesRestr();
            let mesSel = (document.getElementById('restrMes') || {}).value || 'TODOS';
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            const fIso = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
            const DIAS = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];

            // Demanda (unds) por fecha|hora y total por día, de la IA
            let iaHora = {}, iaTotalDia = {};
            (window.IA_PREDICCIONES || []).forEach(r => {
                let u = Number(r['Pred_Unds']) || 0;
                iaHora[r['Fecha'] + '|' + Number(r['Hora'])] = u;
                iaTotalDia[r['Fecha']] = (iaTotalDia[r['Fecha']] || 0) + u;
            });
            const tablaAncla = (typeof _tablaAnclasTienda === 'function') ? _tablaAnclasTienda() : null;
            const hayIA = !!(window.IA_PREDICCIONES && window.IA_PREDICCIONES.length);
            function demHora(iso, h) {
                let v = iaHora[iso + '|' + h];
                if (v != null && v > 0) return v;
                return tablaAncla ? (tablaAncla[h] || 0) : 0;
            }

            // Turnos disponibles con su rango horario
            let turnosRango = turnosDisponibles.map(t => ({ t: t, r: _rangoTurnoDec(t) })).filter(x => x.r);

            // Personas con restricción "HASTA las X"
            let personas = excelData.map((item, idx) => ({
                idx: idx, nombre: item.data[0], restr: (item.restriccion || '').toString(),
                lim: _limiteHasta(item.restriccion), item: item
            })).filter(p => p.lim != null);

            let totalDiasIncumpl = 0, personasConIncumpl = 0, bloques = '';

            personas.forEach(p => {
                let filas = [];
                headers.forEach((hd, i) => {
                    if (hd === 'Total Horas Periodo') return;
                    let colIdx = i + 1;
                    let d = new Date(startDate); d.setDate(d.getDate() + i);
                    if (mesSel !== 'TODOS') {
                        let mk = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
                        if (mk !== mesSel) return;
                    }
                    let t = (p.item.data[colIdx] || '').toString().trim();
                    let r = _rangoTurnoDec(t);
                    if (!r) return;                         // LIBRE/COMP/VC: no aplica
                    if (r.fin <= p.lim + 1e-6) return;      // cumple su límite
                    // --- INCUMPLE: buscar turno que cumpla y sirva a la demanda del día ---
                    let iso = fIso(d);
                    let dur = r.fin - r.ini;
                    let demV = rr => { let s = 0; for (let h = Math.floor(rr.ini); h < Math.ceil(rr.fin); h++) s += demHora(iso, h); return s; };
                    let cands = turnosRango.filter(x => x.r.fin <= p.lim + 1e-6).map(x => ({
                        t: x.t, r: x.r, score: demV(x.r), durDiff: Math.abs((x.r.fin - x.r.ini) - dur)
                    }));
                    // 1º misma duración (±0.25h) para no cambiarle las horas de la semana,
                    // 2º mayor demanda cubierta, 3º duración más parecida
                    cands.sort((a, b) =>
                        (a.durDiff > 0.25 ? 1 : 0) - (b.durDiff > 0.25 ? 1 : 0) ||
                        b.score - a.score || a.durDiff - b.durDiff);
                    let sug = cands[0] || null;
                    // Hora de MAYOR demanda que la persona puede cubrir (apertura..límite)
                    let picoH = null, picoD = -1;
                    for (let h = 8; h < Math.ceil(p.lim); h++) { let dd = demHora(iso, h); if (dd > picoD) { picoD = dd; picoH = h; } }
                    filas.push({ iso: iso, dow: d.getDay(), turnoAct: t, finAct: r.fin, dur: dur, sug: sug, picoH: picoH, ventas: iaTotalDia[iso] || 0 });
                });

                if (!filas.length) return;
                personasConIncumpl++;
                totalDiasIncumpl += filas.length;
                let rows = filas.map(f => {
                    let fecha = f.iso.split('-').reverse().join('/').slice(0, 5);
                    let sugTxt = f.sug
                        ? `<b style="color:#1e7e34;">${f.sug.t}</b> <span style="color:#7f8c8d;">(term. ${_decAHM(f.sug.r.fin)}${f.sug.durDiff > 0.25 ? `, dura ${(f.sug.r.fin - f.sug.r.ini).toFixed(1).replace('.0', '')}h vs ${f.dur.toFixed(1).replace('.0', '')}h` : ''})</span>`
                        : `<span style="color:#c0392b;">— sin turno que cumpla (requiere turno especial hasta ${_decAHM(p.lim)})</span>`;
                    let picoTxt = f.picoH != null ? `${f.picoH}:00` : '—';
                    let ventasTxt = f.ventas ? Math.round(f.ventas).toLocaleString('es') + ' unds' : '—';
                    return `<tr>
                        <td style="padding:5px 8px; border-bottom:1px solid #eee;">${DIAS[f.dow]} ${fecha}</td>
                        <td style="padding:5px 8px; border-bottom:1px solid #eee; color:#c0392b;">${f.turnoAct} <span style="color:#e67e22;">⚠ term. ${_decAHM(f.finAct)}</span></td>
                        <td style="padding:5px 8px; border-bottom:1px solid #eee;">${sugTxt}</td>
                        <td style="padding:5px 8px; border-bottom:1px solid #eee; text-align:center;">${picoTxt}</td>
                        <td style="padding:5px 8px; border-bottom:1px solid #eee; text-align:right;">${ventasTxt}</td>
                    </tr>`;
                }).join('');
                bloques += `<div style="margin-bottom:18px; border:1px solid #e6e6e6; border-radius:8px; overflow:hidden;">
                    <div style="background:#fef5e7; padding:8px 12px; border-left:5px solid #e67e22;">
                        <b style="color:#a04000;">⏰ ${p.nombre}</b>
                        <span style="color:#7f8c8d; font-size:12px;"> · restricción: ${p.restr || 'HASTA ' + _decAHM(p.lim)} · ${filas.length} día(s) con incumplimiento</span>
                    </div>
                    <table style="width:100%; border-collapse:collapse; font-size:12px;">
                        <thead><tr style="background:#f8f9fa; color:#34495e;">
                            <th style="padding:6px 8px; text-align:left;">Día</th>
                            <th style="padding:6px 8px; text-align:left;">Turno actual (incumple)</th>
                            <th style="padding:6px 8px; text-align:left;">Turno sugerido (cumple)</th>
                            <th style="padding:6px 8px; text-align:center;" title="Hora de mayor venta que puede cubrir dentro de su límite">Sirve más a las</th>
                            <th style="padding:6px 8px; text-align:right;">Ventas del día</th>
                        </tr></thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>`;
            });

            let resumen = `<div style="margin-bottom:14px; padding:10px 14px; border-radius:8px; background:${totalDiasIncumpl ? '#fdecea' : '#e8f8f0'}; border-left:5px solid ${totalDiasIncumpl ? '#e74c3c' : '#27ae60'}; font-size:13px; color:#2c3e50;">
                <b>${personas.length}</b> persona(s) con restricción horaria · <b>${personasConIncumpl}</b> con incumplimientos · <b>${totalDiasIncumpl}</b> día(s) por corregir.
                ${!hayIA ? '<br><span style="color:#7d6608;">⚠ Sin el archivo <b>Analisis_Prioridad_IA.xlsx</b>, la sugerencia usa los mínimos de la tienda en vez de las ventas reales.</span>' : ''}</div>`;

            if (!personas.length) {
                cont.innerHTML = '<div style="padding:12px 16px; background:#eaf2f8; border-left:5px solid #2980b9; border-radius:6px; font-size:13px; color:#1a5276;">No hay personas con restricción de horario (columna "Restricción" con texto tipo <b>HASTA las 14:00</b>) en la Matriz cargada.</div>';
                return;
            }
            cont.innerHTML = resumen + (bloques || '<div style="padding:12px 16px; background:#e8f8f0; border-left:5px solid #27ae60; border-radius:6px; font-size:13px; color:#1e6b3a;">🎉 Todas las personas con restricción cumplen su horario en el período/mes seleccionado.</div>');
        }

        // ===== INFORME: CUMPLIMIENTO DEL PERFIL DE COBERTURA (todos los días) =====
        // Recorre TODOS los días de la malla y lista los que NO cumplen el perfil
        // (1º 30% @8:00 · 2º 60% @cierre 21:30 · 3º 100% @11-13 mín. 85% · dom ≥80% tarde),
        // con el detalle de cada hora incumplida y cuántas cajas faltan.

        // ===== EVALUACIÓN DE SERVICIO HORA POR HORA (gráfico circular) =====
        // Clasifica cada celda (día, hora) según el MÍNIMO de cajas de la tienda:
        //   >= mínimo -> EXCELENTE · 89%-99% -> BUENA · 80%-89% y la IA cree que
        //   con esa cantidad hay buen servicio -> BUENA · por debajo -> NO CUMPLE.
        function _clasifCelda(n, req, iaMeta) {
            const ratio = n / req;
            if (ratio >= 1 - 1e-9) return 'exc';
            if (ratio >= 0.89) return 'bue';
            if (ratio >= 0.80 && iaMeta != null && iaMeta <= n) return 'bue';
            return 'no';
        }

        function _evalGaugeCumpl(mesSel) {
            const tabla = _tablaAnclasTienda();
            const tipoSel = window.__CUMPL_TIPO || '';
            let cats = { exc: 0, bue: 0, no: 0 };
            let porHora = {}, mesesSet = new Set();
            const horasOrden = tabla ? Object.keys(tabla).map(Number).sort((a, b) => a - b) : [];
            horasOrden.forEach(h => porHora[h] = { exc: 0, bue: 0, no: 0, sum: 0, min: Infinity, max: 0, n: 0, req: tabla[h] });
            if (!tabla) return { cats, porHora, horasOrden, total: 0, meses: [], sinTabla: true };
            let iaMap = {};
            (window.IA_PREDICCIONES || []).forEach(r => { iaMap[r['Fecha'] + '|' + Number(r['Hora'])] = r; });
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            const fIso = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
            headers.forEach((h, i) => {
                if (h === 'Total Horas Periodo') return;
                let colIdx = i + 1;
                if (!excelData.some(it => _rangoTurnoDec((it.data[colIdx] || '').toString().trim()))) return;
                let d = new Date(startDate); d.setDate(d.getDate() + i);
                let mesKey = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
                mesesSet.add(mesKey);
                if (mesSel && mesKey !== mesSel) return;
                let wd = d.getDay();
                let tipoDia = wd === 0 ? 'DOM' : (wd === 6 ? 'SAB' : 'LV');
                if (tipoSel && tipoDia !== tipoSel) return;
                let esDom = wd === 0;
                let cob = _cobCajasDia(colIdx, esDom);
                let iso = fIso(d);
                horasOrden.forEach(hh => {
                    const req = tabla[hh]; if (req <= 0) return;
                    const n = cob[hh] || 0;
                    let iaMeta = null;
                    const r = iaMap[iso + '|' + hh];
                    if (r) iaMeta = Number(r['Cajas_Meta_Demanda'] !== '' && r['Cajas_Meta_Demanda'] !== undefined ? r['Cajas_Meta_Demanda'] : r['Cajas_Meta']) || null;
                    const cl = _clasifCelda(n, req, iaMeta);
                    cats[cl]++;
                    const ph = porHora[hh];
                    ph[cl]++; ph.sum += n; ph.n++; ph.min = Math.min(ph.min, n); ph.max = Math.max(ph.max, n);
                });
            });
            return { cats, porHora, horasOrden, total: cats.exc + cats.bue + cats.no, meses: [...mesesSet].sort(), sinTabla: false };
        }

        const _MESNOM = ['', 'ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
        function _mesLabel(k) { const p = k.split('-'); return _MESNOM[parseInt(p[1])] + ' ' + p[0]; }

        function _gaugeCumplHTML(g) {
            if (g.sinTabla) return '';
            const mesActual = window.__CUMPL_MES || '';
            let opts = '<option value="">Todos los meses</option>' +
                g.meses.map(k => `<option value="${k}" ${k === mesActual ? 'selected' : ''}>${_mesLabel(k)}</option>`).join('');
            const T = g.total || 1;
            const pOK = Math.round((g.cats.exc + g.cats.bue) / T * 100);
            const pct = x => Math.round(x / T * 100);
            let filas = g.horasOrden.map(h => {
                const ph = g.porHora[h];
                const prom = ph.n ? (ph.sum / ph.n) : 0;
                const okPct = ph.n ? Math.round((ph.exc + ph.bue) / ph.n * 100) : 0;
                const col = okPct >= 100 ? '#27ae60' : (okPct >= 80 ? '#f39c12' : '#c0392b');
                const rango = ph.n ? `${ph.min === Infinity ? 0 : ph.min}–${ph.max}` : '—';
                return `<tr>
                    <td style="padding:3px 8px; font-weight:bold;">${h}:00</td>
                    <td style="padding:3px 8px; text-align:center;">${ph.req}</td>
                    <td style="padding:3px 8px; text-align:center;">${prom.toFixed(1)}</td>
                    <td style="padding:3px 8px; text-align:center; color:#7f8c8d;">${rango}</td>
                    <td style="padding:3px 8px; text-align:center; color:#27ae60;">${ph.exc}</td>
                    <td style="padding:3px 8px; text-align:center; color:#f39c12;">${ph.bue}</td>
                    <td style="padding:3px 8px; text-align:center; color:#c0392b;">${ph.no}</td>
                    <td style="padding:2px 8px;"><div style="height:10px; border-radius:5px; background:${col}; width:${Math.max(6, okPct)}%; min-width:14px;" title="${okPct}% de los días cumplen (excelente o buena)"></div></td>
                  </tr>`;
            }).join('');
            const chip = (c, txt, v) => `<span style="display:inline-flex; align-items:center; gap:6px; font-size:12px; margin:0 8px;">
                <span style="width:12px; height:12px; border-radius:3px; background:${c};"></span> ${txt}: <b>${v}</b> celdas (${pct(v)}%)</span>`;
            return `<div style="background:#fff; border:1px solid #dee2e6; border-radius:10px; padding:16px 18px; margin-bottom:18px;">
                <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:8px;">
                    <h3 style="margin:0; color:#1a5276; font-size:16px;">🎯 Evaluación de servicio BATERÍA DE CAJAS POR hora — ${(window.APP_TIENDA || '').toUpperCase()}</h3>
                    <div style="display:flex; gap:10px; flex-wrap:wrap;">
                      <label style="font-size:13px;"><b>Días:</b> <select id="cumplGaugeTipo" onchange="filtrarCumplTipo()" style="padding:4px; font-size:13px;">
                        <option value="" ${!(window.__CUMPL_TIPO) ? 'selected' : ''}>Todos</option>
                        <option value="LV" ${window.__CUMPL_TIPO === 'LV' ? 'selected' : ''}>Lunes a Viernes</option>
                        <option value="SAB" ${window.__CUMPL_TIPO === 'SAB' ? 'selected' : ''}>Sábado</option>
                        <option value="DOM" ${window.__CUMPL_TIPO === 'DOM' ? 'selected' : ''}>Únicamente domingos</option>
                      </select></label>
                      <label style="font-size:13px;"><b>Mes:</b> <select id="cumplGaugeMes" onchange="filtrarCumplMes()" style="padding:4px; font-size:13px;">${opts}</select></label>
                    </div>
                </div>
                <div style="text-align:center; margin:6px 0 10px 0;">
                    <div style="position:relative; width:300px; height:300px; margin:0 auto;">
                        <canvas id="cumplGaugeCanvas"></canvas>
                        <div style="position:absolute; inset:0; display:flex; flex-direction:column; align-items:center; justify-content:center; pointer-events:none;">
                            <div style="font-size:46px; font-weight:bold; color:${pOK >= 90 ? '#27ae60' : (pOK >= 75 ? '#f39c12' : '#c0392b')};">${pOK}%</div>
                            <div style="font-size:12px; color:#7f8c8d;">servicio bueno o excelente</div>
                            <div style="font-size:11px; color:#95a5a6;">${T} evaluaciones hora·día</div>
                        </div>
                    </div>
                    <div style="margin-top:8px;">
                        ${chip('#27ae60', 'Excelente (≥ mínimo)', g.cats.exc)}
                        ${chip('#f39c12', 'Buena (80-99%)', g.cats.bue)}
                        ${chip('#c0392b', 'No cumple (&lt;80%)', g.cats.no)}
                    </div>
                </div>
                <div style="overflow-x:auto; margin-top:8px;"><table style="border-collapse:collapse; font-size:12px; width:100%; background:#fff;">
                    <thead><tr style="background:#1a5276; color:#fff; font-size:11px;">
                        <th style="padding:4px 8px;">Hora</th><th style="padding:4px 8px;">Mín. cajas</th>
                        <th style="padding:4px 8px;">Prom. cajeros</th><th style="padding:4px 8px;">Rango</th>
                        <th style="padding:4px 8px;">Excel.</th><th style="padding:4px 8px;">Buena</th>
                        <th style="padding:4px 8px;">No cumple</th><th style="padding:4px 8px;">% días OK</th>
                    </tr></thead><tbody>${filas}</tbody></table></div>
                <div style="font-size:11px; color:#7f8c8d; margin-top:6px;">Cada evaluación es una casilla (día × hora). El % central es cuántas quedaron en <b>Excelente</b> o <b>Buena</b>. La banda 80-89% cuenta como Buena solo si la IA (Analisis_Prioridad_IA.xlsx) considera suficiente esa cantidad para la demanda de esa hora.</div>
            </div>`;
        }

        function _dibujarGaugeCumpl(g) {
            const cv = document.getElementById('cumplGaugeCanvas');
            if (!cv || typeof Chart === 'undefined' || g.sinTabla) return;
            if (window.__cumplGaugeChart) { try { window.__cumplGaugeChart.destroy(); } catch (e) {} }
            window.__cumplGaugeChart = new Chart(cv.getContext('2d'), {
                type: 'doughnut',
                data: {
                    labels: ['Excelente', 'Buena', 'No cumple'],
                    datasets: [{
                        data: [g.cats.exc, g.cats.bue, g.cats.no],
                        backgroundColor: ['#27ae60', '#f39c12', '#c0392b'],
                        borderWidth: 2, borderColor: '#fff'
                    }]
                },
                options: {
                    responsive: true, maintainAspectRatio: false, cutout: '68%',
                    plugins: { legend: { display: false } }
                }
            });
        }

        window.filtrarCumplMes = function () {
            const sel = document.getElementById('cumplGaugeMes');
            window.__CUMPL_MES = sel ? sel.value : '';
            renderCumplimiento();
        };
        window.filtrarCumplTipo = function () {
            const sel = document.getElementById('cumplGaugeTipo');
            window.__CUMPL_TIPO = sel ? sel.value : '';
            renderCumplimiento();
        };

        function renderCumplimiento() {
            let cont = document.getElementById('cumplimientoContainer');
            if (!cont) return;
            if (!excelData || excelData.length === 0) {
                cont.innerHTML = '<div style="color:#7f8c8d; padding:10px;">Carga la Matriz de Turnos para ver el cumplimiento del perfil de cobertura.</div>';
                return;
            }
            let bateria = Number(window.APP_BATERIA) || 0;
            if (!(bateria > 0)) {
                cont.innerHTML = `<div style="padding:12px 16px; background:#fff3cd; border-left:5px solid #f1c40f; border-radius:6px; font-size:13px; color:#7d6608;">
                    ⚠️ No hay BATERÍA configurada para esta tienda: genera la malla desde el Python indicando el total de cajas para activar este informe.</div>`;
                return;
            }
            const NOMBRES_DIA = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());

            // Filtros del gráfico (Días y Mes) que también afectan el detalle día a día
            const tipoSelC = window.__CUMPL_TIPO || '';
            const mesSelC = window.__CUMPL_MES || '';
            const tiendaC = (window.APP_TIENDA || '').toUpperCase();
            _minOverrideActivo = (MIN_HORA_OVERRIDE[tiendaC] && tipoSelC) ? (MIN_HORA_OVERRIDE[tiendaC][tipoSelC] || null) : null;

            let totalDias = 0, diasMal = [], diasOrg = [], diasConUltimo = 0;
            headers.forEach((h, i) => {
                if (h === 'Total Horas Periodo') return;
                let colIdx = i + 1;
                // ¿alguien trabaja este día? (si todo es LIBRE/COMP no se evalúa)
                let hayTurnos = excelData.some(it => _rangoTurnoDec((it.data[colIdx] || '').toString().trim()));
                if (!hayTurnos) return;
                let d = new Date(startDate);
                d.setDate(d.getDate() + i);
                // Respetar los filtros de Días (L-V/Sáb/Dom) y Mes del gráfico
                let wdF = d.getDay();
                let tipoF = wdF === 0 ? 'DOM' : (wdF === 6 ? 'SAB' : 'LV');
                if (tipoSelC && tipoF !== tipoSelC) return;
                let mesKeyF = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
                if (mesSelC && mesKeyF !== mesSelC) return;
                totalDias++;
                let esDom = d.getDay() === 0;
                let ev = _evaluarPerfilDia(colIdx, esDom, bateria);
                if (ev.fallos.length) {
                    let jf = _justifPersonalDia(colIdx, esDom, bateria);
                    let vacCaja = _vacCajaDia(colIdx, d);
                    // SOLO ORGANIZACIÓN: el personal alcanza (faltan 0) y NO hay
                    // cajeros/emergentes en vacaciones -> no cuenta como malo.
                    let esOrg = (jf.faltan === 0) && (vacCaja.length === 0);
                    let reg = { colIdx: colIdx, header: h, fecha: d, esDom: esDom, ev: ev, jf: jf, vacCaja: vacCaja };
                    if (esOrg) { diasOrg.push(reg); }
                    else { diasMal.push(reg); if (ev.falloUltimo) diasConUltimo++; }
                }
            });

            // --- PROMEDIOS DE FALTANTES + CRUCE CON LAS PREDICCIONES DE LA IA ---
            // Cada hora incumplida se cruza con la hoja 'Predicciones Horarias'
            // (Analisis_Prioridad_IA.xlsx): si la meta de DEMANDA de la IA también
            // supera las cajas cubiertas, el faltante es real de personal; si no,
            // lo exige el ANCLA del perfil corporativo (nivel de servicio).
            let iaMap = {};
            (window.IA_PREDICCIONES || []).forEach(r => { iaMap[r['Fecha'] + '|' + Number(r['Hora'])] = r; });
            let hayIA = Object.keys(iaMap).length > 0;
            let fIso = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
            let totFaltan = 0, totHorasMal = 0, sumMaxDia = 0, horasDemanda = 0, horasConIA = 0;
            diasMal.forEach(dm => {
                let maxDia = 0, iso = fIso(dm.fecha);
                dm.ev.fallos.forEach(f => {
                    totFaltan += f.faltan; totHorasMal++;
                    if (f.faltan > maxDia) maxDia = f.faltan;
                    let r = iaMap[iso + '|' + f.h];
                    if (r) {
                        horasConIA++;
                        let metaDem = Number(r['Cajas_Meta_Demanda'] !== '' && r['Cajas_Meta_Demanda'] !== undefined
                                             ? r['Cajas_Meta_Demanda'] : r['Cajas_Meta']) || 0;
                        f.iaMetaDem = metaDem;
                        f.iaUnds = Number(r['Pred_Unds']) || 0;
                        f.demandaExige = metaDem > f.n;
                        if (f.demandaExige) horasDemanda++;
                    }
                });
                sumMaxDia += maxDia;
            });
            let promHora = totHorasMal ? totFaltan / totHorasMal : 0;
            let promDia = diasMal.length ? totFaltan / diasMal.length : 0;
            let personasDia = diasMal.length ? sumMaxDia / diasMal.length : 0;

            let _gaugeData = _evalGaugeCumpl(window.__CUMPL_MES || '');
            let htmlOut = _gaugeCumplHTML(_gaugeData) + `<div style="display:flex; gap:14px; flex-wrap:wrap; margin-bottom:16px;">
                <div style="flex:1; min-width:150px; padding:12px 16px; background:#eaf0fb; border-radius:8px; text-align:center;">
                    <div style="font-size:24px; font-weight:bold; color:#2c3e50;">${totalDias}</div>
                    <div style="font-size:11px; color:#7f8c8d;">Días evaluados</div></div>
                <div style="flex:1; min-width:150px; padding:12px 16px; background:${diasMal.length ? '#fdecea' : '#e8f8f0'}; border-radius:8px; text-align:center;">
                    <div style="font-size:24px; font-weight:bold; color:${diasMal.length ? '#c0392b' : '#27ae60'};">${diasMal.length}</div>
                    <div style="font-size:11px; color:#7f8c8d;">Días que NO cumplen el perfil</div></div>
                <div style="flex:1; min-width:150px; padding:12px 16px; background:#e8f8f0; border-radius:8px; text-align:center;">
                    <div style="font-size:24px; font-weight:bold; color:#27ae60;">${totalDias - diasMal.length}</div>
                    <div style="font-size:11px; color:#7f8c8d;">Días OK${diasOrg.length ? ' (incluye ' + diasOrg.length + ' de solo organización)' : ''}</div></div>
            </div>`;

            if (!diasMal.length) {
                htmlOut += `<div style="padding:14px 18px; background:#e8f8f0; border:1px solid #27ae60; border-left:6px solid #27ae60; border-radius:8px; font-size:13px; color:#1e8449;">
                    ✅ Los ${totalDias - diasOrg.length} día(s) con incumplimiento real: NINGUNO. La malla cumple el perfil de cobertura de la batería (${bateria} cajas): ${_descPerfil()}${!_tablaAnclasTienda() ? ' · domingos ≥80% en la tarde' : ''}.</div>`;
                htmlOut += _bloqueDiasOrg(diasOrg, bateria, NOMBRES_DIA);
            } else {
                // --- Promedios de faltantes ---
                htmlOut += `<div style="display:flex; gap:14px; flex-wrap:wrap; margin-bottom:14px;">
                    <div style="flex:1; min-width:170px; padding:12px 16px; background:#fdf2f0; border:1px solid #e74c3c; border-radius:8px; text-align:center;">
                        <div style="font-size:22px; font-weight:bold; color:#c0392b;">${promHora.toFixed(1)}</div>
                        <div style="font-size:11px; color:#7f8c8d;">Cajas que faltan en PROMEDIO<br>por hora incumplida</div></div>
                    <div style="flex:1; min-width:170px; padding:12px 16px; background:#fef9e7; border:1px solid #f1c40f; border-radius:8px; text-align:center;">
                        <div style="font-size:22px; font-weight:bold; color:#b7950b;">👥 ${Math.ceil(personasDia)}</div>
                        <div style="font-size:11px; color:#7f8c8d;">Personas extra estimadas por día<br>(pico simultáneo promedio de faltante)</div></div>
                </div>`;
                // --- Justificación con el análisis de predicciones y ventas de la IA ---
                if (hayIA) {
                    let horasPerfil = horasConIA - horasDemanda;
                    htmlOut += `<div style="margin-bottom:14px; padding:12px 16px; background:#f4ecf7; border:1px solid #8e44ad; border-left:6px solid #8e44ad; border-radius:8px; font-size:12px; color:#4a235a;">
                        🧠 <b>JUSTIFICACIÓN CON LA IA (predicciones y ventas):</b> de las <b>${totHorasMal}</b> horas incumplidas${horasConIA < totHorasMal ? ` (${horasConIA} con predicción disponible)` : ''},
                        en <b>${horasDemanda}</b> la propia demanda prevista por la IA también supera las cajas cubiertas → <b>faltante REAL de personal</b>;
                        en <b>${horasPerfil}</b> la demanda estaría atendida, pero el <b>ancla del perfil corporativo</b> exige más cobertura (nivel de servicio / tiempo de espera).
                        En promedio faltan <b>${promHora.toFixed(1)}</b> cajas por hora incumplida, lo que equivale a ~<b>${Math.ceil(personasDia)}</b> persona(s) extra por día en el momento más crítico.
                        <span style="color:#7d6608;">La meta de la IA ya incluye el ajuste por el tiempo real de atención de este año y el piso del perfil (columna 🛡️ del gráfico de Demanda).</span></div>`;
                } else {
                    htmlOut += `<div style="margin-bottom:14px; padding:10px 16px; background:#fff3cd; border-left:5px solid #f1c40f; border-radius:6px; font-size:12px; color:#7d6608;">
                        💡 Para justificar cada faltante con la demanda prevista y las ventas, carga el archivo <b>Analisis_Prioridad_IA.xlsx</b> (regenerado con el PY actualizado) en el panel de análisis IA.</div>`;
                }
                // --- JUSTIFICACIÓN MATEMÁTICA GLOBAL + RECOMENDACIÓN DE PERSONAL ---
                let recomendado = Math.max(0, ...diasMal.map(dm => dm.jf.faltan));
                let diaPeor = diasMal.find(dm => dm.jf.faltan === recomendado);
                let diasConDeficit = diasMal.filter(dm => dm.jf.faltan > 0).length;
                let diasSoloOrg = diasMal.length - diasConDeficit;
                if (recomendado > 0 && diaPeor) {
                    let j = diaPeor.jf;
                    let detalleCota = j.porHoras >= j.porVentanas
                        ? `la meta del día suma <b>${j.D} horas-caja</b> y la plantilla presente solo aporta <b>${Math.round(j.S)} horas-caja</b> (${j.P} titulares+emergentes) → déficit de ${Math.round(j.deficitHoras)} h-caja ≡ <b>${j.porHoras} persona(s)</b> de 6,5h`
                        : `las metas de las <b>${j.mejorPar[0]}:00</b> (${j.metas[j.mejorPar[0]]}) y las <b>${j.mejorPar[1]}:00</b> (${j.metas[j.mejorPar[1]]}) están a ${j.mejorPar[1] - j.mejorPar[0]}h de distancia — ningún turno (máx. 8h + 1h de almuerzo = 9h de tramo) cubre ambas → se necesitan ≥ <b>${j.reqVentanas}</b> personas y ese día hay <b>${j.P}</b>`;
                    htmlOut += `<div style="margin-bottom:14px; padding:12px 16px; background:#eaf2f8; border:1px solid #2980b9; border-left:6px solid #2980b9; border-radius:8px; font-size:12px; color:#1a5276;">
                        🧮 <b>JUSTIFICACIÓN MATEMÁTICA — FALTA PERSONAL TITULAR:</b> en ${diasConDeficit} de los ${diasMal.length} día(s) incumplidos, NINGÚN acomodo de turnos puede llegar al 100%: el peor caso es <b>${NOMBRES_DIA[diaPeor.fecha.getDay()]} ${diaPeor.header}</b>, donde ${detalleCota}.
                        ${diasSoloOrg > 0 ? `En los otros ${diasSoloOrg} día(s) el personal total sí alcanza: el hueco es de ORGANIZACIÓN de turnos (regenera la malla en el Python para que el optimizador los acomode).` : ''}
                        <div style="margin-top:6px; font-size:13px;">👥 <b>Personal adicional recomendado: ${recomendado} cajero(s)/emergente(s)</b> — usa el simulador de abajo para verificarlo.</div></div>`;
                } else {
                    htmlOut += `<div style="margin-bottom:14px; padding:12px 16px; background:#eaf2f8; border:1px solid #2980b9; border-left:6px solid #2980b9; border-radius:8px; font-size:12px; color:#1a5276;">
                        🧮 <b>JUSTIFICACIÓN MATEMÁTICA:</b> el personal total de cada día alcanza para las metas (las horas-caja disponibles cubren la demanda y ninguna pareja de franjas incompatibles supera la plantilla). Los incumplimientos son de <b>ORGANIZACIÓN de turnos</b>: regenera la malla en el Python para que el optimizador los acomode.</div>`;
                }

                // --- SIMULADOR: cobertura con N personas extra ---
                let SIM = window.__SIM_EXTRA;
                if (SIM && SIM.n > 0) {
                    let filasSim = '', diasOk = 0;
                    diasMal.forEach(dm => {
                        let sim = _simulaDia(dm.ev.cob, dm.jf.metas, SIM.n, SIM.dur);
                        let fallosSim = [];
                        _anclasPerfil(dm.esDom, bateria).forEach(a => {
                            let nn = sim.cobS[a.h] || 0;
                            if (a.objN > 0 && nn < a.objN) fallosSim.push(`${a.h}:00 (${nn}/${a.objN})`);
                        });
                        if (!fallosSim.length) diasOk++;
                        filasSim += `<tr style="background:${fallosSim.length ? '#fdecea' : '#e8f8f0'};">
                            <td style="padding:4px 10px; font-weight:bold;">${NOMBRES_DIA[dm.fecha.getDay()]} ${dm.header}</td>
                            <td style="padding:4px 10px; text-align:center;">${dm.ev.fallos.length}</td>
                            <td style="padding:4px 10px; text-align:center; font-weight:bold; color:${fallosSim.length ? '#c0392b' : '#27ae60'};">${fallosSim.length}</td>
                            <td style="padding:4px 10px; font-size:11px;">${fallosSim.length ? 'Siguen cortas: ' + fallosSim.join(' · ') : '✅ Cumple TODO el perfil'}</td>
                            <td style="padding:4px 10px; font-size:11px; color:#7f8c8d;">entradas: ${sim.colocados.join(', ') || '—'}</td>
                          </tr>`;
                    });
                    htmlOut += `<div style="margin-bottom:14px; padding:12px 16px; background:#fef9e7; border:1px solid #f1c40f; border-left:6px solid #f39c12; border-radius:8px;">
                        <div style="font-weight:bold; color:#7d6608; font-size:13px; margin-bottom:6px;">🧪 SIMULACIÓN: +${SIM.n} cajero(s) de ${SIM.dur}h — ${diasOk}/${diasMal.length} día(s) incumplidos quedarían al 100% del perfil</div>
                        <div style="overflow-x:auto;"><table style="border-collapse:collapse; font-size:12px; background:#fff; border-radius:6px;">
                          <thead><tr style="background:#f39c12; color:#fff; font-size:11px;">
                            <th style="padding:4px 10px;">Día</th><th style="padding:4px 10px;">Horas mal (hoy)</th>
                            <th style="padding:4px 10px;">Horas mal (simulado)</th><th style="padding:4px 10px;">Resultado</th>
                            <th style="padding:4px 10px;">Horas de entrada sugeridas</th>
                          </tr></thead><tbody>${filasSim}</tbody></table></div>
                        <div style="font-size:11px; color:#7f8c8d; margin-top:6px;">La simulación coloca a cada persona extra en el inicio de turno que más huecos tapa (con su hora de almuerzo en la hora de menor déficit). Es una cota optimista: la malla real debe respetar rotaciones y descansos.</div>
                    </div>`;
                }

                htmlOut += _bloqueDiasOrg(diasOrg, bateria, NOMBRES_DIA);
                htmlOut += `<div style="font-size:12px; color:#7f8c8d; margin-bottom:10px;">Perfil exigido cada día (sobre la batería de <b>${bateria}</b> cajas, sumando Cajeros + Emergentes sin personas con X): <b>${_descPerfil()}</b>${!_tablaAnclasTienda() ? ' · domingos además <b>≥80%</b> de 13:00 a 20:00' : ''}.</div>`;
                diasMal.forEach(dm => {
                    let detalle = dm.ev.fallos.map(f => {
                        let celIA = '';
                        if (hayIA) {
                            if (f.iaMetaDem !== undefined) {
                                celIA = f.demandaExige
                                    ? `<td style="padding:3px 10px; font-size:11px; color:#c0392b;"><b>IA: ${f.iaMetaDem} cajas</b> (${f.iaUnds.toLocaleString('es-CO')} unds) — la demanda TAMBIÉN lo exige</td>`
                                    : `<td style="padding:3px 10px; font-size:11px; color:#7d6608;">IA: ${f.iaMetaDem} cajas (${f.iaUnds.toLocaleString('es-CO')} unds) — lo exige el ancla del perfil</td>`;
                            } else {
                                celIA = `<td style="padding:3px 10px; font-size:11px; color:#95a5a6;">sin predicción para esta fecha</td>`;
                            }
                        }
                        return `<tr>
                           <td style="padding:3px 10px; font-weight:bold; color:#c0392b;">${f.h}:00</td>
                           <td style="padding:3px 10px; text-align:center;">${f.n}/${bateria}</td>
                           <td style="padding:3px 10px; text-align:center; color:#c0392b; font-weight:bold;">${Math.round(f.pct * 100)}%</td>
                           <td style="padding:3px 10px; text-align:center;">${Math.round(f.obj * 100)}%</td>
                           <td style="padding:3px 10px; text-align:center; font-weight:bold; color:#c0392b;">faltan ${f.faltan}</td>
                           <td style="padding:3px 10px; font-size:11px; color:#7f8c8d;">${f.txt}${f.ultimo ? ' ⏱️' : ''}</td>
                           ${celIA}
                         </tr>`;
                    }).join('');
                    let notaExtra = dm.ev.falloUltimo
                        ? `<div style="font-size:11px; color:#7b241c; margin-top:6px; padding:6px 10px; background:#fbeee6; border-left:4px solid #e67e22; border-radius:4px;">
                             ⏱️ <b>ÚLTIMO TURNO:</b> la cobertura del cierre (21:00-21:30) debe completarse con <b>HORAS EXTRA de personas que salen un poco antes</b>.</div>`
                        : '';
                    // Personas en VACACIONES este día (cajeros/emergentes)
                    if (dm.vacCaja && dm.vacCaja.length) {
                        notaExtra += `<div style="font-size:11px; color:#1a5276; margin-top:6px; padding:6px 10px; background:#eaf2f8; border-left:4px solid #2980b9; border-radius:4px;">
                             🏖️ <b>EN VACACIONES este día (${dm.vacCaja.length}):</b> ${dm.vacCaja.join(' · ')} — no cuentan en la cobertura.</div>`;
                    }
                    // Personas con RESTRICCIÓN activas este día (X = no entran a caja; ⏰ = solo hasta cierta hora)
                    let restrX = [], restrH = [];
                    (excelData || []).forEach(it => {
                        const trabaja = _rangoTurnoDec((it.data[dm.colIdx] || '').toString().trim());
                        if (it.excluida) restrX.push((it.data[0] || '').toString().trim());
                        else if (trabaja && (it.restriccion || '').toUpperCase().includes('HASTA'))
                            restrH.push((it.data[0] || '').toString().trim() + ' (' + it.restriccion + ')');
                    });
                    if (restrX.length || restrH.length) {
                        notaExtra += `<div style="font-size:11px; color:#7b241c; margin-top:6px; padding:6px 10px; background:#fdecea; border-left:4px solid #c0392b; border-radius:4px;">
                             ⛔ <b>PERSONAL CON RESTRICCIÓN:</b>${restrX.length ? ` <b>X (no entran a caja, no suman):</b> ${restrX.join(' · ')}.` : ''}${restrH.length ? ` <span style="color:#b9770e;">⏰ Solo hasta cierta hora:</span> ${restrH.join(' · ')}.` : ''}</div>`;
                    }
                    let j = dm.jf;
                    if (j.faltan > 0) {
                        notaExtra += `<div style="font-size:11px; color:#1a5276; margin-top:6px; padding:6px 10px; background:#eaf2f8; border-left:4px solid #2980b9; border-radius:4px;">
                             🧮 <b>Matemáticamente faltan ${j.faltan} persona(s) este día:</b> plantilla ${j.P} titulares+emergentes = ${Math.round(j.S)} horas-caja vs ${j.D} horas-caja de meta (déficit ${Math.round(j.deficitHoras)})${j.porVentanas > j.porHoras ? ` · además las franjas de las ${j.mejorPar[0]}:00 y las ${j.mejorPar[1]}:00 son incompatibles para una misma persona y exigen ${j.reqVentanas} personas` : ''}. Ningún acomodo de turnos alcanza el 100%.</div>`;
                    } else {
                        // faltan === 0 pero está en diasMal -> hay cajeros de vacaciones
                        notaExtra += `<div style="font-size:11px; color:#1a5276; margin-top:6px; padding:6px 10px; background:#eaf2f8; border-left:4px solid #2980b9; border-radius:4px;">
                             🏖️ <b>El hueco lo causa la AUSENCIA POR VACACIONES:</b> hay ${dm.vacCaja.length} cajero(s)/emergente(s) en vacaciones este día (${dm.vacCaja.join(', ')}). El personal presente (${j.P} = ${Math.round(j.S)} h-caja) alcanzaría para la meta (${j.D} h-caja), así que no es déficit de plantilla ni de organización: es la afectación de vacaciones.</div>`;
                    }
                    htmlOut += `<div style="margin-bottom:12px; padding:12px 16px; background:#fdecea; border:1px solid #e74c3c; border-left:6px solid #c0392b; border-radius:8px;">
                        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                            <div style="font-weight:bold; color:#c0392b; font-size:13px;">🔴 ${NOMBRES_DIA[dm.fecha.getDay()]} ${dm.header} (${dm.fecha.toLocaleDateString()})${dm.esDom ? ' · DOMINGO' : ''} — ${dm.ev.fallos.length} hora(s) incumplida(s)</div>
                            <button onclick="verDiaEnCobertura(${dm.colIdx})" style="padding:3px 10px; font-size:10px; font-weight:bold; color:#fff; background:#2980b9; border:none; border-radius:4px; cursor:pointer;">🕒 Ver línea de tiempo</button>
                        </div>
                        <div style="overflow-x:auto; margin-top:6px;">
                        <table style="border-collapse:collapse; font-size:12px; background:#fff; border-radius:6px;">
                          <thead><tr style="background:#c0392b; color:#fff; font-size:11px;">
                            <th style="padding:4px 10px;">Hora</th><th style="padding:4px 10px;">Cajas</th>
                            <th style="padding:4px 10px;">Cobertura</th><th style="padding:4px 10px;">Meta</th>
                            <th style="padding:4px 10px;">Faltante</th><th style="padding:4px 10px;">Regla</th>
                            ${hayIA ? '<th style="padding:4px 10px;">🧠 Justificación IA (demanda)</th>' : ''}
                          </tr></thead><tbody>${detalle}</tbody></table></div>
                        ${notaExtra}
                    </div>`;
                });
            }
            cont.innerHTML = htmlOut;
            _dibujarGaugeCumpl(_gaugeData);
            _minOverrideActivo = null;   // no filtrar fuera del informe de Cumplimiento
        }


        // Lee de la pestaña PLANTA del Sheets (window.PLANTA_DATA): quiénes
        // tienen X (NO pueden entrar a caja) y las fechas de vacaciones.
        function _plantaInfo() {
            const pd = window.PLANTA_DATA;
            if (!pd || pd.length < 2) return null;
            const head = pd[0].map(x => (x || '').toString().toUpperCase());
            const iNom = head.findIndex(x => x.includes('NOMBRE'));
            const iRes = head.findIndex(x => x.includes('RESTRIC'));
            const iIni = head.findIndex(x => x.includes('INICIO'));
            const iFin = head.findIndex(x => x.includes('FIN'));
            const iRol = head.findIndex(x => x.includes('ROL'));
            const xSet = new Set(); const vacacionesDe = {};
            for (let r = 1; r < pd.length; r++) {
                const nom = iNom >= 0 ? (pd[r][iNom] || '').toString().trim() : '';
                if (!nom) continue;
                if (iRes >= 0 && (pd[r][iRes] || '').toString().trim().toUpperCase() === 'X') xSet.add(nom);
                if (iIni >= 0 && iFin >= 0 && pd[r][iIni] && pd[r][iFin]) {
                    const d1 = new Date(pd[r][iIni] + 'T12:00:00'), d2 = new Date(pd[r][iFin] + 'T12:00:00');
                    const rol = iRol >= 0 ? (pd[r][iRol] || '').toString() : '';
                    if (!isNaN(d1) && !isNaN(d2)) vacacionesDe[nom] = [d1, d2, rol];
                }
            }
            return { xSet: xSet, vacacionesDe: vacacionesDe };
        }

        // ¿El rol (texto libre) corresponde a una CAJA (cajero titular o emergente)?
        function _esRolCaja(rolStr) {
            const g = (rolStr || '').toString().toUpperCase();
            if (g.includes('EMERGENTE')) return true;
            return g.includes('CAJERO') && !g.includes('CAMBISTA') && !g.includes('VISADO')
                   && !g.includes('SATEL') && !g.includes('SATÉL');
        }

        // Cajeros / emergentes EN VACACIONES un día dado (VC en la malla + rangos
        // de fecha de la pestaña PLANTA). Devuelve la lista de nombres.
        function _vacCajaDia(colIdx, fechaDia) {
            let lista = [];
            (excelData || []).forEach(it => {
                const cat = _rolCategoria(it);
                if (cat !== 'Cajero' && cat !== 'Emergente') return;
                if ((it.data[colIdx] || '').toString().trim().toUpperCase() === 'VC')
                    lista.push((it.data[0] || '').toString().trim());
            });
            const pi = _plantaInfo();
            if (pi) {
                Object.keys(pi.vacacionesDe).forEach(nom => {
                    const v = pi.vacacionesDe[nom];
                    if (v.length >= 3 && !_esRolCaja(v[2])) return;   // solo roles de caja
                    if (fechaDia >= v[0] && fechaDia <= v[1] && lista.indexOf(nom) < 0) lista.push(nom);
                });
            }
            return lista;
        }

        // Marca en la malla cargada a quienes tienen X en la pestaña PLANTA del
        // Sheets (la fuente oficial de "no puede entrar a caja").
        window.sincronizarXDesdePlanta = function () {
            const pi = _plantaInfo();
            if (!pi || !excelData) return 0;
            let n = 0;
            excelData.forEach(it => {
                const nom = (it.data[0] || '').toString().trim();
                if (pi.xSet.has(nom) && !it.excluida) {
                    it.excluida = true;
                    if (!it.restriccion) it.restriccion = 'X';
                    n++;
                }
            });
            if (n > 0) {
                if (typeof renderLineaTiempoCobertura === 'function') renderLineaTiempoCobertura();
                if (typeof renderCumplimiento === 'function') renderCumplimiento();
            }
            return n;
        };

        // ================================================================
        //  PLANTA IDEAL (dimensionamiento WFM) — por ahora solo AKB30.
        //  Calcula la planta óptima con la metodología completa: necesidad
        //  diaria por cinta, cobertura hora a hora, planta por reglas de
        //  descanso (restricción dominante), roles especiales, semana tipo,
        //  deuda de 6h, horas-caja / horas-hombre, planta actual vs ideal y
        //  afectación de vacaciones por mes.
        // ================================================================
        const PLANTA_IDEAL_CFG = {
            'AKB30': {
                bateria: 19,
                cicloDomingo: 4,          // libra 1 de cada 4 domingos (75% disponible)
                propTitular: 0.75,        // titular : emergente = 3 : 1
                horasExtraMax: 1, topeCierreExtra: 11,
                cintaSemana: [            // L-V jornada 7h — 38 cupos (turnos actuales)
                    ['7A8-15',      11, '—'],           // 8:00-15:00 apertura (corrido)
                    ['7i10-18',      8, '14:00-15:00'], // 10:00-18:00 (con almuerzo)
                    ['7C13.3-20.3',  8, '—'],           // 13:30-20:30 tarde (corrido)
                    ['7C14.3-21.3', 11, '—']            // 14:30-21:30 cierre (corrido)
                ],
                cintaSemanaAlta: [        // L-V TEMPORADA ALTA abr/ago/dic — 9/9/9/9 = 36 cupos (jornada 6,5h)
                    ['6,5A8-14.3',  9, '—'],            // 8:00-14:30  apertura (corrido)
                    ['6,5i10.3-18', 9, '14:00-15:00'],  // 10:30-18:00 (con almuerzo)
                    ['6,5C14.3-21', 9, '—'],            // 14:30-21:00 tarde (corrido)
                    ['6,5C15-21.3', 9, '—']             // 15:00-21:30 cierre (corrido)
                ],
                cintaDomingo: [           // Domingo NORMAL — SIEMPRE 8h — 33 cupos (10/4/5/14, igual que el generador)
                    ['8i8-17',      10, '12:30-13:30'], // 8:00-17:00  (con almuerzo)
                    ['8i10-19',      4, '14:00-15:00'], // 10:00-19:00 (con almuerzo)
                    ['8i11-20',      5, '14:30-15:30'], // 11:00-20:00 (con almuerzo)
                    ['8C13.3-21.3', 14, '—']            // 13:30-21:30 cierre (corrido)
                ],
                cintaDomingoAlta: [       // Domingo TEMPORADA ALTA abr/ago/dic — 33 cupos (10/6/5/12)
                    ['8i8-17',      10, '12:30-13:30'], // 8:00-17:00
                    ['8i10-19',      6, '14:00-15:00'], // 10:00-19:00 (+2 vs normal)
                    ['8i11-20',      5, '14:30-15:30'], // 11:00-20:00
                    ['8C13.3-21.3', 12, '—']            // 13:30-21:30 cierre (-2 vs normal)
                ],
                rolesEspeciales: [        // [rol, apertura, cierre, nota, fijo/día (opcional)]
                    ['Visado',    8, 21, 'nunca queda sola: el relevo cubre el almuerzo'],
                    ['Cambista',  8, 21, 'mismo horario que visado'],
                    ['Akomer',   11, 20, 'definido por la tienda: SOLO 1 cajero', 1],
                    ['Satélite 1', 11, 20, ''],
                    ['Satélite 2', 11, 20, '']
                ]
            }
        };

        function _piPresencia(cinta) {
            // presencia por hora (8..21) descontando almuerzos declarados en la cinta
            let pres = {};
            for (let h = 8; h <= 21; h++) pres[h] = 0;
            cinta.forEach(fila => {
                const r = _rangoTurnoDec(fila[0]);
                if (!r) return;
                let a = null;
                const mAlm = /^(\d+):(\d+)-(\d+):/.exec(fila[2] || '');
                if (mAlm) a = parseInt(mAlm[1]) + parseInt(mAlm[2]) / 60;
                for (let h = 8; h <= 21; h++) {
                    if (!(r.ini <= h + 1e-9 && r.fin >= h + 0.5 - 1e-9)) continue;
                    if (a !== null && a <= h && h < a + 1) continue;
                    pres[h] += fila[1];
                }
            });
            return pres;
        }

        // Cobertura real (cajeros+emergentes) por hora, promedio/min/max sobre
        // los días del tipo pedido ('LV' L-Vie, 'SAB' sábados, 'DOM' domingos)
        // dentro del mes seleccionado. Devuelve {h:{prom,min,max,n}}.
        function _piCoberturaReal(tipoDia, mesSel) {
            let acc = {};
            for (let h = 8; h <= 21; h++) acc[h] = { sum: 0, min: Infinity, max: 0, n: 0 };
            if (!excelData || !headers) return acc;
            let startVal = (document.getElementById('fechaInicio') || {}).value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            headers.forEach((hh, i) => {
                if (hh === 'Total Horas Periodo') return;
                let colIdx = i + 1;
                if (!excelData.some(it => _rangoTurnoDec((it.data[colIdx] || '').toString().trim()))) return;
                let d = new Date(startDate); d.setDate(d.getDate() + i);
                let wd = d.getDay();
                let tipo = wd === 0 ? 'DOM' : (wd === 6 ? 'SAB' : 'LV');
                if (tipo !== tipoDia) return;
                let mesKey = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
                if (mesSel && mesKey !== mesSel) return;
                let cob = _cobCajasDia(colIdx, wd === 0);
                for (let h = 8; h <= 21; h++) {
                    let n = cob[h] || 0;
                    acc[h].sum += n; acc[h].n++;
                    acc[h].min = Math.min(acc[h].min, n); acc[h].max = Math.max(acc[h].max, n);
                }
            });
            return acc;
        }

        function _piHoras(cinta) {
            // [horas-caja (laborales), horas-hombre (presencia con almuerzo)]
            let hc = 0, hh = 0;
            cinta.forEach(f => {
                const lab = _horasLaboralesTurno(f[0]) || 0;
                const r = _rangoTurnoDec(f[0]);
                hc += lab * f[1];
                hh += (r ? r.fin - r.ini : lab) * f[1];
            });
            return [hc, hh];
        }

        function renderPlantaIdeal() {
            const cont = document.getElementById('plantaIdealContainer');
            if (!cont) return;
            const tienda = (window.APP_TIENDA || '').toUpperCase();
            const cfgBase = PLANTA_IDEAL_CFG[tienda];
            if (!cfgBase) {
                cont.innerHTML = `<div style="padding:14px 18px; background:#fff3cd; border-left:5px solid #f1c40f; border-radius:6px; font-size:13px; color:#7d6608;">
                    ⚠️ La tienda <b>${tienda || '—'}</b> aún no tiene módulo de Planta Ideal. Por ahora solo está personalizado <b>AKB30</b> (las demás se irán creando).</div>`;
                return;
            }
            // Temporada alta (abr/ago/dic): cintas reforzadas — L-V 9/9/9/9 con
            // entradas 8:00/10:30/14:30/15:00 y domingo 10/6/5/12. Resto del año: normal.
            const _piMesNum = (() => { let v = (document.getElementById('piMes') || {}).value || ''; return v ? parseInt(v.split('-')[1], 10) : 0; })();
            const cfg = [4, 8, 12].includes(_piMesNum)
                ? Object.assign({}, cfgBase, {
                    cintaDomingo: cfgBase.cintaDomingoAlta || cfgBase.cintaDomingo,
                    cintaSemana: cfgBase.cintaSemanaAlta || cfgBase.cintaSemana
                })
                : cfgBase;
            const tabla = _tablaAnclasTienda() || {};
            const fmt1 = x => (Math.round(x * 10) / 10).toLocaleString('es-CO');

            // ---------- 1. Necesidad diaria ----------
            const necSem = cfg.cintaSemana.reduce((a, f) => a + f[1], 0);      // 38
            const necDom = cfg.cintaDomingo.reduce((a, f) => a + f[1], 0);     // 33
            // ¿una persona cubre ambos picos? presencia máx = 7h + 1h almuerzo = 8h;
            // pico AM termina 13:00 y el PM llega hasta las 20:00 -> 8:00→20:00 = 12h > 8h
            const disponibleDom = 1 - 1 / cfg.cicloDomingo;                    // 0.75

            // ---------- 3. Planta por regla de descanso (cajeros) ----------
            const plantaDom = Math.ceil(necDom / disponibleDom);               // 44
            const plantaLV  = Math.ceil(necSem / (1 - disponibleDom / 5));     // 45
            const plantaSab = necSem;                                          // 38 (nadie descansa)
            const plantaCajeros = Math.max(plantaDom, plantaLV, plantaSab);
            const domCaj = plantaLV >= plantaDom && plantaLV >= plantaSab ? 'LUNES-VIERNES'
                          : (plantaDom >= plantaSab ? 'DOMINGO' : 'SÁBADO');
            const compDia = plantaCajeros * disponibleDom / 5;                 // COMP por día L-V

            // ---------- 4. Roles especiales ----------
            const presenciaMax = 8;   // 7h + 1h de almuerzo
            // Conteo de personas REALES por rol (de la malla cargada)
            const _kwRol = nom => { const n = nom.toUpperCase();
                if (n.includes('VISADO')) return 'VISADO';
                if (n.includes('CAMBISTA')) return 'CAMBISTA';
                if (n.includes('AKOMER')) return 'AKOMER';
                if (n.includes('SATEL') || n.includes('SATÉL')) return 'SATEL';
                return null; };
            let actPorKw = {};
            (excelData || []).forEach(it => {
                const kw = _kwRol(it.rolGenerico || it.rol || '');
                if (kw) actPorKw[kw] = (actPorKw[kw] || 0) + 1;
            });
            let filasRoles = '', totEspDia = 0, horasCajaEsp = 0, kwVisto = {};
            cfg.rolesEspeciales.forEach(re => {
                const hCaja = re[2] - re[1];
                const pCalc = Math.max(Math.ceil(hCaja / presenciaMax), hCaja > 7 ? 2 : 1);
                const pDia = re[4] ? re[4] : pCalc;   // la tienda puede FIJAR la cantidad
                totEspDia += pDia; horasCajaEsp += hCaja;
                const cuenta = re[4]
                    ? `fijado por la tienda: <b>${pDia}</b>/día` + (pCalc > pDia ? ` <span style="color:#e67e22;">(el cálculo sugiere ${pCalc})</span>` : '')
                    : `${hCaja} ÷ ${presenciaMax}h → <b>${pDia}</b>/día`;
                const kw = _kwRol(re[0]);
                let actTxt;
                if (kw === 'SATEL') {   // los 2 satélites comparten el grupo
                    actTxt = kwVisto[kw] ? '<span style="color:#95a5a6;">↑ mismo grupo</span>' : `<b>${actPorKw.SATEL || 0}</b> en total`;
                } else {
                    actTxt = `<b>${(kw && actPorKw[kw]) || 0}</b>`;
                }
                if (kw) kwVisto[kw] = true;
                filasRoles += `<tr><td style="padding:4px 10px; font-weight:bold;">${re[0]}</td>
                    <td style="padding:4px 10px; text-align:center;">${re[1]}:00-${re[2]}:00</td>
                    <td style="padding:4px 10px; text-align:center;">${hCaja} h-caja</td>
                    <td style="padding:4px 10px; text-align:center;">${cuenta}</td>
                    <td style="padding:4px 10px; text-align:center;">${actTxt}</td>
                    <td style="padding:4px 10px; font-size:11px; color:#7f8c8d;">${re[3] || ''}</td></tr>`;
            });
            const totalEspActual = (actPorKw.VISADO || 0) + (actPorKw.CAMBISTA || 0) + (actPorKw.AKOMER || 0) + (actPorKw.SATEL || 0);
            const plantaEspDom = Math.ceil(totEspDia / disponibleDom);
            const plantaEspLV  = Math.ceil(totEspDia / (1 - disponibleDom / 5));
            const plantaEsp = Math.max(plantaEspDom, plantaEspLV, totEspDia);
            const domEsp = plantaEspDom >= plantaEspLV ? 'DOMINGO' : 'LUNES-VIERNES';
            const plantaTotal = plantaCajeros + plantaEsp;

            // ---------- Horas-caja / horas-hombre ----------
            const [hcSem, hhSem] = _piHoras(cfg.cintaSemana);
            const [hcDom, hhDom] = _piHoras(cfg.cintaDomingo);
            let hcReq = 0;
            for (let h = 8; h <= 20; h++) hcReq += (tabla[h] !== undefined ? tabla[h] : (_metasHoraDia(cfg.bateria, false)[h] || 0));
            hcReq += (tabla[21] || 0) * 0.5;

            // ---------- Bloques de turnos (con proporción titular/emergente 3:1) ----------
            const bloque = (cinta, titulo) => {
                let filas = '';
                cinta.forEach(f => {
                    const r = _rangoTurnoDec(f[0]);
                    const tit = Math.round(f[1] * cfg.propTitular), eme = f[1] - tit;
                    const hIni = Math.floor(r.ini) + (r.ini % 1 ? ':30' : ':00');
                    const hFin = Math.floor(r.fin) + (r.fin % 1 ? ':30' : ':00');
                    filas += `<tr><td style="padding:4px 10px; font-weight:bold;">${f[0]}</td>
                        <td style="padding:4px 10px; text-align:center;">${hIni}</td>
                        <td style="padding:4px 10px; text-align:center;">${f[2]}</td>
                        <td style="padding:4px 10px; text-align:center;">${hFin}</td>
                        <td style="padding:4px 10px; text-align:center; font-weight:bold;">${f[1]}</td>
                        <td style="padding:4px 10px; text-align:center;">${tit} / ${eme}</td></tr>`;
                });
                const tot = cinta.reduce((a, f) => a + f[1], 0);
                return `<h4 style="margin:14px 0 6px 0; font-size:13px; color:#2c3e50;">${titulo} — ${tot} personas</h4>
                    <div style="overflow-x:auto;"><table style="border-collapse:collapse; font-size:12px; background:#fff;">
                    <thead><tr style="background:#2c3e50; color:#fff; font-size:11px;">
                    <th style="padding:4px 10px;">Turno</th><th style="padding:4px 10px;">Entrada</th>
                    <th style="padding:4px 10px;">Almuerzo</th><th style="padding:4px 10px;">Salida</th>
                    <th style="padding:4px 10px;">Cant.</th><th style="padding:4px 10px;">Tit / Emerg (3:1)</th>
                    </tr></thead><tbody>${filas}</tbody></table></div>`;
            };

            // ---------- Cobertura hora a hora: PROMEDIO / MÍN / MÁX reales vs anclas ----------
            const cobTabla = (cinta, titulo, tipoReal) => {
                const pres = _piPresencia(cinta);
                const mesSelC = (document.getElementById('piMes') || {}).value || '';
                const real = _piCoberturaReal(tipoReal, mesSelC);
                const nDias = real[11] ? real[11].n : 0;
                let head = '', fPres = '', fReq = '', fProm = '', fMin = '', fMax = '', fOk = '';
                for (let h = 8; h <= 21; h++) {
                    const req = tabla[h];
                    const rr = real[h];
                    const prom = rr.n ? rr.sum / rr.n : 0;
                    head += `<th style="padding:3px 6px;">${h}</th>`;
                    fProm += `<td style="padding:3px 6px; text-align:center; font-weight:bold;">${rr.n ? prom.toFixed(1) : '—'}</td>`;
                    fMin += `<td style="padding:3px 6px; text-align:center; color:#c0392b;">${rr.n ? (rr.min === Infinity ? 0 : rr.min) : '—'}</td>`;
                    fMax += `<td style="padding:3px 6px; text-align:center; color:#27ae60;">${rr.n ? rr.max : '—'}</td>`;
                    fPres += `<td style="padding:3px 6px; text-align:center; color:#7f8c8d;">${pres[h]}</td>`;
                    fReq += `<td style="padding:3px 6px; text-align:center; color:#7f8c8d;">${req !== undefined ? req : '·'}</td>`;
                    const ok = req === undefined || (rr.n ? prom >= req - 1e-9 : pres[h] >= req);
                    fOk += `<td style="padding:2px 6px; text-align:center; color:${ok ? '#27ae60' : '#c0392b'}; font-weight:bold;">${req === undefined ? '' : (ok ? '✓' : '✗')}</td>`;
                }
                return `<h4 style="margin:14px 0 6px 0; font-size:13px; color:#2c3e50;">${titulo}</h4>
                    <div style="font-size:11px; color:#7f8c8d; margin-bottom:4px;">Promedio · mínimo · máximo de cajeros por hora sobre <b>${nDias}</b> día(s) reales${mesSelC ? ' de ' + mesSelC : ' (todos los meses)'}. La cinta ideal y el ancla se muestran de referencia.</div>
                    <div style="overflow-x:auto;"><table style="border-collapse:collapse; font-size:11px; background:#fff;">
                    <thead><tr style="background:#34495e; color:#fff;"><th style="padding:3px 8px;">Hora</th>${head}</tr></thead>
                    <tbody>
                    <tr><td style="padding:3px 8px; font-weight:bold;">Promedio real</td>${fProm}</tr>
                    <tr><td style="padding:3px 8px; color:#c0392b;">Mínimo</td>${fMin}</tr>
                    <tr><td style="padding:3px 8px; color:#27ae60;">Máximo</td>${fMax}</tr>
                    <tr><td style="padding:3px 8px; color:#7f8c8d;">Cinta ideal</td>${fPres}</tr>
                    <tr><td style="padding:3px 8px;">Ancla mínima</td>${fReq}</tr>
                    <tr><td style="padding:3px 8px;">Cumple (prom ≥ ancla)</td>${fOk}</tr>
                    </tbody></table></div>`;
            };

            // ---------- Semana tipo (trabajan / descansan) ----------
            const trabDom = Math.floor(plantaCajeros * disponibleDom);
            let semanaTipo = '';
            [['Lunes', necSem], ['Martes', necSem], ['Miércoles', necSem], ['Jueves', necSem],
             ['Viernes', necSem], ['Sábado', necSem], ['Domingo', necDom]].forEach(([dia, nec]) => {
                let disp, desc;
                if (dia === 'Domingo') { disp = trabDom; desc = plantaCajeros - trabDom; }
                else if (dia === 'Sábado') { disp = plantaCajeros; desc = 0; }
                else { desc = Math.ceil(compDia); disp = plantaCajeros - desc; }
                const holg = disp - nec;
                semanaTipo += `<tr style="background:${holg <= 0 ? '#fdecea' : '#fff'};">
                    <td style="padding:4px 10px; font-weight:bold;">${dia}</td>
                    <td style="padding:4px 10px; text-align:center;">${nec}</td>
                    <td style="padding:4px 10px; text-align:center;">${disp}</td>
                    <td style="padding:4px 10px; text-align:center;">${desc}</td>
                    <td style="padding:4px 10px; text-align:center; font-weight:bold; color:${holg < 0 ? '#c0392b' : (holg === 0 ? '#e67e22' : '#27ae60')};">${holg >= 0 ? '+' + holg : holg}${holg === 0 ? ' ⚠' : ''}</td></tr>`;
            });

            // ---------- Deuda de 6h ----------
            const dom8h = cfg.cintaDomingo.filter(f => (_horasLaboralesTurno(f[0]) || 0) === 8)
                                          .reduce((a, f) => a + f[1], 0);
            const deudaDia = Math.ceil(dom8h / 2);

            // ---------- Planta ACTUAL (de la malla o Mi Planta) ----------
            let actCajeros = 0, actEsp = 0, actX = 0;
            (excelData || []).forEach(it => {
                const cat = _rolCategoria(it);
                if (cat === 'Cajero' || cat === 'Emergente') { actCajeros++; if (it.excluida) actX++; }
                else if (cat === 'Cambista' || cat === 'Visado' || cat === 'Satélite') actEsp++;
            });

            // ---------- Vacaciones ----------
            const mesSel = (document.getElementById('piMes') || {}).value || '';
            const extraVac = parseInt((document.getElementById('piVacExtra') || {}).value) || 0;
            let vacMes = 0, vacDetalle = '';
            const pd = window.PLANTA_DATA;
            if (mesSel && pd && pd.length > 1) {
                const head = pd[0].map(x => (x || '').toString().toUpperCase());
                const iIni = head.findIndex(x => x.includes('INICIO')), iFin = head.findIndex(x => x.includes('FIN'));
                const iNom = head.findIndex(x => x.includes('NOMBRE'));
                if (iIni >= 0 && iFin >= 0) {
                    const m0 = new Date(mesSel + '-01T12:00:00');
                    const m1 = new Date(m0.getFullYear(), m0.getMonth() + 1, 0, 12);
                    const nombres = [];
                    for (let rI = 1; rI < pd.length; rI++) {
                        const dIni = pd[rI][iIni] ? new Date(pd[rI][iIni] + 'T12:00:00') : null;
                        const dFin = pd[rI][iFin] ? new Date(pd[rI][iFin] + 'T12:00:00') : null;
                        if (dIni && dFin && !isNaN(dIni) && !isNaN(dFin) && dIni <= m1 && dFin >= m0) {
                            vacMes++;
                            if (iNom >= 0) nombres.push(pd[rI][iNom]);
                        }
                    }
                    vacDetalle = nombres.length ? ' (' + nombres.join(', ') + ')' : '';
                }
            }
            const vacTotal = vacMes + extraVac;
            const colchon = 1;
            const recomendada = plantaTotal + vacTotal + colchon;

            // ---------- Render ----------
            const card = (v, t, c) => `<div style="flex:1; min-width:150px; padding:12px 16px; background:${c || '#eaf0fb'}; border-radius:8px; text-align:center;">
                <div style="font-size:24px; font-weight:bold; color:#2c3e50;">${v}</div>
                <div style="font-size:11px; color:#7f8c8d;">${t}</div></div>`;

            let out = `<div style="display:flex; gap:14px; flex-wrap:wrap; margin-bottom:14px;">
                ${card(plantaCajeros, 'PLANTA ÓPTIMA de cajeros (tit+emerg)', '#e8f8f0')}
                ${card(plantaEsp, 'Roles especiales: ' + cfg.rolesEspeciales.map(r => r[0]).join(' · ') + ` (${totEspDia}/día)`, '#e8f8f0')}
                ${card(plantaTotal, 'PLANTA TOTAL ideal', '#d5f5e3')}
                ${card((actCajeros + actEsp) || '—', `Planta ACTUAL cargada${actX ? ' (' + actX + ' con X)' : ''}`, (actCajeros + actEsp) >= plantaTotal ? '#e8f8f0' : '#fdecea')}
            </div>`;

            out += `<div style="margin-bottom:14px; padding:12px 16px; background:#eaf2f8; border-left:6px solid #2980b9; border-radius:8px; font-size:12px; color:#1a5276;">
                🧮 <b>CÁLCULO (restricción dominante: ${domCaj}):</b><br>
                · Necesidad L-S: <b>${necSem}</b>/día · Domingo: <b>${necDom}</b>/día. Una persona NO cubre ambos picos (8:00→20:00 = 12h &gt; ${presenciaMax}h de presencia máx.): los grupos se SUMAN.<br>
                · <b>Domingo:</b> ${necDom} ÷ ${disponibleDom * 100}% disponibles = ${fmt1(necDom / disponibleDom)} → <b>${plantaDom}</b><br>
                · <b>Lunes-Viernes:</b> planta − COMP ≥ ${necSem}, con COMP/día = planta × ${disponibleDom * 100}% ÷ 5 → planta ≥ ${necSem} ÷ ${1 - disponibleDom / 5} = ${fmt1(necSem / (1 - disponibleDom / 5))} → <b>${plantaLV}</b> ← dominante<br>
                · <b>Sábado:</b> nadie descansa → ${plantaSab}.&nbsp;&nbsp; <b>Cajeros: ${plantaCajeros}</b> (${Math.round(plantaCajeros * cfg.propTitular)} titulares + ${plantaCajeros - Math.round(plantaCajeros * cfg.propTitular)} emergentes, 3:1)<br>
                · <b>Roles especiales:</b> ${totEspDia}/día → Dom: ${totEspDia} ÷ ${disponibleDom * 100}% = ${fmt1(totEspDia / disponibleDom)} → <b>${plantaEspDom}</b> (${domEsp} dominante) · L-V: ${fmt1(totEspDia / (1 - disponibleDom / 5))} → ${plantaEspLV}. <b>Especiales: ${plantaEsp}</b><br>
                · <b>TOTAL: ${plantaCajeros} + ${plantaEsp} = ${plantaTotal} personas</b></div>`;

            out += `<div style="display:flex; gap:14px; flex-wrap:wrap; margin-bottom:14px;">
                ${card(fmt1(hcSem) + ' / ' + fmt1(hhSem), 'HORAS-CAJA / HORAS-HOMBRE por día L-S (cajeros)')}
                ${card(fmt1(hcDom) + ' / ' + fmt1(hhDom), 'HORAS-CAJA / HORAS-HOMBRE domingo')}
                ${card(fmt1(hcReq), 'Horas-caja REQUERIDAS por las anclas/día')}
                ${card(fmt1(horasCajaEsp), 'Horas-caja de roles especiales/día')}
            </div>`;

            const almSem = cfg.cintaSemana.filter(f => f[2] !== '—').reduce((a, f) => a + f[1], 0);
            const almDom = cfg.cintaDomingo.filter(f => f[2] !== '—').reduce((a, f) => a + f[1], 0);
            out += `<div style="margin-bottom:14px; padding:10px 14px; background:#f8f9fa; border-left:5px solid #7f8c8d; border-radius:6px; font-size:12px; color:#2c3e50;">
                ℹ️ <b>Cómo leer las horas:</b> <b>HORAS-CAJA</b> = horas efectivas con caja abierta (jornada laboral, SIN almuerzo). L-S: ${necSem} personas × 7h = <b>${fmt1(hcSem)}</b>.
                <b>HORAS-HOMBRE</b> = horas de presencia en tienda (CON almuerzo): ${fmt1(hcSem)} + ${almSem} almuerzos de 1h = <b>${fmt1(hhSem)}</b>.
                Domingo: ${fmt1(hcDom)} h-caja + ${almDom} almuerzos = ${fmt1(hhDom)} h-hombre.
                Las anclas del día exigen <b>${fmt1(hcReq)}</b> h-caja → holgura de ${fmt1(hcSem - hcReq)} h-caja para absorber ausencias y desfases de almuerzo.</div>`;
            out += bloque(cfg.cintaSemana, '📋 Bloque de turnos LUNES a SÁBADO (7h)');
            out += cobTabla(cfg.cintaSemana, '🕐 Cobertura L-S vs anclas mínimas de ' + tienda, 'LV');
            out += bloque(cfg.cintaDomingo, '📋 Bloque de turnos DOMINGO (75% de la planta)');
            out += cobTabla(cfg.cintaDomingo, '🕐 Cobertura DOMINGO vs anclas', 'DOM');

            out += `<h4 style="margin:14px 0 6px 0; font-size:13px; color:#2c3e50;">🎯 Roles especiales (cajas obligatorias)</h4>
                <div style="overflow-x:auto;"><table style="border-collapse:collapse; font-size:12px; background:#fff;">
                <thead><tr style="background:#7d3c98; color:#fff; font-size:11px;">
                <th style="padding:4px 10px;">Caja</th><th style="padding:4px 10px;">Horario</th>
                <th style="padding:4px 10px;">Horas-caja</th><th style="padding:4px 10px;">Personas/día</th>
                <th style="padding:4px 10px;">Personas actuales</th>
                <th style="padding:4px 10px;">Nota</th></tr></thead><tbody>${filasRoles}
                <tr style="background:#f4ecf7; font-weight:bold;"><td style="padding:4px 10px;">TOTAL</td><td></td>
                <td style="padding:4px 10px; text-align:center;">${horasCajaEsp} h</td>
                <td style="padding:4px 10px; text-align:center;">${totEspDia}/día → planta ${plantaEsp}</td>
                <td style="padding:4px 10px; text-align:center;">${totalEspActual} personas</td><td></td></tr></tbody></table></div>`;

            out += `<h4 style="margin:14px 0 6px 0; font-size:13px; color:#2c3e50;">🔋 BATERÍA DE CAJAS — semana tipo con planta de ${plantaCajeros} cajeros</h4>
                <div style="overflow-x:auto;"><table style="border-collapse:collapse; font-size:12px; background:#fff;">
                <thead><tr style="background:#2c3e50; color:#fff; font-size:11px;">
                <th style="padding:4px 10px;">Día</th><th style="padding:4px 10px;">Necesidad</th>
                <th style="padding:4px 10px;">Disponibles</th><th style="padding:4px 10px;">Descansan</th>
                <th style="padding:4px 10px;">Holgura</th></tr></thead><tbody>${semanaTipo}</tbody></table></div>
                <div style="font-size:11px; color:#7f8c8d; margin-top:4px;">⚠ Holgura 0 = día sin margen: cualquier ausencia rompe la cinta (por eso el colchón +${colchon}). El ciclo dominical completo se repite cada ${cfg.cicloDomingo} semanas.</div>`;

            out += `<div style="margin-top:12px; padding:10px 14px; background:#fef9e7; border-left:5px solid #f39c12; border-radius:6px; font-size:12px; color:#7d6608;">
                🔄 <b>Deuda de 6h:</b> el domingo hacen 8h <b>${dom8h}</b> personas → la semana siguiente deben absorber ${dom8h} turnos de 6h entre martes y jueves (~${deudaDia}/día). Con ${necSem} cupos/día, representan el ${Math.round(deudaDia / necSem * 100)}% de la cinta: absorbible sin romper cobertura (cada 6h resta 1 hora-caja al día).
                &nbsp;·&nbsp; <b>Horas extra:</b> máx ${cfg.horasExtraMax}h/persona solo emergencias de cierre, tope ${cfg.topeCierreExtra} al cierre.</div>`;

            if (mesSel) {
                out += `<div style="margin-top:12px; padding:10px 14px; background:#eaf2f8; border-left:5px solid #2980b9; border-radius:6px; font-size:12px; color:#1a5276;">
                    🏖️ <b>Vacaciones ${mesSel}:</b> ${vacMes} persona(s) de la pestaña PLANTA con vacaciones en el mes${vacDetalle}${extraVac ? ` + ${extraVac} adicional(es) manual(es)` : ''} → la planta debe subir a <b>${plantaTotal} + ${vacTotal} + ${colchon} = ${recomendada}</b> para sostener la cinta ese mes.
                    ${!pd ? '<br>💡 Abre "👥 Mi Planta" una vez para cargar las fechas de vacaciones desde el Sheets.' : ''}</div>`;
            }

            cont.innerHTML = out;
        }
        window.renderPlantaIdeal = renderPlantaIdeal;

        // Bloque de días de SOLO ORGANIZACIÓN (el personal alcanza y no hay
        // cajeros de vacaciones): no cuentan como incumplimiento, se informan
        // aparte en verde con la observación de reorganizar la malla.
        function _bloqueDiasOrg(diasOrg, bateria, NOMBRES_DIA) {
            if (!diasOrg || !diasOrg.length) return '';
            let filas = diasOrg.map(dm => {
                let j = dm.jf;
                let horas = dm.ev.fallos.map(f => `${f.h}:00`).join(', ');
                return `<tr>
                    <td style="padding:4px 10px; font-weight:bold;">${NOMBRES_DIA[dm.fecha.getDay()]} ${dm.header}</td>
                    <td style="padding:4px 10px; text-align:center;">${dm.ev.fallos.length}</td>
                    <td style="padding:4px 10px; font-size:11px; color:#7f8c8d;">${horas}</td>
                    <td style="padding:4px 10px; text-align:center;">${j.P} pers = ${Math.round(j.S)} h-caja ≥ ${j.D} meta</td>
                    <td style="padding:2px 6px;"><button onclick="verDiaEnCobertura(${dm.colIdx})" style="padding:2px 8px; font-size:10px; font-weight:bold; color:#fff; background:#2980b9; border:none; border-radius:4px; cursor:pointer;">🕒 Ver</button></td>
                  </tr>`;
            }).join('');
            return `<div style="margin-bottom:14px; padding:12px 16px; background:#e8f8f0; border:1px solid #27ae60; border-left:6px solid #27ae60; border-radius:8px;">
                <div style="font-weight:bold; color:#1e8449; font-size:13px; margin-bottom:6px;">✅ ${diasOrg.length} día(s) de SOLO ORGANIZACIÓN (no cuentan como incumplimiento)</div>
                <div style="font-size:12px; color:#1e8449; margin-bottom:6px;">El personal del día SÍ alcanza y NO hay cajeros/emergentes en vacaciones: el hueco es solo de ORGANIZACIÓN de turnos — regenera la malla en el Python para que el optimizador lo acomode.</div>
                <div style="overflow-x:auto;"><table style="border-collapse:collapse; font-size:12px; background:#fff; border-radius:6px;">
                  <thead><tr style="background:#27ae60; color:#fff; font-size:11px;">
                    <th style="padding:4px 10px;">Día</th><th style="padding:4px 10px;">Horas bajo ancla</th>
                    <th style="padding:4px 10px;">Cuáles</th><th style="padding:4px 10px;">Personal vs meta</th>
                    <th style="padding:4px 10px;"></th>
                  </tr></thead><tbody>${filas}</tbody></table></div></div>`;
        }

        // Salta a la línea de tiempo de cobertura con ese día ya seleccionado.
        window.verDiaEnCobertura = function (colIdx) {
            // switchView reconstruye el selector de días: primero cambiamos de
            // vista y DESPUÉS fijamos la fecha para que no la resetee.
            let btn = Array.from(document.querySelectorAll('.nav-item')).find(b => b.textContent.includes('Demanda'));
            if (typeof switchView === 'function') switchView('view-demanda', btn || null);
            let sel = document.getElementById('covDaySelector');
            if (sel) {
                if (!sel.options.length && typeof llenarSelectorDiasCobertura === 'function') llenarSelectorDiasCobertura();
                sel.value = String(colIdx);
                if (typeof renderLineaTiempoCobertura === 'function') renderLineaTiempoCobertura();
            }
            let contTl = document.getElementById('covTimelineContainer');
            if (contTl) contTl.scrollIntoView({ behavior: 'smooth', block: 'start' });
        };
        // ================================================================

        function procesarDiaEspecificoSeguro() {
            let selectElem = document.getElementById('daySelector');
            if(!selectElem.value) return;
            let colIdx = parseInt(selectElem.value);
            let rolesHoy = rolesEspecialesData[colIdx] || {};
            
            let personalHoy = excelData.map(item => ({
                nombre: item.data[0], turno: (item.data[colIdx] || '').toString().trim().toUpperCase(), rol: item.rolReal
            })).filter(p => obtenerRango(p.turno)); 

            let selEsp = document.getElementById('empEspecialSelect');
            selEsp.innerHTML = personalHoy.map(p => {
                let info = getRoleInfo(p.nombre, p.rol);
                return `<option value="${p.nombre}">${info.icon} ${p.nombre} (${info.label})</option>`;
            }).join('');
            document.getElementById('rolesEspecialesPanel').style.display = 'block';
            actualizarListaBadges(colIdx);

            personalGlobal = personalHoy;
            rolesHoyGlobal = rolesHoy;

            let cajeros = personalHoy.filter(p => p.rol === 'CAJERO');
            let empacas = personalHoy.filter(p => p.rol === 'EMPACADOR');
            
            let relevosCaj = generarRelevosAvanzados(cajeros);
            let relevosEmp = generarRelevosAvanzados(empacas);
            relevosGlobales =[...relevosCaj, ...relevosEmp];
            relevosGlobales.forEach((r, index) => r.id = index);
            relevosGlobales.forEach(r => r.reemplazo = ""); 
            relevosGlobales.forEach(r => autoCoberturaInteligente(r));

            actualizarVistaDinamicamente();
        }

        function calcHora(grupo, relevos, roles, h) {
            let prog = 0, enDesc = 0, enEsp = 0;
            let mIni = h * 60, mFin = (h + 1) * 60;
            grupo.forEach(p => {
                let r = obtenerRango(p.turno);
                if (r && h >= r.inicio && h < r.fin) {
                    prog++;
                    if (roles[p.nombre]) enEsp++;
                    relevos.filter(rel => rel.nombre === p.nombre).forEach(rel => {
                        let overlap = Math.max(0, Math.min(mFin, rel.minRegreso) - Math.max(mIni, rel.minSalida));
                        enDesc += (overlap / 60);
                    });
                }
            });
            return { prog, desc: enDesc, esp: enEsp, disp: Math.max(0, Math.round(prog - enDesc - enEsp)) };
        }

        function generarBarraHtml(d) {
            if (d.prog === 0) return '-';
            let pDisp = (d.disp / d.prog) * 100;
            let pEsp = (d.esp / d.prog) * 100;
            let pDesc = (d.desc / d.prog) * 100;
            return `<div class="visual-bar-container">
                <div style="width:${pDisp}%; background:var(--success)"></div>
                <div style="width:${pEsp}%; background:var(--special)"></div>
                <div style="width:${pDesc}%; background:var(--danger)"></div>
            </div><small>${d.disp} de ${d.prog} operando</small>`;
        }

        function obtenerRango(t) {
            if (!t) return null;
            let m = t.match(/(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)/);
            return m ? { inicio: parseInt(m[1]), fin: parseInt(m[2]) } : null;
        }

        function generarRelevosAvanzados(grupo) {
            let relevos =[];
            let mStaff = grupo.filter(p => { let r = obtenerRango(p.turno); return r && r.inicio < 11; }); 
            let midStaff = grupo.filter(p => { let r = obtenerRango(p.turno); return r && r.inicio >= 11 && r.inicio < 13; }); 
            let aStaff = grupo.filter(p => { let r = obtenerRango(p.turno); return r && r.inicio >= 13; }); 

            mStaff.sort((a,b) => obtenerRango(a.turno).inicio - obtenerRango(b.turno).inicio);
            aStaff.sort((a,b) => obtenerRango(a.turno).inicio - obtenerRango(b.turno).inicio);

            let aStaffDisp = Array.from(aStaff);
            let coverageMatrix = {}; 

            function isAvailableToCover(nombre, startMin, endMin) {
                 if(!coverageMatrix[nombre]) return true;
                 for(let m = startMin; m < endMin; m++) if(coverageMatrix[nombre][m]) return false;
                 return true;
            }

            function markBusy(nombre, startMin, endMin) {
                 if(!coverageMatrix[nombre]) coverageMatrix[nombre] = {};
                 for(let m = startMin; m < endMin; m++) coverageMatrix[nombre][m] = true;
            }

            let startM = 8 * 60 + 15; let endM = 10 * 60 + 45; let currentMSalida = startM;
            mStaff.forEach((m) => {
                let rM = obtenerRango(m.turno);
                let minPossibleStart = Math.max(currentMSalida, rM.inicio * 60 + 15);
                if (minPossibleStart > endM - 15) minPossibleStart = endM - 15; 
                let minSalida = minPossibleStart; let minRegreso = minSalida + 15; let cover = null;
                while (minRegreso <= endM) {
                    if (isAvailableToCover(m.nombre, minSalida, minRegreso)) {
                        cover = mStaff.find(c => c.nombre !== m.nombre && (obtenerRango(c.turno).inicio * 60) <= minSalida && isAvailableToCover(c.nombre, minSalida, minRegreso));
                        if (cover) break;
                    }
                    minSalida += 15; minRegreso += 15;
                }
                if (cover && minRegreso <= endM) {
                    relevos.push({nombre: m.nombre, turno: m.turno, minSalida: minSalida, minRegreso: minRegreso, badge: 'badge-15', tipo: '15m Mañana', reemplazo: cover.nombre, rol: m.rol});
                    markBusy(m.nombre, minSalida, minRegreso); markBusy(cover.nombre, minSalida, minRegreso); currentMSalida = minSalida + 15;
                } else {
                    let fallSalida = Math.max(startM, rM.inicio * 60 + 15);
                    while (fallSalida + 15 <= endM && !isAvailableToCover(m.nombre, fallSalida, fallSalida + 15)) fallSalida += 15;
                    if (fallSalida + 15 > endM) fallSalida = startM; 
                    relevos.push({nombre: m.nombre, turno: m.turno, minSalida: fallSalida, minRegreso: fallSalida + 15, badge: 'badge-15', tipo: '15m Mañana', reemplazo: 'Líder / Libre', rol: m.rol});
                    markBusy(m.nombre, fallSalida, fallSalida + 15);
                }
            });

            let allMorning =[...mStaff, ...midStaff];
            allMorning.forEach((m) => {
                let rM = obtenerRango(m.turno); let idealLunchStart = 13; if (rM.inicio >= 9) idealLunchStart = 14;
                let matchIdx = aStaffDisp.findIndex(a => obtenerRango(a.turno).inicio === idealLunchStart);
                if (matchIdx === -1 && rM.inicio >= 10) matchIdx = aStaffDisp.findIndex(a => obtenerRango(a.turno).inicio === 13);
                if (matchIdx === -1) matchIdx = aStaffDisp.findIndex(a => obtenerRango(a.turno).inicio >= 13);
                if (matchIdx === -1 && aStaffDisp.length > 0) matchIdx = 0; 
                let aMatch = null; if (matchIdx !== -1) aMatch = aStaffDisp.splice(matchIdx, 1)[0];

                if (aMatch) {
                    let lunchStart = obtenerRango(aMatch.turno).inicio * 60; 
                    relevos.push({nombre: m.nombre, turno: m.turno, minSalida: lunchStart, minRegreso: lunchStart + 60, badge: 'badge-60', tipo: '1h Almuerzo', reemplazo: aMatch.nombre, rol: m.rol});
                    markBusy(m.nombre, lunchStart, lunchStart + 60); markBusy(aMatch.nombre, lunchStart, lunchStart + 60);
                } else {
                    let lunchStart = (rM.inicio + 4) * 60;
                    relevos.push({nombre: m.nombre, turno: m.turno, minSalida: lunchStart, minRegreso: lunchStart + 60, badge: 'badge-60', tipo: '1h Almuerzo', reemplazo: 'Líder / Libre', rol: m.rol});
                    markBusy(m.nombre, lunchStart, lunchStart + 60);
                }
            });

            let startMidTarde = 17 * 60; let endMidTarde = 18 * 60; let currentMidSalida = startMidTarde;
            midStaff.forEach((a) => {
                let minSalida = currentMidSalida; let minRegreso = minSalida + 15; let cover = null;
                while (minRegreso <= endMidTarde) {
                    if (isAvailableToCover(a.nombre, minSalida, minRegreso)) {
                        cover = grupo.find(m => m.nombre !== a.nombre && (obtenerRango(m.turno).inicio * 60) <= minSalida && (obtenerRango(m.turno).fin * 60) >= minRegreso && isAvailableToCover(m.nombre, minSalida, minRegreso));
                        if (cover) break;
                    }
                    minSalida += 15; minRegreso += 15;
                }
                if (cover && minRegreso <= endMidTarde) {
                    relevos.push({nombre: a.nombre, turno: a.turno, minSalida: minSalida, minRegreso: minRegreso, badge: 'badge-15', tipo: '15m Tarde (11am)', reemplazo: cover.nombre, rol: a.rol});
                    markBusy(a.nombre, minSalida, minRegreso); markBusy(cover.nombre, minSalida, minRegreso); currentMidSalida = minSalida + 15;
                } else {
                    let fallSalida = startMidTarde;
                    while(fallSalida + 15 <= endMidTarde && !isAvailableToCover(a.nombre, fallSalida, fallSalida + 15)) fallSalida += 15;
                    if (fallSalida + 15 > endMidTarde) fallSalida = startMidTarde;
                    relevos.push({nombre: a.nombre, turno: a.turno, minSalida: fallSalida, minRegreso: fallSalida + 15, badge: 'badge-15', tipo: '15m Tarde (11am)', reemplazo: 'Líder / Libre', rol: a.rol});
                    markBusy(a.nombre, fallSalida, fallSalida + 15);
                }
            });

            let startTarde = 15 * 60; let endTarde = 18 * 60;   
            aStaff.forEach((a) => {
                let minPossibleStart = Math.max(startTarde, obtenerRango(a.turno).inicio * 60 + 60); 
                let minSalida = minPossibleStart; let minRegreso = minSalida + 15; let cover = null;
                while(minRegreso <= endTarde) {
                    if(isAvailableToCover(a.nombre, minSalida, minRegreso)) {
                        cover = grupo.find(m => m.nombre !== a.nombre && (obtenerRango(m.turno).inicio * 60) <= minSalida && (obtenerRango(m.turno).fin * 60) >= minRegreso && isAvailableToCover(m.nombre, minSalida, minRegreso));
                        if(cover) break; 
                    }
                    minSalida += 15; minRegreso += 15;
                }
                if(cover && minRegreso <= endTarde) {
                    relevos.push({nombre: a.nombre, turno: a.turno, minSalida: minSalida, minRegreso: minRegreso, badge: 'badge-15', tipo: '15m Tarde', reemplazo: cover.nombre, rol: a.rol});
                    markBusy(a.nombre, minSalida, minRegreso); markBusy(cover.nombre, minSalida, minRegreso);
                } else {
                    let fallSalida = minPossibleStart;
                    while(fallSalida < endTarde && !isAvailableToCover(a.nombre, fallSalida, fallSalida + 15)) fallSalida += 15;
                    if (fallSalida + 15 > endTarde) fallSalida = minPossibleStart;
                    relevos.push({nombre: a.nombre, turno: a.turno, minSalida: fallSalida, minRegreso: fallSalida + 15, badge: 'badge-15', tipo: '15m Tarde', reemplazo: 'Líder / Libre', rol: a.rol});
                    markBusy(a.nombre, fallSalida, fallSalida + 15);
                }
            });
            return relevos;
        }

        function renderizarRelevos(arr) {
            arr.sort((a,b) => a.minSalida - b.minSalida);
            document.getElementById('scheduleBody').innerHTML = arr.map(i => {
                let info = getRoleInfo(i.nombre, i.rol);
                let hS = Math.floor(i.minSalida/60).toString().padStart(2, '0'); let mS = (i.minSalida%60).toString().padStart(2, '0');
                let hR = Math.floor(i.minRegreso/60).toString().padStart(2, '0'); let mR = (i.minRegreso%60).toString().padStart(2, '0');
                return `
                <tr class="${i.badge === 'badge-60' ? 'encadenado-row' : ''}">
                    <td style="text-align: left; padding-left: 15px;"><strong style="color: ${info.color};">${i.nombre}</strong> <span class="${info.class}">${info.icon}</span></td>
                    <td><span style="color:#7f8c8d; font-size:12px; font-weight:bold;">${i.turno}</span></td>
                    <td><span class="badge ${i.badge}">${i.tipo}</span></td>
                    <td><input type="time" value="${hS}:${mS}" onchange="modificarHora(${i.id}, 'salida', this.value)" style="border:1px solid #ccc; border-radius:4px; padding:3px; color:var(--danger); font-weight:bold; cursor:pointer;"></td>
                    <td><input type="time" value="${hR}:${mR}" onchange="modificarHora(${i.id}, 'regreso', this.value)" style="border:1px solid #ccc; border-radius:4px; padding:3px; color:var(--success); font-weight:bold; cursor:pointer;"></td>
                    <td style="color:var(--secondary); font-weight:bold; font-size:12px;" id="cobertura-${i.id}">👤 ${i.reemplazo}</td>
                </tr>`;
            }).join('');
        }

        function minutosAHora(m) { return `${Math.floor(m/60)}:${m%60 === 0 ? '00' : m%60}`; }

        function renderChart(labelsX, c, e, d) {
            let ctx = document.getElementById('dailyChart').getContext('2d');
            if (chartDiario) chartDiario.destroy();
            chartDiario = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: labelsX,
                    datasets:[
                        { label: 'Cajeros Disponibles', data: c, borderColor: '#2e7d32', backgroundColor: '#2e7d3220', fill: true, tension: 0.3 },
                        { label: 'Empacadores Disponibles', data: e, borderColor: '#1565c0', backgroundColor: '#1565c020', fill: true, tension: 0.3 }
                    ]
                },
                options: { 
                    responsive: true, maintainAspectRatio: false,
                    scales: {
                        x: {
                            ticks: {
                                color: function(ctx) { return ctx.tick && ctx.tick.label && ctx.tick.label.includes('🚨') ? '#c0392b' : '#333'; },
                                font: function(ctx) { return ctx.tick && ctx.tick.label && ctx.tick.label.includes('🚨') ? {weight: 'bold', size: 14} : {}; }
                            }
                        }
                    }
                }
            });
        }

        function renderVentasChart() {
            let container = document.getElementById('ventasChartContainer');
            if (!window.dataIA || window.dataIA.length === 0) { container.style.display = 'none'; return; }

            let firstRow = window.dataIA[0];
            let keyVenta = Object.keys(firstRow).find(k => {
                let c = k.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
                return c.includes('venta');
            });
            let keyDia = Object.keys(firstRow).find(k => {
                let c = k.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
                return c === 'dia' || c.startsWith('dia') || c === 'day';
            });
            let keyHora = Object.keys(firstRow).find(k => k.toLowerCase().includes('hora'));
            let keyFecha = Object.keys(firstRow).find(k => k.toLowerCase().includes('fecha'));

            if (!keyVenta || !keyHora) { container.style.display = 'none'; return; }

            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            let selectElem = document.getElementById('daySelector');
            let selectedIdx = selectElem.value ? parseInt(selectElem.value) - 1 : 0;
            let d = new Date(startDate);
            d.setDate(d.getDate() + selectedIdx);

            let yyyy = d.getFullYear(), mm = String(d.getMonth()+1).padStart(2,'0'), dd = String(d.getDate()).padStart(2,'0');
            let yy = yyyy.toString().slice(-2), mS = (d.getMonth()+1).toString(), dS = d.getDate().toString();
            let fmts = [
                `${yyyy}-${mm}-${dd}`, `${yyyy}/${mm}/${dd}`, `${dd}/${mm}/${yyyy}`,
                `${dS}/${mS}/${yyyy}`, `${mS}/${dS}/${yyyy}`,
                `${dd}-${mm}-${yyyy}`, `${mS}/${dS}/${yy}`, `${dS}/${mS}/${yy}`
            ];

            let diasNombres = ['dom','lun','mar','mie','jue','vie','sab'];
            let diaSelIdx = d.getDay();
            let diaSelLabel = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'][diaSelIdx];

            function matchFecha(row) {
                let rowDate = keyFecha ? (row[keyFecha] || '').toString().trim() : '';
                if (fmts.some(f => rowDate.includes(f))) return true;
                if (!isNaN(rowDate) && rowDate !== '' && parseFloat(rowDate) > 1000) {
                    let jsD = new Date((parseFloat(rowDate) - (25567+2)) * 86400 * 1000);
                    return jsD.getFullYear()===yyyy && jsD.getMonth()===d.getMonth() && jsD.getDate()===d.getDate();
                }
                return false;
            }

            function getHora(row) {
                let rawH = (row[keyHora] || '').toString().trim();
                if (!isNaN(rawH) && rawH !== '' && parseFloat(rawH) < 1) return Math.floor(parseFloat(rawH) * 24);
                let m = rawH.match(/^(\d+)/);
                return m ? parseInt(m[1]) : -1;
            }

            function getVenta(row) {
                return parseFloat((row[keyVenta] || '').toString().trim().replace(',','.')) || 0;
            }

            function sameDayOfWeek(row) {
                if (keyDia) {
                    let rawDia = (row[keyDia] || '').toString().trim().normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase().substring(0,3);
                    return rawDia === diasNombres[diaSelIdx];
                }
                if (keyFecha) {
                    let rawDate = (row[keyFecha] || '').toString().trim();
                    let jsD = null;
                    let parsed = fmts.map((_, i) => i); // attempt via date parse
                    // try common formats
                    let parts = rawDate.match(/(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
                    if (parts) {
                        let a = parseInt(parts[1]), b = parseInt(parts[2]), c = parseInt(parts[3]);
                        jsD = a > 31 ? new Date(a, b-1, c) : new Date(c < 100 ? 2000+c : c, b-1, a);
                    }
                    if (!isNaN(rawDate) && parseFloat(rawDate) > 1000) jsD = new Date((parseFloat(rawDate)-(25567+2))*86400*1000);
                    return jsD && jsD.getDay() === diaSelIdx;
                }
                return false;
            }

            let currentDay = {}, sameDayAccum = {};
            window.dataIA.forEach(row => {
                let h = getHora(row);
                if (h < 8 || h > 21) return;
                if (matchFecha(row)) {
                    currentDay[h] = getVenta(row);
                } else if (sameDayOfWeek(row)) {
                    if (!sameDayAccum[h]) sameDayAccum[h] = [];
                    sameDayAccum[h].push(getVenta(row));
                }
            });

            let labels = horasEjeX.map(h => h + ':00');
            let dataCurrent = horasEjeX.map(h => currentDay[h] !== undefined ? currentDay[h] : null);
            let dataAvg = horasEjeX.map(h => {
                let arr = sameDayAccum[h];
                if (!arr || arr.length === 0) return null;
                return Math.round(arr.reduce((a,b) => a+b, 0) / arr.length * 10) / 10;
            });

            let hasCurrentData = dataCurrent.some(v => v !== null);
            let hasAvgData = dataAvg.some(v => v !== null);
            if (!hasCurrentData && !hasAvgData) { container.style.display = 'none'; return; }

            container.style.display = 'block';
            let ctx2 = document.getElementById('ventasChart').getContext('2d');
            if (chartVentas) chartVentas.destroy();
            chartVentas = new Chart(ctx2, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: [
                        {
                            label: `Ventas hoy (${diaSelLabel} ${d.getDate()}/${d.getMonth()+1})`,
                            data: dataCurrent,
                            borderColor: '#e67e22',
                            backgroundColor: 'rgba(230,126,34,0.12)',
                            fill: true, tension: 0.35, borderWidth: 3, pointRadius: 4
                        },
                        {
                            label: `Promedio ${diaSelLabel}s histórico`,
                            data: dataAvg,
                            borderColor: '#8e44ad',
                            backgroundColor: 'rgba(142,68,173,0.08)',
                            fill: true, tension: 0.35, borderWidth: 2,
                            borderDash: [6, 4], pointRadius: 3
                        }
                    ]
                },
                options: {
                    responsive: true, maintainAspectRatio: false,
                    plugins: { legend: { position: 'top' } },
                    scales: {
                        y: { beginAtZero: true, title: { display: true, text: 'Ventas' } }
                    }
                }
            });
        }

        function agregarRolEspecial() {
            let idx = document.getElementById('daySelector').value; let n = document.getElementById('empEspecialSelect').value; let r = document.getElementById('rolEspecialSelect').value;
            if(!n) return; if(!rolesEspecialesData[idx]) rolesEspecialesData[idx] = {}; rolesEspecialesData[idx][n] = r; procesarDiaEspecificoSeguro();
        }

        function actualizarListaBadges(idx) {
            let container = document.getElementById('listaRolesEspeciales'); container.innerHTML = ''; let roles = rolesEspecialesData[idx] || {};
            Object.entries(roles).forEach(([nombre, rol]) => { container.innerHTML += `<span class="badge-role">${nombre} - ${rol} <button class="btn-remove-role" onclick="removerRol('${idx}','${nombre}')">×</button></span>`; });
        }

        function removerRol(idx, nombre) { delete rolesEspecialesData[idx][nombre]; procesarDiaEspecificoSeguro(); }

        // Hora decimal (12.5) -> "12:30"
        function _fmtHDec(x) { let h = Math.floor(x); let m = Math.round((x - h) * 60); return h + ':' + String(m).padStart(2, '0'); }

        // Vista PLANILLA autónoma: día propio + filtro por rol. El ALMUERZO se toma
        // de la misma lógica que "Demanda vs Personal" (_almuerzosDelDia, 12:00-16:00).
        function renderPlanillaVista() {
            let container = document.getElementById('planillaContainer');
            if (!container) return;
            if (!excelData || !excelData.length) { container.innerHTML = '<div style="color:#7f8c8d; padding:12px;">Carga la Matriz de Turnos para ver la planilla.</div>'; return; }

            // Poblar el selector de día (si está vacío) y el filtro de rol
            let selDia = document.getElementById('daySelectorPlanilla');
            if (selDia && !selDia.options.length) {
                let sv = document.getElementById('fechaInicio').value;
                let sd = new Date(sv ? (sv + 'T12:00:00') : new Date());
                headers.forEach((h, i) => {
                    if (h === 'Total Horas Periodo') return;
                    let d = new Date(sd); d.setDate(d.getDate() + i);
                    selDia.innerHTML += `<option value="${i + 1}">${h} (${d.toLocaleDateString()})</option>`;
                });
            }
            let colIdx = selDia && selDia.value ? parseInt(selDia.value) : 1;

            let selRol = document.getElementById('planillaRolFilter');
            if (selRol) {
                let cats = [...new Set(excelData.map(it => _rolCategoria(it)))];
                const ORD = ['Cajero', 'Emergente', 'Satélite', 'Empacador', 'Cambista', 'Visado', 'Fundación', 'SADOFE'];
                cats.sort((a, b) => ORD.indexOf(a) - ORD.indexOf(b));
                let cur = selRol.value || 'TODOS';
                selRol.innerHTML = '<option value="TODOS">Todos los roles</option>' +
                    cats.map(c => `<option value="${c}"${c === cur ? ' selected' : ''}>${c}</option>`).join('');
            }
            let rolFiltro = selRol ? selRol.value : 'TODOS';

            // Almuerzos del día (misma lógica que Demanda vs Personal)
            let sv = document.getElementById('fechaInicio').value;
            let sd = new Date(sv ? (sv + 'T12:00:00') : new Date());
            let fecha = new Date(sd); fecha.setDate(fecha.getDate() + (colIdx - 1));
            let almuerzoDe = _almuerzosDelDia(colIdx, fecha.getDay() === 0);

            // Descansos de 15m (del motor de relevos del día completo)
            let descansoDe = {};
            try {
                ['CAJERO', 'EMPACADOR'].forEach(rr => {
                    let grupo = excelData
                        .filter(it => it.rolReal === rr && obtenerRango((it.data[colIdx] || '').toString().trim().toUpperCase()))
                        .map(it => ({ nombre: it.data[0], turno: (it.data[colIdx] || '').toString().trim().toUpperCase(), rol: rr }));
                    (generarRelevosAvanzados(grupo) || []).forEach(r => {
                        if (r.tipo && r.tipo.toLowerCase().includes('15m'))
                            (descansoDe[r.nombre] = descansoDe[r.nombre] || []).push(`${minutosAHora(r.minSalida)} - ${minutosAHora(r.minRegreso)}`);
                    });
                });
            } catch (e) { /* sin descansos */ }

            // Personal del día con turno horario, filtrado por rol
            let personal = excelData.map(item => ({
                nombre: item.data[0], turno: (item.data[colIdx] || '').toString().trim().toUpperCase(),
                rol: item.rolReal, cat: _rolCategoria(item)
            })).filter(p => obtenerRango(p.turno));
            if (rolFiltro !== 'TODOS') personal = personal.filter(p => p.cat === rolFiltro);

            generarPlanillaDiaria(personal, almuerzoDe, descansoDe);
        }
        window.renderPlanillaVista = renderPlanillaVista;

        function generarPlanillaDiaria(personal, almuerzoDe, descansoDe) {
            let container = document.getElementById('planillaContainer'); if (!container) return;
            almuerzoDe = almuerzoDe || {}; descansoDe = descansoDe || {};
            if (!personal || !personal.length) { container.innerHTML = '<div style="color:#7f8c8d; padding:12px;">No hay personal con turno horario para el día/rol seleccionado.</div>'; return; }
            let agrupados = {}; personal.forEach(p => { let t = p.turno || 'SIN TURNO'; if (!agrupados[t]) agrupados[t] = []; agrupados[t].push(p); });
            let turnosOrdenados = Object.keys(agrupados).sort((a, b) => { let rA = obtenerRango(a); let rB = obtenerRango(b); if (rA && rB) return rA.inicio - rB.inicio; return a.localeCompare(b); });
            const tdSt = 'padding:2px 6px; line-height:1.15;';   // celdas compactas (menos altas)
            let html = '';
            turnosOrdenados.forEach(t => {
                html += `<div style="margin-bottom: 14px; border: 1px solid #0277bd; border-radius: 5px; overflow: hidden;"><div style="background-color: #0277bd; color: white; padding: 5px 10px; font-weight: bold; font-size: 13px;">TURNO: ${t} (${agrupados[t].length})</div><div class="table-container" style="margin-bottom: 0; border: none; max-height: none;"><table class="planilla-table" style="width: 100%; font-size:12px;"><thead><tr><th style="${tdSt}">Nombre</th><th style="${tdSt}">Turno</th><th style="${tdSt}">Cerco</th><th style="${tdSt}">Caja-ROL</th><th style="${tdSt}">Hora de Almuerzo</th><th style="${tdSt}">Hora de Descanso</th><th style="${tdSt}">Observación</th></tr></thead><tbody>`;
                agrupados[t].forEach(p => {
                    let a = almuerzoDe[p.nombre];
                    let txtAlmuerzo = a ? `${_fmtHDec(a[0])} - ${_fmtHDec(a[1])}` : '-';
                    let txtDescanso = (descansoDe[p.nombre] || []).join(', ') || '-';
                    let info = getRoleInfo(p.nombre, p.rol);
                    html += `<tr><td class="text-left" style="${tdSt} color:${info.color};">${p.nombre}</td><td style="${tdSt}">${p.turno}</td><td style="${tdSt}"><input type="text" class="planilla-input" placeholder="..."></td><td style="${tdSt}"><input type="text" class="planilla-input" value="${info.label}"></td><td style="${tdSt} font-weight:600; color:#2980b9;">${txtAlmuerzo}</td><td style="${tdSt} font-weight:600; color:#e67e22;">${txtDescanso}</td><td style="${tdSt}"><input type="text" class="planilla-input" placeholder="..."></td></tr>`;
                });
                html += `</tbody></table></div></div>`;
            });
            container.innerHTML = html;
        }

        // ================== CRUCE DE IA EXTREMADAMENTE FLEXIBLE ==================
        function actualizarVistaDinamicamente() {
            renderizarRelevos(relevosGlobales);

            let statsCaj =[], statsEmp =[];
            let tbody = document.getElementById('hourlyBody'); 
            tbody.innerHTML = '';

            let cajeros = personalGlobal.filter(p => p.rol === 'CAJERO');
            let empacas = personalGlobal.filter(p => p.rol === 'EMPACADOR');

            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());
            let selectElem = document.getElementById('daySelector');
            let selectedIdx = selectElem.value ? parseInt(selectElem.value) - 1 : 0; 
            let d = new Date(startDate);
            d.setDate(d.getDate() + selectedIdx);

            // Creamos TODOS los formatos posibles de fecha que podría escupir Excel
            let yyyy = d.getFullYear();
            let mm = String(d.getMonth() + 1).padStart(2, '0');
            let dd = String(d.getDate()).padStart(2, '0');
            let yy = yyyy.toString().slice(-2);
            let mSingle = (d.getMonth() + 1).toString();
            let dSingle = d.getDate().toString();

            let formatosFecha =[
                `${yyyy}-${mm}-${dd}`, `${yyyy}/${mm}/${dd}`, `${dd}/${mm}/${yyyy}`, 
                `${dSingle}/${mSingle}/${yyyy}`, `${mSingle}/${dSingle}/${yyyy}`, 
                `${dd}-${mm}-${yyyy}`, `${mSingle}/${dSingle}/${yy}`, `${dSingle}/${mSingle}/${yy}`
            ];

            let diaTexto = selectElem.options[selectElem.selectedIndex].text;
            let chartLabels =[];

            horasEjeX.forEach(h => {
                let dCaj = calcHora(cajeros, relevosGlobales, rolesHoyGlobal, h);
                let dEmp = calcHora(empacas, relevosGlobales, rolesHoyGlobal, h);
                statsCaj.push(dCaj.disp); statsEmp.push(dEmp.disp);

                let alertaEmergentes = false; let alertaApoyo = false; let textoAlerta =[];

                if (window.dataIA && window.dataIA.length > 0) {
                    let match = window.dataIA.find(row => {
                        let keyFecha = Object.keys(row).find(k => k.toLowerCase().includes('fecha'));
                        let keyHora = Object.keys(row).find(k => k.toLowerCase().includes('hora'));
                        
                        let rowDate = keyFecha ? (row[keyFecha] || '').toString().trim() : '';
                        let rowHora = keyHora ? (row[keyHora] || '').toString().trim() : '';
                        
                        // 1. Verificación de Fecha en 8 formatos distintos
                        let dateMatches = formatosFecha.some(formato => rowDate.includes(formato));
                        
                        // 2. Captura si Excel mandó la fecha en número de serie nativo
                        if (!dateMatches && !isNaN(rowDate) && rowDate !== "" && rowDate > 1000) {
                            let jsDate = new Date((rowDate - (25567 + 2)) * 86400 * 1000);
                            if (jsDate.getFullYear() === yyyy && jsDate.getMonth() === d.getMonth() && jsDate.getDate() === d.getDate()) {
                                dateMatches = true;
                            }
                        }

                        // 3. Captura de Hora
                        let horaMatches = false;
                        let isDecimalTime = !isNaN(rowHora) && rowHora !== "" && rowHora < 1;
                        if (isDecimalTime) {
                            let totalMinutes = Math.round(rowHora * 24 * 60);
                            let parsedH = Math.floor(totalMinutes / 60);
                            if (parsedH === h) horaMatches = true;
                        } else {
                            if (rowHora.startsWith(h + ":") || rowHora === h.toString() || rowHora.startsWith('0' + h + ':')) {
                                horaMatches = true;
                            }
                        }
                        return dateMatches && horaMatches;
                    });

                    if (match) {
                        let key7 = Object.keys(match).find(k => {
                            let clean = k.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                            return clean.includes("7. accion") || clean.includes("emergentes");
                        });
                        let key8 = Object.keys(match).find(k => {
                            let clean = k.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
                            return clean.includes("8. accion") || clean.includes("externo");
                        });
                        
                        let col7 = key7 ? (match[key7] || '').toString().trim().toUpperCase() : '';
                        let col8 = key8 ? (match[key8] || '').toString().trim().toUpperCase() : '';
                        
                        if (col7.includes('SÍ') || col7.includes('SI')) { alertaEmergentes = true; textoAlerta.push("🚨 EME: " + col7); }
                        if (col8.includes('SÍ') || col8.includes('SI')) { alertaApoyo = true; textoAlerta.push("⚠️ EXT: " + col8); }
                    }
                }

                let hasAlert = alertaEmergentes || alertaApoyo;
                chartLabels.push(h + ":00" + (hasAlert ? " 🚨" : ""));
                let cellAlertaHtml = hasAlert ? `<span style="color:#c0392b; font-weight:bold; font-size:11px;">${textoAlerta.join('<br>')}</span>` : `<span style="color:#bdc3c7;">-</span>`;

                tbody.innerHTML += `<tr><td style="${hasAlert ? 'background:#fee2e2; font-weight:bold; color:#c0392b;' : ''}">${h}:00</td><td>${generarBarraHtml(dCaj)}</td><td>${generarBarraHtml(dEmp)}</td><td style="text-align: left;">${cellAlertaHtml}</td></tr>`;
            });
            
            renderChart(chartLabels, statsCaj, statsEmp, diaTexto);
            renderVentasChart();
            if (typeof renderPlanillaVista === 'function') renderPlanillaVista();   // planilla autónoma
        }

        function modificarHora(id, tipo, nuevoValor) {
            let idx = relevosGlobales.findIndex(r => r.id === id); if (idx === -1) return;
            let partes = nuevoValor.split(':'); let nuevosMinutos = parseInt(partes[0]) * 60 + parseInt(partes[1]); let relevo = relevosGlobales[idx];
            if (tipo === 'salida') { relevo.minSalida = nuevosMinutos; let duracion = relevo.tipo.includes('1h') ? 60 : 15; relevo.minRegreso = nuevosMinutos + duracion;
            } else { relevo.minRegreso = nuevosMinutos; }
            autoCoberturaInteligente(relevo); actualizarVistaDinamicamente();
        }

        function getRolGenericoDeMemoria(nombre) { let obj = excelData.find(e => e.data[0] === nombre); return obj ? obj.rolGenerico : ''; }

        function autoCoberturaInteligente(relevo) {
            let reqGenerico = getRolGenericoDeMemoria(relevo.nombre).toUpperCase(); let reqStr = relevo.nombre.toUpperCase() + ' ' + reqGenerico;
            let isCambista = reqStr.includes('CAMBISTA'); let isVisado = reqStr.includes('VISADO'); let isEmergente = reqStr.includes('EMERGENTE'); let isFundacion = reqStr.includes('FUNDACIÓN') || reqStr.includes('FUNDACION');
            let isCajeroNormal = !isCambista && !isVisado && !isEmergente && relevo.rol === 'CAJERO'; let isEmpacadorNormal = !isEmergente && !isFundacion && relevo.rol === 'EMPACADOR';

            let posibles = personalGlobal.filter(p => {
                if (p.nombre === relevo.nombre) return false; 
                let rP = obtenerRango(p.turno); if (!rP) return false;
                let trabajando = (rP.inicio * 60) <= relevo.minSalida && (rP.fin * 60) >= relevo.minRegreso; if (!trabajando) return false;
                let enDescansoPropio = relevosGlobales.some(otro => { if (otro.nombre !== p.nombre) return false; return Math.max(relevo.minSalida, otro.minSalida) < Math.min(relevo.minRegreso, otro.minRegreso); }); if (enDescansoPropio) return false;
                let estaCubriendo = relevosGlobales.some(otro => { if (otro.reemplazo !== p.nombre) return false; return Math.max(relevo.minSalida, otro.minSalida) < Math.min(relevo.minRegreso, otro.minRegreso); }); if (estaCubriendo) return false;
                return true;
            });

            function getScore(candNombre) {
                let cGenerico = getRolGenericoDeMemoria(candNombre).toUpperCase(); let cName = candNombre.toUpperCase() + ' ' + cGenerico;
                let cCambista = cName.includes('CAMBISTA'); let cVisado = cName.includes('VISADO'); let cEmergente = cName.includes('EMERGENTE'); let cFundacion = cName.includes('FUNDACIÓN') || cName.includes('FUNDACION');
                let cCajeroNormal = cName.includes('CAJERO') && !cCambista && !cVisado && !cEmergente; let cEmpacadorNormal = cName.includes('EMPACADOR') && !cEmergente && !cFundacion;
                if (isCambista) { if (cCambista) return 1; if (cCajeroNormal) return 2; return 99; }
                if (isVisado) { if (cVisado) return 1; if (cCajeroNormal) return 2; return 99; }
                if (isCajeroNormal) { if (cCajeroNormal) return 1; if (cEmergente) return 2; return 99; }
                if (isEmergente) { if (cEmergente) return 1; if (cEmpacadorNormal) return 2; if (cFundacion) return 3; return 99; }
                if (isEmpacadorNormal || isFundacion) { if (cEmpacadorNormal) return 1; if (cFundacion) return 2; return 99; } return 99; 
            }

            let candidatosValidos = posibles.map(p => { return { obj: p, score: getScore(p.nombre) }; }).filter(item => item.score < 99).sort((a, b) => a.score - b.score);
            if (candidatosValidos.length > 0) { relevo.reemplazo = candidatosValidos[0].obj.nombre;
            } else {
                if (isCambista) relevo.reemplazo = 'Requiere Camb/Caj / Libre';
                else if (isVisado) relevo.reemplazo = 'Requiere Vis/Caj / Libre';
                else { let esTurno8 = obtenerRango(relevo.turno).inicio === 8; if (esTurno8) { relevo.reemplazo = isCajeroNormal ? 'Cajero Emergente' : 'Otro Empacador'; } else { relevo.reemplazo = 'Líder / Libre'; } }
            }
        }

        function exportarExcelProgramacion() {
            if (!excelData || excelData.length === 0) { alert('No hay datos para exportar. Carga primero la Matriz de Turnos.'); return; }
            if (typeof XLSX === 'undefined') { alert('La librería XLSX no está disponible.'); return; }

            let startVal = document.getElementById('fechaInicio').value;
            let startDate = new Date(startVal ? (startVal + 'T12:00:00') : new Date());

            // Columnas reales (sin "Total Horas Periodo")
            let colIndices = []; // índice i dentro de item.data
            let colHeaders = []; // etiquetas de columna con fecha
            headers.forEach((h, idx) => {
                if (h === 'Total Horas Periodo') return;
                let d = new Date(startDate);
                d.setDate(d.getDate() + idx);
                let tipo = getTipoDia(d);
                let sufijo = tipo === 'festivo' ? ' *' : tipo === 'domingo' ? ' (D)' : tipo === 'sabado' ? ' (S)' : '';
                colHeaders.push(`${h} ${d.getDate()}/${d.getMonth()+1}${sufijo}`);
                colIndices.push(idx + 1); // data[0] es nombre, data[1..] son turnos
            });

            // ===== HOJA 1: PROGRAMACIÓN =====
            let ws1Rows = [['Rol', 'Nombre', ...(hayColumnaNomina ? ['Código Nómina'] : []), ...colHeaders]];
            excelData.forEach(item => {
                let fila = [item.rol || '', item.data[0]];
                if (hayColumnaNomina) fila.push(item.codigoNomina || '');
                colIndices.forEach(ci => fila.push((item.data[ci] || '').toString().trim().toUpperCase()));
                ws1Rows.push(fila);
            });
            let ws1 = XLSX.utils.aoa_to_sheet(ws1Rows);
            ws1['!cols'] = [{ wch: 16 }, { wch: 26 }, ...(hayColumnaNomina ? [{ wch: 14 }] : []), ...colHeaders.map(() => ({ wch: 11 }))];

            // ===== HOJA 2: ANÁLISIS DE EQUIDAD =====
            // 1. Detectar todos los turnos-hora únicos (excluir vacío, LIBRE, COMP, VC, LIC, INC)
            let codigosEspeciales = new Set(['', 'LIBRE', 'COMP', 'VC', 'LIC', 'INC']);
            let shiftsSet = new Set();
            excelData.forEach(item => {
                colIndices.forEach(ci => {
                    let v = (item.data[ci] || '').toString().trim().toUpperCase();
                    if (!codigosEspeciales.has(v)) shiftsSet.add(v);
                });
            });
            let shiftsOrdenados = Array.from(shiftsSet).sort((a, b) => {
                let rA = a.match(/^(\d+)/), rB = b.match(/^(\d+)/);
                if (rA && rB) return parseInt(rA[1]) - parseInt(rB[1]);
                if (rA) return -1; if (rB) return 1;
                return a.localeCompare(b);
            });

            let cabecera2 = [
                'Nombre',
                ...shiftsOrdenados,
                'COMP Lun', 'COMP Mar', 'COMP Mié', 'COMP Jue', 'COMP Vie', 'COMP Sáb', 'Total COMP',
                'LIBRE Dom', 'LIBRE Otros', 'Total LIBRE',
                'VC', 'LIC', 'INC',
                'Días Trabajados', 'Días Totales'
            ];

            let ws2Rows = [cabecera2];

            excelData.forEach(item => {
                let conteoTurnos = {};
                let compPorDia = { 1:0, 2:0, 3:0, 4:0, 5:0, 6:0 }; // Lun-Sáb
                let libreDom = 0, libreOtros = 0;
                let vc = 0, lic = 0, inc = 0, diasTrab = 0, diasTotal = 0;

                colIndices.forEach((ci, colPos) => {
                    let v = (item.data[ci] || '').toString().trim().toUpperCase();
                    if (v === '') return;
                    diasTotal++;
                    let dCol = new Date(startDate);
                    dCol.setDate(dCol.getDate() + colPos);
                    let dow = dCol.getDay(); // 0=Dom … 6=Sáb

                    if (v === 'COMP') {
                        if (dow >= 1 && dow <= 6) compPorDia[dow]++;
                    } else if (v === 'LIBRE') {
                        if (dow === 0) libreDom++; else libreOtros++;
                    } else if (v === 'VC')  { vc++;
                    } else if (v === 'LIC') { lic++;
                    } else if (v === 'INC') { inc++;
                    } else {
                        conteoTurnos[v] = (conteoTurnos[v] || 0) + 1;
                        diasTrab++;
                    }
                });

                let totalComp = Object.values(compPorDia).reduce((a,b) => a+b, 0);
                let totalLibre = libreDom + libreOtros;

                let fila = [item.data[0]];
                shiftsOrdenados.forEach(s => fila.push(conteoTurnos[s] || 0));
                fila.push(compPorDia[1], compPorDia[2], compPorDia[3], compPorDia[4], compPorDia[5], compPorDia[6], totalComp);
                fila.push(libreDom, libreOtros, totalLibre);
                fila.push(vc, lic, inc, diasTrab, diasTotal);
                ws2Rows.push(fila);
            });

            // Fila de TOTALES al final
            let totales = ['TOTAL EQUIPO'];
            for (let col = 1; col < cabecera2.length; col++) {
                let sum = ws2Rows.slice(1).reduce((acc, row) => acc + (Number(row[col]) || 0), 0);
                totales.push(sum);
            }
            ws2Rows.push(totales);

            let ws2 = XLSX.utils.aoa_to_sheet(ws2Rows);
            ws2['!cols'] = [{ wch: 28 }, ...cabecera2.slice(1).map(() => ({ wch: 10 }))];

            // Crear libro y descargar
            let wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws1, 'Programación');
            XLSX.utils.book_append_sheet(wb, ws2, 'Análisis de Equidad');

            let fechaHoy = new Date().toISOString().split('T')[0];
            XLSX.writeFile(wb, `Programacion_Kronomercado_${fechaHoy}.xlsx`);
        }

        function exportarPlanillaPDF() {
            if (typeof html2pdf === 'undefined') { alert("La librería de exportación PDF no está disponible."); return; }
            let element = document.getElementById('seccionPlanilla');
            let inputs = element.querySelectorAll('input'); inputs.forEach(input => { input.setAttribute('value', input.value); });
            let diaSeleccionado = document.getElementById('daySelector'); let txtDia = diaSeleccionado ? diaSeleccionado.options[diaSeleccionado.selectedIndex].text : 'Dia';

            // Inyectar estilos temporales para alinear las tablas correctamente en el PDF
            let styleTag = document.createElement('style');
            styleTag.id = '__pdf_print_style__';
            styleTag.textContent = `
                #seccionPlanilla table.planilla-table {
                    table-layout: fixed !important;
                    width: 100% !important;
                    border-collapse: collapse !important;
                    white-space: normal !important;
                    font-size: 10px !important;
                    word-break: break-word !important;
                }
                #seccionPlanilla table.planilla-table th,
                #seccionPlanilla table.planilla-table td {
                    word-wrap: break-word !important;
                    overflow-wrap: break-word !important;
                    white-space: normal !important;
                    padding: 2px 4px !important;
                    line-height: 1.1 !important;
                    border: 1px solid #ccc !important;
                    vertical-align: middle !important;
                    overflow: hidden !important;
                }
                #seccionPlanilla .planilla-input { padding: 1px 3px !important; height: auto !important; }
                #seccionPlanilla table.planilla-table th:nth-child(1),
                #seccionPlanilla table.planilla-table td:nth-child(1) { width: 21% !important; text-align: left !important; }
                #seccionPlanilla table.planilla-table th:nth-child(2),
                #seccionPlanilla table.planilla-table td:nth-child(2) { width: 9% !important; }
                #seccionPlanilla table.planilla-table th:nth-child(3),
                #seccionPlanilla table.planilla-table td:nth-child(3) { width: 7% !important; }
                #seccionPlanilla table.planilla-table th:nth-child(4),
                #seccionPlanilla table.planilla-table td:nth-child(4) { width: 12% !important; }
                #seccionPlanilla table.planilla-table th:nth-child(5),
                #seccionPlanilla table.planilla-table td:nth-child(5) { width: 13% !important; }
                #seccionPlanilla table.planilla-table th:nth-child(6),
                #seccionPlanilla table.planilla-table td:nth-child(6) { width: 13% !important; }
                #seccionPlanilla table.planilla-table th:nth-child(7),
                #seccionPlanilla table.planilla-table td:nth-child(7) { width: 25% !important; }
                #seccionPlanilla .planilla-input {
                    width: 100% !important;
                    box-sizing: border-box !important;
                    border: none !important;
                    background: transparent !important;
                    font-size: 10px !important;
                    padding: 0 !important;
                    text-align: center !important;
                }
                #seccionPlanilla .table-container {
                    overflow: visible !important;
                    max-height: none !important;
                    border: none !important;
                }
            `;
            document.head.appendChild(styleTag);

            let opt = {
                margin: [0.3, 0.25, 0.3, 0.25],
                filename: `Planilla_Programacion_${txtDia.replace(/[/\\?%*:|"<>]/g, '-')}.pdf`,
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true, letterRendering: true, scrollX: 0, scrollY: 0 },
                jsPDF: { unit: 'in', format: 'legal', orientation: 'landscape' }
            };

            let contenedoresTabla = element.querySelectorAll('.table-container');
            contenedoresTabla.forEach(ct => { ct.style.maxHeight = 'none'; ct.style.overflow = 'visible'; });

            html2pdf().set(opt).from(element).save().then(() => {
                contenedoresTabla.forEach(ct => { ct.style.maxHeight = ''; ct.style.overflow = ''; ct.style.overflowX = 'auto'; });
                let st = document.getElementById('__pdf_print_style__');
                if (st) st.remove();
            });
        }
    