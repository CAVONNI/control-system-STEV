/* ================================================
   VALIDAR LPN — Lógica principal
   ================================================ */

   let datosExcel = [];       // Excel principal (LPN, REFERENCIA, CANTIDAD, NOTA, TIPO...)
   let listaTipoExcel = [];   // Excel de cruce (NOTA -> TIPO/ASN), como lista [{notaNumeros, tipo}]
   
   // ── CARGA DEL EXCEL PRINCIPAL ──
   document.getElementById("excelFile").addEventListener("change", function (e) {
     let archivo = e.target.files[0];
     if (!archivo) return;
   
     let lector = new FileReader();
   
     lector.onload = function (event) {
       let data = new Uint8Array(event.target.result);
       let workbook = XLSX.read(data, { type: "array" });
       let hoja = workbook.Sheets[workbook.SheetNames[0]];
   
       datosExcel = XLSX.utils.sheet_to_json(hoja);
   
       console.log("Excel principal cargado:", datosExcel);
     };
   
     lector.readAsArrayBuffer(archivo);
   });
   
   // ── CARGA DEL EXCEL DE TIPOS (cruce por NOTA) ──
   document.getElementById("excelTipoFile").addEventListener("change", function (e) {
     let archivo = e.target.files[0];
     if (!archivo) return;
   
     let lector = new FileReader();
   
     lector.onload = function (event) {
       let data = new Uint8Array(event.target.result);
       let workbook = XLSX.read(data, { type: "array" });
   
       listaTipoExcel = extraerListaTipos(workbook);
   
       console.log("Excel de tipos cargado:", listaTipoExcel);
   
       if (listaTipoExcel.length === 0) {
         const resultado = document.getElementById("resultado");
         resultado.textContent =
           "⚠ No se encontraron columnas NOTA / TIPO (o ASN) en el Excel de tipos.";
         resultado.className = "resultado advertencia";
       }
     };
   
     lector.readAsArrayBuffer(archivo);
   });
   
   /**
    * Recorre todas las hojas del workbook de "tipos" buscando una que tenga
    * una columna NOTA y una columna TIPO (o ASN, por compatibilidad).
    * Devuelve una lista de objetos { notaNumeros, tipo }.
    */
   function extraerListaTipos(workbook) {
     const nombresColumnaNota = ["nota"];
     const nombresColumnaTipo = ["tipo", "asn"];
   
     for (const nombreHoja of workbook.SheetNames) {
       const hoja = workbook.Sheets[nombreHoja];
       const filas = XLSX.utils.sheet_to_json(hoja, { defval: "" });
   
       if (filas.length === 0) continue;
   
       // Detectar el nombre real de las columnas (case-insensitive) en esta hoja
       const encabezados = Object.keys(filas[0]);
       const colNota = encabezados.find((h) =>
         nombresColumnaNota.includes(h.toString().trim().toLowerCase())
       );
       const colTipo = encabezados.find((h) =>
         nombresColumnaTipo.includes(h.toString().trim().toLowerCase())
       );
   
       if (!colNota || !colTipo) continue; // esta hoja no sirve, seguir buscando
   
       const lista = filas
         .map((fila) => {
           const notaRaw = (fila[colNota] ?? "").toString();
           const notaNumeros = notaRaw.replace(/\D/g, ""); // solo dígitos
           const tipo = (fila[colTipo] ?? "").toString().trim();
           return { notaNumeros, tipo };
         })
         .filter((f) => f.notaNumeros && f.tipo);
   
       if (lista.length > 0) return lista;
     }
   
     return [];
   }
   
   /**
    * Busca el TIPO correspondiente a una nota del Excel principal,
    * comparando por los ÚLTIMOS números de la nota contra el Excel de tipos.
    */
   function obtenerTipoPorNota(notaValor) {
     if (!notaValor || listaTipoExcel.length === 0) return "";
   
     const notaNumeros = notaValor.toString().replace(/\D/g, "");
     if (!notaNumeros) return "";
   
     const match = listaTipoExcel.find((f) => notaNumeros.endsWith(f.notaNumeros));
   
     return match ? match.tipo : "";
   }
   
   // ── BÚSQUEDA DE LPN ──
   function buscarLPN() {
     const valorBusqueda = document
       .getElementById("lpnInput")
       .value.trim()
       .toUpperCase();
   
     const campo = document.getElementById("filtroCampo").value;
     const resultado = document.getElementById("resultado");
   
     if (!datosExcel || datosExcel.length === 0) {
       resultado.textContent = "No hay datos cargados.";
       resultado.className = "resultado advertencia";
       return;
     }
   
     const coincidencias = datosExcel.filter((item) => {
       let valorCampo = item[campo];
       if (!valorCampo) return false;
       return valorCampo.toString().toUpperCase().includes(valorBusqueda);
     });
   
     if (coincidencias.length > 0) {
       resultado.className = "resultado encontrado";
   
       // 👉 SOLO UNA coincidencia
       if (coincidencias.length === 1) {
         let item = coincidencias[0];
         item.TIPO = obtenerTipoPorNota(item.NOTA); // TIPO ahora sale del cruce, no del Excel principal
   
         guardarEnHistorial(item);
   
         resultado.innerHTML = `
           <strong>REFERENCIA:</strong> ${item.REFERENCIA}<br>
           <strong>LPN:</strong> ${item.LPN}<br>
           <strong>CANTIDAD:</strong> ${item.CANTIDAD}<br>
           <strong>NOTA:</strong> ${item.NOTA}<br>
           <strong>TIPO:</strong> ${item.TIPO || "No encontrado"}
         `;
         return;
       }
   
       // 👉 VARIAS coincidencias
       resultado.innerHTML = "<strong>Selecciona cuál guardar:</strong><br><br>";
   
       coincidencias.forEach((item) => {
         const tipoPreview = obtenerTipoPorNota(item.NOTA) || "—";
         resultado.innerHTML += `
           <div style="margin-bottom:10px; border-bottom:1px solid rgba(255,255,255,0.1); padding-bottom:8px;">
             <strong>LPN:</strong> ${item.LPN}<br>
             <strong>REF:</strong> ${item.REFERENCIA}<br>
             <strong>NOTA:</strong> ${item.NOTA}<br>
             <strong>TIPO:</strong> ${tipoPreview}<br>
             <button onclick="seleccionar('${encodeURIComponent(
               JSON.stringify(item)
             )}')">
             Guardar este
           </button>
           </div>
         `;
       });
     } else {
       resultado.textContent = "No se encontraron resultados.";
       resultado.className = "resultado no-encontrado";
     }
   }
   
   function seleccionar(item) {
     item = JSON.parse(decodeURIComponent(item));
     item.TIPO = obtenerTipoPorNota(item.NOTA); // TIPO ahora sale del cruce, no del Excel principal
   
     guardarEnHistorial(item);
   
     const resultado = document.getElementById("resultado");
   
     resultado.innerHTML = `
       <strong>Guardado correctamente:</strong><br><br>
       LPN: ${item.LPN}<br>
       REFERENCIA: ${item.REFERENCIA}<br>
       TIPO: ${item.TIPO || "No encontrado"}
     `;
   }
   
   function guardarEnHistorial(item) {
     let historial = JSON.parse(localStorage.getItem("historialLPN")) || [];
   
     let existe = historial.some(
       (h) => h.LPN === item.LPN && h.REFERENCIA === item.REFERENCIA
     );
   
     if (!existe) {
       historial.push({
         REFERENCIA: item.REFERENCIA || "",
         LPN: item.LPN || "",
         CANTIDAD: item.CANTIDAD || "",
         NOTA: item.NOTA || "",
         TIPO: item.TIPO || "",
       });
   
       localStorage.setItem("historialLPN", JSON.stringify(historial));
     }
   }
   
   /* ================================================
      Funciones de otras páginas (registro.html, etc.)
      — se dejan intactas, no forman parte de este cambio —
      ================================================ */
   
   function exportarExcel() {
     let historial = JSON.parse(localStorage.getItem("historialLPN")) || [];
   
     if (historial.length === 0) {
       alert("No hay registros para exportar");
       return;
     }
   
     let hoja = XLSX.utils.json_to_sheet(historial);
     let libro = XLSX.utils.book_new();
     XLSX.utils.book_append_sheet(libro, hoja, "Historial LPN");
     XLSX.writeFile(libro, "Historial_LPN.xlsx");
   }
   
   function limpiarTabla() {
     if (!confirm("¿Seguro que quieres borrar todo el historial?")) {
       return;
     }
   
     localStorage.removeItem("historialLPN");
   
     let tabla = document.querySelector("#tablaHistorial tbody");
     if (tabla) tabla.innerHTML = "";
   
     const contador = document.getElementById("contadorRegistros");
     if (contador) contador.textContent = "Cantidad de registros: 0";
   }
   
   const btnTop = document.getElementById("btnTop");
   
   if (btnTop) {
     window.addEventListener("scroll", () => {
       if (window.scrollY > 200) {
         btnTop.style.display = "block";
       } else {
         btnTop.style.display = "none";
       }
     });
   }
   
   function irArriba() {
     window.scrollTo({
       top: 0,
       behavior: "smooth",
     });
   }