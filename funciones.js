let datosExcel = [];

document.getElementById("excelFile").addEventListener("change", function (e) {
  let archivo = e.target.files[0];

  let lector = new FileReader();

  lector.onload = function (event) {
    let data = new Uint8Array(event.target.result);

    let workbook = XLSX.read(data, { type: "array" });

    let hoja = workbook.Sheets[workbook.SheetNames[0]];

    datosExcel = XLSX.utils.sheet_to_json(hoja);

    console.log(datosExcel);
  };

  lector.readAsArrayBuffer(archivo);
});

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

  // 🔥 NUEVA LÓGICA
  if (coincidencias.length > 0) {
    resultado.className = "resultado encontrado";

    // 👉 SOLO UNA coincidencia
    if (coincidencias.length === 1) {
      guardarEnHistorial(coincidencias[0]);

      resultado.innerHTML = `
        <strong>REFERENCIA:</strong> ${coincidencias[0].REFERENCIA}<br>
        <strong>LPN:</strong> ${coincidencias[0].LPN}<br>
        <strong>CANTIDAD:</strong> ${coincidencias[0].CANTIDAD}<br>
        <strong>NOTA:</strong> ${coincidencias[0].NOTA}<br>
        <strong>TIPO:</strong> ${coincidencias[0].TIPO}
      `;
      return;
    }

    // 👉 VARIAS coincidencias
    resultado.innerHTML = "<strong>Selecciona cuál guardar:</strong><br><br>";

    coincidencias.forEach((item) => {
      resultado.innerHTML += `
        <div style="margin-bottom:10px; border-bottom:1px solid rgba(255,255,255,0.1); padding-bottom:8px;">
          <strong>LPN:</strong> ${item.LPN}<br>
          <strong>REF:</strong> ${item.REFERENCIA}<br>
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
function exportarExcel() {
  let historial = JSON.parse(localStorage.getItem("historialLPN")) || [];

  if (historial.length === 0) {
    alert("No hay registros para exportar");
    return;
  }

  // Convertir a hoja de Excel
  let hoja = XLSX.utils.json_to_sheet(historial);

  // Crear libro
  let libro = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(libro, hoja, "Historial LPN");

  // Descargar archivo
  XLSX.writeFile(libro, "Historial_LPN.xlsx");
}
function limpiarTabla() {
  if (!confirm("¿Seguro que quieres borrar todo el historial?")) {
    return;
  }

  localStorage.removeItem("historialLPN");

  let tabla = document.querySelector("#tablaHistorial tbody");
  tabla.innerHTML = "";

  // actualizar contador
  document.getElementById("contadorRegistros").textContent =
    "Cantidad de registros: 0";
}
const btnTop = document.getElementById("btnTop");

// Mostrar botón cuando bajas
window.addEventListener("scroll", () => {
  if (window.scrollY > 200) {
    btnTop.style.display = "block";
  } else {
    btnTop.style.display = "none";
  }
});

// Función subir
function irArriba() {
  window.scrollTo({
    top: 0,
    behavior: "smooth",
  });
}
function seleccionar(item) {
  item = JSON.parse(decodeURIComponent(item));

  guardarEnHistorial(item);

  const resultado = document.getElementById("resultado");

  resultado.innerHTML = `
    <strong>Guardado correctamente:</strong><br><br>
    LPN: ${item.LPN}<br>
    REFERENCIA: ${item.REFERENCIA}
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
