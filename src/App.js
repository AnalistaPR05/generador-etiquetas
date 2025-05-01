import './App.css';
import { useState } from 'react';
import * as XLSX from 'xlsx';
function App() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleFileChange = (e) => {
    const uploadedFile = e.target.files[0];
    setFile(uploadedFile);
    console.log('Archivo seleccionado:', uploadedFile?.name);
  };

  const esTalla = (valor) =>
    typeof valor === 'string' && /^\d{2}\s*-\s*\d{2}$/.test(valor);

  const handleProcessClick = () => {
    if (!file) {
      alert('Primero selecciona un archivo Excel.');
      return;
    }
  
    setLoading(true);
  
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
  
      const sheetCurva = workbook.Sheets['CURVA (2)'];
      const sheetEmpaque = workbook.Sheets['EMPAQUE'];
  
      if (!sheetCurva || !sheetEmpaque) {
        console.error('No se encontraron las hojas requeridas');
        setLoading(false);
        return;
      }
  
      const resultados = [];
  
      // EMPAQUE - columnas AO (40) hasta BX (75)
      const mapaDescripciones = {};
      const inicioFila = 4;
      const colNombre = 38;
      const colInicioDescripcion = 40;
      const colFinDescripcion = 75;
  
      for (let row = inicioFila; row <= 1000; row++) {
        const cellNombre = sheetEmpaque[XLSX.utils.encode_cell({ r: row, c: colNombre })];
        const nombreProducto = cellNombre?.v?.toString().trim();
        if (!nombreProducto || nombreProducto.toLowerCase() === '(en blanco)') break;
  
        const descripciones = [];
        for (let col = colInicioDescripcion; col <= colFinDescripcion; col++) {
          const cellDesc = sheetEmpaque[XLSX.utils.encode_cell({ r: row, c: col })];
          const valor = cellDesc?.v?.toString().trim();
          if (valor) {
            descripciones.push(valor);
          }
        }
  
        mapaDescripciones[nombreProducto] = descripciones;
      }
  
      // CURVA - columnas B (1) hasta AX (49), de 2 en 2
      for (let row = 4; row <= 1000; row++) {
        const cellProducto = sheetCurva[XLSX.utils.encode_cell({ r: row, c: 0 })];
        const nombreProducto = cellProducto?.v?.toString().trim();
        if (!nombreProducto || nombreProducto.toLowerCase() === '(en blanco)') break;
  
        const datos = [];
        for (let col = 1; col <= 49; col += 2) {
          const cell1 = sheetCurva[XLSX.utils.encode_cell({ r: row, c: col })];
          const cell2 = sheetCurva[XLSX.utils.encode_cell({ r: row, c: col + 1 })];
  
          const val1 = cell1?.v?.toString().trim();
          const val2 = cell2?.v?.toString().trim();
  
          if (esTalla(val1) && !isNaN(parseFloat(val2))) {
            datos.push({ tallas: val1, precio: parseFloat(val2) });
          }
        }
  
        resultados.push({
          producto: nombreProducto,
          datos,
          descripciones: mapaDescripciones[nombreProducto] || []
        });
      }
  
      // Generar Excel con los resultados
      const datosAOA = [];
      const descripcionesAOA = [];
  
      resultados.forEach((etiqueta, idx) => {
        datosAOA[0] = datosAOA[0] || [];
        datosAOA[0][idx * 2] = etiqueta.producto;
  
        etiqueta.datos.forEach((dato, i) => {
          datosAOA[i + 1] = datosAOA[i + 1] || [];
          datosAOA[i + 1][idx * 2] = dato.tallas;
          datosAOA[i + 1][idx * 2 + 1] = dato.precio;
        });
  
        descripcionesAOA[0] = descripcionesAOA[0] || [];
        descripcionesAOA[0][idx] = etiqueta.producto;
  
        etiqueta.descripciones.forEach((desc, i) => {
          descripcionesAOA[i + 1] = descripcionesAOA[i + 1] || [];
          descripcionesAOA[i + 1][idx] = desc;
        });
      });
  
      const wsDatos = XLSX.utils.aoa_to_sheet(datosAOA);
      const wsDescripciones = XLSX.utils.aoa_to_sheet(descripcionesAOA);
      const newWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWb, wsDatos, 'Datos Tallas y Precios');
      XLSX.utils.book_append_sheet(newWb, wsDescripciones, 'Descripciones');
      XLSX.writeFile(newWb, 'datos_etiquetas.xlsx');
  
      setLoading(false);
    };
  
    reader.readAsArrayBuffer(file);
  };
  

  return (
    <div className="container">
      <header className="app-header">
        <h1>🧾 Generador de Datos para Etiquetas</h1>
        <p>Sube tu archivo Excel y descarga los datos formateados</p>
      </header>
  
      <main className="app-main">
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileChange}
          className="form-control mb-3"
        />
  
        <button
          onClick={handleProcessClick}
          className="btn btn-primary"
          disabled={loading}
        >
          {loading ? 'Formateando datos...' : 'Procesar y generar archivo base'}
        </button>
  
        {loading && (
          <div className="spinner">
            ⏳ Por favor espera, procesando archivo...
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
