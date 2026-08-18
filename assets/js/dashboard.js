// Variables globales para instancias de gráficos
let chartGenero, chartParentesco, chartEnfermedades, chartEducacion, chartLaboral, chartTenencia, chartOcupante;
let chartComunidad, chartComunidadTorre; // Nuevos gráficos para pestaña 5

// Paleta de colores global para los gráficos
const chartColors = [
  '#007bff', '#28a745', '#ffc107', '#dc3545', '#17a2b8', 
  '#6610f2', '#fd7e14', '#20c997', '#6f42c1', '#d63384', 
  '#0dcaf0', '#198754', '#ffcd39', '#0d6efd', '#e83e8c'
];

// Registrar el plugin de Datalabels globalmente
Chart.register(ChartDataLabels);

// Configuración global de Chart.js para Datalabels
Chart.defaults.set('plugins.datalabels', {
  color: '#fff',
  font: { weight: 'bold', size: 11 },
  textShadowColor: 'rgba(0, 0, 0, 0.8)',
  textShadowBlur: 4,
  formatter: (value, ctx) => {
    if (value === 0) return null;
    let sum = 0;
    let dataArr = ctx.chart.data.datasets[ctx.datasetIndex].data;
    dataArr.forEach(data => { sum += Math.abs(data); });
    let percentage = sum > 0 ? ((Math.abs(value) * 100) / sum).toFixed(1) + "%" : "0%";
    return `${Math.abs(value)}\n(${percentage})`;
  },
  textAlign: 'center'
});

document.addEventListener('DOMContentLoaded', () => {
  const loadingStatus = document.getElementById('loadingStatus');
  const exportPdfBtn = document.getElementById('exportPdfBtn');
  const fileInput = document.getElementById('excelUpload');


  // --- Exportar SOLO la pestaña activa ---
  const exportAllPdfBtn = document.getElementById('exportAllPdfBtn');

  // Función auxiliar para forzar fondo blanco de Chart.js antes de captura
  function setChartBackgrounds(tabNode, color) {
    const charts = tabNode.querySelectorAll('canvas');
    charts.forEach(canvas => {
      const ctx = canvas.getContext('2d');
      ctx.save();
      ctx.globalCompositeOperation = 'destination-over';
      ctx.fillStyle = color;
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.restore();
    });
  }

  // Función principal para exportar a PDF (uno o todos)
  async function exportToPDF(exportAll = false) {
    const btn = exportAll ? exportAllPdfBtn : exportPdfBtn;
    btn.disabled = true;
    const originalText = btn.innerHTML;
    btn.textContent = 'Exportando...';

    try {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
      
      // Obtener qué pestañas procesar
      const tabsToProcess = exportAll 
        ? Array.from(document.querySelectorAll('.tab-pane'))
        : [document.querySelector('.tab-pane.active')].filter(Boolean);

      if (tabsToProcess.length === 0) {
        alert('No hay pestañas para exportar.');
        return;
      }

      for (let i = 0; i < tabsToProcess.length; i++) {
        const tab = tabsToProcess[i];
        const tabTitle = document.querySelector(`.tab-btn[onclick*="${tab.id}"]`).textContent;
        const charts = Array.from(tab.querySelectorAll('.chart-box'));
        
        // Guardar estilos originales para restaurar después
        const originalTabStyle = tab.getAttribute('style') || '';
        const wasHidden = tab.style.display === 'none' || !tab.classList.contains('active');
        
        // Configurar la pestaña para que se pueda capturar sin afectar el DOM visible
        if (wasHidden) {
          tab.style.display = 'block';
          tab.style.position = 'absolute';
          tab.style.top = '-9999px';
          tab.style.left = '-9999px';
          tab.style.width = '100%'; 
        }

        for (let j = 0; j < charts.length; j++) {
          const chartBox = charts[j];
          const chartTitle = chartBox.querySelector('h4') ? chartBox.querySelector('h4').textContent : 'Gráfico';
          
          // Asegurar fondo blanco puro para Chart.js y evitar transparencias
          setChartBackgrounds(chartBox, '#ffffff');
          
          // Guardar estilos originales de la caja
          const originalBoxStyle = chartBox.getAttribute('style') || '';
          
          // Aplicar estilos temporales inline a la caja para asegurar fondo blanco 
          chartBox.style.backgroundColor = '#ffffff';
          chartBox.style.padding = '20px';
          chartBox.style.borderRadius = '0'; // Quitar bordes redondeados para el PDF
          chartBox.style.border = 'none';

          // Capturar el canvas
          const canvas = await html2canvas(chartBox, {
            scale: 2,
            useCORS: true,
            backgroundColor: '#ffffff',
            logging: false
          });

          // Restaurar estilos de la caja
          chartBox.setAttribute('style', originalBoxStyle);

          // Agregar página nueva si no es el primer gráfico de la primera pestaña
          if (i !== 0 || j !== 0) {
            doc.addPage();
          }

          // Dibujar Cabecera Elegante
          doc.setFillColor(59, 130, 246); // Color Azul Primario (--primary-color)
          doc.rect(0, 0, 297, 25, 'F'); // A4 landscape width is ~297mm
          
          doc.setTextColor(255, 255, 255);
          doc.setFontSize(16);
          doc.setFont('helvetica', 'bold');
          doc.text(`Módulo: ${tabTitle.toUpperCase()}`, 15, 12);
          
          doc.setFontSize(12);
          doc.setFont('helvetica', 'normal');
          doc.text(`Gráfico: ${chartTitle}`, 15, 20);

          // Agregar logo o texto a la derecha
          doc.setFontSize(14);
          doc.setFont('helvetica', 'bolditalic');
          doc.text('Octava Estrella de Guayana', 282, 16, { align: 'right' });

          // Calcular dimensiones para mantener proporción del gráfico ocupando el resto de la hoja A4
          const imgData = canvas.toDataURL('image/jpeg', 0.98);
          
          const pdfWidth = 297; 
          const pdfHeight = 210;
          const headerHeight = 25;
          const margin = 10;
          
          const maxImgWidth = pdfWidth - (margin * 2);
          const maxImgHeight = pdfHeight - headerHeight - (margin * 2);
          
          let imgWidth = maxImgWidth;
          let imgHeight = (canvas.height * maxImgWidth) / canvas.width;
          
          if (imgHeight > maxImgHeight) {
            imgHeight = maxImgHeight;
            imgWidth = (canvas.width * maxImgHeight) / canvas.height;
          }
          
          // Centrar la imagen en el espacio disponible
          const x = (pdfWidth - imgWidth) / 2;
          const y = headerHeight + margin + ((maxImgHeight - imgHeight) / 2);

          doc.addImage(imgData, 'JPEG', x, y, imgWidth, imgHeight);
        }

        // Restaurar pestaña
        tab.setAttribute('style', originalTabStyle);
      }

      // Descargar PDF final
      const fileName = exportAll ? 'Dashboard_Completo_RioAro.pdf' : 'Dashboard_Pestaña_RioAro.pdf';
      doc.save(fileName);

    } catch (err) {
      console.error('Error generando PDF:', err);
      alert('Ocurrió un error al generar el PDF. Revisa la consola.');
    } finally {
      btn.disabled = false;
      btn.innerHTML = originalText;
    }
  }

  // Asignar listeners
  exportPdfBtn.addEventListener('click', () => exportToPDF(false));
  exportAllPdfBtn.addEventListener('click', () => exportToPDF(true));


  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    loadingStatus.innerHTML = `<i class="bi bi-check-circle-fill" style="color: green;"></i> Procesando archivo local...`;
    
    const reader = new FileReader();
    reader.onload = function(e) {
      procesarArrayBuffer(e.target.result);
    };
    reader.readAsArrayBuffer(file);
  });

  // --- Carga Automática usando fetch() ---
  const urlExcel = '/api/excel/download-full';
  
  fetch(urlExcel)
    .then(response => {
      if (!response.ok) throw new Error(`Error en la red: ${response.status}`);
      return response.arrayBuffer();
    })
    .then(data => {
      loadingStatus.innerHTML = `<i class="bi bi-check-circle-fill" style="color: green;"></i> Excel cargado con éxito`;
      procesarArrayBuffer(data);
    })
    .catch(error => {
      console.error("Error al cargar el Excel automáticamente:", error);
      loadingStatus.innerHTML = `<i class="bi bi-exclamation-triangle-fill" style="color: red;"></i> Error al cargar datos. Use carga manual.`;
    });
});

function procesarArrayBuffer(data) {
  try {
    const buffer = new Uint8Array(data);
    const workbook = XLSX.read(buffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    
    if (jsonData.length > 0) {
        procesarDatos(jsonData);
        document.getElementById('exportPdfBtn').disabled = false;
        document.getElementById('exportAllPdfBtn').disabled = false;
    } else {
        alert("El archivo no contiene datos o la hoja está vacía.");
    }
  } catch (error) {
    console.error("Error procesando el Excel:", error);
    alert("Ocurrió un error leyendo el Excel. Verifique que el archivo sea válido.");
  }
}

// Validación defensiva unificada: Oculta SIEMPRE "NA" y "N/A"
function esCondicionVulnerable(valor) {
  if (valor === undefined || valor === null) return false;
  const texto = String(valor).trim().toUpperCase();
  if (texto === "" || texto === "NINGUNA" || texto === "NINGUNO" || texto === "NO" || texto === "N/A" || texto === "NA") return false;
  return true;
}

function procesarDatos(data) {
  const cleanData = data.map(row => {
    let newRow = {};
    for (let key in row) {
      newRow[key.trim()] = row[key];
    }
    return newRow;
  });

  let totalPoblacion = cleanData.length;
  let totalFamilias = 0;
  let totalVulnerables = 0;
  let totalMenores = 0;
  let totalEmbarazadas = 0;
  let totalLactantes = 0;

  let totalMasculino = 0;
  let totalFemenino = 0;
  const parentescos = {};
  const enfermedadesYDiscapacidades = {};
  const medicamentos = {};
  const educacion = {};
  const ocupacion = {};
  const tenencia = {};
  const condicionOcupante = {};

  // Variables para Pestaña 5: Territorial y Densidad
  const comunidades = {};
  const comunidadesTorres = {};
  const comunidadViviendas = {};
  const comunidadHabitantes = {};

  cleanData.forEach(row => {
    // Familias
    const parentescoStr = String(row['Parentesco'] || row['PARENTESCO'] || '').toUpperCase().trim();
    if (parentescoStr.includes('JEFE DE FAMILIA') || parentescoStr.includes('JEFE')) totalFamilias++;
    if (parentescoStr) parentescos[parentescoStr] = (parentescos[parentescoStr] || 0) + 1;

    // Menores
    const edadStr = String(row['Edad'] || row['EDAD'] || '0').replace(/\D/g, '');
    const edad = parseInt(edadStr) || 0;
    if (edad < 18) totalMenores++;

    // Género Clásico
    let genero = String(row['Genero Biológico'] || row['Genero biológico'] || row['GENERO BIOLOGICO'] || row['Genero'] || '').toUpperCase().trim();
    if (genero === 'M') totalMasculino++;
    if (genero === 'F') totalFemenino++;

    // Distribución Territorial
    const com = String(row['Comunidad'] || '').trim().toUpperCase();
    const tor = String(row['Torre'] || '').trim().toUpperCase();
    const cas = String(row['Numero de Casa / Apto'] || row['Numero de casa / Apto'] || row['Número de Casa / Apto'] || row['Numero de Casa / Apto.'] || '').trim().toUpperCase();

    if (com && com !== 'NA' && com !== 'N/A') {
        comunidades[com] = (comunidades[com] || 0) + 1;
        comunidadHabitantes[com] = (comunidadHabitantes[com] || 0) + 1;
        
        if (!comunidadesTorres[com]) comunidadesTorres[com] = {};
        if (tor && tor !== 'NA' && tor !== 'N/A') {
            comunidadesTorres[com][tor] = (comunidadesTorres[com][tor] || 0) + 1;
        }

        if (tor && cas && tor !== 'NA' && cas !== 'NA') {
            const idVivienda = `${com}|${tor}|${cas}`;
            if (!comunidadViviendas[com]) comunidadViviendas[com] = new Set();
            comunidadViviendas[com].add(idVivienda);
        }
    }

    // Búsqueda dinámica de la columna lactantes
    const lactanteKey = Object.keys(row).find(k => {
      const lowerKey = k.toLowerCase();
      return lowerKey.includes('lactante') || lowerKey.includes('11 meses');
    });

    const enf = String(row['Enfermedades crónicas'] || row['ENFERMEDADES CRONICAS'] || '').trim();
    const disc = String(row['Tipo de discapacidad'] || '').trim();
    const necEsp = String(row['Necesidades especiales'] || row['Necesidades Especiales'] || '').trim();
    
    const embarazo = String(row['Embarazadas'] || row['Embarazada'] || '').trim().toUpperCase();
    const lactante = lactanteKey ? String(row[lactanteKey] || '').trim().toUpperCase() : '';

    // 1. KPI Población Vulnerable (Deduplicación y exclusión global de NA)
    if (
      esCondicionVulnerable(enf) || 
      esCondicionVulnerable(disc) || 
      esCondicionVulnerable(necEsp) || 
      embarazo === 'SI' || 
      lactante === 'SI'
    ) {
      totalVulnerables++;
    }

    // Totales Materno-Infantil
    if (embarazo === 'SI') totalEmbarazadas++;
    if (lactante === 'SI') totalLactantes++;

    // Gráficos de Salud y Atención (Se excluye NA mediante la validación unificada)
    if (esCondicionVulnerable(enf)) {
      const keyEnf = enf.toUpperCase();
      enfermedadesYDiscapacidades[keyEnf] = (enfermedadesYDiscapacidades[keyEnf] || 0) + 1;
    }
    
    if (esCondicionVulnerable(disc)) {
      const keyDisc = disc.toUpperCase();
      enfermedadesYDiscapacidades[keyDisc] = (enfermedadesYDiscapacidades[keyDisc] || 0) + 1;
    }

    // Medicamentos (Se excluye NA)
    const med = String(row['Medicamento consumido'] || row['Medicamentos'] || '');
    if (esCondicionVulnerable(med)) {
      const keyMed = med.trim().toUpperCase();
      medicamentos[keyMed] = (medicamentos[keyMed] || 0) + 1;
    }

    // Educación (Se excluye NA)
    const edu = String(row['Estudio cursado'] || row['ESTUDIO CURSADO'] || '').trim().toUpperCase();
    if (edu && edu !== 'NINGUNO' && edu !== 'NA' && edu !== 'N/A') {
      educacion[edu] = (educacion[edu] || 0) + 1;
    }

    // Situación Laboral (Se excluye NA)
    const ocu = String(row['Ocupación'] || row['Ocupacion'] || '').trim().toUpperCase();
    if (ocu && ocu !== 'NINGUNA' && ocu !== 'NO APLICA' && ocu !== 'NA' && ocu !== 'N/A') {
      ocupacion[ocu] = (ocupacion[ocu] || 0) + 1;
    }

    // Vivienda y Ocupante (Se excluye NA)
    const ten = String(row['Condición de Vivienda'] || row['Condicion de Vivienda'] || row['Condición de vivienda'] || '').trim().toUpperCase();
    if (ten && ten !== 'NA' && ten !== 'N/A') tenencia[ten] = (tenencia[ten] || 0) + 1;

    const estOcu = String(row['Estado de ocupante'] || '').trim().toUpperCase();
    if (estOcu && estOcu !== 'NA' && estOcu !== 'N/A') condicionOcupante[estOcu] = (condicionOcupante[estOcu] || 0) + 1;

  });

  // Actualizar UI - KPIs
  document.getElementById('kpiTotal').textContent = totalPoblacion;
  document.getElementById('kpiFamilias').textContent = totalFamilias;
  document.getElementById('kpiVulnerable').textContent = totalVulnerables;
  document.getElementById('kpiMenores').textContent = totalMenores;
  document.getElementById('kpiEmbarazadas').textContent = totalEmbarazadas;
  document.getElementById('kpiLactantes').textContent = totalLactantes;

  renderTable(medicamentos, 'tableMedicamentos', 10);

  // Renderizar Gráficos
  renderBarChartGenero(totalMasculino, totalFemenino);
  renderPieChart('chartParentesco', parentescos, 6);
  renderPieChart('chartLaboral', ocupacion, 15);
  renderBarChart('chartEnfermedades', enfermedadesYDiscapacidades, 'Casos', 'horizontal', 15);
  renderBarChart('chartEducacion', educacion, 'Personas', 'vertical', 10);
  renderHistogram('chartTenencia', tenencia);
  renderAreaChart('chartOcupante', condicionOcupante);
  
  // Renderizar nuevos gráficos de Densidad y Territorio
  renderDoughnutChart('chartComunidad', comunidades);
  renderStackedBarChart('chartComunidadTorre', comunidadesTorres);
  renderDensidadKPIs(comunidadHabitantes, comunidadViviendas);
}

function getTopN(obj, n = 10) {
  return Object.entries(obj)
    .sort((a, b) => b[1] - a[1])
    .slice(0, n)
    .reduce((r, [k, v]) => ({ ...r, [k]: v }), {});
}

function renderTable(dataObj, tableId, n = 5) {
  const topData = getTopN(dataObj, n);
  const tbody = document.querySelector(`#${tableId} tbody`);
  if(!tbody) return;
  tbody.innerHTML = '';
  
  const keys = Object.keys(topData);
  if (keys.length === 0) {
    tbody.innerHTML = '<tr><td colspan="2">No hay datos</td></tr>';
    return;
  }

  keys.forEach(k => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${k}</td><td>${topData[k]}</td>`;
    tbody.appendChild(tr);
  });
}

function destroyChartIfExists(chartInstance) {
  if (chartInstance) {
    chartInstance.destroy();
  }
}

function renderBarChartGenero(mCount, fCount) {
  const ctx = document.getElementById('chartPiramide').getContext('2d');
  destroyChartIfExists(chartGenero);

  chartGenero = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['Masculino (M)', 'Femenino (F)'],
      datasets: [
        {
          label: 'Total por Género',
          data: [mCount, fCount],
          backgroundColor: ['#3b82f6', '#e83e8c'],
          borderRadius: 4
        }
      ]
    },
    options: {
      indexAxis: 'x', 
      responsive: true,
      animation: { duration: 0 },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => `Total: ${context.raw}`
          }
        },
        datalabels: {
          display: true
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: '10%'
        }
      }
    }
  });
}

function renderPieChart(canvasId, dataObj, limitN = 10) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  
  if (canvasId === 'chartParentesco') destroyChartIfExists(chartParentesco);
  else if (canvasId === 'chartLaboral') destroyChartIfExists(chartLaboral);

  const topData = getTopN(dataObj, limitN);
  const labels = Object.keys(topData);
  const data = Object.values(topData);

  const newChart = new Chart(ctx, {
    type: 'pie', 
    data: {
      labels: labels,
      datasets: [{
        data: data,
        backgroundColor: chartColors
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false, 
      animation: { duration: 0 },
      plugins: {
        legend: { position: 'right' },
        datalabels: {
          display: true
        }
      }
    }
  });

  if (canvasId === 'chartParentesco') chartParentesco = newChart;
  else if (canvasId === 'chartLaboral') chartLaboral = newChart;
}

// === Nuevos Gráficos Pestaña 5 ===

function renderDoughnutChart(canvasId, dataObj) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  destroyChartIfExists(chartComunidad);

  // Ordenar de mayor a menor
  const sortedData = Object.entries(dataObj).sort((a,b) => b[1]-a[1]);
  const labels = sortedData.map(d => d[0]);
  const data = sortedData.map(d => d[1]);

  chartComunidad = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: data,
        backgroundColor: chartColors,
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 0 },
      plugins: {
        legend: { position: 'right' },
        datalabels: { display: true }
      }
    }
  });
}

function renderStackedBarChart(canvasId, dataObj) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  destroyChartIfExists(chartComunidadTorre);

  const comunidades = Object.keys(dataObj).sort(); // Eje X
  // Extraer todas las torres únicas de todas las comunidades para las series
  const todasLasTorres = new Set();
  comunidades.forEach(c => {
      Object.keys(dataObj[c]).forEach(t => todasLasTorres.add(t));
  });
  
  const torresArray = Array.from(todasLasTorres).sort();
  const datasets = torresArray.map((torre, index) => {
      return {
          label: torre,
          backgroundColor: chartColors[index % chartColors.length],
          data: comunidades.map(c => dataObj[c][torre] || 0),
          borderRadius: 4
      };
  });

  chartComunidadTorre = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: comunidades,
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 0 },
      scales: {
        x: { stacked: true },
        y: { stacked: true, grace: '10%' }
      },
      plugins: {
        legend: { display: false },
        datalabels: {
           display: true,
           formatter: (v) => v > 5 ? v : null // Oculta métricas pequeñas para evitar saturación visual en bloques finos
        }
      }
    }
  });
}

function renderDensidadKPIs(habitantesObj, viviendasObj) {
    const container = document.getElementById('kpiDensidadContainer');
    if(!container) return;
    container.innerHTML = '';

    const comunidades = Object.keys(habitantesObj).sort();
    
    comunidades.forEach(com => {
        const hab = habitantesObj[com] || 0;
        const viv = viviendasObj[com] ? viviendasObj[com].size : 0;
        const promedio = viv > 0 ? (hab / viv).toFixed(1) : 0;

        const card = document.createElement('div');
        card.className = 'mini-kpi';
        card.style.flex = '1';
        card.style.minWidth = '220px';
        card.style.textAlign = 'left';
        
        card.innerHTML = `
            <h5 style="color: #3b82f6; font-weight: bold; border-bottom: 1px solid #e5e7eb; padding-bottom: 5px; margin-bottom: 10px; font-size: 1rem;">${com}</h5>
            <div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
                <span style="font-size: 0.9em; color: #6b7280;">Habitantes:</span>
                <strong>${hab}</strong>
            </div>
            <div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
                <span style="font-size: 0.9em; color: #6b7280;">Viviendas únicas:</span>
                <strong>${viv}</strong>
            </div>
            <div style="display: flex; justify-content: space-between; background-color: #f3f4f6; padding: 5px; border-radius: 4px; margin-top: 10px;">
                <span style="font-size: 0.85em; font-weight: bold;">Promedio Hab/Viv:</span>
                <strong style="color: #10b981;">${promedio}</strong>
            </div>
        `;
        container.appendChild(card);
    });
}

// === Fin Nuevos Gráficos ===

function renderBarChart(canvasId, dataObj, labelPrefix, direction, limitN = 8) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  
  if (canvasId === 'chartEnfermedades') destroyChartIfExists(chartEnfermedades);
  if (canvasId === 'chartEducacion') destroyChartIfExists(chartEducacion);

  const topData = getTopN(dataObj, limitN);
  const labels = Object.keys(topData);
  const data = Object.values(topData);

  const bgColors = labels.map((_, i) => chartColors[i % chartColors.length]);

  const newChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: labelPrefix,
        data: data,
        backgroundColor: bgColors,
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: direction === 'horizontal' ? 'y' : 'x',
      responsive: true,
      animation: { duration: 0 },
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true
        }
      }
    }
  });

  if (canvasId === 'chartEnfermedades') chartEnfermedades = newChart;
  if (canvasId === 'chartEducacion') chartEducacion = newChart;
}

function renderHistogram(canvasId, dataObj) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  destroyChartIfExists(chartTenencia);

  const topData = getTopN(dataObj, 10);
  const labels = Object.keys(topData);
  const data = Object.values(topData);
  const bgColors = labels.map((_, i) => chartColors[i % chartColors.length]);

  chartTenencia = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Tenencia',
        data: data,
        backgroundColor: bgColors,
        barPercentage: 1,      
        categoryPercentage: 1  
      }]
    },
    options: {
      responsive: true,
      animation: { duration: 0 },
      layout: {
        padding: { left: 20, right: 20 } 
      },
      scales: {
        x: {
          grid: { offset: false },
          ticks: { autoSkip: false } 
        },
        y: {
          beginAtZero: true,
          grace: '10%' 
        }
      },
      plugins: {
        legend: { display: false },
        datalabels: { display: true }
      }
    }
  });
}

function renderAreaChart(canvasId, dataObj) {
  const ctx = document.getElementById(canvasId).getContext('2d');
  destroyChartIfExists(chartOcupante);

  const topData = getTopN(dataObj, 10);
  const labels = Object.keys(topData);
  const data = Object.values(topData);

  chartOcupante = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: 'Permanencia',
        data: data,
        backgroundColor: 'rgba(59, 130, 246, 0.4)',
        borderColor: '#3b82f6',
        borderWidth: 2,
        fill: true,
        tension: 0.3 
      }]
    },
    options: {
      responsive: true,
      animation: { duration: 0 },
      scales: {
        y: {
          beginAtZero: true,
          grace: '10%'
        }
      },
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true,
          align: 'top',
          color: '#333',
          textShadowBlur: 0
        }
      }
    }
  });
}
