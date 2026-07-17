// =============================================
// BAI05 — exportar.js
// Exportación a Excel con formato tipo formulario (ExcelJS)
// =============================================

function val(id) {
  const el = document.getElementById(id);
  return el ? el.value : '';
}

function checkboxesSeleccionados(name) {
  return Array.from(document.querySelectorAll('input[name="' + name + '"]:checked'))
    .map(function (c) {
      var lbl = c.closest('label');
      return lbl ? lbl.textContent.trim() : c.value;
    }).join('\n');
}

function filasDinamicas(tbodyId, nCols) {
  var tbody = document.getElementById(tbodyId);
  if (!tbody) return [];
  return Array.from(tbody.rows).map(function (row) {
    var celdas = [];
    for (var i = 0; i < nCols; i++) {
      var celda = row.cells[i];
      if (!celda) { celdas.push(''); continue; }
      var inp = celda.querySelector('input, select, textarea');
      celdas.push(inp ? inp.value : celda.textContent.trim());
    }
    return celdas;
  });
}

// ---- Paleta institucional (styles.css) ----
var NEGRO = 'FF1A1A1A';
var BLANCO = 'FFFFFFFF';
var GRIS_CLARO = 'FFF2F2F2';
var GRIS_LINEA = 'FFD8D8D8';
var GRIS_OSCURO = 'FF4A4A4A';
var ROJO = 'FFB30000';
var VERDE = 'FF1A6E3C';
var FUENTE = 'Segoe UI';

var borde = {
  top:    { style: 'thin', color: { argb: GRIS_LINEA } },
  left:   { style: 'thin', color: { argb: GRIS_LINEA } },
  bottom: { style: 'thin', color: { argb: GRIS_LINEA } },
  right:  { style: 'thin', color: { argb: GRIS_LINEA } }
};

// Constructor de hoja con formato
function construirHoja(ws, titulo, subtitulo, cobit, bloques) {
  ws.columns = [
    { width: 34 }, { width: 26 }, { width: 24 }, { width: 22 }, { width: 20 }, { width: 16 }
  ];
  var r = 1;

  // --- Título ---
  ws.mergeCells(r, 1, r, 6);
  var celT = ws.getCell(r, 1);
  celT.value = titulo;
  celT.font = { name: FUENTE, size: 14, bold: true, color: { argb: BLANCO } };
  celT.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NEGRO } };
  celT.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(r).height = 26;
  r++;

  if (subtitulo) {
    ws.mergeCells(r, 1, r, 6);
    var celS = ws.getCell(r, 1);
    celS.value = subtitulo;
    celS.font = { name: FUENTE, size: 10, italic: true, color: { argb: GRIS_OSCURO } };
    celS.alignment = { indent: 1 };
    r++;
  }
  if (cobit) {
    ws.mergeCells(r, 1, r, 6);
    var celC = ws.getCell(r, 1);
    celC.value = cobit;
    celC.font = { name: FUENTE, size: 9, color: { argb: GRIS_OSCURO } };
    celC.alignment = { indent: 1 };
    r++;
  }
  r++; // fila en blanco

  bloques.forEach(function (b) {
    if (b.tipo === 'seccion') {
      ws.mergeCells(r, 1, r, 6);
      var cs = ws.getCell(r, 1);
      cs.value = b.texto;
      cs.font = { name: FUENTE, size: 11, bold: true, color: { argb: BLANCO } };
      cs.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NEGRO } };
      cs.alignment = { vertical: 'middle', indent: 1 };
      ws.getRow(r).height = 20;
      r++;
      return;
    }

    if (b.tipo === 'campo') {
      var cl = ws.getCell(r, 1);
      cl.value = b.label;
      cl.font = { name: FUENTE, size: 10, bold: true, color: { argb: GRIS_OSCURO } };
      cl.alignment = { vertical: 'top', wrapText: true };
      cl.border = borde;
      ws.mergeCells(r, 2, r, 6);
      var cv = ws.getCell(r, 2);
      cv.value = b.value || '';
      cv.font = { name: FUENTE, size: 10 };
      cv.alignment = { vertical: 'top', wrapText: true };
      for (var c = 1; c <= 6; c++) ws.getCell(r, c).border = borde;
      r++;
      return;
    }

    if (b.tipo === 'tabla') {
      // Encabezados
      b.headers.forEach(function (h, i) {
        var ch = ws.getCell(r, i + 1);
        ch.value = h;
        ch.font = { name: FUENTE, size: 10, bold: true, color: { argb: BLANCO } };
        ch.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NEGRO } };
        ch.alignment = { vertical: 'middle', wrapText: true };
        ch.border = borde;
      });
      r++;
      // Filas
      if (b.rows.length === 0) {
        b.headers.forEach(function (h, i) { ws.getCell(r, i + 1).border = borde; });
        r++;
      } else {
        b.rows.forEach(function (fila, idx) {
          fila.forEach(function (celdaVal, i) {
            var cc = ws.getCell(r, i + 1);
            cc.value = celdaVal;
            cc.font = { name: FUENTE, size: 10 };
            cc.alignment = { vertical: 'top', wrapText: true };
            cc.border = borde;
            if (idx % 2 === 1) {
              cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GRIS_CLARO } };
            }
          });
          r++;
        });
      }
      r++; // espacio tras la tabla
      return;
    }
  });

  return r;
}

function descargar(wb, nombre) {
  wb.xlsx.writeBuffer().then(function (buffer) {
    var blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = nombre;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });
}

// Lee el archivo (imagen) adjunto en un input file y agrega una pestaña "Sello de aprobación".
// Devuelve una promesa que se resuelve cuando terminó (haya o no imagen).
function agregarPestanaSello(wb, inputId) {
  return new Promise(function (resolve) {
    var input = document.getElementById(inputId);
    if (!input || !input.files || input.files.length === 0) {
      resolve(false); // no hay archivo adjunto
      return;
    }
    var file = input.files[0];
    var tipo = (file.type || '').toLowerCase();
    // Solo imágenes (el sello es PNG/JPG)
    var esPng = tipo.indexOf('png') !== -1;
    var esJpg = tipo.indexOf('jpeg') !== -1 || tipo.indexOf('jpg') !== -1;
    if (!esPng && !esJpg) {
      resolve(false); // no es imagen, no se puede incrustar
      return;
    }
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var base64 = e.target.result; // data URL
        var imageId = wb.addImage({ base64: base64, extension: esPng ? 'png' : 'jpeg' });
        var hoja = wb.addWorksheet('Sello de aprobación');
        hoja.getCell('A1').value = 'Sello de aprobación adjunto';
        hoja.getCell('A1').font = { name: FUENTE, size: 12, bold: true, color: { argb: NEGRO } };
        hoja.getColumn(1).width = 12;
        // Insertar la imagen a partir de la fila 3
        hoja.addImage(imageId, {
          tl: { col: 0, row: 2 },
          ext: { width: 700, height: 480 }
        });
        resolve(true);
      } catch (err) {
        resolve(false);
      }
    };
    reader.onerror = function () { resolve(false); };
    reader.readAsDataURL(file);
  });
}

// ================= FASE 1 =================
function exportarFase1() {
  var wb = new ExcelJS.Workbook();
  var ws = wb.addWorksheet('F-BAI05-01');

  var disposicion = filasDinamicas('tbody-disposicion', 4);
  var equipo = filasDinamicas('tbody-equipo', 4);

  // Decisión de aprobación
  var decEl = document.querySelector('input[name="f1_decision"]:checked');
  var decision = decEl ? decEl.value : '';
  var bloqueAprob;
  if (decision === 'no') {
    bloqueAprob = [
      { tipo: 'seccion', texto: 'Sección 5 — Decisión de Aprobación' },
      { tipo: 'campo', label: 'Resultado', value: 'RECHAZADO — No se aprueba el avance a la Fase 2' },
      { tipo: 'campo', label: 'Justificación del rechazo', value: val('f1_justificacion_rechazo') }
    ];
  } else {
    bloqueAprob = [
      { tipo: 'seccion', texto: 'Sección 5 — Decisión de Aprobación' },
      { tipo: 'campo', label: 'Resultado', value: 'APROBADO — Se autoriza el avance a la Fase 2' },
      { tipo: 'campo', label: 'OFGECA (Obligatorio) — Nombre y cargo', value: val('f1_firma_ofgeca_nombre') },
      { tipo: 'campo', label: 'Sello de aprobación', value: 'Adjunto en la pestaña "Sello de aprobación"' }
    ];
  }

  var bloques = [
      { tipo: 'seccion', texto: 'Sección 1 — Identificación del Cambio' },
      { tipo: 'campo', label: 'Código del cambio', value: val('f1_codigo') },
      { tipo: 'campo', label: 'Fecha de solicitud', value: val('f1_fecha_solicitud') },
      { tipo: 'campo', label: 'Nombre del cambio', value: val('f1_nombre') },
      { tipo: 'campo', label: 'Solicitante / Promotor', value: val('f1_solicitante') },
      { tipo: 'campo', label: 'Área solicitante', value: val('f1_area') },
      { tipo: 'campo', label: 'Responsable del Cambio (RC)', value: val('f1_rc') },
      { tipo: 'campo', label: 'Fecha estimada de inicio', value: val('f1_fecha_inicio') },

      { tipo: 'seccion', texto: 'Sección 2 — Descripción y Justificación del Cambio' },
      { tipo: 'campo', label: 'Objetivo del cambio', value: val('f1_objetivo') },
      { tipo: 'campo', label: 'Procesos críticos afectados', value: val('f1_procesos') },
      { tipo: 'campo', label: 'Criterios de activación (lineamiento 5.1)', value: checkboxesSeleccionados('criterio') },
      { tipo: 'campo', label: 'Riesgos de no gestionar', value: val('f1_riesgos') },

      { tipo: 'seccion', texto: 'Sección 3 — Evaluación de Disposición al Cambio' },
      { tipo: 'tabla', headers: ['Área afectada', 'Nivel de disposición', 'Riesgos de resistencia', 'Acciones previas recomendadas'], rows: disposicion },

      { tipo: 'seccion', texto: 'Sección 4 — Equipo de Implementación' },
      { tipo: 'tabla', headers: ['Nombre', 'Área', 'Rol en el equipo', 'Disponibilidad (%)'], rows: equipo }
  ].concat(bloqueAprob);

  construirHoja(ws,
    'F-BAI05-01 — Formulario de Gestión del Cambio Organizativo (Fase 1)',
    'Inicio y Aprobación — BAI05 Gestión del Cambio Organizativo',
    'BAI05.01 Establecer el deseo de cambiar  |  BAI05.02 Formar un equipo de implementación eficaz',
    bloques
  );

  var fecha = new Date().toISOString().slice(0, 10);
  var codigo = val('f1_codigo') || 'sin-codigo';
  var sufijo = (decision === 'no') ? '_RECHAZADO' : '';
  var promSello = (decision === 'no') ? Promise.resolve(false) : agregarPestanaSello(wb, 'f1_evidencia_ofgeca');
  promSello.then(function () {
    descargar(wb, 'F-BAI05-01_' + codigo + '_' + fecha + sufijo + '.xlsx');
  });
}

// ================= FASE 2 =================
function exportarFase2() {
  var wb = new ExcelJS.Workbook();
  var ws = wb.addWorksheet('F-BAI05-02');

  var comunicacion = filasDinamicas('tbody-comunicacion', 5);
  var capacitacion = filasDinamicas('tbody-capacitacion', 5);
  var ganancias    = filasDinamicas('tbody-ganancias', 3);
  var resistencia  = filasDinamicas('tbody-resistencia', 3);
  var metricas     = filasDinamicas('tbody-metricas', 5);

  // Decisión de aprobación
  var decEl = document.querySelector('input[name="f2_decision"]:checked');
  var decision = decEl ? decEl.value : '';
  var bloqueAprob;
  if (decision === 'no') {
    bloqueAprob = [
      { tipo: 'seccion', texto: 'Decisión de Aprobación del Plan' },
      { tipo: 'campo', label: 'Resultado', value: 'RECHAZADO — No se aprueba el avance a la Fase 3' },
      { tipo: 'campo', label: 'Justificación del rechazo', value: val('f2_justificacion_rechazo') }
    ];
  } else {
    bloqueAprob = [
      { tipo: 'seccion', texto: 'Decisión de Aprobación del Plan' },
      { tipo: 'campo', label: 'Resultado', value: 'APROBADO — Se autoriza el avance a la Fase 3' },
      { tipo: 'campo', label: 'OFGECA (Obligatorio) — Nombre y cargo', value: val('f2_firma_ofgeca_nombre') },
      { tipo: 'campo', label: 'Sello de aprobación', value: 'Adjunto en la pestaña "Sello de aprobación"' },
      { tipo: 'campo', label: 'Responsable del Cambio — Nombre', value: val('f2_firma_rc_nombre') },
      { tipo: 'campo', label: 'Responsable del Cambio — Fecha', value: val('f2_firma_rc_fecha') }
    ];
  }

  var bloques = [
      { tipo: 'seccion', texto: 'Identificación' },
      { tipo: 'campo', label: 'Código del cambio', value: val('f2_codigo') },
      { tipo: 'campo', label: 'Versión del plan', value: val('f2_version') },
      { tipo: 'campo', label: 'Nombre del cambio', value: val('f2_nombre') },
      { tipo: 'campo', label: 'Fecha de elaboración', value: val('f2_fecha') },

      { tipo: 'seccion', texto: 'Componente A — Plan de Comunicación (BAI05.03)' },
      { tipo: 'tabla', headers: ['Audiencia', 'Mensaje clave', 'Canal', 'Frecuencia', 'Responsable'], rows: comunicacion },

      { tipo: 'seccion', texto: 'Componente B — Capacitación (BAI05.04)' },
      { tipo: 'tabla', headers: ['Grupo a capacitar', 'Temas', 'Modalidad y duración', 'Fecha estimada', 'Ref. Recursos Humanos'], rows: capacitacion },

      { tipo: 'seccion', texto: 'Componente B — Ganancias Rápidas (BAI05.04)' },
      { tipo: 'tabla', headers: ['Ganancia rápida identificada', 'Cómo se comunicó', 'Fecha'], rows: ganancias },

      { tipo: 'seccion', texto: 'Componente B — Gestión de Resistencia (BAI05.04)' },
      { tipo: 'tabla', headers: ['Resistencia identificada', 'Acción tomada', 'Resultado obtenido'], rows: resistencia },

      { tipo: 'seccion', texto: 'Componente C — Métricas de Adopción (BAI05.05)' },
      { tipo: 'tabla', headers: ['Métrica', 'Meta mínima', 'Meta óptima', 'Resultado real', 'Fuente'], rows: metricas }
  ].concat(bloqueAprob);

  construirHoja(ws,
    'F-BAI05-02 — Formulario de Gestión del Cambio Organizativo (Fase 2)',
    'Ejecución — BAI05 Gestión del Cambio Organizativo',
    'BAI05.03 Comunicar la visión  |  BAI05.04 Facultar roles  |  BAI05.05 Habilitar operación',
    bloques
  );

  var fecha = new Date().toISOString().slice(0, 10);
  var codigo = val('f2_codigo') || 'sin-codigo';
  var sufijo = (decision === 'no') ? '_RECHAZADO' : '';
  var promSello = (decision === 'no') ? Promise.resolve(false) : agregarPestanaSello(wb, 'f2_evidencia_ofgeca');
  promSello.then(function () {
    descargar(wb, 'F-BAI05-02_' + codigo + '_' + fecha + sufijo + '.xlsx');
  });
}

// ================= FASE 3 =================
function exportarFase3() {
  var wb = new ExcelJS.Workbook();
  var ws = wb.addWorksheet('F-BAI05-03');

  var resultados = filasDinamicas('tbody-resultados', 5);

  construirHoja(ws,
    'F-BAI05-03 — Formulario de Gestión del Cambio Organizativo (Fase 3)',
    'Cierre y Sostenimiento — BAI05 Gestión del Cambio Organizativo',
    'BAI05.06 Incorporar nuevos enfoques  |  BAI05.07 Sostener cambios',
    [
      { tipo: 'seccion', texto: 'Identificación del Cambio' },
      { tipo: 'campo', label: 'Código del cambio', value: val('f3_codigo') },
      { tipo: 'campo', label: 'Fecha de cierre', value: val('f3_fecha_cierre') },
      { tipo: 'campo', label: 'Nombre del cambio', value: val('f3_nombre') },
      { tipo: 'campo', label: 'Responsable del Cambio', value: val('f3_rc') },

      { tipo: 'seccion', texto: 'Sección 1 — Resultados vs. Metas' },
      { tipo: 'tabla', headers: ['Métrica', 'Meta mínima', 'Meta óptima', 'Resultado final', 'Evaluación'], rows: resultados },

      { tipo: 'seccion', texto: 'Sección 2 — Lecciones Aprendidas' },
      { tipo: 'campo', label: '¿Qué funcionó bien durante este cambio?', value: val('f3_l1') },
      { tipo: 'campo', label: '¿Qué no funcionó bien o representó dificultades?', value: val('f3_l2') },
      { tipo: 'campo', label: '¿Qué se haría diferente en futuros cambios?', value: val('f3_l3') },
      { tipo: 'campo', label: '¿Qué acciones preventivas se recomiendan?', value: val('f3_l4') },
      { tipo: 'campo', label: '¿Cómo se distribuirán estas lecciones?', value: val('f3_l5') },

      { tipo: 'seccion', texto: 'Sección 3 — Aprobación de Cierre Formal' },
      { tipo: 'campo', label: 'OFGECA, Junta Directiva (Obligatorio) y Gerente General cuando aplique — Nombre y cargo', value: val('f3_firma_cierre_nombre') },
      { tipo: 'campo', label: 'Evidencia de aprobación (documento único)', value: 'Documento adjunto en el expediente' },
      { tipo: 'campo', label: 'Responsable del Cambio — Nombre', value: val('f3_firma_rc_nombre') },
      { tipo: 'campo', label: 'Responsable del Cambio — Fecha', value: val('f3_firma_rc_fecha') }
    ]
  );

  var fecha = new Date().toISOString().slice(0, 10);
  var codigo = val('f3_codigo') || 'sin-codigo';
  agregarPestanaSello(wb, 'f3_evidencia_cierre').then(function () {
    descargar(wb, 'F-BAI05-03_' + codigo + '_' + fecha + '.xlsx');
  });
}

// ================= SEGUIMIENTO DE SOSTENIMIENTO =================
function exportarSeguimiento() {
  var wb = new ExcelJS.Workbook();
  var ws = wb.addWorksheet('Seguimiento');

  construirHoja(ws,
    'Formulario de Seguimiento de Sostenimiento — Fase 3',
    'Cierre y Sostenimiento — BAI05 Gestión del Cambio Organizativo',
    'BAI05.06 Incorporar nuevos enfoques  |  BAI05.07 Sostener cambios',
    [
      { tipo: 'seccion', texto: 'Sección 1 — Datos del Cambio' },
      { tipo: 'campo', label: 'Código del cambio', value: val('fs_codigo') },
      { tipo: 'campo', label: 'Nombre del cambio', value: val('fs_nombre') },
      { tipo: 'campo', label: 'Responsable del Cambio (RC)', value: val('fs_rc') },
      { tipo: 'campo', label: 'Fecha de la revisión de sostenimiento', value: val('fs_fecha_revision') },

      { tipo: 'seccion', texto: 'Sección 2 — Criterios de Sostenimiento' },
      { tipo: 'campo', label: 'Criterios que se cumplen', value: checkboxesSeleccionados('fs_criterio') },

      { tipo: 'seccion', texto: 'Sección 3 — Acciones de Refuerzo' },
      { tipo: 'campo', label: 'Acciones de refuerzo definidas (en caso de debilitamiento)', value: val('fs_acciones_refuerzo') },
      { tipo: 'campo', label: 'Nota', value: 'El sostenimiento del cambio en la operación diaria queda a cargo del área dueña del proceso afectado (RDA). Los resultados de esta revisión se anexan al expediente del cambio.' },

      { tipo: 'seccion', texto: 'Sección 4 — Revisión Conjunta' },
      { tipo: 'campo', label: 'Responsable del Cambio (RC) — Nombre y cargo', value: val('fs_rc_nombre') },
      { tipo: 'campo', label: 'Responsable del Cambio (RC) — Fecha', value: val('fs_rc_fecha') },
      { tipo: 'campo', label: 'OFGECA — Nombre y cargo', value: val('fs_ofgeca_nombre') },
      { tipo: 'campo', label: 'OFGECA — Fecha', value: val('fs_ofgeca_fecha') }
    ]
  );

  var fecha = new Date().toISOString().slice(0, 10);
  var codigo = val('fs_codigo') || 'sin-codigo';
  descargar(wb, 'F-BAI05-Seguimiento_' + codigo + '_' + fecha + '.xlsx');
}
