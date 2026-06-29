// =============================================
// BAI05 — exportar.js
// Exportación a Excel por fase con SheetJS
// =============================================

function val(id) {
  const el = document.getElementById(id);
  return el ? el.value : '';
}

function checkboxesSeleccionados(name) {
  return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`))
    .map(c => c.value).join(', ');
}

function radioSeleccionado(name) {
  const el = document.querySelector(`input[name="${name}"]:checked`);
  return el ? el.value : '';
}

function filasDinamicas(tbodyId, cols) {
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return [];
  return Array.from(tbody.rows).map(row => {
    const obj = {};
    cols.forEach((col, i) => {
      const celda = row.cells[i];
      if (!celda) return;
      const inp = celda.querySelector('input, select, textarea');
      obj[col] = inp ? inp.value : celda.textContent.trim();
    });
    return obj;
  });
}

function filaFija(labels, values) {
  return labels.map((l, i) => ({ label: l, value: values[i] || '' }));
}

function crearHoja(titulo, bloques) {
  const data = [];
  const merges = [];

  bloques.forEach(bloque => {
    if (bloque.tipo === 'titulo') {
      data.push([bloque.texto]);
      merges.push({ s: { r: data.length - 1, c: 0 }, e: { r: data.length - 1, c: 5 } });
      return;
    }

    if (bloque.tipo === 'seccion') {
      data.push([]);
      data.push([bloque.texto]);
      merges.push({ s: { r: data.length - 1, c: 0 }, e: { r: data.length - 1, c: 5 } });
      return;
    }

    if (bloque.tipo === 'campo') {
      data.push([bloque.label, bloque.value]);
      return;
    }

    if (bloque.tipo === 'tabla') {
      data.push(bloque.headers);
      bloque.rows.forEach(row => {
        data.push(bloque.headers.map(h => row[h] || ''));
      });
      data.push([]);
      return;
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!merges'] = merges;
  ws['!cols'] = [{ wch: 30 }, { wch: 28 }, { wch: 28 }, { wch: 22 }, { wch: 22 }, { wch: 18 }];
  return ws;
}

// ---- EXPORTAR FASE 1 ----
function exportarFase1() {
  const wb = XLSX.utils.book_new();

  const disposicion = filasDinamicas('tbody-disposicion', ['Área afectada', 'Nivel de disposición', 'Riesgos de resistencia', 'Acciones previas recomendadas']);
  const equipo = filasDinamicas('tbody-equipo', ['Nombre', 'Área', 'Rol en el equipo', 'Disponibilidad (%)']);

  const ws = crearHoja('F-BAI05-01', [
    { tipo: 'titulo', texto: 'F-BAI05-01 — Formulario de Solicitud y Evaluación del Cambio' },
    { tipo: 'campo', label: 'Código del cambio', value: val('f1_codigo') },
    { tipo: 'campo', label: 'Fecha de solicitud', value: val('f1_fecha_solicitud') },
    { tipo: 'campo', label: 'Nombre del cambio', value: val('f1_nombre') },
    { tipo: 'campo', label: 'Nivel de impacto', value: radioSeleccionado('f1_impacto') },
    { tipo: 'campo', label: 'Solicitante / Promotor', value: val('f1_solicitante') },
    { tipo: 'campo', label: 'Área solicitante', value: val('f1_area') },
    { tipo: 'campo', label: 'Responsable del Cambio (RC)', value: val('f1_rc') },
    { tipo: 'campo', label: 'Fecha estimada de inicio', value: val('f1_fecha_inicio') },

    { tipo: 'seccion', texto: 'SECCIÓN 2 — DESCRIPCIÓN Y JUSTIFICACIÓN' },
    { tipo: 'campo', label: 'Objetivo del cambio', value: val('f1_objetivo') },
    { tipo: 'campo', label: 'Procesos críticos afectados', value: val('f1_procesos') },
    { tipo: 'campo', label: 'Criterios de activación', value: checkboxesSeleccionados('criterio') },
    { tipo: 'campo', label: 'Riesgos de no gestionar', value: val('f1_riesgos') },

    { tipo: 'seccion', texto: 'SECCIÓN 3 — EVALUACIÓN DE DISPOSICIÓN AL CAMBIO' },
    { tipo: 'tabla', headers: ['Área afectada', 'Nivel de disposición', 'Riesgos de resistencia', 'Acciones previas recomendadas'], rows: disposicion },

    { tipo: 'seccion', texto: 'SECCIÓN 4 — EQUIPO DE IMPLEMENTACIÓN' },
    { tipo: 'tabla', headers: ['Nombre', 'Área', 'Rol en el equipo', 'Disponibilidad (%)'], rows: equipo },

    { tipo: 'seccion', texto: 'SECCIÓN 5 — APROBACIÓN DE INICIO' },
    { tipo: 'campo', label: 'Junta Directiva — Nombre y cargo', value: val('f1_firma_jd_nombre') },
    { tipo: 'campo', label: 'Junta Directiva — Fecha', value: val('f1_firma_jd_fecha') },
    { tipo: 'campo', label: 'Director de TI — Nombre', value: val('f1_firma_dti_nombre') },
    { tipo: 'campo', label: 'Director de TI — Fecha', value: val('f1_firma_dti_fecha') },
  ]);

  XLSX.utils.book_append_sheet(wb, ws, 'F-BAI05-01');

  const fecha = new Date().toISOString().slice(0, 10);
  const codigo = val('f1_codigo') || 'sin-codigo';
  XLSX.writeFile(wb, `F-BAI05-01_${codigo}_${fecha}.xlsx`);
}

// ---- EXPORTAR FASE 2 ----
function exportarFase2() {
  const wb = XLSX.utils.book_new();

  const comunicacion = filasDinamicas('tbody-comunicacion', ['Audiencia', 'Mensaje clave', 'Canal', 'Frecuencia', 'Responsable']);
  const capacitacion = filasDinamicas('tbody-capacitacion', ['Grupo a capacitar', 'Temas', 'Modalidad y duración', 'Fecha estimada', 'Ref. RRHH']);
  const resistencia  = filasDinamicas('tbody-resistencia',  ['Resistencia identificada', 'Acción tomada', 'Resultado obtenido']);

  const ws = crearHoja('F-BAI05-02', [
    { tipo: 'titulo', texto: 'F-BAI05-02 — Plan Integrado de Gestión del Cambio' },
    { tipo: 'campo', label: 'Código del cambio', value: val('f2_codigo') },
    { tipo: 'campo', label: 'Versión del plan', value: val('f2_version') },
    { tipo: 'campo', label: 'Nombre del cambio', value: val('f2_nombre') },
    { tipo: 'campo', label: 'Fecha de elaboración', value: val('f2_fecha') },

    { tipo: 'seccion', texto: 'COMPONENTE A — PLAN DE COMUNICACIÓN (BAI05.03)' },
    { tipo: 'tabla', headers: ['Audiencia', 'Mensaje clave', 'Canal', 'Frecuencia', 'Responsable'], rows: comunicacion },

    { tipo: 'seccion', texto: 'COMPONENTE B — CAPACITACIÓN (BAI05.04)' },
    { tipo: 'tabla', headers: ['Grupo a capacitar', 'Temas', 'Modalidad y duración', 'Fecha estimada', 'Ref. RRHH'], rows: capacitacion },

    { tipo: 'seccion', texto: 'COMPONENTE B — GESTIÓN DE RESISTENCIA (BAI05.04)' },
    { tipo: 'tabla', headers: ['Resistencia identificada', 'Acción tomada', 'Resultado obtenido'], rows: resistencia },

    { tipo: 'seccion', texto: 'COMPONENTE C — MÉTRICAS DE ADOPCIÓN (BAI05.05)' },
    { tipo: 'campo', label: '% usuarios capacitados — Resultado', value: val('m1_real') },
    { tipo: 'campo', label: '% procesos con nuevo esquema — Resultado', value: val('m2_real') },
    { tipo: 'campo', label: 'Satisfacción de usuarios — Resultado', value: val('m3_real') },
    { tipo: 'campo', label: val('m4_nombre') || 'Métrica adicional — Resultado', value: val('m4_real') },

    { tipo: 'seccion', texto: 'APROBACIÓN DEL PLAN' },
    { tipo: 'campo', label: 'Junta Directiva — Nombre y cargo', value: val('f2_firma_jd_nombre') },
    { tipo: 'campo', label: 'Junta Directiva — Fecha', value: val('f2_firma_jd_fecha') },
    { tipo: 'campo', label: 'Responsable del Cambio — Nombre', value: val('f2_firma_rc_nombre') },
    { tipo: 'campo', label: 'Responsable del Cambio — Fecha', value: val('f2_firma_rc_fecha') },
  ]);

  XLSX.utils.book_append_sheet(wb, ws, 'F-BAI05-02');

  const fecha = new Date().toISOString().slice(0, 10);
  const codigo = val('f2_codigo') || 'sin-codigo';
  XLSX.writeFile(wb, `F-BAI05-02_${codigo}_${fecha}.xlsx`);
}

// ---- EXPORTAR FASE 3 ----
function exportarFase3() {
  const wb = XLSX.utils.book_new();

  const evaluacion = (id) => {
    const el = document.getElementById(id);
    return el ? el.value : '';
  };

  const ws = crearHoja('F-BAI05-03', [
    { tipo: 'titulo', texto: 'F-BAI05-03 — Registro de Cierre y Lecciones Aprendidas' },
    { tipo: 'campo', label: 'Código del cambio', value: val('f3_codigo') },
    { tipo: 'campo', label: 'Fecha de cierre', value: val('f3_fecha_cierre') },
    { tipo: 'campo', label: 'Nombre del cambio', value: val('f3_nombre') },
    { tipo: 'campo', label: 'Responsable del Cambio', value: val('f3_rc') },

    { tipo: 'seccion', texto: 'SECCIÓN 1 — RESULTADOS VS. METAS' },
    { tipo: 'campo', label: '% usuarios capacitados — Resultado', value: val('r1_resultado') },
    { tipo: 'campo', label: '% usuarios capacitados — Evaluación', value: evaluacion('r1_eval') },
    { tipo: 'campo', label: '% procesos con nuevo esquema — Resultado', value: val('r2_resultado') },
    { tipo: 'campo', label: '% procesos con nuevo esquema — Evaluación', value: evaluacion('r2_eval') },
    { tipo: 'campo', label: 'Satisfacción de usuarios — Resultado', value: val('r3_resultado') },
    { tipo: 'campo', label: 'Satisfacción de usuarios — Evaluación', value: evaluacion('r3_eval') },
    { tipo: 'campo', label: 'Reducción de incidentes — Resultado', value: val('r4_resultado') },
    { tipo: 'campo', label: 'Reducción de incidentes — Evaluación', value: evaluacion('r4_eval') },

    { tipo: 'seccion', texto: 'SECCIÓN 2 — LECCIONES APRENDIDAS (BAI05.07)' },
    { tipo: 'campo', label: '¿Qué funcionó bien?', value: val('f3_l1') },
    { tipo: 'campo', label: '¿Qué no funcionó bien?', value: val('f3_l2') },
    { tipo: 'campo', label: '¿Qué se haría diferente?', value: val('f3_l3') },
    { tipo: 'campo', label: '¿Acciones preventivas recomendadas?', value: val('f3_l4') },
    { tipo: 'campo', label: '¿Cómo se distribuirán las lecciones?', value: val('f3_l5') },

    { tipo: 'seccion', texto: 'SECCIÓN 3 — APROBACIÓN DE CIERRE' },
    { tipo: 'campo', label: 'Junta Directiva — Nombre y cargo', value: val('f3_firma_jd_nombre') },
    { tipo: 'campo', label: 'Junta Directiva — Fecha', value: val('f3_firma_jd_fecha') },
    { tipo: 'campo', label: 'Director de TI — Nombre', value: val('f3_firma_dti_nombre') },
    { tipo: 'campo', label: 'Director de TI — Fecha', value: val('f3_firma_dti_fecha') },
    { tipo: 'campo', label: 'Responsable del Cambio — Nombre', value: val('f3_firma_rc_nombre') },
    { tipo: 'campo', label: 'Responsable del Cambio — Fecha', value: val('f3_firma_rc_fecha') },
  ]);

  XLSX.utils.book_append_sheet(wb, ws, 'F-BAI05-03');

  const fecha = new Date().toISOString().slice(0, 10);
  const codigo = val('f3_codigo') || 'sin-codigo';
  XLSX.writeFile(wb, `F-BAI05-03_${codigo}_${fecha}.xlsx`);
}
