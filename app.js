// =============================================
// BAI05 — app.js
// Navegación entre fases y filas dinámicas
// =============================================

// ---- NAVEGACIÓN ----
function mostrarFase(num) {
  document.querySelectorAll('.fase-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.fase-btn').forEach(b => b.classList.remove('active'));

  document.getElementById('fase-' + num).classList.add('active');
  document.querySelector('[data-fase="' + num + '"]').classList.add('active');

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ---- ELIMINAR FILA ----
function eliminarFila(btn) {
  const fila = btn.closest('tr');
  const tbody = fila.parentElement;
  if (tbody.rows.length > 1) {
    fila.remove();
  } else {
    alert('Debe haber al menos una fila en esta tabla.');
  }
}

// ---- FASE 1: DISPOSICIÓN AL CAMBIO ----
function agregarFilaDisposicion() {
  const tbody = document.getElementById('tbody-disposicion');
  const fila = document.createElement('tr');
  fila.className = 'fila-disposicion';
  fila.innerHTML = `
    <td><input type="text" placeholder="Área" /></td>
    <td>
      <select>
        <option value="">Seleccionar</option>
        <option>Alto</option>
        <option>Medio</option>
        <option>Bajo</option>
      </select>
    </td>
    <td><input type="text" placeholder="Describir" /></td>
    <td><input type="text" placeholder="Describir" /></td>
    <td><button class="btn-eliminar" onclick="eliminarFila(this)">✕</button></td>
  `;
  tbody.appendChild(fila);
}

// ---- FASE 1: EQUIPO ----
function agregarFilaEquipo() {
  const tbody = document.getElementById('tbody-equipo');
  const fila = document.createElement('tr');
  fila.className = 'fila-equipo';
  fila.innerHTML = `
    <td><input type="text" placeholder="Nombre" /></td>
    <td><input type="text" placeholder="Área" /></td>
    <td><input type="text" placeholder="Rol" /></td>
    <td><input type="number" placeholder="%" min="0" max="100" /></td>
    <td><button class="btn-eliminar" onclick="eliminarFila(this)">✕</button></td>
  `;
  tbody.appendChild(fila);
}

// ---- FASE 2: COMUNICACIÓN ----
function agregarFilaComunicacion() {
  const tbody = document.getElementById('tbody-comunicacion');
  const fila = document.createElement('tr');
  fila.innerHTML = `
    <td><input type="text" placeholder="Audiencia" /></td>
    <td><input type="text" placeholder="Mensaje" /></td>
    <td><input type="text" placeholder="Canal" /></td>
    <td><input type="text" placeholder="Frec." /></td>
    <td><input type="text" placeholder="Resp." /></td>
    <td><button class="btn-eliminar" onclick="eliminarFila(this)">✕</button></td>
  `;
  tbody.appendChild(fila);
}

// ---- FASE 2: CAPACITACIÓN ----
function agregarFilaCapacitacion() {
  const tbody = document.getElementById('tbody-capacitacion');
  const fila = document.createElement('tr');
  fila.innerHTML = `
    <td><input type="text" placeholder="Grupo" /></td>
    <td><input type="text" placeholder="Temas" /></td>
    <td><input type="text" placeholder="Modalidad" /></td>
    <td><input type="date" /></td>
    <td><input type="text" placeholder="Ref." /></td>
    <td><button class="btn-eliminar" onclick="eliminarFila(this)">✕</button></td>
  `;
  tbody.appendChild(fila);
}

// ---- FASE 2: RESISTENCIA ----
function agregarFilaResistencia() {
  const tbody = document.getElementById('tbody-resistencia');
  const fila = document.createElement('tr');
  fila.innerHTML = `
    <td><input type="text" placeholder="Descripción" /></td>
    <td><input type="text" placeholder="Acción" /></td>
    <td><input type="text" placeholder="Resultado" /></td>
    <td><button class="btn-eliminar" onclick="eliminarFila(this)">✕</button></td>
  `;
  tbody.appendChild(fila);
}
