// ============================================================
// >>> AGREGAR CAMPO NUEVO (lugar 1 de 5)
// Agrega el nombre de la columna al final de la lista.
// Estos nombres se escriben como encabezados en la fila 1.
// ============================================================
var ENCABEZADOS = ["ID", "Fecha y Hora", "Nombre", "¿Que te propones a mejorar en la escuela?", "¿Que cambiarias de la escuela?", "¿En que seccion vas?", "¿Cual es tu materia favorita?", "¿Te gusta la escuela?", "¿Entregas tus tareas?", "¿Como llegas a la escuela?", "¿Participas en clase?", "¿Se te complica la escuela?"];

function obtenerHoja() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(ENCABEZADOS);
  }
  return hoja;
}

function responder(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// CONSULTAR: devuelve todos los registros
// ------------------------------------------------------------
function doGet() {
  var datos = obtenerHoja().getDataRange().getValues();
  datos.shift(); // quita los encabezados

  var registros = datos.map(function (fila) {
    return {
      id: fila[0],
      fecha: fila[1],
      nombre: fila[2],
      mejorar: fila[3],
      cambio: fila[4],
      // >>> AGREGAR CAMPO NUEVO (lugar 2 de 5)
      // Agrega una coma después del campo anterior y mapea el índice de la columna
      seccion: fila[5],
      materia: fila[6],
      school: fila[7],
      tarea: fila[8],
      transporte: fila[9],
      participacion: fila[10],
      dificultad: fila[11]
    };
  });

  return responder(registros);
}

// ------------------------------------------------------------
// GUARDAR, EDITAR o ELIMINAR
// ------------------------------------------------------------
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    var p = e.parameter;
    var hoja = obtenerHoja();

    // ---------- GUARDAR ----------
    if (p.accion === "guardar") {
      // >>> AGREGAR CAMPO NUEVO (lugar 3 de 5)
      // Agrega el parámetro recibido al final del arreglo de inserción.
      hoja.appendRow([Date.now(), new Date(), p.nombre, p.mejorar, p.cambio, p.seccion, p.materia, p.school, p.tarea, p.transporte, p.participacion, p.dificultad]);
      return responder({ ok: true });
    }

    var fila = buscarFila(hoja, p.id);
    if (fila === -1) {
      return responder({ ok: false, mensaje: "No se encontró el registro" });
    }

    // ---------- EDITAR ----------
    if (p.accion === "editar") {
      // >>> AGREGAR CAMPO NUEVO (lugar 4 de 5)
      // a) Agrega el nuevo campo dentro de setValues (ej. p.animal).
      // b) Cambia el ÚLTIMO número del getRange por la cantidad total de
      //    campos que se están actualizando. Ahora son 4 campos (nombre, edad, cargo, animal) -> getRange(fila, 3, 1, 4)
      hoja.getRange(fila, 3, 1, 10).setValues([[p.nombre, p.mejorar, p.cambio, p.seccion, p.materia, p.school, p.tarea, p.transporte, p.participacion, p.dificultad]]);
      return responder({ ok: true });
    }

    // ---------- ELIMINAR ----------
    if (p.accion === "eliminar") {
      hoja.deleteRow(fila);
      return responder({ ok: true });
    }

    return responder({ ok: false, mensaje: "Acción no válida" });

  } catch (error) {
    return responder({ ok: false, mensaje: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------------------------
// Busca el número de fila que tiene ese ID (-1 si no existe)
// ------------------------------------------------------------
function buscarFila(hoja, id) {
  var ids = hoja.getRange(1, 1, hoja.getLastRow(), 1).getValues();
  for (var i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      return i + 1; // +1 porque las filas empiezan en 1
    }
  }
  return -1;
}

// ============================================================
// >>> AGREGAR CAMPO NUEVO (lugar 5 de 5)
// Escribe el encabezado en la fila 1 de la hoja de cálculo,
// en la columna que sigue.
//
// DESPUÉS DE CAMBIAR ESTE ARCHIVO:
// Implementar > Administrar implementaciones > Editar (lápiz)
// > Versión: "Nueva versión" > Implementar
// ============================================================