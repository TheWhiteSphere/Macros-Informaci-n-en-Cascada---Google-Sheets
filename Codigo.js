function buscarConcatenadoEInsertarFilasEnVertical() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. CONFIGURACIÓN DE LAS PESTAÑAS (Verifica que se llamen exactamente así)
  var nombreOrigen = "Pestaña_Informacion"; 
  var nombreDestino = "Resultados";         
  
  var hojaOrigen = libro.getSheetByName(nombreOrigen); 
  var hojaDestino = libro.getSheetByName(nombreDestino);         
  
  if (!hojaOrigen || !hojaDestino) {
    SpreadsheetApp.getUi().alert("Error", "Revisa que los nombres de las pestañas coincidan exactamente con tu libro.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // =========================================================================
  // CONFIGURACIÓN DE COLUMNAS DE ORIGEN A EXTRAER
  // =========================================================================
  var origenK = 11;
  var origenQ = 17;
  var origenS = 19;
  
  // =========================================================================
  // CONFIGURACIÓN DE LA COLUMNA DE DESTINO DONDE CAERÁ EL BLOQUE VERTICAL
  // =========================================================================
  // Modifica este número según la columna que desees (B=2, C=3, D=4...)
  var columnaDestinoVertical = 2; 
  // =========================================================================
  
  // 2. Obtener la fila actual seleccionada por tu cursor en la hoja de destino
  var celdaActiva = hojaDestino.getActiveCell();
  var filaActivaDestino = celdaActiva.getRow();
  
  // 3. Leer la concatenación "XY" desde la COLUMNA A (Columna 1) de tu renglón actual
  var claveBusqueda = hojaDestino.getRange(filaActivaDestino, 1).getValue().toString().trim().toUpperCase(); 
  
  if (claveBusqueda === "") {
    SpreadsheetApp.getUi().alert("Aviso", "La celda de la Columna A en la fila seleccionada está vacía.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // 4. Cargar la hoja de información en memoria para realizar la búsqueda
  var rangoOrigen = hojaOrigen.getDataRange();
  var datosOrigen = rangoOrigen.getValues();
  
  var resultadoEncontrado = false;
  var filaDatosOrigenEncontrada = [];
  
  // BÚSQUEDA CORREGIDA: Escanea la celda de la Columna A de cada renglón en la hoja de origen
  for (var i = 0; i < datosOrigen.length; i++) {
    // Tomar estrictamente el valor de la columna A (índice 0) de la fila actual de origen
    var celdaA_Origen = (datosOrigen[i] && datosOrigen[i][0]) ? datosOrigen[i][0].toString() : "";
    var claveOrigen = celdaA_Origen.trim().toUpperCase(); 
    
    // Comparar las dos concatenaciones de forma limpia
    if (claveOrigen === claveBusqueda) {
      filaDatosOrigenEncontrada = datosOrigen[i]; 
      resultadoEncontrado = true;
      break; 
    }
  }
  
  // 5. ACCIÓN CRUCIAL: Insertar exactamente 3 filas nuevas debajo de la fila activa
  // Esto empuja tus siguientes concatenaciones hacia abajo de forma segura
  hojaDestino.insertRowsAfter(filaActivaDestino, 3);
  
  // Definir las posiciones de los 3 nuevos renglones creados
  var filaNueva1 = filaActivaDestino + 1;
  var filaNueva2 = filaActivaDestino + 2;
  var filaNueva3 = filaActivaDestino + 3;
  
  // 6. Verter los resultados en VERTICAL convirtiendo el texto a MAYÚSCULAS
  if (resultadoEncontrado) {
    
    var datoK = filaDatosOrigenEncontrada[origenK - 1];
    hojaDestino.getRange(filaNueva1, columnaDestinoVertical).setValue(datoK ? datoK.toString().toUpperCase().trim() : "");
    
    var datoQ = filaDatosOrigenEncontrada[origenQ - 1];
    hojaDestino.getRange(filaNueva2, columnaDestinoVertical).setValue(datoQ ? datoQ.toString().toUpperCase().trim() : "");
    
    var datoS = filaDatosOrigenEncontrada[origenS - 1];
    hojaDestino.getRange(filaNueva3, columnaDestinoVertical).setValue(datoS ? datoS.toString().toUpperCase().trim() : "");
    
    libro.toast("¡Bloque vertical insertado con éxito hacia abajo!", "Completado", 3);
    
  } else {
    // Si la combinación no existe, te indica la clave que no encontró para que revises
    hojaDestino.getRange(filaNueva1, columnaDestinoVertical).setValue("NO ENCONTRADO: " + claveBusqueda);
    hojaDestino.getRange(filaNueva2, columnaDestinoVertical).setValue("REVISAR COLUMNA A");
    hojaDestino.getRange(filaNueva3, columnaDestinoVertical).setValue("EN ORIGEN");
    libro.toast("La clave de la Columna A no se encuentra en el origen.", "Aviso", 4);
  }
  
  hojaDestino.autoResizeColumns(1, hojaDestino.getLastColumn());
}

// Crear el acceso directo vinculando la función del menú
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Macro Extraer')
    .addItem('Insertar Filas y Extraer en Vertical', 'buscarConcatenadoEInsertarFilasEnVertical')
    .addToUi();
}
