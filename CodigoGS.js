// ============================================
// CONFIGURACIÓN PRINCIPAL
// ============================================

const CONFIG = {
  CARPETA_DRIVE_ID: '1wHXkMerwlNezPzQ_tSwgnJC9fXkG0iaG',
  DOCUMENTO_PLANTILLA_3_MESES_ID: '1TDdZhnmqR5WU0MLWE2uaXzLbXt-DDeQ56BQs2KGFhas',
  DOCUMENTO_PLANTILLA_6_MESES_ID: '1GjDA7y-6MIr2GxzxVmDZF1YutjIq6NX4v3mH16p7a3s',
  DOCUMENTO_PLANTILLA_9_MESES_ID: '1XnMd2Nl8GbzxIyL7n1J177KQWB9RFxQufzTEmaPbIG4',
  DOCUMENTO_PLANTILLA_ANUAL_ID: '1Q41vx3RBD57Jq4c1C_lZ9e9Fb4QE96IxvzigPILa9_g',
  SPREADSHEET_ID: '1fvf4YFfu4Bu9cCN8EAalkXxaZUS7mSlVEEnsiM8rFNc',
  NOMBRE_HOJA: 'Registro_Prorrogas',
  DIAS_AVISO: 45,
  URL_FORMULARIO: 'https://docs.google.com/forms/d/e/1FAIpQLSfqC_BzKBEUtIBYvJRjrAgpJ5o3_Mi-4ge7fpsviEC-ws75Yw/viewform',
  EMPRESA: 'PRODUCTOS ALIMENTICIOS DORIA S.A.S.',
  REPRESENTANTE_LEGAL: 'Leticia E. Echeverry Velásquez',
  CC_REPRESENTANTE: '30.336.930 de Manizales',
  CIUDAD: 'Mosquera'
};

// ============================================
// FUNCIÓN PRINCIPAL - ABRIR FORMULARIO WEB
// ============================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('formulario')
    .setTitle('Sistema de Prórrogas de Contratos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// PROCESAR FORMULARIO
// ============================================

function procesarFormulario(datos) {
  try {
    // 1. Crear o obtener la hoja de cálculo
    const sheet = obtenerOCrearHoja();
    
    // 2. Calcular fechas de prórrogas
    const fechaInicioParts = datos.fechaInicio.split('-');
    const fechaInicio = new Date(fechaInicioParts[0], fechaInicioParts[1] - 1, fechaInicioParts[2]);
    
    // 3. Guardar en la base de datos
    const fila = [
      new Date(), // Fecha de registro
      datos.nombreTrabajador,
      datos.cedula,
      datos.rol, // Added rol
      datos.cargo,
      Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      datos.jefe,
      'Prórroga Mensual (6 meses)', // Changed from 3 months to 6 months
      Utilities.formatDate(calcularProrrogas(fechaInicio)[0].fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      'Pendiente',
      '' // ID de carpeta Drive
    ];
    
    sheet.appendRow(fila);
    const filaNumero = sheet.getLastRow();
    
    // 4. Crear carpeta en Drive y documentos
    const carpetaId = crearCarpetaYDocumentos(datos, calcularProrrogas(fechaInicio), fechaInicio);
    
    // 5. Actualizar ID de carpeta en la hoja
    sheet.getRange(filaNumero, 11).setValue(carpetaId); // Changed column from 10 to 11
    
    // 6. Programar las siguientes prórrogas
    programarProximasProrrogas(datos, calcularProrrogas(fechaInicio), sheet);
    
    // 7. Enviar correo inicial
    enviarCorreoNotificacion(datos, calcularProrrogas(fechaInicio)[0], carpetaId);
    
    // 8. Configurar trigger para verificaciones diarias
    configurarTriggerDiario();
    
    return `Registro exitoso para ${datos.nombreTrabajador}. Se han programado ${calcularProrrogas(fechaInicio).length} prórrogas y se ha enviado la notificación al jefe.`;
    
  } catch (error) {
    Logger.log('Error en procesarFormulario: ' + error.toString());
    throw new Error('Error al procesar el formulario: ' + error.message);
  }
}

// ============================================
// GESTIÓN DE HOJA DE CÁLCULO
// ============================================

function obtenerOCrearHoja() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.NOMBRE_HOJA);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NOMBRE_HOJA);
    
    // Crear encabezados
    const headers = [
      'Fecha Registro',
      'Nombre Trabajador',
      'Cédula',
      'Rol',
      'Cargo',
      'Fecha Inicio Contrato',
      'Jefe',
      'Tipo Prórroga',
      'Fecha Prórroga',
      'Estado',
      'ID Carpeta Drive'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#667eea');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    
    // Ajustar anchos de columna
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 150); // Rol column
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 250);
    sheet.setColumnWidth(8, 180);
    sheet.setColumnWidth(9, 120);
    sheet.setColumnWidth(10, 100);
    sheet.setColumnWidth(11, 200);
  }
  
  return sheet;
}

function programarProximasProrrogas(datos, prorrogas, sheet) {
  // Guardar las prórrogas adicionales (desde la segunda en adelante)
  for (let i = 1; i < prorrogas.length; i++) {
    const fila = [
      new Date(),
      datos.nombreTrabajador,
      datos.cedula,
      datos.rol, // Added rol
      datos.cargo,
      datos.fechaInicio,
      datos.jefe,
      prorrogas[i].tipo,
      Utilities.formatDate(prorrogas[i].fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      'Programada',
      '' // Se llenará cuando se procese
    ];
    
    sheet.appendRow(fila);
  }
}

// ============================================
// CÁLCULO DE PRÓRROGAS
// ============================================

function calcularProrrogas(fechaInicio) {
  const prorrogas = [];
  const tiposProrrogas = [
    { tipo: 'Prórroga Mensual (6 meses)', meses: 6, plantilla: '6_meses' },
    { tipo: 'Prórroga Trimestral (9 meses)', meses: 9, plantilla: '9_meses' },
    { tipo: 'Prórroga Mensual (12 meses)', meses: 12, plantilla: '12_meses' },
    { tipo: 'Prórroga Anual (12 meses)', meses: 12, plantilla: 'anual' }
  ];
  
  tiposProrrogas.forEach(prorroga => {
    const fecha = new Date(fechaInicio);
    fecha.setMonth(fecha.getMonth() + prorroga.meses);
    
    prorrogas.push({
      tipo: prorroga.tipo,
      fecha: fecha,
      meses: prorroga.meses,
      plantilla: prorroga.plantilla
    });
  });
  
  return prorrogas;
}

// ============================================
// FUNCIONES DE FORMATO DE FECHAS EN ESPAÑOL
// ============================================

function numeroATexto(numero) {
  const unidades = ['', 'uno', 'dos', 'tres', 'cuatro', 'cinco', 'seis', 'siete', 'ocho', 'nueve'];
  const decenas = ['', '', 'veinte', 'treinta', 'cuarenta', 'cincuenta', 'sesenta', 'setenta', 'ochenta', 'noventa'];
  const especiales = ['diez', 'once', 'doce', 'trece', 'catorce', 'quince', 'dieciséis', 'diecisiete', 'dieciocho', 'diecinueve'];
  const veintitantos = ['veinte', 'veintiuno', 'veintidós', 'veintitrés', 'veinticuatro', 'veinticinco', 'veintiséis', 'veintisiete', 'veintiocho', 'veintinueve'];
  
  if (numero === 0) return 'cero';
  if (numero < 10) return unidades[numero];
  if (numero >= 10 && numero < 20) return especiales[numero - 10];
  if (numero >= 20 && numero < 30) return veintitantos[numero - 20];
  if (numero >= 30 && numero < 100) {
    const decena = Math.floor(numero / 10);
    const unidad = numero % 10;
    return unidad === 0 ? decenas[decena] : decenas[decena] + ' y ' + unidades[unidad];
  }
  
  return numero.toString();
}

function mesATexto(mes) {
  const meses = [
    'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'
  ];
  return meses[mes];
}

function añoATexto(año) {
  if (año >= 2000 && año < 3000) {
    const miles = Math.floor(año / 1000);
    const resto = año % 1000;
    
    let texto = 'dos mil';
    
    if (resto > 0) {
      if (resto < 100) {
        texto += ' ' + numeroATexto(resto);
      } else {
        const centenas = Math.floor(resto / 100);
        const decenasUnidades = resto % 100;
        
        const textoCentenas = ['', 'ciento', 'doscientos', 'trescientos', 'cuatrocientos', 'quinientos', 'seiscientos', 'setecientos', 'ochocientos', 'novecientos'];
        
        if (centenas === 1 && decenasUnidades === 0) {
          texto += ' cien';
        } else {
          texto += ' ' + textoCentenas[centenas];
          if (decenasUnidades > 0) {
            texto += ' ' + numeroATexto(decenasUnidades);
          }
        }
      }
    }
    
    return texto;
  }
  
  return año.toString();
}

function formatearFechaTexto(fecha) {
  const dia = fecha.getDate();
  const mes = fecha.getMonth();
  const año = fecha.getFullYear();
  
  const diaTexto = numeroATexto(dia).toUpperCase();
  const mesTexto = mesATexto(mes).toUpperCase();
  const añoTexto = añoATexto(año).toUpperCase();
  const añoFormateado = año.toString().substring(0, año.toString().length - 3) + '.' + año.toString().substring(año.toString().length - 3);
  
  return `${diaTexto} (${dia}) ${mesTexto} DOS MIL ${añoTexto.replace('DOS MIL ', '')} (${añoFormateado})`;
}

function formatearFechaTextoConDE(fecha) {
  const dia = fecha.getDate();
  const mes = fecha.getMonth();
  const año = fecha.getFullYear();
  
  const diaTexto = numeroATexto(dia).toUpperCase();
  const mesTexto = mesATexto(mes).toUpperCase();
  const añoTexto = añoATexto(año).toUpperCase();
  const añoFormateado = año.toString().substring(0, año.toString().length - 3) + '.' + año.toString().substring(año.toString().length - 3);
  
  return `${diaTexto} (${dia}) DE ${mesTexto} DE ${añoTexto.toUpperCase()} (${añoFormateado})`;
}

function formatearFechaTextoMinusculas(fecha) {
  const dia = fecha.getDate();
  const mes = fecha.getMonth();
  const año = fecha.getFullYear();
  
  const diaTexto = numeroATexto(dia);
  const mesTexto = mesATexto(mes);
  const añoTexto = añoATexto(año);
  const añoFormateado = año.toString().substring(0, año.toString().length - 3) + '.' + año.toString().substring(año.toString().length - 3);
  
  return `${diaTexto} (${dia}) de ${mesTexto} de ${añoTexto} (${añoFormateado})`;
}

// ============================================
// GESTIÓN DE GOOGLE DRIVE
// ============================================

function crearCarpetaYDocumentos(datos, prorrogas, fechaInicio) {
  try {
    const carpetaPrincipal = DriveApp.getFolderById(CONFIG.CARPETA_DRIVE_ID);
    
    // Crear carpeta con el nombre del trabajador
    const nombreCarpeta = `${datos.nombreTrabajador} - ${datos.cedula}`;
    const carpetaTrabajador = carpetaPrincipal.createFolder(nombreCarpeta);
    
    // Crear documentos de prórroga
    prorrogas.forEach((prorroga, index) => {
      crearDocumentoProrroga(datos, prorroga, index + 1, carpetaTrabajador, fechaInicio);
    });
    
    return carpetaTrabajador.getId();
    
  } catch (error) {
    Logger.log('Error en crearCarpetaYDocumentos: ' + error.toString());
    throw new Error('Error al crear carpeta en Drive: ' + error.message);
  }
}

// En la función crearDocumentoProrroga, modifica esta sección:

// En la función crearDocumentoProrroga, modifica la sección de anual:

function crearDocumentoProrroga(datos, prorroga, numero, carpeta, fechaInicio) {
  try {
    let plantillaId;
    switch(prorroga.plantilla) {
      case '6_meses':
        plantillaId = CONFIG.DOCUMENTO_PLANTILLA_3_MESES_ID;
        break;
      case '9_meses':
        plantillaId = CONFIG.DOCUMENTO_PLANTILLA_6_MESES_ID;
        break;
      case '12_meses':
        plantillaId = CONFIG.DOCUMENTO_PLANTILLA_9_MESES_ID;
        break;
      case 'anual':
        plantillaId = CONFIG.DOCUMENTO_PLANTILLA_ANUAL_ID;
        break;
      default:
        plantillaId = CONFIG.DOCUMENTO_PLANTILLA_3_MESES_ID;
    }
    
    const plantilla = DriveApp.getFileById(plantillaId);
    const nombreDoc = `${numero}. ${prorroga.tipo} - ${datos.nombreTrabajador}`;
    const copiaDoc = plantilla.makeCopy(nombreDoc, carpeta);
    
    // Open document and replace variables
    const doc = DocumentApp.openById(copiaDoc.getId());
    const body = doc.getBody();
    
    const fechaInicioProrroga = new Date(fechaInicio);
    fechaInicioProrroga.setMonth(fechaInicioProrroga.getMonth() + (prorroga.meses - 3)); // Start of this extension period
    
    const fechaFinProrroga = new Date(fechaInicio);
    fechaFinProrroga.setMonth(fechaFinProrroga.getMonth() + prorroga.meses); // End of this extension period
    
    // MODIFICACIÓN: Restar un día a la fecha de fin para todos los documentos
    fechaFinProrroga.setDate(fechaFinProrroga.getDate() - 1);
    
    const fechaActual = new Date();
    
    const esAnual = prorroga.plantilla === 'anual';
    
    let fechaInicioTexto, fechaFinTexto, fechaFirmaTexto;
    
    if (esAnual) {
      const fechaVencimientoAnterior = new Date(fechaInicio);
      fechaVencimientoAnterior.setMonth(fechaVencimientoAnterior.getMonth() + 9);
      
      const fechaProximoVencimiento = new Date(fechaInicio);
      fechaProximoVencimiento.setFullYear(fechaProximoVencimiento.getFullYear() + 2);
      
      // NUEVA MODIFICACIÓN: Restar un día también al próximo vencimiento de 2 años
      fechaProximoVencimiento.setDate(fechaProximoVencimiento.getDate() - 1);
      
      const fechaVencimientoNuevo = new Date(fechaFinProrroga); // Ya tiene un día menos
      
      const fechaVencimientoAnteriorTexto = formatearFechaTextoConDE(fechaVencimientoAnterior);
      const fechaVencimientoNuevoTexto = formatearFechaTextoConDE(fechaVencimientoNuevo);
      const fechaProximoVencimientoTexto = formatearFechaTextoConDE(fechaProximoVencimiento);
      const fechaIngresoTexto = formatearFechaTextoMinusculas(fechaInicio);
      fechaFirmaTexto = formatearFechaTextoMinusculas(fechaActual);
      
      body.replaceText('{{FECHA_VENCIMIENTO_ANTERIOR_TEXTO}}', fechaVencimientoAnteriorTexto);
      body.replaceText('{{FECHA_VENCIMIENTO_NUEVO_TEXTO}}', fechaVencimientoNuevoTexto);
      body.replaceText('{{FECHA_PROXIMO_VENCIMIENTO_2_AÑOS}}', fechaProximoVencimientoTexto);
      body.replaceText('{{FECHA_INGRESO_TEXTO}}', fechaIngresoTexto);
      
      // ELIMINAR LA LÍNEA DE FIRMA DEL DOCUMENTO ANUAL
      const textoBuscarFirma = 'Se firma en Mosquera, el {{FECHA_FIRMA_TEXTO}}';
      const textoBuscarFirma2 = 'Se firma en Mosquera, el ';
      body.replaceText(textoBuscarFirma, '');
      body.replaceText(textoBuscarFirma2, '');
      body.replaceText('{{FECHA_FIRMA_TEXTO}}', '');
      
    } else {
      fechaInicioTexto = formatearFechaTexto(fechaInicioProrroga);
      fechaFinTexto = formatearFechaTexto(fechaFinProrroga); // Ya tiene un día menos
      fechaFirmaTexto = formatearFechaTexto(fechaActual);
      
      body.replaceText('{{FECHA_INICIO_TEXTO}}', fechaInicioTexto);
      body.replaceText('{{FECHA_FIN_TEXTO}}', fechaFinTexto);
      
      // ELIMINAR LA LÍNEA DE FIRMA DE LOS DOCUMENTOS MENSUALES
      const textoBuscarFirma = 'Se firma en Mosquera, el {{FECHA_FIRMA_TEXTO}}';
      const textoBuscarFirma2 = 'Se firma en Mosquera, el ';
      body.replaceText(textoBuscarFirma, '');
      body.replaceText(textoBuscarFirma2, '');
      body.replaceText('{{FECHA_FIRMA_TEXTO}}', '');
    }
    
    // También keep numeric formats for compatibility
    const fechaInicioNumerico = Utilities.formatDate(fechaInicioProrroga, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const fechaFinNumerico = Utilities.formatDate(fechaFinProrroga, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const fechaFirmaNumerico = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const añoFirma = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'yyyy');
    
    // Simple replacements using placeholders
    body.replaceText('{{NOMBRE_COMPLETO}}', datos.nombreTrabajador);
    body.replaceText('{{NOMBRE_MAYUSCULAS}}', datos.nombreTrabajador.toUpperCase());
    body.replaceText('{{CEDULA}}', datos.cedula);
    body.replaceText('{{ROL}}', datos.rol);
    body.replaceText('{{ROL_MAYUSCULAS}}', datos.rol.toUpperCase());
    body.replaceText('{{CARGO}}', datos.cargo);
    body.replaceText('{{CARGO_MAYUSCULAS}}', datos.cargo.toUpperCase());
    
    // Keep numeric formats for backward compatibility
    body.replaceText('{{FECHA_INICIO}}', fechaInicioNumerico);
    body.replaceText('{{FECHA_FIN}}', fechaFinNumerico);
    body.replaceText('{{FECHA_FIRMA}}', ''); // Dejar vacío
    body.replaceText('{{AÑO}}', añoFirma);
    
    body.replaceText('{{EMPRESA}}', CONFIG.EMPRESA);
    body.replaceText('{{REPRESENTANTE_LEGAL}}', CONFIG.REPRESENTANTE_LEGAL);
    body.replaceText('{{CC_REPRESENTANTE}}', CONFIG.CC_REPRESENTANTE);
    body.replaceText('{{CIUDAD}}', CONFIG.CIUDAD);
    
    doc.saveAndClose();
    
    Logger.log(`Documento creado: ${nombreDoc} con fecha fin ajustada (un día antes)`);
    
  } catch (error) {
    Logger.log('Error en crearDocumentoProrroga: ' + error.toString());
  }
}

// ============================================
// ENVÍO DE CORREOS
// ============================================

function enviarCorreoNotificacion(datos, prorroga, carpetaId) {
  try {
    const urlFormulario = CONFIG.URL_FORMULARIO;
    const urlCarpeta = `https://drive.google.com/drive/folders/${carpetaId}`;
    
    const fechaProrroga = Utilities.formatDate(prorroga.fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    
    const asunto = `⚠️ RECORDATORIO: Prórroga de Contrato - ${datos.nombreTrabajador}`;
    
    const cuerpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; text-align: center; border-radius: 10px 10px 0 0;">
          <h1 style="color: white; margin: 0;">📋 Sistema de Prórrogas</h1>
          <p style="color: white; margin: 10px 0 0 0;">Productos Alimenticios Doria S.A.S.</p>
        </div>
        
        <div style="background: #f9f9f9; padding: 30px; border-radius: 0 0 10px 10px;">
          <h2 style="color: #667eea; margin-top: 0;">Recordatorio de Prórroga de Contrato</h2>
          
          <p>Estimado/a <strong>${datos.jefe}</strong>,</p>
          
          <p>Le recordamos que <strong>faltan ${CONFIG.DIAS_AVISO} días</strong> para la prórroga del contrato del siguiente colaborador:</p>
          
          <div style="background: white; padding: 20px; border-radius: 10px; margin: 20px 0; border-left: 4px solid #667eea;">
            <p style="margin: 5px 0;"><strong>👤 Trabajador:</strong> ${datos.nombreTrabajador}</p>
            <p style="margin: 5px 0;"><strong>🆔 Cédula:</strong> ${datos.cedula}</p>
            <p style="margin: 5px 0;"><strong>💼 Cargo:</strong> ${datos.cargo}</p>
            <p style="margin: 5px 0;"><strong>📅 Fecha de Prórroga:</strong> ${fechaProrroga}</p>
            <p style="margin: 5px 0;"><strong>📝 Tipo:</strong> ${prorroga.tipo}</p>
          </div>
          
          <p><strong>Documentos disponibles:</strong></p>
          <p>Puede acceder a los documentos de prórroga en la siguiente carpeta:</p>
          <p style="text-align: center; margin: 20px 0;">
            <a href="${urlCarpeta}" style="background: #667eea; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; display: inline-block;">📁 Ver Documentos en Drive</a>
          </p>
          
          <p><strong>Registrar nueva prórroga:</strong></p>
          <p>Si necesita registrar una nueva prórroga, puede hacerlo a través del siguiente formulario:</p>
          <p style="text-align: center; margin: 20px 0;">
            <a href="${urlFormulario}" style="background: #764ba2; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; display: inline-block;">📋 Ir al Formulario</a>
          </p>
          
          <hr style="border: none; border-top: 1px solid #ddd; margin: 30px 0;">
          
          <p style="font-size: 12px; color: #666; text-align: center;">
            Este es un correo automático generado por el Sistema de Prórrogas de Contratos.<br>
            Por favor, no responda a este correo.
          </p>
        </div>
      </div>
    `;
    
    MailApp.sendEmail({
      to: datos.jefe,
      subject: asunto,
      htmlBody: cuerpo
    });
    
    Logger.log(`Correo enviado a ${datos.jefe} para ${datos.nombreTrabajador}`);
    
  } catch (error) {
    Logger.log('Error en enviarCorreoNotificacion: ' + error.toString());
  }
}

// ============================================
// VERIFICACIÓN DIARIA DE PRÓRROGAS
// ============================================

function verificarProrrogasPendientes() {
  try {
    const sheet = obtenerOCrearHoja();
    const datos = sheet.getDataRange().getValues();
    
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const estado = fila[9]; // Updated column index for Estado
      
      if (estado === 'Programada' || estado === 'Pendiente') {
        const fechaProrrogaStr = fila[8]; // Updated column index for Fecha Prórroga
        const fechaProrroga = new Date(fechaProrrogaStr);
        fechaProrroga.setHours(0, 0, 0, 0);
        
        const diffTime = fechaProrroga - hoy;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        
        if (diffDays === CONFIG.DIAS_AVISO) {
          const datosEmpleado = {
            nombreTrabajador: fila[1],
            cedula: fila[2],
            rol: fila[3], // Added rol
            cargo: fila[4],
            fechaInicio: fila[5],
            jefe: fila[6]
          };
          
          const prorroga = {
            tipo: fila[7],
            fecha: fechaProrroga
          };
          
          const carpetaId = fila[10]; // Updated column index
          
          enviarCorreoNotificacion(datosEmpleado, prorroga, carpetaId);
          
          sheet.getRange(i + 1, 10).setValue('Notificado'); // Updated column index
          
          Logger.log(`Notificación enviada para ${datosEmpleado.nombreTrabajador} - ${prorroga.tipo}`);
        }
      }
    }
    
  } catch (error) {
    Logger.log('Error en verificarProrrogasPendientes: ' + error.toString());
  }
}

// ============================================
// CONFIGURACIÓN DE TRIGGERS
// ============================================

function configurarTriggerDiario() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'verificarProrrogasPendientes') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('verificarProrrogasPendientes')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
  
  Logger.log('Trigger diario configurado correctamente');
}

// ============================================
// FUNCIÓN DE PRUEBA
// ============================================

function probarSistema() {
  const datosPrueba = {
    nombreTrabajador: 'Juan Pérez García',
    cedula: '1234567890',
    rol: 'Analista', // Added rol
    cargo: 'Analista de Pruebas',
    fechaInicio: '2024-01-15',
    jefe: 'leticia.echeverry@doria.com.co'
  };
  
  const resultado = procesarFormulario(datosPrueba);
  Logger.log(resultado);
}



// Función para limpiar específicamente la línea de firma
function limpiarLineaFirma(doc) {
  const body = doc.getBody();
  
  // Buscar y eliminar cualquier párrafo que contenga texto relacionado con la firma
  const parrafos = body.getParagraphs();
  
  parrafos.forEach(parrafo => {
    const texto = parrafo.getText();
    if (texto.includes('Se firma en') || 
        texto.includes('firma en Mosquera') || 
        texto.includes('{{FECHA_FIRMA_TEXTO}}') ||
        texto.includes('{{FECHA_FIRMA}}')) {
      parrafo.clear(); // Limpiar el contenido del párrafo
    }
  });
  
  return doc;
}