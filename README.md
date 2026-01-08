# Sistema de Prórrogas de Contratos 📋

Un aplicativo web moderno y automatizado para gestionar prórrogas de contratos laborales. Diseñado específicamente para empresas que necesitan mantener un control centralizado de renovaciones contractuales con notificaciones automáticas y generación de documentos.

## 🎯 Características principales

- **Formulario Web Responsivo**: Interfaz moderna y fácil de usar para registrar nuevas prórrogas
- **Base de Datos Automática**: Almacenamiento centralizado en Google Sheets para todos los registros
- **Generación de Documentos**: Creación automática de documentos Word basados en plantillas para diferentes plazos (3, 6, 9 meses y anual)
- **Organización en Google Drive**: Estructura de carpetas automática por trabajador para mantener documentos organizados
- **Notificaciones por Email**: Alertas automáticas 45 días antes de que venza cada prórroga
- **Gestión de Prórrogas**: Seguimiento completo del ciclo de vida de cada prórroga (Pendiente, En Proceso, Completada)
- **Triggerizadores Automáticos**: Verificaciones diarias para detectar próximas fechas de vencimiento

## 🔧 Tecnologías utilizadas

- **Google Apps Script**: Backend serverless
- **Google Sheets API**: Base de datos
- **Google Drive API**: Gestión de documentos
- **Gmail API**: Sistema de notificaciones
- **HTML5 + CSS3 + JavaScript Vanilla**: Frontend interactivo
- **Font Awesome**: Iconografía

## 📋 Requisitos previos

Para utilizar este aplicativo necesitas:

- Una cuenta de Google Workspace (preferentemente)
- Acceso a Google Drive, Google Sheets y Gmail
- Permisos de administrador para crear triggers automáticos
- Plantillas de documentos Word preparadas en Google Drive

## 🚀 Instalación y configuración

### 1. Clonar el repositorio
```bash
git clone https://github.com/tu-usuario/prorrogas-contratos.git
cd prorrogas-contratos
```

### 2. Crear un proyecto en Google Apps Script

- Ve a [script.google.com](https://script.google.com)
- Crea un nuevo proyecto
- Copia el contenido de `CodigoGS.js` en el editor de Apps Script
- Copia el contenido de `formulario.html` en un archivo HTML diferente
- Actualiza la configuración en `appsscript.json`

### 3. Configurar las variables de entorno

En `CodigoGS.js`, actualiza el objeto `CONFIG` con:

```javascript
const CONFIG = {
  CARPETA_DRIVE_ID: 'Tu-ID-de-carpeta-en-Drive',
  DOCUMENTO_PLANTILLA_3_MESES_ID: 'Tu-ID-de-documento',
  DOCUMENTO_PLANTILLA_6_MESES_ID: 'Tu-ID-de-documento',
  DOCUMENTO_PLANTILLA_9_MESES_ID: 'Tu-ID-de-documento',
  DOCUMENTO_PLANTILLA_ANUAL_ID: 'Tu-ID-de-documento',
  SPREADSHEET_ID: 'Tu-ID-de-hoja-de-calculo',
  NOMBRE_HOJA: 'Registro_Prorrogas',
  DIAS_AVISO: 45,
  URL_FORMULARIO: 'Tu-URL-del-formulario',
  EMPRESA: 'Nombre de tu empresa',
  REPRESENTANTE_LEGAL: 'Nombre del representante',
  CC_REPRESENTANTE: 'Cédula del representante',
  CIUDAD: 'Tu ciudad'
};
```

**Para obtener los IDs:**
- **Carpeta/Documento ID**: Se encuentra en la URL. Por ejemplo: `docs.google.com/spreadsheets/d/1ABC2DEF3GHI4JKL5MNO6PQR7STU8VWX9YZ0/`
- El ID es: `1ABC2DEF3GHI4JKL5MNO6PQR7STU8VWX9YZ0`

### 4. Configurar permisos

- En Apps Script, autoriza el acceso a Drive, Gmail y Sheets
- Establece el nivel de acceso a "ANYONE_ANONYMOUS" (ajustable según necesidades)

### 5. Desplegar como aplicación web

- En Apps Script: Nuevo despliegue → Aplicación web
- Configura: Ejecutar como → Yo (tu cuenta)
- Acceso: Cualquiera
- Copia la URL de despliegue

## 📖 Cómo usar

1. **Accede a la aplicación web** usando la URL de despliegue
2. **Completa el formulario** con los datos del trabajador:
   - Nombre del trabajador
   - Número de cédula
   - Rol y cargo
   - Fecha de inicio del contrato
   - Nombre del jefe directo
3. **Envía el formulario**
4. El sistema automáticamente:
   - Crea un registro en la base de datos
   - Genera carpeta en Drive para el trabajador
   - Crea documentos según el tipo de prórroga
   - Envía notificación al jefe directo
   - Programa verificaciones automáticas

## 📊 Estructura de datos

### Hoja de Cálculo (Registro_Prorrogas)

| Campo | Descripción |
|-------|-------------|
| Fecha Registro | Cuándo se registró la prórroga |
| Nombre Trabajador | Nombre completo del empleado |
| Cédula | Número de identificación |
| Rol | Rol del empleado |
| Cargo | Puesto laboral |
| Fecha Inicio | Fecha de inicio del contrato |
| Jefe | Supervisor directo |
| Tipo Prórroga | Duración de la prórroga (6 meses) |
| Fecha Prórroga | Fecha de vencimiento |
| Estado | Pendiente/En Proceso/Completada |
| Carpeta Drive | ID de la carpeta en Drive |

## 🔐 Seguridad y privacidad

- Los datos se almacenan en Google Drive de tu organización
- Solo usuarios autenticados pueden enviar el formulario
- Se utilizan APIs de Google con autenticación OAuth 2.0
- Los datos no se comparten con terceros

## 🐛 Solución de problemas

### El formulario no se carga
- Verifica que has desplegado correctamente como aplicación web
- Comprueba que tienes acceso a la URL de despliegue

### Los emails no se envían
- Verifica que la cuenta de Google tiene Gmail activado
- Comprueba la configuración de dirección de email en el formulario
- Revisa los logs en Apps Script (Ejecuciones)

### Los documentos no se generan
- Asegúrate de que las plantillas existen en Drive
- Verifica que los IDs de las plantillas son correctos
- Comprueba que tienes permisos de edición en las plantillas

## 📝 Archivos del proyecto

```
├── CodigoGS.js          # Backend principal (Google Apps Script)
├── formulario.html      # Interfaz frontend
├── appsscript.json      # Configuración del proyecto
└── README.md            # Este archivo
```

## 🎨 Personalización

### Cambiar colores y estilos
Edita las variables CSS en `formulario.html`:
```css
:root {
    --primary: #2c3e50;
    --secondary: #3498db;
    --accent: #e74c3c;
    /* ... más variables */
}
```

### Cambiar campos del formulario
Modifica tanto `formulario.html` (interfaz) como `procesarFormulario()` en `CodigoGS.js` (lógica)

### Ajustar días de aviso
En `CONFIG`, cambia `DIAS_AVISO: 45` por el número de días deseado

## 📧 Contacto y soporte

Para reportar problemas o sugerencias, abre un issue en el repositorio de GitHub.

## 📄 Licencia

Este proyecto está bajo la licencia MIT. Ver `LICENSE` para más detalles.

---


**Desarrollado para** Productos Alimenticios Doria S.A.S.
**Desarrolado por KEVIN CAMILO DELGADO RESTREPO
