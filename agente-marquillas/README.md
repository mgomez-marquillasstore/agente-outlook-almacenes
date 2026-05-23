# Agente Marquillas S.A.S — Verificador automático de facturas

## ¿Qué hace este sistema?

El agente revisa automáticamente el correo corporativo de Outlook cada 5 minutos (configurable). Cuando encuentra un correo con un archivo ZIP adjunto, lo descarga, extrae el PDF que hay adentro, le pide a la IA de Claude que analice si el NIT y la razón social corresponden a **Marquillas S.A.S (NIT 890900314)**. Si la factura es válida, la reenvía automáticamente al destinatario configurado. Si no es válida, registra el motivo en el log.

**Todo el procesamiento ocurre en memoria** — el ZIP y el PDF nunca se guardan en disco.

---

## Requisitos

- **Python 3.9 o superior** (recomendado: Python 3.12)
- **pip** (viene incluido con Python)
- Cuenta de **Microsoft 365 corporativa** con acceso a la casilla a monitorear
- **API Key de Anthropic** (Claude AI)
- **Credenciales de Azure** configuradas por el área de IT (ver sección de Azure)

---

## Instalación en Windows (paso a paso)

### Paso 1 — Verificar que Python está instalado

Abre una ventana de **PowerShell** o **Símbolo del sistema** y ejecuta:

```
python --version
```

Debe mostrar `Python 3.9.x` o superior. Si no tienes Python, descárgalo desde [python.org](https://python.org) e instálalo marcando la opción "Add Python to PATH".

### Paso 2 — Ir a la carpeta del proyecto

```
cd C:\ruta\donde\descargaste\agente-marquillas
```

### Paso 3 — Crear un entorno virtual de Python

Un entorno virtual aísla las dependencias del proyecto sin afectar tu Python global:

```
python -m venv venv
```

### Paso 4 — Activar el entorno virtual

```
venv\Scripts\activate
```

Verás que el prompt cambia a `(venv) C:\...>`. Esto confirma que el entorno está activo.

### Paso 5 — Instalar las dependencias

```
pip install -r requirements.txt
```

Este comando descarga e instala todas las bibliotecas necesarias. Puede tardar unos minutos la primera vez.

### Paso 6 — Configurar el archivo .env

```
copy .env.example .env
```

Luego abre el archivo `.env` con el Bloc de notas o VS Code y completa cada valor según las instrucciones de la siguiente sección.

---

## Configuración del archivo .env

Abre el archivo `.env` y reemplaza cada valor de ejemplo con los datos reales:

| Variable | Descripción | Ejemplo |
|----------|-------------|---------|
| `ANTHROPIC_API_KEY` | Clave de API de Anthropic | `sk-ant-api03-...` |
| `AZURE_CLIENT_ID` | ID de la app registrada en Azure | `a1b2c3d4-...` |
| `AZURE_CLIENT_SECRET` | Secreto generado en Azure | `abc~XYZ...` |
| `AZURE_TENANT_ID` | ID del directorio de la organización | `f1e2d3c4-...` |
| `EMAIL_MONITOREAR` | Correo donde llegan las facturas | `analista@empresa.com` |
| `EMAIL_DESTINO` | Correo donde se reenvían las aprobadas | `contabilidad@empresa.com` |
| `INTERVALO_MINUTOS` | Frecuencia de revisión en minutos | `5` |

**Cómo obtener la API Key de Anthropic:**
1. Ve a [console.anthropic.com](https://console.anthropic.com)
2. Inicia sesión o crea una cuenta
3. Ve a **API Keys** → **Create Key**
4. Copia la clave y pégala en `ANTHROPIC_API_KEY`

---

## Cómo configurar Azure App Registration (para el área de IT)

El agente accede al correo corporativo usando la API de Microsoft. Para esto se necesita registrar una aplicación en Azure Active Directory. Estos pasos los realiza **el área de IT** con acceso al portal de Azure.

### Paso 1 — Registrar la aplicación

1. Ir a [portal.azure.com](https://portal.azure.com) e iniciar sesión con una cuenta de administrador
2. Buscar **"Azure Active Directory"** en el menú
3. Ir a **Registros de aplicaciones** → **Nuevo registro**
4. Completar:
   - **Nombre**: `Agente Marquillas Facturas`
   - **Tipos de cuenta**: "Solo cuentas de este directorio organizativo"
   - **URI de redirección**: dejar vacío
5. Hacer clic en **Registrar**

### Paso 2 — Anotar las credenciales

En la pantalla de la aplicación recién creada, anotar:
- **Id. de aplicación (cliente)** → es el `AZURE_CLIENT_ID`
- **Id. de directorio (inquilino)** → es el `AZURE_TENANT_ID`

### Paso 3 — Crear el secreto de cliente

1. En el menú izquierdo ir a **Certificados y secretos**
2. Hacer clic en **Nuevo secreto de cliente**
3. Descripción: `Agente Facturas 2025`
4. Duración: 24 meses (o la que corresponda a la política de la organización)
5. Hacer clic en **Agregar**
6. **COPIAR INMEDIATAMENTE** el valor del secreto (solo se muestra una vez)
7. Ese valor es el `AZURE_CLIENT_SECRET`

### Paso 4 — Asignar permisos a la aplicación

1. En el menú izquierdo ir a **Permisos de API**
2. Hacer clic en **Agregar un permiso** → **Microsoft Graph** → **Permisos de aplicación**
3. Buscar y marcar los siguientes permisos:
   - `Mail.Read` — para leer correos
   - `Mail.Send` — para enviar correos
   - `Mail.ReadWrite` — para marcar como leído
4. Hacer clic en **Agregar permisos**
5. Hacer clic en **Conceder consentimiento de administrador para [organización]**
6. Confirmar haciendo clic en **Sí**

> ⚠️ **Importante**: Sin el paso de "Conceder consentimiento de administrador" el agente no podrá acceder al correo aunque tenga las credenciales correctas.

---

## Cómo ejecutar el agente

### Opción A — Prueba local (sin credenciales de Microsoft)

Esta es la forma recomendada para probar mientras se consiguen las credenciales de Azure.
Solo necesita `ANTHROPIC_API_KEY` en el `.env`.

```
python prueba_local.py
```

El script te pedirá la ruta de un archivo ZIP con un PDF adentro, extraerá el texto y lo enviará a Claude para verificación.

### Opción B — Agente completo (con todas las credenciales)

Con el entorno virtual activado (`venv\Scripts\activate`):

```
python agente.py
```

El agente mostrará el banner de inicio, verificará la configuración y comenzará a revisar el correo automáticamente.

**Para detener el agente**: presionar `Ctrl + C` en la consola.

---

## Cómo ver los logs

Los registros de actividad se guardan automáticamente en la carpeta `logs/`:

```
agente-marquillas/
└── logs/
    └── agente.log       ← log actual
    └── agente.log.1     ← respaldo anterior (si superó 5MB)
    └── agente.log.2     ← respaldo más antiguo
```

Para ver el log en tiempo real mientras el agente está corriendo, abre otra ventana de PowerShell y ejecuta:

```powershell
Get-Content logs\agente.log -Wait -Tail 50
```

O simplemente abre el archivo `logs\agente.log` con el Bloc de notas o VS Code.

---

## Errores comunes y sus soluciones

### ❌ `No se encontró 'agente.md'`

**Causa**: El agente no está siendo ejecutado desde la carpeta correcta.

**Solución**: Asegúrate de ejecutar el comando desde la carpeta `agente-marquillas/`:
```
cd C:\ruta\completa\agente-marquillas
python agente.py
```

---

### ❌ `Azure rechazó las credenciales`

**Causa**: El `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET` o `AZURE_TENANT_ID` son incorrectos.

**Solución**:
1. Verificar que los valores en `.env` corresponden exactamente a los del portal de Azure
2. Verificar que el secreto no haya vencido (en Azure → Certificados y secretos)
3. El secreto no se puede ver después de crearlo — si se perdió, hay que generar uno nuevo

---

### ❌ `Error al consultar correos: 401 Unauthorized`

**Causa**: El token expiró o los permisos de la aplicación Azure no fueron otorgados correctamente.

**Solución**:
1. Verificar en Azure → Permisos de API que los permisos `Mail.Read`, `Mail.Send` y `Mail.ReadWrite` estén presentes
2. Confirmar que el administrador haya hecho clic en "Conceder consentimiento de administrador"

---

### ❌ `El PDF no tiene texto seleccionable`

**Causa**: La factura es una imagen escaneada (PDF de imagen, no de texto).

**Solución**: Claude intentará analizar el documento pero con resultados limitados. Para mejores resultados, solicitar al proveedor que envíe la factura en formato PDF generado digitalmente (no escaneado).

---

### ❌ `Claude no respondió con JSON válido`

**Causa**: Muy raramente Claude puede devolver texto adicional antes o después del JSON.

**Solución**: El agente ya tiene un mecanismo de limpieza automática. Si persiste, revisar el log para ver la respuesta exacta de Claude y ajustar las instrucciones en `agente.md`.

---

### ❌ `ModuleNotFoundError: No module named 'fitz'`

**Causa**: Las dependencias no están instaladas o el entorno virtual no está activo.

**Solución**:
```
venv\Scripts\activate
pip install -r requirements.txt
```

---

## Cómo modificar las reglas de validación

Todas las instrucciones que recibe Claude AI están en el archivo **`agente.md`**.

Para cambiar las reglas sin tocar el código Python:

1. Abrir `agente.md` con cualquier editor de texto
2. Modificar lo que necesites:
   - Para cambiar el NIT esperado: editar la línea `NIT exacto que debe aparecer: 890900314`
   - Para cambiar la razón social esperada: editar la línea `MARQUILLAS S.A.S`
   - Para agregar criterios adicionales: agregar instrucciones en la sección "Criterio de aprobación"
3. Guardar el archivo
4. El agente cargará las nuevas instrucciones **en el próximo ciclo de revisión** automáticamente (no es necesario reiniciarlo)

> 💡 **Tip**: Usa `prueba_local.py` para probar los cambios en `agente.md` sin necesitar credenciales de Microsoft.

También puedes cambiar el NIT y la razón social en el código `agente.py` (constantes `NIT_ESPERADO` y `RAZON_SOCIAL_ESPERADA`) aunque esto no afecta la lógica de Claude — solo se usan en los logs. La verificación real la hace Claude según lo que diga `agente.md`.

---

## Notas para Mac y Linux

En Mac/Linux los comandos son ligeramente diferentes:

```bash
# Crear y activar entorno virtual
python3 -m venv venv
source venv/bin/activate

# Instalar dependencias
pip install -r requirements.txt

# Ver logs en tiempo real
tail -f logs/agente.log
```

El resto del uso (configuración .env, ejecución) es idéntico a Windows.

---

## Despliegue en servidor (para más adelante)

Esta sección es para cuando se decida mover el agente a un servidor Linux en lugar de ejecutarlo en una PC local.

### Con systemd (Ubuntu/Debian)

Crear el archivo de servicio `/etc/systemd/system/agente-marquillas.service`:

```ini
[Unit]
Description=Agente Marquillas - Verificador de facturas
After=network.target

[Service]
Type=simple
User=ubuntu
WorkingDirectory=/home/ubuntu/agente-marquillas
ExecStart=/home/ubuntu/agente-marquillas/venv/bin/python agente.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

Habilitar y arrancar el servicio:

```bash
sudo systemctl daemon-reload
sudo systemctl enable agente-marquillas
sudo systemctl start agente-marquillas
sudo systemctl status agente-marquillas
```

### Con nohup (alternativa simple)

```bash
nohup python agente.py > /dev/null 2>&1 &
```

### Ver logs en servidor

```bash
tail -f /home/ubuntu/agente-marquillas/logs/agente.log
```
