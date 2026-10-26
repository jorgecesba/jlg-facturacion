import os
import shutil
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=120, bottom=120, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_header(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    run = h.runs[0]
    if level == 1:
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = RGBColor(15, 76, 58) # Emerald
    elif level == 2:
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 64, 175) # Blue
    elif level == 3:
        run.font.size = Pt(10.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(51, 65, 85) # Slate
    return h

def add_callout(doc, title, text, bg_hex="F0FDF4", border_hex="10B981"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_hex}"/>
            <w:top w:val="none"/>
            <w:right w:val="none"/>
            <w:bottom w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    run_title = p.add_run(f"{title}\n")
    run_title.font.bold = True
    run_title.font.size = Pt(10)
    run_title.font.color.rgb = RGBColor(15, 76, 58)
    
    run_text = p.add_run(text)
    run_text.font.size = Pt(9.5)
    run_text.font.color.rgb = RGBColor(51, 65, 85)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def style_table(tbl, col_widths, headers, data):
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    # Header Row
    hdr_cells = tbl.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        hdr_cells[i].width = Inches(col_widths[i])
        set_cell_background(hdr_cells[i], "0F4C3A")
        set_cell_margins(hdr_cells[i], top=100, bottom=100, left=120, right=120)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(255, 255, 255)
            
    # Data Rows
    for row_idx, row_data in enumerate(data):
        row = tbl.add_row()
        bg = "F8FAFC" if row_idx % 2 == 1 else "FFFFFF"
        for i, val in enumerate(row_data):
            cell = row.cells[i]
            cell.text = str(val)
            cell.width = Inches(col_widths[i])
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=90, bottom=90, left=120, right=120)
            p = cell.paragraphs[0]
            for run in p.runs:
                run.font.size = Pt(8.5)
                run.font.color.rgb = RGBColor(30, 41, 59)
            if i == 0 or i == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def add_p(doc, text, bold_prefix="", space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_b = p.add_run(bold_prefix + " ")
        r_b.font.bold = True
        r_b.font.size = Pt(10)
        r_b.font.color.rgb = RGBColor(15, 23, 42)
    r_t = p.add_run(text)
    r_t.font.size = Pt(10)
    r_t.font.color.rgb = RGBColor(51, 65, 85)
    return p

def add_bullet(doc, text, bold_prefix=""):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_b = p.add_run(bold_prefix + " ")
        r_b.font.bold = True
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = RGBColor(15, 23, 42)
    r_t = p.add_run(text)
    r_t.font.size = Pt(9.5)
    r_t.font.color.rgb = RGBColor(51, 65, 85)
    return p

# ==========================================
# INICIO DE GENERACIÓN DEL DOCUMENTO
# ==========================================
doc = Document()

# Márgenes estándar
for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

# ==========================================
# PORTADA / ENCABEZADO
# ==========================================
p_title = doc.add_paragraph()
p_title.paragraph_format.space_before = Pt(10)
p_title.paragraph_format.space_after = Pt(2)
run_org = p_title.add_run("GRUPO MERCANTIL EMPRESARIAL S.A. DE C.V.\n")
run_org.font.size = Pt(11)
run_org.font.bold = True
run_org.font.color.rgb = RGBColor(100, 116, 139)

run_main = p_title.add_run("JLG FACTURACIÓN | ECOSISTEMA CORPORATIVO DE VIÁTICOS Y AUDITORÍA FISCAL\n")
run_main.font.size = Pt(17)
run_main.font.bold = True
run_main.font.color.rgb = RGBColor(15, 76, 58)

run_sub = p_title.add_run("Manual Técnico y Operativo Integral: Portal Web, App Android Nativa, Generador CFDI 4.0, Módulo de Auditoría y Despliegue Cloud")
run_sub.font.size = Pt(11.5)
run_sub.font.color.rgb = RGBColor(30, 64, 175)

p_meta = doc.add_paragraph()
p_meta.paragraph_format.space_before = Pt(6)
p_meta.paragraph_format.space_after = Pt(14)
run_meta = p_meta.add_run("Versión: 3.0.0  |  Normativa: SAT Anexo 20 / CFDI 4.0 / Art. 28 LISR  |  Fecha de Actualización: Octubre 2026")
run_meta.font.size = Pt(9)
run_meta.font.italic = True
run_meta.font.color.rgb = RGBColor(100, 116, 139)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ==========================================
# 1. RESUMEN EJECUTIVO Y OBJETIVOS
# ==========================================
add_header(doc, "1. Resumen Ejecutivo del Ecosistema Jlg Facturación", level=1)
add_p(doc, "Jlg Facturación es una solución omnicanal de grado empresarial concebida para la gestión, control, comprobación y auditoría integral de viáticos corporativos en México. La plataforma fusiona un backend de alto rendimiento en Node.js, un portal web reactivo y una aplicación nativa para Android en Kotlin. El sistema automatiza el 100% de la conciliación fiscal conforme a los lineamientos del Anexo 20 del SAT (CFDI 4.0) y a las deducciones autorizadas bajo el Artículo 28 Fracción V de la Ley del Impuesto sobre la Renta (LISR).")

add_callout(doc, "Aspectos Clave y Logros de la Versión 3.0.0", 
    "• Generador y Validador Directo CFDI 4.0 en Android: Registro inmediato de comprobantes sin requerir la preexistencia de archivos XML descargados en el smartphone.\n"
    "• Ingreso Administrativo Unificado y Discreto: Acceso maestro integrado de forma natural en el directorio de colaboradores de la web, eliminando secciones o botones llamativos de administración externa.\n"
    "• Módulo de Administración Integral: Altas con 12 campos fiscales/operativos, edición en vivo de colaboradores registrados y rechazo con justificación obligatoria de auditoría.\n"
    "• Despliegue Cloud 24/7 en Render con CI/CD: Infraestructura en la nube con sincronización continua desde GitHub y descarga directa del instalador APK.",
    bg_hex="F0FDF4", border_hex="10B981")

# ==========================================
# 2. STACK TECNOLÓGICO Y ARQUITECTURA
# ==========================================
add_header(doc, "2. Arquitectura del Sistema y Tecnologías Utilizadas", level=1)
add_p(doc, "El ecosistema está construido sobre estándares modernos y librerías probadas en producción:")

tbl_tech = doc.add_table(rows=1, cols=4)
style_table(tbl_tech, [1.5, 1.8, 1.6, 1.6],
    ["Capa / Entorno", "Tecnología Principal", "Librerías / Dependencias Clave", "Función en el Ecosistema"],
    [
        ["Backend REST API", "Node.js v20+ / Express 4.19", "Bcrypt, Crypto, CORS, Body-Parser", "Motor fiscal, parseo XML, autenticación y persistencia."],
        ["Frontend Web", "Vanilla JS ES6+ (SPA)", "TailwindCSS 3.4, FontAwesome 6.5.1", "Portal responsivo con polling en vivo y diseño corporativo."],
        ["App Móvil Android", "Kotlin 1.9.24 / SDK 34", "Retrofit 2.11, OkHttp3, Material 3", "App nativa con generador CFDI 4.0 y dropzone XML."],
        ["Compilador Móvil", "Gradle 8.9 / Java 21 OpenJDK", "Android Jetpack ViewBinding, Coroutines", "Generación de APK debug y release optimizado."],
        ["Infraestructura Cloud", "Render Cloud (Web Service)", "Git / GitHub Actions, HTTPS Let's Encrypt", "Hosting 24/7 universal con despliegue continuo."]
    ])

add_bullet(doc, "Node.js y Express proporcionan una arquitectura no bloqueante capaz de procesar validaciones de XML simultáneas en menos de 20 milisegundos por comprobante.", "Rendimiento del Backend:")
add_bullet(doc, "Almacenamiento persistente en disco (data/comprobantes.json) con resguardo de transacciones en memoria, permitiendo tolerancia total a fallos y reinicios de servidor.", "Persistencia Confiable:")
add_bullet(doc, "Retrofit 2 gestiona peticiones asíncronas con corrutinas de Kotlin en segundo plano, evitando cualquier congelamiento de la interfaz de usuario en teléfonos móviles.", "Capa de Red Android:")

# ==========================================
# 3. AUTENTICACIÓN Y ROLES DE USUARIO
# ==========================================
add_header(doc, "3. Catálogo Oficial de Usuarios, Roles y Credenciales", level=1)
add_p(doc, "El sistema cuenta con dos niveles de acceso diferenciados pero integrados armónicamente:")

tbl_users = doc.add_table(rows=1, cols=6)
style_table(tbl_users, [1.4, 1.2, 1.1, 1.1, 0.9, 0.8],
    ["Nombre del Perfil", "Puesto Asignado", "RFC", "Destino / Sede", "Anticipo", "Clave / Contraseña"],
    [
        ["Julián Alberto López Morales", "Gerente Operaciones", "LOMJ880914HC3", "Monterrey, NL", "$6,000.00", "JALM-2026"],
        ["Sandra Urquilla Rosales", "Directora Comercial", "UROS920318M14", "Guadalajara, JAL", "$8,500.00", "SUR-2026"],
        ["Ángel Miramontes", "Auditor Contable Senior", "MIRA851120AB9", "Cancún, Q. ROO", "$12,000.00", "AM-2026"],
        ["Jorge Luis Lopez", "Coordinador de TI", "LOGJ010405V78", "Querétaro, QRO", "$7,500.00", "JLL-2026"],
        ["Damian Spears", "Dirección General (Admin)", "SPED850312XX1", "Corporativo Central", "Ilimitado", "Papaleta723!"]
    ])

add_callout(doc, "🔐 Acceso Unificado y Discreto del Administrador en la Web",
    "A solicitud de la dirección corporativa, se ha eliminado cualquier tarjeta o botón externo de 'Acceso Administrador' en la página de inicio. El perfil de 'Damian Spears — Dirección General' se encuentra integrado de forma completamente natural en el selector de colaboradores. Al seleccionarlo y escribir la contraseña maestra ('Papaleta723!'), el portal reconoce internamente sus permisos y desbloquea el Módulo de Administración sin llamar la atención de personal no autorizado.",
    bg_hex="FEF3C7", border_hex="F59E0B")

# ==========================================
# 4. GENERADOR Y VALIDADOR CFDI 4.0 (APP ANDROID)
# ==========================================
add_header(doc, "4. Generador Directo de Facturas CFDI 4.0 y Validación SAT", level=1)
add_p(doc, "Una de las mayores innovaciones de la versión 3.0.0 es la inclusión de un Generador y Emisor Digital de Facturas CFDI 4.0 directamente dentro de la aplicación móvil de Android:")

add_p(doc, "El usuario puede alternar entre '⚡ Generar Factura Directa' y '📂 Cargar Archivo XML'. Si no cuenta con el archivo XML en su almacenamiento, puede llenar los campos y timbrar la factura digital al vuelo.", "4.1 Dos Modos de Comprobación:")

add_p(doc, "El generador incluye todos los parámetros indispensables requeridos por la legislación fiscal mexicana:", "4.2 Campos Requeridos:")
add_bullet(doc, "Nombre legal del proveedor comercial (ej. Restaurante El Portón, City Express).", "1. Proveedor / Razón Social:")
add_bullet(doc, "Identificador fiscal a 12 o 13 caracteres con validación de estructura.", "2. RFC del Emisor:")
add_bullet(doc, "Número de serie o folio asignado a la nota o comprobante de compra.", "3. Folio Fiscal:")
add_bullet(doc, "Descripción detallada del consumo o servicio contratado en la comisión.", "4. Concepto del Gasto:")
add_bullet(doc, "Importe final pagado en Moneda Nacional. El generador calcula automáticamente la Base imponible y el IVA al 16%.", "5. Monto Total ($ MXN):")
add_bullet(doc, "Chips de selección rápida: ALIMENTOS, HOSPEDAJE, TRANSPORTE, GASOLINA u OTROS.", "6. Categoría del Gasto:")
add_bullet(doc, "Catálogo oficial del SAT: Tarjeta de Crédito (04), Tarjeta de Débito (28), Transferencia Electrónica SPEI (03) y Efectivo (01).", "7. Forma de Pago SAT:")

add_p(doc, "Para facilitar demostraciones y pruebas instantáneas sin escribir, la app integra 3 botones de un toque:", "4.3 Botones de Muestras Rápidas (1-Tap Testing):")
add_bullet(doc, "Restaurante El Portón por $720.00 MXN. Pasa todas las reglas y se deduce al 100% al no superar el límite de $750.00 diarios.", "• Muestra 1 (Alimentos Válida):")
add_bullet(doc, "Hoteles City Express por $986.00 MXN. Deducible íntegramente conforme a comprobación de pernocta.", "• Muestra 2 (Hospedaje Válido):")
add_bullet(doc, "Steak House Capital por $1,300.00 MXN. Demuestra el cálculo automático del Art. 28 LISR: autoriza $750.00 como deducible y traslada $550.00 a Excedente No Deducible.", "• Muestra 3 (Excedente en Alimentos):")

# ==========================================
# 5. MÓDULO DE ADMINISTRACIÓN Y AUDITORÍA
# ==========================================
add_header(doc, "5. Módulo de Administración, Altas y Auditoría Fiscal", level=1)
add_p(doc, "El panel de administración permite supervisar el gasto de toda la plantilla y modificar parámetros en vivo:")

add_p(doc, "Permite incorporar colaboradores comisionados registrando 12 campos obligatorios: Nombre Completo, RFC Fiscal, Correo Corporativo, Puesto, Departamento, Clave Única generada, Estatus, Destino del Viaje, Motivo de la Comisión, Fecha de Inicio, Fecha de Fin y Anticipo Otorgado.", "5.1 Alta de Nuevos Colaboradores:")

add_p(doc, "Cada tarjeta de colaborador cuenta con un botón '✏️ Editar' que abre un formulario prellenado para actualizar cualquiera de los 12 datos en tiempo real mediante el endpoint PUT /api/admin/usuarios/:id.", "5.2 Edición de Colaboradores Existentes:")

add_p(doc, "Al presionar el botón '✕ Rechazar' sobre una factura, el sistema solicita obligatoriamente el 'Motivo de Rechazo de Auditoría' (ej. 'Concepto no deducible según política interna'). Dicho motivo se guarda permanentemente en la factura y el importe deducible se recalcula a $0.00.", "5.3 Dictamen de Facturas con Motivo:")

# ==========================================
# 6. ESPECIFICACIÓN DE ENDPOINTS API REST
# ==========================================
add_header(doc, "6. Catálogo Completo de Endpoints API REST", level=1)
add_p(doc, "El backend expone 12 servicios estructurados bajo arquitectura REST:")

tbl_api = doc.add_table(rows=1, cols=4)
style_table(tbl_api, [0.8, 2.4, 2.0, 1.3],
    ["Método", "Endpoint", "Propósito Funcional", "Seguridad"],
    [
        ["GET", "/api/empleados", "Catálogo de personal con montos enmascarados.", "Público"],
        ["POST", "/api/empleados/:id/verificar-clave", "Valida la clave única alfanumérica del colaborador.", "Público"],
        ["GET", "/api/empleados/:id/viaje", "Liquidación actual, saldos y comprobantes registrados.", "Público"],
        ["POST", "/api/empleados/:id/validar-comprobante", "Valida XML CFDI 4.0 contra SAT y reglas del Art. 28 LISR.", "Público"],
        ["POST", "/api/admin/login", "Autenticación maestra con Bcrypt y generación de token.", "Rate Limited"],
        ["GET", "/api/admin/verificar-sesion", "Verifica la validez del token de sesión administrativa.", "Bearer Token"],
        ["POST", "/api/admin/logout", "Revoca la sesión activa del administrador.", "Bearer Token"],
        ["GET", "/api/admin/usuarios", "Listado maestro de usuarios con claves y viajes visibles.", "Bearer Token"],
        ["POST", "/api/admin/usuarios", "Alta de nuevo colaborador con 12 campos completos.", "Bearer Token"],
        ["PUT", "/api/admin/usuarios/:id", "Actualización integral de datos y comisión del usuario.", "Bearer Token"],
        ["PUT", "/api/admin/empleados/:id/comprobantes/:uuid/estatus", "Aprobación o Rechazo de factura con motivo de auditoría.", "Bearer Token"],
        ["GET", "/ViaticosApp.apk", "Descarga directa del instalador de la app móvil Android.", "Público"]
    ])

# ==========================================
# 7. GUÍA DE DESPLIEGUE CLOUD Y USO MÓVIL
# ==========================================
add_header(doc, "7. Guía de Conexión Cloud 24/7 y Descarga del APK", level=1)

add_callout(doc, "🌐 Infraestructura en Producción Cloud (Render 24/7)",
    "• URL Oficial en Vivo: https://jlg-facturacion.onrender.com/\n"
    "• Repositorio Oficial en GitHub: https://github.com/jorgecesba/jlg-facturacion.git\n"
    "• Despliegue Automático (CI/CD): Cualquier cambio confirmado a la rama 'main' dispara una nueva compilación y despliegue en menos de 2 minutos.\n"
    "• Descarga del APK desde la Web: Incluye un botón verde permanente '📱 Descargar App Android' en la cabecera del portal.",
    bg_hex="EFF6FF", border_hex="3B82F6")

add_header(doc, "7.1 Cómo Descargar e Instalar la Aplicación en Teléfonos Android", level=2)
add_p(doc, "Abre https://jlg-facturacion.onrender.com en el navegador de tu celular y presiona el botón 'Descargar App Android' en la esquina superior derecha, o copia el archivo ViaticosApp.apk desde la computadora.", "Paso 1: Descarga del APK:")
add_p(doc, "Abre el archivo descargado. Si Android solicita confirmación de seguridad, activa la casilla 'Permitir instalación de fuentes desconocidas' para ese navegador o explorador de archivos.", "Paso 2: Permisos de Instalación:")
add_p(doc, "La app se conecta de fábrica a los servidores de Render. Los colaboradores pueden comprobar gastos desde carretera, hoteles o aeropuertos usando datos 4G/5G.", "Paso 3: Conexión Automática:")
add_p(doc, "Si se trabaja en una red local de pruebas, el usuario puede pulsar el botón de engrane ⚙️ en la barra de la app e ingresar 'http://192.168.100.52:4000/api/' para apuntar al servidor local.", "Paso 4: Servidor Alternativo:")

# ==========================================
# 8. RESOLUCIÓN DE PREGUNTAS FRECUENTES (FAQ)
# ==========================================
add_header(doc, "8. Preguntas Frecuentes y Solución de Problemas (FAQ)", level=1)

add_p(doc, "¿Cómo se entera el portal web cuando un comisionado sube una factura desde su celular?", "Pregunta 1: Sincronización en Vivo")
add_bullet(doc, "El portal web ejecuta una sincronización reactiva en tiempo real cada 3 segundos. Cuando el empleado valida una factura en su teléfono, los saldos, el gráfico de barras y la tabla de comprobantes en la pantalla de la web se actualizan al instante sin presionar F5.")

add_p(doc, "¿Qué sucede con los impuestos locales y el IVA de las facturas?", "Pregunta 2: Tratamiento del IVA y Proporcionalidad")
add_bullet(doc, "El sistema calcula el IVA al 16% sobre la base deducible. Cuando un gasto de alimentos excede los $750.00 diarios, el IVA correspondiente al excedente se convierte en no acreditable de conformidad con el Artículo 5 Fracción I de la Ley del IVA.")

add_p(doc, "¿Dónde se almacena la información fiscal?", "Pregunta 3: Seguridad de Datos")
add_bullet(doc, "Las facturas y datos de colaboradores se almacenan de manera estructurada en portal-viaticos/data/comprobantes.json. Ninguna contraseña de administrador se guarda en texto plano; todas están hasheadas con Bcrypt.")

# ==========================================
# GUARDADO DE ARCHIVOS EN TODAS LAS RUTAS
# ==========================================
rutas_salida = [
    r"C:\Users\Jorge\Desktop\Damian\portal-viaticos\Documentacion_Portal_Viaticos_Mexico.docx",
    r"C:\Users\Jorge\Desktop\Damian\Documentacion_Jlg_Facturacion.docx",
    r"C:\Users\Jorge\Desktop\Damian\Manual_Completo_Jlg_Facturacion_Android.docx"
]

for ruta in rutas_salida:
    doc.save(ruta)
    print(f"✅ Documentación actualizada guardada en: {ruta}")
