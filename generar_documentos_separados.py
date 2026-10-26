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

def add_header(doc, text, level=1, color_rgb=None):
    h = doc.add_heading(text, level=level)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    run = h.runs[0]
    if level == 1:
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = color_rgb or RGBColor(15, 76, 58) # Emerald default
    elif level == 2:
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 64, 175) # Blue
    elif level == 3:
        run.font.size = Pt(10.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(51, 65, 85) # Slate
    return h

def add_callout(doc, title, text, bg_hex="F0FDF4", border_hex="10B981", title_rgb=None):
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
    run_title.font.color.rgb = title_rgb or RGBColor(15, 76, 58)
    
    run_text = p.add_run(text)
    run_text.font.size = Pt(9.5)
    run_text.font.color.rgb = RGBColor(51, 65, 85)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def style_table(doc, tbl, col_widths, headers, data, header_bg="0F4C3A"):
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    # Header Row
    hdr_cells = tbl.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        hdr_cells[i].width = Inches(col_widths[i])
        set_cell_background(hdr_cells[i], header_bg)
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


# =============================================================================
# GENERADOR 1: DOCUMENTACIÓN PORTAL WEB & BACKEND API
# =============================================================================
def generar_doc_web(output_path):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # PORTADA
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(10)
    p_title.paragraph_format.space_after = Pt(2)
    run_org = p_title.add_run("GRUPO MERCANTIL EMPRESARIAL S.A. DE C.V.\n")
    run_org.font.size = Pt(11)
    run_org.font.bold = True
    run_org.font.color.rgb = RGBColor(100, 116, 139)

    run_main = p_title.add_run("JLG FACTURACIÓN | MANUAL TÉCNICO DEL PORTAL WEB Y BACKEND API\n")
    run_main.font.size = Pt(17)
    run_main.font.bold = True
    run_main.font.color.rgb = RGBColor(15, 76, 58)

    run_sub = p_title.add_run("Plataforma Web Reactiva, Servicios REST Node.js, Motor Fiscal CFDI 4.0 SAT, Despliegue Cloud en Render y Auditoría Contable")
    run_sub.font.size = Pt(11.5)
    run_sub.font.color.rgb = RGBColor(30, 64, 175)

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(6)
    p_meta.paragraph_format.space_after = Pt(14)
    run_meta = p_meta.add_run("Versión del Sistema: 3.0.0  |  Normativa Fiscal: SAT Anexo 20 / Art. 28 LISR  |  Entorno: Cloud Render (24/7)")
    run_meta.font.size = Pt(9)
    run_meta.font.italic = True
    run_meta.font.color.rgb = RGBColor(100, 116, 139)

    # 1. RESUMEN EJECUTIVO
    add_header(doc, "1. Resumen Ejecutivo del Portal Web y Servidor Central", level=1)
    add_p(doc, "El Portal Web de Jlg Facturación es una aplicación web empresarial (Single-Page Application) concebida para la gestión, fiscalización y comprobación de viáticos en México. Funciona en simbiosis con un motor backend en Node.js que procesa comprobantes fiscales digitales por internet (CFDI 4.0), contrastando cada factura contra los topes diarios de deducibilidad establecidos en el Artículo 28 Fracción V de la Ley del Impuesto sobre la Renta (LISR).")

    add_callout(doc, "Características Sobresalientes de la Plataforma Web v3.0",
        "• Acceso Directo y Discreto: El ingreso para el Administrador se encuentra 100% integrado dentro del directorio corporativo de colaboradores, eliminando botones externos o formularios visibles que alerten a terceros.\n"
        "• Sincronización Reactiva en Tiempo Real: Polling no intrusivo cada 3 segundos que detecta y refleja cualquier factura o cambio subido desde la app móvil Android sin necesidad de recargar la página.\n"
        "• Panel de Auditoría Fiscal: Módulo para revisar el detalle de comprobantes de cada empleado, autorizarlos o rechazarlos con captura obligatoria del motivo contable de auditoría.\n"
        "• Despliegue Continuo (CI/CD): Hospedado en la nube con Render y sincronizado con GitHub bajo protocolo HTTPS seguro.",
        bg_hex="F0FDF4", border_hex="10B981")

    # 2. STACK TECNOLÓGICO WEB & BACKEND
    add_header(doc, "2. Stack Tecnológico de la Plataforma Web y API", level=1)
    add_p(doc, "El portal web y el backend han sido construidos con tecnologías optimizadas para garantizar tiempos de respuesta ultrarrápidos y nula complejidad de mantenimiento:")

    tbl_tech = doc.add_table(rows=1, cols=4)
    style_table(doc, tbl_tech, [1.5, 1.8, 1.6, 1.6],
        ["Componente", "Tecnología / Framework", "Librerías / Dependencias", "Propósito en la Plataforma"],
        [
            ["Servidor Backend", "Node.js v20+ / Express 4.19", "CORS, Crypto, Body-Parser", "Procesamiento de peticiones REST, orquestación de datos y lógica fiscal."],
            ["Seguridad & Cifrado", "Bcrypt 5.1", "Hash con salteado (10 rondas)", "Protección robusta de contraseñas de administrador en almacenamiento."],
            ["Prevención de Ataques", "Rate Limiter en Memoria", "Control de IPs y ventanas temporales", "Mitigación automática contra ataques de fuerza bruta en logins."],
            ["Frontend Web UI", "Vanilla JavaScript ES6+ (SPA)", "TailwindCSS 3.4, FontAwesome 6.5.1", "Interfaz reactiva modular sin frameworks pesados ni empaquetados lentos."],
            ["Motor Fiscal SAT", "XML Parser Nativo (RegEx)", "Anexo 20 SAT / Art. 28 LISR", "Extracción de atributos CFDI 4.0 y cálculo de deducibilidad en milisegundos."],
            ["Base de Datos", "JSON Database Persistente", "data/comprobantes.json", "Resguardo de comprobantes, usuarios y auditorías tolerante a reinicios."]
        ], header_bg="0F4C3A")

    # 3. ACCESO UNIFICADO Y DISCRETO
    add_header(doc, "3. Mecanismo de Autenticación Unificado y Discreto", level=1)
    add_p(doc, "A diferencia de portales tradicionales donde existe una sección evidente de 'Acceso Administrador' que despierta curiosidad no deseada, la versión 3.0 integra un acceso corporativo uniforme:")
    add_bullet(doc, "La pantalla de bienvenida muestra una única tarjeta de acceso corporativo ('Portal de Comprobación y Control de Gastos') orientada a todo el personal de la empresa.", "Presentación Única:")
    add_bullet(doc, "Dentro del menú desplegable de colaboradores aparece de forma natural el perfil 'Damian Spears — Dirección General', conviviendo con los comisionados técnicos y gerentes.", "Integración en el Directorio:")
    add_bullet(doc, "Si se elige a un colaborador comisionado (ej. Julián López, Sandra Urquilla), el sistema valida su clave única alfanumérica y desbloquea el módulo de 'Comprobación de Viáticos'.", "Flujo para Colaboradores:")
    add_bullet(doc, "Si se elige a 'Damian Spears' y se ingresa la contraseña maestra ('Papaleta723!'), el portal detecta los privilegios de administración, autentica con Bcrypt y abre de inmediato el 'Módulo de Administración de Usuarios y Auditoría Fiscal'.", "Flujo para Administrador:")
    add_bullet(doc, "El botón de inicio de sesión de admin fue retirado de la barra de navegación superior, garantizando confidencialidad absoluta.", "Ocultamiento en Navbar:")

    # 4. MÓDULO DE COMPROBACIÓN Y SALDOS
    add_header(doc, "4. Módulo de Comprobación de Viáticos (Vista Colaborador)", level=1)
    add_p(doc, "Cuando un comisionado accede con su clave, dispone de una suite visual completa para el seguimiento de sus viáticos:")
    add_bullet(doc, "Muestra el anticipo otorgado, el total gastado, el monto deducible validado por el SAT y el saldo pendiente de reintegro o reembolso.", "1. Tarjetas Financieras Dinámicas:")
    add_bullet(doc, "Permite ocultar o mostrar cifras mediante asteriscos (••••••) con un solo clic, garantizando privacidad en espacios públicos.", "2. Modo de Privacidad Multinivel:")
    add_bullet(doc, "Zona de arrastre interactiva que valida archivos XML CFDI 4.0 contra estructura SAT, RFC receptor, temporalidad y topes de gasto.", "3. Validador Drag & Drop de XML:")
    add_bullet(doc, "Bitácora con fecha, emisor, UUID fiscal, categoría y desglose deducible/no deducible de cada factura registrada.", "4. Historial Detallado de Facturas:")

    # 5. MÓDULO ADMINISTRADOR Y AUDITORÍA
    add_header(doc, "5. Módulo de Administración de Personal y Auditoría Fiscal", level=1)
    add_p(doc, "El panel de control administrativo ofrece supervisión integral de toda la nómina corporativa:")
    add_bullet(doc, "Permite registrar colaboradores capturando 12 parámetros obligatorios: Nombre, RFC, Email, Puesto, Departamento, Clave Única, Estatus, Destino, Motivo, Fechas Inicio/Fin y Anticipo.", "Alta de Colaboradores:")
    add_bullet(doc, "Botón '✏️ Editar' en cada tarjeta de usuario para modificar información operativa, claves o presupuesto en vivo sin tocar la base de datos manualmente.", "Edición en Caliente:")
    add_bullet(doc, "Al presionar '🔍 Movimientos', el auditor visualiza las facturas de ese empleado con opción de Aceptar (✓) o Rechazar (✕). Al rechazar, se exige un motivo de auditoría que se guarda y anula el importe deducible a $0.00.", "Auditoría con Motivo de Rechazo:")

    # 6. ESPECIFICACIÓN DE ENDPOINTS
    add_header(doc, "6. Catálogo Oficial de Endpoints API REST", level=1)
    add_p(doc, "El backend expone 12 servicios REST documentados a continuación:")

    tbl_api = doc.add_table(rows=1, cols=4)
    style_table(doc, tbl_api, [0.8, 2.3, 2.1, 1.3],
        ["Método", "Endpoint", "Descripción / Operación", "Seguridad"],
        [
            ["GET", "/api/empleados", "Obtiene la lista de colaboradores con montos enmascarados.", "Público"],
            ["POST", "/api/empleados/:id/verificar-clave", "Valida la clave única alfanumérica del colaborador.", "Público"],
            ["GET", "/api/empleados/:id/viaje", "Obtiene el viaje asignado, saldos y facturas del empleado.", "Público"],
            ["POST", "/api/empleados/:id/validar-comprobante", "Valida y registra XML CFDI 4.0 contra SAT y Art. 28 LISR.", "Público"],
            ["POST", "/api/admin/login", "Autenticación maestra con Bcrypt y generación de token de sesión.", "Rate Limited"],
            ["GET", "/api/admin/verificar-sesion", "Verifica si el token Bearer del administrador sigue vigente.", "Bearer Token"],
            ["POST", "/api/admin/logout", "Revoca la sesión administrativa activa en el servidor.", "Bearer Token"],
            ["GET", "/api/admin/usuarios", "Lista completa de colaboradores con claves y viajes visibles.", "Bearer Token"],
            ["POST", "/api/admin/usuarios", "Alta de nuevo colaborador con 12 campos completos.", "Bearer Token"],
            ["PUT", "/api/admin/usuarios/:id", "Actualización en vivo de datos y viáticos de un colaborador.", "Bearer Token"],
            ["PUT", "/api/admin/empleados/:id/comprobantes/:uuid/estatus", "Dictamen de factura (Aprobado / Rechazado) con motivo.", "Bearer Token"],
            ["GET", "/ViaticosApp.apk", "Descarga directa del instalador de la app móvil Android.", "Público"]
        ], header_bg="0F4C3A")

    # 7. DESPLIEGUE CLOUD Y GITHUB
    add_header(doc, "7. Despliegue en la Nube y Control de Versiones", level=1)
    add_callout(doc, "🌐 Infraestructura de Producción 24/7 en Render Cloud",
        "• URL Oficial del Portal: https://jlg-facturacion.onrender.com/\n"
        "• Repositorio en GitHub: https://github.com/jorgecesba/jlg-facturacion.git\n"
        "• Integración Continua (CI/CD): Todo cambio confirmado a la rama 'main' activa un build y despliegue automático en la nube.\n"
        "• Botón de Descarga APK: La cabecera del portal incluye el botón '📱 Descargar App Android' para instalar la app en cualquier celular.",
        bg_hex="EFF6FF", border_hex="3B82F6")

    doc.save(output_path)
    print(f"[OK] Documento Web guardado en: {output_path}")


# =============================================================================
# GENERADOR 2: DOCUMENTACIÓN APLICACIÓN MÓVIL ANDROID
# =============================================================================
def generar_doc_android(output_path):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # PORTADA
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(10)
    p_title.paragraph_format.space_after = Pt(2)
    run_org = p_title.add_run("GRUPO MERCANTIL EMPRESARIAL S.A. DE C.V.\n")
    run_org.font.size = Pt(11)
    run_org.font.bold = True
    run_org.font.color.rgb = RGBColor(100, 116, 139)

    run_main = p_title.add_run("JLG FACTURACIÓN | MANUAL TÉCNICO DE LA APP ANDROID NATIVA\n")
    run_main.font.size = Pt(17)
    run_main.font.bold = True
    run_main.font.color.rgb = RGBColor(30, 64, 175) # Blue theme for Android

    run_sub = p_title.add_run("Arquitectura en Kotlin, Generador CFDI 4.0 al Vuelo, Muestras Rápidas 1-Tap, Módulo Admin y Conexión Cloud")
    run_sub.font.size = Pt(11.5)
    run_sub.font.color.rgb = RGBColor(15, 76, 58)

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(6)
    p_meta.paragraph_format.space_after = Pt(14)
    run_meta = p_meta.add_run("Versión de la App: 3.0.0 (Build 1)  |  Target SDK: Android 14 (API 34)  |  Paquete APK: ViaticosApp.apk (7.68 MB)")
    run_meta.font.size = Pt(9)
    run_meta.font.italic = True
    run_meta.font.color.rgb = RGBColor(100, 116, 139)

    # 1. RESUMEN EJECUTIVO DE LA APP
    add_header(doc, "1. Resumen Ejecutivo de la Aplicación Móvil Android", level=1, color_rgb=RGBColor(30, 64, 175))
    add_p(doc, "ViaticosAndroidApp es la aplicación móvil nativa desarrollada en Kotlin para Android, diseñada para acompañar a los colaboradores comisionados en viajes de negocios y permitir a los auditores supervisar comprobaciones en tiempo real desde sus teléfonos celulares. La app elimina la necesidad de contar con una computadora portátil para validar gastos fiscales y ofrece una experiencia táctil intuitiva basada en Material Design 3.")

    add_callout(doc, "Novedades y Capacidades Principales en la Versión 3.0.0",
        "• Generador Directo de Facturas CFDI 4.0: Permite comprobar gastos creando la estructura fiscal digital al instante sin requerir un archivo XML previo en el almacenamiento.\n"
        "• Muestras Rápidas en 1 Toque (1-Tap Testing): 3 botones preconfigurados para validar facturas de Alimentos ($720), Hospedaje ($986) y Excedente fiscal ($1,300) en segundos.\n"
        "• Inicio de Sesión Segmentado: Pestañas para 'Colaborador Comisionado' (búsqueda y clave única) y 'Administrador' (usuario maestro y contraseña directa).\n"
        "• Módulo de Administración Completo: Alta de colaboradores con 12 campos, edición en vivo (✏️ Editar) y rechazo con justificación contable obligatoria.",
        bg_hex="EFF6FF", border_hex="3B82F6", title_rgb=RGBColor(30, 64, 175))

    # 2. STACK TECNOLÓGICO Y COMPILACIÓN
    add_header(doc, "2. Arquitectura de Software y Stack Tecnológico Android", level=1, color_rgb=RGBColor(30, 64, 175))
    add_p(doc, "La aplicación utiliza las mejores prácticas de la ingeniería de software para Android:")

    tbl_tech_app = doc.add_table(rows=1, cols=4)
    style_table(doc, tbl_tech_app, [1.5, 1.8, 1.6, 1.6],
        ["Componente", "Tecnología / Librería", "Versión / Configuración", "Función en la Aplicación"],
        [
            ["Lenguaje", "Kotlin Nativo", "v1.9.24", "Desarrollo tipado, conciso y libre de NullPointerExceptions."],
            ["Diseño & UI", "Google Material Design 3", "com.google.android.material:1.12.0", "Tarjetas, chips dinámicos, tabs y diálogos con scroll."],
            ["Vistas & Binding", "Android Jetpack ViewBinding", "viewBinding = true", "Acceso type-safe y eficiente a vistas XML sin findViewById."],
            ["Concurrencia", "Kotlin Coroutines", "lifecycleScope / CoroutineScope", "Ejecución asíncrona de peticiones de red sin bloquear la UI."],
            ["Capa de Red REST", "Retrofit 2 + Gson Converter", "v2.11.0 / com.squareup.okhttp3:4.12.0", "Consumo de endpoints REST con deserialización automática."],
            ["Compilador & JDK", "Gradle 8.9 / OpenJDK 21", "Android Studio JBR / AGP 8.5.1", "Compilación de bytecode moderna para Android 8.0 hasta Android 15."]
        ], header_bg="1E40AF")

    # 3. PANTALLA DE INICIO Y LOGIN SEGMENTADO
    add_header(doc, "3. Pantalla de Inicio y Login Segmentado", level=1, color_rgb=RGBColor(30, 64, 175))
    add_p(doc, "La pantalla de inicio de la app presenta dos pestañas claramente diferenciadas:")
    add_bullet(doc, "Contiene un buscador táctil de colaboradores conectados al servidor en la nube. Al elegir un usuario, muestra su avatar con iniciales, puesto y destino, solicitando su clave única alfanumérica.", "Pestaña 'Colaborador Comisionado':")
    add_bullet(doc, "Diseñada para supervisores y directores. Presenta campos directos para Usuario Maestro (DamianSpears_Owner) y Contraseña (Papaleta723!), accediendo inmediatamente al panel de control.", "Pestaña 'Administrador':")

    # 4. GENERADOR CFDI 4.0 Y MUESTRAS RÁPIDAS
    add_header(doc, "4. Generador Directo de Facturas CFDI 4.0 y Muestras Rápidas", level=1, color_rgb=RGBColor(30, 64, 175))
    add_p(doc, "Una de las innovaciones más destacadas de la aplicación es su capacidad para timbrar y generar comprobantes digitales al vuelo:")

    add_p(doc, "La app incluye dos modos en la pestaña de comprobación:", "4.1 Modos de Comprobación:")
    add_bullet(doc, "Formulario completo para registrar un gasto de inmediato generando el XML internamente.", "1. ⚡ Generar Factura Directa:")
    add_bullet(doc, "Zona dropzone para cargar archivos XML descargados en la memoria del teléfono.", "2. 📂 Cargar Archivo XML:")

    add_p(doc, "El generador incluye los 8 campos obligatorios requeridos por el SAT:", "4.2 Campos del Generador Directo:")
    add_bullet(doc, "Razón social del comercio proveedor.", "1. Emisor / Proveedor:")
    add_bullet(doc, "RFC fiscal del emisor a 12 o 13 caracteres.", "2. RFC Emisor:")
    add_bullet(doc, "Folio interno del comprobante de consumo.", "3. Folio de Factura:")
    add_bullet(doc, "Descripción detallada del bien o servicio adquirido.", "4. Concepto:")
    add_bullet(doc, "Monto total pagado en Moneda Nacional. La app calcula automáticamente el Subtotal y el IVA al 16%.", "5. Monto Total ($ MXN):")
    add_bullet(doc, "Chips de selección: ALIMENTOS, HOSPEDAJE, TRANSPORTE, GASOLINA u OTROS.", "6. Categoría del Gasto:")
    add_bullet(doc, "Chips oficiales SAT: Tarjeta de Crédito (04), Débito (28), SPEI (03) y Efectivo (01).", "7. Forma de Pago SAT:")

    add_p(doc, "Para pruebas ágiles, la app cuenta con 3 botones de un toque:", "4.3 Muestras Rápidas (1-Tap Testing):")
    add_bullet(doc, "Restaurante El Portón por $720.00 MXN. Pasa todas las reglas y se autoriza como 100% deducible.", "🍽️ Muestra 1 (Alimentos Válida):")
    add_bullet(doc, "Hoteles City Express por $986.00 MXN. Deducible íntegramente conforme a comprobación de estancia.", "🏨 Muestra 2 (Hospedaje Válido):")
    add_bullet(doc, "Steak House Capital por $1,300.00 MXN. Comprueba el tope legal del Art. 28 LISR: autoriza $750.00 deducibles y $550.00 como Excedente No Deducible.", "⚠️ Muestra 3 (Excedente en Alimentos):")

    # 5. MÓDULO DE ADMINISTRACIÓN EN ANDROID
    add_header(doc, "5. Módulo de Administración y Auditoría en la App", level=1, color_rgb=RGBColor(30, 64, 175))
    add_p(doc, "Los administradores pueden realizar todas las operaciones de control directamente desde el smartphone:")
    add_bullet(doc, "Formulario con desplazamiento que permite registrar colaboradores con 12 campos completos.", "Alta de Colaboradores:")
    add_bullet(doc, "Botón '✏️ Editar' en cada tarjeta de usuario para modificar información operativa y presupuesto.", "Edición en Vivo (PUT):")
    add_bullet(doc, "Permite inspeccionar comprobantes de cada empleado y dictaminar con Aceptar (✓) o Rechazar (✕), solicitando obligatoriamente el motivo de auditoría.", "Dictamen con Justificación:")

    # 6. INSTALACIÓN Y CONEXIÓN
    add_header(doc, "6. Instalación del APK y Conectividad Cloud 24/7", level=1, color_rgb=RGBColor(30, 64, 175))
    add_callout(doc, "📦 Paquete de Instalación Listo (ViaticosApp.apk)",
        "• Nombre del Archivo: ViaticosApp.apk (app-debug.apk)\n"
        "• Peso: 7.68 MB\n"
        "• Descarga Directa desde la Web: Disponible en la cabecera del portal https://jlg-facturacion.onrender.com/\n"
        "• Conexión de Fábrica: Apunta automáticamente al backend en la nube en Render.",
        bg_hex="EFF6FF", border_hex="3B82F6", title_rgb=RGBColor(30, 64, 175))

    add_header(doc, "6.1 Pasos para Instalar en Teléfonos Android", level=2)
    add_p(doc, "Abre https://jlg-facturacion.onrender.com en el celular y pulsa el botón verde 'Descargar App Android', o transfiere el archivo ViaticosApp.apk vía WhatsApp o Google Drive.", "Paso 1: Descarga:")
    add_p(doc, "Toca el archivo descargado y confirma la instalación. Si Android muestra una advertencia, activa 'Permitir instalación de apps desconocidas' para el navegador.", "Paso 2: Instalación:")
    add_p(doc, "Abre la aplicación. Esta se conectará automáticamente al servidor en Render vía datos móviles o Wi-Fi.", "Paso 3: Inicio Automático:")
    add_p(doc, "Si deseas conectar la app a una red local de pruebas, toca el botón de engrane ⚙️ en la esquina superior derecha e introduce 'http://192.168.100.52:4000/api/'.", "Paso 4: Servidor Local:")

    doc.save(output_path)
    print(f"[OK] Documento Android guardado en: {output_path}")


# =============================================================================
# EJECUCIÓN PRINCIPAL
# =============================================================================
if __name__ == "__main__":
    base_dir = r"C:\Users\Jorge\Desktop\Damian"
    
    # Rutas para el documento Web
    ruta_web_root = os.path.join(base_dir, "Documentacion_Portal_Web_Viaticos.docx")
    ruta_web_portal = os.path.join(base_dir, "portal-viaticos", "Documentacion_Portal_Web_Viaticos.docx")
    
    # Rutas para el documento Android
    ruta_app_root = os.path.join(base_dir, "Documentacion_App_Android_Viaticos.docx")
    ruta_app_portal = os.path.join(base_dir, "portal-viaticos", "Documentacion_App_Android_Viaticos.docx")
    
    print("Generando Documento 1: Portal Web & Backend API...")
    generar_doc_web(ruta_web_root)
    shutil.copyfile(ruta_web_root, ruta_web_portal)
    
    print("\nGenerando Documento 2: Aplicación Móvil Android...")
    generar_doc_android(ruta_app_root)
    shutil.copyfile(ruta_app_root, ruta_app_portal)
    
    print("\n[OK] Ambos documentos Word generados exitosamente!")
