import os
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
            if i == 0 or i == 3:
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

# ==========================================
# INICIO DE GENERACIÓN DEL DOCUMENTO
# ==========================================
doc = Document()

# Configuración de márgenes estándar
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

run_main = p_title.add_run("JLG FACTURACIÓN | ECOSISTEMA DE VIÁTICOS Y COMPROBACIONES MÉXICO\n")
run_main.font.size = Pt(17)
run_main.font.bold = True
run_main.font.color.rgb = RGBColor(15, 76, 58)

run_sub = p_title.add_run("Manual Técnico Integral, Portal Web, App Android, Módulo Administrador y Auditoría Fiscal CFDI 4.0")
run_sub.font.size = Pt(11.5)
run_sub.font.color.rgb = RGBColor(30, 64, 175)

p_meta = doc.add_paragraph()
p_meta.paragraph_format.space_before = Pt(6)
p_meta.paragraph_format.space_after = Pt(14)
run_meta = p_meta.add_run("Versión: 2.5.0  |  Normativa: SAT Anexo 20 / CFDI 4.0 / Art. 28 LISR  |  Fecha: Octubre 2026")
run_meta.font.size = Pt(9)
run_meta.font.italic = True
run_meta.font.color.rgb = RGBColor(100, 116, 139)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ==========================================
# 1. RESUMEN EJECUTIVO
# ==========================================
add_header(doc, "1. Resumen Ejecutivo del Ecosistema Jlg Facturación", level=1)
add_p(doc, "Jlg Facturación es una suite empresarial omnicanal compuesta por una plataforma Web full-stack y una aplicación móvil nativa para Android, desarrollada específicamente para resolver el ciclo integral de comprobación de viáticos y representación en México. Su objetivo central es garantizar que el 100% de los gastos deducibles cumplan a cabalidad con los estándares de validación estructural del SAT (CFDI 4.0) y con los topes legales del Artículo 28 Fracción V de la Ley del Impuesto sobre la Renta (LISR).")

add_callout(doc, "Principales Capacidades y Novedades de la Versión 2.5.0", 
    "• Sincronización Bidireccional Web / Android: Los colaboradores suben sus facturas XML desde su smartphone o la computadora y se reflejan en tiempo real en ambos entornos.\n"
    "• Módulo Administrador con Inspección de Movimientos: El administrador puede auditar el viaje de cada colaborador, revisar facturas individuales y Aceptarlas o Rechazarlas con recálculo dinámico de saldos.\n"
    "• Claves Alfanuméricas y Privacidad Multinivel: Soporte para contraseñas con caracteres alfanuméricos (ej. JALM-2026) y enmascaramiento con asteriscos.\n"
    "• Persistencia de Datos en Disco: Almacenamiento JSON seguro en el servidor para evitar pérdida de registros en reinicios.")

# ==========================================
# 2. TABLA DE USUARIOS Y CREDENCIALES
# ==========================================
add_header(doc, "2. Catálogo Oficial de Usuarios y Credenciales de Acceso", level=1)
add_p(doc, "A continuación se detalla la nómina de colaboradores registrados en el sistema, sus puestos, destinos de viaje asignados y sus claves únicas de acceso confidencial:")

tbl_users = doc.add_table(rows=1, cols=6)
style_table(tbl_users, [1.3, 1.2, 1.1, 1.1, 0.9, 0.9],
    ["Colaborador", "Puesto / Área", "RFC", "Destino", "Anticipo", "Clave Única"],
    [
        ["Julián Alberto López Morales", "Gerente Operaciones", "LOMJ880914HC3", "Monterrey, NL", "$6,000.00", "JALM-2026"],
        ["Sandra Urquilla Delgado", "Directora Comercial", "URDS9204058K2", "Guadalajara, JAL", "$8,500.00", "SUD-2026"],
        ["Ángel Miramontes Castro", "Supervisor Técnico", "MICA851120PX9", "Tijuana, BC", "$4,500.00", "AMC-2026"],
        ["Sofia Cristina Navarro Ruiz", "Coordinadora Proyectos", "NARS900101XYZ", "Querétaro, QRO", "$7,500.00", "SCNR-2026"]
    ])

add_callout(doc, "Credenciales Maestras del Administrador",
    "• Usuario Maestro: DamianSpears_Owner\n"
    "• Contraseña de Acceso: Papaleta723!\n"
    "• Rol: Propietario / Auditor Máximo con facultades para dar de alta colaboradores, auditar movimientos y autorizar o rechazar facturas fiscales.",
    bg_hex="FEF3C7", border_hex="F59E0B")

# ==========================================
# 3. MÓDULO ADMINISTRADOR: INSPECCIÓN Y APROBACIÓN
# ==========================================
add_header(doc, "3. Módulo Administrador: Auditoría Fiscal y Dictamen de Facturas", level=1)
add_p(doc, "Tanto en la aplicación Android como en el portal Web, el Administrador dispone de un panel exclusivo para inspeccionar las comprobaciones de cada perfil y dictaminar sus facturas:")

add_header(doc, "3.1 Inspección Individual de Movimientos", level=2)
add_p(doc, "Al presionar el botón 'Movimientos' (Web) o '🔍 Ver Movimientos y Facturas' (Android), se despliega una ficha financiera detallada que contiene:")
add_p(doc, "Nombre, puesto, departamento, RFC fiscal y clave confidencial asignada.", "1. Perfil del Colaborador:")
add_p(doc, "Destino de viaje autorizado, folio de comisión y período.", "2. Detalles de la Comisión:")
add_p(doc, "Presupuesto inicial entregado al trabajador bajo el concepto de viáticos.", "3. Anticipo Asignado:")
add_p(doc, "Suma total de todas las facturas electrónicas subidas.", "4. Total Gastado Bruto:")
add_p(doc, "Monto validado que cumple estrictamente con el Art. 28 de la LISR.", "5. Deducible SAT:")
add_p(doc, "Diferencia neta entre el anticipo y los gastos deducibles. Si es positivo, se presenta en verde como 'Saldo por Reintegrar a la Empresa'; si es negativo, se presenta como 'Excedente a Reembolsar al Trabajador'.", "6. Saldo de Liquidación:")

add_header(doc, "3.2 Mecanismo de Aceptación y Rechazo de Facturas", level=2)
add_p(doc, "El administrador puede dictaminar individualmente cada comprobante CFDI 4.0 cargado por el trabajador:")
add_p(doc, "Confirma la validez fiscal y administrativa del gasto. La factura se marca como '✅ APROBADA' y su importe deducible se suma a favor del saldo de la liquidación.", "• Acción ACEPTAR (✓):")
add_p(doc, "El administrador marca la factura como '✕ RECHAZADA'. En ese mismo instante, el sistema excluye la factura del cálculo deducible (monto deducible = $0.00), trasladando el importe total a gasto no deducible y aumentando inmediatamente el saldo que el empleado debe devolver a la empresa.", "• Acción RECHAZAR (✕):")
add_p(doc, "Cualquier decisión tomada por el administrador se persiste en disco en comprobantes.json y se notifica al colaborador en su panel en tiempo real.", "• Persistencia e Inmutabilidad:")

# ==========================================
# 4. APLICACIÓN ANDROID NATIVA (JLG FACTURACIÓN)
# ==========================================
add_header(doc, "4. Aplicación Móvil Android (ViaticosAndroidApp)", level=1)
add_p(doc, "La aplicación Android fue desarrollada en Kotlin nativo para brindar el máximo rendimiento, seguridad biométrica y compatibilidad con dispositivos modernos:")

add_header(doc, "4.1 Características de la Aplicación Móvil", level=2)
add_p(doc, "Toda la interfaz visual de inicio solicita primero elegir el perfil de usuario mediante una tarjeta táctil estilizada con gradientes y avatares dinámicos. Ningún saldo o dato sensible es visible antes de validar la clave.", "• Acceso Seguro y Estético:")
add_p(doc, "El campo de contraseña utiliza el teclado alfanumérico completo ('textPassword'), permitiendo ingresar contraseñas seguras con letras, números y símbolos.", "• Teclado Alfanumérico:")
add_p(doc, "Soporta carga directa de archivos .XML del CFDI 4.0 almacenados en el celular o recibidos por correo electrónico o WhatsApp.", "• Carga Local de XML:")
add_p(doc, "Se conecta mediante Retrofit 2 / OkHttp3 al backend Node.js, permitiendo actualizar la IP del servidor dinámicamente desde el botón de configuración.", "• Sincronización REST:")
add_p(doc, "El paquete compilado app-debug.apk (6.6 MB) se encuentra disponible en la carpeta raíz del proyecto bajo el nombre 'JLG_Facturacion.apk'.", "• Instalador Listo:")

tbl_apk = doc.add_table(rows=1, cols=4)
style_table(tbl_apk, [1.6, 1.5, 1.8, 1.6],
    ["Componente", "Tecnología", "Ubicación en Código", "Estado"],
    [
        ["Lenguaje & UI", "Kotlin + Material 3", "ViaticosAndroidApp/app/src", "Compilado exitosamente"],
        ["Capa de Red", "Retrofit 2 + Gson", "network/RetrofitClient.kt", "Endpoints sincronizados"],
        ["Instalador APK", "Android SDK 34 (Android 14)", "c:/Users/Jorge/Desktop/Damian/JLG_Facturacion.apk", "Listo para instalar"],
        ["Dispositivos Soportados", "Android 8.0 a Android 15", "Samsung Galaxy, Xiaomi, Motorola, etc.", "100% Compatible"]
    ])

# ==========================================
# 5. REGLAS FISCALES MEXICANAS (SAT & LISR)
# ==========================================
add_header(doc, "5. Reglas Fiscales del SAT y Límites del Art. 28 de la LISR", level=1)
add_p(doc, "El motor fiscal de Jlg Facturación automatiza el cumplimiento estricto de las leyes fiscales mexicanas aplicables a viáticos y comisiones:")

add_callout(doc, "Topes Cuantitativos de Alimentos (Art. 28 Fracc. V LISR)",
    "• Comisión Nacional: Tope diario deducible de $750.00 MXN por comisionado.\n"
    "• Comisión en el Extranjero: Tope diario deducible de $1,500.00 MXN por comisionado.\n"
    "El sistema agrupa automáticamente por día natural. Si en una misma fecha se suben facturas de alimentos por $1,200.00 MXN, el sistema deduce exactamente $750.00 y los restantes $450.00 los clasifica como Excedente No Deducible.",
    bg_hex="EFF6FF", border_hex="3B82F6")

add_p(doc, "Conforme al Art. 27 Fracc. III de la LISR, el combustible es 100% no deducible si se pagó en efectivo (Forma de Pago 01). Debe liquidarse invariablemente con tarjeta de crédito, débito o monedero electrónico autorizado.", "• Regla de Gasolinas y Combustibles:")
add_p(doc, "El CFDI debe haberse expedido a nombre de Grupo Mercantil Empresarial S.A. de C.V. (RFC GME980512ABC). Facturas a nombre de 'Público en General' (XAXX010101000) o terceros son rechazadas.", "• RFC Receptor Corporativo:")
add_p(doc, "La fecha de timbrado del XML debe ubicarse dentro del intervalo de la comisión de trabajo, con una tolerancia de 24 horas para facturas extemporáneas de hoteles o peajes.", "• Temporalidad de la Comisión:")

# ==========================================
# 6. ESPECIFICACIÓN DE ENDPOINTS API REST
# ==========================================
add_header(doc, "6. Especificación de Endpoints del Backend Node.js", level=1)
add_p(doc, "El servidor opera en el puerto 4000 e implementa los siguientes endpoints protegidos y públicos:")

tbl_api = doc.add_table(rows=1, cols=4)
style_table(tbl_api, [0.8, 2.3, 1.8, 1.6],
    ["Método", "Endpoint", "Propósito", "Autenticación"],
    [
        ["GET", "/api/empleados", "Lista de colaboradores con montos enmascarados", "Público"],
        ["POST", "/api/empleados/:id/verificar-clave", "Valida clave única alfanumérica y desbloquea sesión", "Público"],
        ["GET", "/api/empleados/:id/viaje", "Obtiene liquidación, saldos y comprobantes", "Público"],
        ["POST", "/api/empleados/:id/validar-comprobante", "Parsea y valida XML CFDI 4.0 contra reglas SAT", "Público"],
        ["POST", "/api/admin/login", "Autenticación maestra con Bcrypt (token Bearer)", "Público"],
        ["GET", "/api/admin/usuarios", "Lista completa de usuarios con claves visibles", "Bearer Admin Token"],
        ["PUT", "/api/admin/empleados/:id/comprobantes/:uuid/estatus", "Acepta o Rechaza una factura y recalcula saldos", "Bearer Admin Token"]
    ])

# ==========================================
# 7. GUÍA RÁPIDA DE CONEXIÓN Y USO (CLOUD 24/7)
# ==========================================
add_header(doc, "7. Guía Rápida para Abrir el Portal y Conectar la App", level=1)

add_callout(doc, "Despliegue Cloud en Vivo 24/7 (100% Gratuito)",
    "• URL Oficial en la Nube: https://jlg-facturacion.onrender.com/\n"
    "• Estado: Producción Activa (Live)\n"
    "• Disponibilidad: 24 horas del día, los 365 días del año sin necesidad de encender la computadora del administrador.\n"
    "• Conectividad: Acceso universal desde cualquier red Wi-Fi o datos móviles (4G/5G).",
    bg_hex="F0FDF4", border_hex="10B981")

add_header(doc, "7.1 Cómo Abrir el Portal Web en Cualquier Computadora o Celular", level=2)
add_p(doc, "Abre cualquier navegador web (Chrome, Edge, Safari, Firefox) e ingresa directamente a: https://jlg-facturacion.onrender.com/", "Paso 1: Acceso Directo:")
add_p(doc, "Toca la tarjeta de tu nombre e introduce tu clave confidencial alfanumérica (ej. JALM-2026, SCNR-2026). Tus saldos se revelarán.", "Paso 2: Modo Colaborador:")
add_p(doc, "Haz clic en 'Acceso Administrador' e introduce DamianSpears_Owner / Papaleta723! para acceder a la pestaña de administración, auditar facturas de cada colaborador y Aceptar o Rechazar gastos.", "Paso 3: Modo Administrador:")

add_header(doc, "7.2 Cómo Conectar los Celulares de los Usuarios (App Android)", level=2)
add_p(doc, "Copia el archivo 'JLG_Facturacion.apk' a los celulares de los colaboradores (vía WhatsApp, correo o Google Drive) e instálalo habilitando 'Instalar apps desconocidas'.", "Paso 1: Instalación:")
add_p(doc, "La aplicación ya viene configurada de fábrica con la URL cloud 'https://jlg-facturacion.onrender.com/api/'. En cuanto el colaborador abre la app, esta se conecta automáticamente a la nube.", "Paso 2: Conexión Automática:")
add_p(doc, "Los colaboradores pueden estar en su casa, en la oficina o en carretera con datos móviles 4G/5G. No se requiere estar en la misma red Wi-Fi.", "Paso 3: Conexión Universal:")
add_p(doc, "Si en el futuro se desea apuntar a otro servidor privado o red local, la app cuenta con el botón de engrane ⚙️ en la esquina superior derecha para personalizar la URL del backend.", "Paso 4: Flexibilidad de Servidor:")

# Guardar documento
output_path = r"C:\Users\Jorge\Desktop\Damian\portal-viaticos\Documentacion_Portal_Viaticos_Mexico.docx"
doc.save(output_path)
print(f"Documento Word generado exitosamente en: {output_path}")
