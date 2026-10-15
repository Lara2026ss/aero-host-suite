/**
 * Aero Host Suite - PUBLIC RELEASE V26-0.2-0.1
 * tools/guide.mjs — Comprehensive Tool Catalog, Discovery Engine & Live Capability Manual
 */
import { getVersion, PRODUCT_NAME } from "../core/version.mjs";
import { RISK_LEVELS } from "../core/permissions.mjs";
import { OFFICE_SHORTCUTS } from "./documents.mjs";

export const DOMAIN_SIGNATURES = {
  documents: {
    status: { signature: "{ app?: 'word'|'excel'|'powerpoint'|'all' }", description: "Verifica estado de procesos Office y enlaces COM en vivo." },
    list_open: { signature: "{ app?: 'word'|'excel'|'powerpoint'|'all' }", description: "Lista todos los documentos y libros actualmente abiertos." },
    get_active: { signature: "{ app?: 'word'|'excel'|'powerpoint' }", description: "Metadatos y estado del documento en foco." },
    inspect_structure: { signature: "{ path?: string, app?: 'word'|'excel'|'powerpoint'|'pdf', mode?: 'auto'|'live'|'headless' }", description: "Esquema jerárquico conciso de secciones, títulos, tablas, hojas o diapositivas." },
    create: { signature: "{ app: 'word'|'excel'|'powerpoint', path?: string, template?: string, visible?: boolean, forceNew?: boolean }", description: "Crea un documento en blanco o desde plantilla." },
    open: { signature: "{ path: string, visible?: boolean, readOnly?: boolean }", description: "Abre un documento existente de Word, PowerPoint o Excel." },
    save: { signature: "{ app?: 'word'|'excel'|'powerpoint', path?: string }", description: "Guarda cambios en el documento activo o especificado." },
    save_as: { signature: "{ targetPath: string, path?: string }", description: "Guarda una copia en la ruta de destino (crea directorios padre si no existen)." },
    close: { signature: "{ path?: string, app?: string, saveChanges?: boolean }", description: "Cierra el documento de forma segura." },
    export_pdf: { signature: "{ path: string, outputPath?: string }", description: "Exporta documento de Word, presentación o libro a formato PDF." },
    read: { signature: "{ path?: string, maxChars?: number }", description: "Lee contenido, párrafos y estructura de Word o extrae texto de PDF." },
    insert_text: { signature: "{ text: string, position?: 'end'|'start', alignment?: 'center'|'left'|'right'|'justify', fontSize?: number, fontName?: string, bold?: boolean, italic?: boolean, color?: string, lineSpacing?: number, spaceBefore?: number, spaceAfter?: number, pageBreak?: boolean, style?: string, bullet?: boolean, save?: boolean }", description: "Inserta texto formateado en Word." },
    replace_text: { signature: "{ search: string, replace: string, matchCase?: boolean, matchWholeWord?: boolean, save?: boolean }", description: "Reemplaza texto en Word conservando formato." },
    clear_document: { signature: "{ path?: string, save?: boolean }", description: "Borra el contenido de un documento de Word en vivo." },
    append_section: { signature: "{ heading: string, body?: string, headingLevel?: 1|2|3, headingColor?: string, headingAlignment?: string, bodyColor?: string, bodyAlignment?: string, fontSize?: number, lineSpacing?: number, pageBreak?: boolean, save?: boolean }", description: "Inserta sección estructurada con título H1-H3 y cuerpo." },
    format_text: { signature: "{ search?: string, text?: string, all?: boolean, allOccurrences?: boolean, alignment?: 'center'|'left'|'right'|'justify', fontSize?: number, fontName?: string, bold?: boolean, italic?: boolean, underline?: boolean, color?: string, lineSpacing?: number, spaceBefore?: number, spaceAfter?: number, style?: string, save?: boolean }", description: "Busca y da formato tipográfico a texto en Word." },
    format_paragraphs: { signature: "{ start?: number, end?: number, all?: boolean, alignment?: 'center'|'left'|'right'|'justify', fontSize?: number, fontName?: string, bold?: boolean, italic?: boolean, underline?: boolean, color?: string, lineSpacing?: number, spaceBefore?: number, spaceAfter?: number, style?: string, save?: boolean }", description: "Formatea párrafos por índice o todo el documento." },
    page_setup: { signature: "{ paperSize?: 'letter'|'a4'|'legal', orientation?: 'portrait'|'landscape', margins?: 'normal'|'narrow'|'moderate', topMargin?: number, bottomMargin?: number, leftMargin?: number, rightMargin?: number, save?: boolean }", description: "Configura tamaño de papel, márgenes y orientación en Word." },
    add_page_numbers: { signature: "{ alignment?: 'center'|'right'|'left', firstPage?: boolean, save?: boolean }", description: "Inserta números de página automáticos en pie de página (omite portada con firstPage=false)." },
    insert_page_break: { signature: "{ beforeText?: string, afterText?: string, save?: boolean }", description: "Inserta un salto de página limpio en Word." },
    insert_table: { signature: "{ app?: 'word'|'powerpoint', rows: number, cols: number, data?: any[][], headerBold?: boolean, headerColor?: string, headerBgColor?: string, zebra?: boolean, zebraColor?: string, alignment?: 'center'|'left'|'right', save?: boolean }", description: "Inserta tabla formateada con bordes, cabecera y filas alternadas." },
    word_table_format: { signature: "{ path?: string, tableIndex?: number, headerBg?: string, zebra?: boolean, zebraBg?: string, borderColor?: string, borderWidth?: number, colWidths?: number[], save?: boolean }", description: "Aplica formato avanzado y diseño profesional a tablas de Word." },
    word_comment: { signature: "{ path?: string, searchText: string, commentText: string, author?: string, save?: boolean }", description: "Inserta comentarios de revisión anclados a texto en Word." },
    word_hyperlink: { signature: "{ path?: string, searchText: string, url: string, tooltip?: string, save?: boolean }", description: "Inserta hipervínculos navegables en texto de Word." },
    word_template: { signature: "{ path?: string, preset?: string, theme?: string, title?: string, subtitle?: string, author?: string, sections?: object[], mode?: 'live'|'headless', save?: boolean }", description: "Genera documento profesional con plantillas de diseño y paletas." },
    insert_shape: { signature: "{ app?: 'word'|'powerpoint', shapeType?: string, text?: string, fillColor?: string, textColor?: string, save?: boolean }", description: "Inserta formas vectoriales con texto y color." },
    insert_image: { signature: "{ imagePath: string, width?: number, height?: number, slideIndex?: number, left?: number, top?: number, save?: boolean }", description: "Inserta imagen en Word o PowerPoint." },
    apply_style: { signature: "{ path?: string, search?: string, style: string, save?: boolean }", description: "Aplica estilos nativos ('Heading 1', 'Título 1', 'Title', 'Normal', 'Quote') a texto." },
    insert_list: { signature: "{ path?: string, items: string[], type?: 'bullet'|'number', lineSpacing?: number, spaceAfter?: number, save?: boolean }", description: "Inserta lista con viñetas o numerada formal." },
    highlight_text: { signature: "{ path?: string, search: string, color?: 'yellow'|'brightgreen'|'turquoise'|'pink'|'blue'|'red'|'gray', save?: boolean }", description: "Resalta texto con color de resaltador nativo de Word." },
    insert_callout: { signature: "{ path?: string, text: string, title?: string, borderColor?: string, bgColor?: string, fontSize?: number, save?: boolean }", description: "Inserta recuadro de aviso/destacado con barra lateral de color y fondo suave." },
    insert_divider: { signature: "{ path?: string, color?: string, weight?: number, spaceBefore?: number, spaceAfter?: number, save?: boolean }", description: "Inserta línea divisoria horizontal elegante." },
    insert_header_footer: { signature: "{ path?: string, headerText?: string, footerText?: string, headerAlign?: 'left'|'center'|'right', footerAlign?: 'left'|'center'|'right', differentFirstPage?: boolean, save?: boolean }", description: "Configura encabezados y pies de página en Word." },
    insert_cover_page: { signature: "{ path?: string, institution?: string, school?: string, subject?: string, title: string, subtitle?: string, group?: string, team?: string[], date?: string, primaryColor?: string, save?: boolean }", description: "Genera portada académica institucional con tipografía y diseño jerárquico." },
    set_page_background: { signature: "{ path?: string, color?: string, save?: boolean }", description: "Configura el color de fondo de página en Word." },
    insert_section_break: { signature: "{ path?: string, type?: 'next_page'|'continuous', save?: boolean }", description: "Inserta salto de sección en Word." },
    insert_table_of_contents: { signature: "{ path?: string, pageBreak?: boolean, save?: boolean }", description: "Inserta Tabla de Contenidos nativa de Word basada en Título 1, 2, 3." },
    update_toc: { signature: "{ path?: string, save?: boolean }", description: "Actualiza los campos de índice/tabla de contenidos en Word." },
    insert_footnote: { signature: "{ path?: string, searchText: string, footnoteText: string, save?: boolean }", description: "Añade una nota al pie en el texto encontrado." },
    replace_placeholders: { signature: "{ path?: string, replacements: object, save?: boolean }", description: "Reemplaza múltiples variables en una sola operación." },
    get_document_stats: { signature: "{ path?: string }", description: "Calcula estadísticas: páginas, palabras, caracteres, párrafos y líneas." },
    set_language: { signature: "{ path?: string, languageId?: number, save?: boolean }", description: "Establece el idioma de corrección (2058=Español México, 1034=Español España)." },
    set_view: { signature: "{ path?: string, viewType?: 'print'|'web'|'reading', zoomPercent?: number, showRuler?: boolean }", description: "Configura la vista de edición y zoom en Word." },
    capture_preview: { signature: "{ app?: 'word'|'excel'|'powerpoint', page?: number, outputPath?: string }", description: "Captura de pantalla física en tiempo real de la ventana de la aplicación." },
    screen_capture: { signature: "{ app?: 'word'|'excel'|'powerpoint', page?: number, outputPath?: string }", description: "Alias directo de capture_preview para auditorías visuales." },
    read_slides: { signature: "{ path?: string }", description: "Lee diapositivas y textos de PowerPoint." },
    add_slide: { signature: "{ layout?: string, title?: string, body?: string, save?: boolean }", description: "Añade diapositiva en PowerPoint." },
    set_slide_text: { signature: "{ slideIndex?: number, shapeIndex?: number, text: string, save?: boolean }", description: "Modifica texto de forma en diapositiva." },
    delete_slide: { signature: "{ slideIndex: number, save?: boolean }", description: "Elimina una diapositiva en PowerPoint." },
    ppt_slide_ops: { signature: "{ path?: string, action: 'duplicate'|'move'|'set_size', slideIndex?: number, targetIndex?: number, width?: number, height?: number, orientation?: 'widescreen'|'standard', save?: boolean }", description: "Operaciones avanzadas en diapositivas (duplicar, mover, tamaño 16:9)." },
    ppt_notes: { signature: "{ path?: string, slideIndex?: number, notesText?: string, action?: 'get'|'set', save?: boolean }", description: "Lee o actualiza notas del orador en PowerPoint." },
    ppt_export_slide: { signature: "{ path?: string, slideIndex?: number, outputPath: string, width?: number, height?: number }", description: "Exporta diapositiva a PNG de alta resolución directamente vía COM." },
    ppt_background: { signature: "{ path?: string, slideIndex?: number, allSlides?: boolean, color?: string, save?: boolean }", description: "Configura el color o tema de fondo de diapositivas." },
    ppt_chart: { signature: "{ path?: string, slideIndex?: number, chartType?: 'column'|'line'|'pie'|'bar', title?: string, left?: number, top?: number, width?: number, height?: number, save?: boolean }", description: "Inserta gráfico nativo en diapositiva de PowerPoint." },
    read_range: { signature: "{ path?: string, sheet?: string, range?: string, mode?: 'live'|'headless' }", description: "Lee valores de celdas en Excel en vivo o headless." },
    write_range: { signature: "{ data: any[][], startCell?: string, sheet?: string, borders?: boolean, save?: boolean, mode?: 'live'|'headless' }", description: "Escribe matriz de datos en Excel con fórmulas y bordes." },
    write_cell: { signature: "{ cell?: string, value?: any, formula?: string, bold?: boolean, save?: boolean }", description: "Escribe celda individual en Excel." },
    add_sheet: { signature: "{ name?: string, save?: boolean }", description: "Añade nueva hoja de cálculo en Excel." },
    excel_table: { signature: "{ path?: string, sheet?: string, range: string, name?: string, style?: string, showHeaders?: boolean, showTotals?: boolean, save?: boolean }", description: "Crea tabla estructurada oficial (ListObject) en Excel." },
    excel_chart: { signature: "{ path?: string, sheet?: string, range: string, chartType?: 'column'|'line'|'pie'|'bar', title?: string, left?: number, top?: number, width?: number, height?: number, save?: boolean }", description: "Crea gráfico nativo en Excel vinculado a rango de datos." },
    excel_format_range: { signature: "{ path?: string, sheet?: string, range: string, numberFormat?: string, bold?: boolean, italic?: boolean, fontSize?: number, fontColor?: string, bgColor?: string, align?: string, vAlign?: string, wrapText?: boolean, borderStyle?: string, borderColor?: string, rowHeight?: number, colWidth?: number, merge?: boolean, save?: boolean }", description: "Formatea rangos en Excel: número, porcentaje, moneda, alineación, bordes y rellenos." },
    excel_sheet_ops: { signature: "{ path?: string, sheet?: string, action: 'rename'|'delete'|'hide'|'unhide'|'freeze_panes'|'unfreeze_panes', newName?: string, freezeCell?: string, save?: boolean }", description: "Gestión de hojas: renombrar, ocultar, inmovilizar paneles." },
    excel_calculate: { signature: "{ path?: string, scanErrors?: boolean, save?: boolean }", description: "Fuerza recálculo de fórmulas y detecta errores (#REF!, #DIV/0!)." },
    batch_create: { signature: "{ documents: Array<{ app: string, path: string }> }", description: "Crea varios documentos de forma secuencial." },
    batch_edit: { signature: "{ operations: Array<{ action: string, args: object }>, path?: string }", description: "Ejecuta lote de operaciones atómicas sobre un documento." },
    pdf_inspect: { signature: "{ path: string }", description: "Inspecciona metadatos y páginas de documento PDF." },
    pdf_extract_text: { signature: "{ path: string, startPage?: number, endPage?: number, maxPages?: number }", description: "Extrae texto limpio por página desde un PDF." },
    pdf_merge: { signature: "{ files: string[], outputPath: string }", description: "Fusiona múltiples documentos PDF en uno solo." },
    get_templates: { signature: "{}", description: "Consulta las paletas de diseño y plantillas disponibles." },
    undo: { signature: "{ app?: 'word'|'excel'|'powerpoint', times?: number, path?: string }", description: "Deshace una o más acciones previas en Word, Excel o PowerPoint." },
    redo: { signature: "{ app?: 'word'|'excel'|'powerpoint', times?: number, path?: string }", description: "Rehace una o más acciones deshechas en Word, Excel o PowerPoint." },
    compose_document: { signature: "{ path?: string, title: string, subtitle?: string, theme?: string, coverPage?: boolean, institution?: string, school?: string, subject?: string, team?: string[], date?: string, sections: Array<{ heading: string, body?: string, list?: string[], table?: any[][], callout?: string, divider?: boolean }>, pageNumbers?: boolean, header?: string, footer?: string, save?: boolean }", description: "Composición fluida de documento completo de nivel ejecutivo o académico en una sola llamada." },
    change_case: { signature: "{ path?: string, case?: 'upper'|'lower'|'title'|'sentence', search?: string, paragraph?: number, save?: boolean }", description: "Cambia entre mayúsculas, minúsculas, tipo título o tipo oración en Word." },
    clear_formatting: { signature: "{ path?: string, paragraph?: number, all?: boolean, save?: boolean }", description: "Limpia todo el formato de texto y párrafos volviendo al estilo base." },
    paragraph_indent: { signature: "{ path?: string, action?: 'increase'|'decrease'|'set', paragraph?: number, leftIndent?: number, rightIndent?: number, firstLineIndent?: number, save?: boolean }", description: "Aumenta, disminuye o ajusta sangría exacta en Word." },
    paragraph_shading: { signature: "{ path?: string, paragraph?: number, color?: string, save?: boolean }", description: "Aplica sombreado/color de fondo a párrafos en Word." },
    insert_textbox: { signature: "{ path?: string, text?: string, left?: number, top?: number, width?: number, height?: number, fontSize?: number, fontName?: string, bold?: boolean, color?: string, fillColor?: string, borderColor?: string, borderWidth?: number, save?: boolean }", description: "Inserta un cuadro de texto flotante con relleno, bordes y tipografía personalizada." },
    insert_watermark: { signature: "{ path?: string, text?: string, fontName?: string, fontSize?: number, color?: string, diagonal?: boolean, save?: boolean }", description: "Inserta marca de agua de texto en diagonal u horizontal en todas las páginas." },
    set_columns: { signature: "{ path?: string, columns?: number, spacing?: number, lineBetween?: boolean, section?: number, save?: boolean }", description: "Divide la sección en 1, 2 o 3 columnas con espaciado y línea separadora opcional." },
    track_changes: { signature: "{ path?: string, action?: 'enable'|'disable'|'toggle'|'accept_all'|'reject_all', save?: boolean }", description: "Control de cambios: activa/desactiva revisión o acepta/rechaza todas las revisiones." },
    insert_symbol: { signature: "{ path?: string, symbol?: string, save?: boolean }", description: "Inserta caracteres especiales y símbolos tipográficos." },
    insert_caption: { signature: "{ path?: string, label?: 'Figura'|'Tabla'|'Ecuación'|string, title?: string, position?: 'above'|'below', save?: boolean }", description: "Inserta título/leyenda numerada automática para figuras o tablas." },
    insert_bookmark: { signature: "{ path?: string, name: string, save?: boolean }", description: "Inserta marcador o ancla de navegación en Word." },
    ppt_transition: { signature: "{ path?: string, slideIndex?: number, effect?: 'fade'|'wipe'|'push'|'cut'|string, duration?: number, all?: boolean, save?: boolean }", description: "Aplica efectos de transición y duración a diapositivas en PowerPoint." },
    ppt_align_shapes: { signature: "{ path?: string, slideIndex?: number, align?: 'left'|'center'|'right'|'top'|'middle'|'bottom', save?: boolean }", description: "Alinea y distribuye formas y objetos en una diapositiva de PowerPoint." },
    excel_auto_filter: { signature: "{ path?: string, sheet?: string, range?: string, field?: number, criteria?: string, save?: boolean }", description: "Activa o filtra datos en Excel con autofiltro nativo por columna." },
    excel_data_validation: { signature: "{ path?: string, sheet?: string, range: string, list: string[], title?: string, message?: string, save?: boolean }", description: "Crea listas desplegables y validación de entrada de datos en Excel." },
    excel_sort: { signature: "{ path?: string, sheet?: string, range: string, keyCell?: string, descending?: boolean, hasHeader?: boolean, save?: boolean }", description: "Ordena rangos y tablas de datos en Excel ascendente o descendentemente." },
    get_shortcuts: { signature: "{ app?: 'word'|'excel'|'powerpoint'|'all', category?: string }", description: "Consulta el catálogo maestro de atajos de teclado y Ribbon para Office." }
  },
  terminal: {
    run_command: { signature: "{ command: string, cwd?: string, timeoutMs?: number, maxOutputLines?: number }", description: "Ejecuta comandos de PowerShell o Bash de forma síncrona sin restricciones." },
    run_background: { signature: "{ command: string, cwd?: string }", description: "Lanza procesos en segundo plano devolviendo jobId." },
    list_jobs: { signature: "{}", description: "Lista procesos en segundo plano activos." },
    kill_job: { signature: "{ jobId: string }", description: "Detiene un trabajo en ejecución." },
    run_script: { signature: "{ script: string, language?: 'powershell'|'cmd', args?: string[] }", description: "Guarda y ejecuta scripts temporales en el sistema." }
  },
  files: {
    read_file: { signature: "{ path: string, startLine?: number, endLine?: number }", description: "Lee archivos de texto en UTF-8." },
    write_file: { signature: "{ path: string, content: string, overwrite?: boolean }", description: "Escribe o sobreescribe archivos en disco." },
    append_file: { signature: "{ path: string, content: string }", description: "Añade contenido al final de un archivo." },
    delete_path: { signature: "{ path: string, recursive?: boolean }", description: "Elimina archivos o directorios." },
    move_path: { signature: "{ source: string, destination: string }", description: "Mueve o renombra archivos o carpetas." },
    copy_path: { signature: "{ source: string, destination: string }", description: "Copia archivos o carpetas." },
    list_dir: { signature: "{ path?: string, recursive?: boolean, depth?: number }", description: "Lista contenidos de un directorio." },
    search_files: { signature: "{ path?: string, query: string, recursive?: boolean }", description: "Busca archivos por nombre o extensión." },
    file_info: { signature: "{ path: string }", description: "Obtiene metadatos (tamaño, fechas, permisos)." },
    stash_content: { signature: "{ content: string, label?: string }", description: "Guarda bloques grandes de texto en stash temporal." },
    get_stash: { signature: "{ stashId: string }", description: "Recupera contenido desde el stash." },
    surgical_replace: { signature: "{ path: string, target: string, replacement: string }", description: "Reemplazo quirúrgico de bloques exactos de texto." }
  },
  system: {
    get_info: { signature: "{}", description: "Información del sistema operativo, CPU, RAM y host." },
    get_stats: { signature: "{}", description: "Métricas en tiempo real de uso de CPU y memoria." },
    list_processes: { signature: "{ filter?: string }", description: "Lista procesos en ejecución." },
    kill_process: { signature: "{ pid?: number, name?: string }", description: "Finaliza procesos por PID o nombre." },
    get_clipboard: { signature: "{}", description: "Obtiene el texto del portapapeles de Windows." },
    set_clipboard: { signature: "{ text: string }", description: "Copia texto al portapapeles de Windows." },
    send_notification: { signature: "{ title: string, message: string }", description: "Envía una notificación toast nativa en Windows." },
    rog_ally_status: { signature: "{}", description: "Estado y telemetría de hardware ASUS ROG Ally X." }
  },
  web: {
    fetch_url: { signature: "{ url: string, headers?: object }", description: "Realiza petición HTTP GET a una URL." },
    extract_text: { signature: "{ url: string }", description: "Extrae el texto legible en markdown de una página web." },
    search_web: { signature: "{ query: string, limit?: number }", description: "Búsqueda web rápida de fuentes." },
    download_file: { signature: "{ url: string, targetPath: string }", description: "Descarga un archivo desde la web al disco local." }
  },
  media: {
    play_sound: { signature: "{ path: string, volume?: number }", description: "Reproduce archivos de audio (WAV, MP3)." },
    stop_sound: { signature: "{}", description: "Detiene la reproducción de audio." },
    set_volume: { signature: "{ level: number }", description: "Ajusta el volumen del sistema Windows (0-100)." },
    get_volume: { signature: "{}", description: "Obtiene el nivel actual de volumen del sistema." },
    capture_screen: { signature: "{ outputPath?: string }", description: "Toma captura de la pantalla completa del escritorio." }
  },
  shortcuts: {
    send_hotkey: { signature: "{ keys: string[] }", description: "Simula pulsaciones de teclas combinadas (ej. Ctrl+Shift+Esc)." },
    type_text: { signature: "{ text: string }", description: "Escribe texto emulando pulsaciones de teclado." },
    launch_app: { signature: "{ appName: string, args?: string[] }", description: "Ejecuta una aplicación de Windows instalada." },
    focus_window: { signature: "{ title: string }", description: "Enfoca y trae al frente una ventana por título." },
    minimize_all: { signature: "{}", description: "Minimiza todas las ventanas (Win+D)." }
  },
  upd: {
    check_updates: { signature: "{}", description: "Verifica versiones y actualizaciones del sistema." },
    system_health: { signature: "{}", description: "Diagnóstico completo de salud del entorno." }
  },
  feedback: {
    submit_feedback: { signature: "{ message: string, score?: number }", description: "Registra retroalimentación técnica." },
    get_feedback: { signature: "{}", description: "Lista entradas previas de retroalimentación." }
  },
  guide: {
    get_catalog: { signature: "{ query?: string, filter?: string }", description: "Catálogo completo de todos los dominios y herramientas." },
    get_tool_guide: { signature: "{ tool?: string, query?: string, name?: string }", description: "Guía detallada de acciones, sintaxis y recetas. Soporta 'all' para el manual unificado." },
    get_recipes: { signature: "{ topic?: string }", description: "Recetas de formato y diseño profesional para documentos." },
    get_shortcuts: { signature: "{ app?: 'word'|'excel'|'powerpoint'|'all', category?: string }", description: "Catálogo maestro de atajos de teclado y Ribbon para Office." },
    get_system_status: { signature: "{}", description: "Estado de módulos cargados y salud general." }
  }
};

export const DOCUMENTS_RECIPES = {
  cover_page_and_header: {
    title: "1. Portada profesional y encabezados",
    description: "Inserta portada académica institucional (IPN, CECyT 8, C.B. 1) con formato centrado, tipografía jerárquica y salto de página automático.",
    steps: [
      "1. Usar insert_cover_page con institution, school, subject, title, subtitle, group, team, date.",
      "2. Configurar encabezados con insert_header_footer (differentFirstPage: true omite encabezado/pie en portada).",
      "3. Añadir numeración automática en el pie con add_page_numbers (firstPage: false)."
    ]
  },
  global_formatting_15_margins: {
    title: "2. Formato general de cuerpo, márgenes e interlineado 1.5",
    description: "Configura tamaño Carta, márgenes normales de 2.54cm e interlineado 1.5 en todo el cuerpo del documento.",
    steps: [
      "1. Configurar página: page_setup con paperSize: 'letter', margins: 'normal'.",
      "2. Aplicar fuente e interlineado a todo el texto: format_paragraphs con all: true, fontName: 'Aptos', fontSize: 11, lineSpacing: 1.5, alignment: 'left'.",
      "3. Justificar párrafos de contenido si se requiere: format_paragraphs con start: N, end: M, alignment: 'justify'."
    ]
  },
  heading_styles_for_toc: {
    title: "3. Estilos de Título para esquema e índice",
    description: "Convierte textos simples en títulos reconocidos por Word (Título 1, Título 2) para permitir navegación y tabla de contenido.",
    steps: [
      "1. Títulos de primer nivel: format_text con search: '1. Servicios de Internet', style: 'Título 1', fontSize: 16, bold: true, color: '#6C1D45', spaceBefore: 14, spaceAfter: 6.",
      "2. Subtítulos: format_text con search: '1.1 Navegadores web', style: 'Título 2', fontSize: 13, bold: true, color: '#003366', spaceBefore: 10, spaceAfter: 4.",
      "3. Insertar índice automático: insert_table_of_contents tras la portada."
    ]
  },
  callouts_and_tables: {
    title: "4. Cajas de aviso destacadas y tablas modernas",
    description: "Añade recuadros visuales con insert_callout y tablas con filas alternadas (zebra) y cabecera con color de fondo.",
    steps: [
      "1. Cajas de aviso: insert_callout con title: 'Definición Clave', text: '...', borderColor: '#6C1D45', bgColor: '#F8FAFC'.",
      "2. Tablas profesionales: insert_table con rows, cols, data, headerBold: true, headerBgColor: '#6C1D45', zebra: true."
    ]
  },
  visual_audit_screen_capture: {
    title: "5. Auditoría visual por página con screen_capture / capture_preview",
    description: "Permite tomar capturas de pantalla de la ventana de Word en páginas específicas para verificar el resultado visual.",
    steps: [
      "1. Capturar portada (página 1): capture_preview con app: 'word', page: 1, outputPath: 'C:/temp/pagina1.png'.",
      "2. Capturar página 2: capture_preview con app: 'word', page: 2, outputPath: 'C:/temp/pagina2.png'.",
      "3. Inspeccionar el archivo PNG generado para validar que no haya desbordes ni texto cortado."
    ]
  },
  fluent_document_composition: {
    title: "6. Composición fluida de documentos con compose_document",
    description: "Permite al agente redactar un documento completo con portada, índice, secciones, listas, tablas y recuadros de aviso en una única llamada.",
    steps: [
      "1. Invocar compose_document con title, theme ('modern_blue' o 'executive_burgundy'), coverPage: true.",
      "2. Proporcionar el array de sections con heading, body, list, table y callout.",
      "3. Activar pageNumbers: true y header/footer si se desea para obtener un documento listo para entrega sin pasos intermedios."
    ]
  },
  ribbon_quick_actions: {
    title: "7. Acciones rápidas de cinta (Ribbon) y atajos de teclado",
    description: "Deshacer/rehacer histórico, marcas de agua, columnas de periódico, ordenamiento en Excel y transiciones de diapositivas.",
    steps: [
      "1. Deshacer/Rehacer: undo / redo con app: 'word' y times: 1.",
      "2. Columnas y marcas de agua: set_columns con columns: 2; insert_watermark con text: 'BORRADOR'.",
      "3. Tablas y filtros Excel: excel_auto_filter con range: 'A1:D50', excel_sort con keyCell: 'A2'.",
      "4. Consultar atajos: get_shortcuts con app: 'word' para obtener todos los atajos de teclado oficiales."
    ]
  }
};

export function createGuideDomain({ runtime }) {
  return {
    name: "guide",
    description: "Catálogo exhaustivo de herramientas, documentación técnica de acciones y recetas de formato para Aero Host Suite.",
    permissions: {
      get_catalog: RISK_LEVELS.LOW,
      get_tool_guide: RISK_LEVELS.LOW,
      get_recipes: RISK_LEVELS.LOW,
      get_system_status: RISK_LEVELS.LOW,
      get_shortcuts: RISK_LEVELS.LOW
    },
    actions: {
      async get_catalog(args = {}) {
        const filter = (args.filter || args.query || "").trim().toLowerCase();
        const reg = runtime?.registry;
        const catalog = [];

        if (reg) {
          for (const [name, mod] of reg.modules.entries()) {
            if (filter && filter !== "all" && !name.toLowerCase().includes(filter) && !mod.description.toLowerCase().includes(filter)) {
              continue;
            }
            const actionNames = Object.keys(mod.actions || {}).sort();
            catalog.push({
              tool: name,
              description: mod.description,
              actionsCount: actionNames.length,
              actions: actionNames.map((act) => ({
                action: act,
                risk: "low",
                signature: DOMAIN_SIGNATURES[name]?.[act]?.signature || "{}"
              }))
            });
          }
        }

        return {
          ok: true,
          product: PRODUCT_NAME,
          version: getVersion(),
          toolsCount: catalog.length,
          catalog
        };
      },

      async get_tool_guide(args = {}) {
        const target = (args.tool || args.name || args.query || args.domain || "all").trim().toLowerCase();
        const reg = runtime?.registry;

        // If target is "all", return complete encyclopedic guide of every domain and action
        if (target === "all") {
          const allGuides = {};
          if (reg) {
            for (const [name, mod] of reg.modules.entries()) {
              const actionsDetail = {};
              for (const actName of Object.keys(mod.actions || {}).sort()) {
                const sigInfo = DOMAIN_SIGNATURES[name]?.[actName];
                actionsDetail[actName] = {
                  signature: sigInfo?.signature || mod.actionSignatures?.[actName] || "{}",
                  description: sigInfo?.description || `Acción '${actName}' de la herramienta '${name}'.`,
                  risk: "low"
                };
              }
              allGuides[name] = {
                tool: name,
                description: mod.description,
                actionsCount: Object.keys(actionsDetail).length,
                actions: actionsDetail
              };
            }
          }

          return {
            ok: true,
            product: PRODUCT_NAME,
            version: getVersion(),
            totalTools: Object.keys(allGuides).length,
            tools: allGuides,
            recipes: DOCUMENTS_RECIPES
          };
        }

        const mod = reg?.modules?.get(target);
        if (!mod) {
          const available = reg ? reg.moduleNames().join(", ") : "";
          throw new Error(`Herramienta '${target}' no encontrada en el registro. Disponibles: ${available} o 'all'`);
        }

        const actionsDetail = {};
        for (const actName of Object.keys(mod.actions || {}).sort()) {
          const sigInfo = DOMAIN_SIGNATURES[target]?.[actName];
          actionsDetail[actName] = {
            signature: sigInfo?.signature || mod.actionSignatures?.[actName] || "{}",
            description: sigInfo?.description || `Acción '${actName}' de la herramienta '${target}'.`,
            risk: "low"
          };
        }

        const res = {
          ok: true,
          tool: target,
          description: mod.description,
          actionsCount: Object.keys(actionsDetail).length,
          actions: actionsDetail
        };

        if (target === "documents") {
          res.recipes = DOCUMENTS_RECIPES;
        }

        return res;
      },

      async get_recipes(args = {}) {
        const topic = (args.topic || args.query || "all").trim().toLowerCase();
        if (topic && topic !== "all" && DOCUMENTS_RECIPES[topic]) {
          return { ok: true, topic, recipe: DOCUMENTS_RECIPES[topic] };
        }
        return {
          ok: true,
          recipesCount: Object.keys(DOCUMENTS_RECIPES).length,
          recipes: DOCUMENTS_RECIPES
        };
      },

      async get_system_status() {
        const reg = runtime?.registry;
        return {
          ok: true,
          product: PRODUCT_NAME,
          version: getVersion(),
          toolsLoaded: reg ? reg.modules.size : 0,
          unrestricted: true,
          securityMode: "open",
          status: "healthy"
        };
      },

      async get_shortcuts(args = {}) {
        const app = (args.app || "all").toLowerCase();
        const category = (args.category || "").toLowerCase();

        if (app !== "all" && OFFICE_SHORTCUTS[app]) {
          const shortcuts = OFFICE_SHORTCUTS[app];
          if (category && shortcuts[category]) {
            return {
              ok: true,
              app,
              category,
              shortcuts: shortcuts[category]
            };
          }
          return {
            ok: true,
            app,
            categories: Object.keys(shortcuts),
            shortcuts
          };
        }

        return {
          ok: true,
          apps: Object.keys(OFFICE_SHORTCUTS),
          shortcuts: OFFICE_SHORTCUTS
        };
      }
    }
  };
}
