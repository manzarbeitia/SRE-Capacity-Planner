const ExcelJS = require('exceljs');

// 1. PARSEAR ARGUMENTOS DE LA LÍNEA DE COMANDOS
let fileName = 'Capacity_Planning_Default.xlsx';
let lang = 'en';

process.argv.forEach(arg => {
    if (arg.startsWith('--file_name=')) {
        fileName = arg.split('=')[1];
        if (!fileName.endsWith('.xlsx')) fileName += '.xlsx';
    }
    if (arg.startsWith('--lang=')) {
        const parsedLang = arg.split('=')[1].toLowerCase();
        if (parsedLang === 'es' || parsedLang === 'en') {
            lang = parsedLang;
        } else {
            console.warn(`⚠️ Language '${parsedLang}' not supported. Defaulting to 'en'.`);
        }
    }
});

async function generarExcelMultiidioma() {
    const t = translations[lang]; // Cargar el diccionario del idioma elegido
    const workbook = new ExcelJS.Workbook();
    
    // ==========================================
    // 1. PANEL DE CONTROL (INPUTS)
    // ==========================================
    const sheetInputs = workbook.addWorksheet(t.sheet1, { properties: { tabColor: { argb: 'FF00B0F0' } } });
    sheetInputs.columns = [
        { header: t.s1_h1, key: 'param', width: 48 },
        { header: t.s1_h2, key: 'valor', width: 20 },
        { header: t.s1_h3, key: 'unidad', width: 15 },
        { header: t.s1_h4, key: 'desc', width: 115 }
    ];

    const headerStyle = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } } };
    const sectionStyle = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } } };
    const inputStyle = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } };

    sheetInputs.getRow(1).eachCell(cell => { cell.style = headerStyle; });
    const addSection = (sheet, title) => { const row = sheet.addRow([title, '', '', '']); row.eachCell(cell => { cell.style = sectionStyle; }); };

    // --- BLOQUE A ---
    addSection(sheetInputs, t.s1_secA);
    sheetInputs.addRow(t.s1_a1).getCell(2).style = inputStyle; // B3
    sheetInputs.addRow(t.s1_a2).getCell(2).style = inputStyle; // B4
    sheetInputs.addRow(t.s1_a3).getCell(2).style = inputStyle; // B5
    sheetInputs.addRow(t.s1_a4).getCell(2).style = inputStyle; // B6
    
    // --- BLOQUE B ---
    addSection(sheetInputs, t.s1_secB);
    sheetInputs.addRow(t.s1_b1).getCell(2).style = inputStyle; // B8
    sheetInputs.addRow(t.s1_b2).getCell(2).style = inputStyle; // B9
    sheetInputs.addRow(t.s1_b3).getCell(2).style = inputStyle; // B10
    sheetInputs.addRow(t.s1_b4).getCell(2).style = inputStyle; // B11
    sheetInputs.addRow(t.s1_b5).getCell(2).style = inputStyle; // B12
    sheetInputs.addRow(t.s1_b6).getCell(2).style = inputStyle; // B13

    // --- BLOQUE C ---
    addSection(sheetInputs, t.s1_secC);
    sheetInputs.addRow(t.s1_c1).getCell(2).style = inputStyle; // B15
    sheetInputs.addRow(t.s1_c2).getCell(2).style = inputStyle; // B16
    sheetInputs.addRow(t.s1_c3).getCell(2).style = inputStyle; // B17
    sheetInputs.addRow(t.s1_c4).getCell(2).style = inputStyle; // B18
    sheetInputs.addRow(t.s1_c5).getCell(2).style = inputStyle; // B19
    sheetInputs.addRow(t.s1_c6).getCell(2).style = inputStyle; // B20
    sheetInputs.addRow(t.s1_c7).getCell(2).style = inputStyle; // B21

    // --- BLOQUE D ---
    addSection(sheetInputs, t.s1_secD);
    sheetInputs.addRow(t.s1_d1).getCell(2).style = inputStyle; // B23
    sheetInputs.addRow(t.s1_d2).getCell(2).style = inputStyle; // B24
    sheetInputs.addRow(t.s1_d3).getCell(2).style = inputStyle; // B25
    sheetInputs.addRow(t.s1_d4).getCell(2).style = inputStyle; // B26

    // --- BLOQUE E ---
    addSection(sheetInputs, t.s1_secE);
    sheetInputs.addRow(t.s1_e1).getCell(2).style = inputStyle; // B28
    sheetInputs.addRow(t.s1_e2).getCell(2).style = inputStyle; // B29
    sheetInputs.addRow(t.s1_e3).getCell(2).style = inputStyle; // B30


    // ==========================================
    // 2. DASHBOARD EJECUTIVO
    // ==========================================
    const sheetDash = workbook.addWorksheet(t.sheet2, { properties: { tabColor: { argb: 'FF00B050' } } });
    sheetDash.columns = [{ width: 48 }, { width: 25 }, { width: 65 }];
    sheetDash.getRow(1).eachCell(cell => { cell.style = headerStyle; });

    addSection(sheetDash, t.s2_sec1);
    sheetDash.addRow([t.s2_r1[0], { formula: `('${t.sheet1}'!B3 * '${t.sheet1}'!B4 * '${t.sheet1}'!B5 * '${t.sheet1}'!B6) / 60` }, t.s2_r1[1]]).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };

    addSection(sheetDash, t.s2_sec2);
    sheetDash.addRow([t.s2_r2[0], { formula: `'${t.sheet3}'!B19` }, t.s2_r2[1]]).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    sheetDash.addRow([t.s2_r3[0], { formula: `'${t.sheet3}'!B20` }, t.s2_r3[1]]).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    sheetDash.addRow([t.s2_r4[0], { formula: `'${t.sheet3}'!B21` }, t.s2_r4[1]]).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    sheetDash.addRow([t.s2_r5[0], { formula: `'${t.sheet3}'!B22` }, t.s2_r5[1]]).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    
    sheetDash.addRow([t.s2_r6[0], { formula: `IF(MIN(B5:B8)=B5, "${t.s2_r6[1]}", IF(MIN(B5:B8)=B6, "${t.s2_r6[2]}", IF(MIN(B5:B8)=B7, "${t.s2_r6[3]}", "${t.s2_r6[4]}")))` }, t.s2_r6[5]]).getCell(2).style = { font: { bold: true }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } } };

    sheetDash.addConditionalFormatting({ ref: 'B5:B8', rules: [{ type: 'dataBar', cfvo: [{type: 'min'}, {type: 'max'}], color: {argb: 'FF638EC6'} }] });

    addSection(sheetDash, t.s2_sec3);
    sheetDash.addRow([t.s2_r7[0], { formula: 'MIN(B5:B8)' }, t.s2_r7[1]]).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    sheetDash.addRow([t.s2_r8[0], { formula: 'B11 * 0.80' }, t.s2_r8[1]]).getCell(2).style = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } }, numFmt: '#,##0.00' };

    addSection(sheetDash, t.s2_sec4);
    sheetDash.addRow([t.s2_r9[0], { formula: `INT((B12 * 60) / MAX(1, ('${t.sheet1}'!B5 * '${t.sheet1}'!B6)))` }, t.s2_r9[1]]).getCell(2).style = { font: { bold: true, size: 12 } };
    sheetDash.addRow([t.s2_r10[0], { formula: `INT(B14 / MAX(0.01, '${t.sheet1}'!B4))` }, t.s2_r10[1]]).getCell(2).style = { font: { bold: true, size: 12 } };

    const veredicto = sheetDash.addRow([t.s2_r11[0], { formula: `IF(B3<=B12, "${t.s2_r11[1]}", "${t.s2_r11[2]}")` }, t.s2_r11[3]]);
    veredicto.getCell(2).style = { font: { bold: true, size: 13 } };

    sheetDash.addConditionalFormatting({
        ref: 'B16',
        rules: [
            { type: 'expression', formulae: ['$B$3<=$B$12'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF00B050' } }, font: { bold: true, color: { argb: 'FFFFFFFF' } } } },
            { type: 'expression', formulae: ['$B$3>$B$12'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF0000' } }, font: { bold: true, color: { argb: 'FFFFFFFF' } } } }
        ]
    });


    // ==========================================
    // 3. MOTOR MATEMÁTICO (OCULTO)
    // ==========================================
    const sheetEngine = workbook.addWorksheet(t.sheet3, { properties: { tabColor: { argb: 'FFFFC000' } } }); 
    sheetEngine.columns = [{ width: 45 }, { width: 20 }, { width: 100 }]; 
    
    addSection(sheetEngine, t.s3_secA); 
    sheetEngine.addRow([t.s3_a1[0], { formula: `MAX(1, '${t.sheet1}'!B8 - '${t.sheet1}'!B9)` }, t.s3_a1[1]]); 
    sheetEngine.addRow([t.s3_a2[0], { formula: `'${t.sheet1}'!B11 * 0.15` }, t.s3_a2[1]]); 
    sheetEngine.addRow([t.s3_a3[0], { formula: `'${t.sheet1}'!B11 - B3` }, t.s3_a3[1]]); 

    addSection(sheetEngine, t.s3_secB); 
    sheetEngine.addRow([t.s3_b1[0], { formula: `IF('${t.sheet1}'!B21=1, 0.01, IF('${t.sheet1}'!B21=2, 0.05, 0.1))` }, t.s3_b1[1]]); 
    sheetEngine.addRow([t.s3_b2[0], { formula: `IF('${t.sheet1}'!B21=1, 0.001, IF('${t.sheet1}'!B21=2, 0.005, 0.01))` }, t.s3_b2[1]]); 
    sheetEngine.addRow([t.s3_b3[0], { formula: `MAX(1, '${t.sheet1}'!B10 / (1 + B6*('${t.sheet1}'!B10-1) + B7*'${t.sheet1}'!B10*('${t.sheet1}'!B10-1)))` }, t.s3_b3[1]]); 

    addSection(sheetEngine, t.s3_secC); 
    sheetEngine.addRow([t.s3_c1[0], { formula: `'${t.sheet1}'!B12 * 125` }, t.s3_c1[1]]); 
    sheetEngine.addRow([t.s3_c2[0], { formula: `1 / (1 + ('${t.sheet1}'!B28 * '${t.sheet1}'!B29 * '${t.sheet1}'!B30))` }, t.s3_c2[1]]); 
    sheetEngine.addRow([t.s3_c3[0], { formula: `MIN((B4 - '${t.sheet1}'!B17) / MAX(1, '${t.sheet1}'!B18), '${t.sheet1}'!B13)` }, t.s3_c3[1]]); 

    addSection(sheetEngine, t.s3_secD); 
    sheetEngine.addRow([t.s3_d1[0], { formula: `((B8 * 1000) / MAX(1, '${t.sheet1}'!B15)) * B11` }, t.s3_d1[1]]); 
    sheetEngine.addRow([t.s3_d2[0], { formula: `(B12 * 1000) / MAX(1, '${t.sheet1}'!B16)` }, t.s3_d2[1]]); 
    sheetEngine.addRow([t.s3_d3[0], { formula: `B10 / MAX(1, ('${t.sheet1}'!B19 + '${t.sheet1}'!B20))` }, t.s3_d3[1]]); 
    sheetEngine.addRow([t.s3_d4[0], { formula: `(('${t.sheet1}'!B23 / MAX(1, '${t.sheet1}'!B24)) * (1000 / MAX(1, '${t.sheet1}'!B25))) / MAX(0.01, (1 - MIN(0.9999, '${t.sheet1}'!B26)))` }, t.s3_d4[1]]); 
    
    addSection(sheetEngine, t.s3_secE); 
    sheetEngine.addRow([t.s3_e1[0], { formula: 'B14 * B2' }, t.s3_e1[1]]); 
    sheetEngine.addRow([t.s3_e2[0], { formula: 'B15 * B2' }, t.s3_e2[1]]); 
    sheetEngine.addRow([t.s3_e3[0], { formula: 'B16 * B2' }, t.s3_e3[1]]); 
    sheetEngine.addRow([t.s3_e4[0], { formula: 'B17' }, t.s3_e4[1]]); 

    await workbook.xlsx.writeFile(fileName);
    console.log(`✅ ${lang === 'es' ? 'Archivo generado' : 'File generated'}: ${fileName} (${lang.toUpperCase()})`);
}

// ==========================================
// DICCIONARIO DE TRADUCCIONES (i18n)
// ==========================================
const translations = {
    es: {
        sheet1: '1. PANEL DE CONTROL', sheet2: '2. DASHBOARD EJECUTIVO', sheet3: '3. MOTOR MATEMÁTICO (OCULTO)',
        s1_h1: 'PARÁMETRO DEL SISTEMA', s1_h2: 'VALOR', s1_h3: 'UNIDAD', s1_h4: 'DESCRIPCIÓN TÉCNICA Y REGLAS DE USO',
        s1_secA: 'A. PLANIFICACIÓN DE DEMANDA DEL SISTEMA',
        s1_a1: ['Usuarios Totales Estimados (Censo)', 800, 'Usuarios', 'Número total de clientes registrados o volumen de tráfico general esperado en la plataforma.'],
        s1_a2: ['Porcentaje de Concurrencia en Pico', 0.20, 'Porcentaje', 'Fracción del censo que accederá al mismo tiempo (Ej: 0.20 = 20%). Escribir como decimal.'],
        s1_a3: ['Interacciones por Minuto por Usuario', 2, 'Acciones/min', 'Frecuencia media con la que un usuario activo hace clic o recarga datos en la interfaz.'],
        s1_a4: ['Peticiones HTTP por Interacción', 1.5, 'Peticiones', 'Llamadas internas al servidor (APIs, recursos estáticos) generadas por cada interacción del usuario.'],
        s1_secB: 'B. INFRAESTRUCTURA FÍSICA Y ALTA DISPONIBILIDAD',
        s1_b1: ['Flota Total de Servidores (Nodos)', 3, 'Servidores', 'Cantidad total de máquinas físicas o instancias virtuales conectadas al balanceador de carga.'],
        s1_b2: ['Servidores de Reserva (Tolerancia a Fallos)', 1, 'Servidores', 'Nodos que deben poder aislarse por fallo sin causar la caída del sistema (Arquitectura N+X).'],
        s1_b3: ['Procesadores Físicos/Virtuales por Servidor', 16, 'vCPUs', 'Número de núcleos de procesamiento (Cores) disponibles por cada máquina individual.'],
        s1_b4: ['Memoria RAM Disponible por Servidor', 9011, 'Megabytes', 'Memoria libre por máquina tras el inicio del sistema operativo base.'],
        s1_b5: ['Ancho de Banda de Red por Servidor', 1000, 'Mbps', 'Límite de transferencia de la interfaz de red (Habitualmente 1000 Mbps o superior en entornos Cloud).'],
        s1_b6: ['Hilos Máximos del Servidor Web (Thread Pool)', 200, 'Hilos', 'Conexiones concurrentes máximas aceptadas por el servidor web (Ej: maxThreads en Tomcat, worker_connections en Nginx).'],
        s1_secC: 'C. PERFIL DE RENDIMIENTO DEL SOFTWARE (SÍNCRONO)',
        s1_c1: ['Tiempo de Procesamiento en Procesador (CPU)', 150, 'Milisegundos', 'Tiempo puro de cómputo requerido por petición, excluyendo latencia de red.'],
        s1_c2: ['Tiempo Total de Respuesta al Usuario', 300, 'Milisegundos', 'Duración total de la petición web. Durante este ciclo, la conexión retiene RAM y un Hilo (Thread) del servidor web.'],
        s1_c3: ['Memoria RAM Estática del Servicio', 500, 'Megabytes', 'Consumo base de memoria de la aplicación en estado de reposo.'],
        s1_c4: ['Memoria RAM Dinámica por Petición', 15, 'Megabytes', 'Asignación temporal de memoria exigida para procesar y mantener el contexto de una sola petición.'],
        s1_c5: ['Tamaño Medio de Petición Entrante (Payload In)', 50, 'Kilobytes', 'Peso promedio de los datos transmitidos por el cliente hacia el servidor.'],
        s1_c6: ['Tamaño Medio de Respuesta Saliente (Payload Out)', 200, 'Kilobytes', 'Peso promedio de los datos devueltos por el servidor hacia el cliente.'],
        s1_c7: ['Perfil Arquitectónico de Contención (1 a 3)', 2, 'ID Perfil', 'Penalización matemática por bloqueo de hilos. 1=App Reactiva (Node/Go), 2=Estándar (Java/Python/C#), 3=Monolito (PHP/Ruby).'],
        s1_secD: 'D. CAPA DE PERSISTENCIA Y MEMORIA CACHÉ',
        s1_d1: ['Límite GLOBAL de Conexiones a Base de Datos', 100, 'Conexiones', 'Capacidad máxima centralizada del motor SQL (Ej: max_connections en PostgreSQL). Independiente del pool local de cada nodo.'],
        s1_d2: ['Consultas a Base de Datos por Petición', 3, 'Consultas', 'Ratio de transacciones (lectura/escritura) disparadas hacia la Base de Datos por cada petición HTTP.'],
        s1_d3: ['Tiempo de Retención de Conexión a BD', 300, 'Milisegundos', 'Tiempo que el ORM bloquea la conexión durante el ciclo de vida de la petición HTTP.'],
        s1_d4: ['Porcentaje de Aciertos en Memoria Caché', 0.80, 'Porcentaje', 'Fracción de peticiones procesadas por capas en memoria (Redis/Memcached) eludiendo la Base de Datos.'],
        s1_secE: 'E. PROCESAMIENTO ASÍNCRONO Y BUS DE EVENTOS',
        s1_e1: ['Tráfico Derivado a Bus de Eventos', 0.20, 'Porcentaje', 'Peticiones HTTP que delegan el procesamiento publicando eventos en un Message Broker (Kafka, RabbitMQ, SQS).'],
        s1_e2: ['¿Workers Comparten Servidor Web?', 0, '1=Sí, 0=No', '¿El Worker que lee del Bus está instalado en la misma máquina que la web (1) o en servidores independientes (0)? Si no hay Bus, pon 0.'],
        s1_e3: ['Multiplicador de Esfuerzo Asíncrono', 3, 'Multiplicador', 'Coeficiente relativo que cuantifica el coste computacional de una tarea en segundo plano frente al baseline de una petición HTTP síncrona. En ausencia de profiling, establecer en 1.0.'],
        
        s2_sec1: '1. DEMANDA PROYECTADA DEL SISTEMA',
        s2_r1: ['Volumen de Tráfico Exigido', 'Peticiones por Segundo (RPS) generadas por los usuarios.'],
        s2_sec2: '2. DIAGNÓSTICO DE LÍMITES FÍSICOS DE INFRAESTRUCTURA',
        s2_r2: ['Límite de Procesamiento (Compute CPU)', 'Peticiones por Segundo (RPS)'],
        s2_r3: ['Límite de Memoria RAM o Hilos (Thread Pool)', 'Peticiones por Segundo (RPS)'],
        s2_r4: ['Límite de Ancho de Banda (Tarjeta de Red)', 'Peticiones por Segundo (RPS)'],
        s2_r5: ['Límite de Capa de Persistencia (Base de Datos)', 'Peticiones por Segundo (RPS)'],
        s2_r6: ['CUELLO DE BOTELLA CRÍTICO IDENTIFICADO', 'Límite de Procesador (CPU)', 'Saturación de RAM o Hilos Web', 'Saturación de Tarjeta de Red', 'Saturación de Base de Datos', 'Componente arquitectónico limitante.'],
        s2_sec3: '3. DIMENSIONAMIENTO SEGURO DE OPERACIONES',
        s2_r7: ['Capacidad Teórica Bruta del Clúster', 'Límite técnico absoluto Peticiones/Seg (RPS)'],
        s2_r8: ['Umbral Operativo Seguro (Margen 80% SLO)', 'Límite operativo seguro sin degradación de latencia (RPS)'],
        s2_sec4: '4. RECOMENDACIONES DE VIABILIDAD EMPRESARIAL',
        s2_r9: ['Tolerancia Máxima de Usuarios Simultáneos', 'Concurrencia activa máxima tolerada en la plataforma.'],
        s2_r10: ['Censo Total Soportado Recomendado', 'Volumen de registros en base de datos soportable.'],
        s2_r11: ['ESTADO DE VIABILIDAD DE LA INFRAESTRUCTURA', '✅ LA INFRAESTRUCTURA SOPORTARÁ LA DEMANDA PROYECTADA', '❌ ALERTA CRÍTICA: RIESGO DE DEGRADACIÓN O CAÍDA', 'Evaluación del cruce entre Demanda y Umbral Operativo.'],

        s3_secA: 'A. CÁLCULOS DE REDUNDANCIA Y KERNEL',
        s3_a1: ['Servidores Físicos Activos', 'Resta la redundancia (N+1) garantizando un escenario de tolerancia a fallos de hardware.'],
        s3_a2: ['Memoria RAM OS (15%)', 'Reserva del Kernel de Linux (Page Cache). Evita colapsos por I/O intensivo de disco (Thrashing).'],
        s3_a3: ['Memoria RAM Útil Asignable', 'Memoria física real disponible para procesos de aplicación.'],
        s3_secB: 'B. LEY UNIVERSAL DE ESCALABILIDAD (USL)',
        s3_b1: ['Coeficiente Alpha (Contención)', 'Penalización por Locks de recursos (Gunther USL) basada en el Perfil Arquitectónico.'],
        s3_b2: ['Coeficiente Beta (Coherencia)', 'Penalización por sincronización multihilo (Gunther USL).'],
        s3_b3: ['Cores Efectivos Degradados', 'Rendimiento efectivo del procesador aplicando modelado no lineal de concurrencia.'],
        s3_secC: 'C. MULTIPLICADORES Y LÍMITES NATIVOS',
        s3_c1: ['Ancho de Banda Máximo KBps', 'Normalización a Kilobytes por segundo.'],
        s3_c2: ['Modificador CPU por Workers', 'Media ponderada que deduce la CPU consumida por procesos suscritos al Bus de Eventos.'],
        s3_c3: ['Concurrencia Máxima por Servidor', 'Tope de concurrencia física evaluando disponibilidad de RAM frente a límite de Thread Pool configurado.'],
        s3_secD: 'D. CAPACIDAD BRUTA ABSOLUTA (POR SERVIDOR)',
        s3_d1: ['RPS Máx Procesador (CPU)', 'Capacidad teórica calculada sobre Cores Degradados, ajustada por impacto de Workers locales.'],
        s3_d2: ['RPS Máx Memoria/Hilos Web', 'Conversión de Concurrencia a Peticiones por Segundo utilizando el Tiempo Total de Respuesta.'],
        s3_d3: ['RPS Máx Tarjeta de Red', 'Límite basado en Payloads In/Out.'],
        s3_d4: ['RPS Máx Base de Datos', 'Capacidad de persistencia evaluando límite de conexiones retenidas mitigado por Caché.'],
        s3_secE: 'E. CAPACIDAD AGREGADA DEL CLÚSTER DE ALTA DISPONIBILIDAD',
        s3_e1: ['Clúster CPU (RPS)', 'Multiplicador de CPU x Servidores Físicos Activos.'],
        s3_e2: ['Clúster RAM/Hilos (RPS)', 'Multiplicador de RAM/Hilos x Servidores Físicos Activos.'],
        s3_e3: ['Clúster Red (RPS)', 'Multiplicador de Red x Servidores Físicos Activos.'],
        s3_e4: ['Persistencia BD Global (RPS)', 'Capacidad central de Base de Datos. No se multiplica por la capa web.']
    },
    en: {
        sheet1: '1. CONTROL PANEL', sheet2: '2. EXECUTIVE DASHBOARD', sheet3: '3. MATH ENGINE (HIDDEN)',
        s1_h1: 'SYSTEM PARAMETER', s1_h2: 'VALUE', s1_h3: 'UNIT', s1_h4: 'TECHNICAL DESCRIPTION AND USAGE RULES',
        s1_secA: 'A. SYSTEM DEMAND PLANNING',
        s1_a1: ['Estimated Total Users (Census)', 800, 'Users', 'Total number of registered clients or expected general traffic volume on the platform.'],
        s1_a2: ['Peak Concurrency Percentage', 0.20, 'Percentage', 'Fraction of the census that will access the system simultaneously (e.g., 0.20 = 20%). Enter as decimal.'],
        s1_a3: ['Interactions per Minute per User', 2, 'Actions/min', 'Average frequency at which an active user clicks or reloads data in the interface.'],
        s1_a4: ['HTTP Requests per Interaction', 1.5, 'Requests', 'Internal server calls (APIs, static resources) generated by each user interaction.'],
        s1_secB: 'B. PHYSICAL INFRASTRUCTURE AND HIGH AVAILABILITY',
        s1_b1: ['Total Server Fleet (Nodes)', 3, 'Servers', 'Total amount of physical machines or virtual instances connected to the load balancer.'],
        s1_b2: ['Standby Servers (Fault Tolerance)', 1, 'Servers', 'Nodes that must be able to fail without causing system downtime (N+X Architecture).'],
        s1_b3: ['Physical/Virtual Processors per Server', 16, 'vCPUs', 'Number of processing cores available per individual machine.'],
        s1_b4: ['Available RAM per Server', 9011, 'Megabytes', 'Free memory per machine after base OS boot.'],
        s1_b5: ['Network Bandwidth per Server', 1000, 'Mbps', 'Transfer limit of the network interface (Usually 1000 Mbps or higher in Cloud environments).'],
        s1_b6: ['Max Web Server Threads (Thread Pool)', 200, 'Threads', 'Maximum concurrent connections accepted by the web server (e.g., maxThreads in Tomcat, worker_connections in Nginx).'],
        s1_secC: 'C. SOFTWARE PERFORMANCE PROFILE (SYNCHRONOUS)',
        s1_c1: ['Processor Processing Time (CPU)', 150, 'Milliseconds', 'Pure compute time required per request, excluding network latency.'],
        s1_c2: ['Total User Response Time', 300, 'Milliseconds', 'Total duration of the web request. During this cycle, the connection retains RAM and a web server Thread.'],
        s1_c3: ['Static Service RAM', 500, 'Megabytes', 'Base memory consumption of the application in an idle state.'],
        s1_c4: ['Dynamic RAM per Request', 15, 'Megabytes', 'Temporary memory allocation required to process and maintain the context of a single request.'],
        s1_c5: ['Average Incoming Request Size (Payload In)', 50, 'Kilobytes', 'Average weight of the data transmitted by the client to the server.'],
        s1_c6: ['Average Outgoing Response Size (Payload Out)', 200, 'Kilobytes', 'Average weight of the data returned by the server to the client.'],
        s1_c7: ['Architectural Contention Profile (1 to 3)', 2, 'Profile ID', 'Mathematical penalty for thread blocking. 1=Reactive App (Node/Go), 2=Standard (Java/Python/C#), 3=Monolith (PHP/Ruby).'],
        s1_secD: 'D. PERSISTENCE LAYER AND MEMORY CACHE',
        s1_d1: ['GLOBAL Database Connection Limit', 100, 'Connections', 'Maximum centralized capacity of the SQL engine (e.g., max_connections in PostgreSQL). Independent of each node\'s local pool.'],
        s1_d2: ['Database Queries per Request', 3, 'Queries', 'Ratio of transactions (read/write) triggered to the Database per HTTP request.'],
        s1_d3: ['DB Connection Retention Time', 300, 'Milliseconds', 'Time the ORM locks the connection during the lifecycle of the HTTP request.'],
        s1_d4: ['Cache Hit Rate Percentage', 0.80, 'Percentage', 'Fraction of requests processed by in-memory layers (Redis/Memcached) bypassing the Database.'],
        s1_secE: 'E. ASYNCHRONOUS PROCESSING AND EVENT BUS',
        s1_e1: ['Traffic Routed to Event Bus', 0.20, 'Percentage', 'HTTP requests that delegate processing by publishing events to a Message Broker (Kafka, RabbitMQ, SQS).'],
        s1_e2: ['Do Workers Share Web Server?', 0, '1=Yes, 0=No', 'Is the Worker reading from the Bus installed on the same machine as the web (1) or on independent servers (0)? If no Bus, enter 0.'],
        s1_e3: ['Asynchronous Effort Multiplier', 3, 'Multiplier', 'Relative coefficient quantifying the computational cost of a background task vs the baseline of a synchronous HTTP request. Default to 1.0.'],
        
        s2_sec1: '1. PROJECTED SYSTEM DEMAND',
        s2_r1: ['Required Traffic Volume', 'Requests per Second (RPS) generated by users.'],
        s2_sec2: '2. PHYSICAL INFRASTRUCTURE LIMITS DIAGNOSTIC',
        s2_r2: ['Processing Limit (Compute CPU)', 'Requests per Second (RPS)'],
        s2_r3: ['Memory or Threads Limit (Thread Pool)', 'Requests per Second (RPS)'],
        s2_r4: ['Bandwidth Limit (Network Card)', 'Requests per Second (RPS)'],
        s2_r5: ['Persistence Layer Limit (Database)', 'Requests per Second (RPS)'],
        s2_r6: ['CRITICAL BOTTLENECK IDENTIFIED', 'Processor Limit (CPU)', 'RAM or Web Threads Saturation', 'Network Card Saturation', 'Database Saturation', 'Limiting architectural component.'],
        s2_sec3: '3. SAFE OPERATIONS SIZING',
        s2_r7: ['Theoretical Gross Cluster Capacity', 'Absolute technical limit Requests/Sec (RPS)'],
        s2_r8: ['Safe Operating Threshold (80% SLO Margin)', 'Safe operating limit without latency degradation (RPS)'],
        s2_sec4: '4. BUSINESS VIABILITY RECOMMENDATIONS',
        s2_r9: ['Maximum Simultaneous Users Tolerance', 'Maximum active concurrency tolerated on the platform.'],
        s2_r10: ['Recommended Supported Total Census', 'Maximum bearable volume of database records.'],
        s2_r11: ['INFRASTRUCTURE VIABILITY STATUS', '✅ INFRASTRUCTURE WILL SUPPORT PROJECTED DEMAND', '❌ CRITICAL ALERT: DEGRADATION OR DOWNTIME RISK', 'Evaluation of Demand vs Operating Threshold.'],

        s3_secA: 'A. REDUNDANCY AND KERNEL CALCULATIONS',
        s3_a1: ['Active Physical Servers', 'Subtracts redundancy (N+1) ensuring a hardware fault tolerance scenario.'],
        s3_a2: ['OS RAM Memory (15%)', 'Linux Kernel reservation (Page Cache). Prevents crashes due to intensive disk I/O (Thrashing).'],
        s3_a3: ['Assignable Useful RAM', 'Real physical memory available for application processes.'],
        s3_secB: 'B. UNIVERSAL SCALABILITY LAW (USL)',
        s3_b1: ['Alpha Coefficient (Contention)', 'Resource Lock penalty (Gunther USL) based on Architectural Profile.'],
        s3_b2: ['Beta Coefficient (Coherency)', 'Multithread synchronization penalty (Gunther USL).'],
        s3_b3: ['Degraded Effective Cores', 'Effective processor performance applying non-linear concurrency modeling.'],
        s3_secC: 'C. MULTIPLIERS AND NATIVE LIMITS',
        s3_c1: ['Maximum Bandwidth KBps', 'Normalization to Kilobytes per second.'],
        s3_c2: ['CPU Modifier per Workers', 'Weighted average deducting CPU consumed by processes subscribed to the Event Bus.'],
        s3_c3: ['Maximum Concurrency per Server', 'Physical concurrency cap evaluating RAM availability vs configured Thread Pool limit.'],
        s3_secD: 'D. ABSOLUTE GROSS CAPACITY (PER SERVER)',
        s3_d1: ['Max Processor RPS (CPU)', 'Theoretical capacity calculated on Degraded Cores, adjusted by local Worker impact.'],
        s3_d2: ['Max Memory/Web Threads RPS', 'Conversion of Concurrency to Requests per Second using Total Response Time.'],
        s3_d3: ['Max Network Card RPS', 'Limit based on In/Out Payloads.'],
        s3_d4: ['Max Database RPS', 'Persistence capacity evaluating retained connections limit mitigated by Cache.'],
        s3_secE: 'E. HIGH AVAILABILITY CLUSTER AGGREGATE CAPACITY',
        s3_e1: ['Cluster CPU (RPS)', 'CPU Multiplier x Active Physical Servers.'],
        s3_e2: ['Cluster RAM/Threads (RPS)', 'RAM/Threads Multiplier x Active Physical Servers.'],
        s3_e3: ['Cluster Network (RPS)', 'Network Multiplier x Active Physical Servers.'],
        s3_e4: ['Global DB Persistence (RPS)', 'Central Database Capacity. Not multiplied by the web tier.']
    }
};

generarExcelMultiidioma().catch(console.error);