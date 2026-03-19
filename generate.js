const ExcelJS = require('exceljs');

async function generarExcelV8() {
    const workbook = new ExcelJS.Workbook();
    
    // ==========================================
    // 1. PANEL DE CONTROL (INPUTS)
    // ==========================================
    const sheetInputs = workbook.addWorksheet('1. PANEL DE CONTROL', { properties: { tabColor: { argb: 'FF00B0F0' } } });
    sheetInputs.columns = [
        { header: 'PARÁMETRO DEL SISTEMA', key: 'param', width: 48 },
        { header: 'VALOR', key: 'valor', width: 20 },
        { header: 'UNIDAD', key: 'unidad', width: 15 },
        { header: 'DESCRIPCIÓN TÉCNICA Y REGLAS DE USO', key: 'desc', width: 110 }
    ];

    const headerStyle = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } } };
    const sectionStyle = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } } };
    const inputStyle = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } }, border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } } };

    sheetInputs.getRow(1).eachCell(cell => { cell.style = headerStyle; });
    const addSection = (sheet, title) => { const row = sheet.addRow([title, '', '', '']); row.eachCell(cell => { cell.style = sectionStyle; }); };

    addSection(sheetInputs, 'A. PLANIFICACIÓN DE DEMANDA DEL SISTEMA'); // Fila 2
    sheetInputs.addRow(['Usuarios Totales Estimados (Censo)', 800, 'Usuarios', 'Número total de clientes registrados o volumen de tráfico general esperado en la plataforma.']).getCell(2).style = inputStyle; // B3
    sheetInputs.addRow(['Porcentaje de Concurrencia en Pico', 0.20, 'Porcentaje', 'Fracción del censo que accederá al mismo tiempo (Ej: 0.20 = 20%). Escribir como decimal.']).getCell(2).style = inputStyle; // B4
    sheetInputs.addRow(['Interacciones por Minuto por Usuario', 2, 'Acciones/min', 'Frecuencia media con la que un usuario activo hace clic o recarga datos en la interfaz.']).getCell(2).style = inputStyle; // B5
    sheetInputs.addRow(['Peticiones HTTP por Interacción', 1.5, 'Peticiones', 'Llamadas internas al servidor (APIs, recursos estáticos) generadas por cada interacción del usuario.']).getCell(2).style = inputStyle; // B6
    
    addSection(sheetInputs, 'B. INFRAESTRUCTURA FÍSICA Y ALTA DISPONIBILIDAD'); // Fila 7
    sheetInputs.addRow(['Flota Total de Servidores (Nodos)', 3, 'Servidores', 'Cantidad total de máquinas físicas o instancias virtuales conectadas al balanceador de carga.']).getCell(2).style = inputStyle; // B8
    sheetInputs.addRow(['Servidores de Reserva (Tolerancia a Fallos)', 1, 'Servidores', 'Nodos que deben poder aislarse por fallo sin causar la caída del sistema (Arquitectura N+X).']).getCell(2).style = inputStyle; // B9
    sheetInputs.addRow(['Procesadores Físicos/Virtuales por Servidor', 16, 'vCPUs', 'Número de núcleos de procesamiento (Cores) disponibles por cada máquina individual.']).getCell(2).style = inputStyle; // B10
    sheetInputs.addRow(['Memoria RAM Disponible por Servidor', 9011, 'Megabytes', 'Memoria libre por máquina tras el inicio del sistema operativo base.']).getCell(2).style = inputStyle; // B11
    sheetInputs.addRow(['Ancho de Banda de Red por Servidor', 1000, 'Mbps', 'Límite de transferencia de la interfaz de red (Habitualmente 1000 Mbps o superior en entornos Cloud).']).getCell(2).style = inputStyle; // B12
    sheetInputs.addRow(['Hilos Máximos del Servidor Web (Thread Pool)', 200, 'Hilos', 'Conexiones concurrentes máximas aceptadas por el servidor web (Ej: maxThreads en Tomcat, worker_connections en Nginx).']).getCell(2).style = inputStyle; // B13

    addSection(sheetInputs, 'C. PERFIL DE RENDIMIENTO DEL SOFTWARE (SÍNCRONO)'); // Fila 14
    sheetInputs.addRow(['Tiempo de Procesamiento en Procesador (CPU)', 150, 'Milisegundos', 'Tiempo puro de cómputo requerido por petición, excluyendo latencia de red.']).getCell(2).style = inputStyle; // B15
    sheetInputs.addRow(['Tiempo Total de Respuesta al Usuario', 300, 'Milisegundos', 'Duración total de la petición web. Durante este ciclo, la conexión retiene RAM y un Hilo (Thread) del servidor web.']).getCell(2).style = inputStyle; // B16
    sheetInputs.addRow(['Memoria RAM Estática del Servicio', 500, 'Megabytes', 'Consumo base de memoria de la aplicación en estado de reposo.']).getCell(2).style = inputStyle; // B17
    sheetInputs.addRow(['Memoria RAM Dinámica por Petición', 15, 'Megabytes', 'Asignación temporal de memoria exigida para procesar y mantener el contexto de una sola petición.']).getCell(2).style = inputStyle; // B18
    sheetInputs.addRow(['Tamaño Medio de Petición Entrante (Payload In)', 50, 'Kilobytes', 'Peso promedio de los datos transmitidos por el cliente hacia el servidor.']).getCell(2).style = inputStyle; // B19
    sheetInputs.addRow(['Tamaño Medio de Respuesta Saliente (Payload Out)', 200, 'Kilobytes', 'Peso promedio de los datos devueltos por el servidor hacia el cliente.']).getCell(2).style = inputStyle; // B20
    sheetInputs.addRow(['Perfil Arquitectónico de Contención (1 a 3)', 2, 'ID Perfil', 'Penalización por sincronización de CPU: 1=Arquitectura Reactiva, 2=Arquitectura Estándar multihilo, 3=Monolito de alta contención.']).getCell(2).style = inputStyle; // B21

    addSection(sheetInputs, 'D. CAPA DE PERSISTENCIA Y MEMORIA CACHÉ'); // Fila 22
    sheetInputs.addRow(['Límite GLOBAL de Conexiones a Base de Datos', 100, 'Conexiones', 'Capacidad máxima centralizada del motor SQL (Ej: max_connections en PostgreSQL). Independiente del pool local de cada nodo.']).getCell(2).style = inputStyle; // B23
    sheetInputs.addRow(['Consultas a Base de Datos por Petición', 3, 'Consultas', 'Ratio de transacciones (lectura/escritura) disparadas hacia la Base de Datos por cada petición HTTP.']).getCell(2).style = inputStyle; // B24
    sheetInputs.addRow(['Tiempo de Retención de Conexión a BD', 300, 'Milisegundos', 'Tiempo que el ORM bloquea la conexión durante el ciclo de vida de la petición HTTP.']).getCell(2).style = inputStyle; // B25
    sheetInputs.addRow(['Porcentaje de Aciertos en Memoria Caché', 0.80, 'Porcentaje', 'Fracción de peticiones procesadas por capas en memoria (Redis/Memcached) eludiendo la Base de Datos.']).getCell(2).style = inputStyle; // B26

    addSection(sheetInputs, 'E. PROCESAMIENTO ASÍNCRONO Y BUS DE EVENTOS'); // Fila 27
    sheetInputs.addRow(['Tráfico Derivado a Bus de Eventos', 0.20, 'Porcentaje', 'Peticiones HTTP que delegan el procesamiento publicando eventos en un Message Broker (Kafka, RabbitMQ, SQS).']).getCell(2).style = inputStyle; // B28
    sheetInputs.addRow(['Workers Compartiendo Compute Node', 0, '1=Sí, 0=No', 'Indica si los procesos consumidores (Workers) se ejecutan en los mismos nodos web (1), restando procesador; o en infraestructura dedicada (0).']).getCell(2).style = inputStyle; // B29
    sheetInputs.addRow(['Multiplicador de Esfuerzo Asíncrono', 3, 'Multiplicador', 'Coeficiente relativo que cuantifica el coste computacional de una tarea en segundo plano frente al baseline de una petición HTTP síncrona. En ausencia de profiling, establecer en 1.0.']).getCell(2).style = inputStyle; // B30


    // ==========================================
    // 2. DASHBOARD EJECUTIVO
    // ==========================================
    const sheetDash = workbook.addWorksheet('2. DASHBOARD EJECUTIVO', { properties: { tabColor: { argb: 'FF00B050' } } });
    sheetDash.columns = [{ width: 45 }, { width: 25 }, { width: 65 }];
    sheetDash.getRow(1).eachCell(cell => { cell.style = headerStyle; });

    addSection(sheetDash, '1. DEMANDA PROYECTADA DEL SISTEMA');
    sheetDash.addRow(['Volumen de Tráfico Exigido', { formula: "('1. PANEL DE CONTROL'!B3 * '1. PANEL DE CONTROL'!B4 * '1. PANEL DE CONTROL'!B5 * '1. PANEL DE CONTROL'!B6) / 60" }, 'Peticiones por Segundo (RPS) generadas por los usuarios.']).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };

    addSection(sheetDash, '2. DIAGNÓSTICO DE LÍMITES FÍSICOS DE INFRAESTRUCTURA');
    sheetDash.addRow(['Límite de Procesamiento (Compute CPU)', { formula: "'3. MOTOR MATEMÁTICO (OCULTO)'!B19" }, 'Peticiones por Segundo (RPS)']).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    sheetDash.addRow(['Límite de Memoria RAM o Hilos (Thread Pool)', { formula: "'3. MOTOR MATEMÁTICO (OCULTO)'!B20" }, 'Peticiones por Segundo (RPS)']).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    sheetDash.addRow(['Límite de Ancho de Banda (Tarjeta de Red)', { formula: "'3. MOTOR MATEMÁTICO (OCULTO)'!B21" }, 'Peticiones por Segundo (RPS)']).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    sheetDash.addRow(['Límite de Capa de Persistencia (Base de Datos)', { formula: "'3. MOTOR MATEMÁTICO (OCULTO)'!B22" }, 'Peticiones por Segundo (RPS)']).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    
    sheetDash.addRow(['CUELLO DE BOTELLA CRÍTICO IDENTIFICADO', { formula: 'IF(MIN(B5:B8)=B5, "Límite de Procesador (CPU)", IF(MIN(B5:B8)=B6, "Saturación de RAM o Hilos Web", IF(MIN(B5:B8)=B7, "Saturación de Tarjeta de Red", "Saturación de Base de Datos")))' }, 'Componente arquitectónico limitante.']).getCell(2).style = { font: { bold: true }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } } };

    sheetDash.addConditionalFormatting({ ref: 'B5:B8', rules: [{ type: 'dataBar', cfvo: [{type: 'min'}, {type: 'max'}], color: {argb: 'FF638EC6'} }] });

    addSection(sheetDash, '3. DIMENSIONAMIENTO SEGURO DE OPERACIONES');
    sheetDash.addRow(['Capacidad Teórica Bruta del Clúster', { formula: 'MIN(B5:B8)' }, 'Límite técnico absoluto Peticiones/Seg (RPS)']).getCell(2).style = { font: { bold: true }, numFmt: '#,##0.00' };
    sheetDash.addRow(['Umbral Operativo Seguro (Margen 80% SLO)', { formula: 'B11 * 0.80' }, 'Límite operativo seguro sin degradación de latencia (RPS)']).getCell(2).style = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } }, numFmt: '#,##0.00' };

    addSection(sheetDash, '4. RECOMENDACIONES DE VIABILIDAD EMPRESARIAL');
    sheetDash.addRow(['Tolerancia Máxima de Usuarios Simultáneos', { formula: "INT((B12 * 60) / MAX(1, ('1. PANEL DE CONTROL'!B5 * '1. PANEL DE CONTROL'!B6)))" }, 'Concurrencia activa máxima tolerada en la plataforma.']).getCell(2).style = { font: { bold: true, size: 12 } };
    sheetDash.addRow(['Censo Total Soportado Recomendado', { formula: "INT(B14 / MAX(0.01, '1. PANEL DE CONTROL'!B4))" }, 'Volumen de registros en base de datos soportable.']).getCell(2).style = { font: { bold: true, size: 12 } };

    const veredicto = sheetDash.addRow(['ESTADO DE VIABILIDAD DE LA INFRAESTRUCTURA', { formula: 'IF(B3<=B12, "✅ LA INFRAESTRUCTURA SOPORTARÁ LA DEMANDA PROYECTADA", "❌ ALERTA CRÍTICA: RIESGO DE DEGRADACIÓN O CAÍDA")' }, 'Evaluación del cruce entre Demanda y Umbral Operativo.']);
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
    const sheetEngine = workbook.addWorksheet('3. MOTOR MATEMÁTICO (OCULTO)', { properties: { tabColor: { argb: 'FFFFC000' } } }); 
    sheetEngine.columns = [{ width: 45 }, { width: 20 }, { width: 100 }]; 
    
    addSection(sheetEngine, 'A. CÁLCULOS DE REDUNDANCIA Y KERNEL'); // Fila 1
    sheetEngine.addRow(['Servidores Físicos Activos', { formula: "MAX(1, '1. PANEL DE CONTROL'!B8 - '1. PANEL DE CONTROL'!B9)" }, 'Resta la redundancia (N+1) garantizando un escenario de tolerancia a fallos de hardware.']); // B2
    sheetEngine.addRow(['Memoria RAM OS (15%)', { formula: "'1. PANEL DE CONTROL'!B11 * 0.15" }, 'Reserva del Kernel de Linux (Page Cache). Evita colapsos por I/O intensivo de disco (Thrashing).']); // B3
    sheetEngine.addRow(['Memoria RAM Útil Asignable', { formula: "'1. PANEL DE CONTROL'!B11 - B3" }, 'Memoria física real disponible para procesos de aplicación.']); // B4

    addSection(sheetEngine, 'B. LEY UNIVERSAL DE ESCALABILIDAD (USL)'); // Fila 5
    sheetEngine.addRow(['Coeficiente Alpha (Contención)', { formula: 'IF(\'1. PANEL DE CONTROL\'!B21=1, 0.01, IF(\'1. PANEL DE CONTROL\'!B21=2, 0.05, 0.1))' }, 'Penalización por Locks de recursos (Gunther USL) basada en el Perfil Arquitectónico.']); // B6
    sheetEngine.addRow(['Coeficiente Beta (Coherencia)', { formula: 'IF(\'1. PANEL DE CONTROL\'!B21=1, 0.001, IF(\'1. PANEL DE CONTROL\'!B21=2, 0.005, 0.01))' }, 'Penalización por sincronización multihilo (Gunther USL).']); // B7
    sheetEngine.addRow(['Cores Efectivos Degradados', { formula: "MAX(1, '1. PANEL DE CONTROL'!B10 / (1 + B6*('1. PANEL DE CONTROL'!B10-1) + B7*'1. PANEL DE CONTROL'!B10*('1. PANEL DE CONTROL'!B10-1)))" }, 'Rendimiento efectivo del procesador aplicando modelado no lineal de concurrencia.']); // B8

    addSection(sheetEngine, 'C. MULTIPLICADORES Y LÍMITES NATIVOS'); // Fila 9
    sheetEngine.addRow(['Ancho de Banda Máximo KBps', { formula: "'1. PANEL DE CONTROL'!B12 * 125" }, 'Normalización a Kilobytes por segundo.']); // B10
    sheetEngine.addRow(['Modificador CPU por Workers', { formula: "1 / (1 + ('1. PANEL DE CONTROL'!B28 * '1. PANEL DE CONTROL'!B29 * '1. PANEL DE CONTROL'!B30))" }, 'Media ponderada que deduce la CPU consumida por procesos suscritos al Bus de Eventos.']); // B11
    sheetEngine.addRow(['Concurrencia Máxima por Servidor', { formula: "MIN((B4 - '1. PANEL DE CONTROL'!B17) / MAX(1, '1. PANEL DE CONTROL'!B18), '1. PANEL DE CONTROL'!B13)" }, 'Tope de concurrencia física evaluando disponibilidad de RAM frente a límite de Thread Pool configurado.']); // B12

    addSection(sheetEngine, 'D. CAPACIDAD BRUTA ABSOLUTA (POR SERVIDOR)'); // Fila 13
    sheetEngine.addRow(['RPS Máx Procesador (CPU)', { formula: '((B8 * 1000) / MAX(1, \'1. PANEL DE CONTROL\'!B15)) * B11' }, 'Capacidad teórica calculada sobre Cores Degradados, ajustada por impacto de Workers locales.']); // B14
    sheetEngine.addRow(['RPS Máx Memoria/Hilos Web', { formula: '(B12 * 1000) / MAX(1, \'1. PANEL DE CONTROL\'!B16)' }, 'Conversión de Concurrencia a Peticiones por Segundo utilizando el Tiempo Total de Respuesta.']); // B15
    sheetEngine.addRow(['RPS Máx Tarjeta de Red', { formula: 'B10 / MAX(1, (\'1. PANEL DE CONTROL\'!B19 + \'1. PANEL DE CONTROL\'!B20))' }, 'Límite basado en Payloads In/Out.']); // B16
    sheetEngine.addRow(['RPS Máx Base de Datos', { formula: '((\'1. PANEL DE CONTROL\'!B23 / MAX(1, \'1. PANEL DE CONTROL\'!B24)) * (1000 / MAX(1, \'1. PANEL DE CONTROL\'!B25))) / MAX(0.01, (1 - MIN(0.9999, \'1. PANEL DE CONTROL\'!B26)))' }, 'Capacidad de persistencia evaluando límite de conexiones retenidas por ORM mitigado por Hit Rate de Caché.']); // B17
    
    addSection(sheetEngine, 'E. CAPACIDAD AGREGADA DEL CLÚSTER DE ALTA DISPONIBILIDAD'); // Fila 18
    sheetEngine.addRow(['Clúster CPU (RPS)', { formula: 'B14 * B2' }, 'Multiplicador de CPU x Servidores Físicos Activos.']); // B19
    sheetEngine.addRow(['Clúster RAM/Hilos (RPS)', { formula: 'B15 * B2' }, 'Multiplicador de RAM/Hilos x Servidores Físicos Activos.']); // B20
    sheetEngine.addRow(['Clúster Red (RPS)', { formula: 'B16 * B2' }, 'Multiplicador de Red x Servidores Físicos Activos.']); // B21
    sheetEngine.addRow(['Persistencia BD Global (RPS)', { formula: 'B17' }, 'Capacidad central de Base de Datos. No se multiplica por la capa web.']); // B22

    await workbook.xlsx.writeFile('Capacity_Planning_V8_Final.xlsx');
    console.log('✅ Archivo "Capacity_Planning_V8_Final.xlsx" generado limpio de jerga y sin errores de referencia.');
}

generarExcelV8().catch(console.error);