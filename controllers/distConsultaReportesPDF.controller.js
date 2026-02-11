import { request, response } from 'express';
import PDFDocument from 'pdfkit';
import { CalcularAltoAncho, DibujarTablaPDF, TextoMultiFuente } from '../helpers/ActasPDF.js';
import { anioN, autor, emblemaEC, iecmLogoBN } from '../helpers/Constantes.js';
import { ConsultaClaveColonia, ConsultaDelegacion, ConsultaDistrito, ConsultaTipoEleccion, FechaServer, InformacionConstancia } from '../helpers/Consultas.js';
import { dividirArreglo, NumAMes, NumAText } from '../helpers/Funciones.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Proyectos Participantes Dictaminados Favorablemente

export const ProyectosParticipantes = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia, anio } = req.body;
    try {
        const consulta = (await SICOVACC.sequelize.query(`SELECT num_proyecto, UPPER(folio_proy_web) AS folio_proy_web,
        UPPER(STUFF((
            SELECT ', ' + rubro
            FROM (VALUES
                (CASE WHEN rubro1 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Mejoramiento de espacios públicos' ELSE 'Mejoramiento' END ELSE NULL END),
                (CASE WHEN rubro2 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Equipamiento e infraestructura urbana' ELSE 'Mantenimiento' END ELSE NULL END),
                (CASE WHEN rubro3 = 1 THEN 'Obras' ELSE NULL END),
                (CASE WHEN rubro4 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Servicios' ELSE 'Reparaciones en áreas y bienes de uso común' END ELSE NULL END),
                (CASE WHEN rubro5 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Actividades deportivas' ELSE 'Servicios' END ELSE NULL END),
                (CASE WHEN rubro6 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Actividades recreativas' ELSE 'Actividades deportivas' END ELSE NULL END),
                (CASE WHEN rubro7 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Actividades culturales' ELSE 'Actividades recreativas' END ELSE NULL END),
                (CASE WHEN rubro8 = 1 THEN CASE WHEN tipo_rubro = 1 THEN NULL ELSE 'Actividades culturales' END ELSE NULL END)
            ) AS sub(rubro)
            WHERE rubro IS NOT NULL
            FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
        ) AS rubro_general, UPPER(nom_proyecto) AS nom_proyecto
        FROM consulta_prelacion_proyectos
        WHERE estatus = 1 AND anio = ${anio} AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
        ORDER BY num_proyecto`))[0];
        if (!consulta.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const { nombre_colonia } = await ConsultaClaveColonia(clave_colonia);
        const titulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        let datos = [], buffer = [];
        consulta.forEach(proyecto => {
            let X = [];
            Object.keys(proyecto).forEach(key => {
                X.push([{ text: proyecto[key] ? proyecto[key].toString().trim() : '', font: 'Helvetica', fontSize: 12 }]);
            });
            datos.push(X);
        });
        const subProyectos = dividirArreglo(datos, 15);
        const doc = new PDFDocument({ bufferPages: true, autoFirstPage: false, size: 'A3', layout: 'landscape' });
        doc.info.Author = autor;
        for (let i = 0; i < subProyectos.length; i++) {
            doc.addPage();
            doc.image('./resources/iecm.png', 47, 40, {
                fit: [200, 100],
                align: 'center',
                valign: 'center'
            });
            doc.font('Helvetica', 18).text('DIRECCIÓN EJECUTIVA DE ORGANIZACIÓN ELECTORAL Y GEOESTADÍSTICA', 70, 72, { width: 1050, align: 'center' });
            doc.text(titulo, 70, 114, { width: 1050, align: 'center' });
            doc.font('Helvetica-Bold').text('PROYECTOS PARTICIPANTES DICTAMINADOS FAVORABLEMENTE', 70, 155, { width: 1050, align: 'center', underline: true });
            doc.font('Helvetica').text(`Dirección Distrital: ${id_distrito}`, 70, 198, { align: 'left' });
            doc.text('Nombre de la Unidad Territorial:', 70, 198, { width: 1050, align: 'center' }).font('Helvetica-Bold').text(nombre_colonia, 70, 218, { width: 1050, align: 'center' });
            doc.font('Helvetica').text('Clave de la Unidad Territorial:', 70, 262, { width: 1050, align: 'center' }).font('Helvetica-Bold').text(`(${clave_colonia})`, 70, 282, { width: 1050, align: 'center' });
            doc.font('Helvetica', 14).text(`Fecha: ${fecha}`, 70, 198, { width: 1050, align: 'right' }).text(`Hora: ${hora.substring(0, hora.length - 3)}`, 70, 228, { width: 1050, align: 'right' }).text('FORMATO 1', 70, 258, { width: 1050, align: 'right' });
            DibujarTablaPDF(doc, 70, 322, [
                [{ text: 'CLAVE DEL PROYECTO', font: 'Helvetica-Bold', fontSize: 14, background: '#C0C0C0' }],
                [{ text: 'FOLIO DE REGISTRO', font: 'Helvetica-Bold', fontSize: 14, background: '#C0C0C0' }],
                [{ text: 'RUBRO GENERAL', font: 'Helvetica-Bold', fontSize: 14, background: '#C0C0C0' }],
                [{ text: 'NOMBRE DEL PROYECTO', font: 'Helvetica-Bold', fontSize: 14, background: '#C0C0C0' }]
            ], [
                { width: 130, align: 'center' },
                { width: 200, align: 'center' },
                { width: 340, align: 'center' },
                { width: 380, align: 'center' }
            ], subProyectos[i]);
        }
        let { y } = doc;
        y += 15;
        doc.rect(400, y, 340, 35).fillAndStroke('#C0C0C0', '#000').font('Helvetica-Bold', 14).fillColor('#000').text('TOTAL', 400, y + 12, { width: 340, align: 'center' });
        doc.rect(740, y, 380, 35).fillAndStroke('#C0C0C0', '#000').font('Helvetica-Bold', 14).fillColor('#000').text(consulta.length, 740, y + 12, { width: 380, align: 'center' });
        if (subProyectos.length > 1) {
            const paginas = doc.bufferedPageRange().count;
            for (let i = 0; i < paginas; i++) {
                doc.switchToPage(i);
                doc.font('Helvetica', 12).text(`Hoja ${i + 1} de ${paginas}`, 70, doc.page.height - 90, { width: 1050, align: 'right' });
            }
        }
        doc.end();
        doc.on('data', buffer.push.bind(buffer));
        doc.on('end', () => {
            res.json({
                success: true,
                msg: 'Reporte generado correctamente',
                contentType: 'application/pdf',
                reporte: `Reporte_ProyectosParticipantes_${clave_colonia}_${fecha}-${hora}.pdf`,
                buffer: Buffer.concat(buffer)
            });
        });
    } catch (err) {
        console.error(`Error al generar el reporte PDF en ProyectosParticipantes: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Constancia - En Desuso

export const ConstanciaPDF = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia, anio } = req.body;
    try {
        const { fecha, hora } = await FechaServer();
        const { nombre_colonia, nombre_delegacion, domicilio, mesas, ultimaFecha, ultimaHora, coordinador_puesto, coordinador, secretario_puesto, secretario } = await InformacionConstancia(anio, clave_colonia);
        const dia = NumAText(+ultimaFecha.split('/')[0]).toLowerCase(), mes = NumAMes(+ultimaFecha.split('/')[1]).toLowerCase(), horas = NumAText(+ultimaHora.split(':')[0]).toLowerCase(), minutos = NumAText(+ultimaHora.split(':')[1]).toLowerCase();
        const mesasL = NumAText(mesas);
        const consulta = await SICOVACC.sequelize.query(`SELECT num_proyecto, nom_proyecto, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
        FROM consulta_actas_VVS
        WHERE anio = ${anio} AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}' AND num_proyecto IN (SELECT num_proyecto FROM consulta_prelacion_proyectos WHERE estatus = 1 AND anio = ${anio} AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}')
        GROUP BY num_proyecto, nom_proyecto
        ORDER BY num_proyecto`);
        let datos = [], buffer = [];
        consulta[0].forEach(proyecto => {
            let X = [];
            Object.keys(proyecto).forEach((key, index) => {
                if ([0, 1].includes(index))
                    X.push([{ text: proyecto[key].toString(), font: 'Helvetica', fontSize: 10 }]);
                else if ([2, 3, 4].includes(index)) {
                    X.push([{ text: Intl.NumberFormat('es-MX').format(proyecto[key]), font: 'Helvetica', fontSize: 10 }]);
                    if (index === 4)
                        X.push([{ text: NumAText(proyecto[key]), font: 'Helvetica', fontSize: 10 }]);
                }
            });
            datos.push(X);
        });
        const subProyectos = dividirArreglo(datos, 25);
        const doc = new PDFDocument({ bufferPages: true, autoFirstPage: false, size: 'A3', layout: 'portrait' });
        doc.info.Author = autor;
        doc.addPage();
        doc.image('./resources/iecm.png', 25, 40, {
            fit: [200, 100],
            align: 'center',
            valign: 'center'
        });
        doc.font('Helvetica-Bold', 10).rect(50, 150, 360, 20).stroke().text('UNIDAD TERRITORIAL: (Nombre)', 50, 155, { width: 360, align: 'center' });
        doc.rect(410, 150, 80, 20).stroke().text('UT: (Clave)', 410, 155, { width: 80, align: 'center' });
        doc.rect(490, 150, 100, 20).stroke().text('DISTRITO: (Núm)', 490, 155, { width: 100, align: 'center' });
        doc.rect(590, 150, 200, 20).stroke().text('DEMARCACIÓN: (Nombre)', 590, 155, { width: 200, align: 'center' });
        doc.font('Helvetica').rect(50, 170, 360, 20).stroke().text(nombre_colonia, 50, 175, { width: 360, align: 'center' });
        doc.rect(410, 170, 80, 20).stroke().text(clave_colonia, 410, 175, { width: 80, align: 'center' });
        doc.rect(490, 170, 100, 20).stroke().text(id_distrito, 490, 175, { width: 100, align: 'center' });
        doc.rect(590, 170, 200, 20).stroke().text(nombre_delegacion, 590, 175, { width: 200, align: 'center' });
        doc.rect(50, 200, 300, 50).stroke().text('NÚMERO DE MESAS INSTALADAS PARA LA ELECCIÓN EN ESTA UNIDAD TERRITORIAL', 50, 215, { height: 50, width: 300, align: 'center' });
        doc.rect(350, 200, 440, 50).stroke();
        doc.text(mesas, 390, 216, { width: 150, align: 'center' });
        doc.moveTo(390, 226).lineTo(540, 226).stroke();
        doc.font('Helvetica-Bold').text('(Con Número)', 390, 228, { width: 150, align: 'center' });
        doc.font('Helvetica').text(mesasL, 580, 216, { width: 170, align: 'center' });
        doc.moveTo(580, 226).lineTo(750, 226).stroke();
        doc.font('Helvetica-Bold').text('(Con Letra)', 580, 228, { width: 170, align: 'center' });
        doc.font('Helvetica').text(`En la Ciudad de México, a las ${horas} horas ${minutos} minutos del día ${dia} de ${mes} de ${anioN[anio]}, en el domicilio que ocupa la Dirección Distrital ${id_distrito}, sitío en ${domicilio},
        se realizó la validación de los resultados de la Consulta de Presupuesto Participativo ${anio} de la Unidad Territorial referida en la presente acta, de conformidad con lo dispuesto en los articulos 113,
        fracciones V y XIV, y 367 del Código de Instituciones y Procedimientos Electorales de la Ciudad de México; 116, 117, 118, 119 y 120 de la Ley de Participación Ciudadana de la Ciudad de México; así como el
        numeral 18 de las disposiciones comunes de la Convocatoria única aprobada por el Consejo General del Instituto Electoral de la Ciudad de México mediante Acuerdo IECM/ACU-CG-007/2025 de fecha 15 de enero de 2023,
        y con base en las opiniones emitidas en la Jornada Electiva Única, de forma remota a través del Sistema Electrónico por Internet y de forma presencial en las mesas receptoras de votación y opinión, haciendo constar lo siguiente:`.replace(/\n/g, ''), 50, 260, { width: 740, align: 'justify' });
        for (let i = 0; i < subProyectos.length; i++) {
            doc.font('Helvetica-Bold', 16).text('RESULTADOS', 50, doc.y + 5, { width: 740, align: 'center' });
            DibujarTablaPDF(doc, 50, doc.y + 5, [
                [{ text: 'NO. DE PROYECTO', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }],
                [{ text: 'NOMBRE DEL PROYECTO', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }],
                [{ text: 'RESULTADOS DEL ESCRUTINIO Y CÓMPUTO DE LA MESA', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }],
                [{ text: 'RESULTADO DEL CÓMPUTO DEL SISTEMA ELECTRÓNICO POR INTERNET', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }],
                [{ text: 'TOTAL CON NÚMERO', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }],
                [{ text: 'TOTAL CON LETRA', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }],
            ], [
                { width: 70, align: 'center' },
                { width: 280, align: 'center' },
                { width: 110, align: 'center' },
                { width: 110, align: 'center' },
                { width: 70, align: 'center' },
                { width: 100, align: 'center' },
            ], subProyectos[i]);
            if (i < subProyectos.length - 1)
                doc.addPage();
        }
        const y = doc.page.height - 250;
        doc.font('Helvetica-Bold', 12).text(`Ciudad de México a ${fecha.split('/')[0]} de ${mes} de 2025`, 50, y, { width: 740, align: 'center' });
        doc.moveDown().text('POR LA DIRECCIÓN DISTRITAL', { width: 740, align: 'center' });
        doc.font('Helvetica', 10).text(coordinador_puesto, 50, y + 65, { width: 365, align: 'center' });
        doc.text(coordinador, 50, y + 105, { width: 250, align: 'center' });
        doc.moveTo(50, y + 115).lineTo(300, y + 115).stroke();
        doc.font('Helvetica-Bold', 8).text('NOMBRE COMPLETO', 50, y + 120, { width: 250, align: 'center' });
        doc.moveTo(310, y + 115).lineTo(415, y + 115).stroke();
        doc.font('Helvetica-Bold', 8).text('FIRMA', 310, y + 120, { width: 105, align: 'center' });
        doc.font('Helvetica', 10).text(secretario_puesto, 425, y + 65, { width: 370, align: 'center' });
        doc.text(secretario, 425, y + 105, { width: 250, align: 'center' });
        doc.moveTo(425, y + 115).lineTo(675, y + 115).stroke();
        doc.font('Helvetica-Bold', 8).text('NOMBRE COMPLETO', 425, y + 120, { width: 250, align: 'center' });
        doc.moveTo(685, y + 115).lineTo(790, y + 115).stroke();
        doc.font('Helvetica-Bold', 8).text('FIRMA', 685, y + 120, { width: 105, align: 'center' });
        if (subProyectos.length > 1) {
            const paginas = doc.bufferedPageRange().count;
            for (let i = 0; i < paginas; i++) {
                doc.switchToPage(i);
                doc.font('Helvetica', 10).text(`Hoja ${i + 1} de ${paginas}`, 50, y + 160, { width: 740, align: 'center' });
            }
        }
        doc.end();
        doc.on('data', buffer.push.bind(buffer));
        doc.on('end', () => {
            res.json({
                success: true,
                msg: 'Reporte generado correctamente',
                contentType: 'application/pdf',
                reporte: `Constancia_${clave_colonia}_${fecha}-${hora}.pdf`,
                buffer: Buffer.concat(buffer)
            });
        });
    } catch (err) {
        console.error(`Error al generar la constancia PDF: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar la constancia'
        });
    }
}

//? Acta de Validacion

export const ActaValidacionPDF = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia, anio } = req.body;
    try {
        const { fecha, hora } = await FechaServer();
        const { nombre_delegacion } = await ConsultaDelegacion(id_distrito, clave_colonia);
        const { nombre_colonia } = await ConsultaClaveColonia(clave_colonia);
        const { direccion, coordinador, coordinador_puesto, secretario, secretario_puesto } = await ConsultaDistrito(id_distrito);
        const eleccion = await ConsultaTipoEleccion(anio);
        const consulta = (await SICOVACC.sequelize.query(`SELECT secuencial AS num_proyecto, SUM(total_votos) AS total_votos
        FROM consulta_actas_VVS
        WHERE anio = ${anio} AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
        GROUP BY secuencial
        UNION ALL
        SELECT 0 AS num_proyecto, SUM(bol_nulas) AS total_votos
        FROM consulta_actas
        WHERE anio = ${anio} AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
        ORDER BY secuencial`))[0];
        const { total_votos: bol_nulas } = consulta.find(proyecto => proyecto.num_proyecto == 0);
        const total = consulta.reduce((sum, proyecto) => sum + proyecto.total_votos, 0);
        let datos = [], buffer = [];
        consulta.filter(proyecto => proyecto.num_proyecto != 0).forEach(proyecto => {
            let X = [];
            Object.keys(proyecto).forEach((key, index) => {
                X.push([{ text: proyecto[key].toString(), font: 'Helvetica', fontSize: 10 }]);
                if (index == 1)
                    X.push([{ text: NumAText(proyecto[key]), font: 'Helvetica', fontSize: 10 }]);
            });
            datos.push(X);
        });
        datos.push([
            [{ text: 'Opiniones nulas', font: 'Helvetica-Bold', fontSize: 10, background: '#C0C0C0' }],
            [{ text: String(bol_nulas), font: 'Helvetica-Bold', fontSize: 10, background: '#C0C0C0' }],
            [{ text: NumAText(bol_nulas), font: 'Helvetica-Bold', fontSize: 10, background: '#C0C0C0' }]
        ], [
            [{ text: 'TOTAL', font: 'Helvetica-Bold', fontSize: 14, background: '#000', fillColor: '#FFF' }],
            [{ text: String(total), font: 'Helvetica-Bold', fontSize: 10, background: '#C0C0C0' }],
            [{ text: NumAText(total), font: 'Helvetica-Bold', fontSize: 10, background: '#C0C0C0' }]
        ]);
        const doc = new PDFDocument({ bufferPages: true, autoFirstPage: false, size: 'A3', layout: 'portrait', margin: 30 });
        doc.info.Author = autor;
        doc.addPage();
        doc.rect(50, 50, 658, 70).fillAndStroke('#000', '#000').font('Helvetica-Bold', 16).fillColor('#FFF').text(`ACTA DE VALIDACIÓN DE RESULTADOS PARA LA CONSULTA DE ${eleccion.toUpperCase()} POR UNIDAD TERRITORIAL`, 305, 60, { width: 398, align: 'center' });
        doc.rect(710, 50, 80, 70).fillAndStroke('#000', '#000').fillColor('#FFF').text(`APP\n07`, 710, 70, { width: 80, align: 'center' });
        doc.image(iecmLogoBN, 40, 55, {
            fit: [150, 60],
            align: 'center',
            valign: 'center'
        });
        doc.image(emblemaEC, 160, 55, {
            fit: [150, 60],
            align: 'center',
            valign: 'center'
        });
        doc.rect(50, 140, 740, 20).fillAndStroke('#000', '#000').fontSize(14).fillColor('#FFF').text('INFORMACIÓN DE LA UT', 50, 145, { width: 740, align: 'center' });
        const textos = [
            `DEMARCACIÓN: ${nombre_delegacion}`,
            `DD: ${id_distrito}`,
            `UT (clave): ${clave_colonia}`,
            `UT (nombre): ${nombre_colonia}`
        ];
        doc.font('Helvetica', 10).fillColor('#000');
        const widths = textos.map(t => doc.widthOfString(t));
        const espacio = (740 - widths.reduce((acum, width) => acum + width, 0)) / 3;
        let x = 50;
        for (let i = 0; i < textos.length; i++) {
            doc.text(textos[i], x, 165);
            if (i < 3)
                x += widths[i] + espacio;
        }
        doc.rect(50, 185, 740, 20).fillAndStroke('#000', '#000').font('Helvetica-Bold', 14).fillColor('#FFF').text('INFORMACIÓN DE LA VALIDACIÓN', 50, 190, { width: 740, align: 'center' });
        TextoMultiFuente(doc, 50, 210, 740, 10, [
            { text: 'En la Ciudad de México, siendo las', font: 'Helvetica' },
            { text: `${hora.substring(0, hora.length - 3)}`, font: 'Helvetica', underline: true },
            { text: ' horas del', font: 'Helvetica' },
            { text: `${fecha.split('/')[0]}`, font: 'Helvetica', underline: true },
            { text: ' de', font: 'Helvetica' },
            { text: `${NumAMes(+fecha.split('/')[1]).toLowerCase()}`, font: 'Helvetica', underline: true },
            { text: ' de', font: 'Helvetica' },
            { text: `${fecha.split('/')[2]}`, font: 'Helvetica' },
            { text: ', en el domicilio que ocupa la Dirección Distrital', font: 'Helvetica' },
            { text: `${id_distrito}`, font: 'Helvetica', underline: true },
            { text: ', situada en', font: 'Helvetica' },
            { text: `${direccion}`, font: 'Helvetica', underline: true },
            { text: ', se realizó el', font: 'Helvetica' },
            { text: 'CÓMPUTO TOTAL', font: 'Helvetica-Bold' },
            { text: `de la Unidad Territorial referida en la presente acta, correspondiente a la Consulta de ${eleccion}.`, font: 'Helvetica' }
        ], {
            fillColor: '#000',
            lineHeight: 1.5,
            align: 'justify'
        });
        let y = doc.y + 20;
        doc.font('Helvetica-Bold', 10).text('Por lo anterior, las personas funcionarias que suscriben la presente hacen constar el siguiente resultado:', 50, y);
        y += 25;
        doc.rect(50, y, 740, 20).fillAndStroke('#000', '#000').fontSize(14).fillColor('#FFF').text('RESULTADOS', 50, y + 5, { width: 740, align: 'center' });
        DibujarTablaPDF(doc, 50, doc.y - 1, [
            [{ text: 'Número de proyecto', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }],
            [{ text: 'Total con número', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }],
            [{ text: 'Total con letra', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }]
        ], [
            { width: 100, align: 'center' },
            { width: 100, align: 'center' },
            { width: 540, align: 'center' }
        ], datos);
        y = doc.page.height - 145;
        doc.rect(50, y, 740, 20).fillAndStroke('#A9A9A9', '#000').font('Helvetica-Bold', 14).fillColor('#FFF').text('Por la Dirección Distrital, suscriben:', 50, y + 5, { width: 740, align: 'center' });
        DibujarTablaPDF(doc, 50, y + 20, [
            [{ text: 'Cargo', font: 'Helvetica-Bold', fontSize: 14, fillColor: '#FFF', background: '#A9A9A9' }],
            [{ text: 'Nombre completo', font: 'Helvetica-Bold', fontSize: 14, fillColor: '#FFF', background: '#A9A9A9' }],
            [{ text: 'Firma', font: 'Helvetica-Bold', fontSize: 14, fillColor: '#FFF', background: '#A9A9A9' }]
        ], [
            { width: 280, align: 'center' },
            { width: 330, align: 'center' },
            { width: 130, align: 'center' }
        ], [
            [
                [{ text: coordinador_puesto, font: 'Helvetica', fontSize: 10 }],
                [{ text: coordinador, font: 'Helvetica', fontSize: 10 }],
                [{ text: '', font: 'Helvetica', fontSize: 10 }]
            ],
            [
                [{ text: secretario_puesto, font: 'Helvetica', fontSize: 10 }],
                [{ text: secretario, font: 'Helvetica', fontSize: 10 }],
                [{ text: '', font: 'Helvetica', fontSize: 10 }]
            ]
        ]);
        x = CalcularAltoAncho(doc, [{ text: secretario_puesto, font: 'Helvetica', fontSize: 10 }], 0, 280, 1.15).totalHeight;
        doc.font('Helvetica', 8).text(`Con fundamento en los artículos 36 primer párrafo, 113 fracción V, 366 y 367 segundo párrafo del Código de Instituciones y Procedimientos Electorales de la Ciudad de México 116, 117, 119, 120 inciso e) y 124 fracción IV de la Ley de Participación Ciudadana de la Ciudad de México, el apartado 14.5 del Manual de Geografía, Organización y Capacitación para la Preparación y Desarrollo de la Consulta de ${eleccion}; así como del párrafo tercero de la base décima quinta de las disposiciones comunes de la Convocatoria de la Consulta de ${eleccion}.`, 50, doc.y + (x / 2) + 8, { width: 740, align: 'justify' });
        doc.end();
        doc.on('data', buffer.push.bind(buffer));
        doc.on('end', () => {
            res.json({
                success: true,
                msg: 'Reporte generado correctamente',
                contentType: 'application/pdf',
                reporte: `ActaValidacion_${clave_colonia}_${fecha}-${hora}.pdf`,
                buffer: Buffer.concat(buffer)
            });
        });
    } catch (err) {
        console.error(`Error al generar el Acta de Validación en PDF: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el acta'
        });
    }
}