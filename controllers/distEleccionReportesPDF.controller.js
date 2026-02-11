import { request, response } from 'express';
import PDFDocument from 'pdfkit';
import { CalcularAltoAncho, DibujarTablaPDF, TextoMultiFuente } from '../helpers/ActasPDF.js';
import { autor, emblemaEC, iecmLogoBN } from '../helpers/Constantes.js';
import { ConsultaClaveColonia, ConsultaDelegacion, ConsultaDistrito, FechaServer } from '../helpers/Consultas.js';
import { NumAMes, NumAText } from '../helpers/Funciones.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Acta de Validación

export const ActaValidacionPDF = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia } = req.body;
    try {
        const { fecha, hora } = await FechaServer();
        const { nombre_delegacion } = await ConsultaDelegacion(id_distrito, clave_colonia);
        const { nombre_colonia } = await ConsultaClaveColonia(clave_colonia);
        const { direccion, coordinador, coordinador_puesto, secretario, secretario_puesto } = await ConsultaDistrito(id_distrito);
        const consulta = (await SICOVACC.sequelize.query(`SELECT *
        FROM (
            SELECT dbo.NumeroALetras(secuencial) AS secuencial, SUM(total_votos) AS total_votos
            FROM copaco_actas_VVS V
            WHERE id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
            GROUP BY V.secuencial
        ) A
        UNION ALL
        SELECT '0' AS secuencial, SUM(bol_nulas) AS total_votos
        FROM copaco_actas
        WHERE id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'`))[0];
        const { total_votos: bol_nulas } = consulta.find(participante => participante.secuencial == '0');
        const total = consulta.reduce((sum, participante) => sum + participante.total_votos, 0);
        let datos = [], buffer = [];
        consulta.filter(participante => participante.secuencial != '0').forEach(participante => {
            let X = [];
            Object.keys(participante).forEach((key, index) => {
                X.push([{ text: participante[key].toString(), font: 'Helvetica', fontSize: 10 }]);
                if (index == 1)
                    X.push([{ text: NumAText(participante[key]), font: 'Helvetica', fontSize: 10 }]);
            });
            datos.push(X);
        });
        datos.push([
            [{ text: 'Opiniones nulas', font: 'Helvetica-Bold', fontSize: 10, background: '#C0C0C0' }],
            [{ text: String(bol_nulas), font: 'Helvetica-Bold', fontSize: 10, background: '#C0C0C0' }],
            [{ text: NumAText(bol_nulas), font: 'Helvetica-Bold', fontSize: 10, background: '#C0C0C0' }],
        ], [
            [{ text: 'TOTAL', font: 'Helvetica-Bold', fontSize: 14, background: '#C0C0C0', fillColor: '#FFF' }],
            [{ text: String(total), font: 'Helvetica-Bold', fontSize: 14, background: '#C0C0C0' }],
            [{ text: NumAText(total), font: 'Helvetica-Bold', fontSize: 14, background: '#C0C0C0' }]
        ]);
        const doc = new PDFDocument({ bufferPages: true, autoFirstPage: false, size: 'A3', layout: 'portrait', margin: 30 });
        doc.info.Author = autor;
        doc.addPage();
        doc.rect(50, 50, 658, 70).fillAndStroke('#000', '#000').font('Helvetica-Bold', 16).fillColor('#FFF').text(`ACTA DE VALIDACIÓN DE RESULTADOS PARA LA ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA POR UNIDAD TERRITORIAL`, 305, 60, { width: 398, align: 'center' });
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
            { text: `de la Unidad Territorial referida en la presente acta, correspondiente a la Elección de Comisiones de Participación Comunitaria.`, font: 'Helvetica' }
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
            [{ text: 'Letra del Participante', font: 'Helvetica-Bold', fontSize: 12, background: '#C0C0C0' }],
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
        doc.font('Helvetica', 8).text(`Con fundamento en los artículos 36 primer párrafo, 113 fracción V, 366 y 367 segundo párrafo del Código de Instituciones y Procedimientos Electorales de la Ciudad de México 116, 117, 119, 120 inciso e) y 124 fracción IV de la Ley de Participación Ciudadana de la Ciudad de México, el apartado 14.5 del Manual de Geografía, Organización y Capacitación para la Preparación y Desarrollo de la Elección de Comisiones de Participación Comunitaria; así como del párrafo tercero de la base décima quinta de las disposiciones comunes de la Convocatoria de la Elección de Comisiones de Participación Comunitaria.`, 50, doc.y + (x / 2) + 8, { width: 740, align: 'justify' });
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