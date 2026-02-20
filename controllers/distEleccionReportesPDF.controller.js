import { request, response } from 'express';
import PDFDocument from 'pdfkit';
import { CalcularAltoAncho, DibujarTablaPDF, TextoMultiFuente } from '../helpers/ActasPDF.js';
import { autor, CalendarioAzteca, IECMLogoBN } from '../helpers/Constantes.js';
import { ConsultaClaveColonia, ConsultaDelegacion, ConsultaDistrito, FechaServer } from '../helpers/Consultas.js';
import { DividirArreglo, NumAMes, NumAText } from '../helpers/Funciones.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Acta de Cómputo Total

export const ActaComputoTotalPDF = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia } = req.body;
    try {
        const { fecha, hora } = await FechaServer();
        const { nombre_delegacion } = await ConsultaDelegacion(id_distrito, clave_colonia);
        const { nombre_colonia } = await ConsultaClaveColonia(clave_colonia);
        const { direccion, coordinador, coordinador_puesto, secretario, secretario_puesto } = await ConsultaDistrito(id_distrito);
        const consulta = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia, modalidad, SUM(bol_nulas) AS bol_nulas
            FROM copaco_actas
            WHERE id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
            GROUP BY id_distrito, clave_colonia, modalidad
        ),
        Acta AS (
            SELECT V.secuencial AS orden, dbo.NumeroALetras(V.secuencial) AS secuencial, SUM(V.votos) AS votos, SUM(V.votos_sei) AS votos_sei, SUM(V.total_votos) AS total_votos
            FROM CA
            INNER JOIN copaco_actas_VVS V ON CA.id_distrito = V.id_distrito AND CA.clave_colonia = V.clave_colonia
            WHERE CA.modalidad = 1
            GROUP BY V.secuencial
        )
        SELECT secuencial, votos, votos_sei, total_votos
        FROM (
            SELECT 0 AS orden, '0' AS secuencial, A1.bol_nulas AS votos, COALESCE(A2.bol_nulas, 0) AS votos_sei, A1.bol_nulas + COALESCE(A2.bol_nulas, 0) AS total_votos
            FROM CA A1
            LEFT JOIN CA A2 ON A2.modalidad = 2
            WHERE A1.modalidad = 1
            UNION ALL
            SELECT orden, secuencial, votos, votos_sei, total_votos
            FROM Acta
        ) X
        ORDER BY orden`))[0];
        const { votos: nulas, votos_sei: nulas_sei, total_votos: total_nulas } = consulta.find(participante => participante.secuencial == '0');
        const totalN = consulta.reduce((sum, participante) => sum + participante.votos, 0);
        const totalNS = consulta.reduce((sum, participante) => sum + participante.votos_sei, 0);
        const total = consulta.reduce((sum, participante) => sum + participante.total_votos, 0);
        let datos = [], buffer = [];
        consulta.filter(participante => participante.secuencial != '0').forEach(participante => {
            let X = [];
            Object.keys(participante).forEach((key, index) => {
                X.push([{ text: participante[key].toString(), font: 'Helvetica', fontSize: 10, strokeColor: '#BFBFBF' }]);
                if (index == 3)
                    X.push([{ text: NumAText(participante[key]), font: 'Helvetica', fontSize: 10, strokeColor: '#BFBFBF' }]);
            });
            datos.push(X);
        });
        datos.push([
            [{ text: 'VOTOS NULOS', font: 'Helvetica-Bold', fontSize: 10, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
            [{ text: String(nulas), font: 'Helvetica-Bold', fontSize: 10, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
            [{ text: String(nulas_sei), font: 'Helvetica-Bold', fontSize: 10, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
            [{ text: String(total_nulas), font: 'Helvetica-Bold', fontSize: 10, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
            [{ text: NumAText(total_nulas), font: 'Helvetica-Bold', fontSize: 10, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
        ], [
            [{ text: 'TOTAL', font: 'Helvetica-Bold', fontSize: 14, fillColor: '#FFF', background: '#000' }],
            [{ text: String(totalN), font: 'Helvetica-Bold', fontSize: 14, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
            [{ text: String(totalNS), font: 'Helvetica-Bold', fontSize: 14, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
            [{ text: String(total), font: 'Helvetica-Bold', fontSize: 14, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
            [{ text: NumAText(total), font: 'Helvetica-Bold', fontSize: 14, background: '#F2F2F2', strokeColor: '#BFBFBF' }]
        ]);
        const subDatos = DividirArreglo(datos, 38);
        const doc = new PDFDocument({ bufferPages: true, autoFirstPage: false, size: 'A3', layout: 'portrait', margin: 30 });
        doc.info.Author = autor;
        doc.addPage();
        const textos = [
            `DEMARCACIÓN: ${nombre_delegacion}`,
            `DD: ${id_distrito}`,
            `UT (clave): ${clave_colonia}`,
            `UT (nombre): ${nombre_colonia}`
        ];
        doc.font('Helvetica', 10).fillColor('#000');
        const widths = textos.map(t => doc.widthOfString(t));
        const espacio = (740 - widths.reduce((acum, width) => acum + width, 0)) / 3;
        let x = 0, y = 0;
        for (let i = 0; i < subDatos.length; i++) {
            doc.rect(50, 50, 740, 70).fillAndStroke('#000', '#000');
            TextoMultiFuente(doc, 190, 68, 425, 16, [
                { text: 'ACTA DE CÓMPUTO TOTAL', font: 'Helvetica-Bold' },
                { text: 'DE LA ELECCIÓN DE LAS COMISIONES DE PARTICIPACIÓN COMUNITARIA 2026', font: 'Helvetica' }
            ], {
                fillColor: '#FFF',
                lineHeight: 1.5,
                align: 'center'
            });
            doc.font('Helvetica-Bold', 16).fillColor('#FFF').text(`AC\n05`, 710, 70, { width: 80, align: 'center' });
            doc.image(IECMLogoBN, 40, 55, {
                fit: [150, 60],
                align: 'center',
                valign: 'center'
            });
            doc.image(CalendarioAzteca, 600, 55, {
                fit: [150, 60],
                align: 'center',
                valign: 'center'
            });
            doc.rect(50, 140, 740, 20).fillAndStroke('#F2F2F2', '#BFBFBF').fontSize(14).fillColor('#000').text('INFORMACIÓN DE LA UT', 50, 145, { width: 740, align: 'center' });
            doc.font('Helvetica', 10).fillColor('#000');
            x = 50;
            for (let j = 0; j < textos.length; j++) {
                doc.text(textos[j], x, 165);
                if (j < 3)
                    x += widths[j] + espacio;
            }
            doc.rect(50, 185, 740, 20).fillAndStroke('#F2F2F2', '#BFBFBF').font('Helvetica-Bold', 14).fillColor('#000').text('INFORMACIÓN DE LA VALIDACIÓN', 50, 190, { width: 740, align: 'center' });
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
                { text: 'cómputo total', font: 'Helvetica-Bold' },
                { text: 'de la Unidad Territorial referida en la presente acta, correspondiente a la', font: 'Helvetica' },
                { text: 'Elección de las Comisiones de Participación Comunitaria 2026', font: 'Helvetica-Bold', underline: true },
                { text: '. Por lo anterior,', font: 'Helvetica' },
                { text: 'las personas funcionarias que suscriben la presente, hacen constar los siguientes resultados:', font: 'Helvetica-Bold' }
            ], {
                fillColor: '#000',
                lineHeight: 1.5,
                align: 'justify'
            });
            y = doc.y + 20;
            doc.rect(50, y, 740, 20).fillAndStroke('#F2F2F2', '#BFBFBF').fontSize(14).fillColor('#000').text('RESULTADOS', 50, y + 5, { width: 740, align: 'center' });
            DibujarTablaPDF(doc, 50, doc.y - 1, [
                [{ text: 'LETRA(S) DE CANDIDATURA', font: 'Helvetica-Bold', fontSize: 12, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
                [
                    { text: 'RESULTADOS DEL ESCRUTINIO Y CÓMPUTO', font: 'Helvetica-Bold', fontSize: 12, background: '#F2F2F2', strokeColor: '#BFBFBF' },
                    { text: '(VOTOS SACADOS DE LA URNA)', font: 'Helvetica', fontSize: 12, background: '#F2F2F2', strokeColor: '#BFBFBF' }
                ],
                [
                    { text: 'RESULTADOS DEL COMPUTO DEL SEI', font: 'Helvetica-Bold', fontSize: 12, background: '#F2F2F2', strokeColor: '#BFBFBF' },
                    { text: '(ASENTADOS EN EL ACTA)', font: 'Helvetica', fontSize: 12, background: '#F2F2F2', strokeColor: '#BFBFBF' }
                ],
                [{ text: 'TOTAL CON NÚMERO', font: 'Helvetica-Bold', fontSize: 12, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
                [{ text: 'TOTAL CON LETRA', font: 'Helvetica-Bold', fontSize: 12, background: '#F2F2F2', strokeColor: '#BFBFBF' }]
            ], [
                { width: 100, align: 'center' },
                { width: 100, align: 'center' },
                { width: 100, align: 'center' },
                { width: 100, align: 'center' },
                { width: 340, align: 'center' }
            ], subDatos[i]);
            if (i < subDatos.length - 1)
                doc.addPage();
        }
        y = doc.page.height - 145 - (subDatos.length > 1 ? 15 : 0);
        doc.rect(50, y, 740, 20).fillAndStroke('#F2F2F2', '#BFBFBF').font('Helvetica-Bold', 14).fillColor('#000').text('POR LA DIRECCIÓN DISTRITAL, SUSCRIBEN:', 50, y + 5, { width: 740, align: 'center' });
        DibujarTablaPDF(doc, 50, y + 20, [
            [{ text: 'CARGO', font: 'Helvetica-Bold', fontSize: 14, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
            [{ text: 'NOMBRE COMPLETO', font: 'Helvetica-Bold', fontSize: 14, background: '#F2F2F2', strokeColor: '#BFBFBF' }],
            [{ text: 'FIRMA', font: 'Helvetica-Bold', fontSize: 14, background: '#F2F2F2', strokeColor: '#BFBFBF' }]
        ], [
            { width: 280, align: 'center' },
            { width: 330, align: 'center' },
            { width: 130, align: 'center' }
        ], [
            [
                [{ text: coordinador_puesto, font: 'Helvetica', fontSize: 10, strokeColor: '#BFBFBF' }],
                [{ text: coordinador, font: 'Helvetica', fontSize: 10, strokeColor: '#BFBFBF' }],
                [{ text: '', font: 'Helvetica', fontSize: 10, strokeColor: '#BFBFBF' }]
            ],
            [
                [{ text: secretario_puesto, font: 'Helvetica', fontSize: 10, strokeColor: '#BFBFBF' }],
                [{ text: secretario, font: 'Helvetica', fontSize: 10, strokeColor: '#BFBFBF' }],
                [{ text: '', font: 'Helvetica', fontSize: 10, strokeColor: '#BFBFBF' }]
            ]
        ]);
        x = CalcularAltoAncho(doc, [{ text: secretario_puesto, font: 'Helvetica', fontSize: 10 }], 0, 280, 1.15).totalHeight;
        doc.font('Helvetica', 8).text('SE LEVANTA LA PRESENTE ACTA CON FUNDAMENTO EN LO DISPUESTO EN LOS ARTÍCULOS 6 FRACCIÓN I, 36 PÁRRAFO PRIMERO, 113 FRACCIÓN V, 362 PRIMER Y SEGUNDO PARRAFO, Y 367 DEL CÓDIGO DE INSTITUCIONES Y PROCEDIMIENTOS ELECTORALES DE LA CIUDAD DE MÉXICO; 83, 96 PRIMER PÁRRAFO, 97 Y 106 DE LA LEY DE PARTICIPACIÓN CIUDADANA DE LA CIUDAD DE MÉXICO; ASÍ COMO DEL NUMERAL 16 DE LAS DISPOSICIONES GENERALES DE LA CONVOCATORIA ÚNICA APROBADA POR EL CONSEJO GENERAL DEL INSTITUTO ELECTORAL DE LA CIUDAD DE MÉXICO MEDIANTE ACUERDO IECM/ACU-CG-004/2026 DE FECHA 09 DE ENERO DE 2026.', 50, doc.y + (x / 2) + 9, { width: 740, align: 'justify' });
        if (subDatos.length > 1) {
            const paginas = doc.bufferedPageRange().count;
            for (let i = 0; i < paginas; i++) {
                doc.switchToPage(i);
                doc.font('Helvetica', 10).text(`Hoja ${i + 1} de ${paginas}`, 50, doc.page.height - 45, { width: 740, align: 'center' });
            }
        }
        doc.end();
        doc.on('data', buffer.push.bind(buffer));
        doc.on('end', () => {
            res.json({
                success: true,
                msg: 'Reporte generado correctamente',
                contentType: 'application/pdf',
                reporte: `ActaComputoTotal_${clave_colonia}_${fecha}-${hora}.pdf`,
                buffer: Buffer.concat(buffer)
            });
        });
    } catch (err) {
        console.error(`Error al generar el Acta de Cómputo Total en PDF: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el acta'
        });
    }
}