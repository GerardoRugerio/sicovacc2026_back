import DocxTemplater from 'docxtemplater';
import { request, response } from 'express';
import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import { plantillas } from '../helpers/Constantes.js';
import { ConsultaClaveColonia, ConsultaDelegacion, ConsultaDistrito, FechaServer } from '../helpers/Consultas.js';
import { NumAMes, NumAText } from '../helpers/Funciones.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Acta de Validacion

export const ActaComputoTotalWord = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia } = req.body;
    try {
        const { fecha, hora } = await FechaServer();
        const { nombre_delegacion } = await ConsultaDelegacion(id_distrito, clave_colonia);
        const { nombre_colonia } = await ConsultaClaveColonia(clave_colonia);
        const { direccion, coordinador, coordinador_puesto, secretario, secretario_puesto } = await ConsultaDistrito(id_distrito);
        const consulta = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia, modalidad, bol_nulas
            FROM copaco_actas
            WHERE id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
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
        let participantes = [];
        for (let participante of consulta.filter(participante => participante.secuencial != '0')) {
            const { total_votos } = participante;
            participantes.push({
                ...participante,
                total_votosL: NumAText(total_votos)
            });
        }
        fs.readFile(path.join(plantillas[0], 'Acta_Computo_Total.docx'), 'binary', (err, content) => {
            if (err)
                return res.status(500).json({
                    success: false,
                    msg: 'Error al abrir la plantilla'
                });

            const zip = new PizZip(content);
            const docx = new DocxTemplater(zip, { linebreaks: true, paragraphLoop: true });

            const data = {
                demarcacion: nombre_delegacion,
                dd: id_distrito,
                ut: clave_colonia,
                colonia: nombre_colonia,
                hora: hora.substring(0, hora.length - 3),
                dia: fecha.split('/')[0],
                mes: NumAMes(+fecha.split('/')[1]).toLowerCase(),
                anio: +fecha.split('/')[2],
                direccion,
                participantes,
                nulas, nulas_sei, total_nulas, total_nulasL: NumAText(total_nulas),
                totalN, totalNS, total, totalL: NumAText(total),
                coordinador,
                coordinador_puesto,
                secretario,
                secretario_puesto
            };

            docx.render(data);

            res.json({
                success: true,
                msg: 'Acta de Validación generada correctamente',
                contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                reporte: `ActaComputoTotal_${clave_colonia}_${fecha}-${hora}.docx`,
                buffer: docx.getZip().generate({ type: 'nodebuffer' })
            });
        })
    } catch (err) {
        console.error(`Error al generar el Acta de Validación en WORD: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el acta'
        })
    }
}