import DocxTemplater from 'docxtemplater';
import { request, response } from 'express';
import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import { plantillas } from '../helpers/Constantes.js';
import { ConsultaClaveColonia, ConsultaDelegacion, ConsultaDistrito, ConsultaTipoEleccion, FechaServer } from '../helpers/Consultas.js';
import { NumAMes, NumAText } from '../helpers/Funciones.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Acta de Validacion

export const ActaValidacionWord = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia, anio } = req.body;
    try {
        const { fecha, hora } = await FechaServer();
        const { nombre_delegacion } = await ConsultaDelegacion(id_distrito, clave_colonia);
        const { nombre_colonia } = await ConsultaClaveColonia(clave_colonia);
        const { direccion, coordinador, coordinador_puesto, secretario, secretario_puesto } = await ConsultaDistrito(id_distrito);
        const X = await ConsultaTipoEleccion(anio);
        const eleccion1 = `ELECCIÓN DE ${X.toUpperCase()}`, eleccion2 = `Elección de ${X}`;
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
        const { total_votos: bol_nulas } = consulta.find(proyecto => proyecto.secuencial == '0');
        const total = consulta.reduce((sum, proyecto) => sum + proyecto.total_votos, 0);
        let proyectos = [];
        for (let proyecto of consulta.filter(proyecto => proyecto.secuencial != '0')) {
            const { total_votos } = proyecto;
            proyectos.push({
                ...proyecto,
                total_votosL: NumAText(total_votos)
            });
        }
        fs.readFile(path.join(plantillas[0], 'Acta_Validacion.docx'), 'binary', (err, content) => {
            if (err)
                return res.status(500).json({
                    success: false,
                    msg: 'Error al abrir la plantilla'
                });

            const zip = new PizZip(content);
            const docx = new DocxTemplater(zip, { linebreaks: true, paragraphLoop: true });

            const data = {
                eleccion1,
                eleccion2,
                demarcacion: nombre_delegacion,
                dd: id_distrito,
                ut: clave_colonia,
                colonia: nombre_colonia,
                hora: hora.substring(0, hora.length - 3),
                dia: fecha.split('/')[0],
                mes: NumAMes(+fecha.split('/')[1]).toLowerCase(),
                direccion,
                proyectos,
                nulas: bol_nulas,
                nulasL: NumAText(bol_nulas),
                total,
                totalL: NumAText(total),
                titulo: 'Letra del Participante',
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
                reporte: `ActaValidacion_${clave_colonia}_${fecha}-${hora}.docx`,
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