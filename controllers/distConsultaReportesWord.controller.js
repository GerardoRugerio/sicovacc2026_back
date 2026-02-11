import DocxTemplater from 'docxtemplater';
import { request, response } from 'express';
import fs from 'fs';
import path from 'path';
import PizZip from 'pizzip';
import { anioN, plantillas } from '../helpers/Constantes.js';
import { ConsultaClaveColonia, ConsultaDelegacion, ConsultaDistrito, ConsultaTipoEleccion, FechaServer, InformacionConstancia } from '../helpers/Consultas.js';
import { NumAMes, NumAText } from '../helpers/Funciones.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Constnacia - En Desuso

export const ConstanciaWord = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia, anio } = req.body;
    try {
        const { fecha, hora } = await FechaServer();
        const { nombre_colonia, nombre_delegacion, domicilio, mesas, ultimaFecha, ultimaHora, coordinador_puesto, coordinador, secretario_puesto, secretario } = await InformacionConstancia(anio, clave_colonia);
        const dia = NumAText(ultimaFecha.split('/')[0]).toLowerCase(), mes = NumAMes(+ultimaFecha.split('/')[1]).toLowerCase(), horas = NumAText(ultimaHora.split(':')[0]).toLowerCase(), minutos = NumAText(ultimaHora.split(':')[1]).toLowerCase();
        const mesasL = NumAText(mesas);
        const consProyectos = await SICOVACC.sequelize.query(`SELECT num_proyecto, nom_proyecto, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
        FROM consulta_actas_VVS
        WHERE anio = ${anio} AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}' AND num_proyecto IN (SELECT num_proyecto FROM consulta_prelacion_proyectos WHERE anio = ${anio} AND estatus = 1 AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}')
        GROUP BY num_proyecto, nom_proyecto
        ORDER BY num_proyecto`);
        let proyectos = [];
        for (let proyecto of consProyectos[0]) {
            const { votos, votos_sei, total_votos } = proyecto;
            proyectos.push({
                ...proyecto,
                votos: Intl.NumberFormat('es-MX').format(votos),
                votos_sei: Intl.NumberFormat('es-MX').format(votos_sei),
                total_votos: Intl.NumberFormat('es-MX').format(total_votos),
                totalL: NumAText(total_votos)
            });
        }

        fs.readFile(path.join(plantillas, 'Constancia.docx'), 'binary', (err, content) => {
            if (err)
                return res.status(500).json({
                    success: false,
                    msg: 'Error al abrir la plantilla'
                });

            const zip = new PizZip(content);
            const docx = new DocxTemplater(zip, { linebreaks: true, paragraphLoop: true });

            const data = {
                nombre_colonia,
                clave: clave_colonia,
                distrito: id_distrito,
                nombre_delegacion,
                mesas, mesasL,
                horas, minutos,
                dia,
                mes,
                anio: anioN[anio], domicilio,
                diaH: fecha.split('/')[0],
                mesH: NumAMes(+fecha.split('/')[1]).toLowerCase(),
                coordinador_puesto, coordinador,
                secretario_puesto, secretario,
                proyectos
            };

            docx.render(data);

            res.json({
                success: true,
                msg: 'Constancia generada correctamente',
                contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                reporte: `Constancia_${clave_colonia}_${fecha}-${hora}.docx`,
                buffer: docx.getZip().generate({ type: 'nodebuffer' })
            });
        })
    } catch (err) {
        console.error(`Error al generar la constancia en WORD: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar la constancia'
        });
    }
}

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
        const eleccion1 = `CONSULTA DE ${X.toUpperCase()}`, eleccion2 = `Consulta de ${X}`;
        const consulta = (await SICOVACC.sequelize.query(`SELECT secuencial, SUM(total_votos) AS total_votos
        FROM consulta_actas_VVS
        WHERE anio = ${anio} AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
        GROUP BY secuencial
        UNION ALL
        SELECT 0 AS secuencial, SUM(bol_nulas) AS total_votos
        FROM consulta_actas
        WHERE anio = ${anio} AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
        ORDER BY secuencial`))[0];
        const { total_votos: bol_nulas } = consulta.find(proyecto => proyecto.secuencial == 0);
        const total = consulta.reduce((sum, proyecto) => sum + proyecto.total_votos, 0);
        let proyectos = [];
        for (let proyecto of consulta.filter(proyecto => proyecto.secuencial != 0)) {
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
                titulo: 'Número de proyecto',
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