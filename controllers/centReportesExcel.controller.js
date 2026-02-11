import ExcelJs from 'exceljs';
import { request, response } from 'express';
import path from 'path';
import { autor, contenidoStyle, fill, iecmLogo, plantillas, titulos } from '../helpers/Constantes.js';
import { FechaServer } from '../helpers/Consultas.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Reporte de asistencia de inicio y cierre de la validación

export const InicioCierreValidacion = async (req = request, res = response) => {
    const workbook = new ExcelJs.Workbook();
    try {
        const validacion = (await SICOVACC.sequelize.query(`SELECT id_distrito, CONVERT(VARCHAR(10), fecha_hora_inicio, 103) AS fecha_inicio, CONVERT(VARCHAR(5), fecha_hora_inicio, 114) AS hora_inicio, inicio_asistencia1, inicio_asistencia2, inicio_asistencia4, inicio_asistencia5, inicio_asistencia6, inicio_asistencia7, inicio_total, UPPER(inicio_observaciones) AS inicio_observaciones,
        CONVERT(VARCHAR(10), fecha_hora_cierre, 103) AS fecha_cierre, CONVERT(VARCHAR(5), fecha_hora_cierre, 114) AS hora_cierre, cierre_asistencia1, cierre_asistencia2, cierre_asistencia4, cierre_asistencia5, cierre_asistencia6, cierre_asistencia7, cierre_total, UPPER(cierre_observaciones) AS cierre_observaciones FROM consulta_computo WHERE estatus = 1 ORDER BY id_distrito ASC`))[0];
        if (!validacion.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[0].replace('consulta/', ''), 'Inicio-Cierre_Validacion.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 11;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:M2');
                worksheet.getCell('A3').value = titulos[1];
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:M3');
                worksheet.getCell('A5').value = 'PENDIENTE';
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:M5');
                worksheet.getCell('A6').value = 'REPORTE DE ASISTENCIA DE INICIO Y CIERRE DE LA VALIDACIÓN';
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:M6');
                worksheet.getCell('L7').value = `Fecha: ${fecha}`;
                worksheet.getCell('L8').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                worksheet.getCell('A9').value = 'DEOEyG';
                worksheet.getCell('A9').style = fill;
                worksheet.getCell('B9').value = 'Asistencia de Inicio';
                worksheet.getCell('B9').style = fill;
                if (!worksheet.getCell('B9').isMerged)
                    worksheet.mergeCells('B9:K9')
                worksheet.getCell('L9').value = 'Asistencia de Cierre';
                worksheet.getCell('L9').style = fill;
                if (!worksheet.getCell('L9').isMerged)
                    worksheet.mergeCells('L9:U9')
                worksheet.getCell('A10').value = 'Distrito';
                validacion.forEach(val => {
                    Object.keys(val).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = val[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getColumn(1).width = 13;
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_Central_InicioCierreValidacion-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en InicioCierreValidacion: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en InicioCierreValidacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F3 - Incidentes Presentados Durante la Validación de la Consulta de Presupuesto Participativo

export const Incidentes = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    // const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const incidentes = (await SICOVACC.sequelize.query(`SELECT I.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, I.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(I.num_mro, CASE WHEN TP.mesa IS NOT NULL THEN CONCAT(' ', TP.mesa) END) AS mesa,
        CONCAT(CASE WHEN I.anio = 1 THEN 'ELECCIÓN' ELSE 'CONSULTA' END, ' DE ', UPPER(TE.descripcion)) AS eleccion,
        CASE WHEN I.incidente_1 = 1 THEN 'X' ELSE '' END AS i1, CASE WHEN I.incidente_2 = 1 THEN 'X' ELSE '' END AS i2, CASE WHEN I.incidente_3 = 1 THEN 'X' ELSE '' END AS i3, CASE WHEN I.incidente_4 = 1 THEN 'X' ELSE '' END AS i4,
        CASE WHEN I.incidente_5 = 1 THEN 'X' ELSE '' END AS i5, UPPER(I.participantes) AS participantes, UPPER(I.hechos) AS hechos, UPPER(I.acciones) AS acciones,
        CONCAT(CONVERT(VARCHAR(10), I.fecha_hora, 103), ' ', CONVERT(VARCHAR(5), I.fecha_hora, 114)) AS fecha
        FROM consulta_incidentes I
        LEFT JOIN consulta_tipo_mesa_V TP ON I.tipo_mro = TP.tipo_mro
        LEFT JOIN consulta_cat_delegacion D ON I.id_delegacion = D.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 C ON I.clave_colonia = C.clave_colonia
        LEFT JOIN consulta_cat_tipo_eleccion TE ON I.anio = TE.id_tipo_eleccion
        WHERE I.estatus = 1${id_distrito != 0 ? ` AND I.id_distrito = ${id_distrito}` : ''}
        ORDER BY I.id_distrito, I.anio, D.nombre_delegacion, C.nombre_colonia, I.num_mro, I.tipo_mro`))[0];
        if (!incidentes.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[0], 'Incidentes.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12, inc = [0, 0, 0, 0, 0];
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:O2');
                worksheet.getCell('A3').value = titulos[1];
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:O3');
                worksheet.getCell('A5').value = 'PENDIENTE';
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:O5');
                worksheet.getCell('A7').value = 'INCIDENTES PRESENTADOS DURANTE LA VALIDACIÓN DE LA ELECCIÓN Y LA CONSULTA';
                if (!worksheet.getCell('A7').isMerged)
                    worksheet.mergeCells('A7:O7');
                worksheet.getCell('N9').value = `Fecha: ${fecha}`;
                worksheet.getCell('N10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                worksheet.getCell('G10').value = 'CAUSAS';
                worksheet.getCell('G10').style = fill;
                if (!worksheet.getCell('G10').isMerged)
                    worksheet.mergeCells('G10:K10');
                incidentes.forEach(incidente => {
                    Object.keys(incidente).forEach((key, i) => {
                        worksheet.getCell(fila, i + 1).value = incidente[key];
                        worksheet.getCell(fila, i + 1).style = contenidoStyle;
                        if (['i1', 'i2', 'i3', 'i4', 'i5'].includes(key))
                            switch (key) {
                                case 'i1':
                                    if (incidente[key] == 'X')
                                        inc[0]++;
                                    break;
                                case 'i2':
                                    if (incidente[key] == 'X')
                                        inc[1]++;
                                    break;
                                case 'i3':
                                    if (incidente[key] == 'X')
                                        inc[2]++;
                                    break;
                                case 'i4':
                                    if (incidente[key] == 'X')
                                        inc[3]++;
                                    break;
                                case 'i5':
                                    if (incidente[key] == 'X')
                                        inc[4]++;
                                    break;
                            }
                    });
                    fila++;
                });
                for (let i = 0; i <= 4; i++) {
                    worksheet.getCell(fila, i + 7).value = inc[i];
                    worksheet.getCell(fila, i + 7).style = { ...contenidoStyle, numFmt: '#,##0' };
                }
                fila++;
                worksheet.getCell(fila, 7).value = 'Total de Causas de Incidentes:'
                worksheet.getCell(fila, 7).style = fill;
                worksheet.getCell(fila, 8).value = inc.reduce((sum, i) => sum + i, 0);
                worksheet.getCell(fila, 8).style = { ...contenidoStyle, numFmt: "#,##0" };
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3 || (index >= 10 && index <= 12)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 10)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 14;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 16)
                            column.width = 16;
                        else
                            column.width = maxLength;
                    }
                });
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_Incidentes-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en Incidentes: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en Incidentes: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}