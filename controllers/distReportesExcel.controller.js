import ExcelJs from 'exceljs';
import { request, response } from 'express';
import path from 'path';
import { autor, contenidoStyle, fill, IECMLogo, plantillas, titulos } from '../helpers/Constantes.js';
import { FechaServer } from '../helpers/Consultas.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Inicio - Cierre de Validación

export const InicioCierreValidacion = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const validacion = (await SICOVACC.sequelize.query(`SELECT CONVERT(VARCHAR(10), fecha_hora_inicio, 103) AS fecha_inicio, CONVERT(VARCHAR(5), fecha_hora_inicio, 114) AS hora_inicio, inicio_asistencia1, inicio_asistencia2, inicio_asistencia4, inicio_asistencia5, inicio_asistencia6, inicio_asistencia7, inicio_total, UPPER(inicio_observaciones) AS inicio_observaciones,
        CONVERT(VARCHAR(10), fecha_hora_cierre, 103) AS fecha_cierre, CONVERT(VARCHAR(5), fecha_hora_cierre, 114) AS hora_cierre, cierre_asistencia1, cierre_asistencia2, cierre_asistencia4, cierre_asistencia5, cierre_asistencia6, cierre_asistencia7, cierre_total, UPPER(cierre_observaciones) AS cierre_observaciones
        FROM consulta_computo WHERE estatus = 1 AND id_distrito = ${id_distrito}`))[0];
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
                worksheet.spliceColumns(1, 1);
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                let fila = 13;
                worksheet.getCell('A2').value = titulos[0];
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:T2');
                worksheet.getCell('A3').value = titulos[1];
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:T3');
                worksheet.getCell('A5').value = 'PENDIENTE';
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:T5');
                worksheet.getCell('A6').value = 'REPORTE DE ASISTENCIA DE INICIO Y CIERRE DE LA VALIDACIÓN';
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:T6');
                worksheet.getCell('A8').value = `DIRECCIÓN DISTRITAL: ${id_distrito}`;
                worksheet.getCell('A8').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('S8').value = `Fecha: ${fecha}`;
                worksheet.getCell('S9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                worksheet.getCell('A11').value = 'Asistencia de Inicio';
                worksheet.getCell('A11').style = fill;
                if (!worksheet.getCell('A11').isMerged)
                    worksheet.mergeCells('A11:J11');
                worksheet.getCell('K11').value = 'Asistencia de Cierre';
                worksheet.getCell('K11').style = fill;
                if (!worksheet.getCell('K11').isMerged)
                    worksheet.mergeCells('K11:T11');
                validacion.forEach(val => {
                    Object.keys(val).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = val[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.columns.forEach((column, i) => {
                    if ([10, 21].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 12)
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
                    reporte: `Reporte_InicioCierreValidación-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en InicioCierreValidacion: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            })
    } catch (err) {
        console.error(`Error en InicioCierreValidacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F3 - Listado de Incidentes Presentados en la Validación de la Consulta de Presupuesto Participativo

export const IncidentesDistrito = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    // const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const incidentes = (await SICOVACC.sequelize.query(`SELECT UPPER(D.nombre_delegacion) AS nombre_delegacion, I.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(I.num_mro, CASE WHEN TP.mesa IS NOT NULL THEN CONCAT(' ', TP.mesa) END) AS mesa,
        CONCAT(CASE WHEN I.anio = 1 THEN 'ELECCIÓN' ELSE 'CONSULTA' END, ' DE ', UPPER(TE.descripcion)) AS eleccion,
        CASE WHEN I.incidente_1 = 1 THEN 'X' ELSE '' END AS i1, CASE WHEN I.incidente_2 = 1 THEN 'X' ELSE '' END AS i2, CASE WHEN I.incidente_3 = 1 THEN 'X' ELSE '' END AS i3, CASE WHEN I.incidente_4 = 1 THEN 'X' ELSE '' END AS i4,
        CASE WHEN I.incidente_5 = 1 THEN 'X' ELSE '' END AS i5, UPPER(I.participantes) AS participantes, UPPER(I.hechos) AS hechos, UPPER(I.acciones) AS acciones,
        CONCAT(CONVERT(VARCHAR(10), I.fecha_hora, 103), ' ', CONVERT(VARCHAR(5), I.fecha_hora, 114)) AS fecha
        FROM consulta_incidentes I
        LEFT JOIN consulta_tipo_mesa_V TP ON I.tipo_mro = TP.tipo_mro
        LEFT JOIN consulta_cat_delegacion D ON I.id_delegacion = D.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 C ON I.clave_colonia = C.clave_colonia
        LEFT JOIN consulta_cat_tipo_eleccion TE ON I.anio = TE.id_tipo_eleccion
        WHERE I.estatus = 1 AND I.id_distrito = ${id_distrito}
        ORDER BY I.anio, D.nombre_delegacion, C.nombre_colonia, I.num_mro, I.tipo_mro ASC`))[0];
        /*         CASE WHEN A.incidente_6 = 1 THEN 'X' ELSE '' END AS i6,
        CASE WHEN A.incidente_7 = 1 THEN 'X' ELSE '' END AS i7,
        CASE WHEN A.incidente_8 = 1 THEN 'X' ELSE '' END AS i8, */
        if (!incidentes.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            })
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[0], 'Incidentes.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13, inc = [0, 0, 0, 0, 0];
                worksheet.spliceColumns(1, 1);
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:N2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:N3');
                worksheet.getCell('A5').value = 'PENDIENTE';
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:N5');
                worksheet.getCell('A6').value = 'INCIDENTES PRESENTADOS DURANTE LA VALIDACIÓN DE LA ELECCIÓN Y LA CONSULTA';
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:N6');
                worksheet.getCell('A8').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('A8').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('M8').value = `Fecha: ${fecha}`;
                worksheet.getCell('M9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                if (!worksheet.getCell('A11').isMerged)
                    worksheet.mergeCells('A11:A12');
                if (!worksheet.getCell('B11').isMerged)
                    worksheet.mergeCells('B11:B12');
                if (!worksheet.getCell('C11').isMerged)
                    worksheet.mergeCells('C11:C12');
                if (!worksheet.getCell('D11').isMerged)
                    worksheet.mergeCells('D11:D12');
                if (!worksheet.getCell('E11').isMerged)
                    worksheet.mergeCells('E11:E12');
                worksheet.getCell('F11').value = 'CAUSAS';
                worksheet.getCell('F11').style = fill;
                if (!worksheet.getCell('F11').isMerged)
                    worksheet.mergeCells('F11:J11');
                if (!worksheet.getCell('K11').isMerged)
                    worksheet.mergeCells('K11:K12');
                if (!worksheet.getCell('L11').isMerged)
                    worksheet.mergeCells('L11:L12');
                if (!worksheet.getCell('M11').isMerged)
                    worksheet.mergeCells('M11:M12');
                if (!worksheet.getCell('N11').isMerged)
                    worksheet.mergeCells('N11:N12');
                incidentes.forEach(incidente => {
                    Object.keys(incidente).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = incidente[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                        if (['i1', 'i2', 'i3', 'i4', 'i5'].includes(key))
                            inc[+key.replace('i', '') - 1] += incidente[key] == 'X' ? 1 : 0;
                    });
                    fila++;
                });
                for (let i = 0; i <= 4; i++) {
                    worksheet.getCell(fila, i + 6).value = inc[i];
                    worksheet.getCell(fila, i + 6).style = { ...contenidoStyle, numFmt: '#,##0' };
                }
                fila++;
                worksheet.getCell(fila, 6).value = 'Total de Causas de Incidentes:'
                worksheet.getCell(fila, 6).style = fill;
                worksheet.getCell(fila, 7).value = inc.reduce((sum, i) => sum + i, 0);
                worksheet.getCell(fila, 7).style = { ...contenidoStyle, numFmt: '#,##0' };
                worksheet.columns.forEach((column, i) => {
                    if ([0, 2, 9, 11].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 12)
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
                    reporte: `Reporte_IncidentesDistrito-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en IncidentesDistrito: ${err}`);
                ;
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en IncidentesDistrito: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}