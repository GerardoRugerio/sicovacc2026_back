import ExcelJs from 'exceljs';
import { request, response } from 'express';
import path from 'path';
import { autor, contenidoStyle, fill, IECMLogo, plantillas, titulos } from '../helpers/Constantes.js';
import { FechaServer } from '../helpers/Consultas.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Cómputo total de las Candidaturas por UT

export const ComputoTotalUT = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT DISTINCT id_distrito, clave_colonia
            FROM copaco_actas
            WHERE modalidad = 1${id_distrito != 0 ? ` AND id_distrito = ${id_distrito}` : ''}
        ),
        Info AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, nombre, paterno, materno, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
            FROM copaco_actas_VVS
            WHERE estatus = 1
            GROUP BY id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, nombre, paterno, materno
        )
        SELECT A.id_distrito, nombre_delegacion, A.clave_colonia, nombre_colonia, dbo.NumeroALetras(secuencial) AS secuencial, nombre, paterno, materno, votos, votos_sei, total_votos
        FROM CA A
        LEFT JOIN Info I ON A.id_distrito = I.id_distrito AND A.clave_colonia = I.clave_colonia
        ORDER BY A.id_distrito, nombre_delegacion, nombre_colonia, I.secuencial ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[1], 'Computo_Total_UT.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'CÓMPUTO TOTAL DE LAS CANDIDATURAS POR UNIDADES TERRITORIALES (INCLUYE MRVyO, MECPEP, MECPPP Y SEI)';
                worksheet.getCell('J8').value = `Fecha: ${fecha}`;
                worksheet.getCell('J9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:K2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:K3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:K5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:K6');
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
                if (!worksheet.getCell('F11').isMerged)
                    worksheet.mergeCells('F11:H11');
                if (!worksheet.getCell('I11').isMerged)
                    worksheet.mergeCells('I11:I12');
                if (!worksheet.getCell('J11').isMerged)
                    worksheet.mergeCells('J11:J12');
                if (!worksheet.getCell('K11').isMerged)
                    worksheet.mergeCells('K11:K12');
                actas.forEach(acta => {
                    Object.keys(acta).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = acta[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    })
                    fila++;
                });
                worksheet.columns.forEach((column, i) => {
                    if ([1, 3].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 12)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
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
                    reporte: `Reporte_ComputoTotalUT-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ComputoTotalUT: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ComputoTotalUT: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Resultados del Cómputo Total por Mesa

export const ResultadoComputoTotalMesa = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(A.num_mro, NULLIF(CONCAT(' ', TP.mesa ), '')) AS mesa, A.num_mro, A.tipo_mro, A.levantada_distrito,
            A.bol_sobrantes, A.bol_recibidas, A.total_ciudadanos, A.bol_nulas, A.votacion_total_emitida, A.modalidad
            FROM copaco_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            INNER JOIN consulta_tipo_mesa_V TP ON A.tipo_mro = TP.tipo_mro
            WHERE A.estatus = 1${id_distrito != 0 ? ` AND A.id_distrito = ${id_distrito}` : ''}
        ),
        LD AS (
            SELECT id_distrito, clave_colonia, num_mro, tipo_mro,
            SUM(CASE WHEN levantada_distrito > 0 AND bol_recibidas = 0 THEN 0 ELSE total_ciudadanos END) AS ciudadania,
            SUM(CASE WHEN levantada_distrito > 0 AND bol_recibidas = 0 THEN total_ciudadanos ELSE 0 END) AS distrito
            FROM CA
            GROUP BY id_distrito, clave_colonia, num_mro, tipo_mro
        ),
        MesasEsperadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE estatus = 1
            GROUP BY id_distrito, clave_colonia
        ),
        MesasCapturadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS capturadas
            FROM CA
            WHERE modalidad = 1
            GROUP BY id_distrito, clave_colonia
        ),
        Mesas AS (
            SELECt C.id_distrito, C.clave_colonia
            FROM MesasCapturadas C
            INNER JOIN MesasEsperadas E ON C.id_distrito = E.id_distrito AND C.clave_colonia = E.clave_colonia
            WHERE C.capturadas = E.total
        ),
        ParticipantesJSON AS (
            SELECT id_distrito, clave_colonia, num_mro, tipo_mro, (
                SELECT secuencial, votos, votos_sei, total_votos
                FROM copaco_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.num_mro = V1.num_mro AND V2.tipo_mro = V1.tipo_mro
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS participantes
            FROM copaco_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, num_mro, tipo_mro
        )
        SELECT A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, A1.mesa, A1.bol_sobrantes, LD.ciudadania, LD.distrito, P.participantes, A1.bol_nulas, COALESCE(A2.bol_nulas, 0) AS bol_nulas_sei, A1.bol_nulas + COALESCE(A2.bol_nulas, 0) AS total_nulas,
        A1.votacion_total_emitida, COALESCE(A2.votacion_total_emitida, 0) AS votacion_total_emitida_sei, A1.votacion_total_emitida + COALESCE(A2.votacion_total_emitida, 0) AS votacion_total
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN LD ON A1.id_distrito = LD.id_distrito AND A1.clave_colonia = LD.clave_colonia AND A1.num_mro = LD.num_mro AND A1.tipo_mro = LD.tipo_mro
        LEFT JOIN ParticipantesJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.num_mro = P.num_mro AND A1.tipo_mro = P.tipo_mro 
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_colonia = A1.clave_colonia)
        ORDER BY A1.id_distrito, A1.nombre_delegacion, A1.nombre_colonia, A1.num_mro, A1.tipo_mro ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const max = Math.max(...actas.map(acta => JSON.parse(acta.participantes).length));
        workbook.xlsx.readFile(path.join(plantillas[1], 'Resultado_Computo_Total_Mesa.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const celdasTotales = 14 + (max * 3);
                let fila = 13, celda = 9;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'RESULTADOS DEL CÓMPUTO TOTAL POR MESA (INCLUYE MRVyO, MECPEP, MECPPP Y SEI)';
                worksheet.getCell('M8').value = `Fecha: ${fecha}`;
                worksheet.getCell('M9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                for (let i = 1; i <= max; i++) {
                    for (let j = 1; j <= 3; j++)
                        worksheet.spliceColumns(celda, 0, [null]);
                    if (!worksheet.getCell(11, celda).isMerged)
                        worksheet.mergeCells(11, celda, 11, celda + 2);
                    for (let j = celda; j <= celda + 2; j++)
                        worksheet.getCell(11, j).style = contenidoStyle;
                    worksheet.getCell(11, celda).value = i;
                    worksheet.getCell(11, celda).style = fill;
                    worksheet.getCell(12, celda).value = 'Opiniones Mesa';
                    worksheet.getCell(12, celda).style = fill;
                    worksheet.getCell(12, celda + 1).value = 'Opiniones (SEI)';
                    worksheet.getCell(12, celda + 1).style = fill;
                    worksheet.getCell(12, celda + 2).value = `Total de Opiniones Participante ${i}`;
                    worksheet.getCell(12, celda + 2).style = fill;
                    celda += 3;
                }
                if (!worksheet.getCell(2, 1).isMerged)
                    worksheet.mergeCells(2, 1, 2, celdasTotales);
                if (!worksheet.getCell(3, 1).isMerged)
                    worksheet.mergeCells(3, 1, 3, celdasTotales);
                if (!worksheet.getCell(5, 1).isMerged)
                    worksheet.mergeCells(5, 1, 5, celdasTotales);
                if (!worksheet.getCell(6, 1).isMerged)
                    worksheet.mergeCells(6, 1, 6, celdasTotales);
                if (!worksheet.getCell(11, 1).isMerged)
                    worksheet.mergeCells(11, 1, 12, 1);
                if (!worksheet.getCell(11, 2).isMerged)
                    worksheet.mergeCells(11, 2, 12, 2);
                if (!worksheet.getCell(11, 3).isMerged)
                    worksheet.mergeCells(11, 3, 12, 3);
                if (!worksheet.getCell(11, 4).isMerged)
                    worksheet.mergeCells(11, 4, 12, 4);
                if (!worksheet.getCell(11, 5).isMerged)
                    worksheet.mergeCells(11, 5, 12, 5);
                if (!worksheet.getCell(11, 6).isMerged)
                    worksheet.mergeCells(11, 6, 12, 6);
                if (!worksheet.getCell(11, 7).isMerged)
                    worksheet.mergeCells(11, 7, 12, 7);
                if (!worksheet.getCell(11, 8).isMerged)
                    worksheet.mergeCells(11, 8, 12, 8);
                if (!worksheet.getCell(11, 8 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 8 + (max * 3), 12, 8 + (max * 3));
                if (!worksheet.getCell(11, 9 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 9 + (max * 3), 12, 9 + (max * 3));
                if (!worksheet.getCell(11, 10 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 10 + (max * 3), 12, 10 + (max * 3));
                if (!worksheet.getCell(11, 11 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 11 + (max * 3), 12, 11 + (max * 3));
                if (!worksheet.getCell(11, 12 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 12 + (max * 3), 12, 12 + (max * 3));
                if (!worksheet.getCell(11, 13 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 13 + (max * 3), 12, 13 + (max * 3));
                if (!worksheet.getCell(11, 14 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 14 + (max * 3), 12, 14 + (max * 3));
                const imprimir = (index, text) => {
                    worksheet.getCell(fila, index).value = text;
                    worksheet.getCell(fila, index).style = index > 4 && index < celdasTotales + 1 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                }
                const imprimirParticipantes = (index, participantes) => {
                    let i = index;
                    participantes.forEach(participante => {
                        Object.entries(participante).forEach(([campo, valor]) => {
                            if (!campo.includes('secuencial')) {
                                imprimir(i, valor);
                                i++;
                            }
                        });
                    })
                    return i;
                }
                actas.forEach(acta => {
                    let i = 1;
                    Object.entries(acta).forEach(([campo, valor]) => {
                        if (!campo.match('participantes')) {
                            imprimir(i, valor);
                            i++;
                            return;
                        }
                        i = imprimirParticipantes(i, JSON.parse(valor));
                        const faltantes = max - JSON.parse(valor).length;
                        for (let x = 0; x < faltantes * 3; x++) {
                            imprimir(i, '');
                            i++;
                        }
                    });
                    fila++;
                });
                worksheet.columns.forEach((column, i) => {
                    if ([1, 3].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 11)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
                        else
                            column.width = maxLength;
                    }
                    if (i >= 8 && i <= 8 + (max * 3))
                        column.width = 15;
                });
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ResultadosComputoTotalMesa-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ResultadoComputoTotalMesa: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ResultadoComputoTotalMesa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Resultados del Cómputo total por Unidad Territorial

export const ResultadoComputoTotalUT = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, A.levantada_distrito,
            A.bol_sobrantes, A.bol_recibidas, A.total_ciudadanos, A.bol_nulas, A.votacion_total_emitida, A.modalidad
            FROM copaco_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            WHERE A.estatus = 1${id_distrito != 0 ? ` AND A.id_distrito = ${id_distrito}` : ''}
        ),
        LD AS (
            SELECT id_distrito, clave_colonia,
            SUM(CASE WHEN levantada_distrito > 0 AND bol_recibidas = 0 THEN 0 ELSE total_ciudadanos END) AS ciudadania,
            SUM(CASE WHEN levantada_distrito > 0 AND bol_recibidas = 0 THEN total_ciudadanos ELSE 0 END) AS distrito
            FROM CA
            GROUP BY id_distrito, clave_colonia
        ),
        MesasEsperadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE estatus = 1
            GROUP BY id_distrito, clave_colonia
        ),
        MesasCapturadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS capturadas
            FROM CA
            WHERE modalidad = 1
            GROUP BY id_distrito, clave_colonia
        ),
        Mesas AS (
            SELECt C.id_distrito, C.clave_colonia
            FROM MesasCapturadas C
            INNER JOIN MesasEsperadas E ON C.id_distrito = E.id_distrito AND C.clave_colonia = E.clave_colonia
            WHERE C.capturadas = E.total
        ),
        ParticipantesJSON AS (
            SELECT id_distrito, clave_colonia, (
                SELECT secuencial, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
                FROM copaco_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia
                GROUP BY secuencial
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS participantes
            FROM copaco_actas_VVS V1
            GROUP BY id_distrito, clave_colonia
        )
        SELECT A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, LD.ciudadania, LD.distrito, P.participantes, SUM(A1.bol_nulas) AS bol_nulas, SUM(COALESCE(A2.bol_nulas, 0)) AS bol_nulas_sei, SUM(A1.bol_nulas + COALESCE(A2.bol_nulas, 0)) AS total_nulas,
        SUM(A1.votacion_total_emitida) AS votacion_total_emitida, SUM(COALESCE(A2.votacion_total_emitida, 0)) AS votacion_total_emitida_sei, SUM(A1.votacion_total_emitida + COALESCE(A2.votacion_total_emitida, 0)) AS votacion_total
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A2.modalidad = 2
        LEFT JOIN LD ON A1.id_distrito = LD.id_distrito AND A1.clave_colonia = LD.clave_colonia
        LEFT JOIN ParticipantesJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_colonia = A1.clave_colonia)
        GROUP BY A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, LD.ciudadania, LD.distrito, P.participantes
        ORDER BY A1.id_distrito, A1.nombre_delegacion, A1.nombre_colonia ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const max = Math.max(...actas.map(acta => JSON.parse(acta.participantes).length));
        workbook.xlsx.readFile(path.join(plantillas[1], 'Resultado_Computo_Total_UT.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const celdasTotales = 12 + (max * 3);
                let fila = 13, celda = 7;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'RESULTADOS DEL CÓMPUTO TOTAL POR UNIDAD TERRITORIAL (INCLUYE MRVyO, MECPEP, MECPPP Y SEI)';
                worksheet.getCell('K8').value = `Fecha: ${fecha}`;
                worksheet.getCell('K9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                for (let i = 1; i <= max; i++) {
                    for (let j = 1; j <= 3; j++)
                        worksheet.spliceColumns(celda, 0, [null]);
                    if (!worksheet.getCell(11, celda).isMerged)
                        worksheet.mergeCells(11, celda, 11, celda + 2);
                    for (let j = celda; j <= celda + 2; j++)
                        worksheet.getCell(11, j).style = contenidoStyle;
                    worksheet.getCell(11, celda).value = i;
                    worksheet.getCell(11, celda).style = fill;
                    worksheet.getCell(12, celda).value = 'Opiniones Mesa';
                    worksheet.getCell(12, celda).style = fill;
                    worksheet.getCell(12, celda + 1).value = 'Opiniones (SEI)';
                    worksheet.getCell(12, celda + 1).style = fill;
                    worksheet.getCell(12, celda + 2).value = `Total de Opiniones Participante ${i}`;
                    worksheet.getCell(12, celda + 2).style = fill;
                    celda += 3;
                }
                if (!worksheet.getCell(2, 1).isMerged)
                    worksheet.mergeCells(2, 1, 2, celdasTotales);
                if (!worksheet.getCell(3, 1).isMerged)
                    worksheet.mergeCells(3, 1, 3, celdasTotales);
                if (!worksheet.getCell(5, 1).isMerged)
                    worksheet.mergeCells(5, 1, 5, celdasTotales);
                if (!worksheet.getCell(6, 1).isMerged)
                    worksheet.mergeCells(6, 1, 6, celdasTotales);
                if (!worksheet.getCell(11, 1, 12, 1).isMerged)
                    worksheet.mergeCells(11, 1, 12, 1);
                if (!worksheet.getCell(11, 2, 12, 1).isMerged)
                    worksheet.mergeCells(11, 2, 12, 2);
                if (!worksheet.getCell(11, 3, 12, 1).isMerged)
                    worksheet.mergeCells(11, 3, 12, 3);
                if (!worksheet.getCell(11, 4, 12, 1).isMerged)
                    worksheet.mergeCells(11, 4, 12, 4);
                if (!worksheet.getCell(11, 5, 12, 1).isMerged)
                    worksheet.mergeCells(11, 5, 12, 5);
                if (!worksheet.getCell(11, 6, 12, 1).isMerged)
                    worksheet.mergeCells(11, 6, 12, 6);
                if (!worksheet.getCell(11, 6 + (max * 3), 12, 6 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 6 + (max * 3), 12, 6 + (max * 3));
                if (!worksheet.getCell(11, 7 + (max * 3), 12, 7 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 7 + (max * 3), 12, 7 + (max * 3));
                if (!worksheet.getCell(11, 8 + (max * 3), 12, 8 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 8 + (max * 3), 12, 8 + (max * 3));
                if (!worksheet.getCell(11, 9 + (max * 3), 12, 9 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 9 + (max * 3), 12, 9 + (max * 3));
                if (!worksheet.getCell(11, 10 + (max * 3), 12, 10 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 10 + (max * 3), 12, 10 + (max * 3));
                if (!worksheet.getCell(11, 11 + (max * 3), 12, 11 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 11 + (max * 3), 12, 11 + (max * 3));
                if (!worksheet.getCell(11, 12 + (max * 3), 12, 12 + (max * 3)).isMerged)
                    worksheet.mergeCells(11, 12 + (max * 3), 12, 12 + (max * 3));
                const imprimir = (index, text) => {
                    worksheet.getCell(fila, index).value = text;
                    worksheet.getCell(fila, index).style = index > 3 && index < celdasTotales + 1 ? { ...contenidoStyle, numFmt: "#,##0" } : contenidoStyle;
                }
                const imprimirParticipantes = (index, participantes) => {
                    let i = index;
                    participantes.forEach(participante => {
                        Object.entries(participante).forEach(([campo, valor]) => {
                            if (!campo.includes('secuencial')) {
                                imprimir(i, valor);
                                i++;
                            }
                        });
                    });
                    return i;
                }
                actas.forEach(acta => {
                    let i = 1;
                    Object.entries(acta).forEach(([campo, valor]) => {
                        if (!campo.match('participantes')) {
                            imprimir(i, valor);
                            i++;
                            return;
                        }
                        i = imprimirParticipantes(i, JSON.parse(valor));
                        const faltantes = max - JSON.parse(valor).length;
                        for (let x = 0; x < faltantes * 3; x++) {
                            imprimir(i, '');
                            i++;
                        }
                    });
                    fila++;
                });
                worksheet.columns.forEach((column, i) => {
                    if ([1, 3].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 11)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
                        else
                            column.width = maxLength;
                    }
                    if (i >= 6 && i <= 6 + (max * 3))
                        column.width = 15;
                });
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ResultadosComputoTotalUT-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ResultadoComputoTotalUT: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ResultadoComputoTotalUT: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Concentrado de Candidaturas Participantes

export const ConcentradoParticipantes = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const participantes = (await SICOVACC.sequelize.query(`SELECT F.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, F.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, COUNT(*) AS total
        FROM copaco_formulas F
        LEFT JOIN consulta_cat_delegacion D ON F.id_delegacion = D.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 C ON F.clave_colonia = C.clave_colonia
        WHERE F.secuencial IS NOT NULL${id_distrito != 0 ? ` AND F.id_distrito = ${id_distrito}` : ''}
        GROUP BY F.id_distrito, D.nombre_delegacion, F.clave_colonia, C.nombre_colonia
        ORDER BY F.id_distrito, D.nombre_delegacion, C.nombre_colonia`))[0];
        if (!participantes.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[1], 'Concentrado_Participantes.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'CONCENTRADO CANDIDATURAS PARTICIPANTES';
                worksheet.getCell('E8').value = `Fecha: ${fecha}`;
                worksheet.getCell('E9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:E2')
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:E3')
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:E5')
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:E6')
                participantes.forEach(participante => {
                    Object.keys(participante).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = participante[key];
                        worksheet.getCell(fila, index + 1).style = index == 3 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 4).value = 'TOTAL';
                worksheet.getCell(fila, 4).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 5).value = participantes.reduce((sum, participante) => sum + participante.total, 0);
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ConcentradoParticipantes-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ConcentradoParticipantes: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ConcentradoParticipantes: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Candidaturas en las que se presenta empate

export const CandidaturasEmpate = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const candidatos = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia
            FROM copaco_actas
            ${id_distrito != 0 ? `WHERE id_distrito = ${id_distrito}` : ''}
        ),
        MesasEsperadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE estatus = 1
            GROUP BY id_distrito, clave_colonia
        ),
        MesasCapturadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS capturadas
            FROM CA
            GROUP BY id_distrito, clave_colonia
        ),
        Mesas AS (
            SELECT C.id_distrito, C.clave_colonia
            FROM MesasCapturadas C
            INNER JOIN MesasEsperadas E ON C.id_distrito = E.id_distrito AND C.clave_colonia = E.clave_colonia
            WHERE C.capturadas = E.total
        ),
        ActasValidadas AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, nombreC, total_votos
            FROM copaco_actas_VVS V
            WHERE EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = V.id_distrito AND clave_Colonia = V.clave_colonia) AND estatus = 1
        ),
        Votos AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, nombreC, SUM(total_votos) AS total_votos
            FROM ActasValidadas
            GROUP BY id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, nombreC
        ),
        RANKING AS (
            SELECT *, DENSE_RANK() OVER (PARTITION BY clave_Colonia ORDER BY total_votos DESC) AS DR, COUNT(*) OVER (PARTITION BY clave_colonia, total_votos) AS empate
            FROM Votos
        )
        SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, nombreC, total_votos
        FROM RANKING
        WHERE DR IN (1, 2) AND empate > 1 AND total_votos > 0
        ORDER BY id_distrito, nombre_delegacion, nombre_colonia`))[0];
        if (!candidatos.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[1], 'Candidaturas_Empate.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 1 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'CANDIDATURAS EN LAS QUE SE PRESENTA EMPATE';
                worksheet.getCell('E8').value = `Fecha: ${fecha}`;
                worksheet.getCell('E9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:F2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:F3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:F5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:F6');
                candidatos.forEach(candidato => {
                    Object.keys(candidato).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = candidato[key];
                        worksheet.getCell(fila, index + 1).style = index == 4 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                });
                worksheet.columns.forEach((column, i) => {
                    if ([1, 3, 4].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 10)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
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
                    reporte: `Reporte_CandidaturasEmpate-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en CandidaturasEmpate: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en CandidaturasEmpate: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Resultados de Votos por Mesa

export const ResultadosMesa = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(A.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa, A.num_mro, A.tipo_mro, A.modalidad, A.bol_nulas
            FROM copaco_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            INNER JOIN consulta_tipo_mesa_V TP ON A.tipo_mro = TP.tipo_mro
            WHERE A.estatus = 1${id_distrito != 0 ? ` AND A.id_distrito = ${id_distrito}` : ''}
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE estatus = 1
            GROUP BY id_distrito, clave_colonia
        ),
        MesasCapturadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS capturadas
            FROM CA
            WHERE modalidad = 1
            GROUP BY id_distrito, clave_colonia
        ),
        Mesas AS (
            SELECt C.id_distrito, C.clave_colonia
            FROM MesasCapturadas C
            INNER JOIN MesasEsperadas E ON C.id_distrito = E.id_distrito AND C.clave_colonia = E.clave_colonia
            WHERE C.capturadas = E.total
        ),
        ProyectosJSON AS (
            SELECT id_distrito, clave_colonia, num_mro, tipo_mro, (
                SELECT dbo.NumeroALetras(secuencial) AS secuencial, nombreC, votos, votos_sei, total_votos
                FROM copaco_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.num_mro = V1.num_mro AND V2.tipo_mro = V1.tipo_mro
                ORDER BY V2.secuencial ASC
                FOR JSON PATH
            ) AS participantes
            FROM copaco_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, num_mro, tipo_mro
        )
        SELECT A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, A1.mesa, P.participantes, A1.bol_nulas, COALESCE(A2.bol_nulas, 0) AS bol_nulas_sei
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.num_mro = P.num_mro AND A1.tipo_mro = P.tipo_mro 
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_colonia = A1.clave_colonia)
        ORDER BY A1.id_distrito, A1.nombre_delegacion, A1.nombre_colonia, A1.num_mro, A1.tipo_mro`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[0], 'Resultados_Opi_Mesa.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:J2');
                worksheet.getCell('A3').value = titulos[1];
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:J3');
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:J5');
                worksheet.getCell('A6').value = 'RESULTADOS DE VOTOS POR MESA (INCLUYE MRVyO, MECPEP, MECPPP, SEI)';
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:J6');
                worksheet.getCell('J8').value = fecha;
                worksheet.getCell('J9').value = hora.substring(0, hora.length - 3);
                worksheet.getCell('E11').value = 'Letra del Participante';
                worksheet.getCell('G11').value = 'Nombre del Participante';
                for (let acta of actas) {
                    let sum_votos = 0, sum_votos_sei = 0;
                    const { id_distrito: distrito, nombre_delegacion, clave_colonia, nombre_colonia, mesa, participantes, bol_nulas, bol_nulas_sei } = acta;
                    sum_votos += bol_nulas, sum_votos_sei += bol_nulas_sei;
                    for (let participante of JSON.parse(participantes)) {
                        const { secuencial, nombreC, votos, votos_sei, total_votos } = participante;
                        sum_votos += votos, sum_votos_sei += votos_sei;
                        const X = { distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, mesa, nombreC, votos, votos_sei, total_votos };
                        Object.keys(X).forEach((key, i) => {
                            worksheet.getCell(fila, i + 1).value = X[key];
                            worksheet.getCell(fila, i + 1).style = [4, 7, 8, 9].includes(i) ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                        });
                        fila++;
                    }
                    for (let i = 1; i <= 10; i++) {
                        worksheet.getCell(fila, i).style = i <= 6 ? contenidoStyle : i == 7 ? { ...fill, font: { ...fill.font, bold: false } } : { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                        worksheet.getCell(fila + 1, i).style = i <= 6 ? contenidoStyle : i == 7 ? { ...fill, font: { ...fill.font, bold: false } } : { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                    }
                    worksheet.getCell(fila, 7).value = 'Opiniones Nulas';
                    worksheet.getCell(fila, 8).value = bol_nulas;
                    worksheet.getCell(fila, 9).value = bol_nulas_sei;
                    worksheet.getCell(fila, 10).value = bol_nulas + bol_nulas_sei;
                    worksheet.getCell(fila + 1, 7).value = 'Total por Mesa';
                    worksheet.getCell(fila + 1, 8).value = sum_votos;
                    worksheet.getCell(fila + 1, 9).value = sum_votos_sei;
                    worksheet.getCell(fila + 1, 10).value = sum_votos + sum_votos_sei;
                    fila += 2;
                }
                worksheet.columns.forEach((column, i) => {
                    if ([1, 3, 6].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 8)
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
                    reporte: `Reporte_ResultadoMesa-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ResultadoMesa: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ResultadoMesa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Concentrado de Mesas Computadas

export const MesasComputadas = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const mesas = (await SICOVACC.sequelize.query(`SELECT M.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, M.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(M.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa
        FROM consulta_mros M
        LEFT JOIN consulta_cat_delegacion D ON M.id_delegacion = D.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 C ON M.clave_colonia = C.clave_colonia
        LEFT JOIN consulta_tipo_mesa_V TP ON M.tipo_mro = TP.tipo_mro
        WHERE M.estatus = 1 AND EXISTS (SELECT 1 FROM copaco_actas WHERE modalidad = 1 AND estatus = 1 AND clave_colonia = M.clave_colonia AND num_mro = M.num_mro AND tipo_mro = M.tipo_mro)${id_distrito != 0 ? ` AND M.id_distrito = ${id_distrito}` : ''}
        ORDER BY M.id_distrito, D.nombre_delegacion, C.nombre_colonia, M.num_mro, M.tipo_mro`))[0];
        if (!mesas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[1], 'Mesas_Computadas.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'CONCENTRADO DE MESAS COMPUTADAS (INCLUYE MRVyO, MECPEP, MECPPP Y SEI)';
                worksheet.getCell('D8').value = `Fecha: ${fecha}`;
                worksheet.getCell('D9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:E2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:E3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:E5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:E6');
                mesas.forEach(mesa => {
                    Object.keys(mesa).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = mesa[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 3).value = 'TOTAL';
                worksheet.getCell(fila, 3).style = fill;
                worksheet.getCell(fila, 4).value = mesas.length;
                worksheet.getCell(fila, 4).style = { ...fill, numFmt: '#,##0' };
                worksheet.columns.forEach((column, i) => {
                    if ([1, 3].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 10)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
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
                    reporte: `Reporte_MesasComputadas-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en MesasComputadas: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en MesasComputadas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Concentrado de Mesas que no han sido Computadas

export const MesasNoComputadas = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const mesas = (await SICOVACC.sequelize.query(`SELECT M.id_distrito, M.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(M.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa
        FROM consulta_mros M
        LEFT JOIN consulta_cat_colonia_cc1 C ON M.clave_colonia = C.clave_colonia
        LEFT JOIN consulta_tipo_mesa_V TP ON M.tipo_mro = TP.tipo_mro
        WHERE M.estatus = 1 AND NOT EXISTS (SELECT 1 FROM copaco_actas WHERE modalidad = 1 AND estatus = 1 AND clave_colonia = M.clave_colonia AND num_mro = M.num_mro AND tipo_mro = M.tipo_mro)${id_distrito != 0 ? ` AND M.id_distrito = ${id_distrito}` : ''}
        ORDER BY M.id_distrito, C.nombre_colonia, M.num_mro, M.tipo_mro`))[0];
        if (!mesas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[1], 'Mesas_No_Computadas.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'CONCENTRADO DE MESAS COMPUTADAS (INCLUYE MRVyO, MECPEP, MECPPP Y SEI)';
                worksheet.getCell('D8').value = `Fecha: ${fecha}`;
                worksheet.getCell('D9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:D2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:D3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:D5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:D6');
                mesas.forEach(mesa => {
                    Object.keys(mesa).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = mesa[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 3).value = 'TOTAL';
                worksheet.getCell(fila, 3).style = fill;
                worksheet.getCell(fila, 4).value = mesas.length;
                worksheet.getCell(fila, 4).style = { ...fill, numFmt: '#,##0' };
                worksheet.columns.forEach((column, i) => {
                    if ([2].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 10)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
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
                    reporte: `Reporte_MesasNoComputadas-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en MesasNoComputadas: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en MesasNoComputadas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Unidades Territoriales Con Cómputo Capturado

export const UTConComputo = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const utc = (await SICOVACC.sequelize.query(`SELECT C.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, C.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia
        FROM consulta_cat_colonia_cc1 C
        LEFT JOIN consulta_cat_delegacion D ON C.id_delegacion = D.id_delegacion
        WHERE${id_distrito != 0 ? ` C.id_distrito = ${id_distrito} AND` : ''} EXISTS (SELECT 1 FROM copaco_actas A WHERE modalidad = 1 AND estatus = 1 AND clave_colonia = C.clave_colonia GROUP BY clave_colonia HAVING COUNT(*) = (SELECT COUNT(*) FROM consulta_mros WHERE estatus = 1 AND clave_colonia = A.clave_colonia))
        AND EXISTS (SELECT 1 FROM consulta_mros WHERE id_distrito = C.id_distrito AND clave_colonia = C.clave_colonia AND estatus = 1)
        ORDER BY C.id_distrito, nombre_delegacion, nombre_colonia`))[0];
        if (!utc.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[1], 'UT_Computo.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'UNIDADES TERRITORIALES CON CÓMPUTO CAPTURADO';
                worksheet.getCell('D8').value = `Fecha: ${fecha}`;
                worksheet.getCell('D9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:D2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:D3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:D5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:D6');
                utc.forEach(ut => {
                    Object.keys(ut).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = ut[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 3).value = 'TOTAL';
                worksheet.getCell(fila, 3).style = fill;
                worksheet.getCell(fila, 4).value = utc.length;
                worksheet.getCell(fila, 4).style = { ...fill, numFmt: '#,##0' };
                worksheet.columns.forEach((column, i) => {
                    if ([1, 3].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 10)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
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
                    reporte: `Reporte_UTConComputo-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en UTConComputo: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en UTConComputo: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Unidades Territoriales Sin Cómputo Capturado

export const UTSinComputo = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const utc = (await SICOVACC.sequelize.query(`SELECT C.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, C.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia
        FROM consulta_cat_colonia_cc1 C
        LEFT JOIN consulta_cat_delegacion D ON C.id_delegacion = D.id_delegacion
        WHERE${id_distrito != 0 ? ` C.id_distrito = ${id_distrito} AND` : ''} NOT EXISTS (SELECT 1 FROM copaco_actas A WHERE modalidad = 1 AND estatus = 1 AND clave_colonia = C.clave_colonia GROUP BY clave_colonia HAVING COUNT(*) = (SELECT COUNT(*) FROM consulta_mros WHERE estatus = 1 AND clave_colonia = A.clave_colonia))
        AND EXISTS (SELECT 1 FROM consulta_mros WHERE id_distrito = C.id_distrito AND clave_colonia = C.clave_colonia AND estatus = 1)
        ORDER BY C.id_distrito, nombre_delegacion, nombre_colonia`))[0];
        if (!utc.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[1], 'UT_Computo.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'UNIDADES TERRITORIALES SIN CÓMPUTO CAPTURADO';
                worksheet.getCell('D8').value = `Fecha: ${fecha}`;
                worksheet.getCell('D9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:D2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:D3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:D5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:D6');
                utc.forEach(ut => {
                    Object.keys(ut).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = ut[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 3).value = 'TOTAL';
                worksheet.getCell(fila, 3).style = fill;
                worksheet.getCell(fila, 4).value = utc.length;
                worksheet.getCell(fila, 4).style = { ...fill, numFmt: '#,##0' };
                worksheet.columns.forEach((column, i) => {
                    if ([1, 3].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 10)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
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
                    reporte: `Reporte_UTSinComputo-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en UTSinComputo: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en UTSinComputo: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Concentrado de Unidades Territoriales por Distrito Electoral con Cómputo Capturado (Grado de Avance)

export const UTConComputoGA = async (req = request, res = response) => {
    const workbook = new ExcelJs.Workbook();
    try {
        const avances = (await SICOVACC.sequelize.query(`SELECT D.id_distrito, M.UTTotal, COALESCE(UTV.UTValidadas, 0) AS UTValidadas, ROUND((CAST(COALESCE(UTV.UTValidadas, 0) AS FLOAT) * 100) / CAST(M.UTTotal AS FLOAT), 2) AS avance
        FROM consulta_cat_distrito D
        LEFT JOIN (SELECT id_distrito, COUNT(DISTINCT clave_colonia) AS UTTotal FROM consulta_mros WHERE estatus = 1 GROUP BY id_distrito) AS M ON D.id_distrito = M.id_distrito
        LEFT JOIN (
            SELECT id_distrito, COUNT(*) AS UTValidadas
            FROM consulta_cat_colonia_cc1 C
            WHERE estatus_copaco = 1 AND clave_colonia IN (
                SELECT A.clave_colonia
                FROM (SELECT clave_colonia, COUNT(*) AS total FROM consulta_mros WHERE estatus = 1 GROUP BY clave_colonia) AS A
                LEFT JOIN (SELECT clave_colonia, COUNT(*) AS cantidad FROM copaco_actas WHERE modalidad = 1 AND estatus = 1 GROUP BY clave_colonia) AS B ON A.clave_colonia = B.clave_colonia
                WHERE A.total = B.cantidad
            )
            GROUP BY id_distrito
        ) AS UTV ON D.id_distrito = UTV.id_distrito
        ORDER BY D.id_distrito ASC`))[0];
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[0], 'UT_Avance.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13, UTT = 0, UTCT = 0, total = 0;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A4').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('F9').value = `Fecha: ${fecha}`;
                worksheet.getCell('F10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                avances.forEach(distrito => {
                    Object.keys(distrito).forEach((key, index) => {
                        worksheet.getCell(fila, index + 3).value = distrito[key];
                        worksheet.getCell(fila, index + 3).style = { ...contenidoStyle, numFmt: Number.isInteger(distrito[key]) ? '#,##0' : '#,##0.##' };
                        if (key.match('UTTotal'))
                            UTT += +distrito[key];
                        if (key.match('UTValidadas'))
                            UTCT += +distrito[key];
                    });
                    fila++;
                });
                total = (UTCT * 100) / UTT;
                for (let i = 3; i <= 5; i++)
                    worksheet.getCell(fila, i).style = i == 3 ? fill : { ...fill, numFmt: '#,##0' };
                worksheet.getCell(fila, 3).value = 'TOTAL';
                worksheet.getCell(fila, 4).value = UTT;
                worksheet.getCell(fila, 5).value = UTCT;
                worksheet.getCell(fila, 6).value = total;
                worksheet.getCell(fila, 6).style = { ...fill, numFmt: Number.isInteger(total) ? '#,##0' : '#,##0.##' };
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_UTConnComputo(Avance)-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en UTConComputoGA: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en UTConComputoGA: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Actas Levantadas en Dirección Distrital

export const LevantadaDistrito = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`SELECT A.id_distrito, UPPER(C.nombre_colonia) AS nombre_colonia, A.clave_colonia, CONCAT(A.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa, UPPER(dbo.RazonDistrital(A.razon_distrital )) AS razon_distrital
        FROM copaco_actas A
        LEFT JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
        LEFT JOIN consulta_tipo_mesa_V TP ON A.tipo_mro = TP.tipo_mro
        WHERE A.modalidad = 1 AND A.levantada_distrito = 1${id_distrito != 0 ? ` AND A.id_distrito = ${id_distrito}` : ''}
        ORDER BY A.id_distrito, C.nombre_colonia, A.num_mro , A.tipo_mro`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[1], 'Levantada_Distrito.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 1 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'ACTAS LEVANTADAS EN DIRECCIÓN DISTRITAL (CAUSALES DE RECUENTO)';
                worksheet.getCell('E8').value = `Fecha: ${fecha}`;
                worksheet.getCell('E9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:E2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:E3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:E5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:E6');
                actas.forEach(acta => {
                    Object.keys(acta).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = acta[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 4).value = 'TOTAL';
                worksheet.getCell(fila, 4).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 5).value = actas.length;
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.columns.forEach((column, i) => {
                    if ([1, 4].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 10)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
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
                    reporte: `Reporte_LevantadaDistrito-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en LevantadaDistrito: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en LevantadaDistrito: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Votación Total por Distrito

export const VotacionDistrito = async (req = request, res = response) => {
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, SUM(votacion_total_emitida) AS total_votos, SUM(bol_nulas) AS total_nulas, modalidad
            FROM copaco_actas
            WHERE estatus = 1
            GROUP BY id_distrito, modalidad
        )
        SELECT D.id_distrito, COALESCE(A1.total_votos - A1.total_nulas, 0) AS total_votos, COALESCE(A2.total_votos - A2.total_nulas, 0) AS total_votos_sei, COALESCE(A1.total_nulas, 0) AS total_nulas, COALESCE(A2.total_nulas, 0) AS total_nulas_sei
        FROM consulta_cat_distrito D
        LEFT JOIN CA A1 ON D.id_distrito = A1.id_distrito AND A1.modalidad = 1
        LEFT JOIN CA A2 ON D.id_distrito = A2.id_distrito AND A2.modalidad = 2
        ORDER BY D.id_distrito ASC`))[0];
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[0], 'Opiniones_Distrito.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                let sum1 = 0, sum2 = 0, sum3 = 0, sum4 = 0, sum5 = 0;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 2 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'VOTACIÓN TOTAL POR DISTRITO';
                worksheet.getCell('F8').value = `Fecha: ${fecha}`;
                worksheet.getCell('F9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                for (let res of resultado) {
                    const { id_distrito, total_votos, total_votos_sei, total_nulas, total_nulas_sei } = res;
                    const total = total_votos + total_votos_sei + total_nulas + total_nulas_sei;
                    sum1 += total_votos, sum2 += total_votos_sei, sum3 += total_nulas, sum4 += total_nulas_sei, sum5 += total;
                    for (let i = 1; i <= 6; i++)
                        worksheet.getCell(fila, i).style = { ...contenidoStyle, numFmt: '#,##0' };
                    worksheet.getCell(fila, 1).value = id_distrito;
                    worksheet.getCell(fila, 2).value = total_votos;
                    worksheet.getCell(fila, 3).value = total_votos_sei;
                    worksheet.getCell(fila, 4).value = total_nulas;
                    worksheet.getCell(fila, 5).value = total_nulas_sei;
                    worksheet.getCell(fila, 6).value = total;
                    fila++
                }
                for (let i = 1; i <= 6; i++)
                    worksheet.getCell(fila, i).style = i == 1 ? fill : { ...fill, numFmt: '#,##0' };
                worksheet.getCell(fila, 1).value = 'Totales';
                worksheet.getCell(fila, 2).value = sum1;
                worksheet.getCell(fila, 3).value = sum2;
                worksheet.getCell(fila, 4).value = sum3;
                worksheet.getCell(fila, 5).value = sum4;
                worksheet.getCell(fila, 6).value = sum5;
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_VotacionDistrito-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en VotacionDistrito: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en VotacionDistrito: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Votación Total por Demarcación

export const VotacionDemarcacion = async (req = request, res = response) => {
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_delegacion, SUM(votacion_total_emitida) AS total_votos, SUM(bol_nulas) AS total_nulas, modalidad
            FROM copaco_actas
            WHERE estatus = 1
            GROUP BY id_delegacion, modalidad
        )
        SELECT UPPER(D.nombre_delegacion) AS nombre_delegacion, COALESCE(A1.total_votos - A1.total_nulas, 0) AS total_votos, COALESCE(A2.total_votos - A2.total_nulas, 0) AS total_votos_sei, COALESCE(A1.total_nulas, 0) AS total_nulas, COALESCE(A2.total_nulas, 0) AS total_nulas_sei
        FROM consulta_cat_delegacion D
        LEFT JOIN CA A1 ON D.id_delegacion = A1.id_delegacion AND A1.modalidad = 1
        LEFT JOIN CA A2 ON D.id_delegacion = A2.id_delegacion AND A2.modalidad = 2
        ORDER BY D.nombre_delegacion ASC`))[0];
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[0], 'Opiniones_Distrito.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                let sum1 = 0, sum2 = 0, sum3 = 0, sum4 = 0, sum5 = 0;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 2 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('A6').value = 'OPINIONES POR DEMARCACIÓN';
                worksheet.getCell('F8').value = `Fecha: ${fecha}`;
                worksheet.getCell('F9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                for (let res of resultado) {
                    const { nombre_delegacion, total_votos, total_votos_sei, total_nulas, total_nulas_sei } = res;
                    const total = total_votos + total_votos_sei + total_nulas + total_nulas_sei;
                    sum1 += total_votos, sum2 += total_votos_sei, sum3 += total_nulas, sum4 += total_nulas_sei, sum5 += total;
                    for (let i = 1; i <= 6; i++)
                        worksheet.getCell(fila, i).style = i == 1 ? contenidoStyle : { ...contenidoStyle, numFmt: '#,##0' };
                    worksheet.getCell(fila, 1).value = nombre_delegacion;
                    worksheet.getCell(fila, 2).value = total_votos;
                    worksheet.getCell(fila, 3).value = total_votos_sei;
                    worksheet.getCell(fila, 4).value = total_nulas;
                    worksheet.getCell(fila, 5).value = total_nulas_sei;
                    worksheet.getCell(fila, 6).value = total;
                    fila++;
                }
                for (let i = 1; i <= 6; i++)
                    worksheet.getCell(fila, i).style = i == 1 ? fill : { ...fill, numFmt: '#,##0' };
                worksheet.getCell(fila, 1).value = 'Totales';
                worksheet.getCell(fila, 2).value = sum1;
                worksheet.getCell(fila, 3).value = sum2;
                worksheet.getCell(fila, 4).value = sum3;
                worksheet.getCell(fila, 5).value = sum4;
                worksheet.getCell(fila, 6).value = sum5;
                worksheet.columns.forEach((column, i) => {
                    if ([0].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 11)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 70)
                            column.width = 70;
                        else if (maxLength < 21)
                            column.width = 21;
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
                    reporte: `Reporte_VotacionDistrito-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en VotacionDistrito: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en VotacionDistrito: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Porcentaje Participación por Distrito

export const Participacion = async (req = request, res = response) => {
    const workbook = new ExcelJs.Workbook();
    try {
        const porcentajes = (await SICOVACC.sequelize.query(`SELECT id_distrito, ROUND((COALESCE(total_opiniones, 0) * 100) / lista_nominal, 2) AS porcentaje
        FROM (
            SELECT D.id_distrito, A.total_opiniones, LN.lista_nominal
            FROM consulta_cat_distrito D
            LEFT JOIN (SELECT id_distrito, CAST(SUM(votacion_total_emitida) AS FLOAT) AS total_opiniones FROM copaco_actas WHERE estatus = 1 GROUP BY id_distrito) AS A ON D.id_distrito = A.id_distrito
            LEFT JOIN (SELECT id_distrito, SUM(lista_nominal) AS lista_nominal FROM consulta_mros WHERE estatus = 1 GROUP BY id_distrito) AS LN ON D.id_distrito = LN.id_distrito
        ) AS A
        ORDER BY id_distrito ASC`))[0];
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[0], 'Reporte_Participacion.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(2);
                let fila = 12;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('C8').value = `Fecha: ${fecha}`;
                worksheet.getCell('C9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                porcentajes.forEach(distrito => {
                    Object.keys(distrito).forEach((key, index) => {
                        worksheet.getCell(fila, index + 2).value = distrito[key];
                        worksheet.getCell(fila, index + 2).style = index == 0 ? contenidoStyle : { ...contenidoStyle, numFmt: Number.isInteger(distrito[key]) ? '#,##0' : '#,##0.##' };
                    });
                    fila++;
                });
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_Participacion-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en Participacion: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en Participacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Actas Capturadas con Alertas

export const ActasAlerta = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`SELECT id_distrito, CONCAT(num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa, dbo.Incidente(id_incidencia) AS incidente
        FROM copaco_actas A
        LEFT JOIN consulta_tipo_mesa_V TP ON A.tipo_mro = TP.tipo_mro
        WHERE modalidad = 1 AND id_incidencia IS NOT NULL${id_distrito != 0 ? ` AND id_distrito = ${id_distrito}` : ''}
        ORDER BY id_distrito, num_mro, A.tipo_mro`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        workbook.xlsx.readFile(path.join(plantillas[1], 'Actas_Alerta.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 10;
                const iecm = workbook.addImage({ filename: IECMLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 1 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = 'ELECCIÓN DE COMISIONES DE PARTICIPACIÓN COMUNITARIA';
                worksheet.getCell('B6').value = 'ACTAS CAPTURADAS CON ALERTAS';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:D2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:D3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:D5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:D6');
                actas.forEach(acta => {
                    Object.keys(acta).forEach((key, index) => {
                        worksheet.getCell(fila, index + 2).value = acta[key];
                        worksheet.getCell(fila, index + 2).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 3).value = 'TOTAL';
                worksheet.getCell(fila, 3).style = fill;
                worksheet.getCell(fila, 4).value = actas.length;
                worksheet.getCell(fila, 4).style = { ...contenidoStyle, numFmt: '#,##0' };
                worksheet.columns.forEach((column, i) => {
                    if ([3].includes(i)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, j) => {
                            if (j >= 8)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 10;
                        if (maxLength > 80)
                            column.width = 80;
                        else if (maxLength < 40)
                            column.width = 40;
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
                    reporte: `Reporte_ActasAlerta-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ActasAlerta: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ActasAlerta: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}