import ExcelJs from 'exceljs';
import { request, response } from 'express';
import path from 'path';
import { aniosCAT, autor, contenidoStyle, fill, iecmLogo, plantillas, titulos } from '../helpers/Constantes.js';
import { ConsultaTipoEleccion, FechaServer } from '../helpers/Consultas.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? F1 - Base de Datos

export const BaseDatos = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT UPPER(D.nombre_delegacion) AS nombre_delegacion, CA.id_distrito, CA.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(CA.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa, CA.bol_nulas, CA.num_mro, CA.tipo_mro, CA.modalidad, CA.anio
            FROM consulta_actas CA
            LEFT JOIN consulta_cat_delegacion D ON CA.id_delegacion = D.id_delegacion
            LEFT JOIN consulta_cat_colonia_cc1 C ON CA.clave_colonia = C.clave_colonia
            LEFT JOIN consulta_tipo_mesa_V TP ON CA.tipo_mro = TP.tipo_mro
            WHERE CA.anio = ${anio}${id_distrito != 0 ? ` AND CA.id_distrito = ${id_distrito}` : ''}
        ),
        ProyectosJSON AS (
            SELECT id_distrito, clave_Colonia, num_mro, tipo_mro, anio, (
                SELECT secuencial, nom_proyecto, descripcion, rubro_general votos, votos_sei, total_votos
                FROM consulta_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.num_mro = V1.num_mro AND V2.tipo_mro = V1.tipo_mro AND V2.anio = V1.anio
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS proyectos
            FROM consulta_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, num_mro, tipo_mro, anio
        )
        SELECT CA1.nombre_delegacion, CA1.id_distrito, CA1.clave_colonia, CA1.nombre_colonia, CA1.mesa, CA1.bol_nulas, COALESCE(CA2.bol_nulas, 0) AS bol_nulas_sei, P.proyectos
        FROM CA CA1
        LEFT JOIN CA CA2 ON CA1.id_distrito = CA2.id_distrito AND CA1.clave_colonia = CA2.clave_colonia AND CA1.num_mro = CA2.num_mro AND CA1.tipo_mro = CA2.tipo_mro AND CA2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON CA1.id_distrito = P.id_distrito AND CA1.clave_colonia = P.clave_colonia AND CA1.num_mro = P.num_mro AND CA1.tipo_mro = P.tipo_mro AND CA1.anio = P.anio
        WHERE CA1.modalidad = 1
        ORDER BY CA1.id_distrito, CA1.nombre_delegacion, CA1.nombre_colonia, CA1.num_mro, CA1.tipo_mro ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Base_Datos.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 14;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A6').value = subtitulo;
                worksheet.getCell('L9').value = `Fecha: ${fecha}`;
                worksheet.getCell('L10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                for (let acta of actas) {
                    let sum_total = 0;
                    const { nombre_delegacion, id_distrito: distrito, clave_colonia, nombre_colonia, mesa, bol_nulas, bol_nulas_sei, proyectos } = acta;
                    for (let proyecto of JSON.parse(proyectos)) {
                        const { secuencial, nom_proyecto, descripcion, rubro_general, votos, votos_sei, total_votos } = proyecto;
                        sum_total += total_votos;
                        const X = { nombre_delegacion, distrito, clave_colonia, nombre_colonia, mesa, secuencial, nom_proyecto, descripcion, rubro_general, votos, votos_sei, total_votos };
                        for (let i = 1; i <= 13; i++)
                            worksheet.getCell(fila, i).style = [2, 6, 10, 11, 12].includes(i) ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                        Object.keys(X).forEach((key, i) => {
                            worksheet.getCell(fila, i + 1).value = X[key];
                        });
                        fila++;
                    }
                    for (let i = 1; i <= 13; i++)
                        worksheet.getCell(fila, i).style = i <= 9 ? contenidoStyle : { ...contenidoStyle, numFmt: '#,##0' };
                    worksheet.getCell(fila, 9).value = 'Opiniones Nulas';
                    worksheet.getCell(fila, 10).value = bol_nulas;
                    worksheet.getCell(fila, 11).value = bol_nulas_sei;
                    worksheet.getCell(fila, 12).value = sum_total;
                    worksheet.getCell(fila, 13).value = bol_nulas + bol_nulas_sei + sum_total;
                    fila++;
                }
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 3 || (index >= 6 && index <= 8)) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 12)
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
                    msg: 'Reporte generado correctamente!!!',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_BaseDeDatos-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en BaseDatos: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en BaseDatos: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F2 - Concentrado de Proyectos participantes por Distrito y Unidad Territorial

export const ProyectosParticipantes = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`SELECT CPP.id_distrito, UPPER(CCD.nombre_delegacion) AS nombre_delegacion, CPP.clave_colonia, UPPER(CCC.nombre_colonia) AS nombre_colonia, UPPER(CPP.folio_proy_web) AS folio, CPP.num_proyecto, UPPER(STUFF((
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
        ) AS rubro_general, UPPER(CPP.nom_proyecto) AS nom_proyecto
        FROM consulta_prelacion_proyectos CPP
        LEFT JOIN consulta_cat_delegacion CCD ON CPP.id_delegacion = CCD.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 CCC ON CPP.clave_colonia = CCC.clave_colonia
        WHERE CPP.estatus = 1 AND CPP.anio = ${anio}${id_distrito != 0 ? ` AND CPP.id_distrito = ${id_distrito}` : ''}
        ORDER BY CPP.id_distrito, CCD.nombre_delegacion, CCC.nombre_colonia, CPP.num_proyecto ASC`))[0];
        if (!proyectos.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Proyectos_Participantes.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A4').value = subtitulo;
                worksheet.getCell('H9').value = `Fecha: ${fecha}`;
                worksheet.getCell('H10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                proyectos.forEach(proyecto => {
                    Object.keys(proyecto).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = proyecto[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 7).value = 'Total';
                worksheet.getCell(fila, 7).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 8).value = proyectos.length;
                worksheet.getCell(fila, 8).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3 || index == 6 || index == 7) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 11)
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
                    msg: 'Reporte generado correctamente!!!',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ProyectosParticipantes-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ProyectosParticipantes: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ProyectosParticipantes: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F4 - Validación de Resultados de la Consulta Ciudadana Detalle Mesa

export const ConsultaCiudadanaDetalle = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, A.id_delegacion, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(A.num_mro, NULLIF(CONCAT(' ', TM.mesa), '')) AS mesa, A.num_mro, A.tipo_mro, A.modalidad, A.levantada_distrito,
            A.total_ciudadanos, A.bol_nulas, A.bol_recibidas, A.bol_adicionales, A.bol_sobrantes, A.votacion_total_emitida, A.coordinador_sino, A.num_integrantes, A.observador_sino, A.anio
            FROM consulta_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            INNER JOIN consulta_tipo_mesa_V TM ON A.tipo_mro = TM.tipo_mro
            WHERE A.estatus = 1 AND A.anio = ${anio}${id_distrito != 0 ? ` AND A.id_distrito = ${id_distrito}` : ''}
        ),
        LD AS (
            SELECT id_distrito, clave_colonia, num_mro, tipo_mro,
            SUM(CASE WHEN levantada_distrito > 0 AND bol_recibidas = 0 THEN 0 ELSE total_ciudadanos END) AS ciudadania,
            SUM(CASE WHEN levantada_distrito > 0 AND bol_recibidas = 0 THEN total_ciudadanos ELSE 0 END) AS distrito
            FROM CA
            GROUP BY id_distrito, clave_colonia, num_mro, tipo_mro
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE ${campo} = 1
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
            SELECt id_distrito, clave_colonia, num_mro, tipo_mro, anio, (
                SELECT secuencial, votos, votos_sei, total_votos
                FROM consulta_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.num_mro = V1.num_mro AND V2.tipo_mro = V1.tipo_mro AND V2.anio = V1.anio
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS proyectos
            FROM consulta_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, num_mro, tipo_mro, anio
        )
        SELECT A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, A1.mesa, A1.bol_recibidas, A1.bol_adicionales, A1.bol_sobrantes, LD.ciudadania, LD.distrito, P.proyectos, A1.bol_nulas, COALESCE(A2.bol_nulas, 0) AS bol_nulas_sei, A1.bol_nulas + COALESCE(A2.bol_nulas, 0) AS total_nulas, A1.votacion_total_emitida , COALESCE(A2.votacion_total_emitida, 0) AS votacion_total_emitida_sei, A1.votacion_total_emitida + COALESCE(A2.votacion_total_emitida, 0) AS total_computada,
        CASE A1.coordinador_sino WHEN 1 THEN 'SI' ELSE 'NO' END AS coordinador_sino, COALESCE(A1.num_integrantes, 0) AS num_integrantes, CASE A1.observador_sino WHEN 1 THEN 'SI' ELSE 'NO' END AS observador_sino
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.num_mro = P.num_mro AND A1.tipo_mro = P.tipo_mro AND A1.anio = P.anio
        LEFT JOIN LD ON A1.id_distrito = LD.id_distrito AND A1.clave_colonia = LD.clave_colonia AND A1.num_mro = LD.num_mro AND A1.tipo_mro = LD.tipo_mro
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_Colonia = A1.clave_colonia)
        ORDER BY A1.id_distrito, A1.nombre_delegacion, A1.nombre_colonia, A1.num_mro, A1.tipo_mro ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        const max = Math.max(...actas.map(acta => JSON.parse(acta.proyectos).length));
        workbook.xlsx.readFile(path.join(plantillas[2], 'Validacion_Resultados_Detalle.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const celdasTotales = 19 + (max * 3);
                let fila = 10, celda = 11;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('A6').value = `VALIDACIÓN DE RESULTADOS DE LA CONSULTA CIUDADANA DETALLE MESA ${anio}`;
                worksheet.getCell('S4').value = 'FORMATO 4';
                worksheet.getCell('R7').value = `Fecha: ${fecha}`;
                worksheet.getCell('R8').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                for (let i = 1; i <= max; i++) {
                    for (let j = 1; j <= 3; j++)
                        worksheet.spliceColumns(celda, 0, [null]);
                    if (!worksheet.getCell(8, celda).isMerged)
                        worksheet.mergeCells(8, celda, 8, celda + 2);
                    for (let j = celda; j <= celda + 2; j++)
                        worksheet.getCell(8, j).style = contenidoStyle;
                    worksheet.getCell(8, celda).value = i;
                    worksheet.getCell(8, celda).style = fill;
                    worksheet.getCell(9, celda).value = 'Opiniones Mesa';
                    worksheet.getCell(9, celda).style = fill;
                    worksheet.getCell(9, celda + 1).value = 'Opiniones (SEI: vía remota)';
                    worksheet.getCell(9, celda + 1).style = fill;
                    worksheet.getCell(9, celda + 2).value = `Total de Opiniones Proyecto ${i}`;
                    worksheet.getCell(9, celda + 2).style = fill;
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
                const imprimir = (index, text) => {
                    worksheet.getCell(fila, index).value = text;
                    worksheet.getCell(fila, index).style = (index > 5 && index < celdasTotales - 2) || index == celdasTotales - 1 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                };
                const imprimirProyectos = (index, proyectos) => {
                    let i = index;
                    proyectos.forEach(proyecto => {
                        Object.entries(proyecto).forEach(([campo, valor]) => {
                            if (!campo.includes('secuencial')) {
                                imprimir(i, valor);
                                i++;
                            }
                        });
                    });
                    return i;
                };
                actas.forEach(acta => {
                    let i = 1;
                    Object.entries(acta).forEach(([campo, valor]) => {
                        if (!campo.match('proyectos')) {
                            imprimir(i, valor);
                            i++;
                            return;
                        }
                        i = imprimirProyectos(i, JSON.parse(valor));
                        const faltantes = max - JSON.parse(valor).length;
                        for (let x = 0; x < faltantes * 3; x++) {
                            imprimir(i, '');
                            i++
                        }
                    });
                    fila++;
                });
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 8)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 14;
                        if (index == 0 || index == 2) {
                            if (maxLength > 70)
                                column.width = 70;
                            else if (maxLength < (index == 0 ? 21 : 36))
                                column.width = index == 0 ? 21 : 36;
                            else
                                column.width = maxLength;
                        }
                    }
                    if (index >= 10 && index <= 10 + (max * 3))
                        column.width = 15;
                });
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ConsultaCiudadanaDetalle-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ConsultaCiudadanaDetalle: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ConsultaCiudadanaDetalle: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F5 - Resultado de Opiniones por Mesa

export const OpinionesMesa = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(A.num_mro, NULLIF(CONCAT(' ', TM.mesa), '')) AS mesa, A.num_mro, A.tipo_mro, A.modalidad, A.anio, A.bol_nulas
            FROM consulta_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            INNER JOIN consulta_tipo_mesa_V TM ON A.tipo_mro = TM.tipo_mro
            WHERE A.estatus = 1 AND A.anio = ${anio}${id_distrito != 0 ? ` AND A.id_distrito = ${id_distrito}` : ''}
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE ${campo} = 1
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
            SELECT	id_distrito, clave_colonia, num_mro, tipo_mro, anio, (
                SELECT secuencial, nom_proyecto, votos, votos_sei, total_votos
                FROM consulta_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.num_mro = V1.num_mro AND V2.tipo_mro = V1.tipo_mro AND V2.anio = V1.anio
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS proyectos
            FROM consulta_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, num_mro, tipo_mro, anio
        )
        SELECT A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, A1.mesa, P.proyectos, A1.bol_nulas, COALESCE(A2.bol_nulas, 0) AS bol_nulas_sei
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.num_mro = P.num_mro AND A1.tipo_mro = P.tipo_mro AND A1.anio = P.anio
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_colonia = A1.clave_colonia)
        ORDER BY A1.id_distrito, A1.nombre_delegacion, A1.nombre_colonia, A1.num_mro, A1.tipo_mro ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[0], 'Resultados_Opi_Mesa.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 10;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:J2');
                worksheet.getCell('A3').value = titulos[1];
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:J3');
                worksheet.getCell('A5').value = subtitulo;
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:J5');
                worksheet.getCell('A6').value = `RESULTADOS DE OPINIONES POR MESA ${anio}`;
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:J6');
                worksheet.getCell('J4').value = 'FORMATO 5';
                worksheet.getCell('J7').value = fecha;
                worksheet.getCell('J8').value = hora.substring(0, hora.length - 3);
                worksheet.getCell('E9').value = 'Clave del Proyecto';
                worksheet.getCell('G9').value = 'Nombre del Proyecto Especifico';
                for (let acta of actas) {
                    let sum_votos = 0, sum_votos_sei = 0;
                    const { id_distrito: distrito, nombre_delegacion, clave_colonia, nombre_colonia, mesa, proyectos, bol_nulas, bol_nulas_sei } = acta;
                    sum_votos += bol_nulas, sum_votos_sei += bol_nulas_sei;
                    for (let proyecto of JSON.parse(proyectos)) {
                        const { secuencial, nom_proyecto, votos, votos_sei, total_votos } = proyecto;
                        sum_votos += votos, sum_votos_sei += votos_sei;
                        const X = { distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, mesa, nom_proyecto, votos, votos_sei, total_votos };
                        Object.keys(X).forEach((key, i) => {
                            worksheet.getCell(fila, i + 1).value = X[key];
                            worksheet.getCell(fila, i + 1).style = [4, 7, 8, 9].includes(i) ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                        });
                        fila++
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
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3 || index == 6) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 8)
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
                    reporte: `Reporte_OpinionesMesa-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en OpinionesMesa: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en OpinionesMesa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F6 - Validación de Resultados de la Consulta por Unidad Territorial

export const ConsultaUnidadTerritorial = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, A.id_delegacion, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, A.num_mro, A.tipo_mro, A.modalidad, A.levantada_distrito,
            A.total_ciudadanos, A.bol_nulas, A.votacion_total_emitida, A.coordinador_sino, A.observador_sino, A.anio
            FROM consulta_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            WHERE A.estatus = 1 AND A.anio = ${anio}${id_distrito != 0 ? ` AND A.id_distrito = ${id_distrito}` : ''}
        ),
        LD AS (
            SELECT id_distrito, clave_colonia, SUM(CASE levantada_distrito WHEN 0 THEN total_ciudadanos ELSE 0 END) AS ciudadania, SUM(CASE levantada_distrito WHEN 1 THEN total_ciudadanos ELSE 0 END) AS distrito
            FROM CA
            GROUP BY id_distrito, clave_colonia
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE ${campo} = 1
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
            SELECt id_distrito, clave_colonia, anio, (
                SELECT secuencial, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
                FROM consulta_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.anio = V1.anio
                GROUP BY secuencial
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS proyectos
            FROM consulta_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, anio
        )
        SELECT  A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, LD.ciudadania, LD.distrito, P.proyectos, SUM(A1.bol_nulas) AS bol_nulas, SUM(A2.bol_nulas) AS bol_nulas_sei, SUM(A1.bol_nulas) + SUM(A2.bol_nulas) AS bol_nulas_totales,
        SUM(A1.votacion_total_emitida) AS votacion_total_emitida, SUM(A2.votacion_total_emitida) AS votacion_total_emitida_sei, SUM(A1.votacion_total_emitida) + SUM(A2.votacion_total_emitida) AS total_computada,
        CASE WHEN SUM(CAST(A1.coordinador_sino AS INT)) > 0 THEN 'SI' ELSE 'NO' END AS coordinador_sino, CASE WHEN SUM(CAST(A1.observador_sino AS INT)) > 0 THEN 'SI' ELSE 'NO' END AS observador_sino
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.anio = P.anio
        LEFT JOIN LD ON A1.id_distrito = LD.id_distrito AND A1.clave_colonia = LD.clave_colonia
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_colonia = A1.clave_colonia)
        GROUP BY A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, LD.ciudadania, LD.distrito,P.proyectos , A1.anio
        ORDER BY A1.id_distrito, A1.nombre_delegacion, A1.nombre_colonia ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        const max = Math.max(...actas.map(acta => JSON.parse(acta.proyectos).length));
        workbook.xlsx.readFile(path.join(plantillas[2], 'Validacion_Resultados.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const celdasTotales = 14 + (max * 3);
                let fila = 10, celda = 7;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('A6').value = `VALIDACIÓN DE RESULTADOS DE LA CONSULTA POR UNIDAD TERRITORIAL ${anio}`;
                worksheet.getCell('N4').value = 'FORMATO 6';
                worksheet.getCell('M7').value = `Fecha: ${fecha}`;
                worksheet.getCell('M8').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                for (let i = 1; i <= max; i++) {
                    for (let j = 1; j <= 3; j++)
                        worksheet.spliceColumns(celda, 0, [null]);
                    if (!worksheet.getCell(8, celda).isMerged)
                        worksheet.mergeCells(8, celda, 8, celda + 2);
                    for (let j = celda; j <= celda + 2; j++)
                        worksheet.getCell(8, j).style = contenidoStyle;
                    worksheet.getCell(8, celda).value = i;
                    worksheet.getCell(8, celda).style = fill;
                    worksheet.getCell(9, celda).value = 'Opiniones Mesa';
                    worksheet.getCell(9, celda).style = fill;
                    worksheet.getCell(9, celda + 1).value = 'Opiniones (SEI)';
                    worksheet.getCell(9, celda + 1).style = fill;
                    worksheet.getCell(9, celda + 2).value = `Total de Opiniones Proyecto ${i}`;
                    worksheet.getCell(9, celda + 2).style = fill;
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
                const imprimir = (index, text) => {
                    worksheet.getCell(fila, index).value = text;
                    worksheet.getCell(fila, index).style = index > 4 && index < celdasTotales - 1 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                };
                const imprimirProyectos = (index, proyectos) => {
                    let i = index;
                    proyectos.forEach(proyecto => {
                        Object.entries(proyecto).forEach(([campo, valor]) => {
                            if (!campo.includes('secuencial')) {
                                imprimir(i, valor);
                                i++;
                            }
                        });
                    });
                    return i;
                };
                actas.forEach(acta => {
                    let i = 1;
                    Object.entries(acta).forEach(([campo, valor]) => {
                        if (!campo.match('proyectos')) {
                            imprimir(i, valor);
                            i++;
                            return;
                        }
                        i = imprimirProyectos(i, JSON.parse(valor));
                        const faltantes = max - JSON.parse(valor).length;
                        for (let x = 0; x < faltantes * 3; x++) {
                            imprimir(i, '');
                            i++;
                        }
                    });
                    fila++;
                });
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 9)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 6;
                        if (index == 1 || index == 3) {
                            if (maxLength > 70)
                                column.width = 70;
                            else if (maxLength < (index == 1 ? 21 : 36))
                                column.width = index == 1 ? 21 : 36;
                            else
                                column.width = maxLength;
                        }
                    }
                    if (index > 5 && index <= 5 + (max * 3))
                        column.width = 15;
                });
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ConsultaUnidadTerritorial-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ConsultaUnidadTerritorial: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ConsultaUnidadTerritorial: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F7 - Concentrado de Opiniones por Unidad Territorial

export const OpinionesUT = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT DISTINCT A.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, A.modalidad, A.anio, A.bol_nulas
            FROM consulta_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            WHERE A.estatus = 1 AND A.anio = ${anio}${id_distrito != 0 ? ` AND A.id_distrito = ${id_distrito}` : ''}
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE ${campo} = 1
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
            SELECt id_distrito, clave_colonia, anio, (
                SELECT secuencial, nom_proyecto, rubro_general, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
                FROM consulta_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.anio = V1.anio
                GROUP BY secuencial, nom_proyecto, rubro_general
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS proyectos
            FROM consulta_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, anio
        )
        SELECT A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, P.proyectos, SUM(A1.bol_nulas) AS bol_nulas, SUM(COALESCE(A2.bol_nulas, 0)) AS bol_nulas_sei
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.anio = P.anio
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_colonia = A1.clave_colonia)
        GROUP BY A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, P.proyectos, A1.anio
        ORDER BY A1.id_distrito, A1.nombre_delegacion, A1.nombre_colonia ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Resultados_Opi_Mesa.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 10;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('A6').value = 'CONCENTRADO DE OPINIONES  POR UNIDAD TERRITORIAL';
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:J2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:J3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:J5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:J6');
                worksheet.getCell('J4').value = 'FORMATO 7';
                worksheet.getCell('J7').value = fecha;
                worksheet.getCell('J8').value = hora.substring(0, hora.length - 3);
                worksheet.getCell('F9').value = 'Rubro General o Destino';
                for (let acta of actas) {
                    const { id_distrito: distrito, nombre_delegacion, clave_colonia, nombre_colonia, proyectos, bol_nulas, bol_nulas_sei } = acta;
                    for (let proyecto of JSON.parse(proyectos)) {
                        const { secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos } = proyecto;
                        const X = { distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos };
                        Object.keys(X).forEach((key, i) => {
                            worksheet.getCell(fila, i + 1).value = X[key];
                            worksheet.getCell(fila, i + 1).style = [7, 8, 9].includes(i) ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                        });
                        fila++;
                    }
                    for (let i = 1; i <= 10; i++)
                        worksheet.getCell(fila, i).style = i <= 6 ? contenidoStyle : i == 7 ? { ...fill, font: { ...fill.font, bold: false } } : { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                    worksheet.getCell(fila, 7).value = 'Opiniones Nulas';
                    worksheet.getCell(fila, 8).value = bol_nulas;
                    worksheet.getCell(fila, 9).value = bol_nulas_sei;
                    worksheet.getCell(fila, 10).value = bol_nulas + bol_nulas_sei;
                    fila++;
                }
                worksheet.columns[5].width = 35;
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3 || index == 5 || index == 6) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 8)
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
                    reporte: `Reporte_OpinionesUnidadTerritorial-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en OpinionesUT: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en OpinionesUT: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F8 - Proyectos por Unidad Territorial que Obtuvieron el Primer Lugar en la Consulta de Presupuesto Participativo

export const ProyectosPrimerLugar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia
            FROM consulta_actas
            WHERE modalidad = 1 AND anio = ${anio}${id_distrito != 0 ? ` AND id_distrito = ${id_distrito}` : ''}
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE ${campo} = 1
            GROUP BY id_distrito, clave_colonia
        ),
        MesasCapturadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS capturadas
            FROM CA
            GROUP BY id_distrito, clave_colonia
        ),
        Mesas AS (
            SELECt C.id_distrito, C.clave_colonia
            FROM MesasCapturadas C
            INNER JOIN MesasEsperadas E ON C.id_distrito = E.id_distrito AND C.clave_colonia = E.clave_colonia
            WHERE C.capturadas = E.total
        ),
        ActasValidadas AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
            FROM consulta_actas_VVS V
            WHERE EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = V.id_distrito AND clave_colonia = V.clave_colonia) AND anio = ${anio} AND ${campo} = 1
        ),
        Votos AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
            FROM ActasValidadas
            GROUP BY id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto
        ),
        RANKING AS (
            SELECT *, DENSE_RANK() OVER (PARTITION BY clave_colonia ORDER BY total_votos DESC) AS DR, COUNT(*) OVER (PARTITION BY clave_colonia, total_votos) AS empate
            FROM Votos
        )
        SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM RANKING
        WHERE DR = 1 AND empate <= 1 AND total_votos > 0
        ORDER BY id_distrito, nombre_delegacion, nombre_colonia, secuencial ASC`))[0];
        if (!proyectos.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Proyectos-GE.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'PROYECTOS POR UNIDAD TERRITORIAL QUE OBTUVIERON EL PRIMER LUGAR EN LA CONSULTA DE PRESUPUESTO PARTICIPATIVO';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:I2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:I3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:I5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:I6');
                worksheet.getCell('J4').value = 'FORMATO 8';
                worksheet.getCell('J7').value = fecha;
                worksheet.getCell('J8').value = hora.substring(0, hora.length - 3);
                let fila = 11;
                let colonias = [];
                proyectos.forEach(res => {
                    if (!colonias.includes(res.clave_colonia))
                        colonias.push(res.clave_colonia);
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = index + 1 >= 8 && index + 1 <= 10 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                });
                if (!worksheet.getCell(fila, 2).isMerged)
                    worksheet.mergeCells(fila, 2, fila, 3);
                for (let i = 1; i <= 10; i++)
                    worksheet.getCell(fila, i).style = [4, 7].includes(i) ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                worksheet.getCell(fila, 2).value = 'Total de Unidades Territoriales';
                worksheet.getCell(fila, 2).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.getCell(fila, 4).value = colonias.length;
                worksheet.getCell(fila, 6).value = 'Total de Proyectos';
                worksheet.getCell(fila, 6).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.getCell(fila, 7).value = proyectos.length;
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3 || index == 5 || index == 6) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 9)
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
                    reporte: `Reporte_GanadoresPrimerLugar-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ProyectosPrimerLugar: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ProyectosPrimerLugar: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F9 - Casos de Empate de los Proyectos que Obtuvieron el Primer Lugar

export const ProyectosEmpatePrimerLugar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia
            FROM consulta_actas
            WHERE modalidad = 1 AND anio = ${anio}${id_distrito != 0 ? ` AND id_distrito = ${id_distrito}` : ''}
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE ${campo} = 1
            GROUP BY id_distrito, clave_colonia
        ),
        MesasCapturadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS capturadas
            FROM CA
            GROUP BY id_distrito, clave_colonia
        ),
        Mesas AS (
            SELECt C.id_distrito, C.clave_colonia
            FROM MesasCapturadas C
            INNER JOIN MesasEsperadas E ON C.id_distrito = E.id_distrito AND C.clave_colonia = E.clave_colonia
            WHERE C.capturadas = E.total
        ),
        ActasValidadas AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
            FROM consulta_actas_VVS V
            WHERE EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = V.id_distrito AND clave_colonia = V.clave_colonia) AND anio = ${anio} AND ${campo} = 1
        ),
        Votos AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
            FROM ActasValidadas
            GROUP BY id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto
        ),
        Ranking AS (
            SELECT *, DENSE_RANK() OVER (PARTITION BY clave_colonia ORDER BY total_votos DESC) AS DR, COUNT(*) OVER (PARTITION BY clave_colonia, total_votos) AS empate
            FROM Votos
        )
        SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM Ranking
        WHERE DR = 1 AND empate > 1 AND total_votos > 0
        ORDER BY id_distrito, nombre_delegacion, nombre_colonia, secuencial ASC`))[0];
        if (!proyectos.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Proyectos-GE.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'CASOS DE EMPATE DE LOS PROYECTOS QUE OBTUVIERON EL PRIMER LUGAR';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:I2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:I3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:I5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:I6');
                worksheet.getCell('J4').value = 'FORMATO 9';
                worksheet.getCell('J7').value = fecha;
                worksheet.getCell('J8').value = hora.substring(0, hora.length - 3);
                let fila = 11;
                let colonias = [];
                proyectos.forEach(res => {
                    if (!colonias.includes(res.clave_colonia))
                        colonias.push(res.clave_colonia);
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = index + 1 >= 8 && index + 1 <= 10 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                });
                if (!worksheet.getCell(fila, 2).isMerged)
                    worksheet.mergeCells(fila, 2, fila, 3);
                for (let i = 1; i <= 10; i++)
                    worksheet.getCell(fila, i).style = [4, 7].includes(i) ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                worksheet.getCell(fila, 2).value = 'Total de Unidades Territoriales';
                worksheet.getCell(fila, 2).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.getCell(fila, 4).value = colonias.length;
                worksheet.getCell(fila, 6).value = 'Total de Proyectos';
                worksheet.getCell(fila, 6).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.getCell(fila, 7).value = proyectos.length;
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3 || index == 5 || index == 6) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 9)
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
                    reporte: `Reporte_EmpatadosPrimerLugar-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ProyectosEmpatePrimerLugar: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ProyectosEmpatePrimerLugar: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F10 - Concentrado de Unidades Territoriales que NO Recibieron Opiniones

export const ProyectosSinOpiniones = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia, modalidad, anio
            FROM consulta_actas
            WHERE anio = ${anio}${id_distrito != 0 ? ` AND id_distrito = ${id_distrito}` : ''} AND votacion_total_emitida = 0
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE ${campo} = 1
            GROUP BY id_distrito, clave_colonia
        ),
        MesasCapturadas AS (
            SELECT id_distrito, clave_colonia, modalidad, anio, COUNT(*) AS capturadas
            FROM CA
            GROUP BY id_distrito, clave_colonia, modalidad, anio
        ),
        Mesas AS (
            SELECt C.id_distrito, C.clave_colonia, C.modalidad, C.anio
            FROM MesasCapturadas C
            INNER JOIN MesasEsperadas E ON C.id_distrito = E.id_distrito AND C.clave_colonia = E.clave_colonia
            WHERE C.capturadas = E.total
        ),
        V AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos, anio
            FROM consulta_actas_VVS
        )
        SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM V
        WHERE EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = V.id_distrito AND clave_colonia = V.clave_colonia AND anio = V.anio AND modalidad = 1)
        AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = V.id_distrito AND clave_colonia = V.clave_colonia AND anio = V.anio AND modalidad = 2)
        ORDER BY id_distrito, nombre_delegacion, nombre_colonia, secuencial ASC`))[0];
        if (!proyectos.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'UTNo_Recibieron_Opiniones.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('A6').value = 'CONCENTRADO DE UNIDADES TERRITORIALES QUE NO RECIBIERON OPINIONES';
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:J2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:J3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:J5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:J6');
                worksheet.getCell('J8').value = fecha;
                worksheet.getCell('J9').value = hora.substring(0, hora.length - 3);
                worksheet.getCell('J10').value = 'FORMATO 10';
                let fila = 12;
                proyectos.forEach(res => {
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 6).value = 'Total';
                worksheet.getCell(fila, 6).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 7).value = proyectos.length;
                worksheet.getCell(fila, 7).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3 || index == 5 || index == 6) {
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
                    reporte: `Reporte_UTNoRecibieronOpiniones-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ProyectosSinOpiniones: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ProyectosSinOpiniones: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F11 - Reporte Asistencia por Unidad Territorial

export const AsistenciaUT = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`SELECT CA.id_distrito, UPPER(CCD.nombre_delegacion) AS nombre_delegacion, CA.clave_colonia, UPPER(CCC.nombre_colonia) AS nombre_colonia, CONCAT(CA.num_mro, CASE WHEN TP.mesa IS NOT NULL THEN CONCAT(' ', TP.mesa) END) AS mesa, CONVERT(VARCHAR(10), CA.fecha_alta, 103) AS fecha_alta, CONVERT(VARCHAR(8), CA.fecha_alta, 114) AS hora_alta,
        CONVERT(VARCHAR(10), CA.fecha_modif, 103) AS fecha_modif, CONVERT(VARCHAR(8), CA.fecha_modif, 114) AS hora_modif, COALESCE(CA.num_integrantes, 0) AS num_integrantes, CASE WHEN CA.observador_sino = 1 THEN 'SI' ELSE 'NO' END AS observador_sino
        FROM consulta_actas CA
        LEFT JOIN consulta_tipo_mesa_V TP ON CA.tipo_mro = TP.tipo_mro
        LEFT JOIN consulta_cat_delegacion CCD ON CA.id_delegacion = CCD.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 CCC ON CA.clave_colonia = CCC.clave_colonia
        WHERE CA.modalidad = 1 AND CA.estatus = 1 AND CA.anio = ${anio} ${id_distrito != 0 ? `AND CA.id_distrito = ${id_distrito}` : ''}
        ORDER BY CA.id_distrito, CCD.nombre_delegacion, CCC.nombre_colonia, CA.num_mro, CA.tipo_mro ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Reporte_Asistencia.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 12;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A4').value = subtitulo;
                worksheet.getCell('J7').value = `Fecha: ${fecha}`;
                worksheet.getCell('J8').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                actas.forEach(acta => {
                    Object.keys(acta).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = acta[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 5).value = 'Total';
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 6).value = actas.length;
                worksheet.getCell(fila, 6).style = { ...contenidoStyle, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3) {
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
                    reporte: `Reporte_AsistenciaUT-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en AsistenciaUT: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en AsistenciaUT: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F12 - Mesas con Cómputo Capturado

export const MesasConComputo = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = (await SICOVACC.sequelize.query(`SELECT M.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, M.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(M.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa
        FROM consulta_mros M
        LEFT JOIN consulta_tipo_mesa_V TP ON M.tipo_mro = TP.tipo_mro
        LEFT JOIN consulta_cat_delegacion D ON M.id_delegacion = D.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 C ON M.clave_colonia = C.clave_colonia
        WHERE M.${campo} = 1 AND EXISTS (SELECT 1 FROM consulta_actas WHERE modalidad = 1 AND estatus = 1 AND anio = ${anio} AND clave_colonia = M.clave_colonia AND num_mro = M.num_mro AND tipo_mro = M.tipo_mro)${id_distrito != 0 ? ` AND M.id_distrito = ${id_distrito}` : ''}
        ORDER BY M.id_distrito, D.nombre_delegacion, C.nombre_colonia, M.num_mro, M.tipo_mro ASC`))[0];
        if (!resultado.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Mesas_CSC.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'MESAS RECEPTORAS DE OPINIÓN CON CÓMPUTO CAPTURADO';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:E2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:E3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:E5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:E6');
                worksheet.getCell('D7').value = 'FORMATO 12';
                worksheet.getCell('D9').value = `Fecha: ${fecha}`;
                worksheet.getCell('D10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                resultado.forEach(res => {
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 3).value = 'TOTAL';
                worksheet.getCell(fila, 3).style = fill;
                worksheet.getCell(fila, 4).value = resultado.length;
                worksheet.getCell(fila, 4).style = { ...fill, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 1) {
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
                    contentType: 'application/vnd-openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_MesasConComputo-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en MesasConComputo: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en MesasConComputo: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F13 - Mesas sin Cómputo Capturado

export const MesasSinComputo = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = (await SICOVACC.sequelize.query(`SELECT M.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, M.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(M.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa
        FROM consulta_mros M
        LEFT JOIN consulta_tipo_mesa_V TP ON M.tipo_mro = TP.tipo_mro
        LEFT JOIN consulta_cat_delegacion D ON M.id_delegacion = D.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 C ON M.clave_colonia = C.clave_colonia
        WHERE M.${campo} = 1 AND NOT EXISTS (SELECT 1 FROM consulta_actas WHERE modalidad = 1 AND estatus = 1 AND anio = ${anio} AND clave_colonia = M.clave_colonia AND num_mro = M.num_mro AND tipo_mro = M.tipo_mro)${id_distrito != 0 ? ` AND M.id_distrito = ${id_distrito}` : ''}
        ORDER BY M.id_distrito, D.nombre_delegacion, C.nombre_colonia, M.num_mro, M.tipo_mro ASC`))[0];
        if (!resultado.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Mesas_CSC.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'MESAS RECEPTORAS DE OPINIÓN SIN CÓMPUTO CAPTURADO';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:E2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:E3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:E5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:E6');
                worksheet.getCell('D7').value = 'FORMATO 12';
                worksheet.getCell('D9').value = `Fecha: ${fecha}`;
                worksheet.getCell('D10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                resultado.forEach(res => {
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 3).value = 'TOTAL';
                worksheet.getCell(fila, 3).style = fill;
                worksheet.getCell(fila, 4).value = resultado.length;
                worksheet.getCell(fila, 4).style = { ...fill, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 1) {
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
                    contentType: 'application/vnd-openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_MesasSinComputo-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en MesasSinComputo: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en MesasSinComputo: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F14 - Concentrado de Unidades Territoriales por Distrito Electoral con Cómputo Capturado (Grado de Avance)

export const UTConComputoGA = async (req = request, res = response) => {
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const avances = (await SICOVACC.sequelize.query(`SELECT D.id_distrito, M.UTTotal, COALESCE(UTV.UTValidadas, 0) AS UTValidadas, ROUND((CAST(COALESCE(UTV.UTValidadas, 0) AS FLOAT) * 100) / CAST(M.UTTotal AS FLOAT), 2) AS avance
        FROM consulta_cat_distrito D
        LEFT JOIN (SELECT id_distrito, COUNT(DISTINCT clave_colonia) AS UTTotal FROM consulta_mros WHERE ${campo} = 1 GROUP BY id_distrito) AS M ON D.id_distrito = M.id_distrito
        LEFT JOIN (
            SELECT id_distrito, COUNT(*) AS UTValidadas
            FROM consulta_cat_colonia_cc1 C
            WHERE ${campo} = 1 AND clave_colonia IN (
                SELECT A.clave_colonia
                FROM (SELECT clave_colonia, COUNT(*) AS total FROM consulta_mros WHERE ${campo} = 1 GROUP BY clave_colonia) AS A
                LEFT JOIN (SELECT clave_colonia, COUNT(*) AS cantidad FROM consulta_actas WHERE modalidad = 1 AND estatus = 1 AND anio = ${anio} GROUP BY clave_colonia) AS B ON A.clave_colonia = B.clave_colonia
                WHERE A.total = B.cantidad
            )
            GROUP BY id_distrito
        ) AS UTV ON D.id_distrito = UTV.id_distrito
        ORDER BY D.id_distrito ASC`))[0];
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[0], 'UT_Avance.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13, UTT = 0, UTCT = 0, total = 0;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A4').value = subtitulo;
                worksheet.getCell('H8').value = 'FORMATO 14';
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
                worksheet.getCell(fila, 6).style = { ...fill, numFmt: Number.isInteger(total) ? '#,##0' : '#,##0.##' }
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd-openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_UTConComputo(Avance)-${fecha}-${hora}.xlsx`,
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

//? F15 - Opiniones por Distrito

export const OpinionesDistrito = async (req = request, res = response) => {
    const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, SUM(votacion_total_emitida) AS total_votos, SUM(bol_nulas) AS total_nulas, modalidad
            FROM consulta_actas
            WHERE estatus = 1 AND anio = ${anio}
            GROUP BY id_distrito, modalidad
        )
        SELECT D.id_distrito, COALESCE(A1.total_votos - A1.total_nulas, 0) AS total_votos, COALESCE(A2.total_votos - A2.total_nulas, 0) AS total_votos_sei, COALESCE(A1.total_nulas, 0) AS total_nulas, COALESCE(A2.total_nulas, 0) AS total_nulas_sei
        FROM consulta_cat_distrito D
        LEFT JOIN CA A1 ON D.id_distrito = A1.id_distrito AND A1.modalidad = 1
        LEFT JOIN CA A2 ON D.id_distrito = A2.id_distrito AND A2.modalidad = 2
        ORDER BY D.id_distrito ASC`))[0];
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[0], 'Opiniones_Distrito.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13;
                let sum1 = 0, sum2 = 0, sum3 = 0, sum4 = 0, sum5 = 0;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 2 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A4').value = subtitulo;
                worksheet.getCell('A6').value = 'OPINIONES POD DISTRITO';
                worksheet.getCell('F5').value = 'FORMATO 15';
                worksheet.getCell('F9').value = `Fecha: ${fecha}`;
                worksheet.getCell('F10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
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
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd-openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_OpinionesDistrito-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en OpinionesDistrito: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en OpinionesDistrito: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F16 - Opiniones por Demarcación

export const OpinionesDemarcacion = async (req = request, res = response) => {
    const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_delegacion, SUM(votacion_total_emitida) AS total_votos, SUM(bol_nulas) AS total_nulas, modalidad
            FROM consulta_actas
            WHERE estatus = 1 AND anio = 2
            GROUP BY id_delegacion, modalidad
        )
        SELECT UPPER(D.nombre_delegacion) AS nombre_delegacion, COALESCE(A1.total_votos - A1.total_nulas, 0) AS total_votos, COALESCE(A2.total_votos - A2.total_nulas, 0) AS total_votos_sei, COALESCE(A1.total_nulas, 0) AS total_nulas, COALESCE(A2.total_nulas, 0) AS total_nulas_sei
        FROM consulta_cat_delegacion D
        LEFT JOIN CA A1 ON D.id_delegacion = A1.id_delegacion AND A1.modalidad = 1
        LEFT JOIN CA A2 ON D.id_delegacion = A2.id_delegacion AND A2.modalidad = 2
        ORDER BY D.nombre_delegacion ASC`))[0];
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[0], 'Opiniones_Demarcacion.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13;
                let sum1 = 0, sum2 = 0, sum3 = 0, sum4 = 0, sum5 = 0;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 2 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A4').value = subtitulo;
                worksheet.getCell('A6').value = 'OPINIONES POR DEMARCACIÓN';
                worksheet.getCell('F5').value = 'FORMATO 16';
                worksheet.getCell('F9').value = `Fecha: ${fecha}`;
                worksheet.getCell('F10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
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
                    contentType: 'application/vnd-openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_OpinionesDemarcacion-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en OpinionesDemarcacion: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en OpinionesDemarcacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Proyectos por Unidad Territorial que Obtuvieron el Segundo Lugar en la Consulta de Presupuesto Participativo

export const ProyectosSegundoLugar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia
            FROM consulta_actas
            WHERE modalidad = 1 AND anio = ${anio}${id_distrito != 0 ? ` AND id_distrito = ${id_distrito}` : ''}
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE ${campo} = 1
            GROUP BY id_distrito, clave_colonia
        ),
        MesasCapturadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS capturadas
            FROM CA
            GROUP BY id_distrito, clave_colonia
        ),
        Mesas AS (
            SELECt C.id_distrito, C.clave_colonia
            FROM MesasCapturadas C
            INNER JOIN MesasEsperadas E ON C.id_distrito = E.id_distrito AND C.clave_colonia = E.clave_colonia
            WHERE C.capturadas = E.total
        ),
        ActasValidadas AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
            FROM consulta_actas_VVS V
            WHERE EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = V.id_distrito AND clave_colonia = V.clave_colonia) AND anio = ${anio} AND ${campo} = 1
        ),
        Votos AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
            FROM ActasValidadas
            GROUP BY id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto
        ),
        RANKING AS (
            SELECT *, DENSE_RANK() OVER (PARTITION BY clave_colonia ORDER BY total_votos DESC) AS DR, COUNT(*) OVER (PARTITION BY clave_colonia, total_votos) AS empate
            FROM Votos
        )
        SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM RANKING
        WHERE DR = 2 AND empate <= 1 AND total_votos > 0
        ORDER BY id_distrito, nombre_delegacion, nombre_colonia, secuencial ASC`))[0];
        if (!proyectos.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Proyectos-GE.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'PROYECTOS POR UNIDAD TERRITORIAL QUE OBTUVIERON EL SEGUNDO LUGAR EN CONSULTA DE PRESUPUESTO PARTICIPATIVO';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:I2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:I3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:I5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:I6');
                worksheet.getCell('J4').value = 'FORMATO 9';
                worksheet.getCell('J7').value = fecha;
                worksheet.getCell('J8').value = hora.substring(0, hora.length - 3);
                let fila = 11;
                let colonias = [];
                proyectos.forEach(res => {
                    if (!colonias.includes(res.clave_colonia))
                        colonias.push(res.clave_colonia);
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = index + 1 >= 8 && index + 1 <= 10 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                });
                if (!worksheet.getCell(fila, 2).isMerged)
                    worksheet.mergeCells(fila, 2, fila, 3);
                for (let i = 1; i <= 10; i++)
                    worksheet.getCell(fila, i).style = [4, 7].includes(i) ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                worksheet.getCell(fila, 2).value = 'Total de Unidades Territoriales';
                worksheet.getCell(fila, 2).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 4).value = colonias.length;
                worksheet.getCell(fila, 6).value = 'Total de Proyectos';
                worksheet.getCell(fila, 6).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 7).value = proyectos.length;
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3 || index == 5 || index == 6) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 9)
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
                    reporte: `Reporte_GanadoresSegundoaLugar-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ProyectosSegundoLugar: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ProyectosSegundoLugar: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Casos de Empates de los Proyectos que Obtuvieron el Segundo Lugar

export const ProyectosEmpateSegundoLugar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia
            FROM consulta_actas
            WHERE modalidad = 1 AND anio = ${anio}${id_distrito != 0 ? ` AND id_distrito = ${id_distrito}` : ''}
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE ${campo} = 1
            GROUP BY id_distrito, clave_colonia
        ),
        MesasCapturadas AS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS capturadas
            FROM CA
            GROUP BY id_distrito, clave_colonia
        ),
        Mesas AS (
            SELECt C.id_distrito, C.clave_colonia
            FROM MesasCapturadas C
            INNER JOIN MesasEsperadas E ON C.id_distrito = E.id_distrito AND C.clave_colonia = E.clave_colonia
            WHERE C.capturadas = E.total
        ),
        ActasValidadas AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
            FROM consulta_actas_VVS V
            WHERE EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = V.id_distrito AND clave_colonia = V.clave_colonia) AND anio = ${anio} AND ${campo} = 1
        ),
        Votos AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
            FROM ActasValidadas
            GROUP BY id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto
        ),
        Ranking AS (
            SELECT *, DENSE_RANK() OVER (PARTITION BY clave_colonia ORDER BY total_votos DESC) AS DR, COUNT(*) OVER (PARTITION BY clave_colonia, total_votos) AS empate
            FROM Votos
        )
        SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM Ranking
        WHERE DR = 2 AND empate > 1 AND total_votos > 0
        ORDER BY id_distrito, nombre_delegacion, nombre_colonia, secuencial ASC`);
        if (proyectos[1] == 0)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Proyectos-GE.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'CASOS DE EMPATE DE LOS PROYECTOS QUE OBTUVIERON EL SEGUNDO LUGAR';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:I2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:I3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:I5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:I6');
                worksheet.getCell('J4').value = 'FORMATO 11';
                worksheet.getCell('J7').value = fecha;
                worksheet.getCell('J8').value = hora.substring(0, hora.length - 3);
                let fila = 11;
                let colonias = [];
                proyectos[0].forEach(res => {
                    if (!colonias.includes(res.clave_colonia))
                        colonias.push(res.clave_colonia);
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = index + 1 >= 8 && index + 1 <= 10 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                })
                if (!worksheet.getCell(fila, 2).isMerged)
                    worksheet.mergeCells(fila, 2, fila, 3);
                for (let i = 1; i <= 10; i++)
                    worksheet.getCell(fila, i).style = [4, 7].includes(i) ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                worksheet.getCell(fila, 2).value = 'Total de Unidades Territoriales';
                worksheet.getCell(fila, 2).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 4).value = colonias.length;
                worksheet.getCell(fila, 6).value = 'Total de Proyectos';
                worksheet.getCell(fila, 6).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 7).value = proyectos.length;
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 3 || index == 5 || index == 6) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 9)
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
                    reporte: `Reporte_EmpatadosSegundoLugar-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ProyectosEmpateSegundoLugar: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ProyectosEmpateSegundoLugar: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Actas Levantadas en Dirección Distrital

export const LevantadaDistrito = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = (await SICOVACC.sequelize.query(`SELECT CA.id_distrito, ROW_NUMBER() OVER (ORDER BY CA.id_distrito, CCC.nombre_colonia, CA.num_mro) AS NC, UPPER(CCC.nombre_colonia) AS nombre_colonia, CA.clave_colonia, CA.num_mro,
        UPPER(dbo.RazonDistrital(CA.razon_distrital)) AS razon_distrital
        FROM consulta_actas CA
        LEFT JOIN consulta_cat_colonia_cc1 CCC ON CA.clave_colonia = CCC.clave_colonia
        WHERE CA.modalidad = 1 AND CA.estatus = 1 AND CA.anio = ${anio} AND CA.levantada_distrito = 1 ${id_distrito != 0 ? `AND CA.id_distrito = ${id_distrito}` : ''}
        ORDER BY CA.id_distrito, nombre_colonia ASC`))[0];
        if (!resultado.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Reporte_Levantada_Distrito.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 11;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:F2');
                worksheet.getCell('A3').value = titulos[1];
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:F3');
                worksheet.getCell('A5').value = subtitulo;
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:F5');
                worksheet.getCell('A6').value = 'ACTAS LEVANTADAS EN DIRECCIÓN DISTRITAL (CAUSALES DE RECUENTO)';
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:F6');
                worksheet.getCell('F8').value = `Fecha: ${fecha}`;
                worksheet.getCell('F9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                resultado.forEach(res => {
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = index == 1 || index == 4 ? +res[key] : res[key];
                        worksheet.getCell(fila, index + 1).style = index == 1 || index == 4 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 5).value = 'Total';
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 6).value = resultado.length;
                worksheet.getCell(fila, 6).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 2 || index == 5) {
                        let maxLength = 0;
                        column.eachCell({ includeEmpty: false }, (cell, index) => {
                            if (index >= 9)
                                if (cell.value) {
                                    const length = cell.value.toString().length;
                                    if (length > maxLength)
                                        maxLength = length;
                                }
                        });
                        maxLength += 14;
                        if (maxLength > 75)
                            column.width = 75;
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

//? Porcentaje de Participación

export const Participacion = async (req = request, res = response) => {
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const porcentajes = (await SICOVACC.sequelize.query(`SELECT id_distrito, ROUND((COALESCE(total_opiniones, 0) * 100) / lista_nominal, 2) AS porcentaje
        FROM (
            SELECT D.id_distrito, CAST(SUM(votacion_total_emitida) AS FLOAT) AS total_opiniones, LN.lista_nominal
            FROM consulta_cat_distrito D
            LEFT JOIN (SELECT id_distrito, votacion_total_emitida FROM consulta_actas WHERE estatus = 1 AND anio = ${anio}) AS A ON D.id_distrito = A.id_distrito
            LEFT JOIN (SELECT id_distrito, SUM(lista_nominal) AS lista_nominal FROM consulta_mros WHERE ${campo} = 1 GROUP BY id_distrito) AS LN ON D.id_distrito = LN.id_distrito
            GROUP BY D.id_distrito, LN.lista_nominal
        ) AS A
        ORDER BY id_distrito ASC`))[0];
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[0], 'Reporte_Participacion.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(2);
                let fila = 11;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B5').value = subtitulo;
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