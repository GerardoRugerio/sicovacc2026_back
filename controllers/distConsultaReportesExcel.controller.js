import ExcelJs from 'exceljs';
import { request, response } from 'express';
import path from 'path';
import { aniosCAT, autor, contenidoStyle, fill, iecmLogo, plantillas, titulos, tituloStyle } from '../helpers/Constantes.js';
import { ConsultaTipoEleccion, FechaServer } from '../helpers/Consultas.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Consulta de Resultados Por Unidad Territorial

export const UTValidadas = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const colonias = (await SICOVACC.sequelize.query(`SELECT UPPER(D.nombre_delegacion) AS nombre_delegacion, C.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia
        FROM consulta_cat_colonia_cc1 C
        LEFT JOIN consulta_cat_delegacion D ON C.id_delegacion = D.id_delegacion
        WHERE C.id_distrito = ${id_distrito} AND EXISTS (
            SELECT 1
            FROM (
                SELECT clave_colonia
                FROM consulta_actas A
                WHERE modalidad = 1 AND estatus = 1 AND anio = ${anio}
                GROUP BY clave_colonia
                HAVING COUNT(*) = (SELECT COUNT(*) FROM consulta_mros WHERE ${campo} = 1 AND clave_colonia = A.clave_colonia)
            ) X
            WHERE X.clave_colonia = C.clave_colonia
        ) AND EXISTS (SELECT 1 FROM consulta_mros WHERE id_distrito = C.id_distrito AND clave_colonia = C.clave_colonia AND ${campo} = 1)
        ORDER BY nombre_delegacion, nombre_colonia`))[0]
        if (!colonias.length)
            return res.status(404).json({
                success: false,
                msg: "¡No existe información!"
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Reporte_UT-VPV.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'UNIDADES TERRITORIALES VALIDADAS';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:C2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:C3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:C5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:C6');
                worksheet.getCell('A10').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('C7').value = 'FORMATO 8';
                worksheet.getCell('C9').value = `Fecha: ${fecha}`;
                worksheet.getCell('C10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                colonias.forEach(colonia => {
                    Object.keys(colonia).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = colonia[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 2).value = 'Total:';
                worksheet.getCell(fila, 2).style = fill;
                worksheet.getCell(fila, 3).value = colonias.length;
                worksheet.getCell(fila, 3).style = { ...contenidoStyle, numFmt: '#,##0' };
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado corectamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_UTValidadas-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en UTValidadas: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en UTValidadas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

export const UTPorValidar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const colonias = (await SICOVACC.sequelize.query(`SELECT UPPER(D.nombre_delegacion) AS nombre_delegacion, C.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia
        FROM consulta_cat_colonia_cc1 C
        LEFT JOIN consulta_cat_delegacion D ON C.id_delegacion = D.id_delegacion
        WHERE C.id_distrito = ${id_distrito} AND NOT EXISTS (
            SELECT 1
            FROM (
                SELECT clave_colonia
                FROM consulta_actas A
                WHERE modalidad = 1 AND estatus = 1 AND anio = ${anio}
                GROUP BY clave_colonia
                HAVING COUNT(*) = (SELECT COUNT(*) FROM consulta_mros WHERE ${campo} = 1 AND clave_colonia = A.clave_colonia)
            ) X
            WHERE X.clave_colonia = C.clave_colonia
        ) AND EXISTS (SELECT 1 FROM consulta_mros WHERE id_distrito = C.id_distrito AND clave_colonia = C.clave_colonia AND ${campo} = 1)
        ORDER BY nombre_delegacion, nombre_colonia`))[0];
        if (!colonias.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Reporte_UT-VPV.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 13;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'UNIDADES TERRITORIALES POR VALIDAR';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:C2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:C3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:C5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:C6');
                worksheet.getCell('A10').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('C7').value = 'FORMATO 9';
                worksheet.getCell('C9').value = `Fecha: ${fecha}`;
                worksheet.getCell('C10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                colonias.forEach(colonia => {
                    Object.keys(colonia).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = colonia[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 2).value = 'Total:';
                worksheet.getCell(fila, 2).style = fill;
                worksheet.getCell(fila, 3).value = colonias.length;
                worksheet.getCell(fila, 3).style = { ...contenidoStyle, numFmt: '#,##0' };
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_UTPorValidar-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en UTPorValidar: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en UTPorValidar: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F2 - Concentrado de Proyectos Participantes por Unidad Territorial

export const ListadoProyectos = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`SELECT UPPER(Del.nombre_delegacion) AS nombre_delegacion, Proys.clave_colonia, UPPER(Cols.nombre_colonia) AS nombre_colonia, Proys.num_proyecto, UPPER(Proys.folio_proy_web) AS folio_proy_web,
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
        ) AS rubro_general, UPPER(Proys.nom_proyecto) AS nom_proyecto
        FROM consulta_prelacion_proyectos Proys
        LEFT JOIN consulta_cat_colonia_cc1 Cols ON Proys.clave_colonia = Cols.clave_colonia
        LEFT JOIN consulta_cat_delegacion Del ON Cols.id_delegacion = Del.id_delegacion
        WHERE Proys.estatus = 1 AND Proys.anio = ${anio} AND Cols.id_distrito = ${id_distrito}
        ORDER BY Del.nombre_delegacion, Cols.nombre_colonia, Proys.num_proyecto`))[0];
        if (!proyectos.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Listado_Proyectos.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 11;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('A6').value = 'CONCENTRADO DE PROYECTOS PARTICIPANTES POR UNIDAD TERRITORIAL';
                worksheet.getCell('B8').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('G8').value = `Fecha: ${fecha}`;
                worksheet.getCell('G9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                proyectos.forEach(proyecto => {
                    Object.keys(proyecto).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = proyecto[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 2).value = 'Total:';
                worksheet.getCell(fila, 2).style = fill;
                worksheet.getCell(fila, 3).value = proyectos.length;
                worksheet.getCell(fila, 3).style = { ...contenidoStyle, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 2 || index == 5 || index == 6) {
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
                    reporte: `Reporte_ListadoProyectos-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ListadoProyectos: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ListadoProyectos: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F4 - Validación de Resultados de la Consulta por Unidad Territorial

export const ValidacionResultados = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia, anio } = req.body;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, A.id_delegacion, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, A.num_mro, A.tipo_mro, A.modalidad, A.levantada_distrito,
            A.total_ciudadanos, A.bol_nulas, A.votacion_total_emitida, A.coordinador_sino, A.observador_sino, A.anio
            FROM consulta_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            WHERE A.id_distrito = ${id_distrito}${clave_colonia ? ` AND A.clave_colonia = '${clave_colonia}'` : ''} AND A.estatus = 1 AND A.anio = ${anio}
        ),
        LD AS (
            SELECT id_distrito, clave_colonia, SUM(CASE levantada_distrito WHEN 0 THEN total_ciudadanos ELSE 0 END) AS ciudadania, SUM(CASE levantada_distrito WHEN 1 THEN total_ciudadanos ELSE 0 END) AS distrito
            FROM CA
            GROUP BY id_distrito, clave_colonia
        ),
        MesasEsperadas aS (
            SELECT id_distrito, clave_colonia, COUNT(*) AS total
            FROM consulta_mros
            WHERE estatus_cc1 = 1
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
        SELECT A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, LD.ciudadania, LD.distrito, P.proyectos, SUM(A1.bol_nulas) AS bol_nulas, SUM(A2.bol_nulas) AS bol_nulas_sei, SUM(A1.bol_nulas) + SUM(A2.bol_nulas) AS bol_nulas_totales,
        SUM(A1.votacion_total_emitida) AS votacion_total_emitida, SUM(A2.votacion_total_emitida) AS votacion_total_emitida_sei, SUM(A1.votacion_total_emitida) + SUM(A2.votacion_total_emitida) AS total_computada,
        CASE WHEN SUM(CAST(A1.coordinador_sino AS INT)) > 0 THEN 'SI' ELSE 'NO' END AS coordinador_sino, CASE WHEN SUM(CAST(A1.observador_sino AS INT)) > 0 THEN 'SI' ELSE 'NO' END AS observador_sino
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.anio = P.anio
        LEFT JOIN LD ON A1.id_distrito = LD.id_distrito AND A1.clave_colonia = LD.clave_colonia
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_colonia = A1.clave_colonia)
        GROUP BY A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, LD.ciudadania, LD.distrito, P.proyectos, A1.anio
        ORDER BY A1.nombre_delegacion, A1.nombre_colonia ASC`))[0];
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
                const celdasTotales = 13 + (max * 3);
                let fila = 10, celda = 6;
                worksheet.spliceColumns(1, 1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('A6').value = 'VALIDACIÓN DE RESULTADOS DE LA CONSULTA POR UNIDAD TERRITORIAL';
                worksheet.getCell('A7').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('A7').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('L7').value = `Fecha: ${fecha}`;
                worksheet.getCell('L8').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
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
                    worksheet.getCell(fila, index).style = index > 3 && index < celdasTotales - 1 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
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
                    if (index == 0 || index == 2) {
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
                    if (index > 4 && index <= 4 + max * 3)
                        column.width = 15;
                });
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ValidacionResultado${clave_colonia ? `_${clave_colonia}_` : '-'}${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ValidacionResultados: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ValidacionResultados: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F5 - Validación de Resultados de la Consulta Detalle Mesa

export const ValidacionResultadosDetalle = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia, anio } = req.body;
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
            WHERE A.id_distrito = ${id_distrito}${clave_colonia ? ` AND A.clave_colonia = '${clave_colonia}'` : ''} AND A.estatus = 1 AND A.anio = ${anio}
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
        SELECT A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, A1.mesa, A1.bol_recibidas, A1.bol_adicionales, A1.bol_sobrantes, LD.ciudadania, LD.distrito, P.proyectos, A1.bol_nulas, COALESCE(A2.bol_nulas, 0) AS bol_nulas_sei,
        A1.bol_nulas + COALESCE(A2.bol_nulas, 0) AS total_nulas, A1.votacion_total_emitida , COALESCE(A2.votacion_total_emitida, 0) AS votacion_total_emitida_sei, A1.votacion_total_emitida + COALESCE(A2.votacion_total_emitida, 0) AS total_computada,
        CASE A1.coordinador_sino WHEN 1 THEN 'SI' ELSE 'NO' END AS coordinador_sino, COALESCE(A1.num_integrantes, 0) AS num_integrantes, CASE A1.observador_sino WHEN 1 THEN 'SI' ELSE 'NO' END AS observador_sino
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.num_mro = P.num_mro AND A1.tipo_mro = P.tipo_mro AND A1.anio = P.anio
        LEFT JOIN LD ON A1.id_distrito = LD.id_distrito AND A1.clave_colonia = LD.clave_colonia AND A1.num_mro = LD.num_mro AND A1.tipo_mro = LD.tipo_mro
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_Colonia = A1.clave_colonia)
        ORDER BY A1.nombre_delegacion, A1.nombre_colonia, A1.num_mro, A1.tipo_mro ASC`))[0];
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
                const celdasTotales = 18 + (max * 3);
                let fila = 10, celda = 10;
                worksheet.spliceColumns(1, 1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('A6').value = 'VALIDACIÓN DE RESULTADOS DE LA CONSULTA DETALLE MESA';
                worksheet.getCell('A7').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('A7').style = fill;
                worksheet.getCell('Q7').value = `Fecha: ${fecha}`;
                worksheet.getCell('Q8').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
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
                    worksheet.getCell(fila, index).style = (index > 4 && index < celdasTotales - 2) || index == celdasTotales - 1 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
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
                    if (index == 0 || index == 2) {
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
                    if (index > 8 && index <= 8 + (max * 3))
                        column.width = 15;
                });
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ValidacionResultadoDetalle${clave_colonia ? `_${clave_colonia}_` : '-'}${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ValidacionResultadosDetalle: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ValidacionResultadosDetalle: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F6 - Validación de Resultados de la Consulta por Nombre del Proyecto

export const ValidacionResultadosNombre = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia, anio } = req.body;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, A.id_delegacion, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, A.num_mro, A.tipo_mro, A.modalidad, A.levantada_distrito,
            A.total_ciudadanos, A.bol_nulas, A.votacion_total_emitida, A.coordinador_sino, A.observador_sino, A.anio
            FROM consulta_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            WHERE A.id_distrito = ${id_distrito} AND A.clave_colonia = '${clave_colonia}' AND A.estatus = 1 AND A.anio = ${anio}
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
                SELECT nom_proyecto AS proyecto, SUM(votos) AS votos, SUM(votos_sei) AS votos_sei, SUM(total_votos) AS total_votos
                FROM consulta_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.anio = V1.anio
                GROUP BY secuencial, nom_proyecto
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS proyectos
            FROM consulta_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, anio
        )
        SELECT A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, LD.ciudadania, LD.distrito, P.proyectos, SUM(A1.bol_nulas) AS bol_nulas, SUM(A2.bol_nulas) AS bol_nulas_sei, SUM(A1.bol_nulas) + SUM(A2.bol_nulas) AS bol_nulas_totales,
        SUM(A1.votacion_total_emitida) AS votacion_total_emitida, SUM(A2.votacion_total_emitida) AS votacion_total_emitida_sei, SUM(A1.votacion_total_emitida) + SUM(A2.votacion_total_emitida) AS total_computada,
        CASE WHEN SUM(CAST(A1.coordinador_sino AS INT)) > 0 THEN 'SI' ELSE 'NO' END AS coordinador_sino, CASE WHEN SUM(CAST(A1.observador_sino AS INT)) > 0 THEN 'SI' ELSE 'NO' END AS observador_sino
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.anio = P.anio
        LEFT JOIN LD ON A1.id_distrito = LD.id_distrito AND A1.clave_colonia = LD.clave_colonia
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_colonia = A1.clave_colonia)
        GROUP BY A1.id_distrito, A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, LD.ciudadania, LD.distrito, P.proyectos, A1.anio
        ORDER BY A1.nombre_delegacion, A1.nombre_colonia ASC`))[0];
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Validacion_Resultados_Nombre.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const celdasTotales = 13 + (JSON.parse(actas[0].proyectos).length * 3);
                let fila = 10, celda = 6;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('B7').value = id_distrito;
                worksheet.getCell('L7').value = `Fecha: ${fecha}`;
                worksheet.getCell('L8').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                for (const [i, proyecto] of JSON.parse(actas[0].proyectos).entries()) {
                    const { proyecto: nom } = proyecto;
                    for (let j = 1; j <= 3; j++)
                        worksheet.spliceColumns(celda, 0, [null]);
                    if (!worksheet.getCell(8, celda).isMerged)
                        worksheet.mergeCells(8, celda, 8, celda + 2);
                    for (let j = celda; j <= celda + 2; j++)
                        worksheet.getCell(8, j).style = contenidoStyle;
                    worksheet.getCell(8, celda).value = nom.trim();
                    worksheet.getCell(8, celda).style = fill
                    worksheet.getCell(9, celda).value = 'Opiniones Mesa';
                    worksheet.getCell(9, celda).style = fill;
                    worksheet.getCell(9, celda + 1).value = 'Opiniones (SEI: vía remota)';
                    worksheet.getCell(9, celda + 1).style = fill;
                    worksheet.getCell(9, celda + 2).value = `Total de Opiniones Proyecto ${i + 1}`;
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
                    worksheet.getCell(fila, index).style = index > 3 && index < celdasTotales - 3 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                };
                const imprimirProyectos = (index, proyectos) => {
                    let i = index;
                    proyectos.forEach(proyecto => {
                        Object.entries(proyecto).forEach(([campo, valor]) => {
                            if (!campo.includes('proyecto')) {
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
                    });
                    fila++;
                });
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 2) {
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
                    if (index > 4 && index <= 4 + (actas[0].proyectos.length * 3))
                        column.width = 15;
                });
                return workbook.xlsx.writeBuffer();
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ValidacionResultadoNombre_${clave_colonia}_${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ValidacionResultadosNombre: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ValidacionResultadosNombre: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F7 - Validación de Resultados de la Consulta por Nombre del Proyecto (Detalle por Mesa)

export const ValidacionResultadosNombreDetalle = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { clave_colonia, anio } = req.body;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, A.id_delegacion, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(A.num_mro, NULLIF(CONCAT(' ', TM.mesa), '')) AS mesa, A.num_mro, A.tipo_mro, A.modalidad,
            A.bol_recibidas, A.bol_adicionales, A.bol_sobrantes, A.levantada_distrito, A.total_ciudadanos, A.bol_nulas, A.votacion_total_emitida, A.coordinador_sino, A.num_integrantes, A.observador_sino, A.anio
            FROM consulta_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            INNEr JOIN consulta_tipo_mesa_V TM ON A.tipo_mro = TM.tipo_mro
            WHERE A.id_distrito = ${id_distrito} AND A.clave_colonia = '${clave_colonia}' AND A.estatus = 1 AND A.anio = ${anio}
        ),
        LD AS (
            SELECT id_distrito, clave_colonia, num_mro, tipo_mro, CASE WHEN levantada_distrito > 0 AND bol_recibidas = 0 THEN 0 ELSE total_ciudadanos END AS ciudadania, CASE WHEN levantada_distrito > 0 AND bol_recibidas = 0 THEN total_ciudadanos ELSE 0 END AS distrito
            FROM CA
            WHERE modalidad = 1
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
                SELECT nom_proyecto AS proyecto, votos, votos_sei, total_votos
                FROM consulta_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.num_mro = V1.num_mro AND V2.tipo_mro = V1.tipo_mro AND V2.anio = V1.anio
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS proyectos
            FROM consulta_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, num_mro, tipo_mro, anio
        )
        SELECT A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, A1.mesa, A1.bol_recibidas, A1.bol_adicionales, A1.bol_sobrantes, LD.ciudadania, LD.distrito, P.proyectos, A1.bol_nulas, COALESCE(A2.bol_nulas, 0) AS bol_nulas_sei, A1.bol_nulas + COALESCE(A2.bol_nulas, 0) AS bol_nulas_totales, A1.votacion_total_emitida, COALESCE(A2.votacion_total_emitida, 0) AS votacion_total_emitida_sei, A1.votacion_total_emitida + COALESCE(A2.votacion_total_emitida, 0) AS total_computada,
        CASE A1.coordinador_sino WHEN 1 THEN 'SI' ELSE 'NO' END AS coordinador_sino, CASE WHEN A1.num_integrantes IS NULL THEN 0 ELSE A1.num_integrantes END AS num_integrantes, CASE A1.observador_sino WHEN 1 THEN 'SI' ELSE 'NO' END AS observador_sino
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.num_mro = P.num_mro AND A1.tipo_mro = P.tipo_mro AND A1.anio = P.anio
        LEFT JOIN LD ON A1.id_distrito = LD.id_distrito AND A1.clave_colonia = LD.clave_colonia AND A1.num_mro = LD.num_mro AND A1.tipo_mro = LD.tipo_mro
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_Colonia = A1.clave_colonia)
        ORDER BY A1.nombre_delegacion, A1.nombre_colonia, A1.num_mro, A1.tipo_mro ASC`))[0]
        if (!actas.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const subtitulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Validacion_Resultados_Nombre_Detalle.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                const celdasTotales = 18 + (JSON.parse(actas[0].proyectos).length * 3);
                let fila = 10, celda = 10;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('B7').value = id_distrito;
                worksheet.getCell('Q7').value = `Fecha: ${fecha}`;
                worksheet.getCell('Q8').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                for (const [i, proyecto] of JSON.parse(actas[0].proyectos).entries()) {
                    const { proyecto: nom } = proyecto;
                    for (let j = 1; j <= 3; j++)
                        worksheet.spliceColumns(celda, 0, [null]);
                    if (!worksheet.getCell(8, celda).isMerged)
                        worksheet.mergeCells(8, celda, 8, celda + 2);
                    for (let j = celda; j <= celda + 2; j++)
                        worksheet.getCell(8, j).style = contenidoStyle;
                    worksheet.getCell(8, celda).value = nom.trim();
                    worksheet.getCell(8, celda).style = fill;
                    worksheet.getCell(9, celda).value = 'Opiniones Mesa';
                    worksheet.getCell(9, celda).style = fill;
                    worksheet.getCell(9, celda + 1).value = 'Opiniones (SEI: vía remota)';
                    worksheet.getCell(9, celda + 1).style = fill;
                    worksheet.getCell(9, celda + 2).value = `total de Opiniones Proyecto ${i}`;
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
                    worksheet.getCell(fila, index).style = index > 3 && index < celdasTotales - 3 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                };
                const imprimirProyectos = (index, proyectos) => {
                    let i = index;
                    proyectos.forEach(proyecto => {
                        Object.entries(proyecto).forEach(([campo, valor]) => {
                            if (!campo.includes('proyecto')) {
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
                    });
                    fila++;
                });
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 2) {
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
                            column.width = 70
                        else if (maxLength < 16)
                            column.width = 16
                        else
                            column.width = maxLength;
                    }
                    if (index > 8 && index <= 8 + (actas[0].proyectos.length * 3))
                        column.width = 15;
                });
                return workbook.xlsx.writeBuffer()
            })
            .then(buffer => {
                res.json({
                    success: true,
                    msg: 'Reporte generado correctamente',
                    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    reporte: `Reporte_ValidacionResultadosNombreDetalle_${clave_colonia}_${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ValidacionResultadosNombreDetalle: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ValidacionResultadosNombreDetalle: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F8 - Mesas Con Cómputo Capturado

export const MesasConComputo = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = (await SICOVACC.sequelize.query(`SELECT UPPER(D.nombre_delegacion) AS nombre_delegacion, M.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(M.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa
        FROM consulta_mros M
        LEFT JOIN consulta_tipo_mesa_V TP ON M.tipo_mro = TP.tipo_mro
        LEFT JOIN consulta_cat_delegacion D ON M.id_delegacion = D.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 C ON M.clave_colonia = C.clave_colonia
        WHERE M.${campo} = 1 AND M.id_distrito = ${id_distrito} AND EXISTS (SELECT 1 FROM consulta_actas WHERE modalidad = 1 AND estatus = 1 AND anio = ${anio} AND clave_colonia = M.clave_colonia AND num_mro = M.num_mro AND tipo_mro = M.tipo_mro)
        ORDER BY D.nombre_delegacion, C.nombre_colonia, M.num_mro, M.tipo_mro ASC`))[0];
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
                worksheet.spliceColumns(1, 1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'MESAS RECEPTORAS DE OPINIÓN CON CÓMPUTO CAPTURADO';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:C2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:C3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:C5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:C6');
                worksheet.getCell('C7').value = 'FORMATO 8';
                worksheet.getCell('A10').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('A10').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('C9').value = `Fecha: ${fecha}`;
                worksheet.getCell('C10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                resultado.forEach(res => {
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 2).value = 'TOTAL';
                worksheet.getCell(fila, 2).style = fill;
                worksheet.getCell(fila, 3).value = resultado.length;
                worksheet.getCell(fila, 3).style = { ...fill, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 0) {
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

//? F9 - Mesas Sin Cómputo Capturado

export const MesasSinComputo = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = (await SICOVACC.sequelize.query(`SELECT UPPER(D.nombre_delegacion) AS nombre_delegacion, M.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(M.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa
        FROM consulta_mros M
        LEFT JOIN consulta_tipo_mesa_V TP ON M.tipo_mro = TP.tipo_mro
        LEFT JOIN consulta_cat_delegacion D ON M.id_delegacion = D.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 C ON M.clave_colonia = C.clave_colonia
        WHERE M.${campo} = 1 AND M.id_distrito = ${id_distrito} AND NOT EXISTS (SELECT 1 FROM consulta_actas WHERE modalidad = 1 AND estatus = 1 AND anio = ${anio} AND clave_colonia = M.clave_colonia AND num_mro = M.num_mro AND tipo_mro = M.tipo_mro)
        ORDER BY D.nombre_delegacion, C.nombre_colonia, M.num_mro, M.tipo_mro ASC`))[0];
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
                worksheet.spliceColumns(1, 1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                worksheet.getCell('B3').value = titulos[1];
                worksheet.getCell('B5').value = subtitulo;
                worksheet.getCell('B6').value = 'MESAS RECEPTORAS DE OPINIÓN SIN CÓMPUTO CAPTURADO';
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:C2');
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:C3');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:C5');
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:C6');
                worksheet.getCell('C7').value = 'FORMATO 9';
                worksheet.getCell('A10').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('A10').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('C9').value = `Fecha: ${fecha}`;
                worksheet.getCell('C10').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                resultado.forEach(res => {
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 2).value = 'TOTAL';
                worksheet.getCell(fila, 2).style = fill;
                worksheet.getCell(fila, 3).value = resultado.length;
                worksheet.getCell(fila, 3).style = { ...fill, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 0) {
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

//? F10 - Resultados de Opiniones por Mesa

export const ResultadosOpiMesa = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const actas = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT A.id_distrito, UPPER(D.nombre_delegacion) AS nombre_delegacion, A.clave_colonia, UPPER(C.nombre_colonia) AS nombre_colonia, CONCAT(A.num_mro, NULLIF(CONCAT(' ', TP.mesa), '')) AS mesa, A.num_mro, A.tipo_mro, A.modalidad, A.anio, A.bol_nulas
            FROM consulta_actas A
            INNER JOIN consulta_cat_delegacion D ON A.id_delegacion = D.id_delegacion
            INNER JOIN consulta_cat_colonia_cc1 C ON A.clave_colonia = C.clave_colonia
            INNER JOIN consulta_tipo_mesa_V TP ON A.tipo_mro = TP.tipo_mro
            WHERE A.id_distrito = ${id_distrito} AND A.estatus = 1 AND A.anio = ${anio}
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
                SELECT secuencial, nom_proyecto, votos, votos_sei, total_votos
                FROM consulta_actas_VVS V2
                WHERE V2.id_distrito = V1.id_distrito AND V2.clave_colonia = V1.clave_colonia AND V2.num_mro = V1.num_mro AND V2.tipo_mro = V1.tipo_mro AND V2.anio = V1.anio
                ORDER BY secuencial ASC
                FOR JSON PATH
            ) AS proyectos
            FROM consulta_actas_VVS V1
            GROUP BY id_distrito, clave_colonia, num_mro, tipo_mro, anio
        )
        SELECT A1.nombre_delegacion, A1.clave_colonia, A1.nombre_colonia, A1.mesa, P.proyectos, A1.bol_nulas, COALESCE(A2.bol_nulas, 0) AS bol_nulas_sei
        FROM CA A1
        LEFT JOIN CA A2 ON A1.id_distrito = A2.id_distrito AND A1.clave_colonia = A2.clave_colonia AND A1.num_mro = A2.num_mro AND A1.tipo_mro = A2.tipo_mro AND A2.modalidad = 2
        LEFT JOIN ProyectosJSON P ON A1.id_distrito = P.id_distrito AND A1.clave_colonia = P.clave_colonia AND A1.num_mro = P.num_mro AND A1.tipo_mro = P.tipo_mro AND A1.anio = P.anio
        WHERE A1.modalidad = 1 AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = A1.id_distrito AND clave_colonia = A1.clave_colonia)
        ORDER BY A1.nombre_delegacion, A1.nombre_colonia, A1.num_mro, A1.tipo_mro ASC`))[0];
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
                worksheet.spliceColumns(1, 1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:I2');
                worksheet.getCell('A3').value = titulos[1];
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:I3');
                worksheet.getCell('A5').value = subtitulo;
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:I5');
                worksheet.getCell('A6').value = 'RESULTADOS DE OPINIONES POR MESA';
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:I6');
                worksheet.getCell('A7').value = 'Dirección Distrital';
                worksheet.getCell('A7').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('B7').value = id_distrito;
                worksheet.getCell('B7').style = { ...contenidoStyle, font: { ...contenidoStyle.font, size: 12 } };
                worksheet.getCell('I4').value = 'FORMATO 10';
                worksheet.getCell('I7').value = fecha;
                worksheet.getCell('I8').value = hora.substring(0, hora.length - 3);
                worksheet.getCell('D9').value = 'Clave del Proyecto'
                worksheet.getCell('F9').value = 'Nombre del Proyecto Específico';
                for (let acta of actas) {
                    let sum_votos = 0, sum_votos_sei = 0;
                    const { nombre_delegacion, clave_colonia, nombre_colonia, mesa, proyectos, bol_nulas, bol_nulas_sei } = acta;
                    sum_votos += bol_nulas, sum_votos_sei += bol_nulas_sei;
                    for (let proyecto of JSON.parse(proyectos)) {
                        const { secuencial, nom_proyecto, votos, votos_sei, total_votos } = proyecto;
                        sum_votos += votos, sum_votos_sei += votos_sei;
                        const X = { nombre_delegacion, clave_colonia, nombre_colonia, secuencial, mesa, nom_proyecto, votos, votos_sei, total_votos };
                        Object.keys(X).forEach((key, i) => {
                            worksheet.getCell(fila, i + 1).value = X[key];
                            worksheet.getCell(fila, i + 1).style = [3, 6, 7, 8].includes(i) ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                        });
                        fila++;
                    }
                    for (let i = 1; i <= 9; i++) {
                        worksheet.getCell(fila, i).style = i <= 5 ? contenidoStyle : i == 6 ? { ...fill, font: { ...fill.font, bold: false } } : { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                        worksheet.getCell(fila + 1, i).style = i <= 5 ? contenidoStyle : i == 6 ? { ...fill, font: { ...fill.font, bold: false } } : { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                    }
                    worksheet.getCell(fila, 6).value = 'Opiniones Nulas';
                    worksheet.getCell(fila, 7).value = bol_nulas;
                    worksheet.getCell(fila, 8).value = bol_nulas_sei;
                    worksheet.getCell(fila, 9).value = bol_nulas + bol_nulas_sei;
                    worksheet.getCell(fila + 1, 6).value = 'Total por Mesa';
                    worksheet.getCell(fila + 1, 7).value = sum_votos;
                    worksheet.getCell(fila + 1, 8).value = sum_votos_sei;
                    worksheet.getCell(fila + 1, 9).value = sum_votos + sum_votos_sei;
                    fila += 2;
                }
                worksheet.columns.forEach((column, i) => {
                    if ([0, 2, 5].includes(i)) {
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
                    reporte: `Reporte_ResultadosOpiMesa-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ResultadosOpiMesa: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ResultadosOpiMesa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? F11 - Proyectos por Unidad Territorial que Obtuvieron el Primer Lugar

export const ProyectosPrimerLugar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia
            FROM consulta_actas
            WHERE modalidad = 1 AND anio = ${anio} AND id_distrito = ${id_distrito}
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
        SELECT nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM RANKING
        WHERE DR = 1 AND empate <= 1 AND total_votos > 0
        ORDER BY nombre_delegacion, nombre_colonia, secuencial ASC`))[0];
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
                worksheet.spliceColumns(1, 1);
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
                worksheet.getCell('I4').value = 'FORMATO 11';
                worksheet.getCell('A8').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('A8').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('I7').value = fecha;
                worksheet.getCell('I8').value = hora.substring(0, hora.length - 3);
                let fila = 11;
                let colonias = [];
                proyectos.forEach(res => {
                    if (!colonias.includes(res.clave_colonia))
                        colonias.push(res.clave_colonia);
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = index >= 6 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                });
                if (!worksheet.getCell(fila, 1).isMerged)
                    worksheet.mergeCells(fila, 1, fila, 2);
                for (let i = 1; i <= 9; i++)
                    worksheet.getCell(fila, i).style = i == 3 || i == 6 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                worksheet.getCell(fila, 1).value = 'Total de Unidades Territoriales';
                worksheet.getCell(fila, 1).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 3).value = colonias.length;
                worksheet.getCell(fila, 5).value = 'Total de Proyectos';
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 6).value = proyectos.length;
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 2 || index == 4 || index == 5) {
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
                return workbook.xlsx.writeBuffer()
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

//? F12 - Proyectos por Unidad Territorial que Obtuvieron el Segundo Lugar

export const ProyectosSegundoLugar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia
            FROM consulta_actas
            WHERE modalidad = 1 AND anio = ${anio} AND id_distrito = ${id_distrito}
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
        SELECT nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM RANKING
        WHERE DR = 2 AND empate <= 1 AND total_votos > 0
        ORDER BY nombre_delegacion, nombre_colonia, secuencial ASC`);
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
                worksheet.spliceColumns(1, 1);
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
                worksheet.getCell('I4').value = 'FORMATO 12';
                worksheet.getCell('A8').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('A8').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('I7').value = fecha;
                worksheet.getCell('I8').value = hora.substring(0, hora.length - 3);
                let fila = 11;
                let colonias = [];
                proyectos[0].forEach(res => {
                    if (!colonias.includes(res.clave_colonia))
                        colonias.push(res.clave_colonia);
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = index >= 6 && index <= 8 ? Intl.NumberFormat('es-MX').format(res[key]) : res[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                if (!worksheet.getCell(fila, 1).isMerged)
                    worksheet.mergeCells(fila, 1, fila, 2);
                for (let i = 1; i <= 9; i++)
                    worksheet.getCell(fila, i).style = contenidoStyle;
                worksheet.getCell(fila, 1).value = 'Total de Unidades Territoriales';
                worksheet.getCell(fila, 1).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 3).value = Intl.NumberFormat('es-MX').format(colonias.length);
                worksheet.getCell(fila, 5).value = 'Total de Proyectos';
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 6).value = Intl.NumberFormat('es-MX').format(proyectos[1]);
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 2 || index == 4 || index == 5) {
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

//? F13 - Proyectos Empatados que Obtuvieron el Primer Lugar

export const ProyectosEmpatePrimerLugar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia
            FROM consulta_actas
            WHERE modalidad = 1 AND anio = ${anio} AND id_distrito = ${id_distrito}
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
        SELECT nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM RANKING
        WHERE DR = 1 AND empate > 1 AND total_votos > 0
        ORDER BY nombre_delegacion, nombre_colonia, secuencial ASC`))[0];
        if (!proyectos.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const titulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas, 'Proyectos-GE.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                worksheet.spliceColumns(1, 1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = 'DIRECCIÓN EJECUTIVA DE ORGANIZACIÓN ELECTORAL Y GEOSTADÍSTICA';
                worksheet.getCell('B2').style = { ...tituloStyle, font: { ...tituloStyle.font, size: 16 } };
                worksheet.getCell('B4').value = titulo;
                worksheet.getCell('B4').style = { ...tituloStyle, font: { ...tituloStyle.font, size: 14 } };
                worksheet.getCell('B5').value = 'CASOS DE EMPATES DE LOS PROYECTOS QUE OBTUVIERON EL PRIMER LUGAR';
                worksheet.getCell('B5').style = { ...tituloStyle, font: { ...tituloStyle.font, size: 14 } };
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:I2');
                if (!worksheet.getCell('B4').isMerged)
                    worksheet.mergeCells('B4:I4');
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:I5');
                worksheet.getCell('I3').value = 'FORMATO 13';
                worksheet.getCell('A7').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('A7').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('I6').value = fecha;
                worksheet.getCell('I7').value = hora.substring(0, hora.length - 3);
                let fila = 10;
                let colonias = [];
                proyectos.forEach(res => {
                    if (!colonias.includes(res.clave_colonia))
                        colonias.push(res.clave_colonia);
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = index >= 6 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                });
                if (!worksheet.getCell(fila, 1).isMerged)
                    worksheet.mergeCells(fila, 1, fila, 2);
                for (let i = 1; i <= 9; i++)
                    worksheet.getCell(fila, i).style = i == 3 || i == 6 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                worksheet.getCell(fila, 1).value = 'Total de Unidades Territoriales';
                worksheet.getCell(fila, 1).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 3).value = colonias.length;
                worksheet.getCell(fila, 5).value = 'Total de Proyectos';
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 6).value = proyectos.length;
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 2 || index == 4 || index == 5) {
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

//? F14 - Proyectos Empatados que Obtuvieron el Segundo Lugar

export const ProyectosEmpateSegundoLugar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = await SICOVACC.sequelize.query(`;WITH CA AS (
            SELECT id_distrito, clave_colonia
            FROM consulta_actas
            WHERE modalidad = 1 AND anio = ${anio} AND id_distrito = ${id_distrito}
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
        SELECT nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM RANKING
        WHERE DR = 2 AND empate > 1 AND total_votos > 0
        ORDER BY nombre_delegacion, nombre_colonia, secuencial ASC`);
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
                worksheet.spliceColumns(1, 1);
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
                worksheet.getCell('I4').value = 'FORMATO 14';
                worksheet.getCell('A8').value = `Dirección Distrital: ${id_distrito}`;
                worksheet.getCell('A8').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('I7').value = fecha;
                worksheet.getCell('I8').value = hora.substring(0, hora.length - 3);
                let fila = 11;
                let colonias = [];
                proyectos[0].forEach(res => {
                    if (!colonias.includes(res.clave_colonia))
                        colonias.push(res.clave_colonia);
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = index >= 6 && index <= 8 ? Intl.NumberFormat('es-MX').format(res[key]) : res[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                })
                if (!worksheet.getCell(fila, 1).isMerged)
                    worksheet.mergeCells(fila, 1, fila, 2);
                for (let i = 1; i <= 9; i++)
                    worksheet.getCell(fila, i).style = contenidoStyle;
                worksheet.getCell(fila, 1).value = 'Total de Unidades Territoriales';
                worksheet.getCell(fila, 1).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 3).value = Intl.NumberFormat('es-MX').format(colonias.length);
                worksheet.getCell(fila, 5).value = 'Total de Proyectos';
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 6).value = Intl.NumberFormat('es-MX').format(proyectos[1]);
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 2 || index == 4 || index == 5) {
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

//? F15 - Unidades Territoriales que NO Recibieron Opiniones

export const ProyectosUTSinOpiniones = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const campo = aniosCAT[0][anio];
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`;WITH Mesas AS (
            SELECT id_distrito, clave_colonia, modalidad, anio
            FROM consulta_actas CA
            WHERE estatus = 1 AND votacion_total_emitida = 0
            GROUP BY id_distrito, clave_colonia, modalidad, anio
            HAVING COUNT(*) = (
                SELECT COUNT(*)
                FROM consulta_mros
                WHERE id_distrito = CA.id_distrito AND clave_colonia = CA.clave_colonia AND ${campo} = 1
            )
        ),
        V AS (
            SELECT id_distrito, nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos, anio
            FROM consulta_actas_VVS
            WHERE anio = ${anio} AND id_distrito = ${id_distrito}
        )
        SELECT nombre_delegacion, clave_colonia, nombre_colonia, secuencial, rubro_general, nom_proyecto, votos, votos_sei, total_votos
        FROM V
        WHERE EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = V.id_distrito AND clave_colonia = V.clave_colonia AND anio = V.anio AND modalidad = 1)
        AND EXISTS (SELECT 1 FROM Mesas WHERE id_distrito = V.id_distrito AND clave_colonia = V.clave_colonia AND anio = V.anio AND modalidad = 2)
        ORDER BY nombre_delegacion, nombre_colonia, secuencial ASC`))[0];
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
                worksheet.spliceColumns(1, 1);
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A2').value = titulos[0];
                worksheet.getCell('A3').value = titulos[1];
                worksheet.getCell('A5').value = subtitulo;
                worksheet.getCell('A6').value = 'CONCENTRADO DE UNIDADES TERRITORIALES QUE NO RECIBIERON OPINIONES EN NINGUNO DE SUS PROYECTOS SOMETIDOS A OPINIÓN';
                if (!worksheet.getCell('A2').isMerged)
                    worksheet.mergeCells('A2:I2');
                if (!worksheet.getCell('A3').isMerged)
                    worksheet.mergeCells('A3:I3');
                if (!worksheet.getCell('A5').isMerged)
                    worksheet.mergeCells('A5:I5');
                if (!worksheet.getCell('A6').isMerged)
                    worksheet.mergeCells('A6:I6');
                worksheet.getCell('A8').value = 'Dirección Distrital:';
                worksheet.getCell('A8').style = { ...fill, font: { ...fill.font, size: 12 } };
                worksheet.getCell('B8').value = id_distrito;
                worksheet.getCell('B8').style = contenidoStyle;
                worksheet.getCell('I8').value = fecha;
                worksheet.getCell('I9').value = hora.substring(0, hora.length - 3);
                worksheet.getCell('I10').value = 'FORMATO 15';
                let fila = 12;
                proyectos.forEach(res => {
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = index >= 6 ? { ...contenidoStyle, numFmt: '#,##0' } : contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 5).value = 'Total';
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 6).value = proyectos.length;
                worksheet.getCell(fila, 6).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 2 || index == 4 || index == 5) {
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
                return workbook.xlsx.writeBuffer()
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
                console.error(`Error al procesar el archivo Excel en ProyectosUTSinOpiniones: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ProyectosUTSinOpiniones: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Proyectos a Opinar   

export const ProyectosOpinar = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const proyectos = (await SICOVACC.sequelize.query(`SELECT UPPER(CCD.nombre_delegacion) AS nombre_delegacion, CPP.clave_colonia, UPPER(CCC.nombre_colonia) AS nombre_colonia, UPPER(CPP.folio_proy_web) AS folio_proy_web, CPP.num_proyecto, UPPER(CPP.nom_proyecto) AS nom_proyecto
        FROM consulta_prelacion_proyectos CPP
        LEFT JOIN consulta_cat_delegacion CCD ON CPP.id_delegacion = CCD.id_delegacion
        LEFT JOIN consulta_cat_colonia_cc1 CCC ON CPP.clave_colonia = CCC.clave_colonia
        WHERE CPP.estatus = 1 AND CPP.anio = ${anio} AND CPP.id_distrito = ${id_distrito} AND CPP.clave_colonia IN (SELECT clave_colonia FROM consulta_mros WHERE estatus = 1)
        ORDER BY CCD.nombre_delegacion, CCC.nombre_colonia, CPP.num_proyecto ASC`))[0];
        if (!proyectos.length)
            return res.status(404).json({
                success: false,
                msg: '¡No existe información!'
            });
        const { fecha, hora } = await FechaServer();
        const titulo = `CONSULTA DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`;
        workbook.xlsx.readFile(path.join(plantillas[2], 'Proyectos_Opinar.xlsx'))
            .then(() => {
                workbook.creator = autor;
                const worksheet = workbook.getWorksheet(1);
                let fila = 11;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('A5').value = titulo;
                worksheet.getCell('A6').value = 'PROYECTOS A OPINAR';
                worksheet.getCell('B7').value = id_distrito;
                worksheet.getCell('F7').value = `Fecha: ${fecha}`;
                worksheet.getCell('F8').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                proyectos.forEach(proyecto => {
                    Object.keys(proyecto).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = proyecto[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 2).value = 'Total:';
                worksheet.getCell(fila, 2).style = fill;
                worksheet.getCell(fila, 3).value = proyectos.length;
                worksheet.getCell(fila, 3).style = { ...contenidoStyle, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 0 || index == 2 || index == 5) {
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
                    reporte: `Reporte_ProyectosOpinar-${fecha}-${hora}.xlsx`,
                    buffer
                });
            })
            .catch(err => {
                console.error(`Error al procesar el archivo Excel en ProyectosOpinar: ${err}`);
                res.status(500).json({
                    success: false,
                    msg: `Error al procesar el archivo Excel: ${err}`
                });
            });
    } catch (err) {
        console.error(`Error en ProyectosOpinar: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al generar el reporte'
        });
    }
}

//? Levantada en Distrito

export const LevantadaDistrito = async (req = request, res = response) => {
    const { id_distrito } = req.params;
    const { anio } = req.query;
    const workbook = new ExcelJs.Workbook();
    try {
        const resultado = await SICOVACC.sequelize.query(`
            SELECT
                ROW_NUMBER() OVER(ORDER BY CCC.nombre_colonia, CCC.clave_colonia, CA.num_mro, CA.razon_distrital ASC) AS Consecutivo,
                UPPER(CCC.nombre_colonia) AS nombre_colonia, 
                UPPER(CCC.clave_colonia) AS clave_colonia, 
                CA.num_mro, 
            UPPER(dbo.RazonDistrital(CA.razon_distrital)) AS razon_distrital
            FROM consulta_actas CA
            LEFT JOIN consulta_cat_colonia_cc1 CCC ON CA.clave_colonia = CCC.clave_colonia
            WHERE CA.modalidad = 1 AND CA.anio = ${anio} AND CA.levantada_distrito = 1 AND CCC.estatus = 1 AND CA.id_distrito = ${id_distrito}
            ORDER BY CCC.nombre_colonia, CA.num_mro ASC
        `);
        if (resultado[1] == 0)
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
                worksheet.spliceColumns(1, 1);
                let fila = 11;
                const iecm = workbook.addImage({ filename: iecmLogo, extension: 'png' });
                worksheet.addImage(iecm, { tl: { col: 0, row: 0 }, ext: { width: 231, height: 140 }, editAs: 'absolute' });
                worksheet.getCell('B2').value = titulos[0];
                if (!worksheet.getCell('B2').isMerged)
                    worksheet.mergeCells('B2:E2');
                worksheet.getCell('B3').value = titulos[1];
                if (!worksheet.getCell('B3').isMerged)
                    worksheet.mergeCells('B3:E3');
                worksheet.getCell('B5').value = subtitulo;
                if (!worksheet.getCell('B5').isMerged)
                    worksheet.mergeCells('B5:E5');
                worksheet.getCell('B6').value = 'ACTAS LEVANTADAS EN DIRECCIÓN DISTRITAL (CAUSALES DE RECUENTO)';
                if (!worksheet.getCell('B6').isMerged)
                    worksheet.mergeCells('B6:E6');
                worksheet.getCell('E8').value = `Fecha: ${fecha}`;
                worksheet.getCell('E9').value = `Hora: ${hora.substring(0, hora.length - 3)}`;
                worksheet.getCell('A8').value = `Dirección Distrital:`;
                worksheet.getCell('B8').value = id_distrito;
                worksheet.getCell('B8').style = contenidoStyle;
                worksheet.getCell('A8').style = { ...fill, font: { ...fill.font, size: 12 } };
                resultado[0].forEach(res => {
                    Object.keys(res).forEach((key, index) => {
                        worksheet.getCell(fila, index + 1).value = res[key];
                        worksheet.getCell(fila, index + 1).style = contenidoStyle;
                    });
                    fila++;
                });
                worksheet.getCell(fila, 4).value = 'Total: ';
                worksheet.getCell(fila, 4).style = { ...fill, font: { ...fill.font, bold: false } };
                worksheet.getCell(fila, 5).value = resultado[1];
                worksheet.getCell(fila, 5).style = { ...fill, font: { ...fill.font, bold: false }, numFmt: '#,##0' };
                worksheet.columns.forEach((column, index) => {
                    if (index == 1 || index == 4) {
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