import { request, response } from 'express';
import { Audit } from '../helpers/Audit.js';
import { TipoMesa } from '../helpers/Constantes.js';
import { ConsultaClaveColonia, ConsultaDelegacion, ConsultaExistenciaActas, ConsultaVerificaInicioCierre, ConsultaVerificaProyectos } from '../helpers/Consultas.js';
import { Comillas, EncryptData, LetrasANumero } from '../helpers/Funciones.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Estado de la Base de Datos

export const EstadoBaseDatos = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    try {
        const [verifica, infoG, conteoG] = await Promise.all([
            ConsultaVerificaInicioCierre(id_distrito),
            SICOVACC.sequelize.query(`EXEC InfoGeneral ${id_distrito}`),
            SICOVACC.sequelize.query(`EXEC ConteoGeneral ${id_distrito}`)
        ]);
        res.json({
            success: true,
            datos: {
                datosSEI: infoG[0][0].datosSEI,
                ...verifica,
                mesasNI: infoG[0][0].mesasNI,
                incidentes: {
                    incidentes_C: infoG[0][0].incidentes_C,
                    incidentes_CC1: infoG[0][0].incidentes_CC1,
                    incidentes_CC2: infoG[0][0].incidentes_CC2
                },
                conteo: {
                    conteo_C: { ...(conteoG[0][0]) },
                    conteo_CC1: { ...(conteoG[0][1]) },
                    conteo_CC2: { ...(conteoG[0][2]) }
                }
            }
        });
    } catch (err) {
        console.error(`Error en EstadoBaseDatos: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

//? Inicio de Validación

export const DatosInicioValidacion = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT inicio_asistencia1 AS MSPEN, inicio_asistencia2 AS COPACO, inicio_asistencia3 AS personasCandidatas, inicio_asistencia4 AS personasObservadoras, inicio_asistencia5 AS presentaronProyecto,
        inicio_asistencia6 AS mediosComunicacion, inicio_asistencia7 AS otros, inicio_total AS total,
        CONVERT(VARCHAR(10), fecha_hora_inicio, 20) AS fecha, CONVERT(VARCHAR(5), fecha_hora_inicio, 8) AS hora, inicio_observaciones AS observaciones FROM consulta_computo WHERE estatus = 1 AND id_distrito = ${id_distrito}`))[0][0];
        if (datos[1] == 0)
            return res.status(404).json({
                success: false,
                msg: 'No hay información'
            });
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en DatosInicioValidacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const GuardarInicioValidacion = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { MSPEN, COPACO, personasCandidatas, personasObservadoras, presentaronProyecto, mediosComunicacion, otros, total, fecha, hora, observaciones } = req.body;
    try {
        const { inicioValidacion, cierreValidacion } = await ConsultaVerificaInicioCierre(id_distrito);
        if (inicioValidacion)
            return res.status(400).json({
                success: false,
                msg: 'El inicio de validación ya se inicio'
            });
        if (cierreValidacion)
            return res.status(400).json({
                success: false,
                msg: 'El cierre de validación ya se hizo, no se puede modificar'
            });
        await SICOVACC.sequelize.query(`INSERT consulta_computo (id_distrito, inicio_asistencia1, inicio_asistencia2, inicio_asistencia3, inicio_asistencia4, inicio_asistencia5, inicio_asistencia6, inicio_asistencia7, inicio_total, fecha_hora_inicio, inicio_observaciones, estatus, fecha_alta)
        VALUES (${id_distrito}, ${MSPEN}, ${COPACO}, ${personasCandidatas ? personasCandidatas : 'NULL'}, ${personasObservadoras}, ${presentaronProyecto}, ${mediosComunicacion}, ${otros}, ${total}, '${fecha} ${hora}:00', UPPER('${Comillas(observaciones)}'), 1, CURRENT_TIMESTAMP)`);
        await Audit(id_transaccion, id_usuario, id_distrito, 'REALIZÓ EL INICIO DE VALIDACIÓN');
        res.json({
            success: true,
            msg: 'Información guardada',
            inicioValidacion: !inicioValidacion
        });
    } catch (err) {
        console.error(`Error en GuardarInicioValidacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const ActualizarInicioValidacion = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { MSPEN, COPACO, personasCandidatas, personasObservadoras, presentaronProyecto, mediosComunicacion, otros, total, fecha, hora, observaciones } = req.body;
    try {
        // const { cierreValidacion } = await ConsultaVerificaInicioCierre(id_distrito);
        // if (cierreValidacion)
        //     return res.status(400).json({
        //         success: false,
        //         msg: 'El cierre de validación ya se hizo, no se puede modificar'
        //     });
        await SICOVACC.sequelize.query(`UPDATE consulta_computo SET inicio_asistencia1 = ${MSPEN}, inicio_asistencia2 = ${COPACO}, inicio_asistencia3 = ${personasCandidatas ? personasCandidatas : 'NULL'}, inicio_asistencia4 = ${personasObservadoras}, inicio_asistencia5 = ${presentaronProyecto}, inicio_asistencia6 = ${mediosComunicacion}, inicio_asistencia7 = ${otros}, inicio_total = ${total},
        fecha_hora_inicio = '${fecha} ${hora}:00', inicio_observaciones = '${Comillas(observaciones)}', fecha_modif = CURRENT_TIMESTAMP WHERE estatus = 1 AND id_distrito = ${id_distrito}`);
        await Audit(id_transaccion, id_usuario, id_distrito, 'ACTUALIZÓ LA INFORMACIÓN DEL INICIO DE VALIDACIÓN');
        res.json({
            success: true,
            msg: 'Información guardada'
        });
    } catch (err) {
        console.error(`Error en ActualizarInicioValidacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

//? Mesas Instaladas

export const MesasInstaladas = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT UPPER(CCC.nombre_colonia) AS nombre_colonia, CM.clave_colonia, CM.num_mro, CM.tipo_mro,
        CAST(CASE WHEN CM.estatus = 1 THEN 0 ELSE 1 END AS BIT) AS noInstalada
        FROM consulta_mros CM
        LEFT JOIN consulta_cat_colonia_cc1 CCC ON CM.clave_colonia = CCC.clave_colonia
        WHERE CCC.nombre_colonia IS NOT NULL AND CM.estatus IN (0, 1) AND CM.id_distrito = ${id_distrito}
        ORDER BY CCC.nombre_colonia`))[0];
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en MesasInstaladas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const GuardarMesasInstaladas = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { mesas } = req.body;
    try {
        for (let mesa of mesas) {
            const { clave_colonia, num_mro, tipo_mro, noInstalada } = mesa;
            await SICOVACC.sequelize.query(`UPDATE consulta_mros SET estatus = ${noInstalada ? 0 : 1} WHERE clave_colonia = '${clave_colonia}' AND num_mro = '${num_mro}' AND tipo_mro = ${tipo_mro}`);
        }
        await Audit(id_transaccion, id_usuario, id_distrito, 'ACTUALIZÓ EL ESTADO DE LAS MESAS');
        res.json({
            success: true,
            msg: 'Información guardada'
        });
    } catch (err) {
        console.error(`Error en GuardarMesasInstaladas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

//? Registros de Incidentes

export const RegistrosIncidentes = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { anio } = req.query;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT CI.id_incidente, CI.id_distrito, CI.id_delegacion, CI.id_colonia, CI.clave_colonia, UPPER(CCC.nombre_colonia) AS nombre_colonia, CI.num_mro, CI.tipo_mro, UPPER(CCD.nombre_delegacion) AS nombre_delegacion,
        CONCAT('M', RIGHT('00' + num_mro, 2)) AS mro, CI.incidente_1, CI.incidente_2, CI.incidente_3, CI.incidente_4, CI.incidente_5, CI.incidente_6, CI.incidente_7, CI.incidente_8,
        CONVERT(VARCHAR(10), CI.fecha_hora, 20) AS fecha, CONVERT(VARCHAR(5), CI.fecha_hora, 8) AS hora, UPPER(REPLACE(REPLACE(REPLACE(REPLACE(CI.participantes, 'µþ34µþ', CHAR(34)), '@@39@@', CHAR(39)), 'ËÈ13ËÈ', CHAR(13)), 'ÎÏ10ÎÏ', CHAR(10))) AS participantes,
        UPPER(REPLACE(REPLACE(REPLACE(REPLACE(CI.hechos, 'µþ34µþ', CHAR(34)), '@@39@@', CHAR(39)), 'ËÈ13ËÈ', CHAR(13)), 'ÎÏ10ÎÏ', CHAR(10))) AS hechos, UPPER(REPLACE(REPLACE(REPLACE(REPLACE(CI.acciones, 'µþ34µþ', CHAR(34)), '@@39@@', CHAR(39)), 'ËÈ13ËÈ', CHAR(13)), 'ÎÏ10ÎÏ', CHAR(10))) AS acciones
        FROM consulta_incidentes CI
        LEFT JOIN consulta_cat_colonia_cc1 CCC ON CI.clave_colonia = CCC.clave_colonia
        LEFT JOIN consulta_cat_delegacion CCD ON CI.id_delegacion = CCD.id_delegacion
        WHERE CI.estatus = 1 AND CI.anio = ${anio} AND CI.id_distrito = ${id_distrito}
        ORDER BY CCC.nombre_colonia, CI.num_mro`))[0];
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en RegistrosIncidentes: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const GuardarIncidente = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { clave_colonia, num_mro, tipo_mro, incidente_1, incidente_2, incidente_3, incidente_4, incidente_5, incidente_6, incidente_7, incidente_8, fecha, hora, participantes, hechos, acciones, anio } = req.body;
    try {
        const { id_delegacion } = await ConsultaDelegacion(id_distrito, clave_colonia);
        const { id_colonia } = await ConsultaClaveColonia(clave_colonia);
        await SICOVACC.sequelize.query(`INSERT consulta_incidentes (id_distrito, id_delegacion, id_colonia, clave_colonia, num_mro, tipo_mro, incidente_1, incidente_2, incidente_3, incidente_4, incidente_5, incidente_6, incidente_7, incidente_8, fecha_hora, participantes, hechos, acciones, anio, estatus, fecha_alta, id_usuario)
        VALUES (${id_distrito}, ${id_delegacion}, ${id_colonia}, '${clave_colonia}', '${num_mro}', ${tipo_mro}, ${incidente_1 ? 1 : 0}, ${incidente_2 ? 1 : 0}, ${incidente_3 ? 1 : 0}, ${incidente_4 ? 1 : 0}, ${incidente_5 ? 1 : 0}, 0, 0, 0, '${fecha} ${hora.padEnd(8, ':00')}', UPPER('${Comillas(participantes)}'), UPPER('${Comillas(hechos)}'), UPPER('${Comillas(acciones)}'), '${anio}', 1, CURRENT_TIMESTAMP, ${id_usuario})`);
        await Audit(id_transaccion, id_usuario, id_distrito, `REGISTRÓ UN INCIDENTE DE LA ${anio == 1 ? 'ELECCIÓN' : 'CONSULTA'}, EN LA UT ${clave_colonia}, MESA M${String(num_mro).padStart(2, '0')}${tipo_mro != 1 ? `, DE TIPO DE MESA ${TipoMesa(tipo_mro)}` : ''}`);
        res.json({
            success: true,
            msg: 'Incidente Registrado'
        });
    } catch (err) {
        console.error(`Error en GuardarIncidente: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const EditarIncidente = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { id_incidente, num_mro, tipo_mro, incidente_1, incidente_2, incidente_3, incidente_4, incidente_5, incidente_6, incidente_7, incidente_8, fecha, hora, participantes, hechos, acciones } = req.body;
    try {
        const { clave_colonia, mro } = (await SICOVACC.sequelize.query(`SELECT clave_colonia, num_mro AS mro FROM consulta_incidentes WHERE id_incidente = ${id_incidente}`))[0][0];
        await SICOVACC.sequelize.query(`UPDATE consulta_incidentes SET num_mro = ${num_mro}, tipo_mro = ${tipo_mro}, incidente_1 = ${incidente_1 ? 1 : 0}, incidente_2 = ${incidente_2 ? 1 : 0}, incidente_3 = ${incidente_3 ? 1 : 0}, incidente_4 = ${incidente_4 ? 1 : 0}, incidente_5 = ${incidente_5 ? 1 : 0}, incidente_6 = 0, incidente_7 = 0, incidente_8 = 0,
        fecha_hora = '${fecha} ${hora.padEnd(8, ':00')}', participantes = UPPER('${Comillas(participantes)}'), hechos = UPPER('${Comillas(hechos)}'), acciones = UPPER('${Comillas(acciones)}'), fecha_modif = CURRENT_TIMESTAMP
        WHERE estatus = 1 AND id_incidente = ${id_incidente}`);
        await Audit(id_transaccion, id_usuario, id_distrito, `EDITÓ EL INICIDENTE DE LA ${anio == 1 ? 'ELECCIÓN' : 'CONSULTA'}, ${id_incidente} DE LA UT ${clave_colonia}, ${num_mro == mro ? `MESA M${String(num_mro).padStart(2, '0')}` : ` CAMBIO A LA MESA M${String(num_mro).padStart(2, '0')}`}${tipo_mro != 1 ? `, DE TIPO DE MESA ${TipoMesa(tipo_mro)}` : ''}`);
        res.json({
            success: true,
            msg: 'Incidente Actualizado'
        });
    } catch (err) {
        console.error(`Error en EditarIncidente: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const EliminarIncidente = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { id_incidente } = req.body;
    try {
        const { clave_colonia, num_mro, tipo_mro } = (await SICOVACC.sequelize.query(`SELECT clave_colonia, num_mro, tipo_mro FROM consulta_incidentes WHERE id_incidente = ${id_incidente}`))[0][0];
        const resp = await SICOVACC.sequelize.query(`DELETE FROM consulta_incidentes WHERE id_incidente = ${id_incidente}`);
        if (resp[1] == 0)
            return res.status(404).json({
                success: false,
                msg: 'Incidente no encontrado'
            });
        await Audit(id_transaccion, id_usuario, id_distrito, `ELIMINÓ EL INCIDENTE DE LA ${anio == 1 ? 'ELECCIÓN' : 'CONSULTA'}, ${id_incidente} DE LA UT ${clave_colonia}, MESA M${String(num_mro).padStart(2, '0')}${tipo_mro != 1 ? `, DE TIPO DE MESA ${TipoMesa(tipo_mro)}` : ''}`);
        res.json({
            success: true,
            msg: 'Incidente Eliminado'
        });
    } catch (err) {
        console.error(`Error en EliminarIncidente: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

//? Captura de Resultados de Consulta por Mesa

export const ResultadoConsultaMesa = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { anio } = req.query;
    try {
        let datos = { actas: [], actasCapturadas: 0, actasPorCapturar: 0, UTValidadas: 0, UTPorValidar: 0 };
        datos.actas = (await SICOVACC.sequelize.query(`SELECT CA.id_acta, UPPER(CCC.nombre_colonia) AS nombre_colonia, CCC.clave_colonia, CONCAT('M', RIGHT('00' + num_mro, 2)) AS num_mro, tipo_mro AS tipo
        FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas CA
        LEFT JOIN consulta_cat_colonia_cc1 CCC ON CA.clave_colonia = CCC.clave_colonia
        WHERE CA.modalidad = 1 AND CA.estatus = 1 AND CA.id_distrito = ${id_distrito}${anio != 1 ? ` AND CA.anio = ${anio}` : ''}
        ORDER BY nombre_colonia, num_mro, tipo_mro`))[0];
        datos = { ...datos, ...(await SICOVACC.sequelize.query(`EXEC Conteo ${anio}, ${id_distrito}`))[0][0] };
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en ResultadoConsultaMesa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const VerificarConsultaMesa = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { clave_colonia, num_mro, tipo_mro, anio } = req.body;
    try {
        if (await ConsultaExistenciaActas(id_distrito, clave_colonia, num_mro, tipo_mro, anio) != 0)
            return res.status(400).json({
                success: false,
                msg: `Esta Unidad Territorial Ya tiene una ${anio == 1 ? 'Acta' : 'Consulta'} En esta Mesa`
            });
        const TOTPROY = await ConsultaVerificaProyectos(id_distrito, clave_colonia, anio, false);
        if (TOTPROY == 0)
            return res.status(400).json({
                success: false,
                msg: `Esta Unidad Territorial NO tiene ${anio == 1 ? 'participantes' : 'proyectos opinados favorables'}`
            });
        if (anio != 1) {
            const TOTPROYSORTEO = await ConsultaVerificaProyectos(id_distrito, clave_colonia, anio, true);
            if (TOTPROY != TOTPROYSORTEO)
                return res.status(400).json({
                    success: false,
                    msg: 'Esta Unidad Territorial tiene proyectos Sin Sortear'
                });
        }
        const datos = JSON.parse((await SICOVACC.sequelize.query(`;WITH I AS (SELECT ${id_distrito} AS id_distrito, '${clave_colonia}' AS clave_colonia, ${num_mro} AS num_mro, ${tipo_mro} AS tipo_mro${anio != 1 ? `, ${anio} AS anio` : ''}),
        C AS (
            SELECT UPPER(D.nombre_delegacion) AS nombre_delegacion, COALESCE(A2.bol_nulas, 0) AS bol_nulas_sei, COALESCE(A2.votacion_total_emitida, 0) AS opi_total_sei, I.id_distrito, I.clave_colonia, I.num_mro, I.tipo_mro${anio != 1 ? ', I.anio' : ''}
            FROM I
            LEFT JOIN consulta_mros M ON I.id_distrito = M.id_distrito AND I.clave_colonia = M.clave_colonia AND I.num_mro = M.num_mro AND I.tipo_mro = M.tipo_mro
            LEFT JOIN ${anio == 1 ? 'copaco' : 'consulta'}_actas A2 ON I.id_distrito = A2.id_distrito AND I.clave_colonia = A2.clave_colonia AND I.num_mro = A2.num_mro AND I.tipo_mro = A2.tipo_mro AND A2.modalidad = 2 AND A2.estatus = 2${anio != 1 ? ' AND I.anio = A2.anio' : ''}
            LEFT JOIN consulta_cat_delegacion D ON M.id_delegacion = D.id_delegacion
        ),
        V AS (
            SELECT secuencial, ${anio == 1 ? 'nombreC' : 'nom_proyecto'} AS nom_p${anio != 1 ? ', rubro_general' : ''}, '' AS votos, votos_sei, id_distrito, clave_colonia, num_mro, tipo_mro${anio != 1 ? ', anio' : ''}
            FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas_VVS
            WHERE estatus = 1
        )
        SELECT (
            SELECT nombre_delegacion, bol_nulas_sei, opi_total_sei, (
                SELECT ${anio == 1 ? 'dbo.NumeroALetras(secuencial) AS ' : ''}secuencial, nom_p${anio != 1 ? ', rubro_general' : ''}, votos, votos_sei
                FROM V
                WHERE id_distrito = C.id_distrito AND clave_colonia = C.clave_colonia AND num_mro = C.num_mro AND tipo_mro = C.tipo_mro${anio != 1 ? ' AND anio = C.anio' : ''}
                ORDER BY V.secuencial ASC
                FOR JSON PATH
            ) AS integraciones
            FROM C
            FOR JSON PATH
        ) AS data`))[0][0].data)[0];
        res.json({
            success: true,
            datos: EncryptData(datos)
        })
    } catch (err) {
        console.error(`Error en VerificarConsultaMesa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const DatosActa = async (req = request, res = response) => {
    const { id_acta } = req.params;
    const { anio } = req.query;
    try {
        const datos = JSON.parse((await SICOVACC.sequelize.query(`;WITH C AS (
            SELECT CA1.id_acta, CA1.clave_colonia, UPPER(CCC.nombre_colonia) AS nombre_colonia, CA1.id_delegacion, UPPER(CCD.nombre_delegacion) AS nombre_delegacion, CA1.num_mro, CA1.tipo_mro, CONCAT('M', RIGHT('00' + CA1.num_mro, 2)) AS mro,
            CA1.coordinador_sino, CASE WHEN CA1.num_integrantes IS NULL THEN '' ELSE CAST(CA1.num_integrantes AS VARCHAR) END AS num_integrantes, CA1.observador_sino, CA1.levantada_distrito, CA1.razon_distrital,
            CA1.bol_recibidas, CA1.bol_adicionales, CA1.total_ciudadanos, CA1.bol_sobrantes, CA1.bol_nulas, COALESCE(CA2.bol_nulas, 0) AS bol_nulas_sei, CA1.votacion_total_emitida AS bol_total_emitidas,
            COALESCE(CA2.votacion_total_emitida, 0) AS opi_total_sei, CA1.id_distrito${anio != 1 ? ', CA1.anio' : ''}
            FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas CA1
            LEFT JOIN ${anio == 1 ? 'copaco' : 'consulta'}_actas CA2 ON CA1.id_distrito = CA2.id_distrito AND CA1.id_delegacion = CA2.id_delegacion AND CA1.clave_colonia = CA2.clave_colonia AND CA1.num_mro = CA2.num_mro AND CA1.tipo_mro = CA2.tipo_mro AND CA2.modalidad = 2${anio != 1 ? ' AND CA1.anio = CA2.anio' : ''}
            LEFT JOIN consulta_cat_colonia_cc1 CCC ON CA1.clave_colonia = CCC.clave_colonia
            LEFT JOIN consulta_cat_delegacion CCD ON CA1.id_delegacion = CCD.id_delegacion
            WHERE CA1.id_acta = ${id_acta} AND CA1.modalidad = 1 AND CA1.estatus = 1${anio != 1 ? ` AND CA1.anio = ${anio}` : ''}
        ),
        V AS (
            SELECT secuencial, ${anio == 1 ? 'nombreC' : 'nom_proyecto'} AS nom_p,${anio != 1 ? ' rubro_general,' : ''} votos, votos_sei, id_distrito, clave_colonia, num_mro, tipo_mro${anio != 1 ? ', anio' : ''}
            FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas_VVS
            WHERE estatus = 1
        )
        SELECT COALESCE((
            SELECT id_acta, clave_colonia, nombre_colonia, id_delegacion, nombre_delegacion, num_mro, tipo_mro, mro, coordinador_sino, num_integrantes, observador_sino, levantada_distrito, razon_distrital, bol_recibidas, bol_adicionales,
            total_ciudadanos, bol_sobrantes, bol_nulas, bol_nulas_sei, bol_total_emitidas, opi_total_sei, (
                SELECT ${anio == 1 ? 'dbo.NumeroALetras(secuencial) AS ' : ''}secuencial, nom_p,${anio != 1 ? ' rubro_general,' : ''} votos, votos_sei
                FROM V
                WHERE id_distrito = C.id_distrito AND clave_colonia = C.clave_colonia AND num_mro = C.num_mro AND tipo_mro = C.tipo_mro${anio != 1 ? ' AND anio = C.anio' : ''}
                ORDER BY V.secuencial ASC
                FOR JSON PATH
            ) AS integraciones
            FROM C
            FOR JSON PATH
        ), '[]') AS data`))[0][0].data)[0];
        if (!datos)
            return res.status(404).json({
                success: false,
                msg: 'Acta no encontrada'
            });
        res.json({
            success: true,
            datos: EncryptData(datos)
        });
    } catch (err) {
        console.error(`Error en DatosActa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const RegistrarActa = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { clave_colonia, tipo_mro, num_mro, levantada_distrito, forzar, coordinador_sino, num_integrantes, observador_sino, bol_recibidas, bol_adicionales, bol_sobrantes, total_ciudadanos, bol_nulas, bol_total_emitidas, opi_total_sei, razon_distrital, id_incidencia, integraciones, anio } = req.body;
    try {
        if (await ConsultaExistenciaActas(id_distrito, clave_colonia, num_mro, tipo_mro, anio) != 0)
            return res.status(400).json({
                success: false,
                msg: 'Esta UT ya tiene un Acta con esta MRO'
            });
        let insertVotos = '', valuesVotos = '';
        for (let X of integraciones) {
            const { secuencial: sec, votos } = X;
            if (isNaN(+votos))
                return res.status(400).json({
                    success: false,
                    msg: 'Favor de llenar todos los votos con opiniones'
                });
            insertVotos += `${anio == 1 ? `participante${LetrasANumero(sec)}` : `proyecto${sec}_votos`}, `
            valuesVotos += `${votos}, `;
        }
        // if (!levantada_distrito && !forzar) {
        //     if (bol_recibidas + bol_adicionales != bol_total_emitidas + bol_sobrantes)  //? id_incidente 2 = El total dde boletas recibidas más las adicionales no coinciden con la suma del total de opiniones emitidas más las boletas sobrantes
        //         return res.status(400).json({
        //             success: false,
        //             msg: `El número de boletas sobrantes más los resultados de la votación en mesa 'No es igual' a las boletas entregadas. (No incluye Datos SEI)`
        //         });
        //     if (total_ciudadanos != (opi_total_sei + bol_total_emitidas)) //? id_incidente 1 = La suma de los votos de mesa más los del SEI no coinciden con el total de votos de la ciudadania
        //         return res.status(400).json({
        //             success: false,
        //             msg: `La suma de personas que votaron 'No es igual' a los resultados de la votación asentados en el acta. (Incluye Datos SEI) ¿Deseas Levantar el Acta en Dirección Distrital?`
        //         });
        // }
        const { id_delegacion } = await ConsultaDelegacion(id_distrito, clave_colonia);
        const { id_acta } = (await SICOVACC.sequelize.query(`INSERT ${anio == 1 ? 'copaco' : 'consulta'}_actas (id_distrito, id_delegacion, clave_colonia, num_mro, tipo_mro, modalidad, coordinador_sino, num_integrantes, bol_recibidas, total_ciudadanos, bol_sobrantes, bol_nulas, opi_total_computada, votacion_total_emitida, ${insertVotos.substring(0, insertVotos.length - 2)}, bol_adicionales, levantada_distrito, observador_sino, razon_distrital,${anio != 1 ? ' anio,' : ''} id_incidencia, id_usuario, fecha_alta, estatus)
        OUTPUT INSERTED.id_acta
        VALUES (${id_distrito}, ${id_delegacion}, '${clave_colonia}', ${num_mro}, ${tipo_mro}, 1, ${coordinador_sino ? 1 : 0}, ${num_integrantes ? num_integrantes : 'NULL'}, ${bol_recibidas}, ${total_ciudadanos}, ${bol_sobrantes}, ${bol_nulas}, 0, ${bol_total_emitidas}, ${valuesVotos.substring(0, valuesVotos.length - 2)}, ${bol_adicionales}, ${levantada_distrito ? 1 : 0}, ${observador_sino ? 1 : 0}, ${razon_distrital ? `'${razon_distrital}'` : 'NULL'},${anio != 1 ? ` ${anio},` : ''} ${id_incidencia ? id_incidencia : 'NULL'}, ${id_usuario}, CURRENT_TIMESTAMP, 1)`))[0][0];
        await Audit(id_transaccion, id_usuario, id_distrito, `REGISTRÓ EL ACTA DE LA ${anio == 1 ? 'ELECCIÓN' : 'CONSULTA'}, DE LA UT ${clave_colonia}, MESA M${String(num_mro).padStart(2, '0')}${tipo_mro != 1 ? `, DE TIPO DE MESA ${TipoMesa(tipo_mro)}` : ''}`);
        res.json({
            success: true,
            msg: 'Acta Registrada',
            id_acta: EncryptData(id_acta)
        });
    } catch (err) {
        console.error(`Error en RegistrarActa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const ActualizarActa = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { id_acta, levantada_distrito, forzar, coordinador_sino, num_integrantes, observador_sino, bol_recibidas, bol_adicionales, bol_sobrantes, total_ciudadanos, bol_nulas, bol_total_emitidas, opi_total_sei, razon_distrital, id_incidencia, anio, integraciones } = req.body;
    try {
        let updateVotos = '';
        for (let X of integraciones) {
            const { secuencial: sec, votos } = X;
            if (isNaN(+votos))
                return res.status(400).json({
                    success: false,
                    msg: 'Favor de llenar todos los votos con opiniones'
                });
            updateVotos += `${anio == 1 ? `participante${LetrasANumero(sec)}` : `proyecto${sec}_votos`} = ${votos}, `;
        }
        // if (!levantada_distrito && !forzar) {
        //     if (bol_recibidas + bol_adicionales != bol_total_emitidas + bol_sobrantes)
        //         return res.status(400).json({
        //             success: false,
        //             msg: `El número de boletas sobrantes más los resultados de la votación en mesa 'no es igual' a las boletas entregadas. (No incluye Datos SEI)`
        //         });
        //     if (total_ciudadanos != (opi_total_sei + bol_total_emitidas))
        //         return res.status(400).json({
        //             success: false,
        //             msg: `La suma de personas que votaron 'No es igual' a los resultados de la votación asentados en el acta. (Incluye Datos SEI)`
        //         });
        // }
        // Array.from({ length: proyectos }, (_, idx) => ({ num: idx + 1 }))
        let select = '';
        Array.from({ length: anio == 1 ? 100 : 50 }, (_, idx) => ({ num: idx + 1 })).forEach(({ num }, i) => select += `${anio == 1 ? `participante${num}` : `proyecto${num}_votos`}${i != (anio == 1 ? 100 : 50) - 1 ? ', ' : ''}`)
        const acta = (await SICOVACC.sequelize.query(`SELECT id_acta,${anio != 1 ? ` anio,` : ''} id_distrito, id_delegacion, clave_colonia, num_mro, tipo_mro, modalidad, CAST(coordinador_sino AS INTEGER) AS coordinador_sino, num_integrantes, bol_recibidas, total_ciudadanos, bol_sobrantes, bol_nulas, opi_total_computada, votacion_total_emitida, CAST(levantada_distrito AS INTEGER) AS levantada_distrito, ${select}, CAST(observador_sino AS INTEGER) AS observador_sino, bol_adicionales, razon_distrital, id_incidencia, id_usuario, CONVERT(VARCHAR(19), fecha_alta, 120) AS fecha_alta, CONVERT(VARCHAR(19), fecha_modif, 120) AS fecha_modif, estatus
        FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas WHERE id_acta = ${id_acta}`))[0][0];
        if (!acta)
            return res.status(404).json({
                success: false,
                msg: 'Acta no encontrada'
            });
        const varchar = ['clave_colonia', 'num_mro', 'observaciones', 'razon_distrital', 'fecha_alta', 'fecha_modif'];
        let insert = '', values = '';
        Object.keys(acta).forEach(key => {
            insert += `${key}${!key.match('estatus') ? ', ' : ''}`;
            values += `${varchar.includes(key) && acta[key] ? `'${acta[key]}'` : acta[key]}${!key.match('estatus') ? ', ' : ''}`;
        });
        const { clave_colonia, num_mro, tipo_mro } = acta;
        // await SIVACC.sequelize.query(`INSERT consulta_actas_hist (${insert}) VALUES (${values})`);
        await SICOVACC.sequelize.query(`UPDATE ${anio == 1 ? 'copaco' : 'consulta'}_actas SET coordinador_sino = ${coordinador_sino ? 1 : 0}, num_integrantes = ${num_integrantes ? num_integrantes : 'NULL'}, bol_recibidas = ${bol_recibidas}, total_ciudadanos = ${total_ciudadanos}, bol_sobrantes = ${bol_sobrantes}, bol_nulas = ${bol_nulas}, votacion_total_emitida = ${bol_total_emitidas}, ${updateVotos.substring(0, updateVotos.length - 2)},
        bol_adicionales = ${bol_adicionales}, levantada_distrito = ${levantada_distrito ? 1 : 0}, razon_distrital = ${razon_distrital ? `'${razon_distrital}'` : 'NULL'}, id_incidencia = ${id_incidencia ? id_incidencia : 'NULL'}, observador_sino = ${observador_sino ? 1 : 0}, fecha_modif = CURRENT_TIMESTAMP
        WHERE modalidad = 1 AND estatus = 1 AND id_acta = ${id_acta}`);
        await Audit(id_transaccion, id_usuario, id_distrito, `ACTUALIZÓ EL ACTA DE LA ${anio == 1 ? 'ELECCIÓN' : 'CONSULTA'}, DE LA UT ${clave_colonia}, MESA M${String(num_mro).padStart(2, '0')}${tipo_mro != 1 ? `, DE TIPO DE MESA ${TipoMesa(tipo_mro)}` : ''}`);
        res.json({
            success: true,
            msg: 'Acta Actualizada'
        });
    } catch (err) {
        console.error(`Error en ActualizarActa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

//? Cierre de Validación

export const DatosCierreValidacion = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT cierre_asistencia1 AS MSPEN, cierre_asistencia2 AS COPACO, cierre_asistencia3 AS personasCandidatas, cierre_asistencia4 AS personasObservadoras, cierre_asistencia5 AS presentaronProyecto,
        cierre_asistencia6 AS mediosComunicacion, cierre_asistencia7 AS otros, cierre_total AS total,
        CONVERT(VARCHAR(10), fecha_hora_cierre, 20) AS fecha, CONVERT(VARCHAR(5), fecha_hora_cierre, 8) AS hora, cierre_observaciones AS observaciones FROM consulta_computo WHERE estatus = 1 AND fecha_hora_cierre IS NOT NULL AND id_distrito = ${id_distrito}`))[0][0];
        if (!datos)
            return res.status(404).json({
                success: false,
                msg: 'No hay información'
            });
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en DatosCierreValidacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const GuardarActualizarCierreValidacion = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { MSPEN, COPACO, personasCandidatas, personasObservadoras, presentaronProyecto, mediosComunicacion, otros, total, fecha, hora, observaciones } = req.body;
    try {
        const { cierreValidacion } = await ConsultaVerificaInicioCierre(id_distrito);
        await SICOVACC.sequelize.query(`UPDATE consulta_computo set cierre_asistencia1 = ${MSPEN}, cierre_asistencia2 = ${COPACO}, cierre_asistencia3 = ${personasCandidatas ? personasCandidatas : 'NULL'},
        cierre_asistencia4 = ${personasObservadoras}, cierre_asistencia5 = ${presentaronProyecto}, cierre_asistencia6 = ${mediosComunicacion},
        cierre_asistencia7 = ${otros}, cierre_total = ${total}, fecha_hora_cierre = '${fecha} ${hora}:00', cierre_observaciones = UPPER('${Comillas(observaciones)}'), fecha_modif = CURRENT_TIMESTAMP
        WHERE estatus = 1 AND id_distrito = ${id_distrito}`);
        await Audit(id_transaccion, id_usuario, id_distrito, `${!cierreValidacion ? 'REALIZÓ' : 'ACTUALIZÓ'} EL CIERRE DE VALIDACIÓN`);
        res.json({
            success: true,
            msg: 'Información guardada',
            cierreValidacion: !cierreValidacion
        });
    } catch (err) {
        console.error(`Error en GuardarActualizarCierreValidacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}