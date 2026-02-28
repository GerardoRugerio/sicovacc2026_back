import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Regresa la fecha del servidor
export const FechaServer = async () => {
    const fecha_server = await SICOVACC.sequelize.query(`SELECT CONVERT(VARCHAR(10), GETDATE(), 103) AS fecha, CONVERT(VARCHAR(8), GETDATE(), 114) AS hora`);
    return fecha_server[0][0]
}

//? Regresa el número de proyecto mas alto para saber cuantos proyectos tiene el distrito y/o la Unidad Territorial (Se necesita tener captradas todas las mesas de la UT)
export const ProyMax = async (anio, distrito, clave_colonia = undefined) => {
    const total = await SICOVACC.sequelize.query(`SELECT COALESCE(MAX(CPP.num_proyecto), 0) AS num_proyectos
    FROM consulta_actas CA
    LEFT JOIN consulta_prelacion_proyectos CPP ON CA.id_distrito = CPP.id_distrito AND CA.clave_colonia = CPP.clave_colonia AND CA.anio = CPP.anio AND CA.estatus = CPP.estatus
    WHERE CA.modalidad = 1 AND CA.estatus = 1 AND CA.anio = ${anio} ${distrito == 0 ? '' : `AND CA.id_distrito = ${distrito}`} ${clave_colonia ? `AND CA.clave_colonia = '${clave_colonia}'` : ''}AND CA.clave_colonia IN (
        SELECT A.clave_colonia
        FROM (SELECT clave_colonia, COUNT(clave_colonia) AS cantidad FROM consulta_actas WHERE modalidad = 1 AND estatus = 1 AND anio = ${anio} GROUP BY clave_colonia) AS A
        LEFT JOIN (SELECT clave_colonia, COUNT(clave_colonia) AS total FROM consulta_mros WHERE estatus = 1 GROUP BY clave_colonia) AS B ON A.clave_colonia = B.clave_colonia
        WHERE A.cantidad = B.total)`);
    return total[1] != 0 ? total[0][0].num_proyectos : 0;
}

//? Regresa el número de proyecto mas alto para saber cuantos proyectos tiene el distrito y/o la Unidad Territorial
export const ProyMaxG = async (anio, distrito, clave_colonia = undefined) => {
    const total = await SICOVACC.sequelize.query(`SELECT COALESCE(MAX(CPP.num_proyecto), 0) AS num_proyectos
    FROM consulta_actas CA
    LEFT JOIN consulta_prelacion_proyectos CPP ON CA.id_distrito = CPP.id_distrito AND CA.clave_colonia = CPP.clave_colonia AND CA.anio = CPP.anio AND CA.estatus = CPP.estatus
    WHERE CA.modalidad = 1 AND CA.estatus = 1 AND CA.anio = ${anio} ${distrito == 0 ? '' : `AND CA.id_distrito = ${distrito}`} ${clave_colonia ? `AND CA.clave_Colonia = '${clave_colonia}'` : ''}`);
    return total[1] != 0 ? total[0][0].num_proyectos : 0;
}

//? Regresa una lista de proyectos de la Unidad Territorial (anio 1 y 2) o de los participantes de la COPACO (anio 1)
export const Listado = async (distrito, clave_colonia, anio) => {
    const proyectos = await SICOVACC.sequelize.query(`SELECT ${anio == 1 ? 'secuencial' : 'num_proyecto AS secuencial'}, ${anio == 1 ? `UPPER(CONCAT(nombre, ' ', paterno, ' ', materno))` : 'nom_proyecto'} AS nom_p
    FROM ${anio == 1 ? 'copaco_formulas' : 'consulta_prelacion_proyectos'}
    WHERE ${anio == 1 ? 'secuencial IS NOT NULL' : 'estatus = 1'} AND id_distrito = ${distrito} AND clave_colonia = '${clave_colonia}'${anio != 1 ? ` AND anio = ${anio}` : ''}
    ORDER BY ${anio == 1 ? 'secuencial' : 'num_proyecto'} ASC`);
    return proyectos[0];
}

//? Regresa una lista de proyectos de la Unidad Territorial (Se necesita tener capturadas todas las mesas de la UT)
export const ProyectosUT = async (anio, distrito, clave_colonia) => {
    const proyectos = await SICOVACC.sequelize.query(`SELECT DISTINCT CPP.num_proyecto, CPP.nom_proyecto
    FROM consulta_actas CA
    LEFT JOIN consulta_prelacion_proyectos CPP ON CA.id_distrito = CPP.id_distrito AND CA.clave_colonia = CPP.clave_colonia AND CA.anio = CPP.anio AND CA.estatus = CPP.estatus
    WHERE CA.modalidad = 1 AND CA.estatus = 1 AND CA.anio = ${anio} AND CA.id_distrito = ${distrito} AND CA.clave_colonia = '${clave_colonia}' AND CA.clave_colonia IN (
        SELECT A.clave_colonia
        FROM (SELECT clave_colonia, COUNT(clave_colonia) AS cantidad FROM consulta_actas WHERE modalidad = 1 AND estatus = 1 AND anio = ${anio} GROUP BY clave_colonia) AS A
        LEFT JOIN (SELECT clave_colonia, COUNT(clave_colonia) AS total FROM consulta_mros WHERE estatus = 1 GROUP BY clave_colonia) AS B ON A.clave_colonia = B.clave_colonia
        WHERE A.cantidad = B.total
    )`);
    return proyectos[0];
}

export const ConsultaTipoEleccion = async anio => (await SICOVACC.sequelize.query(`SELECT descripcion FROM consulta_cat_tipo_eleccion WHERE id_tipo_eleccion = ${anio}`))[0][0].descripcion;

//? Regresa el inicio y cierre de la validación del distrito
export const ConsultaVerificaInicioCierre = async distrito => {
    const consulta = await SICOVACC.sequelize.query(`SELECT CAST(CASE WHEN SUM(CASE WHEN inicio_asistencia1 IS NOT NULL THEN 1 ELSE 0 END) > 0 THEN 1 ELSE 0 END AS BIT) AS inicioValidacion,
    CAST(CASE WHEN SUM(CASE WHEN cierre_asistencia1 IS NOT NULL THEN 1 ELSE 0 END) > 0 THEN 1 ELSE 0 END AS BIT) AS cierreValidacion
    FROM consulta_computo WHERE estatus = 1 AND id_distrito = ${distrito}`);
    return consulta[0][0];
}

//? Regresa el id de la Delegación de la Unidad Territorial
export const ConsultaClaveColonia = async (clave_colonia) => {
    const consulta = await SICOVACC.sequelize.query(`SELECT id_colonia, nombre_colonia, id_delegacion FROM consulta_cat_colonia_cc1 WHERE clave_colonia = '${clave_colonia}'`);
    return consulta[0][0];
}

//? Regresa el nombre de la Delegación de la Unidad Territorial
export const ConsultaDelegacion = async (distrito, clave_colonia) => {
    const consulta = await SICOVACC.sequelize.query(`SELECT DISTINCT D.id_delegacion, UPPER(D.nombre_delegacion) AS nombre_delegacion
    FROM consulta_mros M
    LEFT JOIN consulta_cat_delegacion D ON M.id_delegacion = D.id_delegacion
    WHERE M.estatus = 1 AND M.id_distrito = ${distrito} AND M.clave_colonia = '${clave_colonia}'`);
    return consulta[0][0];
}

//? Regresa la información del distrito
export const ConsultaDistrito = async distrito => {
    const consulta = await SICOVACC.sequelize.query(`SELECT UPPER(domicilio) AS direccion, COALESCE(UPPER(coordinador), '') AS coordinador, COALESCE(UPPER(coordinador_puesto), '') AS coordinador_puesto, COALESCE(UPPER(secretario), '') AS secretario, COALESCE(UPPER(secretario_puesto), '') AS secretario_puesto FROM consulta_cat_distrito WHERE id_distrito = ${distrito}`);
    return consulta[0][0];
}

//? Regresa informacion de la existencia de la acta
export const ConsultaExistenciaActas = async (id_distrito, clave_colonia, num_mro, tipo_mro, anio) => {
    const consulta = await SICOVACC.sequelize.query(`SELECT * FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas WHERE modalidad = 1 AND estatus = 1 AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}' AND num_mro = ${num_mro} AND tipo_mro = ${tipo_mro}${anio != 1 ? ` AND anio = ${anio}` : ''}`);
    return consulta[1];
}

//? Regresa informacion de la verificación de los proyecytos o si tienen sorteo hecho
export const ConsultaVerificaProyectos = async (id_distrito, clave_colonia, anio, sorteo) => {
    const consulta = await SICOVACC.sequelize.query(`SELECT * FROM ${anio == 1 ? 'copaco_formulas' : 'consulta_prelacion_proyectos'} WHERE ${anio == 1 ? 'secuencial IS NOT NULL' : 'estatus = 1'} AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'${anio != 1 ? ` AND anio = ${anio}${sorteo ? ' AND id_sorteo > 0' : ''}` : ''}`);
    return consulta[1];
}

//? Regresa el total de mesas instaladas en la Unidad Territorial
export const MesasI = async (distrito, clave_colonia, anio) => {
    const consulta = await SICOVACC.sequelize.query(`SELECT CAST(CASE WHEN COUNT(num_mro) > 0 THEN 1 ELSE 0 END AS BIT) AS mesasI FROM consulta_mros WHERE estatus = 1 AND id_distrito = ${distrito} AND clave_colonia = '${clave_colonia}'`);
    return consulta[0][0];
}

//? Regresa el estado de la Unidad Territorial, si ya fueron capturadas todas las actas
export const EstadoUT = async (clave_colonia, anio) => {
    const consulta = await SICOVACC.sequelize.query(`SELECT CAST(CASE WHEN A.cantidad = B.cantidad THEN 1 ELSE 0 END AS BIT) AS validada
    FROM (
        SELECT id_distrito, id_delegacion, clave_colonia, COUNT(num_mro) AS cantidad
        FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas
        WHERE modalidad = 1 AND estatus = 1${anio != 1 ? ` AND anio = ${anio}` : ''}
        GROUP BY id_distrito, id_delegacion, clave_colonia
    ) AS A
    LEFT JOIN (
        SELECT id_distrito, id_delegacion, clave_colonia, COUNT(num_mro) AS cantidad
        FROM consulta_mros
        WHERE estatus = 1
        GROUP BY id_distrito, id_delegacion, clave_colonia
    ) AS B ON A.id_distrito = B.id_distrito AND A.id_delegacion = B.id_delegacion AND A.clave_colonia = B.clave_colonia
    WHERE A.clave_colonia = '${clave_colonia}'`);
    return consulta[1] != 0 ? consulta[0][0].validada : false;
}

//? Regresa información de las mesas faltantes de la Unidad Territorial
export const MesasFalt = async (id_distrito, clave_colonia, anio) => {
    const consulta = await SICOVACC.sequelize.query(`SELECT * FROM consulta_mros M WHERE estatus = 1 AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
    AND NOT EXISTS (SELECT 1 FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas WHERE modalidad = 1 AND estatus = 1${anio != 1 ? ` AND anio = ${anio}` : ''} AND id_distrito = M.id_distrito AND clave_colonia = M.clave_colonia AND num_mro = M.num_mro AND tipo_mro = M.tipo_mro)`);
    return consulta[1];
}

//? Regresa información para las constancias
export const InformacionConstancia = async (anio, clave_colonia) => {
    const consulta = await SICOVACC.sequelize.query(`SELECT TOP 1 CA.nombre_colonia, CA.nombre_delegacion, CA.domicilio, COALESCE(CM.mesas, 0) AS mesas, CA.ultimaFecha, CA.ultimaHora, CA.coordinador_puesto, CA.coordinador, CA.secretario_puesto, CA.secretario
    FROM (
        SELECT A.nombre_colonia, UPPER(C.nombre_delegacion) AS nombre_delegacion, CASE WHEN B.domicilio IS NULL THEN 'SIN INFORMACIÓN' ELSE CONCAT(UPPER(B.domicilio), ', C.P. ', B.cp) END AS domicilio, CONVERT(VARCHAR(10), A.fecha, 103) AS ultimaFecha, CONVERT(VARCHAR(5), A.fecha, 114) AS ultimaHora,
        CASE WHEN B.coordinador_puesto IS NULL THEN 'SIN INFORMACIÓN' ELSE UPPER(B.coordinador_puesto) END AS coordinador_puesto, CASE WHEN B.coordinador IS NULL THEN 'SIN INFORMACIÓN' ELSE UPPER(B.coordinador) END AS coordinador,
        CASE WHEN B.secretario_puesto IS NULL THEN 'SIN INFORMACIÓN' ELSE UPPER(B.secretario_puesto) END AS secretario_puesto, CASE WHEN B.secretario IS NULL THEN 'SIN INFORMACIÓN' ELSE UPPER(B.secretario) END AS secretario,
        A.anio, A.clave_colonia
        FROM consulta_actas_VVS A
        LEFT JOIN consulta_cat_distrito B ON A.id_distrito = B.id_distrito
        LEFT JOIN consulta_cat_delegacion C ON A.id_delegacion = C.id_delegacion
    ) AS CA
    LEFT JOIN (SELECT clave_colonia, COUNT(clave_colonia) AS mesas FROM consulta_mros WHERE estatus = 1 GROUP BY clave_colonia) AS CM ON CA.clave_colonia = CM.clave_colonia
    WHERE CA.anio = ${anio} AND CA.clave_colonia = '${clave_colonia}'
    ORDER BY CA.ultimaFecha, CA.ultimaHora DESC`);
    return consulta[0][0];
}

export const FechaHoraActa = async (id_distrito, clave_colonia, anio) => (await SICOVACC.sequelize.query(`SELECT CONVERT(VARCHAR(10), fecha_alta, 103) AS fechaActa, CONVERT(VARCHAR(5), fecha_alta, 114) AS horaActa
    FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas
    WHERE modalidad = 1 AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'${anio != 1 ? ` AND anio = ${anio}` : ''}`))[0][0];