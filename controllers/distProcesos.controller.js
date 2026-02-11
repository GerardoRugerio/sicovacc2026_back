import { request, response } from 'express';
import { Audit } from '../helpers/Audit.js';
import { Comillas } from '../helpers/Funciones.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Actualización de Datos del Disrtito

export const DatosDistrito = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT UPPER(CCDe.nombre_delegacion) AS nombre_delegacion, UPPER(CCDi.domicilio) AS domicilio, CCDi.cp AS codigo_postal, COALESCE(UPPER(CCDi.coordinador), '') AS coordinador, COALESCE(UPPER(CCDi.coordinador_puesto), '') AS coordinador_puesto, COALESCE(CCDi.coordinador_genero, '') AS coordinador_genero,
        COALESCE(UPPER(CCDi.secretario), '') AS secretario, COALESCE(UPPER(CCDi.secretario_puesto), '') AS secretario_puesto, COALESCE(CCDi.secretario_genero, '') AS secretario_genero
        FROM consulta_cat_distrito CCDi
        LEFT JOIN consulta_cat_delegacion CCDe ON CCDi.id_delegacion = CCDe.id_delegacion
        WHERE CCDi.id_distrito = ${id_distrito}`))[0][0];
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en DatosDistrito: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const ActualizarDatosDistrito = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { domicilio, codigo_postal, coordinador, coordinador_puesto, coordinador_genero, secretario, secretario_puesto, secretario_genero } = req.body;
    try {
        await SICOVACC.sequelize.query(`UPDATE consulta_cat_distrito SET domicilio = UPPER('${Comillas(domicilio)}'), coordinador = ${coordinador ? `UPPER('${Comillas(coordinador)}')` : 'NULL'}, coordinador_puesto = ${coordinador_puesto ? `UPPER('${Comillas(coordinador_puesto)}')` : 'NULL'}, coordinador_genero = ${coordinador_genero ? `'${coordinador_genero}'` : 'NULL'},
        secretario = ${secretario ? `UPPER('${Comillas(secretario)}')` : 'NULL'}, secretario_puesto = ${secretario_puesto ? `UPPER('${Comillas(secretario_puesto)}')` : 'NULL'}, secretario_genero = ${secretario_genero ? `'${secretario_genero}'` : 'NULL'}, fecha_modif = CURRENT_TIMESTAMP WHERE id_distrito = ${id_distrito}`);
        await Audit(id_transaccion, id_usuario, id_distrito, 'ACTUALIZÓ LA INFORMACIÓN DEL DISTRITO');
        res.json({
            success: true,
            msg: 'Actualización de Datos Realizado Correctamente '
        });
    } catch (err) {
        console.error(`Error en ActualizarDatosDistrito: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error, Registro no Guardado'
        });
    }
}

//? Limpiar la Base de Datos

export const LimpiarBD = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    try {
        const consulta = await SICOVACC.sequelize.query(`SELECT CASE WHEN perfil = 1 THEN COALESCE(ocultar_opcion, 0) ELSE 0 END AS opcion FROM consulta_usuarios_sivacc WHERE id_usuario = ${id_usuario}`);
        const { opcion } = consulta[0][0];
        if (opcion == 0)
            return res.json({
                success: true,
                msg: undefined
            });
        await SICOVACC.sequelize.query(`EXEC LimpiarBD ${id_distrito}`);
        // await SIVACC.sequelize.query(`UPDATE consulta_usuarios_sivacc SET ocultar_opcion = 0 WHERE id_usuario = ${id_usuario}`);
        await Audit(id_transaccion, id_usuario, id_distrito, 'LIMPIÓ SU BD');
        res.json({
            success: true,
            msg: '¡Se realizó la limpieza de la base de datos!'
        });
    } catch (err) {
        console.error(`Error en LimpiarBD: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error al Limpiar la BD'
        });
    }
}