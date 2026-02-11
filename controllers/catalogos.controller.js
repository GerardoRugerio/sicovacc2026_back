import { request, response } from 'express';
import { aniosCAT } from '../helpers/Constantes.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

export const TipoEleccion = async (req = request, res = response) => {
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT id_tipo_eleccion AS id, descripcion AS nombre FROM consulta_cat_tipo_eleccion WHERE estatus = 1`))[0];
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en TipoEleccion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const Colonias = async (req = request, res = response) => {
    const { id_distrito } = req.data.id_distrito == 0 ? req.body : req.data;
    const { anio } = req.body;
    const campo = aniosCAT[0][anio];
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT clave_colonia AS id, UPPER(nombre_colonia) AS nombre FROM consulta_cat_colonia_cc1 WHERE ${campo} = 1 AND id_distrito = ${id_distrito} ORDER BY nombre_colonia`))[0];
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en Colonias: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const ColoniasConActas = async (req = request, res = response) => {
    const { id_distrito } = req.data.id_distrito == 0 ? req.body : req.data;
    const { anio } = req.body;
    const campo = aniosCAT[0][anio];
    try {
        const colonias = await SICOVACC.sequelize.query(`SELECT clave_colonia AS id, UPPER(nombre_colonia) AS nombre
        FROM consulta_cat_colonia_cc1
        WHERE ${campo} = 1 AND id_distrito = ${id_distrito} AND clave_colonia IN (SELECT clave_colonia FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas WHERE modalidad = 1 AND estatus = 1${anio != 1 ? ` AND anio = ${anio}` : ''}) AND 
        clave_colonia IN (SELECT clave_colonia FROM consulta_mros WHERE ${campo} = 1)
        ORDER BY nombre_colonia`);
        if (colonias[1] == 0)
            return res.status(404).json({
                success: false,
                msg: 'No hay UT con cómputo capturado'
            });
        res.json({
            success: true,
            datos: colonias[0]
        });
    } catch (err) {
        console.error(`Error en ColoniasConActas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const ColoniasSinActas = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { anio } = req.body;
    const campo = aniosCAT[0][anio];
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT clave_colonia AS id, UPPER(nombre_colonia) AS nombre
        FROM consulta_cat_colonia_cc1
        WHERE ${campo} = 1 AND id_distrito = ${id_distrito} AND clave_colonia NOT IN (
            SELECT A.clave_colonia
            FROM (SELECT clave_colonia, COUNT(clave_colonia) AS cantidad FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas WHERE modalidad = 1 AND estatus = 1${anio != 1 ? ` AND anio = ${anio}` : ''} GROUP BY clave_colonia) AS A
            LEFT JOIN (SELECT clave_colonia, COUNT(clave_colonia) AS total FROM consulta_mros WHERE ${campo} = 1 GROUP BY clave_colonia) AS B ON A.clave_colonia = B.clave_colonia
            WHERE A.cantidad = B.total
        ) AND clave_colonia IN (SELECT clave_colonia FROM consulta_mros WHERE ${campo} = 1)
        ORDER BY nombre_colonia`))[0];
        if (!datos)
            return res.status(404).json({
                success: false,
                msg: 'No hay UT sin cómputo capturado'
            });
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en ColoniasSinActas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const ColoniasValidadas = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { anio } = req.body;
    const campo = aniosCAT[0][anio];
    try {
        const colonias = await SICOVACC.sequelize.query(`SELECT clave_colonia AS id, UPPER(nombre_colonia) AS nombre
        FROM consulta_cat_colonia_cc1
        WHERE ${campo} = 1 AND id_distrito = ${id_distrito} AND clave_colonia IN (
            SELECT A.clave_colonia
            FROM (SELECT clave_colonia, COUNT(clave_colonia) AS cantidad FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas WHERE modalidad = 1 AND estatus = 1${anio != 1 ? ` AND anio = ${anio}` : ''} GROUP BY clave_colonia) AS A
            LEFT JOIN (SELECT clave_colonia, COUNT(clave_colonia) AS total FROM consulta_mros WHERE ${campo} = 1 GROUP BY clave_colonia) AS B ON A.clave_colonia = B.clave_colonia
            WHERE A.cantidad = B.total
        ) AND clave_colonia IN (SELECT clave_colonia FROM consulta_mros WHERE ${campo} = 1)
        ORDER BY nombre_colonia`);
        if (colonias[1] == 0)
            return res.status(404).json({
                success: false,
                msg: 'Sin UT Validadas'
            });
        res.json({
            success: true,
            datos: colonias[0]
        });
    } catch (err) {
        console.error(`Error en ColoniasValidadas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const Delegacion = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { clave_colonia } = req.body;
    try {
        const { delegacion } = (await SICOVACC.sequelize.query(`SELECT DISTINCT UPPER(CCD.nombre_delegacion) AS delegacion
        FROM consulta_mros CM
        LEFT JOIN consulta_cat_delegacion CCD ON CM.id_delegacion = CCD.id_delegacion
        WHERE CM.id_distrito = ${id_distrito} AND CM.clave_colonia = '${clave_colonia}'`))[0][0];
        res.json({
            success: true,
            delegacion
        });
    } catch (err) {
        console.error(`Error en Delegacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const Mesas = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { clave_colonia, anio } = req.body;
    const campo = aniosCAT[0][anio];
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT num_mro AS id, tipo_mro AS tipo, CONCAT('M', RIGHT('00' + num_mro, 2)) AS nombre FROM consulta_mros WHERE ${campo} = 1 AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}' ORDER BY num_mro, tipo_mro`))[0];
        if (!datos)
            return res.status(404).json({
                success: false,
                msg: 'No hay mesas disponibles'
            });
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en Mesas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const MesasConActas = async (req = request, res = response) => {
    const { id_distrito, clave_colonia, anio } = req.body;
    const campo = aniosCAT[0][anio];
    try {
        const mesas = await SICOVACC.sequelize.query(`SELECT num_mro AS id, tipo_mro AS tipo, CONCAT('M', RIGHT('00' + num_mro, 2)) AS nombre FROM consulta_mros M
        WHERE ${campo} = 1 AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}' AND EXISTS (
            SELECT 1 FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas A
            WHERE A.modalidad = 1 AND A.estatus = 1 AND A.id_distrito = M.id_distrito AND A.clave_colonia = M.clave_colonia
            AND A.num_mro = M.num_mro AND A.tipo_mro = M.tipo_mro${anio != 1 ? ` AND A.anio = ${anio}` : ''}
        )`);
        if (mesas[1] == 0)
            return res.status(404).json({
                success: false,
                msg: 'No hay mesas disponibles'
            });
        res.json({
            success: true,
            datos: mesas[0]
        });
    } catch (err) {
        console.error(`Error en MesasConActas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const MesaSinActas = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { clave_colonia, anio } = req.body;
    const campo = aniosCAT[0][anio];
    try {
        const mesas = await SICOVACC.sequelize.query(`SELECT num_mro AS id, tipo_mro AS tipo, CONCAT('M', RIGHT('00' + num_mro, 2)) AS nombre FROM consulta_mros M
        WHERE ${campo} = 1 AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}' AND NOT EXISTS (
            SELECT 1 FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas A
            WHERE A.modalidad = 1 AND A.estatus = 1 AND A.id_distrito = M.id_distrito AND A.clave_colonia = M.clave_colonia
            AND A.num_mro = M.num_mro AND A.tipo_mro = M.tipo_mro${anio != 1 ? ` AND A.anio = ${anio}` : ''}
        )`);
        if (mesas[1] == 0)
            return res.status(404).json({
                success: false,
                msg: 'No hay mesas disponibles'
            });
        res.json({
            success: true,
            datos: mesas[0]
        });
    } catch (err) {
        console.error(`Error en MesaSinActas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const ColoniasVotacion = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { anio } = req.body;
    const campo = aniosCAT[0][anio];
    try {
        const colonias = await SICOVACC.sequelize.query(`SELECT DISTINCT CA.clave_colonia, UPPER(CCC.nombre_colonia) AS nombre_colonia, CA.id_delegacion, UPPER(CCD.nombre_delegacion) AS nombre_delegacion
        FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas CA
        LEFT JOIN consulta_cat_colonia_cc1 CCC ON CA.clave_colonia = CCC.clave_colonia AND CCC.${campo} = 1
        LEFT JOIN consulta_cat_delegacion CCD ON CA.id_delegacion = CCD.id_delegacion
        WHERE CA.estatus = 1 AND CA.id_distrito = ${id_distrito}${anio != 1 ? ` AND CA.anio = ${anio}` : ''}
        ORDER BY nombre_colonia`);
        res.json({
            success: true,
            datos: colonias[0]
        });
    } catch (err) {
        console.error(`Error en ColoniasVotacion: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}