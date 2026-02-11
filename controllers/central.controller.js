import { request, response } from 'express';
import { Audit } from '../helpers/Audit.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Captura de Resultados de Consulta por Mesa

export const Actas = async (req = request, res = response) => {
    const { id_distrito, clave_colonia, anio } = req.body;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT CA.id_acta, UPPER(CCC.nombre_colonia) AS nombre_colonia, CCC.clave_colonia, CONCAT('M', RIGHT('00' + CA.num_mro, 2)) AS num_mro, CA.tipo_mro
        FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas CA
        LEFT JOIN consulta_cat_colonia_cc1 CCC ON CA.clave_colonia = CCC.clave_colonia
        WHERE CA.modalidad = 1 AND CA.estatus = 1${anio != 1 ? ` AND anio = ${anio}` : ''} AND CA.id_distrito = ${id_distrito} AND CA.clave_colonia = '${clave_colonia}'
        ORDER BY CA.num_mro, CA.tipo_mro`))[0];
        if (!datos)
            return res.status(404).json({
                success: false,
                msg: 'Esta UT no tiene actas capturadas'
            });
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en Actas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const EliminarActa = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { id_acta, anio } = req.body;
    try {
        let select = '';
        Array.from({ length: anio == 1 ? 100 : 50 }, (_, idx) => ({ num: idx + 1 })).forEach(({ num }, i) => select += `${anio == 1 ? `participante${num}` : `proyecto${num}_votos`}${i != (anio == 1 ? 100 : 50) - 1 ? ', ' : ''}`);
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
        if (anio != 1)
            SICOVACC.sequelize.query(`INSERT consulta_Actas_hist (${insert}) VALUES (${values})`);
        await SICOVACC.sequelize.query(`DELETE FROM ${anio == 1 ? 'copaco' : 'consulta'}_actas WHERE id_acta = ${id_acta}`);
        await Audit(id_transaccion, id_usuario, id_distrito, `ELIMINÓ EL ACTA DE LA ${anio == 1 ? 'ELECCIÓN' : 'CONSULTA'}, DE LA UT ${clave_colonia}, MESA M${String(num_mro).padStart(2, '0')}${tipo_mro != 1 ? `, DE TIPO DE MESA ${TipoMesa(tipo_mro)}` : ''}`);
        res.json({
            success: true,
            msg: 'Acta eliminada'
        });
    } catch (err) {
        console.error(`Error en EliminarActa: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}