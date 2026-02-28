import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Registra el inicio de sesión junto con el navegador y sistema operativo en el que se esta usando
export const IniciarSesion = async (id_usuario, id_distrito, ip_cliente, useragent, descripcion) => {
    try {
        const so_cliente = useragent.split('/')[1].trim(), navegador_cliente = useragent.split('/')[0].trim();
        await SICOVACC.sequelize.query(`INSERT consulta_audit (id_usuario, id_distrito, ip_cliente, so_cliente, navegador_cliente, fecha_inicio, estatus) VALUES (${id_usuario}, ${id_distrito}, '${ip_cliente}', '${so_cliente}', '${navegador_cliente}', CURRENT_TIMESTAMP, 1)`);
        const consulta = await SICOVACC.sequelize.query(`SELECT TOP 1 ID, CONVERT(VARCHAR(25), fecha_inicio, 121) AS fecha FROM consulta_audit WHERE estatus = 1 AND id_usuario = ${id_usuario} ORDER BY fecha_inicio DESC`);
        const { ID, fecha } = consulta[0][0];
        await Audit(ID, id_usuario, id_distrito, descripcion, fecha);
        return ID;
    } catch (err) {
        console.error(`Error al procesar al usuario ${id_usuario}`);
    }
}

//? Actualiza el token
export const ActualizarInicio = async (id_transaccion, token) => await SICOVACC.sequelize.query(`UPDATE consulta_audit SET token = '${token}' WHERE ID = ${id_transaccion}`);

//? Elimina el token, actualiza la fecha de cierre y el estatus, usado para cuando pierde conexión o cierra sesión
export const DesactivarInicio = async id_transaccion => await SICOVACC.sequelize.query(`UPDATE consulta_audit SET token = NULL, fecha_cierre = CURRENT_TIMESTAMP, estatus = 0 WHERE estatus = 1 AND ID = ${id_transaccion}`);

//? En dado caso que el usuario tenga información activa, actualiza esos registros
export const DesactivarUsuario = async id_usuario => await SICOVACC.sequelize.query(`UPDATE consulta_audit SET token = NULL, fecha_cierre = CURRENT_TIMESTAMP, estatus = 0 WHERE estatus = 1 AND id_usuario = ${id_usuario}`);

//? Registra el cierre de sesión o perdida de conexion
export const CerrarSesion = async (id_transaccion, descripcion) => {
    try {
        await DesactivarInicio(id_transaccion);
        const consulta = await SICOVACC.sequelize.query(`SELECT id_usuario, id_distrito, CONVERT(VARCHAR(25), fecha_cierre, 121) AS fecha FROM consulta_audit WHERE ID = ${id_transaccion}`);
        const { id_usuario, id_distrito, fecha } = consulta[0][0];
        await Audit(id_transaccion, id_usuario, id_distrito, descripcion, fecha);
    } catch (err) {
        console.error(`Error con la transaccion ${id_transaccion}`);
    }
}

//? Registra todas las acciones que el usuario haga
export const Audit = async (id_transaccion_det, id_usuario, id_distrito, descripcion, fecha_registro = 'CURRENT_TIMESTAMP') => await SICOVACC.sequelize.query(`INSERT consulta_audit_det (id_transaccion_det, id_usuario, id_distrito, descripcion, fecha_registro) VALUES (${id_transaccion_det}, ${id_usuario}, ${id_distrito}, UPPER('${descripcion}'), ${!fecha_registro.match('CURRENT_TIMESTAMP') ? `'${fecha_registro}'` : fecha_registro})`);