import { request, response } from 'express';
import useragent from 'useragent';
import { ActualizarInicio, Audit, DesactivarUsuario, IniciarSesion } from '../helpers/Audit.js';
import { ConsultaVerificaInicioCierre } from '../helpers/Consultas.js';
import { Capitalizar } from '../helpers/Funciones.js';
import { genToken } from '../helpers/genToken.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';
// import { BuscarUsuarioID } from '../sockets/usuarios.js';
// import { BorrarUsuario, getSocketServerInstance } from '../sockets/sockets.js';

export const Login = async (req = request, res = response) => {
    const { usuario, contrasena } = req.body;
    try {
        const userData = await SICOVACC.sequelize.query(`SELECT CUS.id_usuario, CUS.usuario, CUS.contrasena, CUS.id_distrito, CCDe.nombre_delegacion, CONCAT(CUS.nombre, ' ', CUS.ape_paterno, ' ', CUS.ape_materno) AS nombre, CUS.id_admin, CUS.perfil, CUS.estatus, CUS.ocultar_opcion
        FROM consulta_usuarios_sivacc CUS
        LEFT JOIN (SELECT DISTINCT id_distrito, id_delegacion FROM consulta_cat_distrito) AS CCDi ON CUS.id_distrito = CCDi.id_distrito
        LEFT JOIN consulta_cat_delegacion CCDe ON CCDi.id_delegacion = CCDe.id_delegacion
        WHERE CUS.usuario = '${usuario}'`);
        if (userData[1] == 0)
            return res.status(401).json({
                success: false,
                msg: 'Usuario incorrecto'
            });
        if (userData[0][0].contrasena != contrasena)
            return res.status(401).json({
                success: false,
                msg: 'Contraseña incorrecta'
            });
        if (!userData[0][0].estatus)
            return res.status(403).json({
                success: false,
                msg: 'El usuario no esta activo'
            });
        const datos = {
            id_transaccion: null,
            id_usuario: userData[0][0].id_usuario,
            nombre: Capitalizar(userData[0][0].nombre.trim()),
            usuario: userData[0][0].usuario,
            id_distrito: userData[0][0].id_distrito,
            nombre_delegacion: userData[0][0].nombre_delegacion,
            perfil: userData[0][0].perfil,
            id_admin: userData[0][0].id_admin
        };
        //? Busca si el usuario ya se encuentra conectado desde otro dispositivo o navegador, si es asi, cierra sesión forzosamente
        const userConnect = await SICOVACC.sequelize.query(`SELECT ID FROM consulta_audit WHERE estatus = 1 AND id_usuario = ${datos.id_usuario} ORDER BY ID DESC`);
        if (userConnect[1] >= 1) {
            const { ID } = userConnect[0][0];
            try {
                const isProd = process.env.NODE_ENV?.match('prod') ?? false;
                const sufix = isProd ? '-redis' : '';
                const { BuscarUsuarioID } = await import(`../sockets/usuarios${sufix}.js`);
                const { BorrarUsuario, getSocketServerInstance } = await import(`../sockets/sockets${sufix}.js`);
                const { id } = isProd ? await BuscarUsuarioID(ID) : BuscarUsuarioID(ID);
                const io = getSocketServerInstance();
                isProd ? await BorrarUsuario(id, 1, 'SE CERRÓ SESIÓN FORZOSAMENTE') : BorrarUsuario(id, 1, 'SE CERRÓ SESIÓN FORZOSAMENTE');
                io.to(id).emit('usuario-activo', 'Se inició sesión en otro dispositivo. Esta sesión se terminará');
            } catch (err) {
                DesactivarUsuario(datos.id_usuario);
            }
        }
        const ID = await IniciarSesion(datos.id_usuario, datos.id_distrito, req.ip, useragent.parse(req.headers['user-agent']).toString(), 'Inició Sesión');
        datos.id_transaccion = ID;
        const token = await genToken(datos);
        await ActualizarInicio(ID, token);
        const { inicioValidacion, cierreValidacion } = await ConsultaVerificaInicioCierre(datos.id_distrito);
        res.json({
            success: true,
            token,
            inicioValidacion,
            cierreValidacion,
            opcion: userData[0][0].ocultar_opcion,
            msg: `Bienvenido, ${datos.nombre}`
        });
    } catch (err) {
        console.error(`Error en Login: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido al autenticar usuario'
        });
    }
}

export const LoginTitular = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { usuario, contrasena } = req.body;
    try {
        const userData = await SICOVACC.sequelize.query(`SELECT CUS.id_usuario, CUS.usuario, CUS.contrasena, CUS.id_distrito, CCDe.nombre_delegacion, CONCAT(CUS.nombre, ' ', CUS.ape_paterno, ' ', CUS.ape_materno) AS nombre, CUS.id_admin, CUS.perfil, CUS.estatus
        FROM consulta_usuarios_sivacc CUS
        LEFT JOIN (SELECT DISTINCT id_distrito, id_delegacion FROM consulta_cat_distrito) AS CCDi ON CUS.id_distrito = CCDi.id_distrito
        LEFT JOIN consulta_cat_delegacion CCDe ON CCDi.id_delegacion = CCDe.id_delegacion
        WHERE CUS.perfil = 1 AND CUS.id_distrito = ${id_distrito} AND CUS.usuario = '${usuario}'`);
        if (userData[1] == 0)
            return res.status(401).json({
                success: false,
                msg: 'Usuario incorrecto'
            });
        if (userData[0][0].contrasena != contrasena)
            return res.status(401).json({
                success: false,
                msg: 'Contraseña incorrecta'
            });
        if (!userData[0][0].estatus)
            return res.status(403).json({
                success: false,
                msg: 'El usuario no esta activo'
            });
        const datos = {
            id_transaccion,
            id_usuario: userData[0][0].id_usuario,
            nombre: Capitalizar(userData[0][0].nombre.trim()),
            usuario: userData[0][0].usuario,
            id_distrito: userData[0][0].id_distrito,
            nombre_delegacion: userData[0][0].nombre_delegacion,
            perfil: userData[0][0].perfil,
            id_admin: userData[0][0].id_admin
        };
        const token = await genToken(datos);
        await Audit(id_transaccion, id_usuario, id_distrito, 'SOLICITÓ PERMISO AL TITULAR');
        res.json({
            success: true,
            msg: `Permiso concedido por: ${datos.nombre}`,
            token
        });
    } catch (err) {
        console.error(`Error en LoginTitular: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido al autenticar usuario'
        });
    }
}

export const Verify = async (req = request, res = response) => {
    const { id_usuario, id_distrito, exp } = req.data;
    try {
        const { inicioValidacion, cierreValidacion } = await ConsultaVerificaInicioCierre(id_distrito);
        const { opcion } = (await SICOVACC.sequelize.query(`SELECT ocultar_opcion AS opcion FROM consulta_usuarios_sivacc WHERE id_usuario = ${id_usuario}`))[0][0];
        if (!(await SICOVACC.sequelize.query(`SELECT estatus FROM consulta_usuarios_sivacc WHERE id_usuario = ${id_usuario}`))[0][0].estatus)
            return res.status(403).json({
                success: false,
                msg: 'Usuario inactivo'
            });
        //? Verifica si el usuario ya se encuentra conectado al sistema
        const validarToken = await SICOVACC.sequelize.query(`SELECT token FROM consulta_audit WHERE estatus = 1 AND id_usuario = ${id_usuario}`);
        //? En dado caso de que no se encuentre coincidencia se inicia sesión con el token
        if (validarToken[1] == 0) {
            const ID = await IniciarSesion(id_usuario, id_distrito, req.ip, useragent.parse(req.headers['user-agent']).toString(), 'INICIÓ SESIÓN CON EL TOKEN');
            const expire = exp - (Math.floor(Date.now() / 1000));
            req.data.id_transaccion = ID;
            delete req.data.exp;
            const token = await genToken(req.data, expire);
            await ActualizarInicio(ID, token);
            return res.json({
                success: true,
                token,
                inicioValidacion,
                cierreValidacion,
                opcion,
                msg: 'Token verificado'
            });
        } else //? En dado caso que ya se encuentre conectado, verifica que sea el mismo token con el que inicio, si no es asi no prosigue y se le indica que ya se encuentra iniciado desde otro dispositivo
            if (validarToken[0][0].token != req.token)
                return res.status(409).json({
                    success: false,
                    msg: 'Este usuario ya se encuentra conectado desde otro dispositivo'
                });
        res.json({
            success: true,
            inicioValidacion,
            cierreValidacion,
            opcion,
            msg: 'Token verificado'
        });
    } catch (err) {
        console.error(`Error en Verify: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido al verificar el usuario'
        })
    }
}