import { config } from 'dotenv';
import { request, response } from 'express';
import JWT from 'jsonwebtoken';
import { BuscarUsuarioID } from '../sockets/usuarios.js';
import { BorrarUsuario } from '../sockets/sockets.js';

//! Roles
//?   1 - Titular
//?   2 - Capturista
//?   3 - Central
//?   4 - DEOEyG
//?   99 - Administrador

export const chkToken = (rol = undefined) => {
    return async (req = request, res = response, next) => {
        const { authorization } = req.headers;
        if (!authorization || authorization.split(' ')[0] != 'Bearer' || ['null', 'undefined'].includes(authorization.split(' ')[1]))
            return res.status(401).json({
                success: false,
                msg: 'No hay Token'
            });
        const token = authorization.split(' ')[1];
        try {
            const { perfil } = JWT.verify(token, config().parsed.SECRET);
            if (rol && !rol.includes(perfil))
                return res.status(403).json({
                    success: false,
                    msg: 'No tienes permiso'
                })
            req.token = token;
            next();
        } catch (err) {
            if (err.name.match('TokenExpiredError')) {
                try {
                    const { id_transaccion } = JWT.decode(token);
                    const { id } = BuscarUsuarioID(id_transaccion);
                    BorrarUsuario(id, 1, 'TOKEN EXPIRADO');
                } catch (e) { }
                return res.status(401).json({
                    success: false,
                    msg: 'Token expirado'
                });
            }
            res.status(401).json({
                success: false,
                msg: 'Token invalido'
            });
        }
    }
}

export const dataToken = async (req = request, res = response, next) => {
    const token = req.token;
    const { id_transaccion, id_usuario, nombre, usuario, id_distrito, nombre_delegacion, perfil, id_admin, exp } = JWT.decode(token);
    const datos = { id_transaccion, id_usuario, nombre, usuario, id_distrito, nombre_delegacion, perfil, id_admin, exp };
    req.data = datos;
    next();
}