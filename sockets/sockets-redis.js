import { createAdapter } from '@socket.io/redis-adapter';
import { createClient } from 'redis';

import { CerrarSesion } from '../helpers/Audit.js';
import { ActualizarUsuario, AgregarUsuario, BuscarUsuario, EliminarUsuario, getUsuarios } from './usuarios-redis.js';

let io = null;

//? Se crea la instancia del socket
export const setSocketServerInstanceRedis = async serverInstance => {
    io = serverInstance;
    //? Configurar Redis adapter
    const pubClient = createClient({ url: 'redis://127.0.0.1:6379' });
    const subClient = pubClient.duplicate();
    pubClient.on('error', err => console.error(`Error en Redis Pub: ${err}`));
    subClient.on('error', err => console.error(`Error en Redis Sub: ${err}`));
    await pubClient.connect();
    await subClient.connect();
    io.adapter(createAdapter(pubClient, subClient));
    await pubClient.flushDb();
    console.log('Socket.IO conectado a Redis');
    Sockets(io);
}

//? Regresa la instancia del socket
export const getSocketServerInstance = () => io;

//? Eventos y funciones para el socket
const Sockets = io => {
    io.on('connect', socket => {
        AgregarUsuario(socket.id);
        ConfigurarUsuarios(socket);
        DesconectarCliente(socket);
    });
};

//? Configura el usuario, actualiza su id_transaccion
const ConfigurarUsuarios = cliente => {
    cliente.on('configurar-usuario', async (payload, callback) => {
        await ActualizarUsuario(cliente.id, payload.id_transaccion);
        console.log('==> Usuarios <==\n', await getUsuarios());
        if (callback)
            callback({
                success: true,
                msg: 'Usuario configurado correctamente'
            });
    });
}

//? Si cierra sesión o si pierde la conexión ejecuta la función de BorrarUsuario
const DesconectarCliente = socket => {
    socket.on('logout', () => BorrarUsuario(socket.id, 1, 'CERRÓ SESIÓN'));
    socket.on('disconnect', () => BorrarUsuario(socket.id, 0, 'PERDIÓ LA CONEXIÓN'));
}

//? Borra al usuario o actualiza al usuario quitandole su id_transaccion (solo si cierra sesión)
export const BorrarUsuario = async (id, motivo, descripcion) => {
    const { id_transaccion } = await BuscarUsuario(id);
    if (id_transaccion)
        CerrarSesion(id_transaccion, descripcion);
    if (motivo == 1) {
        await ActualizarUsuario(id);
        return;
    }
    await EliminarUsuario(id);
}