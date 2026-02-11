import { CerrarSesion } from '../helpers/Audit.js';
import { ActualizarUsuario, AgregarUsuario, BuscarUsuario, EliminarUsuario, getUsuarios } from './usuarios.js';

let io = null;

//? Se crea la instancia del socket
export const setSocketServerInstance = serverInstance => {
    io = serverInstance;
    Sockets(serverInstance);
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
    cliente.on('configurar-usuario', (payload, callback) => {
        ActualizarUsuario(cliente.id, payload.id_transaccion);
        console.log('==> Usuarios <==\n', getUsuarios());
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
export const BorrarUsuario = (id, motivo, descripcion) => {
    const { id_transaccion } = BuscarUsuario(id);
    if (id_transaccion)
        CerrarSesion(id_transaccion, descripcion);
    if (motivo == 1) {
        ActualizarUsuario(id);
        return;
    }
    EliminarUsuario(id);
}