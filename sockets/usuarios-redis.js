import { createClient } from 'redis';

const client = createClient({ url: 'redis://127.0.0.1:6379' });
client.on('error', err => console.error(`Error en Redis (usuarios-redis.js): ${err}`));

await client.connect();

const USUARIOS_KEY = 'usuarios_conectados';

const limpiarUsuariosConectados = async () => await client.del(USUARIOS_KEY);

limpiarUsuariosConectados();

export const getUsuarios = async () => {
    const data = await client.hGetAll(USUARIOS_KEY);
    return Object.values(data).map(u => JSON.parse(u));
}

export const AgregarUsuario = async id => await client.hSet(USUARIOS_KEY, id, JSON.stringify({ id, id_transaccion: null }));

export const BuscarUsuario = async id => {
    const data = await client.hGet(USUARIOS_KEY, id);
    return data ? JSON.parse(data) : {};
}

export const BuscarUsuarioID = async id_transaccion => {
    const data = await getUsuarios();
    return data.find(u => u.id_transaccion == id_transaccion);
}

export const ActualizarUsuario = async (id, id_transaccion = null) => {
    const usuario = await BuscarUsuario(id);
    if (usuario) {
        usuario.id_transaccion = id_transaccion;
        await client.hSet(USUARIOS_KEY, id, JSON.stringify(usuario));
    }
}

export const EliminarUsuario = async id => await client.hDel(USUARIOS_KEY, id);