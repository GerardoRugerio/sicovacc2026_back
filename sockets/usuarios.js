let usuarios = [];

export const getUsuarios = () => usuarios;

export const AgregarUsuario = id => usuarios.push({ id, id_transaccion: null });

export const BuscarUsuario = id => usuarios.find(usuario => usuario.id == id);

export const BuscarUsuarioID = id => usuarios.find(usuario => usuario.id_transaccion == id);

export const ActualizarUsuario = (id, id_transaccion = null) => {
    for (let usuario of usuarios)
        if (usuario.id == id) {
            usuario.id_transaccion = id_transaccion;
            break;
        }
}

export const EliminarUsuario = id => usuarios = usuarios.filter(usuario => usuario.id != id);