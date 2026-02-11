import { DataTypes } from 'sequelize';
import { sqlConnector_SICOVACC } from '../database/config.js';

export const SICOVACC = sqlConnector_SICOVACC.define('consulta_usuarios_sivacc', {
    id_usuario: { type: DataTypes.INTEGER },
    nombre: { type: DataTypes.STRING },
    ape_paterno: { type: DataTypes.STRING },
    ape_materno: { type: DataTypes.STRING },
    id_distrito: { type: DataTypes.INTEGER },
    usuario: { type: DataTypes.STRING },
    contrasena: { type: DataTypes.STRING },
    id_admin: { type: DataTypes.INTEGER },
    fecha_alta: { type: DataTypes.STRING },
    fecha_modif: { type: DataTypes.STRING },
    status: { type: DataTypes.STRING },
    perfil: { type: DataTypes.INTEGER }
}, { freezeTableName: true });