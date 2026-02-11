import { DataTypes } from 'sequelize';
import { sqlConnector_VSEI } from '../database/config.js';

export const VTA_RES_PROYECTOS_2 = sqlConnector_VSEI.define('VTA_RES_PROYECTOS_2', {
    ID_DISTRITO: { type: DataTypes.INTEGER },
    ID_DELEGACION: { type: DataTypes.INTEGER },
    CLAVE_COLONIA: { type: DataTypes.STRING },
    ID_MRO: { type: DataTypes.STRING },
    NUM_PROYECTO: { type: DataTypes.INTEGER },
    TOTAL: { type: DataTypes.INTEGER },
    ANIO: { type: DataTypes.INTEGER }
}, { freezeTableName: true });