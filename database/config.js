import { config } from 'dotenv';
import { Sequelize } from 'sequelize';

/**
 * Instancia de Sequelize para la base de datos
 * 
 * Opciones:
 * - dialect: 'mssql' → Motor de base de datos
 * - dialectOptions: Configuración específica para SQL Server
 *   - encrypt: Uso de conexión cifrada
 * - define.timestamps = false → Desactiva campos automáticos createdAt/updatedAt
 * - pool: Configuración de conexiones:
 *   - min: Número mínimo de conexiones
 *   - max: Número maximo de conexiones
 *   - acquire: Tiempo máximo (ms) que sequelize intentará obtener una conexión antes de lanzar error
 *   - idle: Tiempo máximo (ms) que una conexión puede estar inactiva antes de liberarse
 * - logging: Habilita logs de Sequelize
 */
const crearConexion = (nombre, envConfig) => {
    console.log(`Ambiente ${envConfig.tipo}: ${nombre}`);
    return new Sequelize(envConfig.database, envConfig.user, envConfig.pass, {
        host: envConfig.host,
        port: envConfig.port,
        dialect: 'mssql',
        dialectOptions: {
            options: {
                encrypt: envConfig.encrypt
            }
        },
        define: {
            timestamps: false
        },
        pool: {
            min: 0,
            max: 10,
            acquire: 10000,
            idle: 5000
        },
        // logging: true,
        retry: {
            match: [
                /ConnectionError/,
                /ConnectionRefusedError/,
                /TimedOutError/,
                /HostNotReachableError/,
                /AccessDeniedError/
            ],
            max: 3
        }
    });
}

const conectarBD = async (nombre, conexion, maxRetries = 5) => {
    let retries = maxRetries;
    while (retries > 0) {
        try {
            await conexion.authenticate();
            console.log(`Conexión establecida: BD ${nombre}`);
            listenForDisconnect(nombre, conexion);
            return true;
        } catch (err) {
            retries--;
            console.error(`Error al conectar a la BD ${nombre}: ${err}`);
            if (retries == 0) throw err;
            console.error(`Reintentando conexión a BD ${nombre}`);
            await new Promise(res => setTimeout(res, 3000));
        }
    }
}

const listenForDisconnect = (nombre, conexion) => {
    const cm = conexion.connectionManager;
    cm.on?.('error', async err => {
        console.error(`Error en conexion con la BD ${nombre}: ${err.message}`);
        if (err.message.includes('ECONNRESET') || err.message.includes('SequelizeConnectionError')) {
            console.log(`Intentando reconectar a la BD ${nombre}...`);
            await reconnect(nombre, conexion);
        }
    });
}

const reconnect = async (nombre, conexion) => {
    let retries = 5;
    while (retries > 0) {
        try {
            await conexion.authenticate();
            console.log(`Reconexión exitosa a la BD ${nombre}`);
            return;
        } catch (err) {
            retries--;
            console.error(`Error al reconectar a la BD ${nombre}: ${err.message}`);
            if (retries > 0)
                await new Promise(res => setTimeout(res, 3000));
            else {
                console.error(`No se pudo reconectar a la BD ${nombre}`);
                process.exit(1);
            }
        }
    }
}

/**
 * Configuración de conexión a la base de datos (SICOVACC y SEI)
 * 
 * - Se obtienen las variables desde .env
 * - El objeto contiene:
 *   - tipo: Tipo de ambiente (Desarrollo, Producción, Contingencia, etc)
 *   - database: Nombre de la base de datos
 *   - user: Usuario de conexión
 *   - pass: Contraseña
 *   - host: Dirección del servidor SQL
 *   - port: Puerto del servidor SQL
 *   - encrypt: Indica si se requiere encriptar la conexión
 */
export const sqlConnector_SICOVACC = crearConexion('SICOVACC', JSON.parse(config().parsed.DB_SICOVACC));
export const sqlConnector_VSEI = crearConexion('SEI', JSON.parse(config().parsed.DB_SEI));

export const conectarTodas = async () => await Promise.all([
    conectarBD('SICOVACC', sqlConnector_SICOVACC),
    conectarBD('SEI', sqlConnector_VSEI)
]);