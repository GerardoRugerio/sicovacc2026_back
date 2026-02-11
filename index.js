import cors from 'cors';
import { config } from 'dotenv';
import express, { json } from 'express';
import fs from 'fs';
import { createServer } from 'https';
import path, { dirname } from 'path';
import { Server } from 'socket.io';
import { fileURLToPath } from 'url';
import useragent from 'useragent';
import { conectarTodas } from './database/config.js';

import { administradorRouter } from './routes/administrador.router.js';
import { catRouter } from './routes/catalogos.router.js';
import { centralRouter } from './routes/central.router.js';
import { distritalRouter } from './routes/distrital.router.js';
import { userRoutes } from './routes/user.router.js';

useragent(true); //? Sirve para formatear el useragent enviado en la cabecera de las peticiones

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const sslOption = {
    key: fs.readFileSync(path.join(__dirname, 'certificate/server.key')),
    cert: fs.readFileSync(path.join(__dirname, 'certificate/server.cert'))
};
const app = express();
const server = createServer(sslOption, app);
const HOST = config().parsed.HOST;
const PORT = config().parsed.PORT;
const io = new Server(server, { cors: { origin: true, credentials: true } });

app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    optionsSuccessStatus: 204
}));
app.options('*', cors());

app.use(cors({ origin: '*' }));
app.use(json());
app.use(express.static(path.join(__dirname, 'views')));

await (async () => {
    try {
        await conectarTodas();
    } catch (err) {
        console.error(`Error fatal al iniciar conexiones ${err}`);
        process.exit(1);
    }
})();

if (process.env.NODE_ENV && process.env.NODE_ENV.match('prod')) {
    const { setSocketServerInstanceRedis } = await import('./sockets/sockets-redis.js');
    (async () => await setSocketServerInstanceRedis(io))();
} else {
    const { setSocketServerInstance } = await import('./sockets/sockets.js');
    setSocketServerInstance(io);
}

app.get('/', (_, res) => res.sendFile(path.join(__dirname, 'views/index.html')));

app.use('/api', userRoutes);

app.use('/api/administrador', administradorRouter);

app.use('/api/distrital', distritalRouter);

app.use('/api/central', centralRouter);

app.use('/api/cat', catRouter);

server.listen(PORT, () => console.log(`Servidor escuchando en ${HOST}:${PORT}`));