const envType = process.env.NODE_ENV || process.argv.includes('--env') && process.argv[process.argv.indexOf('--env') + 1] || 'prod';
const isProd = envType === 'prod';

module.exports = {
    apps: [{
        name: 'SICOVACC',                           //? Nombre de la app
        namespace: 'IECM',                          //? Opcional: agrupa procesos
        script: './index.js',                       //? Punto de entrada del backend
        node_args: '--max-old-space-size=2048',
        exec_mode: isProd ? 'cluster' : 'fork',      //? Modo fork = un solo proceso; 'cluster' para escalar
        ...(isProd ? { instances: 2 } : {}),         //? max para usar todos los nucleos 
        watch: !isProd,
        ignore_watch: [
            'logs',
            'node_modules',
            'plantillas',
            'resources',
            'views'
        ],
        restart_delay: 5000,                        //? Espera 5s antes de reiniciar tras un crasheo
        error_file: './logs/err.log',               //? Log de errores
        out_file: './logs/out.log',                 //? Log normal
        log_date_format: 'DD/MM/YYYY HH:mm:ss',
        autorestart: true,                          //? Reinicia si crashea o sale con error
        min_uptime: 5000,                           //? Debe estar vivo al menos 5s para considerarse estable
        kill_timeout: 3000,
        env: { NODE_ENV: 'prod' },
        env_dev: { NODE_ENV: 'dev' }
    }]
}