import { request, response } from 'express';
import { EstadoUT, MesasFalt, MesasI } from '../helpers/Consultas.js';

//?   undefined - Dato pasado por parametro ('todos', o del 1 al 33)
//?   0 - Central - Dato pasado por body
//?   1 -> 33 - Distritos - Dato pasado por body

export const chkDistrito = (distrito = undefined, todos = false) => {
    return (req = request, res = response, next) => {
        const { id_distrito } = distrito ? req.params : req.body;
        if (!distrito) {
            if (!id_distrito || (isNaN(id_distrito) && id_distrito.toLowerCase() != 'todos') || (id_distrito < 1 || id_distrito > 33))
                return res.status(400).json({
                    success: false,
                    msg: `El id_distrito debe de ser del 1 al 33${todos ? ` o en su defecto 'TODOS'` : ''}`
                });
        } else {
            if (!(+id_distrito >= distrito && +id_distrito <= 33))
                return res.status(400).json({
                    success: false,
                    msg: `${id_distrito} no existe`
                });
        }
        next();
    }
}

export const StatusReporte = (required = true) => {
    return async (req = request, res = response, next) => {
        const { id_distrito } = req.params;
        const { clave_colonia } = req.body;
        const anio = req.body.anio ?? 1;
        if (!clave_colonia) {
            if (required)
                return res.status(500).json({
                    success: false,
                    msg: 'La clave_colonia es requerida'
                });
            return next();
        }
        const { mesasI } = await MesasI(id_distrito, clave_colonia, anio);
        if (!mesasI)
            return res.status(400).json({
                success: false,
                msg: 'Esta Unidad Territorial no tiene ninguna mesa instalada'
            });
        const validada = await EstadoUT(clave_colonia, anio);
        const mesasFalt = await MesasFalt(id_distrito, clave_colonia, anio);
        if (!validada)
            return res.status(400).json({
                success: false,
                msg: `Esta Unidad Territorial todavía no está validada, ${mesasFalt == 1 ? `falta ${mesasFalt} Mesa` : `faltan ${mesasFalt} Mesas`} por Validar`
            });
        next();
    }
    // const { id_distrito } = req.params;
    // const { clave_colonia } = req.body;
    // const anio = req.body.anio ?? 1;
    // const { mesasI } = await MesasI(id_distrito, clave_colonia, anio);
    // if (mesasI) {
    //     const validada = await EstadoUT(clave_colonia, anio);
    //     const mesasFalt = await MesasFalt(id_distrito, clave_colonia, anio);
    //     if (!validada)
    //         return res.status(400).json({
    //             success: false,
    //             msg: `Esta Unidad Territorial todavía no está validada, ${mesasFalt == 1 ? `falta ${mesasFalt} Mesa` : `faltan ${mesasFalt} Mesas`} por Validar`
    //         });
    // } else
    //     return res.status(400).json({
    //         success: false,
    //         msg: 'Esta Unidad Territorial no tiene ninguna mesa instalada'
    //     })
    // next();
}