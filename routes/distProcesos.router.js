import { Router } from 'express';
import { body } from 'express-validator';
import { ActualizarDatosDistrito, DatosDistrito, LimpiarBD } from '../controllers/distProcesos.controller.js';
import { chkToken, dataToken } from '../middlewares/auth-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

//? Actualización de Datos del Disrtito

router.get('/datosDistrito', [
    chkToken(),
    dataToken
], DatosDistrito);

router.put('/datosDistrito', [
    chkToken(),
    dataToken,
    body('domicilio').exists().notEmpty().isString(),
    // body('codigo_postal').exists().notEmpty().isNumeric(),
    body('coordinador').isString(),
    body('coordinador_puesto').isString(),
    body('coordinador_genero').isString(),
    body('secretario').isString(),
    body('secretario_puesto').isString(),
    body('secretario_genero').isString(),
    Validator
], ActualizarDatosDistrito);

//? Limpiar la Base de Datos

router.get('/limpiarBD', [
    chkToken([1]),
    dataToken
], LimpiarBD);

export { router as distProcesosRouter };