import { Router } from 'express';
import { body } from 'express-validator';
import { Actas, EliminarActa } from '../controllers/central.controller.js';
import { chkToken, dataToken } from '../middlewares/auth-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

//? Captura de Resultados de Consulta por Mesa

router.post('/acta', [
    chkToken([4]),
    body('id_distrito').exists().notEmpty().isNumeric(),
    body('clave_colonia').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], Actas);

router.delete('/acta', [
    chkToken([4]),
    dataToken,
    body('id_acta').exists().notEmpty().isNumeric(),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], EliminarActa);

export { router as centSeguimientoRouter };