import { Router } from 'express';
import { body, param } from 'express-validator';
import { ListaFormulas, ListaProyectos } from '../controllers/distReportes.controller.js';
import { IncidentesDistrito, InicioCierreValidacion } from '../controllers/distReportesExcel.controller.js';
import { chkToken, dataToken } from '../middlewares/auth-middleware.js';
import { chkDistrito } from '../middlewares/dist-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

//? Consulta de Proyectos

router.post('/proyectos', [
    chkToken(),
    dataToken,
    body('clave_colonia').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 o 3'),
    Validator
], ListaProyectos);

//? Consulta de Fórmulas

router.post('/formulas', [
    chkToken(),
    dataToken,
    body('clave_colonia').exists().notEmpty().isString(),
    Validator
], ListaFormulas);

//? Inicio - Cierre de Validación - Excel

router.get('/inicioCierreValidacion/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], InicioCierreValidacion);

//? F3 - Listado de Incidentes Presentados en la Validación de la Consulta de Presupuesto Participativo - Excel

router.get('/incidentesDistrito/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    // query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], IncidentesDistrito);

export { router as distReportesRouter };

