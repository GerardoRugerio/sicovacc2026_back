import { Router } from 'express';
import { param } from 'express-validator';
import { Incidentes, InicioCierreValidacion } from '../controllers/centReportesExcel.controller.js';
import { chkDistrito } from '../middlewares/dist-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

//? Reporte de asistencia de inicio y cierre de la validación - Excel

router.get('/inicioCierreValidacion', InicioCierreValidacion);

//? F3 - Incidentes Presentados Durante la Validación de la Consulta de Presupuesto Participativo - Excel

router.get('/incidentes/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    // query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], Incidentes);

export { router as centReportesRouter };

