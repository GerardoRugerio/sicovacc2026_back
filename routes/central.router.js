import { Router } from 'express';
import { centConsultaReportesRouter } from './centConsultaReportes.router.js';
import { centEleccionReportesRouter } from './centEleccionReportes.router.js';
import { centReportesRouter } from './centReportes.router.js';
import { centSeguimientoRouter } from './centSeguimiento.router.js';

const router = Router();

router.use('/seguimiento', centSeguimientoRouter);

router.use('/reportes', centReportesRouter);

//? COPACO

router.use('/reportes/eleccion', centEleccionReportesRouter);

//? Presupuesto Participativo

router.use('/reportes/consulta', centConsultaReportesRouter);

export { router as centralRouter };

