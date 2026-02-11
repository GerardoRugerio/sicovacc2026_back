import { Router } from 'express';
import { distConsultaReportesRouter } from './distConsultaReportes.router.js';
import { distEleccionReportesRouter } from './distEleccionReportes.router.js';
import { distProcesosRouter } from './distProcesos.router.js';
import { distReportesRouter } from './distReportes.router.js';
import { distSeguimientoRouter } from './distSeguimiento.router.js';

const router = Router();

router.use('/seguimiento', distSeguimientoRouter);

router.use('/reportes', distReportesRouter);

//? COPACO

router.use('/reportes/eleccion', distEleccionReportesRouter);

//? Presupuesto Participativo

router.use('/reportes/consulta', distConsultaReportesRouter);

router.use('/procesos', distProcesosRouter);

export { router as distritalRouter };

