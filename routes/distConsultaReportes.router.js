import { request, response, Router } from 'express';
import { body, param, query } from 'express-validator';
import { LevantadaDistrito, ListadoProyectos, MesasConComputo, MesasSinComputo, ProyectosEmpatePrimerLugar, ProyectosEmpateSegundoLugar, ProyectosOpinar, ProyectosPrimerLugar, ProyectosSegundoLugar, ProyectosUTSinOpiniones, ResultadosOpiMesa, UTPorValidar, UTValidadas, ValidacionResultados, ValidacionResultadosDetalle, ValidacionResultadosNombre, ValidacionResultadosNombreDetalle } from '../controllers/distConsultaReportesExcel.controller.js';
import { ActaValidacionPDF, ProyectosParticipantes } from '../controllers/distConsultaReportesPDF.controller.js';
import { ActaValidacionWord } from '../controllers/distConsultaReportesWord.controller.js';
import { chkDistrito, StatusReporte } from '../middlewares/dist-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

//? Consulta de Resultados Por Unidad Territorial - Excel

router.get('/UTValidadas/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], UTValidadas);

router.get('/UTPorValidar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], UTPorValidar);

//? F2 - Concentrado de Proyectos Participantes por Unidad Territorial - Excel

router.get('/listadoProyectos/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ListadoProyectos);

//? F4 - Validación de Resultados de la Consulta por Unidad Territorial - Excel

router.post('/validacionResultados/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    body('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator,
    StatusReporte(false)
], ValidacionResultados);

//? F5 - Validación de Resultados de la Consulta Detalle Mesa - Excel

router.post('/validacionResultadosDetalle/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    body('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator,
    StatusReporte(false)
], ValidacionResultadosDetalle);

//? F6 - Validación de Resultados de la Consulta por Nombre del Proyecto - Escel

router.post('/validacionResultadosNombre/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    body('clave_colonia').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator,
    StatusReporte()
], ValidacionResultadosNombre);

//? F7 - Validación de Resultados de la Consulta por Nombre del Proyecto (Detalle por Mesa) - Excel

router.post('/validacionResultadosNombreDetalle/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    body('clave_colonia').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator,
    StatusReporte()
], ValidacionResultadosNombreDetalle);

//? F8 - Mesas Con Cómputo Capturado - Excel

router.get('/MesasConComputo/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], MesasConComputo);

//? F9 - Mesas Sin Cómputo Capturado - Excel

router.get('/MesasSinComputo/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], MesasSinComputo);

//? F10 - Resultados de Opiniones por Mesa - Excel

router.get('/resultadosOpiMesa/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ResultadosOpiMesa);

//? F11 - Proyectos por Unidad Territorial que Obtuvieron el Primer Lugar - Excel

router.get('/proyectosPrimerLugar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosPrimerLugar);

//? F12 - Proyectos por Unidad Territorial que Obtuvieron el Segundo Lugar - Excel

router.get('/proyectosSegundoLugar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosSegundoLugar);

//? F13 - Proyectos Empatados que Obtuvieron el Primer Lugar - Excel

router.get('/proyectosEmpatePrimerLugar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosEmpatePrimerLugar);

//? F14 - Proyectos Empatados que Obtuvieron el Segundo Lugar - Excel

router.get('/proyectosEmpateSegundoLugar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosEmpateSegundoLugar);

//? F15 - Unidades Territoriales que NO Recibieron Opiniones - Excel

router.get('/proyectosUTSinOpiniones/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosUTSinOpiniones);

//? Proyectos a Opinar - Excel

router.get('/proyectosOpinar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosOpinar);

//? Levantada en Distrito - Excel

router.get('/levantadaDistrito/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], LevantadaDistrito);

//? Proyectos Participantes Dictaminados Favorablemente - PDF

router.post('/proyectosParticipantes/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    body('clave_colonia').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosParticipantes);

//? Constancia - En Desuso - PDF - Word

router.post('/constancia/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    body('clave_colonia').exists().notEmpty().isString(),
    body('tipo').exists().notEmpty().isString().isIn(['PDF', 'WORD']).withMessage(`El tipo debe de ser 'PDF' o 'WORD'`),
    body('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator,
    StatusReporte()
], async (req = request, res = response) => {
    const { tipo } = req.body;
    if (tipo.toLowerCase() == 'pdf')
        ConstanciaPDF(req, res);
    else
        ConstanciaWord(req, res);
});

//? Acta de Validación - PDF - Word

router.post('/actaValidacion/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    body('clave_colonia').exists().notEmpty().isString(),
    body('tipo').exists().notEmpty().isString().isIn(['PDF', 'WORD']).withMessage(`El tipo debe de ser 'PDF' o 'WORD'`),
    body('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator,
    StatusReporte()
], async (req = request, res = response) => {
    const { tipo } = req.body;
    if (tipo.toLowerCase() == 'pdf')
        ActaValidacionPDF(req, res);
    else
        ActaValidacionWord(req, res);
});

export { router as distConsultaReportesRouter };

