import { Router } from 'express';
import { param, query } from 'express-validator';
import { AsistenciaUT, BaseDatos, ConsultaCiudadanaDetalle, ConsultaUnidadTerritorial, LevantadaDistrito, MesasConComputo, MesasSinComputo, OpinionesDemarcacion, OpinionesDistrito, OpinionesMesa, OpinionesUT, Participacion, ProyectosEmpatePrimerLugar, ProyectosEmpateSegundoLugar, ProyectosParticipantes, ProyectosPrimerLugar, ProyectosSegundoLugar, ProyectosSinOpiniones, UTConComputoGA } from '../controllers/centConsultaReportesExcel.controller.js';
import { chkDistrito } from '../middlewares/dist-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

//? F1 - Base de Datos - Excel

router.get('/baseDatos/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], BaseDatos);

//? F2 - Concentrado de Proyectos participantes por Distrito y Unidad Territorial - Excel

router.get('/proyectosParticipantes/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosParticipantes);

//? F4 - Validación de Resultados de la Consulta Ciudadana Detalle Mesa - Excel

router.get('/consultaCiudadanaDetalle/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ConsultaCiudadanaDetalle);

//? F5 - Resultado de Opiniones por Mesa - Excel

router.get('/opinionesMesa/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], OpinionesMesa);

//? F6 - Validación de Resultados de la Consulta por Unidad Territorial - Excel

router.get('/consultaUnidadTerritorial/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ConsultaUnidadTerritorial);

//? F7 - Concentrado de Opiniones por Unidad Territorial - Escel

router.get('/opinionesUT/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], OpinionesUT)

//? F8 - Proyectos por Unidad Territorial que Obtuvieron el Primer Lugar en la Consulta de Presupuesto Participativo - Excel

router.get('/proyectosPrimerLugar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosPrimerLugar);

//? F9 - Casos de Empate de los Proyectos que Obtuvieron el Primer Lugar - Excel

router.get('/proyectosEmpatePrimerLugar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosEmpatePrimerLugar);

//? F10 - Concentrado de Unidades Territoriales que NO Recibieron Opiniones - Excel

router.get('/proyectosSinOpiniones/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosSinOpiniones);

//? F11 - Reporte Asistencia por Unidad Territorial - Excel

router.get('/asistenciaUT/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], AsistenciaUT);

//? F12 - Mesas con Cómputo Capturado - Excel

router.get('/MesasConComputo/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], MesasConComputo);

//? F13 - Mesas sin Cómputo Capturado - Excel

router.get('/MesasSinComputo/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], MesasSinComputo);

//? F14 - Concentrado de Unidades Territoriales por Distrito Electoral con Cómputo Capturado (Grado de Avance) - Excel

router.get('/UTConComputoGA', [
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], UTConComputoGA);

//? F15 - Opiniones por Distrito - Excel

router.get('/opinionesDistrito', [
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], OpinionesDistrito);

//? F16 - Opiniones por Demarcación - Excel

router.get('/opinionesDemarcacion', [
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], OpinionesDemarcacion);

//? Proyectos por Unidad Territorial que Obtuvieron el Segundo Lugar en la Consulta de Presupuesto Participativo - Excel

router.get('/proyectosSegundoLugar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosSegundoLugar);

//? Casos de Empates de los Proyectos que Obtuvieron el Segundo Lugar - Excel

router.get('/proyectosEmpateSegundoLugar/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], ProyectosEmpateSegundoLugar);

//? Actas Levantadas en Dirección Distrital - Excel

router.get('/levantadaDistrito/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], LevantadaDistrito);

//? Porcentaje de Participación - Excel

router.get('/participacion', [
    query('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator
], Participacion);

export { router as centConsultaReportesRouter };

