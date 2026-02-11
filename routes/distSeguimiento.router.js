import { Router } from 'express';
import { body, query } from 'express-validator';
import { ActualizarActa, ActualizarInicioValidacion, DatosActa, DatosCierreValidacion, DatosInicioValidacion, EditarIncidente, EliminarIncidente, EstadoBaseDatos, GuardarActualizarCierreValidacion, GuardarIncidente, GuardarInicioValidacion, GuardarMesasInstaladas, MesasInstaladas, RegistrarActa, RegistrosIncidentes, ResultadoConsultaMesa, VerificarConsultaMesa } from '../controllers/distSeguimiento.controller.js';
import { chkToken, dataToken } from '../middlewares/auth-middleware.js';
import { decryptPayload } from '../middlewares/encrypt-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

//? Estado de la Base de Datos

router.get('/estadoBD', [
    chkToken(),
    dataToken,
], EstadoBaseDatos);

//? Inicio de Validación

router.get('/inicioValidacion', [
    chkToken(),
    dataToken
], DatosInicioValidacion);

router.post('/inicioValidacion', [
    chkToken(),
    dataToken,
    body('MSPEN').exists().notEmpty().isNumeric(),
    body('COPACO').exists().notEmpty().isNumeric(),
    // body('personasCandidatas').exists().notEmpty().isNumeric(),
    body('personasObservadoras').exists().notEmpty().isNumeric(),
    body('presentaronProyecto').exists().notEmpty().isNumeric(),
    body('mediosComunicacion').exists().notEmpty().isNumeric(),
    body('otros').exists().notEmpty().isNumeric(),
    body('total').exists().notEmpty().isNumeric(),
    body('fecha').exists().notEmpty().isDate(),
    body('hora').exists().notEmpty().isTime(),
    body('observaciones').exists().notEmpty().isString(),
    Validator
], GuardarInicioValidacion);

router.put('/inicioValidacion', [
    chkToken(),
    dataToken,
    body('MSPEN').exists().notEmpty().isNumeric(),
    body('COPACO').exists().notEmpty().isNumeric(),
    // body('personasCandidatas').exists().notEmpty().isNumeric(),
    body('personasObservadoras').exists().notEmpty().isNumeric(),
    body('presentaronProyecto').exists().notEmpty().isNumeric(),
    body('mediosComunicacion').exists().notEmpty().isNumeric(),
    body('otros').exists().notEmpty().isNumeric(),
    body('total').exists().notEmpty().isNumeric(),
    body('fecha').exists().notEmpty().isDate(),
    body('hora').exists().notEmpty().isTime(),
    body('observaciones').exists().notEmpty().isString(),
    Validator
], ActualizarInicioValidacion);

//? Mesas Instaladas

router.get('/mesasInstaladas', [
    chkToken(),
    dataToken,
    query('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], MesasInstaladas);

router.put('/mesasInstaladas', [
    chkToken(),
    dataToken,
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    body('mesas').exists().notEmpty().isArray({ min: 1 }).withMessage('El minimo de objetos es 1. Debe de contener clave_colonia, num_mro, tipo_mro y noInstaladas'),
    body('mesas.*.clave_colonia').exists().notEmpty().isString(),
    body('mesas.*.num_mro').exists().notEmpty().isNumeric(),
    body('mesas.*.tipo_mro').exists().notEmpty().isNumeric(),
    body('mesas.*.noInstalada').exists().notEmpty().isBoolean(),
    Validator
], GuardarMesasInstaladas);

//? Registros de Incidentes

router.get('/incidentes', [
    chkToken(),
    dataToken,
    query('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], RegistrosIncidentes);

router.post('/incidentes', [
    chkToken(),
    dataToken,
    body('clave_colonia').exists().notEmpty().isString(),
    body('num_mro').exists().notEmpty().isNumeric(),
    body('tipo_mro').exists().notEmpty().isNumeric(),
    body('incidente_1').exists().notEmpty().isBoolean(),
    body('incidente_2').exists().notEmpty().isBoolean(),
    body('incidente_3').exists().notEmpty().isBoolean(),
    body('incidente_4').exists().notEmpty().isBoolean(),
    body('incidente_5').exists().notEmpty().isBoolean(),
    // body('incidente_6').exists().notEmpty().isBoolean(),
    // body('incidente_7').exists().notEmpty().isBoolean(),
    // body('incidente_8').exists().notEmpty().isBoolean(),
    body('fecha').exists().notEmpty().isDate(),
    body('hora').exists().notEmpty().isTime(),
    body('participantes').exists().notEmpty().isString(),
    body('hechos').exists().notEmpty().isString(),
    body('acciones').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], GuardarIncidente);

router.put('/incidentes', [
    chkToken(),
    dataToken,
    body('id_incidente').exists().notEmpty().isNumeric(),
    body('num_mro').exists().notEmpty().isNumeric(),
    body('tipo_mro').exists().notEmpty().isNumeric(),
    body('incidente_1').exists().notEmpty().isBoolean(),
    body('incidente_2').exists().notEmpty().isBoolean(),
    body('incidente_3').exists().notEmpty().isBoolean(),
    body('incidente_4').exists().notEmpty().isBoolean(),
    body('incidente_5').exists().notEmpty().isBoolean(),
    // body('incidente_6').exists().notEmpty().isBoolean(),
    // body('incidente_7').exists().notEmpty().isBoolean(),
    // body('incidente_8').exists().notEmpty().isBoolean(),
    body('fecha').exists().notEmpty().isDate(),
    body('hora').exists().notEmpty().isTime(),
    body('participantes').exists().notEmpty().isString(),
    body('hechos').exists().notEmpty().isString(),
    body('acciones').exists().notEmpty().isString(),
    Validator
], EditarIncidente);

router.delete('/incidentes', [
    chkToken([1]),
    dataToken,
    body('id_incidente').exists().notEmpty().isNumeric(),
    Validator
], EliminarIncidente);

//? Captura de Resultados de Consulta por Mesa

router.get('/resultadoConsultaMesa', [
    chkToken(),
    dataToken,
    query('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], ResultadoConsultaMesa);

router.post('/resultadoConsultaMesa', [
    chkToken(),
    dataToken,
    body('clave_colonia').exists().notEmpty().isString(),
    body('num_mro').exists().notEmpty().isNumeric(),
    body('tipo_mro').exists().notEmpty().isNumeric(),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], VerificarConsultaMesa);

router.get('/acta/:id_acta', [
    chkToken(),
    dataToken,
    query('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], DatosActa);

router.post('/acta',
    /* decryptPayload, */[
        chkToken(),
        dataToken,
        body('clave_colonia').exists().notEmpty().isString(),
        body('num_mro').exists().notEmpty().isNumeric(),
        body('tipo_mro').exists().notEmpty().isNumeric(),
        body('levantada_distrito').exists().notEmpty().isBoolean(),
        body('forzar').exists().notEmpty().isBoolean(),
        body('coordinador_sino').exists().notEmpty().isBoolean(),
        body('num_integrantes').exists(),
        body('observador_sino').exists().notEmpty().isBoolean(),
        body('bol_recibidas').exists().notEmpty().isNumeric(),
        body('bol_adicionales').exists().notEmpty().isNumeric(),
        body('bol_sobrantes').exists().notEmpty().isNumeric(),
        body('total_ciudadanos').exists().notEmpty().isNumeric(),
        body('bol_nulas').exists().notEmpty().isNumeric(),
        body('bol_total_emitidas').exists().notEmpty().isNumeric(),
        body('opi_total_sei').exists().notEmpty().isNumeric(),
        body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
        body('integraciones').exists().notEmpty().isArray({ min: 1 }).withMessage('El minimo de objetos es 1. Debe de contener num_proyecto y votos'),
        body('integraciones.*.secuencial').exists().notEmpty().trim().matches(/^[A-Za-z]+$|^[0-9]+$/).withMessage('Solo letras o números son permitidos'),
        body('integraciones.*.votos').exists().notEmpty().isNumeric(),
        Validator
    ], RegistrarActa);

router.put('/acta',
    /* decryptPayload, */[
        chkToken([1]),
        dataToken,
        body('id_acta').exists().notEmpty().isNumeric(),
        body('levantada_distrito').exists().notEmpty().isBoolean(),
        body('forzar').exists().notEmpty().isBoolean(),
        body('coordinador_sino').exists().notEmpty().isBoolean(),
        body('num_integrantes').exists(),
        body('observador_sino').exists().notEmpty().isBoolean(),
        body('bol_recibidas').exists().notEmpty().isNumeric(),
        body('bol_adicionales').exists().notEmpty().isNumeric(),
        body('bol_sobrantes').exists().notEmpty().isNumeric(),
        body('total_ciudadanos').exists().notEmpty().isNumeric(),
        body('bol_nulas').exists().notEmpty().isNumeric(),
        body('bol_total_emitidas').exists().notEmpty().isNumeric(),
        body('opi_total_sei').exists().notEmpty().isNumeric(),
        body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
        body('integraciones').exists().notEmpty().isArray({ min: 1 }).withMessage('El minimo de objetos es 1. Debe de contener num_proyecto y votos'),
        body('integraciones.*.secuencial').exists().notEmpty().trim().matches(/^[A-Za-z]+$|^[0-9]+$/).withMessage('Solo letras o números son permitidos'),
        body('integraciones.*.votos').exists().notEmpty().isNumeric(),
        Validator
    ], ActualizarActa);

//? Cierre de Validación

router.get('/cierreValidacion', [
    chkToken(),
    dataToken
], DatosCierreValidacion);

router.post('/cierreValidacion', [
    chkToken(),
    dataToken,
    body('MSPEN').exists().notEmpty().isNumeric(),
    body('COPACO').exists().notEmpty().isNumeric(),
    // body('personasCandidatas').exists().notEmpty().isNumeric(),
    body('personasObservadoras').exists().notEmpty().isNumeric(),
    body('presentaronProyecto').exists().notEmpty().isNumeric(),
    body('mediosComunicacion').exists().notEmpty().isNumeric(),
    body('otros').exists().notEmpty().isNumeric(),
    body('total').exists().notEmpty().isNumeric(),
    body('fecha').exists().notEmpty().isDate(),
    body('hora').exists().notEmpty().isTime(),
    body('observaciones').exists().notEmpty().isString(),
    Validator
], GuardarActualizarCierreValidacion);

export { router as distSeguimientoRouter };

