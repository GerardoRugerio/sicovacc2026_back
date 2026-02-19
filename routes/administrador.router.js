import { Router } from 'express';
import { body } from 'express-validator';
import { DatosParticipantes, DatosProyectos, EliminarParticipante, EliminarProyecto, ImportarParticipantesAprobados, ImportarProyectosAprobados, ImportarVotosSEI, ListaUsuarios } from '../controllers/administrador.controller.js';
import { chkToken, dataToken } from '../middlewares/auth-middleware.js';
import { chkDistrito } from '../middlewares/dist-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

router.post('/importarVotosSEI', [
    chkToken([99]),
    dataToken,
    body('id_distrito').exists().notEmpty().isNumeric(),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator,
    chkDistrito(undefined, true)
], ImportarVotosSEI);

router.post('/importarProyectos', [
    chkToken([99]),
    dataToken,
    body('id_distrito').exists().notEmpty().isNumeric(),
    Validator,
    chkDistrito(undefined, true)
], ImportarProyectosAprobados);

router.post('/importarParticipantes', [
    chkToken([99]),
    dataToken,
    body('id_distrito').exists().notEmpty().isNumeric(),
    Validator,
    chkDistrito(undefined, true)
], ImportarParticipantesAprobados);

router.post('/eliminarProyectos', [
    chkToken([99]),
    body('id_distrito').exists().notEmpty().isNumeric(),
    body('clave_colonia').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 2, max: 3 }).withMessage('El valor debe de ser 2 al 3'),
    Validator,
    chkDistrito()
], DatosProyectos);

router.delete('/eliminarProyectos', [
    chkToken([99]),
    dataToken,
    body('id_proyecto').exists().notEmpty().isNumeric(),
    Validator
], EliminarProyecto);

router.post('/eliminarParticipantes', [
    chkToken([99]),
    body('id_distrito').exists().notEmpty().isNumeric(),
    body('clave_colonia').exists().notEmpty().isString(),
    Validator
], DatosParticipantes);

router.delete('/eliminarParticipantes', [
    chkToken([99]),
    dataToken,
    body('idFormulas').exists().notEmpty().isNumeric(),
    Validator
], EliminarParticipante);

router.get('/listaUsuarios', [
    chkToken([99]),
    dataToken
], ListaUsuarios);

export { router as administradorRouter };

