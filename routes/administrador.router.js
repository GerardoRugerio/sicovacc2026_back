import { Router } from 'express';
import { body } from 'express-validator';
import { DatosProyectos, EliminarProyecto, ImportarProyectosAprobados, ImportarVotosSEI, ListaUsuarios } from '../controllers/administrador.controller.js';
import { chkToken, dataToken } from '../middlewares/auth-middleware.js';
import { chkDistrito } from '../middlewares/dist-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

router.post('/importarVotosSEI', [
    chkToken([99]),
    dataToken,
    body('id_distrito').exists().notEmpty(),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator,
    chkDistrito(undefined, true)
], ImportarVotosSEI);

router.post('/importarProyectos', [
    chkToken([99]),
    dataToken,
    body('id_distrito').exists().notEmpty(),
    Validator,
    chkDistrito(undefined, true)
], ImportarProyectosAprobados);

router.post('/eliminarProyectos', [
    chkToken([99]),
    body('id_distrito').exists().notEmpty().isNumeric(),
    body('clave_colonia').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator,
    chkDistrito()
], DatosProyectos);

router.delete('/eliminarProyectos', [
    chkToken([99]),
    dataToken,
    body('id_proyecto').exists().notEmpty().isNumeric(),
    Validator
], EliminarProyecto);

router.get('/listaUsuarios', [
    chkToken([99]),
    dataToken
], ListaUsuarios);

export { router as administradorRouter };

