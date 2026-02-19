import { Router } from 'express';
import { body } from 'express-validator';
import { Colonias, ColoniasConActas, ColoniasSinActas, ColoniasValidadas, ColoniasVotacion, Delegacion, Mesas, MesasConActas, MesaSinActas, TipoEleccion } from '../controllers/catalogos.controller.js';
import { chkToken, dataToken } from '../middlewares/auth-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

router.get('/tipoEleccion', [
    chkToken()
], TipoEleccion);

router.post('/colonias', [
    chkToken(),
    dataToken,
    body('id_distrito').custom((value, { req }) => {
        const { perfil } = req.data;
        if (![1, 2].includes(perfil)) {
            if (!value)
                throw new Error('Es obligatorio');
            if (isNaN(value))
                throw new Error('Valor invalido');
        }
        return true;
    }),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], Colonias);

router.post('/coloniasConActas', [
    chkToken(),
    dataToken,
    body('id_distrito').custom((value, { req }) => {
        const { perfil } = req.data;
        if (![1, 2].includes(perfil)) {
            if (!value)
                throw new Error('Es obligatorio');
            if (isNaN(value))
                throw new Error('Valor invalido');
        }
        return true;
    }),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], ColoniasConActas);

router.post('/coloniasSinActas', [
    chkToken(),
    dataToken,
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], ColoniasSinActas);

router.post('/coloniasValidadas', [
    chkToken(),
    dataToken,
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], ColoniasValidadas);

router.post('/delegacion', [
    chkToken(),
    dataToken,
    body('clave_colonia').exists().notEmpty().isString(),
    Validator
], Delegacion);

router.post('/mesas', [
    chkToken(),
    dataToken,
    body('clave_colonia').exists().notEmpty().isString(),
    Validator
], Mesas);

router.post('/mesasConActas', [
    chkToken(),
    body('id_distrito').exists().notEmpty().isNumeric(),
    body('clave_colonia').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], MesasConActas);

router.post('/mesasSinActas', [
    chkToken(),
    dataToken,
    body('clave_colonia').exists().notEmpty().isString(),
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], MesaSinActas);

router.post('/coloniasVotacion', [
    chkToken(),
    dataToken,
    body('anio').exists().notEmpty().isInt({ min: 1, max: 3 }).withMessage('El valor debe de ser 1 al 3'),
    Validator
], ColoniasVotacion);

export { router as catRouter };