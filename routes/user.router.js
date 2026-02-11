import { Router } from 'express';
import { body } from 'express-validator';
import { Login, LoginTitular, Verify } from '../controllers/user.controller.js';
import { chkToken, dataToken } from '../middlewares/auth-middleware.js';
import { Validator } from '../validators/validator.js';
import { decryptPayload } from '../middlewares/encrypt-middleware.js';

const router = Router();

router.post('/login',
    decryptPayload, [
    body('usuario').exists().notEmpty().isLength({ min: 4, max: 20 }).isString(),
    body('contrasena').exists().notEmpty().isString(),
    Validator
], Login);

router.post('/loginTitular',
    decryptPayload, [
    chkToken(),
    dataToken,
    body('usuario').exists().notEmpty().isLength({ min: 4, max: 20 }).isString(),
    body('contrasena').exists().notEmpty().isString(),
    Validator
], LoginTitular);

router.get('/verify', [
    chkToken(),
    dataToken
], Verify);

export { router as userRoutes };