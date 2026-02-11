import { Router } from 'express';
import { param } from 'express-validator';
import { ActasAlerta, CandidaturasEmpate, ComputoTotalUT, ConcentradoParticipantes, LevantadaDistrito, MesasComputadas, MesasNoComputadas, Participacion, ResultadoComputoTotalMesa, ResultadoComputoTotalUT, ResultadosMesa, UTConComputo, UTConComputoGA, UTSinComputo, VotacionDemarcacion, VotacionDistrito } from '../controllers/centEleccionReportesExcel.controller.js';
import { chkDistrito } from '../middlewares/dist-middleware.js';
import { Validator } from '../validators/validator.js';

const router = Router();

//? Cómputo total de las Candidaturas por UT - Excel

router.get('/computoTotalUT/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], ComputoTotalUT);

//? Resultados del Cómputo Total por Mesa - Excel

router.get('/resultadoComputoTotalMesa/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], ResultadoComputoTotalMesa);

//? Resultados del Cómputo Total por Unidad Territorial - Excel

router.get('/resultadoComputoTotalUT/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], ResultadoComputoTotalUT);

//? Concentrado de Candidaturas Participantes - Excel

router.get('/concentradoParticipantes/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], ConcentradoParticipantes);

//? Candidaturas en las que se presenta empate - Excel

router.get('/candidaturasEmpate/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], CandidaturasEmpate);

//? Resultados de Votos por Mesa - Excel

router.get('/resultadosMesa/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], ResultadosMesa);

//? Concentrado de Mesas Computadas - Excel

router.get('/MesasComputadas/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], MesasComputadas);

//? Concentrado de Mesas que no han sido Computadas - Excel

router.get('/MesasNoComputadas/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], MesasNoComputadas);

//? Unidades Territoriales Con Cómputo Capturado - Excel

router.get('/UTConComputo/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], UTConComputo);

//? Unidades Territoriales Sin Cómputo Capturado - Excel

router.get('/UTSinComputo/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], UTSinComputo);

//? Concentrado de Unidades Territoriales por Distrito Electoral con Cómputo Capturado (Grado de Avance) - Excel

router.get('/UTConComputoGA', UTConComputoGA);

//? Actas Levantadas en Dirección Distrital - Excel

router.get('/levantadaDistrito/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], LevantadaDistrito);

//? Votación Total por Distrito - Excel

router.get('/votacionDistrito', VotacionDistrito);

//? Votación Total por Demarcación - Excel

router.get('/votacionDemarcacion', VotacionDemarcacion);

//? Porcentaje Participación por Distrito - Excel

router.get('/participacion', Participacion);

//? Actas Capturadas con Alertas - Excel

router.get('/actasAlerta/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('0'),
    Validator
], ActasAlerta);

export { router as centEleccionReportesRouter };

