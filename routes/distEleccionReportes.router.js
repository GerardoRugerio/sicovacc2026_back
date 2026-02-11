import { request, Router } from 'express';
import { body, param } from 'express-validator';
import { ActasAlerta, CandidaturasEmpate, ComputoTotalUT, ConcentradoParticipantes, LevantadaDistrito, MesasComputadas, MesasNoComputadas, ResultadoComputoTotalMesa, ResultadoComputoTotalUT, ResultadosMesa, UTConComputo, UTSinComputo } from '../controllers/distEleccionReportesExcel.controller.js';
import { chkDistrito, StatusReporte } from '../middlewares/dist-middleware.js';
import { Validator } from '../validators/validator.js';
import { ActaValidacionPDF } from '../controllers/distEleccionReportesPDF.controller.js';
import { ActaValidacionWord } from '../controllers/distEleccionReporteWord.controller.js';

const router = Router();

//? Cómputo Total de las Candidaturas por UT - Excel

router.get('/computoTotalUT/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], ComputoTotalUT);

//? Resultados del Cómputo Total por Mesa - Excel

router.get('/resultadoComputoTotalMesa/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], ResultadoComputoTotalMesa);

//? Resultados del Cómputo Total por Unidad Territorial - Excel

router.get('/resultadoComputoTotalUT/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], ResultadoComputoTotalUT);

//? Concentrado de Candidaturas Participantes - Excel

router.get('/concentradoParticipantes/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], ConcentradoParticipantes);

//? Candidaturas en las que se presenta empate - Excel

router.get('/candidaturasEmpate/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], CandidaturasEmpate);

//? Resultados de Votos por Mesa - Excel

router.get('/resultadosMesa/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], ResultadosMesa);

//? Concentrado de Mesas Computadas - Excel

router.get('/MesasComputadas/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], MesasComputadas);

//? Concentrado de Mesas que no han sido Computadas - Excel

router.get('/MesasNoComputadas/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], MesasNoComputadas);

//? Unidades Territoriales Con Cómputo Capturado - Excel

router.get('/UTConComputo/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], UTConComputo);

//? Unidades Territoriales Sin Cómputo Capturado - Excel

router.get('/UTSinComputo/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], UTSinComputo);

//? Actas Levantadas en Dirección Distrital - Excel

router.get('/levantadaDistrito/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], LevantadaDistrito);

//? Actas Capturadas con Alertas - Excel

router.get('/actasAlerta/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    Validator
], ActasAlerta);

//? Acta de Validación - PDF - Word

router.post('/actaValidacion/:id_distrito', [
    param('id_distrito').exists().notEmpty(),
    chkDistrito('1'),
    body('clave_colonia').exists().notEmpty().isString(),
    body('tipo').exists().notEmpty().isString().isIn(['PDF', 'WORD']).withMessage(`El tipo debe de ser 'PDF' o 'WORD'`),
    Validator,
    StatusReporte
], async (req = request, res = response) => {
    const { tipo } = req.body;
    if (tipo.toLowerCase() == 'pdf')
        ActaValidacionPDF(req, res);
    else
        ActaValidacionWord(req, res);
});

export { router as distEleccionReportesRouter };

