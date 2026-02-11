import { request, response } from 'express';
import { validationResult } from 'express-validator';

export const Validator = (req = request, res = response, next) => {
    const errores = validationResult(req);
    if (!errores.isEmpty()) {
        let array = [];
        Object.keys(errores.mapped()).forEach(key => array.push(`${key} (${errores.mapped()[key].value == undefined ? 'Es obligatorio' : errores.mapped()[key].msg.match('Invalid value') ? 'Valor invalido' : errores.mapped()[key].msg})`));
        return res.status(400).json({
            success: false,
            msg: Error(array)
        });
    }
    next();
}

const Error = (Array) => {
    if (Array.length >= 3) {
        const cant = Array.length - 1;
        let res = '';
        for (let i = 0; i < cant; i++)
            res += `${Array[i]}, `
        return `Error en los campos: ${res.substring(0, res.length - 2)} y ${Array[Array.length - 1]}`;
    } else
        return `${Array.length == 1 ? 'Error en el campo:' : 'Error en los campos:'} ${Array.join(' y ')}`
};