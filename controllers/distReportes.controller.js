import { request, response } from 'express';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Consulta de Proyectos

export const ListaProyectos = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { clave_colonia, anio } = req.body;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT secuencial AS num_proyecto, fecha, costo_aproximado, folio, rubro_general, nom_proyecto, ciudadano_presenta, poblacion_benef, ubicacion_exacta, descripcion
        FROM consulta_prelacion_proyectos_VVS
        WHERE id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}' AND anio = ${anio}
        ORDER BY num_proyecto, folio ASC`))[0];
        if (!datos.length)
            return res.status(404).json({
                success: false,
                msg: 'No se encotnro ningún proyecto'
            });
        res.json({
            success: true,
            datos
        });
    } catch (err) {
        console.error(`Error en ListaProyectos: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

//? Consulta de Fórmulas

export const ListaFormulas = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { clave_colonia } = req.body;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT dbo.NumeroALetras(secuencial) AS secuencial, UPPER(CONCAT(nombre, ' ', paterno, ' ', materno)) AS nombre, edad, CASE genero WHEN 'F' THEN 'FEMENINO' ELSE 'MASCULINO' END AS genero, UPPER(cargo) AS cargo, folio
        FROM copaco_formulas F
        WHERE F.secuencial IS NOT NULL AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
        ORDER BY F.secuencial ASC`))[0];
        if (!datos.length)
            return res.status(404).json({
                success: false,
                msg: 'No se encotnro ningún candidato'
            });
        res.json({
            success: true,
            datos
            // datos: EncryptData(acta)
        })
    } catch (err) {
        console.error(`Error en ListaFormulas: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}