import { request, response } from 'express';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

//? Consulta de Proyectos

export const ListaProyectos = async (req = request, res = response) => {
    const { id_distrito } = req.data;
    const { clave_colonia, anio } = req.body;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT num_proyecto, fecha_presenta AS fecha, FORMAT(CAST(costo_aproximado AS INT), 'C', 'en-US') AS costo_aproximado, UPPER(folio_proy_web) AS folio_proy_web,
        UPPER(STUFF((
            SELECT ', ' + rubro
            FROM (VALUES
                (CASE WHEN rubro1 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Mejoramiento de espacios públicos' ELSE 'Mejoramiento' END ELSE NULL END),
                (CASE WHEN rubro2 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Equipamiento e infraestructura urbana' ELSE 'Mantenimiento' END ELSE NULL END),
                (CASE WHEN rubro3 = 1 THEN 'Obras' ELSE NULL END),
                (CASE WHEN rubro4 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Servicios' ELSE 'Reparaciones en áreas y bienes de uso común' END ELSE NULL END),
                (CASE WHEN rubro5 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Actividades deportivas' ELSE 'Servicios' END ELSE NULL END),
                (CASE WHEN rubro6 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Actividades recreativas' ELSE 'Actividades deportivas' END ELSE NULL END),
                (CASE WHEN rubro7 = 1 THEN CASE WHEN tipo_rubro = 1 THEN 'Actividades culturales' ELSE 'Actividades recreativas' END ELSE NULL END),
                (CASE WHEN rubro8 = 1 THEN CASE WHEN tipo_rubro = 1 THEN NULL ELSE 'Actividades culturales' END ELSE NULL END)
            ) AS sub(rubro)
            WHERE rubro IS NOT NULL
            FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
        ) AS rubro_general,
        UPPER(nom_proyecto) AS nom_proyecto, UPPER(ciudadano_presenta) AS ciudadano_presenta,
        UPPER(STUFF((
            SELECT ', ' + pob
            FROM (VALUES
                (CASE WHEN pob1 = 1 THEN 'Toda la población' ELSE NULL END),
                (CASE WHEN pob2 = 1 THEN 'Personas mayores (60 años o más)' ELSE NULL END),
                (CASE WHEN pob3 = 1 THEN 'Personas con discapacidad' ELSE NULL END),
                (CASE WHEN pob4 = 1 THEN 'Infancias y adolescencias (menores de 18 años)' ELSE NULL END),
                (CASE WHEN pob5 = 1 THEN 'Jóvenes' ELSE NULL END),
                (CASE WHEN pob6 = 1 THEN 'Mujeres' ELSE NULL END),
                (CASE WHEN pob7 = 1 THEN 'Hombres' ELSE NULL END),
                (CASE WHEN pob8 IS NOT NULL THEN pob8 ELSE NULL END)
            ) AS sub(pob)
            WHERE pob IS NOT NULL
            FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
        ) AS poblacion_benef,
        UPPER(ubicacion_exacta) AS ubicacion_exacta, UPPER(descripcion) AS descripcion
        FROM consulta_prelacion_proyectos
        WHERE anio = ${anio} AND estatus = 1 AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
        ORDER BY num_proyecto, folio_proy_web ASC`))[0];
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