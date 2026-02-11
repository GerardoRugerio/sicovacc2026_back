import { request, response } from 'express';
import { Audit } from '../helpers/Audit.js';
import { aniosCAT } from '../helpers/Constantes.js';
import { Comillas } from '../helpers/Funciones.js';
import { VTA_RES_PROYECTOS_2 } from '../models/VTA_RES_PROYECTOS_2.model.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';

export const ImportarVotosSEI = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito: distrito } = req.data;
    const { id_distrito, anio } = req.body;
    try {
        //? Se trae todos los votos desde SEI del distrito seleccionado, si no se le manda distrito trae todos
        const votos = await VTA_RES_PROYECTOS_2.sequelize.query(`SELECT VS.id_distrito, VS.id_delegacion, VS.clave_colonia, CAST(RIGHT(VS.id_mesa, 1) AS INT) AS id_mesa, VS.num_proyecto, COALESCE(VS.total, 0) AS votos,
        COALESCE(NS.total_nulos, 0) AS bol_nulas, COALESCE(CS.total_computados, 0) AS votacion_total_emitida
        FROM ${aniosCAT[1][anio]} VS
        LEFT JOIN VTA_NULOS_P NS ON VS.id_distrito = NS.id_distrito AND VS.clave_colonia = NS.clave_colonia AND CAST(RIGHT(VS.id_mesa, 1) AS INT) = NS.id_mesa
        LEFT JOIN VTA_COMPUTADOS_P CS ON VS.id_distrito = CS.id_distrito AND VS.clave_colonia = CS.clave_colonia AND CAST(RIGHT(VS.id_mesa, 1) AS INT) = CS.id_mesa
        ${!isNaN(id_distrito) ? `WHERE VS.id_distrito = ${id_distrito}` : ''}`);
        if (votos[1] == 0)
            return res.status(404).json({
                success: false,
                msg: `No se encontraron votos SEI${!isNaN(id_distrito) ? ` para el distrito ${id_distrito}` : ''}`
            });
        await SICOVACC.sequelize.query(`EXEC BorrarVotosSEI ${id_distrito == 'TODOS' ? 0 : id_distrito}`);
        let proyectos = [];
        //? Se formatea para que los votos puedan estar en una sola linea
        for (let voto of votos[0]) {
            const { id_distrito, id_delegacion, clave_colonia, id_mesa: num_mro, num_proyecto: num, votos, bol_nulas, votacion_total_emitida } = voto;
            const index = proyectos.findIndex(proyecto => proyecto.id_distrito == id_distrito && proyecto.id_delegacion == id_delegacion && proyecto.clave_colonia == clave_colonia && proyecto.num_mro == num_mro);
            if (index != -1)
                proyectos[index] = { ...proyectos[index], [`proyecto${num}_votos`]: votos };
            else
                proyectos.push({ id_distrito, id_delegacion, clave_colonia, num_mro, bol_nulas, votacion_total_emitida, [`proyecto${num}_votos`]: votos });
        }
        for (let proyecto of proyectos) {
            const { id_distrito, id_delegacion, clave_colonia, num_mro } = proyecto;
            let insert = '', values = '';
            Object.keys(proyecto).forEach(key => {
                insert += `${key}, `;
                values += `${key == 'clave_colonia' || key == 'num_mro' ? `'${proyecto[key]}'` : proyecto[key]}, `;
            });
            const existe = await SICOVACC.sequelize.query(`SELECT * FROM consulta_actas WHERE id_distrito = ${id_distrito} AND id_delegacion = ${id_delegacion} AND clave_colonia = '${clave_colonia}' AND num_mro = ${num_mro} AND modalidad = 2`);
            //? Si no existe se guarda en la base de datos
            if (existe[1] == 0)
                await SICOVACC.sequelize.query(`INSERT consulta_actas (${insert.substring(0, insert.length - 2)}, tipo_mro, modalidad, anio, fecha_alta, estatus) VALUES (${values.substring(0, values.length - 2)}, 1, 2, ${anio}, CURRENT_TIMESTAMP, 1)`);
        }
        await Audit(id_transaccion, id_usuario, distrito, `IMPORTÓ VOTOS SEI ${isNaN(id_distrito) ? 'DE TODOS LOS DISTRITO' : `DEL DISTRITO ${id_distrito}`}`);
        res.json({
            success: true,
            msg: 'Proceso Terminado Correctamente'
        });
    } catch (err) {
        console.error(`Error en ImportarVotosSEI: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const ImportarProyectosAprobados = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito: distrito } = req.data;
    const { id_distrito } = req.body;
    try {
        let proyectos = {};
        const myHeaders = new Headers();
        //? Configuración del header con la API
        myHeaders.append('X-API-KEY', 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiYWRtaW4iOnRydWUsImlhdCI6MTUxNjIzOTAyMn0.KMUFsIDTnFmyG3nMiGM6H9FNFUROf3wh7SmqJp-QM39');
        const requestOptions = {
            method: 'GET',
            headers: myHeaders,
            redirect: 'follow'
        };
        //? Se hace la peticion al WebService y se transforma a JSON
        proyectos = await fetch('https://app.iecm.mx/siproe-aleatorio2025/api/reportdata/exportar', requestOptions).then(response => response.json());
        // proyectos = await fetch('http://145.0.46.37:4000/api/reportdata/exportar', requestOptions).then(response => response.json());
        await SICOVACC.sequelize.query(`EXEC BorrarProyectos ${id_distrito == 'TODOS' ? 0 : id_distrito}`);
        if (id_distrito != 'TODOS')
            proyectos = proyectos.filter(proyecto => proyecto.distrito === id_distrito);
        for (const proyecto of proyectos) {
            const { id, distrito, id_demarcacion, ut, numero_aleatorio, nombre, destino_recursos, ut_mejoramiento, ut_infraestructura, ut_obras, ut_servicios, ut_act_dep, ut_act_rec, ut_act_cul,
                uh_mejoramiento, uh_mantenimiento, uh_obras, uh_reparaciones, uh_servicios, uh_act_dep, uh_act_rec, uh_act_cul, ciudadano_proponente, presupuesto_aut, pe_toda,
                pe_per_may, pe_nna, pe_jovenes, pe_mujeres, pe_hombres, pe_otra, pe_desc_otra, pe_per_disc, tipo_ubicacion, calles, num_ext, fecha_alta, descripcion, folio, sorteo, ejer_fis
            } = proyecto;
            const existe = (await SICOVACC.sequelize.query(`SELECT * FROM consulta_prelacion_proyectos WHERE id_prelacion = ${id} AND id_distrito = ${distrito} AND id_delegacion = ${id_demarcacion} AND clave_colonia = '${ut}'`))[0];
            if (existe == 0)
                await SICOVACC.sequelize.query(`INSERT consulta_prelacion_proyectos (id_prelacion, id_distrito, id_delegacion, clave_colonia, num_proyecto, nom_proyecto, tipo_rubro, rubro1, rubro2, rubro3, rubro4, rubro5, rubro6, rubro7, rubro8, propuesto_por, ciudadano_presenta,
                costo_aproximado, poblacion_benef, pob1, pob2, pob3, pob4, pob5, pob6, pob7, pob8, ubicacion_exacta, fecha_presenta, opinion_favorable, descripcion, folio_proy_web, id_sorteo, anio, fecha_alta, id_usuario, estatus) VALUES (${id}, ${distrito}, ${id_demarcacion}, '${ut}', ${numero_aleatorio}, UPPER('${Comillas(nombre)}'),
                ${destino_recursos}, ${destino_recursos == 1 ? ut_mejoramiento : uh_mejoramiento}, ${destino_recursos == 1 ? ut_infraestructura : uh_mantenimiento}, ${destino_recursos == 1 ? ut_obras : uh_obras}, ${destino_recursos == 1 ? ut_servicios : uh_reparaciones}, ${destino_recursos == 1 ? ut_act_dep : uh_servicios},
                ${destino_recursos == 1 ? ut_act_rec : uh_act_dep}, ${destino_recursos == 1 ? ut_act_cul : uh_act_rec}, ${destino_recursos == 1 ? 0 : uh_act_cul}, UPPER('${ciudadano_proponente}'), UPPER('${ciudadano_proponente}'), '${presupuesto_aut}', NULL, ${pe_toda}, ${pe_per_may}, ${pe_per_disc}, ${pe_nna},
                ${pe_jovenes}, ${pe_mujeres}, ${pe_hombres}, ${pe_otra ? `UPPER('${pe_desc_otra}')` : 'NULL'}, ${tipo_ubicacion == 1 ? "'TODA LA UT'" : `UPPER('${calles}, ${num_ext}')`}, '${fecha_alta}', 'SI', UPPER('${Comillas(descripcion)}'), '${folio}', ${sorteo}, ${ejer_fis}, CURRENT_TIMESTAMP, ${id_usuario}, 1)`);
        }
        await SICOVACC.sequelize.query(`UPDATE consulta_prelacion_proyectos SET nom_proyecto = UPPER('${Comillas(`"Arte en revolución y expresión: 'Revolucionarte'"`)}') WHERE clave_colonia = '14-043' AND num_proyecto = 10`);
        await Audit(id_transaccion, id_usuario, distrito, `IMPORTÓ LOS PROYECTOS DE SIPROE${id_distrito != 'TODOS' ? ` DEL DISTRITO ${id_distrito}` : ''}`);
        res.json({
            success: true,
            msg: 'Proyectos importados con exito'
        });
    } catch (err) {
        console.error(`Error en ImportarProyectosAprobados: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        })
    }
}

export const DatosProyectos = async (req = request, res = response) => {
    const { id_distrito, clave_colonia, anio } = req.body;
    try {
        const proyectos = await SICOVACC.sequelize.query(`SELECT CPP.id_proyecto, UPPER(CCD.nombre_delegacion) AS nombre_delegacion, CPP.num_proyecto, UPPER(CPP.nom_proyecto) AS nom_proyecto, UPPER(CPP.folio_proy_web) AS folio
        FROM consulta_prelacion_proyectos CPP
        LEFT JOIN consulta_cat_delegacion CCD ON CPP.id_delegacion = CCD.id_delegacion
        LEFT JOIN ${aniosCAT[anio]} CCC ON CPP.clave_colonia = CCC.clave_colonia
        WHERE CPP.estatus = 1 AND CPP.anio = ${anio} AND CPP.id_distrito = ${id_distrito} AND CPP.clave_colonia = '${clave_colonia}'
        ORDER BY CPP.num_proyecto`);
        if (proyectos[1] == 0)
            return res.status(404).json({
                success: false,
                msg: 'No se han encontrado proyectos'
            });
        res.json({
            success: true,
            datos: proyectos[0]
        });
    } catch (err) {
        console.error(`Error en DatosProyectos: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const EliminarProyecto = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { id_proyecto } = req.body;
    try {
        const resp = await SICOVACC.sequelize.query(`UPDATE consulta_prelacion_proyectos SET estatus = 0, fecha_modif = CURRENT_TIMESTAMP WHERE id_proyecto = ${id_proyecto}`);
        // const resp = await SICOVACC.sequelize.query(`DELETE FROM consulta_prelacion_proyectos WHERE id_proyecto = ${id_proyecto}`);
        if (resp[1] == 0)
            return res.status(404).json({
                success: false,
                msg: 'Proyecto no encontrado'
            });
        await Audit(id_transaccion, id_usuario, id_distrito, `ELIMINÓ EL PROYECTO ${id_proyecto}`);
        res.json({
            success: true,
            msg: 'Proyecto eliminado'
        });
    } catch (err) {
        console.error(`Error en EliminarProyecto: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const ListaUsuarios = async (req = request, res = response) => {
    const { id_usuario } = req.data;
    try {
        const usuarios = (await SICOVACC.sequelize.query(`SELECT CONCAT(U.nombre, ' ', U.ape_paterno, ' ', U.ape_materno) AS usuario, TC.tipo_cuenta
        FROM consulta_usuarios_sivacc U
        LEFT JOIN (
            SELECT 1 AS perfil, 'Titular' AS tipo_cuenta
            UNION ALL SELECT 2 AS perfil, 'Capturista' AS tipo_cuenta
            UNION ALL SELECT 3 AS perfil, 'Central' AS tipo_cuenta
            UNION ALL SELECT 4 AS perfil, 'DEOEyG' AS tipo_cuenta
            UNION ALL SELECT 99 AS perfil, 'Administrador' AS tipo_cuenta
        ) AS TC ON U.perfil = TC.perfil
        LEFT JOIN consulta_audit A ON U.id_usuario = A.id_usuario
        WHERE A.token IS NOT NULL AND A.estatus = 1 AND (A.fecha_inicio IS NOT NULL AND A.fecha_cierre IS NULL) AND U.id_usuario <> ${id_usuario}
        ORDER BY U.id_usuario`))[0];
        res.json({
            success: true,
            datos: usuarios.map(({usuario, tipo_cuenta}) => `${usuario}${!usuario.toLowerCase().match(tipo_cuenta.toLowerCase()) ? ` - ${tipo_cuenta}` : ''}`)
        })
    } catch (err) {
        console.error(`Error en ListaUsuarios: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}