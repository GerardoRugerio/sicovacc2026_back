import { request, response } from 'express';
import { Agent } from 'undici';
import { Audit } from '../helpers/Audit.js';
import { anioN, aniosCAT, Letras } from '../helpers/Constantes.js';
import { Comillas } from '../helpers/Funciones.js';
import { VTA_RES_PROYECTOS_2 } from '../models/VTA_RES_PROYECTOS_2.model.js';
import { SICOVACC } from '../models/consulta_usuarios_sicovacc.model.js';
import { ConsultaTipoEleccion } from '../helpers/Consultas.js';

export const ImportarVotosSEI = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito: distrito } = req.data;
    const { id_distrito, anio } = req.body;
    try {
        const sei = (await VTA_RES_PROYECTOS_2.sequelize.query(`;WITH Candidatos AS (
            SELECT VS.id_distrito, VS.id_delegacion, VS.clave_colonia, VS.ID_MESA AS num_mro, VS.secuencial, COALESCE(VS.TOTAL, 0) AS votos, COALESCE(NS.TOTAL_NULOS, 0) AS bol_nulas,1 AS anio
            FROM VTA_RES_FORMULAS VS
            LEFT JOIN VTA_NULOS_P_C2 NS ON VS.ID_DISTRITO = NS.ID_DISTRITO AND VS.CLAVE_COLONIA = NS.CLAVE_COLONIA AND VS.ID_MESA = NS.ID_MESA
        ),
        Proyectos AS (
            SELECT VS.id_distrito, VS.id_delegacion, VS.clave_colonia, VS.ID_MESA AS num_mro, CAST(VS.NUM_PROYECTO AS VARCHAR) AS secuencial, COALESCE(VS.TOTAL, 0) AS votos, COALESCE(NS.TOTAL_NULOS, 0) AS bol_nulas,
            CASE VS.ANIO WHEN 2026 THEN 2 ELSE 3 END AS anio
            FROM VTA_RES_PROYECTOS_2 VS
            LEFT JOIN VTA_NULOS_P NS ON VS.ID_DISTRITO = NS.ID_DISTRITO AND VS.CLAVE_COLONIA = NS.CLAVE_COLONIA AND VS.ID_MESA = NS.ID_MESA
        )
        SELECT *
        FROM (
        	SELECT * FROM Candidatos
            UNION ALL
        	SELECT * FROM Proyectos
        ) AS X
        ${!isNaN(id_distrito) ? `WHERE X.ID_DISTRITO = ${id_distrito}${anio != 0 ? ` AND X.ANIO = ${anio}` : ''}` : `${anio != 0 ? `WHERE X.ANIO = ${anio}` : ''}`}`))[0];
        if (!sei.length)
            return res.status(404).json({
                success: false,
                msg: `No se encontraron votos SEI`
            });
        await SICOVACC.sequelize.query(`EXEC BorrarVotosSEI ${!isNaN(id_distrito) ? id_distrito : 0}, ${anio}`);
        let P = [];
        for (const voto of sei) {
            const { id_distrito, id_delegacion, clave_colonia, num_mro, secuencial, votos, bol_nulas, anio } = voto;
            const index = P.findIndex(p => p.id_distrito === id_distrito && p.clave_colonia === clave_colonia && p.num_mro === num_mro && p.anio === anio);
            const idx = Letras.indexOf(secuencial.trim());
            const key = anio == 1 ? `participante${idx + 1}` : `proyecto${secuencial}_votos`;
            if (index != -1) {
                P[index].votos = P[index].votos || {};
                P[index].votos[key] = (P[index].votos[key] || 0) + votos;
                P[index].votacion_total_emitida += votos;
            } else
                P.push({ id_distrito, id_delegacion, clave_colonia, num_mro, bol_nulas, votacion_total_emitida: votos + bol_nulas, votos: { [key]: votos }, anio });
        }
        for (const p of P) {
            const { id_distrito, id_delegacion, clave_colonia, num_mro, bol_nulas, votacion_total_emitida, votos, anio } = p;
            let insert = '', values = '';
            Object.entries(votos).forEach(([campo, valor]) => {
                insert += `${campo}, `;
                values += `${valor}, `;
            });
            await SICOVACC.sequelize.query(`INSERT ${anio == 1 ? 'copaco' : 'consulta'}_actas (${anio != 1 ? `anio, ` : ''}id_distrito, id_delegacion, clave_colonia, num_mro, tipo_mro, modalidad, bol_nulas, votacion_total_emitida, ${insert.substring(0, insert.length - 2)}, fecha_alta, estatus)
            VALUES (${anio != 1 ? `${anio}, ` : ''}${id_distrito}, ${id_delegacion}, '${clave_colonia}', '${num_mro}', 1, 2, ${bol_nulas}, ${votacion_total_emitida}, ${values.substring(0, values.length - 2)}, CURRENT_TIMESTAMP, 1)`);
        }
        await Audit(id_transaccion, id_usuario, distrito, `IMPORTÓ VOTOS SEI ${isNaN(id_distrito) ? 'DE TODOS LOS DISTRITOS' : `DEL DISTRITO ${id_distrito}`}, ${anio == 0 ? 'DE TODOS LOS TIPOS DE ELECCIÓN' : `DE LA ELECCIÓN DE ${(await ConsultaTipoEleccion(anio)).toUpperCase()}`}`);
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
        const agent = new Agent({ connect: { rejectUnauthorized: false } });
        const myHeaders = new Headers();
        //? Configuración del header con la API
        myHeaders.append('X-API-KEY', 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiYWRtaW4iOnRydWUsImlhdCI6MTUxNjIzOTAyMn0.KMUFsIDTnFmyG3nMiGM6H9FNFUROf3wh7SmqJp-QV30');
        const requestOptions = {
            method: 'GET',
            headers: myHeaders,
            dispatcher: agent
        };
        //? Se hace la peticion al WebService y se transforma a JSON
        proyectos = await fetch('https://aplicaciones2.iecm.mx/siproe-aleatorio-2026-2027/api/reportdata/exportar', requestOptions).then(response => response.json());
        // await SICOVACC.sequelize.query(`EXEC BorrarProyectos ${id_distrito == 'TODOS' ? 0 : id_distrito}`);
        if (id_distrito != 'TODOS')
            proyectos = proyectos.filter(proyecto => proyecto.distrito == id_distrito);
        for (const proyecto of proyectos) {
            const { id, distrito, id_demarcacion, ut, numero_aleatorio, nombre, destino_recursos, ut_mejoramiento, ut_infraestructura, ut_obras, ut_servicios, ut_act_dep, ut_act_rec, ut_act_cul,
                uh_mejoramiento, uh_mantenimiento, uh_obras, uh_reparaciones, uh_servicios, uh_act_dep, uh_act_rec, uh_act_cul, ciudadano_proponente, presupuesto_aut, destRecursos, destRecursosMedia, destRecursosFinal, pe_toda,
                pe_per_may, pe_nna, pe_jovenes, pe_mujeres, pe_hombres, pe_otra, pe_desc_otra, pe_per_disc, tipo_ubicacion, calles, num_ext, fecha_alta, descripcion, folio, sorteo, ejer_fis
            } = proyecto;
            // console.log(`INSERT consulta_prelacion_proyectos (id_prelacion, id_distrito, id_delegacion, clave_colonia, num_proyecto, nom_proyecto, destino_recursos, destino_recursos_media, destino_recursos_final, tipo_rubro, rubro1, rubro2, rubro3, rubro4, rubro5, rubro6,
            // rubro7, rubro8,propuesto_por, ciudadano_presenta, costo_aproximado, poblacion_benef, pob1, pob2, pob3, pob4, pob5, pob6, pob7, pob8, ubicacion_exacta, fecha_presenta, opinion_favorable, descripcion, folio_proy_web, id_sorteo, anio, fecha_alta, id_usuario, estatus) 
            // VALUES (${id}, ${distrito}, ${id_demarcacion}, '${ut}', ${numero_aleatorio}, UPPER('${Comillas(nombre)}'), UPPER('${destRecursos}'), UPPER('${destRecursosMedia}'), UPPER('${destRecursosFinal}'), ${destino_recursos}, ${destino_recursos == 1 ? ut_mejoramiento : uh_mejoramiento},
            // ${destino_recursos == 1 ? ut_infraestructura : uh_mantenimiento}, ${destino_recursos == 1 ? ut_obras : uh_obras}, ${destino_recursos == 1 ? ut_servicios : uh_reparaciones}, ${destino_recursos == 1 ? ut_act_dep : uh_servicios}, ${destino_recursos == 1 ? ut_act_rec : uh_act_dep},
            // ${destino_recursos == 1 ? ut_act_cul : uh_act_rec}, ${destino_recursos == 1 ? 0 : uh_act_cul}, UPPER('${ciudadano_proponente}'), UPPER('${ciudadano_proponente}'), '${presupuesto_aut}', NULL, ${pe_toda}, ${pe_per_may}, ${pe_per_disc}, ${pe_nna}, ${pe_jovenes}, ${pe_mujeres},
            // ${pe_hombres}, ${pe_otra ? `UPPER('${pe_desc_otra}')` : 'NULL'}, ${tipo_ubicacion == 1 ? "'TODA LA UT'" : `UPPER('${calles}, ${num_ext}')`}, '${fecha_alta}', 'SI', UPPER('${Comillas(descripcion)}'), '${folio}', ${sorteo}, ${Object.keys(anioN).find(key => anioN[key] === ejer_fis)}, CURRENT_TIMESTAMP, ${id_usuario}, 1)`);
            const existe = (await SICOVACC.sequelize.query(`SELECT * FROM consulta_prelacion_proyectos WHERE id_prelacion = ${id} AND id_distrito = ${distrito} AND id_delegacion = ${id_demarcacion} AND clave_colonia = '${ut}'`))[0];
            if (existe == 0)
                await SICOVACC.sequelize.query(`INSERT consulta_prelacion_proyectos (id_prelacion, id_distrito, id_delegacion, clave_colonia, num_proyecto, nom_proyecto, tipo_rubro, rubro1, rubro2, rubro3, rubro4, rubro5, rubro6, rubro7, rubro8, propuesto_por, ciudadano_presenta,
                costo_aproximado, poblacion_benef, pob1, pob2, pob3, pob4, pob5, pob6, pob7, pob8, ubicacion_exacta, fecha_presenta, opinion_favorable, descripcion, folio_proy_web, id_sorteo, anio, fecha_alta, id_usuario, estatus) VALUES (${id}, ${distrito}, ${id_demarcacion}, '${ut}', ${numero_aleatorio}, UPPER('${Comillas(nombre)}'),
                ${destino_recursos}, ${destino_recursos == 1 ? ut_mejoramiento : uh_mejoramiento}, ${destino_recursos == 1 ? ut_infraestructura : uh_mantenimiento}, ${destino_recursos == 1 ? ut_obras : uh_obras}, ${destino_recursos == 1 ? ut_servicios : uh_reparaciones}, ${destino_recursos == 1 ? ut_act_dep : uh_servicios},
                ${destino_recursos == 1 ? ut_act_rec : uh_act_dep}, ${destino_recursos == 1 ? ut_act_cul : uh_act_rec}, ${destino_recursos == 1 ? 0 : uh_act_cul}, UPPER('${ciudadano_proponente}'), UPPER('${ciudadano_proponente}'), '${presupuesto_aut}', NULL, ${pe_toda}, ${pe_per_may}, ${pe_per_disc}, ${pe_nna},
                ${pe_jovenes}, ${pe_mujeres}, ${pe_hombres}, ${pe_otra ? `UPPER('${pe_desc_otra}')` : 'NULL'}, ${tipo_ubicacion == 1 ? "'TODA LA UT'" : `UPPER('${calles}, ${num_ext}')`}, '${fecha_alta}', 'SI', UPPER('${Comillas(descripcion)}'), '${folio}', ${sorteo},
                ${Object.keys(anioN).find(key => anioN[key] === ejer_fis)}, CURRENT_TIMESTAMP, ${id_usuario}, 1)`);
        }
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

export const ImportarParticipantesAprobados = async (req = request, res = response) => {
    res.json({
        success: false,
        msg: 'En preparación me encuentro'
    });
}

export const DatosProyectos = async (req = request, res = response) => {
    const { id_distrito, clave_colonia, anio } = req.body;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT id_proyecto, UPPER(D.nombre_delegacion) AS nombre_delegacion, secuencial, nom_proyecto, folio
        FROM consulta_prelacion_proyectos_VVS P
        LEFT JOIN consulta_cat_delegacion D ON P.id_delegacion = D.id_delegacion
        WHERE id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}' AND anio = ${anio}
        ORDER BY secuencial ASC`))[0];
        if (!datos.length)
            return res.status(404).json({
                success: false,
                msg: 'No se han encontrado proyectos'
            });
        res.json({
            success: true,
            datos
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

export const DatosParticipantes = async (req = request, res = response) => {
    const { id_distrito, clave_colonia } = req.body;
    try {
        const datos = (await SICOVACC.sequelize.query(`SELECT idFormulas, UPPER(D.nombre_delegacion) AS nombre_delegacion, secuencial, CONCAT(nombre, ' ', paterno, ' ', materno) AS nombre, folio
        FROM copaco_formulas F
        LEFT JOIN consulta_cat_delegacion D ON F.id_delegacion = D.id_delegacion
        WHERE estatus = 1 AND secuencial IS NOT NULL AND id_distrito = ${id_distrito} AND clave_colonia = '${clave_colonia}'
        ORDER BY secuencial ASC`))[0];
        if (!datos.length)
            return res.status(404).json({
                success: false,
                msg: 'No se han encontrado participantes'
            });
        res.json({
            success: true,
            datos
        })
    } catch (err) {
        console.error(`Error en DatosParticipantes: ${err}`);
        res.status(500).json({
            success: false,
            msg: 'Error desconocido'
        });
    }
}

export const EliminarParticipante = async (req = request, res = response) => {
    const { id_transaccion, id_usuario, id_distrito } = req.data;
    const { idFormulas } = req.body;
    try {
        const resp = SICOVACC.sequelize.query(`UPDATE copaco_formulas SET estatus = 0 WHERE idFormulas = ${idFormulas}`);
        if (resp[1] == 0)
            return res.status(404).json({
                success: false,
                msg: 'Participante no encontrado'
            });
        await Audit(id_transaccion, id_usuario, id_distrito, `ELIMINÓ EL PARTICIPANTE ${idFormulas}`);
        res.json({
            success: true,
            msg: 'Participante eliminado'
        });
    } catch (err) {
        console.error(`Error en EliminarParticipante: ${err}`);
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