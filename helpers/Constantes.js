export const iecmLogo = './resources/Emblema IECM byn.png';
export const iecmLogoBN = './resources/Emblema IECM Calado Blanco.png';
export const emblemaEC = './resources/Emblema Enchula tu colonia byn.png';

export const anioN = {
    2: 2026,
    3: 2027
}; //! ACTUALIZARLO SI ES NECESARIO

export const autor = 'SICOVACC';

export const titulos = {
    0: 'DIRECCIÓN EJECUTIVA DE ORGANIZACIÓN ELECTORAL Y GEOESTADÍSTICA',
    1: 'SISTEMA DE CÓMPUTO Y VALIDACIONES PARA LAS CONSULTAS CIUDADANAS (SICOVACC)'
}

export const plantillas = {
    0: './plantillas/',
    1: './plantillas/eleccion/',
    2: './plantillas/consulta/'
}

export const alignment = {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true
};

export const border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    right: { style: 'thin' },
    bottom: { style: 'thin' }
};

export const tituloStyle = {
    font: {
        name: 'Arial',
        bold: true
    },
    alignment
}

export const contenidoStyle = {
    font: {
        name: 'Arial',
        size: 11,
    },
    alignment,
    border
};

export const fill = {
    font: {
        name: 'Arial',
        size: 11,
        bold: true
    },
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'C0C0C0' }
    },
    border,
    alignment
};

export const aniosCAT = {
    0: {
        1: 'estatus_copaco',
        2: 'estatus_cc1',
        3: 'estatus_cc2'
    },
    1: {
        1: '',
        2: '',
        3: ''
    }
}

export const TipoMesa = tipo => {
    switch (tipo) {
        case 3: return 'MECPEP';
        case 4: return 'MECPPP';
        default: return '';
    }
}