import { config } from 'dotenv';
import Pako from 'pako';
import CryptoJS from 'crypto-js';

const formatear = new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
});

/**
 * Función que transforma un numero a formato de dinero (Dolares)
 * 
 * @param {string} num Número a transformar
 * @returns {string} Dinero tranformado al formato de Dolares
 */
export const Convertir = num => formatear.format(num);

/**
 * Función que transforma la cadena de texto a una con la primera en mayusculas
 * 
 * @param {string} cadena Cadena de texto a Capitalizar (Solo la primera letra en mayusculas)
 * @returns {string} Cadena capitalizada
 */
export const Capitalizar = cadena => cadena.replace(/\w\S*/g, (texto) => texto.charAt(0).toUpperCase() + texto.substring(1).toLowerCase());

/**
 * Función que transfroma la cadena de texto para poder usarla en un insert en la base de datos
 * 
 * @param {string} cadena Cadena de texto que convertira una comilla simple a una comilla simple doble
 * @returns {string} Cadena de texto modificada
 */
export const Comillas = cadena => cadena.replace(/'/g, "''");

/**
 * Función que divide un arreglo a un tamaño deseado
 * 
 * @param {object} original Arreglo que se dividira
 * @param {number} tamanio Número en que se dividira el arreglo 
 * @returns {object} Arreglo con sub arreglos
 */
export const dividirArreglo = (original, tamanio) => {
    let subArreglo = [];
    for (let i = 0; i < original.length; i += tamanio)
        subArreglo.push(original.slice(i, i + tamanio));
    return subArreglo;
}

/**
 * Función que transforma un número a texto
 * 
 * @param {number} num Número a transformar
 * @returns {string} Número en texto
 */
export const NumAText = num => {
    if (num == 0) return 'CERO';
    const uniYesp = ['', 'UNO', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE', 'DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'DIECISÉIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE', 'VEINTE', 'VEINTIUNO', 'VEINTIDÓS', 'VEINTITRÉS', 'VEINTICUATRO', 'VEINTICINCO', 'VEINTISÉIS', 'VEINTISIETE', 'VEINTIOCHO', 'VEINTINUEVE'];
    const decenas = ['', '', '', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA'];
    const centenas = ['', 'CIEN', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS'];
    //? Convierte el número en string y separa en grupos de 3 cifras
    const strNum = String(num).split('').reverse();
    let numeros = [];
    let palabra = '';
    for (let i = 0; i < strNum.length; i += 3) {
        const grupo = strNum.slice(i, i + 3).reverse().join('');
        numeros.push(Number(grupo));
    }
    //? Procesa cada grupo de 3 cifras
    for (let i = numeros.length - 1; i >= 0; i--) {
        const original = numeros[i];
        let aux = 0;
        let aux2 = '';
        //? Procesa centenas
        if (numeros[i] >= 100) {
            aux = numeros[i] / 100;
            aux2 += `${aux > 1 && aux < 2 ? `${centenas[Math.floor(aux)]}TO` : centenas[Math.floor(aux)]} `;
            numeros[i] %= 100;
        }
        //? Procesa decenas
        if (numeros[i] >= 30 && numeros[i] < 100) {
            aux = numeros[i] / 10;
            aux2 += `${Number.isInteger(aux) ? decenas[aux] : `${decenas[Math.floor(aux)]} Y`} `;
            numeros[i] %= 10;
        }
        //? Procesa unidades y números especiales
        if (numeros[i] > 0) {
            aux2 += `${uniYesp[numeros[i]]} `;
            numeros[i] -= numeros[i];
        }
        //? Añade sufijos segun la posición
        switch (i) {
            case 4: aux2 = original === 1 ? 'UN BILLÓN ' : `${aux2.trim()} BILLONES `; break;
            case 2: aux2 = original === 1 ? 'UN MILLÓN ' : `${aux2.trim()} MILLONES `; break;
            default:
                if (i == 3 || i == 1)
                    aux2 = original === 1 ? 'MIL ' : `${aux2.replace('UNO', 'UN')}MIL `;
                break;
        }
        palabra += aux2; //? Concatena el resultado final
    }
    return palabra.trim();
}

/**
 * Función que transforma un número a su equivalente al mes del año
 * 
 * @param {number} num Número del mes (1 al 12)
 * @returns {string} Mes en texto
 */
export const NumAMes = num => {
    const meses = ['', 'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
    return meses[num];
}

/**
 * Función para encriptar información
 * 
 * @param {object} data Información a encriptar
 * @returns {string} Información encriptada
 */
export const EncryptData = data => CryptoJS.AES.encrypt(CryptoJS.lib.WordArray.create(Pako.deflate(JSON.stringify(data))), config().parsed.SECRET_KEY).toString();

/**
 * Función que convierte un Número a Letra/s
 * 
 * @param {int} num Número a convertir
 * @returns {string} Número convertido a letra
 */
export const NumeroALetras = num => {
    let res = '';
    while (num > 0) {
        num--;
        res = String.fromCharCode(65 + (num % 26)) + res;
        num = Math.floor(num / 26);
    }
    return res;
}

/**
 * Función que convierte Letra/s a un Número
 * 
 * @param {string} txt Letra a convertir
 * @returns {int} Letra convertido a número
 */
export const LetrasANumero = txt => {
    let res = 0;
    for (let i = 0; i < txt.length; i++)
        res = res * 26 + (txt.charCodeAt(i) - 64);
    return res;
}