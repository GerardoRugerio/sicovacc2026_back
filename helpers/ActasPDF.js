import PDFDocument from 'pdfkit';

export const TextoMultiFuente = (doc = PDFDocument, x, y, width, fontSize, fontBlocks, options = {}) => {
    const {
        fillColor = '#000',
        lineHeight = 1.15,
        maxHeight = Infinity,
        backgroundColor = null,
        align = 'left' //? 'left', 'right', 'justify', 'center'
    } = options;
    //? 1. Separar texto en bloques de palabras, detectando saltos de línea explícitos (\n)
    fontBlocks = Array.isArray(fontBlocks) ? fontBlocks : [fontBlocks];
    const palabras = fontBlocks.flatMap(block => {
        return block.text.split('\n').flatMap((linea, index, array) => {
            const palabrasLinea = linea.split(' ').map(palabra => ({
                texto: palabra,
                font: block.font,
                underline: block.underline || false,
                fontSize: block.fontSize ?? fontSize,
                fillColor: block.fillColor ?? fillColor,
                espacioAntes: index !== 0,
                lineBreakForced: false
            }));
            //? Si no es la última línea, agregamos una marca de salto de línea forzado
            if (index < array.length - 1)
                palabrasLinea.push({ texto: '', font: block.font, fontSize: block.fontSize ?? fontSize, underline: false, lineBreakForced: true });
            return palabrasLinea;
        });
    });
    //? 2. Agrupar palabras en líneas según el ancho máximo o salto forzado
    const lineas = [];
    let lineaActual = [];
    let anchoActual = 0;
    for (const palabra of palabras) {
        doc.font(palabra.font).fontSize(palabra.fontSize);
        const siguientePalabra = palabras[palabras.indexOf(palabra) + 1];
        const agregarEspacio = siguientePalabra && !siguientePalabra.texto.startsWith(',');
        const palabraWidth = doc.widthOfString(palabra.texto + (agregarEspacio ? ' ' : ''));
        //? Si la palabra marca un salto forzado o ya no cabe en la línea actual
        if (palabra.lineBreakForced || (anchoActual + palabraWidth > width && lineaActual.length > 0)) {
            lineas.push(lineaActual);
            lineaActual = palabra.lineBreakForced ? [] : [palabra];
            anchoActual = palabra.lineBreakForced ? 0 : palabraWidth;
        } else {
            lineaActual.push(palabra);
            anchoActual += palabraWidth;
        }
    }
    if (lineaActual.length > 0)
        lineas.push(lineaActual);
    //? 3. Dibujar fondo si se requiere
    const totalHeight = lineas.reduce((acc, linea) => {
        const maxFontSize = Math.max(...linea.map(p => p.fontSize));
        return acc + maxFontSize * lineHeight;
    }, 0);
    if (backgroundColor)
        doc.rect(x, y, width, Math.min(totalHeight, maxHeight)).fillAndStroke(backgroundColor, backgroundColor);
    //? 4. Dibujar texto línea por línea
    let offsetY = y;
    for (let i = 0; i < lineas.length; i++) {
        const linea = lineas[i];
        const isLastLine = i === lineas.length - 1;
        const maxFontSizeLinea = Math.max(...linea.map(p => p.fontSize));
        const anchoLinea = linea.reduce((acc, palabra, index) => {
            const siguiente = linea[index + 1];
            const agregarEspacio = siguiente && !siguiente.texto.startsWith(',');
            doc.font(palabra.font).fontSize(palabra.fontSize);
            return acc + doc.widthOfString(palabra.texto + (agregarEspacio ? ' ' : ''));
        }, 0);
        let offsetX = x;
        if (align === 'center')
            offsetX = x + (width - anchoLinea) / 2;
        else if (align === 'right')
            offsetX = x + (width - anchoLinea);
        let extraSpace = 0;
        if (align === 'justify' && linea.length > 1 && !isLastLine) {
            const espaciosValidos = linea.slice(0, -1).filter((_, i) => {
                const siguiente = linea[i + 1];
                return siguiente && !/^[,.;:)\]]/.test(siguiente.texto.trim());
            });
            const textoWidth = linea.reduce((acc, palabra) => {
                doc.font(palabra.font).fontSize(palabra.fontSize);
                return acc + doc.widthOfString(palabra.texto);
            }, 0);
            extraSpace = espaciosValidos.length > 0 ? (width - textoWidth) / espaciosValidos.length : 0;
        }
        let cursorX = offsetX;
        let underlineStartX = null;
        for (let j = 0; j < linea.length; j++) {
            const palabra = linea[j];
            const siguiente = linea[j + 1];
            const comienzaConSigno = siguiente && /^[,.;:)\]]/.test(siguiente.texto.trim());
            const espacio = j < linea.length - 1 ? (comienzaConSigno ? 0 : (align === 'justify' && !isLastLine ? extraSpace : doc.widthOfString(' '))) : 0;
            doc.font(palabra.font).fontSize(palabra.fontSize);
            const palabraWidth = doc.widthOfString(palabra.texto);
            //? Dibujar palabra
            doc.fillColor(palabra.fillColor).text(palabra.texto, palabra.underline ? cursorX + (espacio / 2) : cursorX, offsetY, { lineBreak: false });
            //? Subrayado continuo
            if (palabra.underline) {
                if (underlineStartX === null)
                    underlineStartX = cursorX;
            } else if (underlineStartX !== null) {
                const underlineY = offsetY + palabra.fontSize * 0.85;
                doc.moveTo(underlineStartX, underlineY).lineTo(cursorX, underlineY).strokeColor(fillColor).lineWidth(0.5).stroke();
                underlineStartX = null;
            }
            cursorX += palabraWidth + espacio;
        }
        //? Si termina con subrayado activo
        if (underlineStartX !== null) {
            const underlineY = offsetY + maxFontSizeLinea * 0.85;
            doc.moveTo(underlineStartX, underlineY).lineTo(cursorX, underlineY).strokeColor(fillColor).lineWidth(0.5).stroke();
        }
        underlineStartX = null;
        offsetY += maxFontSizeLinea * lineHeight;
        if (offsetY > y + maxHeight)
            break;
    }
}

export const DibujarTablaPDF = (doc = PDFDocument, x, y, encabezados, columnas, datos, options = {}) => {
    const { margen = 2, fontSize = 7 } = options;
    let offsetY = y;
    //? Función interna para dibujas una fila
    const dibujarFila = (blocks) => {
        let offsetX = x, height = 0;
        //? Calcula la altura de cada celda según su contenido
        const alturas = blocks.map((block, index) => CalcularAltoAncho(doc, block, block[0].fontSize ?? fontSize, columnas[index].width - margen + 1).totalHeight);
        const maxHeight = Math.max(...alturas) + (margen * 2);
        //? Dibuja cada celda de la fila
        blocks.map((block, index) => {
            const colWidth = columnas[index].width;
            height = alturas[index];
            doc.rect(offsetX, offsetY, colWidth, maxHeight).fillAndStroke(block[0].background ?? '#FFF', block[0].strokeColor ?? '#000');
            //? Escribir el texto dentro de la celda usanto la función TextoMultiFuente
            TextoMultiFuente(doc, offsetX + (margen / 2), offsetY + ((maxHeight - height) / 2) + 1.5, colWidth - margen, block[0].fontSize ?? fontSize, block, {
                fillColor: block[0].fillColor ?? '#000',
                maxHeight: maxHeight - (margen * 2),
                align: block[0].align ?? columnas[index].align
            });
            offsetX += colWidth;
        });
        offsetY += maxHeight;
    };
    dibujarFila(encabezados);
    datos.map(blocks => dibujarFila(blocks));
}

export const CalcularAltoAncho = (doc = PDFDocument, fontBlocks, fontSize, width, lineHeight = 1.15) => {
    //? Asegura que fontBlocks siempre sea un arreglo
    fontBlocks = Array.isArray(fontBlocks) ? fontBlocks : [fontBlocks];
    //? Dividir el texto en palabras
    const palabras = fontBlocks.flatMap(block => {
        return block.text.split('\n').flatMap((linea, index, array) => {
            //? Cada línea se separa en palabras
            const palabrasLinea = linea.split(' ').map(palabra => ({
                texto: palabra,
                font: block.font,
                fontSize: block.fontSize ?? fontSize,
                lineBreakForced: false
            }));
            //? Si hay un salto de línea forzado, se agrega un marcador
            if (index < array.length - 1)
                palabrasLinea.push({
                    texto: '',
                    font: block.font,
                    fontSize: block.fontSize ?? fontSize,
                    lineBreakForced: true
                });
            return palabrasLinea;
        });
    });
    const lineas = [];
    let lineaActual = [];
    let anchoActual = 0;
    //? Recorremos todas las palabras
    for (const palabra of palabras) {
        doc.font(palabra.font).fontSize(palabra.fontSize);
        const palabraWidth = doc.widthOfString(palabra.texto + ' ');
        //? Si la palabra no cabe en la línea o hay un salto forzado
        if (palabra.lineBreakForced || (anchoActual + palabraWidth > width && lineaActual.length > 0)) {
            lineas.push(lineaActual);
            lineaActual = palabra.lineBreakForced ? [] : [palabra];
            anchoActual = palabra.lineBreakForced ? 0 : palabraWidth;
        } else { //? Agrega palabra a la línea actual
            lineaActual.push(palabra);
            anchoActual += palabraWidth;
        }
    }
    //? Agrega la última línea si quedó con palabras
    if (lineaActual.length > 0)
        lineas.push(lineaActual);
    //? Calcula el alto total sumando el alto de cada línea
    const totalHeight = lineas.reduce((acc, linea) => {
        const maxFontSize = Math.max(...linea.map(p => p.fontSize));
        return acc + (maxFontSize * lineHeight);
    }, 0);
    //? Calcula ancho máximop de las lineas
    const maxWidth = Math.max(...lineas.map(linea =>
        linea.reduce((acc, palabra) => {
            doc.font(palabra.font).fontSize(palabra.fontSize);
            return acc + doc.widthOfString(palabra.texto + ' ');
        }, 0)
    ));
    return { totalHeight, maxWidth };
}

//? Suma el ancho para tener el ancho total de una tabla
const SumarWidth = arr => arr.slice(0, -2).reduce((acum, item) => acum + item.width, 0);