import { request, response } from 'express';
import { config } from 'dotenv';
import Pako from 'pako';
import CryptoJS from 'crypto-js';

export const decryptPayload = (req = request, res = response, next) => {
    const { payload } = req.body;
    if (!payload)
        return res.status(400).json({
            success: false,
            msg: 'Payload faltante'
        });
    try {
        const compressed = Buffer.from(CryptoJS.AES.decrypt(payload, config().parsed.SECRET_KEY).toString(CryptoJS.enc.Base64), 'base64');
        const data = JSON.parse(Pako.inflate(compressed, { to: 'string' }));
        req.body = data;
        delete req.body.payload;
        next();
    } catch (err) {
        res.status(500).json({
            success: false,
            msg: 'Error al descifrar datos'
        });
    }
}