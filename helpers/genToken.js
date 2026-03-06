import JWT from 'jsonwebtoken';
import { config } from 'dotenv';

export const genToken = (datos, expire = '12h') => new Promise((resolve, reject) => {
    try {
        const token = JWT.sign(datos, config({ quiet: true }).parsed.SECRET, { expiresIn: expire });
        resolve(token);
    } catch (err) {
        reject('Sin JWT');
    }
})