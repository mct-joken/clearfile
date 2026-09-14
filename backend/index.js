require('dotenv').config();
const mysql = require('mysql2/promise');

async function main() {
    try {
        const connection = await mysql.createConnection({
            host: process.env.DB_HOST,
            port: process.env.DB_PORT,
            user: process.env.DB_USER,
            password: process.env.DB_PASSWORD,
            database: process.env.DB_NAME
        });

        console.log('MySQL接続成功');

        await connection.end();
    } catch (error) {
        console.error('MySQL接続失敗');
        console.error(error.message);
    }
}

main();
