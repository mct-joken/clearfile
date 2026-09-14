require('dotenv').config();

const express = require('express');
const mysql = require('mysql2/promise');

const app = express();
const PORT = 3000;

// JSON形式のデータを受け取れるようにする
app.use(express.json());

// MySQLとの接続を作る
const pool = mysql.createPool({
    host: process.env.DB_HOST,
    port: process.env.DB_PORT,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME
});

// ファイル登録API
app.post('/api/files', async (req, res) => {
    try {
        const { fileName, address } = req.body;

        const [result] = await pool.execute(
            'INSERT INTO files (file_name, address) VALUES (?, ?)',
            [fileName, address]
        );

        res.status(201).json({
            message: 'ファイル登録完了',
            fileId: result.insertId
        });

    } catch (error) {
        console.error(error);

        res.status(500).json({
            message: 'ファイル登録失敗'
        });
    }
});

app.listen(PORT, () => {
    console.log(`サーバー起動: http://localhost:${PORT}`);
});
