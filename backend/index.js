require('dotenv').config();

const express = require('express');
const cors = require('cors');
const mysql = require('mysql2/promise');

const app = express();
app.use(cors({
    origin: 'http://localhost:5173'
}));

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

// ファイル一覧取得API
app.get('/api/files', async (req, res) => {
    try {
        const [rows] = await pool.execute(
            'SELECT * FROM files'
        );

        res.json(rows);

    } catch (error) {
        console.error(error);
        res.status(500).json({
            message: 'ファイル取得失敗'
        });
    }
});

// ファイルにタグを追加するAPI
app.post('/api/files/:fileId/tags', async (req, res) => {
    try {
        const { fileId } = req.params;
        const { name } = req.body;

        // タグがすでに存在するか確認
        const [tags] = await pool.execute(
            'SELECT id FROM tags WHERE name = ?',
            [name]
        );

        let tagId;

        if (tags.length > 0) {
            // すでに存在する場合
            tagId = tags[0].id;
        } else {
            // 存在しない場合は新しく作る
            const [result] = await pool.execute(
                'INSERT INTO tags (name) VALUES (?)',
                [name]
            );

            tagId = result.insertId;
        }

        // ファイルとタグを関連付ける
        await pool.execute(
            'INSERT INTO file_tags (file_id, tag_id) VALUES (?, ?)',
            [fileId, tagId]
        );

        res.status(201).json({
            message: 'タグ追加',
            fileId: Number(fileId),
            tagId: tagId
        });

    } catch (error) {
        console.error(error);

        res.status(500).json({
            message: 'タグ追加失敗'
        });
    }
});

// ファイルのタグを取得するAPI
app.get('/api/files/:fileId/tags', async (req, res) => {
    try {
        const { fileId } = req.params;

        const [rows] = await pool.execute(
            `SELECT tags.id, tags.name
             FROM tags
             INNER JOIN file_tags
             ON tags.id = file_tags.tag_id
             WHERE file_tags.file_id = ?`,
            [fileId]
        );

        res.json(rows);

    } catch (error) {
        console.error(error);

        res.status(500).json({
            message: 'タグ取得失敗'
        });
    }
});

// ファイルからタグを削除するAPI
app.delete('/api/files/:fileId/tags/:tagId', async (req, res) => {
    try {
        const { fileId, tagId } = req.params;

        await pool.execute(
            `DELETE FROM file_tags
             WHERE file_id = ? AND tag_id = ?`,
            [fileId, tagId]
        );

        res.json({
            message: 'タグ削除'
        });

    } catch (error) {
        console.error(error);

        res.status(500).json({
            message: 'タグ削除失敗'
        });
    }
});

app.listen(PORT, () => {
    console.log(`サーバー起動: http://localhost:${PORT}`);
});
