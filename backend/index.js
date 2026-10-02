require('dotenv').config({ path: require('path').join(__dirname, '.env') });

const express = require('express');
const cors = require('cors');
const mysql = require('mysql2/promise');

const app = express();
app.use(cors({
    origin: 'http://localhost:5173'
}));

const PORT = 3000;

// Microsoft Graphからログインユーザーを取得する
async function authenticateUser(req, res, next) {
    try {
        const authHeader = req.headers.authorization;

        if (!authHeader || !authHeader.startsWith('Bearer ')) {
            return res.status(401).json({
                message: '認証情報がありません'
            });
        }

        const accessToken = authHeader.substring('Bearer '.length);

        // Microsoft Graphにアクセストークンを渡して
        // 現在のユーザー情報を取得する
        const response = await fetch(
            'https://graph.microsoft.com/v1.0/me?$select=id',
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`
                }
            }
        );

        if (!response.ok) {
            return res.status(401).json({
                message: '認証に失敗しました'
            });
        }

        const user = await response.json();

        // usersテーブルにユーザーが存在するか確認
        const [users] = await pool.execute(
            `SELECT id
             FROM users
             WHERE microsoft_user_id = ?`,
            [user.id]
        );

        let userId;

        if (users.length > 0) {
            // 既存ユーザー
            userId = users[0].id;
        } else {
            // 初めて利用するユーザー
            const [result] = await pool.execute(
                `INSERT INTO users (microsoft_user_id)
                 VALUES (?)`,
                [user.id]
            );

            userId = result.insertId;
        }

        // 後続のAPIから使えるようにする
        req.userId = userId;

        next();

    } catch (error) {
        console.error(error);

        res.status(500).json({
            message: 'ユーザー認証失敗'
        });
    }
}

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
app.post('/api/files', authenticateUser, async (req, res) => {
    try {
        const { oneDriveId, fileName, address } = req.body;

        // すでに同じOneDriveファイルが登録されているか確認
        const [existingFiles] = await pool.execute(
            `SELECT id
            FROM files
            WHERE user_id = ?
                AND onedrive_id = ?`,
            [req.userId, oneDriveId]
        );

        // すでに登録されている場合
        if (existingFiles.length > 0) {
            return res.json({
                message: 'ファイルはすでに登録されています',
                fileId: existingFiles[0].id
            });
        }

        // 未登録の場合は新しく登録
        const [result] = await pool.execute(
            `INSERT INTO files
                (user_id, onedrive_id, file_name, address)
            VALUES (?, ?, ?, ?)`,
            [req.userId, oneDriveId, fileName, address]
        );

        res.json({
            message: 'ファイル登録',
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
app.get('/api/files', authenticateUser, async (req, res) => {
    try {
        const [rows] = await pool.execute(
            `SELECT *
            FROM files
            WHERE user_id = ?`,
            [req.userId]
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
app.post('/api/files/:fileId/tags', authenticateUser, async (req, res) => {
    try {
        const { fileId } = req.params;

        const [files] = await pool.execute(
            `SELECT id
            FROM files
            WHERE id = ?
                AND user_id = ?`,
            [fileId, req.userId]
        );

        if (files.length === 0) {
            return res.status(404).json({
                message: 'ファイルが見つかりません'
            });
        }

        const { name } = req.body;

        // タグがすでに存在するか確認
        const [tags] = await pool.execute(
            `SELECT id
            FROM tags
            WHERE user_id = ?
                AND name = ?`,
            [req.userId, name]
        );

        let tagId;

        if (tags.length > 0) {
            // すでに存在する場合
            tagId = tags[0].id;
        } else {
            // 存在しない場合は新しく作る
            const [result] = await pool.execute(
                `INSERT INTO tags (user_id, name)
                VALUES (?, ?)`,
                [req.userId, name]
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
app.get('/api/files/:fileId/tags', authenticateUser, async (req, res) => {
    try {
        const { fileId } = req.params;

        const [rows] = await pool.execute(
            `SELECT tags.id, tags.name
            FROM tags
            INNER JOIN file_tags
                ON tags.id = file_tags.tag_id
            INNER JOIN files
                ON files.id = file_tags.file_id
            WHERE file_tags.file_id = ?
            AND files.user_id = ?
            AND tags.user_id = ?`,
            [fileId, req.userId, req.userId]
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
app.delete(
    '/api/files/:fileId/tags/:tagId',
    authenticateUser,
    async (req, res) => {
    try {
        const { fileId, tagId } = req.params;

       await pool.execute(
            `DELETE FROM file_tags
            WHERE file_id = ?
            AND tag_id = ?
            AND EXISTS (
                SELECT 1
                FROM files
                WHERE files.id = file_tags.file_id
                    AND files.user_id = ?
            )`,
            [fileId, tagId, req.userId]
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
