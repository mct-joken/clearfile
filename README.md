# clearfile

OneDrive上のファイルをタグで管理するWebアプリケーション。

## Stack

- Frontend: React / TypeScript / Vite
- Backend: Node.js / Express
- Database: MySQL
- OneDrive: Microsoft Graph API / MSAL

## Setup

### 1. Install

```bash
npm install

### 2. Environment variables

`.env.example` を `.env` にコピーして設定する。

VITE_CLIENT_ID=
DB_HOST=
DB_PORT=3306
DB_USER=
DB_PASSWORD=
DB_NAME=clearfile_db

`.env` はGit管理対象外。

### 3. Database

`database/schema.sql` を `clearfile_db` に適用する。

mysql -u <DB_USER> -p clearfile_db < database/schema.sql

### 4. Run

Backend:

node backend/index.js

Frontend:

npm run dev

Frontend: http://localhost:5173
Backend: http://localhost:3000

## Database

- `files` - OneDriveファイル
- `tags` - タグ
- `file_tags` - ファイルとタグの関連

## API

- `GET /api/files`
- `POST /api/files`
- `POST /api/files/:fileId/tags`
- `GET /api/files/:fileId/tags`
- `DELETE /api/files/:fileId/tags/:tagId`
