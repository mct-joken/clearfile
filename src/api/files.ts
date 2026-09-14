const API_URL = "http://localhost:3000";

export async function getFiles() {
  const response = await fetch(`${API_URL}/api/files`);

  if (!response.ok) {
    throw new Error("ファイル取得失敗");
  }

  return await response.json();
}

export async function getFileTags(fileId: number) {
  const response = await fetch(
    `${API_URL}/api/files/${fileId}/tags`
  );

  if (!response.ok) {
    throw new Error("タグ取得失敗");
  }

  return await response.json();
}