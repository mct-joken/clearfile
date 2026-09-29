const API_URL = "http://localhost:3000";

// ファイル一覧を取得
export async function getFiles() {
  const response = await fetch(`${API_URL}/api/files`);

  if (!response.ok) {
    throw new Error("ファイルの取得に失敗しました");
  }

  return await response.json();
}

// ファイルのタグを取得
export async function getFileTags(fileId: number) {
  const response = await fetch(
    `${API_URL}/api/files/${fileId}/tags`
  );

  if (!response.ok) {
    throw new Error("タグの取得に失敗しました");
  }

  return await response.json();
}

// ファイルにタグを追加
export async function addFileTag(fileId: number, tagName: string) {
  const response = await fetch(
    `${API_URL}/api/files/${fileId}/tags`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        name: tagName,
      }),
    }
  );

  if (!response.ok) {
    throw new Error("タグの追加に失敗しました");
  }

  return await response.json();
}

// ファイルからタグを削除
export async function deleteFileTag(
  fileId: number,
  tagId: number
) {
  const response = await fetch(
    `${API_URL}/api/files/${fileId}/tags/${tagId}`,
    {
      method: "DELETE",
    }
  );

  if (!response.ok) {
    throw new Error("タグの削除に失敗しました");
  }

  return await response.json();
}