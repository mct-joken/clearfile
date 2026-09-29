const API_URL = "http://localhost:3000";

// ファイル一覧を取得
export async function getFiles() {
  const response = await fetch(`${API_URL}/api/files`);

  if (!response.ok) {
    throw new Error("ファイル取得失敗");
  }

  return await response.json();
}

// ファイルのタグを取得
export async function getFileTags(fileId: number) {
  const response = await fetch(
    `${API_URL}/api/files/${fileId}/tags`
  );

  if (!response.ok) {
    throw new Error("タグ取得失敗");
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
    throw new Error("タグ追加失敗");
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
    throw new Error("タグ削除失敗");
  }

  return await response.json();
}

export async function addFile(
  oneDriveId: string,
  fileName: string,
  address: string
) {
  const response = await fetch(`${API_URL}/api/files`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      oneDriveId,
      fileName,
      address,
    }),
  });

  if (!response.ok) {
    throw new Error("ファイル登録失敗");
  }

  return await response.json();
}