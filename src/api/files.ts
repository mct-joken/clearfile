import { Providers } from "@microsoft/mgt-element";
const API_URL = "http://172.21.7.163:3000";

async function getAccessToken() {
  const provider = Providers.globalProvider;

  const accessToken = await provider.getAccessToken({
    scopes: ["User.Read"],
  });

  if (!accessToken) {
    throw new Error("アクセストークンを取得できませんでした");
  }

  return accessToken;
}

// ファイル一覧を取得
export async function getFiles() {
  const accessToken = await getAccessToken();

  const response = await fetch(`${API_URL}/api/files`, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    throw new Error("ファイル取得失敗");
  }

  return await response.json();
}

// ファイルのタグを取得
export async function getFileTags(fileId: number) {
  const accessToken = await getAccessToken();

    const response = await fetch(
      `${API_URL}/api/files/${fileId}/tags`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

  if (!response.ok) {
    throw new Error("タグ取得失敗");
  }

  return await response.json();
}

// ファイルにタグを追加
export async function addFileTag(fileId: number, tagName: string) {
  const accessToken = await getAccessToken();

  const response = await fetch(
    `${API_URL}/api/files/${fileId}/tags`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${accessToken}`,
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
  const accessToken = await getAccessToken();

  const response = await fetch(
    `${API_URL}/api/files/${fileId}/tags/${tagId}`,
    {
      method: "DELETE",
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    throw new Error("タグ削除失敗");
  }

  return await response.json();
}

//ファイルを登録
export async function addFile(
  oneDriveId: string,
  fileName: string,
  address: string
) {
  const accessToken = await getAccessToken();

  const response = await fetch(`${API_URL}/api/files`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${accessToken}`,
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