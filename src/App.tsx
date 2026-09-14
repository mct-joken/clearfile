import { Providers } from "@microsoft/mgt-element";
import { Msal2Provider } from "@microsoft/mgt-msal2-provider";

import { Login, FileList } from "@microsoft/mgt-react";

import { useEffect, useRef, useState } from "react";

import "./App.css";

// アプリの起動時に一度だけ実行
Providers.globalProvider = new Msal2Provider({
  clientId: import.meta.env.VITE_CLIENT_ID as string,
  scopes: ["Files.Read", "Files.Read.All", "User.Read"], // OneDriveを読む許可をもらう
});

function App() {
  // 現在表示しているフォルダのIDを管理する（初期値は 'root'）
  const [currentFolderId, setCurrentFolderId] = useState<string>("root");
  // 過去に移動したフォルダIDの履歴（「戻る」ボタン用）
  const [folderHistory, setFolderHistory] = useState<string[]>([]);
  const fileListRef = useRef<any>(null);
  //タグ管理用
  const [selectedFile, setSelectedFile] = useState<any>(null);
  const [localTags, setLocalTags] = useState<Record<string, string[]>>({});
  const [actionModalFile, setActionModalFile] = useState<any>(null);
  useEffect(() => {
    const fileList = fileListRef.current;
    if (!fileList) return;

    const originalShowChildren = fileList.showChildren?.bind(fileList);

    fileList.showChildren = (fileId: string) => {
      const itemDOM = fileList.renderRoot?.getElementById?.(
        `file-list-item-${fileId}`,
      );

      if (itemDOM) {
        fileList.renderChildren(fileId, itemDOM);
        return;
      }

      if (typeof originalShowChildren === "function") {
        originalShowChildren(fileId);
        }
    };

useEffect(() => {
  // Document または DocumentFragment の querySelector 全体に安全パッチを当てる
  const patchQuerySelector = (proto: any) => {
    if (!proto || proto.__querySelectorPatched) return;
    const original = proto.querySelector;

    proto.querySelector = function (selector: string) {
      try {
        return original.call(this, selector);
      } catch (e) {
        // #file-list-item- 系の記号入りIDでエラーが出たら CSS.escape で修復して再実行
        if (typeof selector === "string" && selector.startsWith("#")) {
          try {
            const safeSelector = "#" + CSS.escape(selector.slice(1));
            return original.call(this, safeSelector);
          } catch {
            // エスケープしてもダメな場合は元の例外を投げる
          }
        }
        throw e;
      }
    };
    proto.__querySelectorPatched = true;
  };

  // 通常のDOM（Document）と Shadow DOM（DocumentFragment）の両方に適用
  patchQuerySelector(Document.prototype);
  patchQuerySelector(DocumentFragment.prototype);
}, []);

  // ファイルやフォルダがクリックされた時の処理
  const handleItemClick = (e: any) => {
    e?.preventDefault?.();
    const clickedItem = e?.detail?.item ?? e?.detail ?? e;
    if (!clickedItem) return;

    if (clickedItem.folder) {
      const nextFolderId = clickedItem.id;
      setFolderHistory((prev) => [...prev, currentFolderId]);
      setCurrentFolderId(nextFolderId);
      setSelectedFile(null);
    } else {
      // ファイルクリック時: 「開く」か「タグ編集」かを選択するダイアログを表示
      setActionModalFile(clickedItem);
    }
  };

  // 「前のフォルダに戻る」ボタンの処理
const handleBackClick = () => {
    if (folderHistory.length === 0) return;
    const previousFolderId = folderHistory[folderHistory.length - 1];
    setFolderHistory((prev) => prev.slice(0, -1));
    setCurrentFolderId(previousFolderId);
    setSelectedFile(null); // 戻る時も選択を解除（追加）
  };

  useEffect(() => {
    //tagデータ読み込み
    const savedTags = localStorage.getItem("my_onedrive_tags");
    if (savedTags) {
      setLocalTags(JSON.parse(savedTags));
    }
  }, []);

  //タグの更新処理
  useEffect(() => {
    localStorage.setItem("my_onedrive_tags", JSON.stringify(localTags));
  }, [localTags]);

  // タグの追加処理
  const addTag = (fileId: string, tag: string) => {
    setLocalTags((prev) => {
      const currentFileTags = prev[fileId] || [];
      if (currentFileTags.includes(tag)) return prev; // 重複防止
      return { ...prev, [fileId]: [...currentFileTags, tag] };
    });
  };

  //タグの削除処理
  const removeTag = (fileId: string, tagToRemove: string) => {
    setLocalTags((prev) => {
      const currentFileTags = prev[fileId] || [];
      return {
        ...prev,
        [fileId]: currentFileTags.filter((tag) => tag !== tagToRemove),
      };
    });
  };

  return (
    <div style={{ display: "flex", gap: "20px", position: "relative" }}>
      {/* 左側: ファイル一覧 */}
      <div style={{ flex: 2 }}>
        <h1>マイ OneDrive</h1>
        <Login />
        <hr />

        {currentFolderId !== "root" && (
          <button onClick={handleBackClick} style={{ marginBottom: "10px" }}>
            ⬅ 前のフォルダに戻る
          </button>
        )}

        <FileList
          ref={fileListRef}
          itemId={currentFolderId}
          itemClick={handleItemClick}
        />
      </div>

      {/* 右側: タグ編集パネル */}
      <div style={{ flex: 1, borderLeft: "1px solid #ccc", paddingLeft: "20px" }}>
        <h2>タグ編集パネル</h2>
        {selectedFile ? (
          <div>
            <p>
              <strong>選択中:</strong> {selectedFile.name}
            </p>
            
            <div>
              <strong>現在のタグ:</strong>
              <ul style={{ paddingLeft: "20px" }}>
                {(localTags[selectedFile.id] || []).map((tag, idx) => (
                  <li key={idx} style={{ marginBottom: "5px" }}>
                    {tag} {/* タグ削除ボタン */}
                    <button
                      onClick={() => removeTag(selectedFile.id, tag)}
                      style={{
                        marginLeft: "8px",
                        padding: "2px 6px",
                        cursor: "pointer",
                        backgroundColor: "#ff4d4f",
                        color: "#fff",
                        border: "none",
                        borderRadius: "3px",
                      }}
                    >
                      削除
                    </button>
                  </li>
                ))}
              </ul>
            </div>

            <div style={{ marginTop: "15px" }}>
              <strong>タグを追加:</strong>
              <br />
              <button 
                onClick={() => addTag(selectedFile.id, "重要")}
                style={{ marginTop: "5px", marginRight: "5px" }}
              >
                + 重要
              </button>
              <button
                onClick={() => addTag(selectedFile.id, "確認済み")}
                style={{ marginTop: "5px" }}
              >
                + 確認済み
              </button>
            </div>
          </div>
        ) : null}
          
      {/* ファイルクリック時に表示される選択ダイアログ（モーダル） */}
      {actionModalFile && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            width: "100vw",
            height: "100vh",
            backgroundColor: "rgba(0, 0, 0, 0.5)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 1000,
          }}
        >
          <div
            style={{
              backgroundColor: "#fff",
              padding: "20px",
              borderRadius: "8px",
              minWidth: "300px",
              textAlign: "center",
              boxShadow: "0 4px 10px rgba(0,0,0,0.3)",
            }}
          >
            <h3>操作を選択</h3>
            <p style={{ wordBreak: "break-all" }}>
              <strong>{actionModalFile.name}</strong>
            </p>
            
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                gap: "10px", marginTop: "20px"
              }}
            >
              {/* 1. OneDriveで開く */}
              <button
                onClick={() => {
                  if (actionModalFile.webUrl) {
                    window.open(actionModalFile.webUrl, "_blank", "noopener,noreferrer");
                  }
                  setActionModalFile(null);
                }}
                style={{ padding: "8px", cursor: "pointer", backgroundColor: "#0078d4", color: "#fff", border: "none", borderRadius: "4px" }}
              >
                開く
              </button>

              {/* 2. タグを編集する */}
              <button
                onClick={() => {
                  setSelectedFile(actionModalFile);
                  setActionModalFile(null);
                }}
                style={{ padding: "8px", cursor: "pointer", backgroundColor: "#28a745", color: "#fff", border: "none", borderRadius: "4px" }}
              >
                タグを編集
              </button>

              {/* 3. キャンセル */}
              <button
                onClick={() => setActionModalFile(null)}
                style={{ padding: "5px", cursor: "pointer", backgroundColor: "#ccc", border: "none", borderRadius: "4px", marginTop: "5px" }}
              >
                キャンセル
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
