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
  // タグ管理用
  const [localTags, setLocalTags] = useState<Record<string, string[]>>({});
  const [actionModalFile, setActionModalFile] = useState<any>(null);
  const [tagModalFile, setTagModalFile] = useState<any>(null);
  const [fileNameMap, setFileNameMap] = useState<Record<string, string>>({});
  const [newTagName, setNewTagName] = useState<string>("");
  const [newTagAddToQuickAdd, setNewTagAddToQuickAdd] = useState<boolean>(false);
  const [quickAddTags, setQuickAddTags] = useState<string[]>(["重要", "確認済み"]);
  const [searchPool, setSearchPool] = useState<string[]>([]);
  // タグ検索用
  const [selectedTags, setSelectedTags] = useState<string[]>([]);

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

    return () => {
      if (typeof originalShowChildren === "function") {
        fileList.showChildren = originalShowChildren;
      }
    };
  }, [currentFolderId]);


  // ファイルやフォルダがクリックされた時の処理
  const handleItemClick = (e: any) => {
    e?.preventDefault?.();
    e?.stopPropagation?.();
    const clickedItem = e?.detail?.item ?? e?.detail ?? e;
    if (!clickedItem) return;
    if (clickedItem.folder) {
      const nextFolderId = clickedItem.id;
      setFolderHistory((prev) => [...prev, currentFolderId]);
      setCurrentFolderId(nextFolderId);
      setTagModalFile(null);
    } else {
      setFileNameMap((prev) => ({ ...prev, [clickedItem.id]: clickedItem.name }));
      setActionModalFile(clickedItem);
    }
  };

  // 「前のフォルダに戻る」ボタンの処理
const handleBackClick = () => {
    if (folderHistory.length === 0) return;
    const previousFolderId = folderHistory[folderHistory.length - 1];
    setFolderHistory((prev) => prev.slice(0, -1));
    setCurrentFolderId(previousFolderId);
    setTagModalFile(null);
  };

  useEffect(() => {//tagデータ読み込み
    const savedTags = localStorage.getItem("my_onedrive_tags");
    if (savedTags) {
      setLocalTags(JSON.parse(savedTags));
    }

    const savedQuickAddTags = localStorage.getItem("my_onedrive_quick_add_tags");
    if (savedQuickAddTags) {
      setQuickAddTags(JSON.parse(savedQuickAddTags));
    }

    const savedFileNameMap = localStorage.getItem("my_onedrive_file_names");
    if (savedFileNameMap) {
      setFileNameMap(JSON.parse(savedFileNameMap));
    }
  }, []);

   useEffect(() => {
    localStorage.setItem("my_onedrive_file_names", JSON.stringify(fileNameMap));
  }, [fileNameMap]);

  useEffect(() => {
    const derivedSearchPool = Array.from(
      new Set(Object.values(localTags).flat().map((tag) => tag.trim())),
    ).sort((a, b) => a.localeCompare(b));
    setSearchPool(derivedSearchPool);
  }, [localTags]);

  // タグの更新処理
  useEffect(() => {
    localStorage.setItem("my_onedrive_tags", JSON.stringify(localTags));
  }, [localTags]);

  useEffect(() => {
    localStorage.setItem("my_onedrive_quick_add_tags", JSON.stringify(quickAddTags));
  }, [quickAddTags]);

  // タグの追加処理
  const addTag = (fileId: string, tag: string) => {
    const cleanedTag = tag.trim();
    if (!cleanedTag) return;

    setLocalTags((prev) => {
      const currentFileTags = prev[fileId] || [];
      const alreadyExists = currentFileTags.some(
        (existingTag) => existingTag.toLowerCase() === cleanedTag.toLowerCase(),
      );

      if (alreadyExists) return prev; // 重複防止
      return { ...prev, [fileId]: [...currentFileTags, cleanedTag] };
    });
  };

  const addQuickAddTag = (tag: string) => {
    const cleanedTag = tag.trim();
    if (!cleanedTag) return;

    setQuickAddTags((prev) => {
      const alreadyExists = prev.some(
        (existingTag) => existingTag.toLowerCase() === cleanedTag.toLowerCase(),
      );

      if (alreadyExists) return prev;
      return [...prev, cleanedTag].sort((a, b) => a.localeCompare(b));
    });
  };

  // タグの削除処理
  const removeTag = (fileId: string, tagToRemove: string) => {
    setLocalTags((prev) => {
      const currentFileTags = prev[fileId] || [];
      return {
        ...prev,
        [fileId]: currentFileTags.filter((tag) => tag !== tagToRemove),
      };
    });
  };

  // 各タグの件数を集計する処理
  const tagsCounts = Object.values(localTags).flat().reduce<Record<string, number>>((acc, tag) => {
    const cleanedTags = tag.trim();
    if (!cleanedTags) return acc;
    acc[cleanedTags] = (acc[cleanedTags] || 0) + 1;
    return acc;
  }, {});
  
  const allTags = searchPool;

  const toggleSelectedTag = (tag: string) => {
    setSelectedTags((prev) => {
      if (prev.includes(tag)) {
        return prev.filter((item) => item !== tag);
      }

      return [...prev, tag];
    });
  };

  const tagSearchResults = Object.entries(localTags)
    .filter(([, fileTags]) => {
      if (selectedTags.length === 0) return true;

      return selectedTags.every((selectedTag) =>
        fileTags.some((fileTag) => fileTag.toLowerCase() === selectedTag.toLowerCase()),
      );
    })
    .map(([fileId, fileTags]) => ({
      id: fileId,
      name: fileNameMap[fileId] ?? fileId,
      tags: fileTags,
    }));

  return (
    <div style={{ padding: "20px", textAlign: "left" }}>
      
      {/* ヘッダー */}
      <div style={{  flex: 2 }}>
        <header style={{ display: "flex", alignItems: "center" }}>
          <h1>マイ OneDrive</h1>
          <div style={{ transform: "scale(1.2)" , position: "absolute", right: "10%" }}>
            <Login />
          </div>
        </header>
          <hr />
      </div>
        {currentFolderId !== "root" && (
          <button onClick={handleBackClick} style={{ marginBottom: "10px" }}>
            ⬅ 前のフォルダに戻る
          </button>
        )}

        {/* ヘッダー下を二分割  左:タグ一覧 右:検索とファイル一覧*/}
        <div style={{ display: "flex", gap: "20px", alignItems: "flex-start"}}>

          {/* タグ一覧 */}
          <aside style={{ flex: "0 0 200px",padding: "16px",border: "1px solid #ddd",borderRadius: "8px",backgroundColor: "#ffffff",}}>
            <h3 style={{ marginTop: "0", marginBottom: "12px" }}>タグ一覧</h3>
              {Object.keys(tagsCounts).length === 0 ? (
              <span style={{ color: "#666", fontSize: "14px" }}>タグはありません</span>
              ) : (
                <ul style={{ paddingLeft: "20px", margin: 0 }}>
                  {Object.entries(tagsCounts).map(([tag, count]) => (
                   <li key={tag} style={{ marginBottom: "6px" }}>
                    <button onClick={() => toggleSelectedTag(tag)} 
                      style={{
                              background: "none",
                              border: "none",
                              color: "#0078d4",
                              cursor: "pointer",
                              padding: 0,
                              textDecoration: selectedTags.includes(tag) ? "underline" : "none",
                              fontWeight: selectedTags.includes(tag) ? "bold" : "normal",
                            }}>
                      #{tag} ({count})
                    </button>
                   </li>
                    ))}
                  </ul>
                  )}
          </aside>

          {/* 右側の検索&ファイル一覧まとめたボックス */}
          <main style={{ flex: 1 , padding: "20px", border: "1px solid #ddd", borderRadius: "8px", backgroundColor: "#ffffff"}}>        
            {/* 検索プール */}
            <div style={{ fontWeight: "bold", marginBottom: "8px" }}>検索プール</div>
            <div style={{ marginTop: "10px", display: "flex", gap: "8px", flexWrap: "wrap" }}>
              {allTags.length === 0 ? (
                <span style={{ color: "#666" }}>タグがまだありません</span>
              ) : (
                allTags.map((tag) => (
                  <button
                    key={tag}
                    onClick={() => toggleSelectedTag(tag)}
                    style={{
                      padding: "5px 10px",
                      cursor: "pointer",
                      borderRadius: "20px",
                      border: selectedTags.includes(tag) ? "1px solid #0078d4" : "1px solid #bbb",
                      backgroundColor: selectedTags.includes(tag) ? "#d9edff" : "#fff",
                      color: selectedTags.includes(tag) ? "#005a9e" : "#333",
                    }}>
                    #{tag}
                  </button>
                ))
              )}
            </div>

            {selectedTags.length > 0 && (
            <div style={{ marginTop: "10px" }}>
              <span style={{ marginRight: "8px", color: "#666" }}>
                選択中: {selectedTags.map((tag) => `#${tag}`).join(" / ")}
              </span>
              <button
                onClick={() => setSelectedTags([])}
                style={{
                  padding: "6px 10px",
                  cursor: "pointer",
                  borderRadius: "4px",
                  border: "1px solid #aaa",
                  backgroundColor: "#737272",
                }}
              >
                条件をクリア
              </button>
            </div>
          )}

          <div style={{ marginTop: "14px" }}>
            {/* 検索結果 */}
              <div style={{ fontWeight: "bold", marginBottom: "8px" }}>検索結果</div>
              {tagSearchResults.length === 0 ? (
                <div style={{ color: "#666" }}>タグに一致するファイルはありません</div>
              ) : (
                <ul style={{ margin: 0, paddingLeft: "20px", textAlign: "left" }}>
                 {tagSearchResults.map((result) => (
                    <li key={result.id} style={{ marginBottom: "8px" }}>
                     <span style={{ fontWeight: "bold" }}>{result.name}</span>
                     <span style={{ color: "#666", marginLeft: "8px" }}>
                       {result.tags.map((tag) => `#${tag}`).join(" ")}
                     </span>
                     <button
                        onClick={() => {
                          setTagModalFile({ id: result.id, name: result.name });
                       }}
                        style={{
                          marginLeft: "8px",
                         padding: "2px 8px",
                         cursor: "pointer",
                         borderRadius: "4px",
                          border: "1px solid #28a745",
                          backgroundColor: "#fff",
                          color: "#28a745",
                        }}>
                        タグ編集
                      </button>
                    </li>
                 ))}
                </ul>
              )}
            </div>

            {/* ファイル一覧 */}
            <FileList
             ref={fileListRef}
              itemId={currentFolderId}
              itemClick={handleItemClick}/>
             </main>
          </div>


      {/* ファイルクリック時のモーダル*/}
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
            <p style={{ wordBreak: "break-all" }}><strong>{actionModalFile.name}</strong></p>
            
            <div style={{ display: "flex", flexDirection: "column", gap: "10px", marginTop: "20px" }}>
              {/* 1. OneDriveで開く */}
              <button
                onClick={() => {
                  
                    window.open(actionModalFile.webUrl, "_blank", "noopener,noreferrer");
                  
                  setActionModalFile(null);
                }}
                style={{ padding: "8px", cursor: "pointer", backgroundColor: "#0078d4", color: "#fff", border: "none", borderRadius: "4px" }}
              >
                開く
              </button>

              {/* 2. タグを編集する */}
              <button
                onClick={() => {
                  setFileNameMap((prev) => ({ ...prev, [actionModalFile.id]: actionModalFile.name }));
                  setTagModalFile(actionModalFile);
                  setActionModalFile(null);
                }}
                style={{ padding: "8px", cursor: "pointer", backgroundColor: "#28a745", color: "#fff", border: "none", borderRadius: "4px" }}
              >
                タグを編集する
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

      {tagModalFile && (
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
            zIndex: 1001,
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
            <h3>タグを編集</h3>
            <p style={{ wordBreak: "break-all" }}><strong>{tagModalFile.name}</strong></p>

            <div>
              <strong>現在のタグ:</strong>
              <ul style={{ paddingLeft: "20px", textAlign: "left" }}>
                {(localTags[tagModalFile.id] || []).map((tag, idx) => (
                  <li key={idx} style={{ marginBottom: "5px" }}>
                    {tag}{" "}
                    <button
                      onClick={() => removeTag(tagModalFile.id, tag)}
                      style={{
                        marginLeft: "8px",
                        padding: "2px 6px",
                        cursor: "pointer",
                        backgroundColor: "#d01215",
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
              <strong>タグを追加:</strong><br />
              <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", marginTop: "8px", justifyContent: "center" }}>
                {quickAddTags.length === 0 ? (
                  <span style={{ color: "#000000" }}>候補はまだありません</span>
                ) : (
                  quickAddTags.map((tag) => (
                    <button
                      key={tag}
                      onClick={() => addTag(tagModalFile.id, tag)}
                      style={{
                        padding: "6px 10px",
                        cursor: "pointer",
                        borderRadius: "4px",
                        border: "1px solid #bbb",
                        backgroundColor: "#575555",
                      }}
                    >
                      + {tag}
                    </button>
                  ))
                )}
              </div>
            </div>

            <div style={{ marginTop: "15px" }}>
              <strong>新規タグを入力:</strong><br />
              <input
                value={newTagName}
                onChange={(e) => setNewTagName(e.target.value)}
                placeholder="例: 会議"
                style={{
                  padding: "8px",
                  borderRadius: "4px",
                  border: "1px solid #ccc",
                  marginTop: "8px",
                  minWidth: "180px",
                }}
              />
              <div style={{ marginTop: "10px", fontSize: "14px", color: "#333" }}>
                <label>
                  <input
                    type="checkbox"
                    checked={newTagAddToQuickAdd}
                    onChange={(e) => setNewTagAddToQuickAdd(e.target.checked)}
                  />
                  追加タグに追加する
                </label>
              </div>
              <button
                onClick={() => {
                  const trimmedTag = newTagName.trim();
                  if (!trimmedTag) return;
                  addTag(tagModalFile.id, trimmedTag);
                  if (newTagAddToQuickAdd) {
                    addQuickAddTag(trimmedTag);
                  }
                  setNewTagName("");
                  setNewTagAddToQuickAdd(false);
                }}
                style={{
                  padding: "8px 12px",
                  cursor: "pointer",
                  backgroundColor: "#0078d4",
                  color: "#fff",
                  border: "none",
                  borderRadius: "4px",
                  marginLeft: "8px",
                }}
              >
                追加
              </button>
            </div>

            <button
              onClick={() => {
                setNewTagName("");
                setTagModalFile(null);
              }}
              style={{ padding: "5px", cursor: "pointer", backgroundColor: "#ccc", border: "none", borderRadius: "4px", marginTop: "15px" }}
            >
              閉じる
            </button>
          </div>
        </div>
      )}
    </div>
  );
}


export default App;