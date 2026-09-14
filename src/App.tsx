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

  // タグの削除処理 
  const removeTag = (fileId: string, tag: string) => { 
    setLocalTags((prev) => { 
      const currentFileTags = prev[fileId] || []; 
      return { 
        ...prev, 
        [fileId]: currentFileTags.filter((item) => item !== tag), 
      }; 
    }); 
  }; 

  // タグリスト
  const tagsCounts = Object.values(localTags).flat().reduce((acc, tag) => {acc[tag] = (acc[tag] || 0) + 1; return acc;}, {} as Record<string, number>);

  return ( 
    <div style={{ display: "flex", gap: "20px",  }}> 
      {/* ヘッダー */} 
      <div style={{ flex: 2 }}> 
        <header className="header"> 
          <h1>マイ OneDrive</h1> 
          <div className=""></div> 
          <Login /> 
        </header> 
        <hr /> 
      
      {/* メインコンテンツ領域 */}
      <div style={{ display: "flex", gap: "20px", alignItems: "flex-start",}}>

        {/*  設定したタグ一覧 */}
        <div style={{ width: "180px", border: "1px solid #ccc", borderRadius: "10px", padding: "16px" }}>
          <h3 style={{ marginTop: 0, textAlign: "center" }}>タグ一覧</h3>
          <ul style={{ listStyle: "none", padding: 0, margin: 0 }}>
            {Object.entries(tagsCounts).map(([tag, count], idx) => (
              <li key={idx} style={{ marginBottom: "8px" }}>
                #{tag}({count})
              </li>
            ))}
          </ul>
        </div>

        {/* 右側: ファイル一覧 */}
        <div style={{ flex: 1 }}>
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
      </div>
      </div>

      {/* タグ編集選択時: タグ編集パネル */} 
      {selectedFile && (
        <div style={{ position: "fixed", top: "50%", left: "50%", transform: "translate(-50%, -50%)", width: "90%", maxWidth: "400px", backgroundColor: "#fff", padding: "24px", borderRadius: "12px", boxShadow: "0 10px 25px rgba(0, 0, 0, 0.3)", zIndex: 9999, boxSizing: "border-box" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <h2>タグ編集パネル</h2>
          </div>

          <p> 
            <strong>選択中:</strong> <br />
            <span>{selectedFile.name}</span>
          </p> 

          <div> 
            <strong>現在のタグ:</strong> 
            <ul style={{ paddingLeft: "20px" }}> 
              {(localTags[selectedFile.id] || []).map((tag, idx) => ( 
                <li key={idx} style={{ marginBottom: "5px" }}> 
                  {tag}{" "} 
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

          <button onClick={() => setSelectedFile(null)}>
            完了
          </button>
        </div> 
      )}

      {/* ファイルクリック時に表示される選択ダイアログ モーダル */} 
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
                gap: "10px", 
                marginTop: "20px", 
              }} 
            > 
              <button 
                onClick={() => { 
                  if (actionModalFile.webUrl) { 
                    window.open( 
                      actionModalFile.webUrl, 
                      "_blank", 
                      "noopener,noreferrer" 
                    ); 
                  } 
                  setActionModalFile(null); 
                }} 
                style={{ 
                  padding: "8px", 
                  cursor: "pointer", 
                  backgroundColor: "#0078d4", 
                  color: "#fff", 
                  border: "none", 
                  borderRadius: "4px", 
                }} 
              > 
                開く 
              </button> 

              <button 
                onClick={() => { 
                  setSelectedFile(actionModalFile); 
                  setActionModalFile(null); 
                }} 
                style={{ 
                  padding: "8px", 
                  cursor: "pointer", 
                  backgroundColor: "#28a745", 
                  color: "#fff", 
                  border: "none", 
                  borderRadius: "4px", 
                }} 
              > 
                タグを編集 
              </button> 

              <button 
                onClick={() => setActionModalFile(null)} 
                style={{ 
                  padding: "5px", 
                  cursor: "pointer", 
                  backgroundColor: "#ccc", 
                  border: "none", 
                  borderRadius: "4px", 
                  marginTop: "5px", 
                }} 
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