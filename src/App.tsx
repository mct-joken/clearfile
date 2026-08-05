import { Providers } from "@microsoft/mgt-element";
import { Msal2Provider } from "@microsoft/mgt-msal2-provider";

import { Login, FileList } from "@microsoft/mgt-react";

import { useEffect, useRef, useState } from "react";

import "./App.css";

// アプリの起動時に一度だけ実行
Providers.globalProvider = new Msal2Provider({
  clientId: "6c7057ca-aa6e-463d-b858-55c78248ce07",
  scopes: ["Files.Read", "Files.Read.All", "User.Read"], // OneDriveを読む許可をもらう
});

function App() {
  // 現在表示しているフォルダのIDを管理する（初期値は 'root'）
  const [currentFolderId, setCurrentFolderId] = useState<string>("root");
  // 過去に移動したフォルダIDの履歴（「戻る」ボタン用）
  const [folderHistory, setFolderHistory] = useState<string[]>([]);
  const fileListRef = useRef<any>(null);

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
    // 💡 リダイレクト（OneDriveのページが開くの）を阻止する
    e?.preventDefault?.();

    const clickedItem = e?.detail?.item ?? e?.detail ?? e;

    if (!clickedItem) {
      return;
    }

    // クリックされたのが「フォルダ」だった場合のみ、その中身に遷移する
    if (clickedItem.folder) {
      const nextFolderId = clickedItem.id;

      // 履歴に現在のフォルダを追加し、新しいフォルダIDに切り替える
      setFolderHistory((prev) => [...prev, currentFolderId]);
      setCurrentFolderId(nextFolderId);
    } else {
      // ファイルだった場合の処理（後回しでOKですが、確認用にログを出します）
      console.log("ファイルがクリックされました:", clickedItem.name);
    }
  };

  // 「前のフォルダに戻る」ボタンの処理
  const handleBackClick = () => {
    if (folderHistory.length === 0) return;

    // 履歴の最後から1つ取り出す
    const previousFolderId = folderHistory[folderHistory.length - 1];

    setFolderHistory((prev) => prev.slice(0, -1)); // 最後の要素を履歴から消す
    setCurrentFolderId(previousFolderId); // フォルダIDを戻す
  };

  return (
    <div>
      <h1>マイ OneDrive</h1>
      <Login />

      <hr />

      {/* ルート（最上位）ではない時だけ「戻る」ボタンを表示する */}
      {currentFolderId !== "root" && (
        <button
          onClick={handleBackClick}
          style={{
            marginBottom: "10px",
            padding: "5px 10px",
            cursor: "pointer",
          }}
        >
          ⬅ 前のフォルダに戻る
        </button>
      )}

      {/* itemId: 指定したIDのフォルダ内を表示（デフォルトは root）
        itemClick: クリックされた時のイベントを横取りする
      */}
      <FileList
        ref={fileListRef}
        itemId={currentFolderId}
        itemClick={handleItemClick}
      />
    </div>
  );
}

export default App;
