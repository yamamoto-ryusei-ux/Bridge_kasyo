# Tauriアプリ統合手順書 — KENBAN → COMIC-Bridge タブ移植

別のTauriアプリ（React+Rust）を、既存のTauriアプリの新規タブとして丸ごと移植する際の手順と注意点をまとめたもの。

---

## 前提

- **ベースアプリ（COMIC-Bridge）**: Tauri 2 + React 18 + TypeScript + Tailwind CSS 3 + Zustand + Vite
- **移植元アプリ（KENBAN）**: Tauri 2 + React 19 + TypeScript + Tailwind CSS 4 + Vite
- ベースアプリはタブベースのナビゲーション（ViewRouter + viewStore）を持つ
- 移植元アプリは単独で動作する全画面アプリ（App.tsx に全状態管理）

---

## 手順

### 1. ベースアプリをコピー

```
ベースアプリのリポジトリをクローン → 作業フォルダにコピー
```

### 2. 移植元のフロントエンドファイルをコピー

移植元の構造をベースアプリ内に分離配置する：

| 移植元パス | コピー先 |
|-----------|---------|
| `src/App.tsx` | `src/components/kenban/KenbanApp.tsx` |
| `src/components/*.tsx` | `src/components/kenban/` （ファイル名にプレフィックス付与で衝突回避） |
| `src/utils/*.ts` | `src/kenban-utils/` |
| `src/hooks/*.ts` | `src/kenban-hooks/` |
| `src/workers/*.ts` | `src/kenban-workers/` |
| `src/types.ts` | `src/kenban-utils/kenbanTypes.ts` |
| `src/assets/*` | `src/kenban-assets/` |
| `src/index.css`, `src/App.css` | `src/kenban-utils/kenban.css`, `kenbanApp.css` |
| `public/pdfjs-wasm/` | `public/pdfjs-wasm/` |

### 3. import パスを全修正

コピーしたファイル内のimportパスを新しいディレクトリ構造に合わせてsedで一括置換：

```bash
# 例: KenbanApp.tsx 内
sed -i "s|from './utils/pdf'|from '../../kenban-utils/pdf'|g" KenbanApp.tsx
sed -i "s|from './components/Header'|from './KenbanHeader'|g" KenbanApp.tsx
sed -i "s|from './types'|from '../../kenban-utils/kenbanTypes'|g" KenbanApp.tsx
# ... 他のファイルも同様
```

**見落としやすいポイント:**
- `import('./types')` 形式のインライン型import（sedの通常パターンでは引っかからない）
- `new URL('../workers/...', import.meta.url)` 形式のWeb Worker URL（Viteビルド時に解決される）
- アセットimport（`import logo from '../assets/logo.png'`）

### 4. ビュー・タブの作成

#### 4-1. viewStore.ts に新しいタブIDを追加

```typescript
export type AppView =
  | "specCheck" | "layers" | ...
  | "kenban";  // 追加
```

#### 4-2. KenbanView.tsx（ラッパーコンポーネント）を作成

```tsx
import KenbanApp from "../kenban/KenbanApp";
import "../../kenban-utils/kenban.css";
import "../../kenban-utils/kenbanApp.css";

export function KenbanView() {
  return (
    <div className="flex h-full w-full overflow-hidden kenban-scope"
         style={{ position: 'absolute', inset: 0 }}>
      <KenbanApp />
    </div>
  );
}
```

#### 4-3. ViewRouter.tsx で状態保持型のルーティング

**重要**: 通常の条件付きレンダリング（`{active === "kenban" && <KenbanView />}`）ではタブ切替時にアンマウント→再マウントされ、全状態が失われる。

```tsx
// 一度マウントしたらアンマウントしない方式
const [kenbanMounted, setKenbanMounted] = useState(false);

useEffect(() => {
  if (activeView === "kenban") setKenbanMounted(true);
}, [activeView]);

return (
  <div className="flex-1 overflow-hidden relative">
    {/* 他のタブは通常の条件付きレンダリング */}
    {activeView === "specCheck" && <SpecCheckView />}

    {/* KenbanViewはdisplay切替で状態保持 */}
    {kenbanMounted && (
      <div style={{ display: activeView === "kenban" ? "contents" : "none" }}>
        <KenbanView />
      </div>
    )}
  </div>
);
```

#### 4-4. TopNav.tsx にタブボタンを追加

VIEW_TABS配列に新しいタブエントリを追加。

### 5. 移植元App.tsxの修正

#### 5-1. レイアウト修正

```diff
- <div className="h-screen flex flex-col ...">
+ <div className="h-full w-full flex flex-col ...">
```

`h-screen`はタブ内ではオーバーフローするため`h-full`に変更。

#### 5-2. onCloseRequested を無効化

単独アプリでは`getCurrentWebviewWindow().onCloseRequested()`でウィンドウ閉じ時の未保存確認を行っていたが、タブとして埋め込まれるとウィンドウ全体の×ボタンをブロックしてしまう。削除またはコメントアウト。

#### 5-3. 自動更新チェックの無効化（任意）

移植元のupdater `check()` 呼び出しはベースアプリ側のupdaterと競合する可能性がある。タブとして動作する場合は不要。

### 6. Rustバックエンドの統合

#### 6-1. 移植元の lib.rs → kenban.rs モジュールとして作成

- `pub fn run()` を削除（ベースアプリが独自に持つ）
- `AppState` を `KenbanState` にリネーム（名前衝突回避）
- 構造体名の衝突を回避（例: `CropBounds` → `KenbanCropBounds`）
- 全 `#[tauri::command]` 関数を `pub` にする
- ベースアプリと重複するコマンドは `kenban_` プレフィックスを付ける

**リネームが必要なコマンド例:**
| 元の名前 | リネーム後 | 理由 |
|---------|-----------|------|
| `parse_psd` | `kenban_parse_psd` | ベースに同名コマンドあり |
| `list_files_in_folder` | `kenban_list_files_in_folder` | ベースに類似コマンドあり |
| `render_pdf_page` | `kenban_render_pdf_page` | ベースに同名コマンドあり |
| `open_file_in_photoshop` | `kenban_open_file_in_photoshop` | 引数が異なる |
| `save_screenshot` | `kenban_save_screenshot` | ベースにない独自機能 |

重複しないコマンド（`compute_diff_simple`等）はそのまま使用可能。

#### 6-2. lib.rs にモジュール登録

```rust
pub mod kenban;

// Builder に追加
.manage(kenban::KenbanState { ... })
.invoke_handler(tauri::generate_handler![
    // 既存コマンド...
    kenban::kenban_parse_psd,
    kenban::compute_diff_simple,
    // ...
])
```

#### 6-3. Cargo.toml に依存関係を追加

移植元にあってベースにない依存関係のみ追加。既存の依存関係はfeatureフラグの確認だけ行う。

#### 6-4. pdfium.dll の探索パス統一

移植元が `exe隣のみ検索` だった場合、ベースアプリのパターン（`CARGO_MANIFEST_DIR/resources/pdfium/` → exe隣 → `resources/pdfium/` → システム）に統一する。dev環境ではexe隣にdllがないため。

### 7. フロントエンドのinvokeコマンド名を修正

**最も見落としやすいステップ。** リネームしたRustコマンド名に合わせて、フロントエンドの `invoke('...')` 呼び出しを全て修正する。

```bash
# 検索して見つける
grep -rn "invoke(" src/components/kenban/ | grep "'parse_psd'"
grep -rn "invoke(" src/components/kenban/ | grep "'list_files_in_folder'"
# ...

# 一括置換
sed -i "s|invoke('parse_psd'|invoke('kenban_parse_psd'|g" KenbanApp.tsx
sed -i "s|invoke<string\[\]>('list_files_in_folder'|invoke<string[]>('kenban_list_files_in_folder'|g" KenbanApp.tsx
```

**注意:** `invoke<Type>('command_name'` のようにジェネリクス付きの呼び出しもあるため、sed パターンは複数必要。

### 8. CSS/テーマの統合

#### Tailwind バージョンが異なる場合（v4 → v3）

移植元のTailwind v4 `@theme` ブロックをCSS変数に変換し、`.kenban-scope`でスコープ化：

```css
/* Tailwind v4の@themeブロックを変換 */
.kenban-scope {
  --color-surface-base: #0e0e10;
  --color-text-primary: #ececf0;
  background-color: #0e0e10;
  color: #ececf0;
}
```

標準のTailwindユーティリティクラス（`bg-neutral-900`等）はv3でもそのまま動作する。

#### `@import "tailwindcss"` の削除

Tailwind v4の`@import "tailwindcss"`はv3では不要（globals.cssで`@tailwind`ディレクティブが既にある）。

### 9. package.json 依存関係のマージ

移植元にあってベースにない依存関係のみ追加。**Reactバージョンはベース側に合わせる**（ダウングレード）。

### 10. TypeScript型の互換性対応

React 19 → 18 ダウングレード時の型非互換性：

- `useRef<HTMLDivElement>(null)` の戻り値型が異なる（React 19: `RefObject<HTMLDivElement | null>`, React 18: `MutableRefObject<HTMLDivElement | null>`）
- 解決策: 移植元コンポーネントに `// @ts-nocheck` を追加

その他:
- PNG/SVGアセットの型宣言ファイル（`assets.d.ts`）を作成
- UTIFなどの型宣言ファイル（`utif.d.ts`）を作成

### 11. Tauri設定の更新

#### tauri.conf.json

- **CSP**: `worker-src blob:`, `script-src 'unsafe-eval' blob:` を追加（pdfjs-dist等のWeb Worker用）
- **フォント**: Google Fontsを使う場合は `style-src`と`font-src`にドメインを追加

#### capabilities/default.json

- 移植元が使うプラグイン権限を確認し、不足分を追加
- **注意**: ベースアプリにインストールされていないプラグインの権限を追加するとビルドエラーになる（例: `opener:default` はプラグインなしでは使えない）

#### index.html

- Google Fontsの`<link>`タグを追加（移植元がWebフォントを使う場合）

### 12. .gitignore の作成

```
node_modules/
dist/
src-tauri/target/
src-tauri/gen/
src-tauri/resources/pdfium/pdfium.dll
package-lock.json
src-tauri/Cargo.lock
```

---

## よくある問題と対処法

### ビルドは通るがアプリが起動しない
- `pdfium.dll` が配置されているか確認（`src-tauri/resources/pdfium/`）
- capabilities に未インストールプラグインの権限がないか確認

### タブ切替で状態が消える
- ViewRouterで条件付きレンダリング（`&&`）ではなく`display: none/contents`切替にする

### ×ボタン（ウィンドウクローズ）が効かない
- 移植元の`onCloseRequested` + `event.preventDefault()`がウィンドウ全体のクローズをブロックしている。無効化する。

### invokeコマンドが見つからない
- Rust側でリネームしたコマンド名とフロントエンドの`invoke()`呼び出し名が一致しているか全件確認
- `invoke<Type>('name'` のようなジェネリクス付きパターンも見落とさない
- インライン`import('./types')` 形式のimportも忘れがち

### PDF が表示されない
- Rust側のpdfium.dll探索パスがdev環境（`CARGO_MANIFEST_DIR`）を含んでいるか確認
- CSPに`worker-src blob:`, `script-src blob:`があるか確認

### React 18/19 型エラー
- `useRef` の戻り値型が異なるため、移植元ファイルに `// @ts-nocheck` を追加
- ビルドコマンドが`tsc && vite build`の場合、tscが通らなければビルド失敗するため必須

### Web Worker パスエラー
- `new URL('../workers/file.ts', import.meta.url)` のパスをコピー先に合わせて修正
- Viteビルド時に解決されるため、tscは通ってもビルドで失敗する

---

## チェックリスト

- [ ] フロントエンドファイルのコピーとimportパス修正
- [ ] KenbanView ラッパーコンポーネント作成
- [ ] viewStore にタブID追加
- [ ] ViewRouter に状態保持型ルーティング追加
- [ ] TopNav にタブボタン追加
- [ ] KenbanApp.tsx の `h-screen` → `h-full` 修正
- [ ] `onCloseRequested` の無効化
- [ ] Rust kenban.rs モジュール作成（コマンドリネーム含む）
- [ ] lib.rs にモジュール登録・state管理・コマンド登録
- [ ] Cargo.toml 依存関係マージ
- [ ] pdfium.dll 探索パス統一
- [ ] フロントエンドのinvokeコマンド名を全件修正
- [ ] CSS スコープ化（`.kenban-scope`）
- [ ] package.json 依存関係マージ（Reactバージョンは変えない）
- [ ] TypeScript型対応（`@ts-nocheck`、型宣言ファイル）
- [ ] tauri.conf.json CSP更新
- [ ] capabilities 権限確認
- [ ] index.html フォント追加
- [ ] .gitignore 作成
- [ ] `npm run build` でビルド確認
- [ ] `npm run tauri dev` で起動確認
