# Supabase 雲端同步設定（多裝置共用同一份資料）

跟住做，所有教練／管理員無論用邊部機，都會睇到**同一份即時資料**。
全部免費（Supabase 有免費 tier，呢類小型用途用唔晒）。

---

## 1. 開 Supabase project

1. 去 https://supabase.com → Sign in（用 GitHub 登入最方便）
2. 撳 **New project**
3. 填 project name（例如 `gymily`）、set 一個 database password（自己記住，呢個唔使填落 app）
4. 揀近你嘅 region（例如 Singapore / Hong Kong）
5. 撳 Create，等 1–2 分鐘佢開好

## 2. 建立資料表（貼 SQL）

1. 左邊揀 **SQL Editor** → **New query**
2. 將下面成段貼入去，撳 **Run**：

```sql
-- 主資料表（成個 app 嘅資料存喺一行 JSON）
create table if not exists public.app_state (
  id text primary key,
  data jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

-- 開 Row Level Security，容許讀寫（app 內部自己有登入層）
alter table public.app_state enable row level security;
create policy "allow read"   on public.app_state for select using (true);
create policy "allow insert" on public.app_state for insert with check (true);
create policy "allow update" on public.app_state for update using (true) with check (true);

-- 即時同步：將表加入 realtime，並令更新帶完整資料
alter table public.app_state replica identity full;
alter publication supabase_realtime add table public.app_state;
```

> 如果最後一行出現「already member of publication」之類嘅訊息，唔緊要，代表已經加咗。

## 3. 攞 URL 同 anon key

1. 左下角 **Project Settings**（齒輪）→ **API**
2. 喺呢頁搵：
   - **Project URL**（例如 `https://abcdxyz.supabase.co`）
   - **Project API keys** 下面嘅 **anon / public** key（好長一串）

## 4. 填落 app

打開 `src/config.js`，將兩個值填入：

```js
export const SUPABASE_URL = "https://abcdxyz.supabase.co";
export const SUPABASE_ANON_KEY = "eyJhbGciOi....（你抄低嗰串）";
```

- 喺 GitHub 直接改 `src/config.js`（撳鉛筆 icon 編輯）再 Commit，Vercel 會自動重新 deploy。
- 填好之後，app 右上角會顯示「☁️ 已同步」。留空就會顯示「📱 本機」。

## 5. 完成

之後任何裝置開個網址、登入，睇到嘅都係同一份資料；一部機改咗，其他機幾秒內自動更新。

---

## 常見問題

**右上角顯示「⚠️ 同步失敗」？**
通常係 SQL 未行成功，或 URL/key 填錯。返 SQL Editor 確認三段都 Run 過，再核對 `config.js` 兩個值。

**安全性點？**
anon key 係設計上公開嘅（會喺前端出現）。呢個版本為咗簡單，任何有網址嘅人理論上可以讀寫資料表。對小型私人工作室一般夠用，但如果你好注重保安，可以稍後再升級做「真正帳號驗證（Supabase Auth）+ 更嚴格 RLS」。需要嘅話再搵我。

**會唔會撞數據（兩部機同時改）？**
採用「最後寫入為準」。對單一場地、幾個教練嚟講好少會啱啱同一秒改同一格，影響好微。

**想清空所有資料？**
管理員 → 設定 → 重設所有資料（會同步清空雲端）。或喺 Supabase Table Editor 直接刪 `app_state` 嗰行。
