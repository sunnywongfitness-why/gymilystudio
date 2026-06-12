// ====== Supabase 設定 ======
// 喺 Supabase 開咗 project 之後，去 Project Settings → API，
// 將 Project URL 同 anon public key 填落去下面。
//
// 兩個都填咗，app 就會自動用雲端同步（所有裝置共用同一份資料）。
// 留空嘅話，app 會自動改用本機儲存（localStorage，唔會跨裝置同步）。

export const SUPABASE_URL = "";       // 例如 "https://abcdxyz.supabase.co"
export const SUPABASE_ANON_KEY = "";  // 例如 "eyJhbGciOi....（好長一串）"
