// ====== Supabase 設定 ======
// 喺 Supabase 開咗 project 之後，去 Project Settings → API，
// 將 Project URL 同 anon public key 填落去下面。
//
// 兩個都填咗，app 就會自動用雲端同步（所有裝置共用同一份資料）。
// 留空嘅話，app 會自動改用本機儲存（localStorage，唔會跨裝置同步）。

export const SUPABASE_URL = "https://eztpnmadqbxiznzgvlzu.supabase.co";
export const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImV6dHBubWFkcWJ4aXpuemd2bHp1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODEyMzI2MTMsImV4cCI6MjA5NjgwODYxM30.XnN6hbuEoPGh-CEr-kOLHqIrzD7qjpIt70e6H_vL1FM";
