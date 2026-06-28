// 模組層級嘅常數：純資料，同公司品牌無關，任何同類型 app 都用得到
// 公司專屬資料（logo、價錢、預設帳戶等）已經搬去 brand.js
import { BRAND_SLUG, OPEN_HOUR, CLOSE_HOUR } from "./brand.js";

export const LS_KEY = `${BRAND_SLUG}_data_v1`;
export const SESSION_KEY = `${BRAND_SLUG}_session_v1`;
export const CALSCALE_KEY = `${BRAND_SLUG}_calscale_v1`;
export const ADMIN_TAB_KEYS = ["overview", "schedule", "coaches", "ledger", "records", "settings"];
export const COLORS = ["#FF6B6B", "#4ECDC4", "#FFE66D", "#A8E6CF", "#C9B1FF", "#FFB347", "#FF8FA3", "#6BCB77"];

// 15-min grid，根據 brand.js 嘅營業時間自動生成
export const TIME_SLOTS = [];
for (let h = OPEN_HOUR; h < CLOSE_HOUR; h++) for (let m = 0; m < 60; m += 15) TIME_SLOTS.push(`${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`);
