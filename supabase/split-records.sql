-- ===== Gymily：拆分資料（2026-11）=====
-- 以前：所有資料放喺 app_state 一行（id = 'main'），每次改動成份覆蓋。
-- 而家：app_records 每筆記錄一行；app_state.main 改為由資料庫自動砌出嚟嘅「快照」（俾日曆 feed 同舊版 app 讀），app 唔會再直接寫佢。
--
-- 分兩部分執行：
--   PART A：建表、權限、快照函數（隨時可以行，唔影響現有運作）
--   PART B：切換（鎖住 app_state 寫入 + 安裝 trigger）——搬完資料之後先行

-- ================= PART A =================
create table if not exists public.app_records (
  collection text not null,          -- "s" 設定值、"k" 登記、"l:<key>" list 項目、"d:<key>" dict entry
  rid        text not null,          -- 記錄 id
  data       jsonb,
  updated_at timestamptz not null default now(),
  primary key (collection, rid)
);
alter table public.app_records enable row level security;
drop policy if exists "app_records read" on public.app_records;
drop policy if exists "app_records insert" on public.app_records;
drop policy if exists "app_records update" on public.app_records;
drop policy if exists "app_records delete" on public.app_records;
create policy "app_records read"   on public.app_records for select using (true);
create policy "app_records insert" on public.app_records for insert with check (true);
create policy "app_records update" on public.app_records for update using (true) with check (true);
create policy "app_records delete" on public.app_records for delete using (true);
grant select, insert, update, delete on public.app_records to anon, authenticated;

-- updated_at 一律由資料庫蓋（唔信裝置時鐘），app 用佢做「最近有冇改動」嘅增量檢查
create or replace function public.app_records_touch() returns trigger
language plpgsql as $$ begin new.updated_at := now(); return new; end $$;
drop trigger if exists app_records_touch on public.app_records;
create trigger app_records_touch before insert or update on public.app_records
  for each row execute function public.app_records_touch();
create index if not exists app_records_updated_at on public.app_records (updated_at);

-- 每日自動備份（以前放喺 app_state 嘅 backup-0…6，而家獨立一個表）
create table if not exists public.app_backups (
  id         text primary key,
  data       jsonb not null,
  updated_at timestamptz not null default now()
);
alter table public.app_backups enable row level security;
drop policy if exists "app_backups read" on public.app_backups;
drop policy if exists "app_backups insert" on public.app_backups;
drop policy if exists "app_backups update" on public.app_backups;
create policy "app_backups read"   on public.app_backups for select using (true);
create policy "app_backups insert" on public.app_backups for insert with check (true);
create policy "app_backups update" on public.app_backups for update using (true) with check (true);
grant select, insert, update on public.app_backups to anon, authenticated;

-- 由一行行記錄砌返成份資料（同 app 入面 recordSync.implode 一樣嘅規則）
create or replace function public.gymily_snapshot() returns jsonb
language sql stable set search_path = public as $$
  with reg as (
    select rid as key, data from app_records where collection = 'k'
  ),
  lists as (
    select reg.key,
      coalesce((
        select jsonb_agg(r.data order by (o.ord is not null), o.ord, r.rid collate "C")
        from app_records r
        left join jsonb_array_elements_text(reg.data->'order') with ordinality as o(rid, ord) on o.rid = r.rid
        where r.collection = 'l:' || reg.key
      ), '[]'::jsonb) as val
    from reg where reg.data->>'kind' = 'list'
  ),
  dicts as (
    select reg.key,
      coalesce((select jsonb_object_agg(r.rid, r.data) from app_records r where r.collection = 'd:' || reg.key), '{}'::jsonb) as val
    from reg where reg.data->>'kind' = 'dict'
  ),
  sets as (
    select rid as key, data as val from app_records where collection = 's'
  )
  select coalesce(jsonb_object_agg(key, val), '{}'::jsonb)
  from (select * from sets union all select * from lists union all select * from dicts) x
$$;

create or replace function public.gymily_refresh_snapshot() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  -- 有登記記錄先更新快照（防止空表時將 main 洗成空白）
  if exists (select 1 from app_records where collection = 'k') then
    update app_state set data = gymily_snapshot(), updated_at = now() where id = 'main';
  end if;
  return null;
end $$;

-- ================= PART B（切換時先行）=================
-- 1) 舊版 app 唔可以再寫 app_state（仍然可以讀，所以未更新嘅裝置都見到最新資料）
-- revoke insert, update, delete, truncate on public.app_state from anon, authenticated;
-- 2) 搬資料（由 app 用同一套 explode 規則寫入 app_records）
-- 3) 安裝 trigger：之後 app_records 每次改動，app_state.main 自動更新做快照
-- drop trigger if exists gymily_snapshot_trigger on public.app_records;
-- create trigger gymily_snapshot_trigger after insert or update or delete on public.app_records
--   for each statement execute function public.gymily_refresh_snapshot();
-- 4) 實時同步（已經喺測試時加咗）
-- alter publication supabase_realtime add table public.app_records;
