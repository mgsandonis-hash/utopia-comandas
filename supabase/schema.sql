-- Utopía Comandas: base de datos en Supabase
-- Una sola tabla de documentos (colección + id + datos JSON), con acceso solo para los correos del equipo.

create table if not exists public.equipo (
  email text primary key
);

create table if not exists public.docs (
  col text not null,
  id text not null,
  data jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now(),
  primary key (col, id)
);
create index if not exists docs_col_idx on public.docs (col);

-- ¿El usuario conectado es del equipo?
create or replace function public.es_equipo() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.equipo where lower(email) = lower(coalesce(auth.jwt() ->> 'email', '')));
$$;

alter table public.equipo enable row level security;
alter table public.docs enable row level security;
alter table public.docs replica identity full;

drop policy if exists "equipo lee docs" on public.docs;
drop policy if exists "equipo crea docs" on public.docs;
drop policy if exists "equipo cambia docs" on public.docs;
drop policy if exists "equipo borra docs" on public.docs;
create policy "equipo lee docs" on public.docs for select to authenticated using (public.es_equipo());
create policy "equipo crea docs" on public.docs for insert to authenticated with check (public.es_equipo());
create policy "equipo cambia docs" on public.docs for update to authenticated using (public.es_equipo()) with check (public.es_equipo());
create policy "equipo borra docs" on public.docs for delete to authenticated using (public.es_equipo());

drop policy if exists "equipo se ve" on public.equipo;
create policy "equipo se ve" on public.equipo for select to authenticated using (public.es_equipo());

-- Mezcla profunda de JSON: los objetos se combinan, el resto se sustituye
create or replace function public.jsonb_deep_merge(a jsonb, b jsonb) returns jsonb
language plpgsql immutable as $$
declare k text; v jsonb; r jsonb;
begin
  if a is null or jsonb_typeof(a) <> 'object' or b is null or jsonb_typeof(b) <> 'object' then
    return b;
  end if;
  r := a;
  for k, v in select * from jsonb_each(b) loop
    if r ? k and jsonb_typeof(r -> k) = 'object' and jsonb_typeof(v) = 'object' then
      r := jsonb_set(r, array[k], public.jsonb_deep_merge(r -> k, v));
    else
      r := jsonb_set(r, array[k], v, true);
    end if;
  end loop;
  return r;
end $$;

-- Actualización atómica de una parte de un documento (evita pisar cambios del otro dispositivo)
create or replace function public.merge_doc(p_col text, p_id text, p_patch jsonb) returns jsonb
language plpgsql security invoker set search_path = public as $$
declare r jsonb;
begin
  update public.docs set data = public.jsonb_deep_merge(data, p_patch), updated_at = now()
   where col = p_col and id = p_id
   returning data into r;
  if r is null then
    raise exception 'no existe' using errcode = 'P0002';
  end if;
  return r;
end $$;

grant execute on function public.merge_doc(text, text, jsonb) to authenticated;
grant execute on function public.es_equipo() to authenticated;

-- Cambios en tiempo real entre el móvil y el iPad
do $$ begin
  alter publication supabase_realtime add table public.docs;
exception when duplicate_object then null; end $$;
