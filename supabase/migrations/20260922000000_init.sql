-- QG Real Estate — schema iniziale
--
-- L'app salva l'intero workspace (liquidità, operazioni, immobili a reddito, bilancio) come
-- un unico documento JSON, esattamente come faceva l'artifact. Il formato coincide con il
-- backup "Esporta/Importa", quindi i dati esistenti si caricano con il pulsante Importa.
--
-- Sicurezza: i dati contengono saldi bancari e nomi degli investitori. Possono leggerli e
-- modificarli SOLO gli utenti autenticati la cui email è in `team_members`.

-- ---------------------------------------------------------------------------
-- Team autorizzato
-- ---------------------------------------------------------------------------
create table if not exists public.team_members (
  email      text primary key,
  nome       text,
  created_at timestamptz not null default now()
);

alter table public.team_members enable row level security;
-- Nessuna policy: la tabella si gestisce solo dal SQL editor / service role.
-- Esempio (sostituisci con le email reali dei soci):
--   insert into public.team_members (email, nome) values
--     ('marco@qgrealestate.it', 'Marco');

create or replace function public.is_team_member()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.team_members
    where lower(email) = lower(coalesce(auth.jwt() ->> 'email', ''))
  );
$$;

revoke all on function public.is_team_member() from public;
grant execute on function public.is_team_member() to authenticated;

-- ---------------------------------------------------------------------------
-- Stato dell'applicazione
-- ---------------------------------------------------------------------------
create table if not exists public.app_state (
  key        text primary key,
  data       jsonb not null,
  version    bigint not null default 1,
  updated_at timestamptz not null default now(),
  updated_by uuid references auth.users (id)
);

alter table public.app_state enable row level security;

create policy "team legge app_state" on public.app_state
  for select to authenticated using (public.is_team_member());
create policy "team inserisce app_state" on public.app_state
  for insert to authenticated with check (public.is_team_member());
create policy "team aggiorna app_state" on public.app_state
  for update to authenticated using (public.is_team_member()) with check (public.is_team_member());
-- Nessuna policy di delete: i dati non si cancellano dal client.

-- ---------------------------------------------------------------------------
-- Storico: ogni salvataggio conserva la versione precedente.
-- Serve a recuperare i dati dopo un errore (es. un'importazione sbagliata).
-- ---------------------------------------------------------------------------
create table if not exists public.app_state_history (
  id         bigint generated always as identity primary key,
  key        text not null,
  data       jsonb not null,
  version    bigint not null,
  updated_at timestamptz not null,
  updated_by uuid,
  archived_at timestamptz not null default now()
);

create index if not exists app_state_history_key_idx on public.app_state_history (key, archived_at desc);

alter table public.app_state_history enable row level security;
create policy "team legge storico" on public.app_state_history
  for select to authenticated using (public.is_team_member());

create or replace function public.archive_app_state()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.app_state_history (key, data, version, updated_at, updated_by)
  values (old.key, old.data, old.version, old.updated_at, old.updated_by);
  return new;
end;
$$;

drop trigger if exists app_state_archive on public.app_state;
create trigger app_state_archive
  before update on public.app_state
  for each row
  when (old.data is distinct from new.data)
  execute function public.archive_app_state();

-- ---------------------------------------------------------------------------
-- Salvataggio con controllo di versione (optimistic locking).
-- Il client passa la versione che ha letto: se nel frattempo un altro socio ha salvato,
-- il salvataggio viene rifiutato con CONFLICT invece di sovrascrivere le sue modifiche.
-- ---------------------------------------------------------------------------
create or replace function public.save_app_state(p_key text, p_data jsonb, p_expected_version bigint)
returns bigint
language plpgsql
security invoker
set search_path = public
as $$
declare
  v_current record;
  v_new bigint;
begin
  select version, data into v_current from public.app_state where key = p_key;

  if not found then
    if coalesce(p_expected_version, 0) <> 0 then
      raise exception 'CONFLICT' using errcode = '40001';
    end if;
    insert into public.app_state (key, data, version, updated_by)
    values (p_key, p_data, 1, auth.uid())
    on conflict (key) do nothing
    returning version into v_new;
    if v_new is null then
      raise exception 'CONFLICT' using errcode = '40001';
    end if;
    return v_new;
  end if;

  if v_current.version <> coalesce(p_expected_version, 0) then
    raise exception 'CONFLICT' using errcode = '40001';
  end if;

  -- Nessuna modifica reale: niente nuova versione né voce di storico.
  if v_current.data = p_data then
    return v_current.version;
  end if;

  update public.app_state
     set data = p_data, version = version + 1, updated_at = now(), updated_by = auth.uid()
   where key = p_key and version = p_expected_version
  returning version into v_new;

  if v_new is null then
    raise exception 'CONFLICT' using errcode = '40001';
  end if;
  return v_new;
end;
$$;

revoke all on function public.save_app_state(text, jsonb, bigint) from public;
grant execute on function public.save_app_state(text, jsonb, bigint) to authenticated;
