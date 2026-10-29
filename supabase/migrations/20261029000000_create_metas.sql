-- Tabela de metas compartilhadas do "Relatório do cliente"
-- Execute no Supabase: SQL Editor > New query > cole este arquivo > Run
-- (ou use `supabase db push` se estiver usando o CLI)

create table if not exists public.metas (
  id          uuid primary key default gen_random_uuid(),
  nome        text not null,
  tipo        text not null check (tipo in ('rec', 'res', 'desp')),
  valor       numeric not null check (valor > 0),
  de          date not null,
  ate         date not null,
  criado_por  text default (auth.jwt() ->> 'email'),
  created_at  timestamptz not null default now(),
  check (de <= ate)
);

alter table public.metas enable row level security;

-- Mesma lista de e-mails autorizados da Edge Function fetch-google-sheets.
-- Para liberar outra pessoa, inclua o e-mail aqui E na Edge Function.
drop policy if exists "metas_select" on public.metas;
drop policy if exists "metas_insert" on public.metas;
drop policy if exists "metas_update" on public.metas;
drop policy if exists "metas_delete" on public.metas;

create policy "metas_select" on public.metas for select to authenticated
  using (lower(auth.jwt() ->> 'email') in ('contabilidade@dawadm.net', 'fiscal1@dawadm.net', 'maicon@dawadm.net'));

create policy "metas_insert" on public.metas for insert to authenticated
  with check (lower(auth.jwt() ->> 'email') in ('contabilidade@dawadm.net', 'fiscal1@dawadm.net', 'maicon@dawadm.net'));

create policy "metas_update" on public.metas for update to authenticated
  using (lower(auth.jwt() ->> 'email') in ('contabilidade@dawadm.net', 'fiscal1@dawadm.net', 'maicon@dawadm.net'))
  with check (lower(auth.jwt() ->> 'email') in ('contabilidade@dawadm.net', 'fiscal1@dawadm.net', 'maicon@dawadm.net'));

create policy "metas_delete" on public.metas for delete to authenticated
  using (lower(auth.jwt() ->> 'email') in ('contabilidade@dawadm.net', 'fiscal1@dawadm.net', 'maicon@dawadm.net'));
