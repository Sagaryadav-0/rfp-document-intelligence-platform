-- Create a table for RFP session metadata and enable row-level security.
-- This table is intended to isolate records per browser session or per authenticated user.

create table if not exists public.rfp_sessions (
  id uuid primary key default gen_random_uuid(),
  session_id text not null,
  file_name text,
  file_size bigint,
  status text not null default 'created',
  metadata jsonb default '{}'::jsonb,
  created_at timestamptz not null default now()
);

alter table public.rfp_sessions enable row level security;

create policy "Allow access by session_id" on public.rfp_sessions
  for select using (session_id = current_setting('request.session_id', true));

create policy "Allow insert by session" on public.rfp_sessions
  for insert with check (session_id = current_setting('request.session_id', true));

-- If you later add Supabase Auth, switch to auth.uid() based rules like:
-- create policy "Allow authenticated users" on public.rfp_sessions
--   for select using (auth.uid() = current_setting('request.user_id', true));

-- Create a private storage bucket for RFP files
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('rfp-files', 'rfp-files', false, 26214400, array['application/pdf']);

-- Enable RLS on storage.objects for the bucket
alter table storage.objects enable row level security;

-- Policy: Allow upload to /session_id/ paths
create policy "Allow upload by session" on storage.objects
  for insert with check (
    bucket_id = 'rfp-files' and
    (storage.foldername(name))[1] = current_setting('request.session_id', true)
  );

-- Policy: Allow read (signed URLs) by session
create policy "Allow read by session" on storage.objects
  for select using (
    bucket_id = 'rfp-files' and
    (storage.foldername(name))[1] = current_setting('request.session_id', true)
  );
