-- Correos que pueden entrar en la app de Utopía
insert into public.equipo (email) values
  ('mgsandonis@gmail.com'),
  ('lauracvacca@gmail.com')
on conflict (email) do nothing;
