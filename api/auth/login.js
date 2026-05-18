import { createClient } from '@supabase/supabase-js';
import bcrypt from 'bcryptjs';
import { SignJWT } from 'jose';

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);
const JWT_SECRET = new TextEncoder().encode(process.env.JWT_SECRET);

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const { email, password } = req.body ?? {};
  const ip = req.headers['x-forwarded-for']?.split(',')[0] ?? req.socket?.remoteAddress;
  const ua = req.headers['user-agent'];

  if (!email || !password) {
    return res.status(400).json({ error: 'Email e password richiesti' });
  }

  const { data: user } = await supabase
    .from('kpi_users')
    .select('*')
    .eq('email', email.toLowerCase().trim())
    .eq('is_active', true)
    .single();

  if (!user || !(await bcrypt.compare(password, user.password_hash))) {
    await supabase.from('kpi_access_logs').insert({
      user_id: user?.id ?? null,
      email: email.toLowerCase().trim(),
      ip, user_agent: ua, action: 'failed_login'
    });
    return res.status(401).json({ error: 'Credenziali non valide' });
  }

  const token = await new SignJWT({ sub: user.id, email: user.email, name: user.name, role: user.role })
    .setProtectedHeader({ alg: 'HS256' })
    .setExpirationTime('7d')
    .sign(JWT_SECRET);

  await Promise.all([
    supabase.from('kpi_users').update({ last_login: new Date().toISOString() }).eq('id', user.id),
    supabase.from('kpi_access_logs').insert({ user_id: user.id, email: user.email, ip, user_agent: ua, action: 'login' })
  ]);

  res.setHeader('Set-Cookie', `kpi_session=${token}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=${7 * 24 * 60 * 60}`);
  res.status(200).json({ name: user.name });
}
