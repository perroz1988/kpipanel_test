import { jwtVerify } from 'jose';

const JWT_SECRET = new TextEncoder().encode(process.env.JWT_SECRET);

function parseCookies(cookieHeader = '') {
  return Object.fromEntries(
    cookieHeader.split(';').map(c => {
      const idx = c.indexOf('=');
      return idx < 0 ? [c.trim(), ''] : [c.slice(0, idx).trim(), c.slice(idx + 1).trim()];
    })
  );
}

export default async function handler(req, res) {
  const token = parseCookies(req.headers.cookie)['kpi_session'];

  if (!token) return res.status(401).json({ error: 'Non autenticato' });

  try {
    const { payload } = await jwtVerify(token, JWT_SECRET);
    res.status(200).json({ name: payload.name, email: payload.email, role: payload.role });
  } catch {
    res.setHeader('Set-Cookie', 'kpi_session=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0');
    res.status(401).json({ error: 'Sessione scaduta' });
  }
}
