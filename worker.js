// Anomaly Quantix - Pro License Worker
// Cloudflare Worker with KV storage
// KV Binding: QUANTIX_PRO
// Environment Variables: GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, WORKER_URL, APP_URL

const ALLOWED_ORIGIN = 'https://anomaly-surround.github.io';

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // CORS
    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders() });
    }

    try {
      // Google OAuth callback
      if (url.pathname === '/auth/google/callback') {
        return handleGoogleCallback(url, env);
      }

      // Get Google auth URL
      if (url.pathname === '/auth/google') {
        return handleGoogleRedirect(env);
      }

      // Check pro status (requires token)
      if (url.pathname === '/api/status') {
        return handleStatus(request, env);
      }

      // Activate license key (requires token + key)
      if (url.pathname === '/api/activate' && request.method === 'POST') {
        return handleActivate(request, env);
      }

      // Deactivate
      if (url.pathname === '/api/deactivate' && request.method === 'POST') {
        return handleDeactivate(request, env);
      }

      return json({ error: 'Not found' }, 404);
    } catch (err) {
      return json({ error: err.message }, 500);
    }
  }
};

// Google OAuth: redirect to Google login
function handleGoogleRedirect(env) {
  const params = new URLSearchParams({
    client_id: env.GOOGLE_CLIENT_ID,
    redirect_uri: env.WORKER_URL + '/auth/google/callback',
    response_type: 'code',
    scope: 'email profile',
    access_type: 'online',
    prompt: 'select_account'
  });

  return Response.redirect('https://accounts.google.com/o/oauth2/v2/auth?' + params.toString(), 302);
}

// Google OAuth: handle callback, exchange code for token, get user info
async function handleGoogleCallback(url, env) {
  const code = url.searchParams.get('code');
  const appUrl = env.APP_URL || ALLOWED_ORIGIN + '/anomaly-quantix';

  if (!code) return Response.redirect(appUrl + '?error=no_code', 302);

  // Exchange code for tokens
  const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      code,
      client_id: env.GOOGLE_CLIENT_ID,
      client_secret: env.GOOGLE_CLIENT_SECRET,
      redirect_uri: env.WORKER_URL + '/auth/google/callback',
      grant_type: 'authorization_code'
    })
  });

  const tokens = await tokenRes.json();
  if (!tokens.access_token) {
    return Response.redirect(appUrl + '?error=login_failed', 302);
  }

  // Get user info
  const userRes = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
    headers: { Authorization: 'Bearer ' + tokens.access_token }
  });

  const user = await userRes.json();
  if (!user.email) {
    return Response.redirect(appUrl + '?error=no_email', 302);
  }

  // Generate a session token
  const sessionToken = crypto.randomUUID();

  // Store session with 30-day expiry
  await env.QUANTIX_PRO.put('session:' + sessionToken, JSON.stringify({
    email: user.email,
    name: user.name || '',
    picture: user.picture || ''
  }), { expirationTtl: 60 * 60 * 24 * 30 });

  // Redirect back to app with token only (pro status fetched via /api/status)
  return Response.redirect(appUrl + '?token=' + sessionToken, 302);
}

// Check pro status
async function handleStatus(request, env) {
  const token = getToken(request);
  if (!token) return json({ error: 'Not authenticated' }, 401);

  const session = await getSession(token, env);
  if (!session) return json({ error: 'Invalid session' }, 401);

  const proData = await env.QUANTIX_PRO.get('pro:' + session.email);

  return json({
    email: session.email,
    name: session.name,
    picture: session.picture,
    pro: proData ? true : false,
    licenseKey: proData ? JSON.parse(proData).key : null
  });
}

// Activate license key — validate with LemonSqueezy first
async function handleActivate(request, env) {
  const token = getToken(request);
  if (!token) return json({ error: 'Not authenticated' }, 401);

  const session = await getSession(token, env);
  if (!session) return json({ error: 'Invalid session' }, 401);

  const body = await request.json();
  const key = body.key?.trim();
  if (!key || key.length < 6) return json({ error: 'Invalid license key' }, 400);

  // Check if key is already used by someone else
  const existingUser = await env.QUANTIX_PRO.get('key:' + key);
  if (existingUser && existingUser !== session.email) {
    return json({ error: 'This license key is already in use' }, 400);
  }

  // Validate with LemonSqueezy API
  const valid = await validateWithLemonSqueezy(key);
  if (!valid) {
    return json({ error: 'Invalid or expired license key' }, 400);
  }

  // Store pro status
  await env.QUANTIX_PRO.put('pro:' + session.email, JSON.stringify({
    key,
    activatedAt: Date.now()
  }));

  // Store key -> email mapping
  await env.QUANTIX_PRO.put('key:' + key, session.email);

  return json({ success: true, pro: true });
}

// Validate license key with LemonSqueezy
async function validateWithLemonSqueezy(key) {
  try {
    const res = await fetch('https://api.lemonsqueezy.com/v1/licenses/validate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Accept': 'application/json' },
      body: 'license_key=' + encodeURIComponent(key)
    });

    const data = await res.json();
    // Valid if the key exists and is active
    if (data.valid) return true;
    if (data.license_key && data.license_key.status === 'active') return true;
    return false;
  } catch {
    // If LemonSqueezy is unreachable, deny activation (fail closed)
    return false;
  }
}

// Deactivate
async function handleDeactivate(request, env) {
  const token = getToken(request);
  if (!token) return json({ error: 'Not authenticated' }, 401);

  const session = await getSession(token, env);
  if (!session) return json({ error: 'Invalid session' }, 401);

  const proData = await env.QUANTIX_PRO.get('pro:' + session.email);
  if (proData) {
    const { key } = JSON.parse(proData);
    await env.QUANTIX_PRO.delete('key:' + key);
  }
  await env.QUANTIX_PRO.delete('pro:' + session.email);

  return json({ success: true, pro: false });
}

// Helpers
function getToken(request) {
  const auth = request.headers.get('Authorization');
  if (auth?.startsWith('Bearer ')) return auth.slice(7);
  return null;
}

async function getSession(token, env) {
  if (!token || token.length < 10) return null;
  const data = await env.QUANTIX_PRO.get('session:' + token);
  return data ? JSON.parse(data) : null;
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...corsHeaders() }
  });
}

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization'
  };
}
