// Anomaly Quantix - Pro License Worker
// Cloudflare Worker with KV storage
// KV Binding: QUANTIX_PRO
// Environment Variables: GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, WORKER_URL

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
  if (!code) return json({ error: 'No code' }, 400);

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
    return redirectToApp(env, null, 'Google login failed');
  }

  // Get user info
  const userRes = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
    headers: { Authorization: 'Bearer ' + tokens.access_token }
  });

  const user = await userRes.json();
  if (!user.email) {
    return redirectToApp(env, null, 'Could not get email');
  }

  // Generate a simple session token
  const sessionToken = crypto.randomUUID();

  // Store session: token -> { email, name, picture }
  await env.QUANTIX_PRO.put('session:' + sessionToken, JSON.stringify({
    email: user.email,
    name: user.name || '',
    picture: user.picture || ''
  }), { expirationTtl: 60 * 60 * 24 * 365 }); // 1 year

  // Check if user has pro
  const proData = await env.QUANTIX_PRO.get('pro:' + user.email);

  // Redirect back to app with token
  return redirectToApp(env, sessionToken, null, proData ? true : false);
}

function redirectToApp(env, token, error, isPro) {
  const appUrl = env.APP_URL || 'https://anomaly-surround.github.io/anomaly-quantix';
  const params = new URLSearchParams();
  if (token) params.set('token', token);
  if (error) params.set('error', error);
  if (isPro) params.set('pro', '1');
  return Response.redirect(appUrl + '?' + params.toString(), 302);
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

// Activate license key and tie to Google account
async function handleActivate(request, env) {
  const token = getToken(request);
  if (!token) return json({ error: 'Not authenticated' }, 401);

  const session = await getSession(token, env);
  if (!session) return json({ error: 'Invalid session' }, 401);

  const body = await request.json();
  const key = body.key?.trim();
  if (!key || key.length < 6) return json({ error: 'Invalid license key' }, 400);

  // Store pro status
  await env.QUANTIX_PRO.put('pro:' + session.email, JSON.stringify({
    key,
    activatedAt: Date.now()
  }));

  // Also store key -> email mapping (to prevent reuse if needed later)
  await env.QUANTIX_PRO.put('key:' + key, session.email);

  return json({ success: true, pro: true });
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
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization'
  };
}
