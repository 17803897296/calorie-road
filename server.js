const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT = __dirname;
const PORT = process.env.PORT || 3001;
const DATA_DIR = path.join(ROOT, 'data');
const USERS_FILE = path.join(DATA_DIR, 'users.json');

function readKey() {
  if (process.env.ZHIPU_KEY) return process.env.ZHIPU_KEY.trim();
  try { return fs.readFileSync(path.join(ROOT, 'zhipu.key'), 'utf8').trim(); } catch (e) { return ''; }
}
const ZHIPU_KEY = readKey();

/* ===== OAuth 配置（oauth.json，请勿提交，已 gitignore） =====
   {
     "wechat": { "appid": "", "secret": "" },   // 微信开放平台 网站应用/移动应用
     "qq":     { "appid": "", "secret": "" },   // QQ互联
     "baseUrl": "https://你的域名"               // 回调域名（须与平台登记一致）
   }
*/
let OAUTH = {};
try { OAUTH = JSON.parse(fs.readFileSync(path.join(ROOT, 'oauth.json'), 'utf8')); } catch (e) { OAUTH = {}; }
function oauthBase(req) {
  if (OAUTH.baseUrl) return OAUTH.baseUrl.replace(/\/+$/, '');
  const host = req.headers['host'] || ('localhost:' + PORT);
  const proto = req.headers['x-forwarded-proto'] || 'http';
  return proto + '://' + host;
}

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  'Access-Control-Allow-Methods': 'POST, GET, OPTIONS'
};
function sendJson(res, code, obj) {
  res.writeHead(code, Object.assign({ 'Content-Type': 'application/json; charset=utf-8' }, CORS));
  res.end(JSON.stringify(obj));
}
function readBody(req, limit = 12 * 1024 * 1024) {
  return new Promise((resolve) => {
    let b = '';
    req.on('data', c => { b += c; if (b.length > limit) req.destroy(); });
    req.on('end', () => resolve(b));
  });
}

/* ================= 用户库（JSON 文件持久化） ================= */
let db = { users: {} };
function loadDb() { try { db = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8')) || { users: {} }; } catch (e) { db = { users: {} }; } }
function saveDb() { try { fs.mkdirSync(DATA_DIR, { recursive: true }); fs.writeFileSync(USERS_FILE, JSON.stringify(db, null, 2)); } catch (e) { } }
loadDb();

function hashPass(user, pass) {
  return crypto.createHash('sha256').update(user + '::' + pass + '::cardroute').digest('hex');
}
function newToken() { return crypto.randomBytes(24).toString('hex'); }
function authUser(req) {
  const h = req.headers['authorization'] || '';
  const m = h.match(/^Bearer\s+(.+)$/i);
  if (!m) return null;
  const t = m[1];
  for (const u in db.users) { if (db.users[u].token === t) return u; }
  return null;
}

/* ================= 智谱调用 ================= */
async function zhipu(model, messages) {
  const r = await fetch('https://open.bigmodel.cn/api/paas/v4/chat/completions', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + ZHIPU_KEY, 'Content-Type': 'application/json' },
    body: JSON.stringify({ model, messages })
  });
  const j = await r.json();
  return (j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) || '';
}
async function recognize(image) {
  const content = await zhipu('glm-4v-flash', [{ role: 'user', content: [
    { type: 'image_url', image_url: { url: image } },
    { type: 'text', text: '你是食物卡路里识别助手。识别图片里的食物或饮料，估算这一份的热量。只返回 JSON，格式：{"name":"食物名","kcal":数字}，kcal 为千卡数。不要输出任何其它文字。' }
  ] }]);
  const m = content.match(/\{[\s\S]*\}/);
  if (m) { try { const o = JSON.parse(m[0]); return { name: String(o.name || '未知食物'), kcal: Math.round(Number(o.kcal) || 0) }; } catch (e) { } }
  return { name: '识别失败', kcal: 0, raw: content };
}

const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.svg': 'image/svg+xml', '.json': 'application/json; charset=utf-8', '.key': 'text/plain' };

const server = http.createServer(async (req, res) => {
  if (req.method === 'OPTIONS') { res.writeHead(204, CORS); res.end(); return; }
  const url = (req.url || '/').split('?')[0];

  try {
    /* ---- 注册 ---- */
    if (url === '/api/register' && req.method === 'POST') {
      const { user, pass } = JSON.parse(await readBody(req) || '{}');
      if (!user || !pass) return sendJson(res, 400, { error: '用户名和密码必填' });
      if (db.users[user]) return sendJson(res, 409, { error: '用户名已存在' });
      const token = newToken();
      db.users[user] = { pass: hashPass(user, pass), token, state: null, created: Date.now() };
      saveDb();
      return sendJson(res, 200, { ok: true, user, token });
    }
    /* ---- 登录 ---- */
    if (url === '/api/login' && req.method === 'POST') {
      const { user, pass } = JSON.parse(await readBody(req) || '{}');
      const u = db.users[user];
      if (!u || u.pass !== hashPass(user, pass)) return sendJson(res, 401, { error: '用户名或密码错误' });
      u.token = newToken(); saveDb();
      return sendJson(res, 200, { ok: true, user, token: u.token });
    }
    /* ---- 取状态 ---- */
    if (url === '/api/state' && req.method === 'GET') {
      const user = authUser(req);
      if (!user) return sendJson(res, 401, { error: '未登录' });
      return sendJson(res, 200, { ok: true, state: db.users[user].state || null });
    }
    /* ---- 存状态 ---- */
    if (url === '/api/state' && req.method === 'POST') {
      const user = authUser(req);
      if (!user) return sendJson(res, 401, { error: '未登录' });
      const body = JSON.parse(await readBody(req) || '{}');
      db.users[user].state = body.state || null;
      db.users[user].updated = Date.now();
      saveDb();
      return sendJson(res, 200, { ok: true });
    }
    /* ---- 拍照识卡 ---- */
    if (url === '/api/food-recognize' && req.method === 'POST') {
      if (!ZHIPU_KEY) return sendJson(res, 500, { error: '未配置 ZHIPU_KEY' });
      const { image } = JSON.parse(await readBody(req) || '{}');
      if (!image) return sendJson(res, 400, { error: 'no image' });
      return sendJson(res, 200, await recognize(image));
    }
    /* ---- 给食谱提意见（大模型修正） ---- */
    if (url === '/api/recipe-feedback' && req.method === 'POST') {
      if (!ZHIPU_KEY) return sendJson(res, 500, { error: '未配置 ZHIPU_KEY' });
      const { text, meals, dislikes, wants } = JSON.parse(await readBody(req) || '{}');
      const sys = '你是减肥食谱助手。用户会对当前食谱提意见。请分别理解用户的【忌口/厌倦】和【想吃】。关键词要简短具体（如 面、烤肉、奶茶、水饺）。只返回 JSON：{"avoid":["要避开的食物关键词"],"want":["想吃的食物关键词"],"note":"一句话说明"}。没有的用空数组，不要输出任何其它文字。';
      const usr = `当前三餐：${JSON.stringify(meals || [])}\n已有忌口：${JSON.stringify(dislikes || [])}\n已有想吃：${JSON.stringify(wants || [])}\n用户意见：${text || ''}`;
      const content = await zhipu('glm-4-flash', [{ role: 'system', content: sys }, { role: 'user', content: usr }]);
      const m = content.match(/\{[\s\S]*\}/);
      let out = { avoid: [], want: [], note: '' };
      if (m) { try { const o = JSON.parse(m[0]); out = { avoid: o.avoid || [], want: o.want || [], note: o.note || '' }; } catch (e) { } }
      return sendJson(res, 200, out);
    }

    /* ---- 微信网页扫码登录 ---- */
    if (url === '/api/oauth/wechat/start') {
      const c = OAUTH.wechat || {};
      if (!c.appid) return sendJson(res, 500, { error: '未配置微信 oauth.json' });
      const redirect = encodeURIComponent(oauthBase(req) + '/api/oauth/wechat/callback');
      const state = crypto.randomBytes(8).toString('hex');
      res.writeHead(302, { Location: 'https://open.weixin.qq.com/connect/qrconnect?appid=' + c.appid + '&redirect_uri=' + redirect + '&response_type=code&scope=snsapi_login&state=' + state + '#wechat_redirect' });
      res.end(); return;
    }
    if (url === '/api/oauth/wechat/callback') {
      const c = OAUTH.wechat || {};
      const code = (req.url.split('code=')[1] || '').split('&')[0];
      if (!code) return sendJson(res, 400, { error: 'no code' });
      const r1 = await fetch('https://api.weixin.qq.com/sns/oauth2/access_token?appid=' + c.appid + '&secret=' + c.secret + '&code=' + code + '&grant_type=authorization_code');
      const j1 = await r1.json();
      if (!j1.openid) return sendJson(res, 400, { error: 'wechat: ' + (j1.errmsg || 'failed') });
      const r2 = await fetch('https://api.weixin.qq.com/sns/userinfo?access_token=' + j1.access_token + '&openid=' + j1.openid);
      const j2 = await r2.json();
      const uname = 'wx_' + j1.openid;
      if (!db.users[uname]) db.users[uname] = { pass: '', state: null, created: Date.now(), oauth: 'wechat', nickname: j2.nickname || '微信用户' };
      db.users[uname].token = newToken(); saveDb();
      res.writeHead(302, { Location: '/?oauth_token=' + db.users[uname].token + '&oauth_user=' + encodeURIComponent(db.users[uname].nickname) });
      res.end(); return;
    }
    /* ---- QQ 网页登录 ---- */
    if (url === '/api/oauth/qq/start') {
      const c = OAUTH.qq || {};
      if (!c.appid) return sendJson(res, 500, { error: '未配置 QQ oauth.json' });
      const redirect = encodeURIComponent(oauthBase(req) + '/api/oauth/qq/callback');
      const state = crypto.randomBytes(8).toString('hex');
      res.writeHead(302, { Location: 'https://graph.qq.com/oauth2.0/authorize?response_type=code&client_id=' + c.appid + '&redirect_uri=' + redirect + '&state=' + state });
      res.end(); return;
    }
    if (url === '/api/oauth/qq/callback') {
      const c = OAUTH.qq || {};
      const code = (req.url.split('code=')[1] || '').split('&')[0];
      if (!code) return sendJson(res, 400, { error: 'no code' });
      const redirect = oauthBase(req) + '/api/oauth/qq/callback';
      const r1 = await fetch('https://graph.qq.com/oauth2.0/token?grant_type=authorization_code&client_id=' + c.appid + '&client_secret=' + c.secret + '&code=' + code + '&redirect_uri=' + encodeURIComponent(redirect) + '&fmt=json');
      const j1 = await r1.json();
      if (!j1.access_token) return sendJson(res, 400, { error: 'qq token failed' });
      const r2 = await fetch('https://graph.qq.com/oauth2.0/me?access_token=' + j1.access_token + '&fmt=json');
      const j2 = await r2.json();
      if (!j2.openid) return sendJson(res, 400, { error: 'qq openid failed' });
      const r3 = await fetch('https://graph.qq.com/user/get_user_info?access_token=' + j1.access_token + '&oauth_consumer_key=' + c.appid + '&openid=' + j2.openid);
      const j3 = await r3.json();
      const uname = 'qq_' + j2.openid;
      if (!db.users[uname]) db.users[uname] = { pass: '', state: null, created: Date.now(), oauth: 'qq', nickname: j3.nickname || 'QQ用户' };
      db.users[uname].token = newToken(); saveDb();
      res.writeHead(302, { Location: '/?oauth_token=' + db.users[uname].token + '&oauth_user=' + encodeURIComponent(db.users[uname].nickname) });
      res.end(); return;
    }
  } catch (e) {
    return sendJson(res, 500, { error: String(e && e.message || e) });
  }

  /* ---- 静态文件 ---- */
  let p = decodeURIComponent(url);
  if (p === '/') p = '/index.html';
  const fp = path.normalize(path.join(ROOT, p));
  if (!fp.startsWith(ROOT)) { res.writeHead(403); res.end('forbidden'); return; }
  fs.readFile(fp, (err, data) => {
    if (err) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(fp).toLowerCase()] || 'application/octet-stream' });
    res.end(data);
  });
});

server.listen(PORT, '0.0.0.0', () => {
  console.log('卡路里之路 server: http://localhost:' + PORT + (ZHIPU_KEY ? '  [AI已启用]' : '  [未配置 ZHIPU_KEY]'));
});
