const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const PORT = process.env.PORT || 3001;

function readKey() {
  if (process.env.ZHIPU_KEY) return process.env.ZHIPU_KEY.trim();
  try { return fs.readFileSync(path.join(ROOT, 'zhipu.key'), 'utf8').trim(); } catch (e) { return ''; }
}
const ZHIPU_KEY = readKey();

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'POST, GET, OPTIONS'
};

function sendJson(res, code, obj) {
  res.writeHead(code, Object.assign({ 'Content-Type': 'application/json; charset=utf-8' }, CORS));
  res.end(JSON.stringify(obj));
}

async function recognize(image) {
  const payload = {
    model: 'glm-4v-flash',
    messages: [{
      role: 'user',
      content: [
        { type: 'image_url', image_url: { url: image } },
        { type: 'text', text: '你是食物卡路里识别助手。识别图片里的食物或饮料，估算这一份的热量。只返回 JSON，格式：{"name":"食物名","kcal":数字}，kcal 为千卡数。不要输出任何其它文字。' }
      ]
    }]
  };
  const r = await fetch('https://open.bigmodel.cn/api/paas/v4/chat/completions', {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + ZHIPU_KEY, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  const j = await r.json();
  const content = (j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) || '';
  const m = content.match(/\{[\s\S]*\}/);
  if (m) {
    try {
      const o = JSON.parse(m[0]);
      return { name: String(o.name || '未知食物'), kcal: Math.round(Number(o.kcal) || 0) };
    } catch (e) { }
  }
  return { name: '识别失败', kcal: 0, raw: content };
}

const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.svg': 'image/svg+xml', '.json': 'application/json; charset=utf-8' };

const server = http.createServer((req, res) => {
  if (req.method === 'OPTIONS') { res.writeHead(204, CORS); res.end(); return; }

  if (req.url === '/api/food-recognize' && req.method === 'POST') {
    let body = '';
    req.on('data', c => { body += c; if (body.length > 12 * 1024 * 1024) req.destroy(); });
    req.on('end', async () => {
      try {
        if (!ZHIPU_KEY) return sendJson(res, 500, { error: '未配置 ZHIPU_KEY（请创建 zhipu.key 文件）' });
        const { image } = JSON.parse(body || '{}');
        if (!image) return sendJson(res, 400, { error: 'no image' });
        const out = await recognize(image);
        sendJson(res, 200, out);
      } catch (e) {
        sendJson(res, 500, { error: String(e && e.message || e) });
      }
    });
    return;
  }

  let p = decodeURIComponent((req.url || '/').split('?')[0]);
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
  console.log('卡路里之路 server: http://localhost:' + PORT + (ZHIPU_KEY ? '  [AI识卡已启用]' : '  [未配置 ZHIPU_KEY]'));
});
