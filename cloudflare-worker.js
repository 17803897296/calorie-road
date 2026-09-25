// 卡路里之路 · 拍照识卡 Cloudflare Worker（免费、免实名、全球可访问）
// 部署：dash.cloudflare.com → Workers & Pages → 创建 Worker → 粘贴本文件全部内容 → 部署
const ZHIPU_KEY = '请填入你的智谱API Key';
const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS'
};
function json(o, s) {
  return new Response(JSON.stringify(o), {
    status: s || 200,
    headers: Object.assign({ 'Content-Type': 'application/json; charset=utf-8' }, CORS)
  });
}

export default {
  async fetch(request) {
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: CORS });
    if (request.method !== 'POST') return json({ error: 'use POST' }, 405);
    try {
      const { image } = await request.json();
      if (!image) return json({ error: 'no image' }, 400);
      const r = await fetch('https://open.bigmodel.cn/api/paas/v4/chat/completions', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + ZHIPU_KEY, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'glm-4v-flash',
          messages: [{
            role: 'user',
            content: [
              { type: 'image_url', image_url: { url: image } },
              { type: 'text', text: '你是食物卡路里识别助手。识别图片里的食物或饮料，估算这一份的热量。只返回 JSON，格式：{"name":"食物名","kcal":数字}，kcal 为千卡数。不要输出任何其它文字。' }
            ]
          }]
        })
      });
      const j = await r.json();
      const content = (j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) || '';
      const m = content.match(/\{[\s\S]*\}/);
      let out = { name: '识别失败', kcal: 0 };
      if (m) {
        try {
          const o = JSON.parse(m[0]);
          out = { name: String(o.name || '未知食物'), kcal: Math.round(Number(o.kcal) || 0) };
        } catch (e) { }
      }
      return json(out, 200);
    } catch (e) {
      return json({ error: String(e && e.message || e) }, 500);
    }
  }
};
