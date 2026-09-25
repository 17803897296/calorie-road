'use strict';
// 卡路里之路 · 拍照识卡云函数（腾讯云 SCF / 事件函数 + 函数URL 或 API网关触发）
// 零依赖，直接粘贴到云函数代码框即可。Node.js 运行环境。
const https = require('https');

const ZHIPU_KEY = process.env.ZHIPU_KEY || '请填入你的智谱API Key';

function callZhipu(payload) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(payload);
    const req = https.request({
      hostname: 'open.bigmodel.cn',
      path: '/api/paas/v4/chat/completions',
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + ZHIPU_KEY,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data)
      }
    }, res => {
      let body = '';
      res.on('data', c => body += c);
      res.on('end', () => { try { resolve(JSON.parse(body)); } catch (e) { reject(e); } });
    });
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

exports.main_handler = async (event) => {
  const cors = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS'
  };

  const method = event.httpMethod || (event.requestContext && event.requestContext.httpMethod) || 'POST';
  if (method === 'OPTIONS') {
    return { isBase64Encoded: false, statusCode: 204, headers: cors, body: '' };
  }

  try {
    let raw = event.body || '{}';
    if (event.isBase64Encoded) raw = Buffer.from(raw, 'base64').toString('utf8');
    const { image } = JSON.parse(raw);
    if (!image) {
      return { isBase64Encoded: false, statusCode: 400, headers: cors, body: JSON.stringify({ error: 'no image' }) };
    }

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

    const j = await callZhipu(payload);
    const content = (j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) || '';
    const m = content.match(/\{[\s\S]*\}/);
    let out = { name: '识别失败', kcal: 0 };
    if (m) {
      try {
        const o = JSON.parse(m[0]);
        out = { name: String(o.name || '未知食物'), kcal: Math.round(Number(o.kcal) || 0) };
      } catch (e) { }
    }
    return { isBase64Encoded: false, statusCode: 200, headers: cors, body: JSON.stringify(out) };
  } catch (e) {
    return { isBase64Encoded: false, statusCode: 500, headers: cors, body: JSON.stringify({ error: String(e && e.message || e) }) };
  }
};
