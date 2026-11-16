// Посредник для тех, у кого Google недоступен (например, из России): Yandex Cloud Functions, среда Node.js 18 или новее.
// Пересылает запросы сайта в скрипт таблицы и возвращает ответ как есть. Пароль проверяет сам скрипт — посредник его не знает.
// Настройка — README.md, раздел «Доступ из России».
const API = 'https://script.google.com/macros/s/AKfycbw2fj5R4qPbx6eQ1jQoq0rtN3sbnRUDEmLnlLrDSazVF2KhZX8EnZxORVCDcdZ2oHXe/exec';
const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type' };

module.exports.handler = async event => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: CORS, body: '' };
  if (event.httpMethod !== 'GET' && event.httpMethod !== 'POST') return { statusCode: 405, headers: CORS, body: '' };
  const body = event.isBase64Encoded ? Buffer.from(event.body || '', 'base64').toString('utf8') : event.body;
  // Google отвечает переадресацией на googleusercontent.com — fetch проходит её сам (после POST — запросом GET, как и браузер)
  const res = await fetch(API, event.httpMethod === 'POST' ? { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body } : {});
  return { statusCode: res.status, headers: { ...CORS, 'Content-Type': 'application/json; charset=utf-8' }, body: await res.text() };
};
