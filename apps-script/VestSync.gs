/**
 * 出金中心：Vest 收信自動同步（Google Apps Script）
 * 每 5 分鐘讀一次 Gmail 的 Vest 收據與出金通知（含垃圾桶），新的就寫進 Firestore。
 * 第一次：執行 setup() 並完成授權。之後全自動。
 */
const PROJECT = 'trade-sync-e41ce';
const UID = '5vXUJX9CiUeEIye29neWlDCnbIj2';
const BASE = 'https://firestore.googleapis.com/v1/projects/' + PROJECT + '/databases/(default)/documents/payoutCenter/' + UID + '/records';
const QUERY = 'from:noreply@spark.vestmarkets.com (subject:"Vest receipt" OR subject:"Payout Completed") in:anywhere';

function setup() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncVest').timeBased().everyMinutes(5).create();
  syncVest();
}

function api_(method, url, body) {
  const res = UrlFetchApp.fetch(url, {
    method: method, muteHttpExceptions: true, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'x-goog-user-project': PROJECT },
    payload: body ? JSON.stringify(body) : undefined
  });
  if (res.getResponseCode() >= 300) throw new Error(method + ' ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText() || '{}');
}

function val_(v) {
  if (!v) return null;
  if ('stringValue' in v) return v.stringValue;
  if ('doubleValue' in v) return v.doubleValue;
  if ('integerValue' in v) return +v.integerValue;
  return null;
}

function listRecords_() {
  const out = []; let token = '';
  do {
    const j = api_('get', BASE + '?pageSize=300' + (token ? '&pageToken=' + token : ''));
    (j.documents || []).forEach(d => {
      const f = d.fields || {}; const r = { id: d.name.split('/').pop() };
      Object.keys(f).forEach(k => r[k] = val_(f[k]));
      out.push(r);
    });
    token = j.nextPageToken || '';
  } while (token);
  return out;
}

function writeRecord_(r) {
  const f = {};
  ['date', 'firm', 'size', 'type', 'cat', 'note'].forEach(k => f[k] = { stringValue: String(r[k] || '') });
  f.amount = { doubleValue: +r.amount };
  f.split = r.split == null ? { nullValue: null } : { doubleValue: +r.split };
  api_('patch', BASE + '/' + encodeURIComponent(r.id), { fields: f });
}

function spec_(prod) {
  if (/Instant/i.test(prod)) return { size: '500 Instant', spec: 'Instant 500' };
  const parts = prod.split('|').map(x => x.trim());
  const goal = (parts.find(x => /Goal/i.test(x)) || '').replace(/\s*Goal/i, '');
  const dlRaw = parts.find(x => /Daily Loss/i.test(x)) || '';
  const dl = /No Daily/i.test(dlRaw) ? '無日損' : dlRaw.replace(/\s*Daily Loss/i, '日損');
  const sp = parts.find(x => /^\d+\/\d+$/.test(x)) || '';
  return { size: '5K', spec: 'Eval ' + goal + '目標 · ' + dl + ' · ' + sp };
}

const tw_ = d => Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd');
const num_ = s => +String(s || '0').replace(/,/g, '');

function syncVest() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    // 讀 Gmail
    let threads = [], start = 0;
    while (true) { const t = GmailApp.search(QUERY, start, 100); threads = threads.concat(t); if (t.length < 100) break; start += 100; }
    const receipts = [], payouts = [];
    threads.forEach(th => th.getMessages().forEach(m => {
      const from = m.getFrom(), subj = m.getSubject(), t = m.getPlainBody();
      if (!/vestmarkets\.com/i.test(from)) return;
      if (/Vest receipt/i.test(subj)) {
        const rc = (t.match(/Receipt (VST-\d+-\d+)/) || [])[1];
        const paid = num_((t.match(/(?:Confirmed purchase amount \(USD\)|Total paid \(USD\)|Amount paid|Confirmed purchase amount):\s*\$([\d.,]+)/) || [])[1]);
        const price = num_((t.match(/Price per account:\s*\$([\d.,]+)/) || [])[1]);
        const qty = +((t.match(/Number of accounts:\s*(\d+)/) || [])[1] || 1);
        const prod = ((t.match(/\n(Vest Capital[^\n]+)/) || [])[1] || '').trim();
        if (rc && paid > 0) receipts.push({ date: m.getDate(), rc: rc, paid: paid, price: price, qty: qty, prod: prod });
      } else if (/Payout Completed/i.test(subj)) {
        const amt = num_((t.match(/([\d.,]+)\s*USDC/) || [])[1]);
        if (amt > 0) payouts.push({ date: m.getDate(), id: m.getId(), amt: amt });
      }
    }));

    // 比對現有資料
    const allRecs = listRecords_();
    const recs = allRecs.filter(r => /vest/i.test(r.firm || ''));
    const known = new Set(recs.map(r => ((r.note || '').match(/VST-\d+-\d+/) || [])[0]).filter(Boolean));
    let maxAcc = 0; recs.forEach(r => { const m = /^vest-acc-(\d+)/.exec(r.id); if (m) maxAcc = Math.max(maxAcc, +m[1]); });
    const pays = recs.filter(r => r.type === 'payout');
    const ids = new Set(recs.map(r => r.id));
    let added = 0;

    // 舊資料（沒有收據編號）：依購買順序對回同一個 Account 編號，補上收據資訊而不是新增
    const legacy = {}; recs.forEach(r => { if (/^vest-acc-\d+$/.test(r.id) && !/VST-/.test(r.note || '')) legacy[r.id] = r; });
    let seq = 0;
    receipts.sort((a, b) => a.date - b.date).forEach(r => {
      const firstSeq = seq + 1; seq += r.qty;
      if (known.has(r.rc)) return;
      const legacyId = 'vest-acc-' + String(firstSeq).padStart(2, '0');
      const lp = legacy[legacyId] && +(((legacy[legacyId].note || '').match(/原價 \$([\d.]+)/) || [])[1] || 0);
      if (r.qty === 1 && legacy[legacyId] && lp === r.price) {
        const sp0 = spec_(r.prod), nn0 = String(firstSeq).padStart(2, '0');
        writeRecord_({ id: legacyId, date: tw_(r.date), firm: 'Vest Markets', size: sp0.size, type: 'cost', cat: '購買帳號',
          amount: r.paid, split: null, note: 'Account ' + nn0 + ' · ' + sp0.spec + ' · 原價 $' + r.price + ' · 收據 ' + r.rc });
        known.add(r.rc); delete legacy[legacyId]; return;
      }
      const sp = spec_(r.prod), unit = Math.round(r.paid / r.qty * 100) / 100;
      for (let k = 0; k < r.qty; k++) {
        maxAcc++; const nn = String(maxAcc).padStart(2, '0');
        notify_('🧾 Vest 購買帳號 Account ' + nn + '\n' + sp.spec + '\n成本 $' + unit + '（' + tw_(r.date) + '）');
        writeRecord_({ id: 'vest-acc-' + nn, date: tw_(r.date), firm: 'Vest Markets', size: sp.size, type: 'cost', cat: '購買帳號',
          amount: unit, split: null, note: 'Account ' + nn + ' · ' + sp.spec + ' · 原價 $' + r.price + ' · 收據 ' + r.rc });
        added++;
      }
      known.add(r.rc);
    });

    payouts.forEach(p => {
      const id = 'vest-pay-mail-' + p.id, d = tw_(p.date);
      if (ids.has(id)) return;
      if (pays.some(e => e.date === d && Math.abs(e.amount * (e.split == null ? 100 : e.split) / 100 - p.amt) < 0.05)) return;
      writeRecord_({ id: id, date: d, firm: 'Vest Markets', size: '', type: 'payout', cat: '出金', amount: p.amt, split: 100, note: 'Gmail 出金通知 · 實收 ' + p.amt + ' USDC' });
      pays.push({ date: d, amount: p.amt, split: 100 }); added++;
      notify_('💰 Vest 出金到帳 $' + p.amt + ' USDC（' + d + '）');
    });

    if (added) console.log('Vest 同步：新增 ' + added + ' 筆');
    syncLucidTdy_(allRecs);
  } finally { lock.releaseLock(); }
}


/* ---------- Lucid / Tradeify ---------- */
function syncLucidTdy_(allRecs) {
  const ids = new Set(allRecs.map(r => r.id));
  const pays = allRecs.filter(r => r.type === 'payout' && /lucid/i.test(r.firm || ''));
  const dayDiff = (a, b) => Math.abs((new Date(a) - new Date(b)) / 864e5);
  let added = 0;

  // Lucid 訂單（Order Processing 信）
  GmailApp.search('from:lucidtrading.com subject:"Order Processing" in:anywhere', 0, 200).forEach(th => th.getMessages().forEach(m => {
    const t = m.getPlainBody();
    const ord = (t.match(/Order number:\s*(\d+)/) || [])[1];
    const total = num_((t.match(/Total:\s*\|?\s*\$([\d.,]+)/) || [])[1]);
    if (!ord || !(total > 0)) return;
    const id = 'lucid-o-' + ord;
    if (ids.has(id)) return;
    const prod = ((t.match(/\|\s*(Lucid[A-Za-z]*\s+\d+K[^|]*)\|/) || [])[1] || '').trim();
    const size = ((prod.match(/(\d+)K/i) || [])[1] || '') ;
    writeRecord_({ id: id, date: tw_(m.getDate()), firm: 'Lucid Trading', size: size ? size + 'k' : '', type: 'cost', cat: '購買帳號', amount: total, split: null, note: '#' + ord + ' · ' + prod });
    ids.add(id); added++;
    notify_('🧾 Lucid 購買 ' + prod + '\n成本 $' + total + '（' + tw_(m.getDate()) + '）');
  }));

  // Lucid 出金（Payout Sent! 信，金額為實收，預設 90% 分成）
  GmailApp.search('from:lucidtrading.com subject:"Payout Sent" in:anywhere', 0, 100).forEach(th => th.getMessages().forEach(m => {
    const net = num_((m.getPlainBody().match(/Payment Amount\s*\$([\d.,]+)/) || [])[1]);
    if (!(net > 0)) return;
    const id = 'lucid-pay-mail-' + m.getId(), d = tw_(m.getDate());
    if (ids.has(id)) return;
    if (pays.some(e => dayDiff(e.date, d) <= 2 && Math.abs(e.amount * (e.split == null ? 100 : e.split) / 100 - net) < 0.05)) return;
    const gross = Math.round(net / 0.9 * 100) / 100;
    writeRecord_({ id: id, date: d, firm: 'Lucid Trading', size: '', type: 'payout', cat: '出金', amount: gross, split: 90, note: 'Gmail 出金通知 · 實收 $' + net });
    pays.push({ date: d, amount: gross, split: 90 }); ids.add(id); added++;
    notify_('💰 Lucid 出金到帳 $' + net + '（' + d + '）');
  }));

  // Tradeify 訂單（account is active 信）
  GmailApp.search('from:orders@tradeify.co subject:"account is active" in:anywhere', 0, 100).forEach(th => th.getMessages().forEach(m => {
    const t = m.getPlainBody();
    const ord = (t.match(/Order ID:\s*(\d+)/) || [])[1];
    const amt = num_((t.match(/Order Amount:\s*\$?([\d.,]+)/) || [])[1]);
    if (!ord || !(amt > 0)) return;
    const id = 'tdfy-o-' + ord;
    if (ids.has(id)) return;
    const acct = (t.match(/account name is\s*([A-Z0-9]+)/) || [])[1] || '';
    const size = ((acct.match(/^[A-Z]*?(\d{2,3})\d{6,}/) || [])[1] || '');
    writeRecord_({ id: id, date: tw_(m.getDate()), firm: 'Tradeify', size: size ? (+size) + 'k' : '', type: 'cost', cat: '購買帳號', amount: amt, split: null, note: '#' + ord + (acct ? ' · ' + acct : '') });
    ids.add(id); added++;
    notify_('🧾 Tradeify 購買 ' + acct + '\n成本 $' + amt);
  }));

  if (added) console.log('Lucid/Tradeify 同步：新增 ' + added + ' 筆');
}

/* ---------- Telegram 通知 ----------
 * 在「專案設定 → 指令碼屬性」新增：
 *   TG_TOKEN = 你的 Telegram Bot Token
 *   TG_CHAT  = 你的 Telegram chat id
 * 沒設定就不會發通知。 */
function notify_(text) {
  const p = PropertiesService.getScriptProperties();
  const token = p.getProperty('TG_TOKEN'), chat = p.getProperty('TG_CHAT');
  if (!token || !chat) return;
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post', muteHttpExceptions: true, contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chat, text: '出金中心\n' + text })
    });
  } catch (e) { console.log('Telegram 通知失敗：' + e); }
}

function testNotify() { notify_('✅ 測試通知：出金中心已連上 Telegram'); }
