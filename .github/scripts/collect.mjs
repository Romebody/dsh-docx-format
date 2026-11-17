#!/usr/bin/env node
/**
 * 采集「网上的反馈」+ 使用信号快照。
 *
 * 只做只读 GET，不写任何仓库文件（落盘交给 render.mjs）。
 * 设计要点：
 *   - traffic 端点需要 push 权限；403 时整段写 null，不影响其它数据。
 *   - 反馈（issue / 评论 / 第三方提及）按 since 增量抓，去重后交给 render 合并。
 *   - --offline 用本地 fixture 回放，便于测试与复现。
 *
 * 用法：
 *   node .github/scripts/collect.mjs --out .stats-out
 *   node .github/scripts/collect.mjs --out .stats-out --offline _test/fixtures/stats.json
 *
 * 环境变量：
 *   GH_TOKEN / GITHUB_TOKEN   访问令牌（read: repo / issues）
 *   REPO                      owner/name，默认 Romebody/dsh-docx-format
 */
import { mkdir, writeFile, readFile } from 'node:fs/promises';
import { join } from 'node:path';

const PKG = 'dsh-docx-format';
const API = 'https://api.github.com';
const NPM = 'https://api.npmjs.org';

function arg(name, dflt = undefined) {
  const i = process.argv.indexOf(`--${name}`);
  if (i === -1) return dflt;
  const v = process.argv[i + 1];
  return v && !v.startsWith('--') ? v : true;
}

const OUT = String(arg('out', '.stats-out'));
const OFFLINE = arg('offline', null);
const REPO = process.env.REPO || 'Romebody/dsh-docx-format';
const TOKEN = process.env.GH_TOKEN || process.env.GITHUB_TOKEN || '';

const warnings = [];

function isoWeek(d = new Date()) {
  const t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((t - y0) / 86400000 + 1) / 7);
  return `${t.getUTCFullYear()}-W${String(week).padStart(2, '0')}`;
}

async function getJson(url, { auth = true, allowFail = false } = {}) {
  const headers = { Accept: 'application/vnd.github+json', 'User-Agent': 'dsh-docx-format-stats' };
  if (auth && TOKEN) headers.Authorization = `Bearer ${TOKEN}`;
  const res = await fetch(url, { headers });
  if (!res.ok) {
    if (allowFail) {
      warnings.push(`${url} -> HTTP ${res.status}`);
      return null;
    }
    throw new Error(`${url} -> HTTP ${res.status} ${await res.text().catch(() => '')}`.slice(0, 300));
  }
  return res.json();
}

async function collectLive(since) {
  const out = { npm: {}, repo: {}, traffic: {}, feedback: {}, market: {} };

  for (const [key, period] of [['day', 'last-day'], ['week', 'last-week'], ['month', 'last-month']]) {
    const d = await getJson(`${NPM}/downloads/point/${period}/${PKG}`, { auth: false, allowFail: true });
    out.npm[key] = d?.downloads ?? null;
  }
  const range = await getJson(`${NPM}/downloads/range/last-month/${PKG}`, { auth: false, allowFail: true });
  out.npm.days = Array.isArray(range?.downloads) ? range.downloads.slice(-14) : [];

  const repo = await getJson(`${API}/repos/${REPO}`);
  out.repo = {
    stars: repo.stargazers_count ?? 0,
    forks: repo.forks_count ?? 0,
    issues: repo.open_issues_count ?? 0,
    watchers: repo.subscribers_count ?? 0,
    hasIssues: repo.has_issues === true,
    hasDiscussions: repo.has_discussions === true,
  };

  // traffic：需要 push 权限，403 时降级为 null
  const views = await getJson(`${API}/repos/${REPO}/traffic/views`, { allowFail: true });
  const clones = await getJson(`${API}/repos/${REPO}/traffic/clones`, { allowFail: true });
  const refs = await getJson(`${API}/repos/${REPO}/traffic/popular/referrers`, { allowFail: true });
  out.traffic = {
    views: views?.count ?? null,
    viewUniques: views?.uniques ?? null,
    clones: clones?.count ?? null,
    cloneUniques: clones?.uniques ?? null,
    referrers: Array.isArray(refs) ? refs.slice(0, 5) : null,
  };

  // 反馈：issue + 评论 + 第三方提及
  const issues = await getJson(`${API}/repos/${REPO}/issues?state=all&since=${encodeURIComponent(since)}&per_page=100`, { allowFail: true });
  const comments = await getJson(`${API}/repos/${REPO}/issues/comments?since=${encodeURIComponent(since)}&per_page=100`, { allowFail: true });
  const mentions = await getJson(`${API}/search/issues?q=${encodeURIComponent(`"${PKG}" in:title,body`)}&per_page=50`, { allowFail: true });

  const items = [];
  for (const it of Array.isArray(issues) ? issues : []) {
    items.push({
      type: it.pull_request ? 'pr' : 'issue',
      url: it.html_url,
      title: (it.title || '').slice(0, 160),
      state: it.state,
      author: it.user?.login ?? '',
      at: it.created_at,
      body: (it.body || '').replace(/\s+/g, ' ').slice(0, 300),
    });
  }
  for (const c of Array.isArray(comments) ? comments : []) {
    items.push({
      type: 'comment',
      url: c.html_url,
      title: (c.body || '').replace(/\s+/g, ' ').slice(0, 160),
      state: 'open',
      author: c.user?.login ?? '',
      at: c.created_at,
      body: '',
    });
  }
  for (const m of Array.isArray(mentions?.items) ? mentions.items : []) {
    // 自己仓库的不算“外部提及”：三种字段任一命中就跳过（search API 的字段并不总是齐全）
    const full = m.repository_url?.split('/repos/')[1] ?? m.repository?.full_name ?? '';
    const ownRepo = full.toLowerCase() === REPO.toLowerCase()
      || String(m.html_url || '').toLowerCase().includes('/' + REPO.toLowerCase() + '/');
    if (ownRepo) continue;
    // 自己在别人仓库里开的 PR / issue 也不算“外部反馈”
    if ((m.user?.login ?? '').toLowerCase() === REPO.split('/')[0].toLowerCase()) continue;
    // GitHub 搜索会把包名按连字符拆词（dsh / docx / format），
    // 光有 “DOCX … format” 也会被搜出来 —— 这里再做一次精确子串校验。
    const hay = `${m.title || ''}\n${m.body || ''}`.toLowerCase();
    if (!hay.includes(PKG.toLowerCase())) continue;
    items.push({
      type: 'mention',
      url: m.html_url,
      title: (m.title || '').slice(0, 160),
      state: m.state,
      author: m.user?.login ?? '',
      at: m.created_at,
      body: `${m.repository_url?.split('/').slice(-2).join('/') ?? ''} · ${(m.body || '').replace(/\s+/g, ' ').slice(0, 200)}`,
    });
  }

  out.feedback = { since, items, counts: countFeedback(items) };
  out.market = { listed: true, version: null };
  return out;
}

function countFeedback(items) {
  const c = { issues: 0, comments: 0, mentions: 0, prs: 0 };
  for (const i of items) {
    if (i.type === 'issue') c.issues += 1;
    else if (i.type === 'comment') c.comments += 1;
    else if (i.type === 'mention') c.mentions += 1;
    else if (i.type === 'pr') c.prs += 1;
  }
  return c;
}

async function main() {
  const now = new Date();
  const week = isoWeek(now);
  const since = new Date(now.getTime() - 8 * 86400000).toISOString(); // 覆盖一周 + 1 天缓冲

  let data;
  if (OFFLINE) {
    data = JSON.parse(await readFile(String(OFFLINE), 'utf8'));
  } else {
    data = await collectLive(since);
  }

  const row = {
    week,
    date: now.toISOString().slice(0, 10),
    npm: { day: data.npm?.day ?? null, week: data.npm?.week ?? null, month: data.npm?.month ?? null },
    repo: data.repo ?? {},
    traffic: data.traffic ?? {},
    market: data.market ?? {},
    feedback: data.feedback?.counts ?? { issues: 0, comments: 0, mentions: 0, prs: 0 },
    warnings,
  };

  await mkdir(OUT, { recursive: true });
  await writeFile(join(OUT, 'row.json'), JSON.stringify(row, null, 2), 'utf8');
  await writeFile(join(OUT, 'feedback.json'), JSON.stringify({ week, items: data.feedback?.items ?? [] }, null, 2), 'utf8');
  await writeFile(join(OUT, 'snapshot.json'), JSON.stringify(data, null, 2), 'utf8');

  const fb = row.feedback;
  const total = fb.issues + fb.comments + fb.mentions + fb.prs;
  console.log(`[stats] ${week} npm(w=${row.npm.week} m=${row.npm.month}) stars=${row.repo.stars ?? '?'} feedback=${total}`);
  if (warnings.length) console.log(`[stats] warnings: ${warnings.join('; ')}`);
}

main().catch((e) => {
  console.error('[stats] failed:', e.message);
  process.exit(1);
});
