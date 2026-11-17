#!/usr/bin/env node
/**
 * 把本周采集结果并进 stats 分支的账本，并重新生成报表与徽章。
 *
 * 规则：
 *   1) history.jsonl 按 week 去重覆盖（同一周跑多次只留一行）。
 *   2) 只保留最近 13 周的周行；更早的周行合并成一条 monthly 行（永久保留）。
 *   3) feedback.json 按 url 去重累积，永久保留（反馈本身不清）。
 *   4) 生成的 README.md 只展示最近 13 周 + 月度表，避免文件无限变长。
 *
 * 用法：
 *   node .github/scripts/render.mjs --row .stats-out/row.json --feedback .stats-out/feedback.json --stats-dir <stats 分支工作区>
 */
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { join } from 'node:path';

function arg(name, dflt = undefined) {
  const i = process.argv.indexOf(`--${name}`);
  if (i === -1) return dflt;
  const v = process.argv[i + 1];
  return v && !v.startsWith('--') ? v : true;
}

const ROW = String(arg('row', '.stats-out/row.json'));
const FB = String(arg('feedback', '.stats-out/feedback.json'));
const DIR = String(arg('stats-dir', '.'));

const KEEP_WEEKS = 13;

async function readJsonl(path) {
  if (!existsSync(path)) return [];
  const text = await readFile(path, 'utf8');
  const rows = [];
  for (const line of text.split('\n')) {
    const s = line.trim();
    if (!s) continue;
    try {
      rows.push(JSON.parse(s));
    } catch {
      /* 坏行跳过，不中断 */
    }
  }
  return rows;
}

function num(v) {
  return typeof v === 'number' && Number.isFinite(v) ? v : null;
}

function fmt(v) {
  return v === null || v === undefined ? '—' : String(v);
}

function spark(values) {
  const nums = values.filter((v) => typeof v === 'number');
  if (!nums.length) return '';
  const max = Math.max(...nums, 1);
  const chars = '▁▂▃▄▅▆▇█';
  return values
    .map((v) => (typeof v !== 'number' ? ' ' : chars[Math.min(chars.length - 1, Math.round((v / max) * (chars.length - 1)))]))
    .join('');
}

function monthOf(dateStr) {
  return (dateStr || '').slice(0, 7);
}

async function main() {
  const row = JSON.parse(await readFile(ROW, 'utf8'));
  const fb = existsSync(FB) ? JSON.parse(await readFile(FB, 'utf8')) : { items: [] };

  await mkdir(DIR, { recursive: true });
  const historyPath = join(DIR, 'history.jsonl');
  const payloadPath = join(DIR, 'feedback.json');

  // 1) 周行去重覆盖
  const rows = await readJsonl(historyPath);
  const byWeek = new Map();
  for (const r of rows) if (r.week) byWeek.set(r.week, r);
  byWeek.set(row.week, row);

  // 2) 只留最近 KEEP_WEEKS 周，更早的折成月度行
  const weekly = [...byWeek.values()].filter((r) => r.kind !== 'month').sort((a, b) => (a.week < b.week ? -1 : 1));
  const months = new Map(
    [...byWeek.values()].filter((r) => r.kind === 'month').map((r) => [r.month, r]),
  );

  const keep = weekly.slice(-KEEP_WEEKS);
  const fold = weekly.slice(0, Math.max(0, weekly.length - KEEP_WEEKS));
  for (const r of fold) {
    const m = monthOf(r.date);
    if (!m) continue;
    const prev = months.get(m);
    const add = (a, b) => (num(a) ?? 0) + (num(b) ?? 0);
    months.set(m, {
      kind: 'month',
      month: m,
      npm: { week: r.npm?.week ?? null, month: r.npm?.month ?? null, day: null },
      repo: r.repo ?? {},
      traffic: {
        views: add(prev?.traffic?.views, r.traffic?.views),
        clones: add(prev?.traffic?.clones, r.traffic?.clones),
      },
      feedback: {
        issues: add(prev?.feedback?.issues, r.feedback?.issues),
        comments: add(prev?.feedback?.comments, r.feedback?.comments),
        mentions: add(prev?.feedback?.mentions, r.feedback?.mentions),
        prs: add(prev?.feedback?.prs, r.feedback?.prs),
      },
      weeks: (prev?.weeks ?? 0) + 1,
    });
  }

  const finalRows = [...[...months.values()].sort((a, b) => (a.month < b.month ? -1 : 1)), ...keep];
  await writeFile(historyPath, finalRows.map((r) => JSON.stringify(r)).join('\n') + '\n', 'utf8');

  // 3) 反馈累积去重
  const known = new Map();
  if (existsSync(payloadPath)) {
    try {
      const prev = JSON.parse(await readFile(payloadPath, 'utf8'));
      for (const it of prev.items ?? []) if (it.url) known.set(it.url, it);
    } catch {
      /* 坏文件按空处理 */
    }
  }
  const fresh = [];
  for (const it of fb.items ?? []) {
    if (!it?.url) continue;
    if (known.has(it.url)) continue;
    known.set(it.url, { ...it, firstSeen: row.week });
    fresh.push(it);
  }
  const allFb = [...known.values()].sort((a, b) => String(b.at).localeCompare(String(a.at)));
  await writeFile(payloadPath, JSON.stringify({ updatedAt: new Date().toISOString(), items: allFb }, null, 2) + '\n', 'utf8');

  // 4) 报表
  const recent = keep.slice(-KEEP_WEEKS);
  const lines = [];
  lines.push('# 下载与反馈');
  lines.push('');
  lines.push('> 由 `.github/workflows/stats.yml` 每周自动更新，**请勿手改**。');
  lines.push('> 原始采集件在 Actions artifacts 里保留 30 天后自动删除；本文件只留聚合值。');
  lines.push('');
  lines.push(`更新于 ${row.date}（${row.week}）`);
  lines.push('');

  lines.push('## 最近 13 周');
  lines.push('');
  lines.push('| 周 | npm 周下载 | npm 月下载 | stars | views | clones | issue | 评论 | 外部提及 |');
  lines.push('| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |');
  for (const r of recent) {
    lines.push(
      `| ${r.week} | ${fmt(r.npm?.week)} | ${fmt(r.npm?.month)} | ${fmt(r.repo?.stars)} | ${fmt(r.traffic?.views)} | ${fmt(r.traffic?.clones)} | ${fmt(r.feedback?.issues)} | ${fmt(r.feedback?.comments)} | ${fmt(r.feedback?.mentions)} |`,
    );
  }
  lines.push('');

  const series = recent.map((r) => num(r.npm?.week));
  if (series.some((v) => v !== null)) {
    lines.push('```text');
    lines.push(`npm 周下载走势（近 ${recent.length} 周）`);
    lines.push(spark(series));
    lines.push('```');
    lines.push('');
  }

  const monthRows = [...months.values()].sort((a, b) => (a.month < b.month ? -1 : 1));
  if (monthRows.length) {
    lines.push('## 月度（周行滚动 13 周后折成月行，永久保留）');
    lines.push('');
    lines.push('| 月 | 周数 | 期末 npm 月下载 | stars | views 合计 | clones 合计 | issue | 评论 | 外部提及 |');
    lines.push('| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |');
    for (const r of monthRows) {
      lines.push(
        `| ${r.month} | ${fmt(r.weeks)} | ${fmt(r.npm?.month)} | ${fmt(r.repo?.stars)} | ${fmt(r.traffic?.views)} | ${fmt(r.traffic?.clones)} | ${fmt(r.feedback?.issues)} | ${fmt(r.feedback?.comments)} | ${fmt(r.feedback?.mentions)} |`,
      );
    }
    lines.push('');
  }

  lines.push('## 反馈明细');
  lines.push('');
  if (!allFb.length) {
    lines.push('目前 **0 条**：Issues、Discussions、第三方提及都为空。有反馈时这里会自动列出来。');
  } else {
    lines.push(`共 ${allFb.length} 条（按时间倒序）：`);
    lines.push('');
    lines.push('| 类型 | 标题 | 作者 | 时间 |');
    lines.push('| --- | --- | --- | --- |');
    for (const it of allFb.slice(0, 50)) {
      const title = String(it.title || '').replace(/\|/g, '\\|').slice(0, 70);
      lines.push(`| ${it.type} | [${title}](${it.url}) | @${it.author || '?'} | ${String(it.at || '').slice(0, 10)} |`);
    }
    if (allFb.length > 50) lines.push('', `（只显示最近 50 条，完整列表见 \`feedback.json\`）`);
  }
  lines.push('');
  lines.push('## 口径');
  lines.push('');
  lines.push('- **npm 下载量**含镜像、CI、爬虫，只看趋势，不当用户数。');
  lines.push('- **views / clones** 是 GitHub 自己统计的 14 天滚动值，需 push 权限，取不到时显示 —。');
  lines.push('- **外部提及**指别的仓库在 issue/PR 正文里提到本包（自己仓库的不算）。');
  lines.push('- 本周无反馈时不写任何反馈记录，只保留这一行快照。');
  lines.push('');

  await writeFile(join(DIR, 'README.md'), lines.join('\n'), 'utf8');

  // 5) 徽章（shields endpoint 格式）
  const weekDl = num(row.npm?.week);
  const total = (row.feedback?.issues ?? 0) + (row.feedback?.comments ?? 0) + (row.feedback?.mentions ?? 0);
  const badge = {
    schemaVersion: 1,
    label: 'npm downloads / week',
    message: weekDl === null ? 'n/a' : String(weekDl),
    color: weekDl === null ? 'lightgrey' : 'blue',
  };
  await writeFile(join(DIR, 'badge.json'), JSON.stringify(badge, null, 2) + '\n', 'utf8');

  const fbBadge = {
    schemaVersion: 1,
    label: 'open feedback',
    message: String(total),
    color: total > 0 ? 'green' : 'lightgrey',
  };
  await writeFile(join(DIR, 'badge-feedback.json'), JSON.stringify(fbBadge, null, 2) + '\n', 'utf8');

  console.log(`[stats] rows=${finalRows.length} weeks=${keep.length} months=${months.size} feedback=${allFb.length} new=${fresh.length}`);
}

main().catch((e) => {
  console.error('[stats] render failed:', e.message);
  process.exit(1);
});
