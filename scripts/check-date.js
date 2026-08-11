const now = new Date();
const parts = new Intl.DateTimeFormat('en-CA', {
  timeZone: 'America/Santiago',
  year: 'numeric', month: '2-digit', day: '2-digit'
}).formatToParts(now);
const o = {};
for (const p of parts) o[p.type] = p.value;
const date = `${o.year}-${o.month}-${o.day}`;

const ok = date >= '2026-08-10' && date <= '2026-09-11';

console.log(`Fecha Chile: ${date} | actualizar=${ok}`);

if (process.env.GITHUB_OUTPUT) {
  require('fs').appendFileSync(process.env.GITHUB_OUTPUT, `run=${ok}\n`);
}
