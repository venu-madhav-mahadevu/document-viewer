// Writes public/samples/sizes.json ({ "path": bytes }) so the sample gallery can
// show exact file sizes without downloading anything. Runs before dev and build.
import { readdirSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const dir = join(dirname(fileURLToPath(import.meta.url)), '..', 'public', 'samples');
const sizes = {};
const walk = (d) => {
  for (const name of readdirSync(d)) {
    const p = join(d, name);
    if (statSync(p).isDirectory()) walk(p);
    else if (name !== 'sizes.json') sizes[relative(dir, p).split(sep).join('/')] = statSync(p).size;
  }
};
walk(dir);
writeFileSync(join(dir, 'sizes.json'), JSON.stringify(sizes, null, 2) + '\n');
console.log(`sample sizes written (${Object.keys(sizes).length} files)`);
