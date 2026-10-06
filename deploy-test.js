/**
 * Raccourci pour déployer l'environnement de TEST de Kanaga Portail Évaluation.
 * Usage: node deploy-test.js [description]
 */

const { execSync } = require('child_process');
const customDesc = process.argv.slice(2).join(' ');
const cmd = customDesc ? `node deploy.js test "${customDesc}"` : 'node deploy.js test';

try {
  execSync(cmd, { stdio: 'inherit' });
} catch (e) {
  process.exit(1);
}
