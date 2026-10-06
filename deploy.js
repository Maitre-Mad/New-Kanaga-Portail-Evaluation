/**
 * Script de déploiement universel pour Kanaga Portail Évaluation.
 * 
 * Usage :
 *   node deploy.js                -> Déploie la PRODUCTION (URL permanente)
 *   node deploy.js prod           -> Déploie la PRODUCTION
 *   node deploy.js test           -> Déploie le TEST (URL dédiée de test)
 *   node deploy.js all            -> Déploie TEST puis PRODUCTION
 * 
 * Exemples avec description personnalisée :
 *   node deploy.js test "Mise a jour module evaluation"
 *   node deploy.js prod "Version officielle v1.0.8"
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const PROD_DEPLOYMENT_ID = 'AKfycbxeoDfu8Uh9plmQXjud3N1cXUBmSAaIpgdrXQxgWEnp1jst3K3N2puD1pF3zl1XYsRntA';
const TEST_DEPLOYMENT_ID = 'AKfycbwauIpOqIbIiuEO55xTs6KOuTiqcOKlj90gr8R4sljI8UrM9EfLA4xUccfOzWKvlqn3lA';

const PROD_URL = `https://script.google.com/macros/s/${PROD_DEPLOYMENT_ID}/exec`;
const TEST_URL = `https://script.google.com/macros/s/${TEST_DEPLOYMENT_ID}/exec`;

const codeJsPath = path.join(__dirname, 'Code.js');

function setDeploymentEnv(targetEnv) {
  let content = fs.readFileSync(codeJsPath, 'utf8');
  content = content.replace(
    /var ACTIVE_DEPLOYMENT_ENV\s*=\s*['"][^'"]+['"];/,
    `var ACTIVE_DEPLOYMENT_ENV = '${targetEnv}';`
  );
  fs.writeFileSync(codeJsPath, content, 'utf8');
  console.log(`🔧 Environnement configuré dans Code.js : [${targetEnv}]`);
}

function deployTarget(env, customDesc) {
  const isTest = env === 'TEST';
  const deployId = isTest ? TEST_DEPLOYMENT_ID : PROD_DEPLOYMENT_ID;
  const url = isTest ? TEST_URL : PROD_URL;
  const tag = isTest ? 'TEST' : 'PRODUCTION';
  const desc = customDesc || `${tag} - ${new Date().toLocaleDateString('fr-FR')} ${new Date().toLocaleTimeString('fr-FR')}`;

  console.log('\n====================================================');
  console.log(`🚀 DÉPLOIEMENT ${tag} EN COURS`);
  console.log('====================================================');
  console.log(`📌 Description : "${desc}"`);
  console.log(`🔑 ID Déploiement : ${deployId}`);

  // 1. Configurer la variable d'environnement
  setDeploymentEnv(tag);

  // 2. Clasp push
  console.log('📦 Envoi des fichiers vers Google Apps Script (clasp push)...');
  execSync('npx clasp push', { stdio: 'inherit' });

  // 3. Clasp deploy vers l'identifiant cible
  console.log(`🚀 Mise à jour du déploiement ${tag}...`);
  execSync(`npx clasp deploy -i ${deployId} -d "${desc}"`, { stdio: 'inherit' });

  console.log('====================================================');
  console.log(`✅ DÉPLOIEMENT ${tag} RÉUSSI AVEC SUCCÈS !`);
  console.log(`🔗 URL : ${url}`);
  console.log('====================================================\n');
}

// Analyse des arguments
const rawArgs = process.argv.slice(2);
let mode = 'PROD';
let descArgs = [];

if (rawArgs.length > 0) {
  const first = rawArgs[0].toLowerCase();
  if (first === 'test' || first === '--test') {
    mode = 'TEST';
    descArgs = rawArgs.slice(1);
  } else if (first === 'prod' || first === 'production' || first === '--prod') {
    mode = 'PROD';
    descArgs = rawArgs.slice(1);
  } else if (first === 'all' || first === 'both' || first === '--all') {
    mode = 'ALL';
    descArgs = rawArgs.slice(1);
  } else {
    // Premier argument est déjà la description
    mode = 'PROD';
    descArgs = rawArgs;
  }
}

const desc = descArgs.join(' ').trim();

try {
  if (mode === 'ALL') {
    deployTarget('TEST', desc ? `TEST - ${desc}` : null);
    deployTarget('PRODUCTION', desc ? `PROD - ${desc}` : null);
  } else if (mode === 'TEST') {
    deployTarget('TEST', desc);
    // Restaurer le code local en mode PRODUCTION par défaut
    setDeploymentEnv('PRODUCTION');
  } else {
    deployTarget('PRODUCTION', desc);
  }
} catch (error) {
  // En cas d'erreur, s'assurer que le code local reste en mode PRODUCTION
  try { setDeploymentEnv('PRODUCTION'); } catch(e) {}
  console.error('\n❌ Échec du déploiement :', error.message);
  process.exit(1);
}
