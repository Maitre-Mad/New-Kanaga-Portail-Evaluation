/**
 * Script de déploiement à URL Fixe / Permanente pour Kanaga Portail Évaluation.
 * Ce script pousse le code (clasp push) et met à jour TOUJOURS le même identifiant de déploiement de production.
 * L'URL publique NE CHANGE DONC JAMAIS.
 */

const { execSync } = require('child_process');

const PROD_DEPLOYMENT_ID = 'AKfycbxeoDfu8Uh9plmQXjud3N1cXUBmSAaIpgdrXQxgWEnp1jst3K3N2puD1pF3zl1XYsRntA';
const PERMANENT_URL = `https://script.google.com/macros/s/${PROD_DEPLOYMENT_ID}/exec`;

const customDesc = process.argv.slice(2).join(' ') || `Déploiement ${new Date().toLocaleDateString('fr-FR')} ${new Date().toLocaleTimeString('fr-FR')}`;

console.log('====================================================');
console.log('🔄 DÉPLOIEMENT EN COURS SUR L\'URL PERMANENTE');
console.log('====================================================');
console.log(`📌 Description : "${customDesc}"`);
console.log(`🔑 ID Déploiement Fixe : ${PROD_DEPLOYMENT_ID}\n`);

try {
  console.log('📦 1/2. Envoi des fichiers vers Google Apps Script (clasp push)...');
  execSync('npx clasp push', { stdio: 'inherit' });

  console.log('\n🚀 2/2. Mise à jour du déploiement de production...');
  execSync(`npx clasp deploy -i ${PROD_DEPLOYMENT_ID} -d "${customDesc}"`, { stdio: 'inherit' });

  console.log('\n====================================================');
  console.log('✅ DÉPLOIEMENT RÉUSSI AVEC SUCCÈS !');
  console.log('====================================================');
  console.log('🔗 VOTRE LIEN PERMANENT (NE CHANGE JAMAIS) :');
  console.log(`   ${PERMANENT_URL}`);
  console.log('====================================================\n');
} catch (error) {
  console.error('\n❌ Échec du déploiement :', error.message);
  process.exit(1);
}
