/**
 * Fonction Apps Script pour générer directement le rapport complet au format Google Doc natif dans Google Drive.
 * 
 * Instructions d'exécution :
 * 1. Ouvrez l'éditeur Apps Script : https://script.google.com/u/0/home/projects/13SutwXPTMEdEpQl0EbojEyS-8CxVy_NHpY65SVRthpii8wraRMxAIogM/edit
 * 2. Dans la barre supérieure, sélectionnez la fonction 'generatePlatformEvaluationReportGoogleDoc'
 * 3. Cliquez sur 'Exécuter' (Run)
 * 4. L'URL du document Google Doc créé s'affichera directement dans le journal d'exécution (Execution Log).
 */
function generatePlatformEvaluationReportGoogleDoc() {
  const docTitle = "Rapport Complet - Processus & Fonctionnalités d'Évaluation - Kanaga Consulting";
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  // Marges de page
  body.setMarginLeft(40);
  body.setMarginRight(40);
  body.setMarginTop(40);
  body.setMarginBottom(40);

  const primaryColor = '#7C4C26';   // Marron Kanaga
  const secondaryColor = '#B8860B'; // Or / Bronze
  const darkTextColor = '#2D1A0A';  // Texte foncé
  const headerBgColor = '#7C4C26';  // Fond en-têtes de tableau
  const rowAltBgColor = '#FBF9F6';  // Fond rangées alternées

  // 1. En-tête principal
  const titleP = body.appendParagraph("KANAGA CONSULTING");
  titleP.setHeading(DocumentApp.ParagraphHeading.TITLE);
  titleP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  titleP.setForegroundColor(primaryColor);
  titleP.getRuns()[0].setFontSize(22).setBold(true);

  const subP = body.appendParagraph("Portail d'Évaluation de la Performance\nRapport Complet : Processus Métier & Points Fondamentaux des Fonctionnalités");
  subP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subP.getRuns()[0].setFontSize(12).setItalic(true).setForegroundColor('#555555');

  body.appendHorizontalRule();
  body.appendParagraph("");

  // Helpers de mise en page
  function addH1(title) {
    const p = body.appendParagraph(title);
    p.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    p.setForegroundColor(primaryColor);
    p.getRuns()[0].setFontSize(14).setBold(true);
    p.setSpacingBefore(14).setSpacingAfter(4);
    return p;
  }

  function addH2(title) {
    const p = body.appendParagraph(title);
    p.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    p.setForegroundColor('#444444');
    p.getRuns()[0].setFontSize(12).setBold(true);
    p.setSpacingBefore(10).setSpacingAfter(3);
    return p;
  }

  function addBody(text) {
    const p = body.appendParagraph(text);
    p.getRuns()[0].setFontSize(10.5).setForegroundColor(darkTextColor);
    p.setSpacingAfter(4);
    p.setLineSpacing(1.15);
    return p;
  }

  function addBullet(boldPrefix, text) {
    const p = body.appendListItem("");
    p.setGlyphType(DocumentApp.GlyphType.BULLET);
    p.setSpacingAfter(3);
    p.setLineSpacing(1.15);
    const r1 = p.appendText(boldPrefix);
    r1.setBold(true).setFontSize(10.5).setForegroundColor(darkTextColor);
    const r2 = p.appendText(text);
    r2.setBold(false).setFontSize(10.5).setForegroundColor(darkTextColor);
    return p;
  }

  function styleTable(table, headerBg) {
    const headerRow = table.getRow(0);
    for (let j = 0; j < headerRow.getNumCells(); j++) {
      const cell = headerRow.getCell(j);
      cell.setBackgroundColor(headerBg || headerBgColor);
      cell.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
      const cellP = cell.getChild(0).asParagraph();
      cellP.getRuns()[0].setBold(true).setForegroundColor('#FFFFFF').setFontSize(10);
    }
    for (let i = 1; i < table.getNumRows(); i++) {
      const row = table.getRow(i);
      const bg = (i % 2 === 1) ? rowAltBgColor : '#FFFFFF';
      for (let j = 0; j < row.getNumCells(); j++) {
        const cell = row.getCell(j);
        cell.setBackgroundColor(bg);
        cell.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8).setPaddingRight(8);
        const cellP = cell.getChild(0).asParagraph();
        if (cellP.getRuns() && cellP.getRuns().length > 0) {
          cellP.getRuns()[0].setFontSize(9.5);
        }
      }
    }
  }

  // Section 1
  addH1("1. Introduction & Vision Globale");
  addBody("Le module d'évaluation du Portail Kanaga Consulting est une solution intégrée et collaborative conçue pour piloter la performance individuelle et collective des équipes. Développé sur l'écosystème Google Apps Script avec stockage sur Google Sheets et intégration Google Drive/Docs, il garantit une fluidité totale entre l'auto-évaluation du collaborateur, la contribution d'évaluateurs pairs (secondaires) et l'arbitrage managérial final.");
  addBody("Ce dispositif assure la conformité déontologique, la traçabilité des compétences, l'alignement sur les standards de performance de l'entreprise et la génération instantanée de comptes-rendus contractuels au format PDF.");

  // Section 2
  addH1("2. Acteurs & Rôles dans le Processus");
  addBody("Le système articule 4 rôles opérationnels distincts :");
  const rolesTable = body.appendTable([
    ["Rôle", "Responsabilités & Droits dans le Système"],
    ["Collaborateur (Évalué)", "Remplit son auto-évaluation (notation, commentaires, réalisations de la période, souhaits de formation et aspirations). Accède à l'historique de ses évaluations et télécharge ses comptes-rendus PDF."],
    ["Évaluateur(s) Secondaire(s)", "Pairs, chefs de missions ou collaborateurs mandatés. Fournissent une appréciation et une notation intermédiaire avant l'entretien final. Disposent de sauvegardes de brouillons indépendantes."],
    ["Évaluateur Principal (Manager)", "Initie l'évaluation, pilote l'entretien d'évaluation, consulte les avis du collaborateur et des secondaires en miroir, réalise la synthèse finale, fixe les objectifs SMART et clôture le dossier."],
    ["Administrateur RH / IT", "Administre le Form Builder dynamique (ajout de profils métiers, questions, pagination), personnalise les modèles d'e-mails automatiques, gère les habilitations et supervise les logs d'activité."]
  ]);
  styleTable(rolesTable);

  body.appendParagraph("");

  // Section 3
  addH1("3. Cycle de Vie & Workflow d'Évaluation de Bout en Bout");
  
  addH2("Phase 1 : L'Initiation du Dossier");
  addBullet("Déclencheur : ", "Le Manager ou l'Administrateur démarre le processus depuis le Portail Manager.");
  addBullet("Paramètres saisis : ", "Période d'évaluation, sélection des collaborateurs (support d'initiation en lot), profil de poste, évaluateur principal et évaluateurs secondaires.");
  addBullet("Statut : ", "Initiée.");
  addBullet("Automatismes : ", "Notification par e-mail au collaborateur avec lien vers son auto-évaluation, alerte aux secondaires et confirmation au manager.");

  addH2("Phase 2 : L'Auto-évaluation du Collaborateur");
  addBullet("Accès sécurisé : ", "Le collaborateur accède à son espace 'Mes Évaluations'. Les données de cadrage sont verrouillées en lecture seule.");
  addBullet("Saisie du collaborateur : ", "Évaluation des compétences fondamentales, des compétences spécifiques et bilan de conclusion.");
  addBullet("Brouillon : ", "Sauvegarde intermédiaire autonome sans notification.");
  addBullet("Statut après soumission : ", "'Attente Évaluateur Secondaire' ou 'Attente Évaluateur Principal'.");

  addH2("Phase 3 : L'Évaluation Secondaire (Multi-évaluateurs)");
  addBullet("Accès partagé : ", "Chaque secondaire accède au dossier et voit l'auto-évaluation du collaborateur.");
  addBullet("Brouillons individuels : ", "Chaque secondaire dispose d'une sauvegarde isolée.");
  addBullet("Transition : ", "Passage automatique au manager principal quand tous les secondaires ont soumis, ou avance manuelle par le bouton 'Passer au Principal →'.");

  addH2("Phase 4 : L'Évaluation Managériale & Entretien d'Évaluation");
  addBullet("Vue comparative miroir : ", "Affichage côte à côte des notes et commentaires de l'employé et des secondaires.");
  addBullet("Pré-évaluation pour réunion : ", "Le manager peut préparer ses notes en amont sans clôturer le dossier.");
  addBullet("Retour aux secondaires : ", "Possibilité de renvoyer le dossier aux secondaires avec conservation intégrale du brouillon manager.");
  addBullet("Validation finale : ", "Notation arbitrée, objectifs SMART futurs et appréciation globale.");

  addH2("Phase 5 : Clôture, Génération PDF & Archivage");
  addBullet("Statut : ", "Complétée.");
  addBullet("Génération PDF : ", "Création automatique du compte-rendu officiel avec tableaux comparatifs et barème.");
  addBullet("Archivage Drive : ", "Stockage dans le dossier Drive sécurisé et lien inséré dans le dossier.");
  addBullet("Notification finale : ", "Envoi des e-mails avec lien direct vers le compte-rendu PDF.");

  body.appendParagraph("");

  // Section 4
  addH1("4. Architecture du Formulaire d'Évaluation (L'Entonnoir en 4 Étapes)");
  const structTable = body.appendTable([
    ["Étape / Section", "Intitulé", "Contenu & Rôle Fonctionnel"],
    ["Étape 1", "Section 0 : Entonnoir de départ", "Informations générales : Période d'évaluation, employé(s) ciblé(s), profil de poste, désignation de l'évaluateur principal et des évaluateurs secondaires."],
    ["Étape 2", "Section 1 : Tronc Commun", "8 compétences fondamentales partagées par l'ensemble des collaborateurs de Kanaga Consulting (Éthique, Communication, Travail d'équipe, Gestion du temps, Initiative, Adaptabilité, Contexte local, Innovation)."],
    ["Étape 3", "Section 2 : Aiguillage Profil", "Compétences spécifiques au profil métier avec pagination dynamique (Junior Privé, Public, Consultant, Chef de Projet, Comptable, Auditeur, etc.)."],
    ["Étape 4", "Section 3 : Conclusion Commune", "Objectifs passés, points forts, axes d'amélioration, note globale, besoins de formation, objectifs SMART futurs, aspirations professionnelles et commentaires libres."]
  ]);
  styleTable(structTable);

  body.appendParagraph("");

  // Section 5
  addH1("5. Le Barème Harmonisé d'Évaluation");
  const scaleTable = body.appendTable([
    ["Niveau / Sigle", "Intitulé", "Définition & Critère d'Attribution"],
    ["5 - EXCEPTI", "Performance Exceptionnelle", "Dépasse constamment les attentes et constitue une référence pour l'organisation."],
    ["4 - SUPERIE", "Performance Supérieure", "Dépasse régulièrement les attentes dans les domaines clés de la mission."],
    ["3 - CONFORM", "Performance Conforme aux Attentes", "Atteint les objectifs et les standards de performance attendus pour le poste."],
    ["2 - AMELIOR", "Performance à Améliorer", "N'atteint pas toujours les objectifs et nécessite une amélioration dans certains domaines."],
    ["1 - INSATIS", "Performance Insatisfaisante", "N'atteint pas les objectifs de manière significative ; nécessite une amélioration immédiate."],
    ["NONAPPL", "N/A Non Applicable", "Non applicable au poste ou non évaluable sur la période de référence."]
  ]);
  styleTable(scaleTable);

  body.appendParagraph("");

  // Section 6
  addH1("6. Points Fondamentaux & Fonctionnalités Techniques Clés");
  addBullet("Form Builder dynamique : ", "Gestion flexible des formulaires, pages et questions via Google Sheets 'EvaluationQuestions'.");
  addBullet("Centre de notifications paramétrable : ", "5 modèles d'e-mails HTML personnalisables avec balises dynamiques ({nom_employe}, {periode}, {profil}, etc.).");
  addBullet("Brouillons multi-niveaux : ", "Isolation des données de saisie entre collaborateur, évaluateurs secondaires et manager.");
  addBullet("Compte-rendu PDF automatique : ", "Génération Google Docs vers PDF avec archivage Drive et mise en page corporate.");
  addBullet("Traçabilité et journalisation : ", "Journalisation systématique dans Cloud Logging et dans l'onglet 'Logs'.");

  doc.saveAndClose();

  // Déplacement optionnel vers le dossier Drive de l'application
  try {
    const folderId = '10y8QMWCUlWL1frurah6lXg2d8j1ukGlj';
    const folder = DriveApp.getFolderById(folderId);
    const docFile = DriveApp.getFileById(doc.getId());
    docFile.moveTo(folder);
  } catch (err) {
    Logger.log("Info: Le fichier est conservé à la racine de Google Drive (" + err.message + ")");
  }

  const fileUrl = doc.getUrl();
  Logger.log("✅ Document Google Doc créé avec succès !");
  Logger.log("Lien du Google Doc : " + fileUrl);
  return fileUrl;
}
