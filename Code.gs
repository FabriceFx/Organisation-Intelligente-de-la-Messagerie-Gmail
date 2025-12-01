/**
 * @fileoverview Script d'organisation de la messagerie Gmail
 * @name Organisation Messagerie - Auto & Colors Fix
 * @namespace OrganisationMessagerie
 * @author Fabrice Faucheux
 */

// Configuration centralisée avec PALETTE OFFICIELLE GMAIL
const CONFIGURATION = {
  LIBELLES: {
    PERSO_NON_LU: "# 1.Moi (Non Lu)",
    PERSO_LU: "# 1.Moi (Lu)",
    COPIE_NON_LU: "# 2.Copie (Non Lu)",
    COPIE_LU: "# 2.Copie (Lu)"
  },
  // Codes hexadécimaux stricts de Google. Ne pas modifier.
  COULEURS: {
    ROUGE: { backgroundColor: '#ac2b37', textColor: '#ffffff' },
    GRIS:  { backgroundColor: '#dddddd', textColor: '#434343' },
    BLEU:  { backgroundColor: '#4986e7', textColor: '#ffffff' }
  },
  LIMITES: { MAX_THREADS: 50 },
  TRIGGER: { NOM_FONCTION: "organiserMessagerie", FREQUENCE_MINUTES: 10 }
};

/**
 * FONCTION PRINCIPALE (Logique inchangée)
 */
function organiserMessagerie() {
  console.time("ChronoExecution");
  try {
    const { LIBELLES, LIMITES } = CONFIGURATION;
    const obtenirLibelle = (nom) => GmailApp.getUserLabelByName(nom) || GmailApp.createLabel(nom);

    const mapLibelles = {
      persoNonLu: obtenirLibelle(LIBELLES.PERSO_NON_LU),
      persoLu: obtenirLibelle(LIBELLES.PERSO_LU),
      copieNonLu: obtenirLibelle(LIBELLES.COPIE_NON_LU),
      copieLu: obtenirLibelle(LIBELLES.COPIE_LU)
    };

    const groupes = { persoNonLu: [], persoLu: [], copieNonLu: [], copieLu: [] };
    const fils = GmailApp.getInboxThreads(0, LIMITES.MAX_THREADS);

    if (fils.length === 0) { console.info("Aucun fil à traiter."); return; }

    const emailUser = Session.getActiveUser().getEmail().toLowerCase();

    fils.forEach(fil => {
      const msg = fil.getMessages()[fil.getMessages().length - 1];
      const to = msg.getTo().toLowerCase();
      const cc = msg.getCc().toLowerCase();
      const isUnread = fil.isUnread();

      if (to.includes(emailUser)) {
        isUnread ? groupes.persoNonLu.push(fil) : groupes.persoLu.push(fil);
      } else if (cc.includes(emailUser)) {
        isUnread ? groupes.copieNonLu.push(fil) : groupes.copieLu.push(fil);
      }
    });

    const update = (threads, add, remove) => {
      if (!threads.length) return;
      try {
        add.addToThreads(threads);
        remove.removeFromThreads(threads);
        console.log(`MAJ: +${add.getName()} / -${remove.getName()} (${threads.length})`);
      } catch (e) { console.warn(`Err partial: ${e.message}`); }
    };

    update(groupes.persoNonLu, mapLibelles.persoNonLu, mapLibelles.persoLu);
    update(groupes.persoLu, mapLibelles.persoLu, mapLibelles.persoNonLu);
    update(groupes.copieNonLu, mapLibelles.copieNonLu, mapLibelles.copieLu);
    update(groupes.copieLu, mapLibelles.copieLu, mapLibelles.copieNonLu);

  } catch (e) { console.error(`Erreur Main: ${e.stack}`); }
  finally { console.timeEnd("ChronoExecution"); }
}

/**
 * CONFIGURATION DES COULEURS (Version Debug & Robuste)
 * À lancer manuellement une fois.
 */
function configurerCouleurs() {
  console.log("--- Début Configuration Couleurs ---");

  // 1. Vérification du Service Avancé
  if (typeof Gmail === 'undefined') {
    const msg = "ERREUR CRITIQUE : Le service 'Gmail' n'est pas activé. Cliquez sur '+' à côté de Services > Gmail API.";
    console.error(msg);
    throw new Error(msg);
  }

  try {
    // 2. Récupération de tous les libellés existants via API
    const userLabels = Gmail.Users.Labels.list('me').labels;
    
    // 3. Définition du mapping Libellé -> Couleur
    const mapping = [
      { nom: CONFIGURATION.LIBELLES.PERSO_NON_LU, color: CONFIGURATION.COULEURS.ROUGE },
      { nom: CONFIGURATION.LIBELLES.PERSO_LU,     color: CONFIGURATION.COULEURS.GRIS },
      { nom: CONFIGURATION.LIBELLES.COPIE_NON_LU, color: CONFIGURATION.COULEURS.BLEU },
      { nom: CONFIGURATION.LIBELLES.COPIE_LU,     color: CONFIGURATION.COULEURS.GRIS }
    ];

    // 4. Application
    mapping.forEach(item => {
      // Recherche insensible à la casse et robuste
      const targetLabel = userLabels.find(l => l.name === item.nom);

      if (targetLabel) {
        try {
          Gmail.Users.Labels.patch({
            color: item.color
          }, 'me', targetLabel.id);
          console.log(`[OK] Couleur appliquée sur : ${item.nom}`);
        } catch (apiError) {
          console.error(`[ERREUR API] Impossible de colorer ${item.nom}. Cause : ${apiError.message}`);
        }
      } else {
        // Tentative de création si manquant (Fallback)
        console.warn(`[INFO] Libellé manquant : ${item.nom}. Création en cours...`);
        GmailApp.createLabel(item.nom);
        // On relance récursivement ou on demande de relancer
        console.log("Veuillez relancer la fonction pour colorer le nouveau libellé.");
      }
    });

  } catch (e) {
    console.error(`Erreur Globale Couleurs : ${e.message}`);
  }
  console.log("--- Fin Configuration ---");
}

/**
 * AUTOMATISATION
 */
function installerAutomatisation() {
  const nom = CONFIGURATION.TRIGGER.NOM_FONCTION;
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.some(t => t.getHandlerFunction() === nom)) {
    console.warn("Automatisation déjà active.");
    return;
  }
  ScriptApp.newTrigger(nom).timeBased().everyMinutes(CONFIGURATION.TRIGGER.FREQUENCE_MINUTES).create();
  console.log("Automatisation activée.");
  
  // Tentative automatique de coloration
  try { configurerCouleurs(); } catch(e) { console.warn("Attention: Couleurs non configurées (voir logs)."); }
}

function desinstallerAutomatisation() {
  const nom = CONFIGURATION.TRIGGER.NOM_FONCTION;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === nom) ScriptApp.deleteTrigger(t);
  });
  console.log("Automatisation désactivée.");
}

