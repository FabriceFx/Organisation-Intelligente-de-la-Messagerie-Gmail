/**
 * @fileoverview Script d'organisation de la messagerie Gmail avec gestion d'exclusivité, automatisation et coloration.
 * @name Organisation Messagerie - Auto & Colors
 * @namespace OrganisationMessagerie
 * @version 3.0
 * @author Fabrice Faucheux
 */

// Configuration centralisée
const CONFIGURATION = {
  LIBELLES: {
    PERSO_NON_LU: "# 1.Moi (Non Lu)",
    PERSO_LU: "# 1.Moi (Lu)",
    COPIE_NON_LU: "# 2.Copie (Non Lu)",
    COPIE_LU: "# 2.Copie (Lu)"
  },
  // Palette de couleurs (Format API Gmail)
  COULEURS: {
    PERSO_NON_LU: { backgroundColor: "#ac2b37", textColor: "#ffffff" }, // Rouge Vif
    PERSO_LU:     { backgroundColor: "#dddddd", textColor: "#434343" }, // Gris
    COPIE_NON_LU: { backgroundColor: "#4986e7", textColor: "#ffffff" }, // Bleu
    COPIE_LU:     { backgroundColor: "#e1e1e1", textColor: "#666666" }  // Gris Clair
  },
  LIMITES: {
    MAX_THREADS: 50
  },
  TRIGGER: {
    NOM_FONCTION: "organiserMessagerie",
    FREQUENCE_MINUTES: 10
  }
};

/**
 * FONCTION PRINCIPALE
 * Analyse les fils récents et met à jour les libellés et statuts.
 */
function organiserMessagerie() {
  console.time("ChronoExecution");

  try {
    const { LIBELLES, LIMITES } = CONFIGURATION;
    
    // Helper pour garantir l'existence du libellé
    const obtenirLibelle = (nom) => GmailApp.getUserLabelByName(nom) || GmailApp.createLabel(nom);

    const mapLibelles = {
      persoNonLu: obtenirLibelle(LIBELLES.PERSO_NON_LU),
      persoLu: obtenirLibelle(LIBELLES.PERSO_LU),
      copieNonLu: obtenirLibelle(LIBELLES.COPIE_NON_LU),
      copieLu: obtenirLibelle(LIBELLES.COPIE_LU)
    };

    const groupesTraitement = {
      persoNonLu: [], persoLu: [],
      copieNonLu: [], copieLu: []
    };

    const filsDiscussion = GmailApp.getInboxThreads(0, LIMITES.MAX_THREADS);

    if (filsDiscussion.length === 0) {
      console.info("Aucun fil de discussion à traiter.");
      return;
    }

    const emailUtilisateur = Session.getActiveUser().getEmail().toLowerCase();

    // Classification
    filsDiscussion.forEach(fil => {
      const messages = fil.getMessages();
      const dernierMessage = messages[messages.length - 1];
      
      const destinatairesDirects = dernierMessage.getTo().toLowerCase();
      const destinatairesCopie = dernierMessage.getCc().toLowerCase();
      const estNonLu = fil.isUnread();

      const estDestinataire = destinatairesDirects.includes(emailUtilisateur);
      const estEnCopie = destinatairesCopie.includes(emailUtilisateur);

      if (estDestinataire) {
        estNonLu ? groupesTraitement.persoNonLu.push(fil) : groupesTraitement.persoLu.push(fil);
      } else if (estEnCopie) {
        estNonLu ? groupesTraitement.copieNonLu.push(fil) : groupesTraitement.copieLu.push(fil);
      }
    });

    // Application par lots
    const { persoNonLu, persoLu, copieNonLu, copieLu } = groupesTraitement;
    const { persoNonLu: lblPersoNonLu, persoLu: lblPersoLu, copieNonLu: lblCopieNonLu, copieLu: lblCopieLu } = mapLibelles;

    appliquerChangementLibelles(persoNonLu, lblPersoNonLu, lblPersoLu);
    appliquerChangementLibelles(persoLu, lblPersoLu, lblPersoNonLu);
    appliquerChangementLibelles(copieNonLu, lblCopieNonLu, lblCopieLu);
    appliquerChangementLibelles(copieLu, lblCopieLu, lblCopieNonLu);

  } catch (erreur) {
    console.error(`Erreur critique : ${erreur.stack}`);
  } finally {
    console.timeEnd("ChronoExecution");
  }
}

/**
 * Utilitaire : Applique les modifications de libellés (Batch).
 */
const appliquerChangementLibelles = (groupeFils, libelleAAppliquer, libelleANettoyer) => {
  if (!groupeFils || groupeFils.length === 0) return;
  try {
    libelleAAppliquer.addToThreads(groupeFils);
    libelleANettoyer.removeFromThreads(groupeFils);
    console.log(`Mise à jour : +[${libelleAAppliquer.getName()}] / -[${libelleANettoyer.getName()}] (${groupeFils.length} fils)`);
  } catch (e) {
    console.warn(`Erreur partielle : ${e.message}`);
  }
};

/**
 * CONFIGURATION VISUELLE (Nécessite le Service Avancé "Gmail")
 * Applique les couleurs définies dans la configuration aux libellés.
 * À exécuter manuellement une seule fois ou lors de l'installation.
 */
function configurerCouleurs() {
  console.log("Début de la configuration des couleurs...");
  
  try {
    // Vérification de la présence du service avancé
    if (typeof Gmail === 'undefined') {
      throw new Error("Le service avancé 'Gmail' n'est pas activé. Veuillez l'ajouter via le menu 'Services'.");
    }

    const listeLibelles = Gmail.Users.Labels.list('me').labels;
    const definitionsCouleurs = [
      { nom: CONFIGURATION.LIBELLES.PERSO_NON_LU, couleur: CONFIGURATION.COULEURS.PERSO_NON_LU },
      { nom: CONFIGURATION.LIBELLES.PERSO_LU, couleur: CONFIGURATION.COULEURS.PERSO_LU },
      { nom: CONFIGURATION.LIBELLES.COPIE_NON_LU, couleur: CONFIGURATION.COULEURS.COPIE_NON_LU },
      { nom: CONFIGURATION.LIBELLES.COPIE_LU, couleur: CONFIGURATION.COULEURS.COPIE_LU }
    ];

    definitionsCouleurs.forEach(def => {
      const libelleTrouve = listeLibelles.find(l => l.name === def.nom);
      
      if (libelleTrouve) {
        Gmail.Users.Labels.patch({
          color: def.couleur
        }, 'me', libelleTrouve.id);
        console.log(`Couleur appliquée pour : ${def.nom}`);
      } else {
        console.warn(`Libellé introuvable : ${def.nom}. Lancez d'abord 'organiserMessagerie' pour les créer.`);
      }
    });

  } catch (e) {
    console.error(`Erreur configuration couleurs : ${e.message}`);
  }
}

/**
 * GESTION DE L'AUTOMATISATION
 * Installe le déclencheur temporel.
 */
function installerAutomatisation() {
  try {
    const nomFonction = CONFIGURATION.TRIGGER.NOM_FONCTION;
    const frequence = CONFIGURATION.TRIGGER.FREQUENCE_MINUTES;
    const declencheursExistants = ScriptApp.getProjectTriggers();
    
    if (declencheursExistants.some(trigger => trigger.getHandlerFunction() === nomFonction)) {
      console.warn(`Automatisation déjà active.`);
      return;
    }

    ScriptApp.newTrigger(nomFonction).timeBased().everyMinutes(frequence).create();
    console.log(`Succès : Automatisation installée (${frequence} min).`);
    
    // Tentative de configuration des couleurs après l'installation
    configurerCouleurs();

  } catch (erreur) {
    console.error(`Erreur installation : ${erreur.message}`);
  }
}

/**
 * NETTOYAGE
 * Supprime les déclencheurs.
 */
function desinstallerAutomatisation() {
  const nomFonction = CONFIGURATION.TRIGGER.NOM_FONCTION;
  const declencheurs = ScriptApp.getProjectTriggers();
  let compteur = 0;
  
  declencheurs.forEach(trigger => {
    if (trigger.getHandlerFunction() === nomFonction) {
      ScriptApp.deleteTrigger(trigger);
      compteur++;
    }
  });
  console.log(`Nettoyage : ${compteur} déclencheur(s) supprimé(s).`);
}
