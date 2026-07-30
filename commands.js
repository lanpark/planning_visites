Office.onReady(() => {
  // Le framework Office de la commande principale est prêt
  console.log("CommandsRuntime prêt.");
});

function actionOpenPlanning(event) {
	
  // Ton URL Power Apps exacte, appelée DIRECTEMENT
  const powerAppsUrl = "https://apps.powerapps.com/play/e/default-1142a072-6d78-4789-a846-e69a5abb61b4/a/415f9d61-b3be-418d-9421-869452d86e9b?tenantId=1142a072-6d78-4789-a846-e69a5abb61b4";
	console.log("Action appelée");
  Office.context.ui.displayDialogAsync(
    powerAppsUrl,
    {
      height: 85,             // Prend une bonne partie de l'écran
      width: 65,
      displayInIframe: false, // OBLIGATOIRE : ouvre une vraie fenêtre pop-up, pas d'iframe
      promptBeforeOpen: false
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Échec de l'ouverture du dialogue : " + asyncResult.error.message);
        // Libère le bouton Outlook si ça échoue au lancement
        event.completed();
        return;
      }

      // L'ouverture a réussi, on récupère le contrôle de la fenêtre de dialogue
      const dialog = asyncResult.value;

      // On écoute l'événement de fermeture pour savoir quand l'utilisateur a fini
      dialog.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
        if (args.error === 12006) { // 12006 = L'utilisateur a cliqué sur la croix (X)
          console.log("Fenêtre Power Apps fermée par l'utilisateur.");
          // C'est seulement à ce moment précis qu'on dit à Outlook que l'action est terminée !
          event.completed();
        }
      });
    }
  );
}

// Liaison de la fonction au manifeste
Office.actions.associate("actionOpenPlanning", actionOpenPlanning);