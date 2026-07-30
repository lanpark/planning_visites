Office.onReady(() => {
  // Initialisation standard
});

let loginDialog; // Variable pour garder une référence à la fenêtre

function actionOpenPlanning(event) {
  const dialogUrl = "https://lanpark.github.io/planning_visites/dialog.html";

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    {
      height: 80,
      width: 60,
      displayInIframe: false, // Ouvre une vraie fenêtre de navigation popup
      promptBeforeOpen: false
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Impossible d'ouvrir le planning : " + asyncResult.error.message);
        // En cas d'échec direct d'ouverture, on libère immédiatement le bouton
        event.completed();
      } else {
        // Le dialogue s'est ouvert avec succès !
        loginDialog = asyncResult.value;

        // TRÈS IMPORTANT : On écoute les événements de cette fenêtre
        // Notamment si l'utilisateur clique sur la croix "X" pour la fermer
        loginDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          if (arg.error === 12006) { 
            // 12006 = L'utilisateur a fermé la boîte de dialogue
            console.log("L'utilisateur a fermé la fenêtre.");
            // C'est SEULEMENT ICI qu'on signale à Outlook que tout est fini
            event.completed();
          }
        });

        // Optionnel : Si votre dialog.html envoie un message pour dire "j'ai fini"
        loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          // Si le dialogue envoie un signal de fermeture volontaire
          loginDialog.close();
          event.completed();
        });
      }
    }
  );
}

// Associe la fonction JS à l'action déclarée dans le manifest
Office.actions.associate("actionOpenPlanning", actionOpenPlanning);