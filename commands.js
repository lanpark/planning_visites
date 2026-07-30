let isOfficeReady = false;

// Sécurité : On attend que l'API Office soit totalement opérationnelle
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    isOfficeReady = true;
    console.log("Office.js est pleinement initialisé et prêt.");
  }
});

/**
 * Action déclenchée par le bouton du ruban.
 */
function actionOpenPlanning(event) {
  console.log("Clic détecté sur le bouton Planning Visites.");

  // Si l'API n'est pas encore prête, on attend un court instant avant de retenter
  if (!isOfficeReady) {
    console.warn("L'API Office n'est pas encore prête, mise en attente...");
    setTimeout(() => actionOpenPlanning(event), 500);
    return;
  }

  const dialogUrl = "https://lanpark.github.io/planning_visites/dialog.html";

  try {
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      {
        height: 80,
        width: 60,
        displayInIframe: false, // Force l'ouverture d'une vraie popup séparée
        promptBeforeOpen: false
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Échec displayDialogAsync : " + asyncResult.error.message);
          event.completed();
        } else {
          console.log("Demande d'ouverture de la popup envoyée avec succès.");
          
          const loginDialog = asyncResult.value;
          
          // On écoute la fermeture manuelle par l'utilisateur (croix X)
          loginDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            if (arg.error === 12006) {
              console.log("L'utilisateur a fermé la boîte de dialogue.");
              event.completed();
            }
          });
        }
      }
    );
  } catch (err) {
    console.error("Erreur fatale lors de l'appel au dialogue : ", err);
    event.completed();
  }
}

// Enregistrement de l'action
Office.actions.associate("actionOpenPlanning", actionOpenPlanning);