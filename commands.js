Office.onReady(() => {
  // Rien à initialiser : ce fichier ne contient que l'action du bouton.
});

/**
 * Action déclenchée par le bouton "Planning Visites" du ruban.
 * Ouvre une fenêtre de dialogue Office qui navigue DIRECTEMENT vers
 * l'URL de lecture de l'app Power Apps (pas d'iframe, pas d'embedding),
 * ce qui contourne les restrictions d'embedding dans une app native.
 */
function actionOpenPlanning(event) {
  const dialogUrl = "https://lanpark.github.io/planning_visites/dialog.html";

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    {
      height: 80,          // % de l'écran
      width: 60,            // % de l'écran
      displayInIframe: false, // force une vraie fenêtre, pas un cadre imbriqué
      promptBeforeOpen: false
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(
          "Impossible d'ouvrir le planning : " + asyncResult.error.message
        );
      }
      // On ne garde pas de référence au dialog : il vit sa vie indépendamment.
    }
  );

  // Indispensable : signale à Outlook que l'action est terminée.
  event.completed();
}

// Associe la fonction JS à l'action déclarée dans le manifest.json
Office.actions.associate("actionOpenPlanning", actionOpenPlanning);
