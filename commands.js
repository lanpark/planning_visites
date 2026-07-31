Office.onReady(() => {
  // Le framework Office de la commande principale est prêt
  console.log("CommandsRuntime prêt.");
});

function actionOpenPlanning(event) {
  // Ouvre une page intermédiaire sur GitHub Pages, puis redirige vers Power Apps.
  const dialogUrl = "https://lanpark.github.io/planning_visites/dialog.html";
  console.log("ActionOpenPlanning appelée");

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    {
      height: 80,
      width: 60,
      displayInIframe: false,
      promptBeforeOpen: false
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Échec de l'ouverture du dialogue : " + asyncResult.error.message);
        event.completed();
        return;
      }

      console.log("Dialogue ouvert avec succès.");
      event.completed();
    }
  );
}

// Liaison de la fonction au manifeste
Office.actions.associate("actionOpenPlanning", actionOpenPlanning);