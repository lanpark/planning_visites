// Fonction globale de secours pour attraper les erreurs de syntaxe ou d'exécution
self.onerror = function(message, source, lineno, colno, error) {
    const txt = "Erreur Globale: " + message + " à " + source + ":" + lineno;
    alert(txt); // Tente d'afficher une boîte système
    if(self.localStorage) localStorage.setItem('addin_error', txt);
    return false;
};

function actionOpenPlanning(event) {
  try {
    // Étape 1 : Vérifier si l'API Office est bien chargée dans ce contexte
    if (typeof Office === 'undefined') {
      alert("Erreur : L'objet 'Office' n'est pas défini !");
      if(event) event.completed();
      return;
    }

    var dialogUrl = "https://lanpark.github.io/planning_visites/dialog.html";

    alert("Tentative d'ouverture du dialogue..."); // Pour valider que le clic est bien reçu

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 80, width: 60, displayInIframe: false, promptBeforeOpen: false },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          const err = "Échec displayDialogAsync: " + asyncResult.error.message + " (Code: " + asyncResult.error.code + ")";
          alert(err);
          if(self.localStorage) localStorage.setItem('addin_error', err);
        } else {
          alert("Dialogue ouvert avec succès !");
        }
        // Toujours fermer l'événement à la toute fin du callback
        if(event) event.completed();
      }
    );
  } catch (e) {
    const catchErr = "Erreur dans le bloc try: " + e.message;
    alert(catchErr);
    if(self.localStorage) localStorage.setItem('addin_error', catchErr);
    if(event) event.completed();
  }
}