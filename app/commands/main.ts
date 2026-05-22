Office.onReady(() => {
  // Reserved for future ribbon commands.
});

function action(event: Office.AddinCommands.Event) {
  event.completed();
}

Office.actions.associate("action", action);
