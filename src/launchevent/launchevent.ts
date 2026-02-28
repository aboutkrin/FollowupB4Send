/*
 * OnMessageSend event handler.
 * Runs in the JS-only runtime (no DOM). Must call event.completed() in every path.
 */

function onMessageSendHandler(event: Office.AddinCommands.Event): void {
  // Block send and prompt the user with a Smart Alert dialog.
  // "Set Reminder" opens the task pane; "Send Anyway" lets the mail through.
  event.completed({
    allowEvent: false,
    cancelLabel: "Set Reminder",
    commandId: "msgComposeOpenPaneButton",
    sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
    errorMessageMarkdown:
      "**Would you like to set a follow-up reminder?**\n\n" +
      'Click **Set Reminder** to pick a date, or **Send Anyway** to send without a reminder.',
  } as Office.AddinCommands.EventCompletedOptions);
}

// Register the handler so the Office runtime can find it by name.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
