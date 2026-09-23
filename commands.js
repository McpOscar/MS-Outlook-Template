/*
 * Commands.js - Event-Based Activation for Meeting Template
 * FIXED: Bruger OnNewAppointmentOrganizer event
 */

// Møde skabelon i HTML format med Neue Hass Grotesk font
const MEETING_TEMPLATE = `
<div style="font-family: Arial, sans-serif; color: #333333;">
<p><strong>Meeting Purpose</strong></p>
<p>Brief description of why we are meeting and what we aim to achieve.</p>
<br>
<p><strong>Agenda</strong></p>
<p>1.&nbsp;<br>2.&nbsp;<br>3.&nbsp;</p>
<br>
<p><strong>Roles</strong></p>
<p>Meeting lead:&nbsp;&nbsp;&nbsp; Notes:&nbsp;</p>
<br>
<p><strong>Decisions and Next Steps</strong></p>
<p>To be completed during the meeting.</p>
<br>
<p><strong>Other Business</strong></p>
<p>Questions or items raised at the meeting.</p>
</div>
`;

// Initialize Office.js
Office.onReady(() => {
  console.log('Office.js ready - Meeting Template Add-in loaded');
});

/**
 * Event handler for OnNewAppointmentOrganizer
 * Kaldes automatisk når en ny mødeaftale oprettes
 *
 * VIGTIGT: event.completed() SKAL kaldes når færdig!
 */
function onNewAppointmentOrganizer(event) {
  console.log('OnNewAppointmentOrganizer event triggered');

  const item = Office.context.mailbox.item;

  // Check om body allerede har indhold (undgå duplicate insert)
  item.body.getAsync(Office.CoercionType.Text, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {

      const bodyText = result.value.trim();

      // Kun indsæt hvis body er tom eller næsten tom
      if (bodyText.length < 10) {

        console.log('Body is empty - inserting template');

        // Indsæt skabelon i møde body
        item.body.setAsync(
          MEETING_TEMPLATE,
          { coercionType: Office.CoercionType.Html },
          function(setResult) {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('Template inserted successfully');
            } else {
              console.error('Error inserting template:', setResult.error);
            }

            // KRITISK: Signaler at event handler er færdig
            event.completed();
          }
        );

      } else {
        console.log('Body already has content - skipping template insert');
        // KRITISK: event.completed() SKAL kaldes selv hvis vi springer insert over
        event.completed();
      }

    } else {
      console.error('Error reading body:', result.error);
      // KRITISK: Kald event.completed() selv ved fejl
      event.completed();
    }
  });
}

// Registrer event handler globalt så manifest kan kalde den
// VIGTIGT: Function navn skal matche præcis det der står i manifest
if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate("onNewAppointmentOrganizer", onNewAppointmentOrganizer);
  console.log('Event handler registered: onNewAppointmentOrganizer');
}
