import xapi from 'xapi';

// Global Parameters
const gridDefault = true; // Define Grid View Default option
const silentDtmf = true; // Change to false if you want to hear the DTMF Feedback tones;
const debugMode = false; // Enable debug logging

// ----- EDIT BELOW THIS LINE AT OWN RISK ----- //

const vimtDomain = '@m.webex.com';
let vimtSetup = false;
let vimtActive = false;

const sleep = (timeout) => new Promise((resolve) => {
  setTimeout(resolve, timeout);
});

// Function to send the DTMF tones during the call.
// This receives the DTMF tones from the Event Listener below.
async function sendDTMF(code, message) {
  if (debugMode) console.debug(message);

  try {
    const Level = await xapi.Status.Audio.Volume.get();
    await xapi.Command.Audio.Volume.Set({
      Level: '30',
    });
    if (debugMode) console.debug('Volume set to 30');

    await sleep(200);
    const Feedback = silentDtmf ? 'Silent' : 'Audible';
    xapi.Command.Call.DTMFSend({
      DTMFString: code,
      Feedback,
    });
    if (debugMode) console.debug('DTMF Values Sent');

    await sleep(750);
    await xapi.Command.Audio.Volume.Set({
      Level,
    });
    if (debugMode) console.debug(`Volume Set Back to ${Level}`);
  } catch (error) {
    console.error(error);
  }
}
// Init function - Listen for Calls and check if matches VIMT Uri
function init() {
  // Reset Call Status to false

  xapi.Event.OutgoingCallIndication.on(async () => {
    const call = await xapi.Status.Call.get();
    if (call.length < 1) {
      // No Active Call
      return;
    }
    if (debugMode) console.debug(call[0].CallbackNumber);
    if (call[0].CallbackNumber.match(vimtDomain)) {
      // Matched VIMT Call, update Global Variable
      vimtSetup = true;
      if (debugMode) console.debug(`VIMT Setup: ${vimtSetup}`);
      const vimtRegex = new RegExp(`.*[0-9].*..*${vimtDomain}`);
      if (call[0].CallbackNumber.match(vimtRegex)) {
        vimtActive = true;
        if (debugMode) console.debug(`VIMT Active: ${vimtActive}`);
      }
    }
  });

  async function performActions() {
    // Successfully joined VIMT Call, perform optimizations
    // Pause 1 second just to be sure
    await sleep(1000);
    const messageText = [];
    // Grid Default
    if (gridDefault) {
      try {
        await xapi.Command.Video.Layout.SetLayout({ LayoutName: 'Grid' });
        messageText.push('Grid View Layout Enabled');
      } catch (error) {
        console.error('Unable to set Grid');
        console.debug(error);
      }
    }
  }

  xapi.Event.CallSuccessful.on(async () => {
    const call = await xapi.Status.Call.get();
    if (call.length < 1) {
      // No Active Call
      return;
    }
    // Check if active VIMT Call
    if (vimtActive) {
      await performActions();
      return;
    }
    // Check if VIMT Call setup (aka VTC Conference DTMF Menu)
    if (vimtSetup) {
      if (debugMode) console.debug('Pending Meeting Join');
      // Wait for Participant Added Event which is triggered once Device is Admitted
      xapi.Event.Conference.ParticipantList.ParticipantAdded.on(async () => {
        if (vimtSetup && !vimtActive) {
          vimtActive = true;
          if (debugMode) console.debug(`VIMT Active: ${vimtActive}`);
          performActions();
        }
      });
    }
  });

  xapi.Event.CallDisconnect.on(() => {
    // Call Disconnect detected, remove Panel
    if (debugMode) console.debug('Invoke Remove Panel');
    removePanel();
    // Restore VIMT Global Variables
    vimtSetup = false;
    if (debugMode) console.debug(`VIMT Setup: ${vimtSetup}`);
    vimtActive = false;
    if (debugMode) console.debug(`VIMT Active: ${vimtActive}`);
  });
}

// Initialize Function
init();
