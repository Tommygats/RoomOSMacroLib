// eslint-disable-next-line import/no-unresolved
import xapi from 'xapi';

// Global Parameters
const hideNonVideo = true; // Hide Non-Video Participants by default
const addNonVideoButton = true; // Adds option during meeting to toggle non-video participants
const silentDtmf = true; // Change to false if you want to hear the DTMF Feedback tones;
const debugMode = false; // Enable debug logging

// ----- EDIT BELOW THIS LINE AT OWN RISK ----- //

const vimtDomain = '@m.webex.com';
const panelId = 'HideNonVid';
const messageTitle = 'Hide Non-Video Participants';
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

// Event Listener for Panel Clicks
xapi.event.on('UserInterface Extensions Panel Clicked', (event) => {
  if (event.PanelId === panelId) {
    try {
      sendDTMF('#5', 'Hide Non-Video was Pressed');
      xapi.Command.UserInterface.Message.Alert.Display({
        Title: messageTitle,
        Text: 'Hide Non-Video Participants Toggled',
        Duration: messageTimeout,
      });
    } catch (error) {
      console.error('Unable to toggle non-video participants');
      console.debug(error);
    }
  }
});

async function addPanel() {
  const config = await xapi.Command.UserInterface.Extensions.List();
  if (config.Extensions.Panel) {
    const panelExist = config.Extensions.Panel.find(
      (panel) => panel.PanelId === panelId,
    );
    if (panelExist) {
      if (debugMode) console.debug('Panel already added');
      return;
    }
  }

  if (debugMode) console.debug('Adding Panel');
  const xml = `<?xml version="1.0"?>
  <Extensions>
  <Version>1.10</Version>
  <Panel>
    <Order>1</Order>
    <PanelId>${panelId}</PanelId>
    <Location>CallControls</Location>
    <Icon>Camera</Icon>
    <Color>#262626</Color>
    <Name>Toggle Non-Video</Name>
    <ActivityType>Custom</ActivityType>
  </Panel>
  </Extensions>`;

  await xapi.Command.UserInterface.Extensions.Panel.Save(
    {
      PanelId: panelId,
    },
    xml,
  );
}

async function removePanel() {
  const config = await xapi.Command.UserInterface.Extensions.List();
  if (config.Extensions.Panel) {
    const panelExist = config.Extensions.Panel.find(
      (panel) => panel.PanelId === panelId,
    );
    if (!panelExist) {
      if (debugMode) console.debug('Panel does not exist');
      return;
    }
  }

  if (debugMode) console.debug('Removing Panel');
  await xapi.Command.UserInterface.Extensions.Panel.Close();
  await xapi.Command.UserInterface.Extensions.Panel.Remove({
    PanelId: panelId,
  });
}

// Init function - Listen for Calls and check if matches VIMT Uri
function init() {
  // Remove lingering Button during init.
  removePanel();
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
    // Hide Non-Video Participants
    if (hideNonVideo) {
      try {
        await sendDTMF('#5', 'Hide Non-Video was Pressed');
        messageText.push('Non-Video Participants Hidden');
      } catch (error) {
        console.error('Unable to Hide Non-Video');
        console.debug(error);
      }
    }
    // Add Non-Video Participants Toggle
    if (addNonVideoButton) {
      try {
        await addPanel();
        messageText.push('Non-Video Toggle added to Controls');
      } catch (error) {
        console.error('Unable to add Panel');
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
