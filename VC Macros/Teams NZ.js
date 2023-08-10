const VIMT_TENANT = '460393446' // Customer VIMT Tenant
const WEBEX_DOMAIN = '@t.plcm.vc' // VIMT Webex Domain
const SHOW_BUTTON = true; // Display Teams button on Device
const DIALPAD_ID = 'teamsNZdialpad'; // Dial Pad Identifier
const PANEL_ID = 'teamsjoinNZ'; // Action Button Identifier

async function addButton() {
  const config = await xapi.Command.UserInterface.Extensions.List();
  if (config.Extensions && config.Extensions.Panel) {
    const buttonExist = config.Extensions.Panel.find(
      (panel) => panel.PanelId === PANEL_ID
    );
    if (buttonExist) {
      console.debug('Teams Button already added');
      return;
    }
  }

  console.debug(`Adding Teams NZ Button`);
  const xml = `<?xml version="1.0"?>
  <Extensions>
  <Version>1.8</Version>
  <Panel>
    <Order>1</Order>
    <PanelId>${PANEL_ID}</PanelId>
    <Type>Home</Type>
    <Icon>Custom</Icon>
    <Name>Join Teams NZ</Name>
    <ActivityType>Custom</ActivityType>
    <CustomIcon>
    <Content>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</Content>
      <Id>18b5b5656820422f47afa93147e7fd23e4cfe7678efaa9df31187a3c60b21665</Id>
    </CustomIcon>
  </Panel>
  </Extensions>`;

  await xapi.Command.UserInterface.Extensions.Panel.Save(
    {
      PanelId: PANEL_ID,
    },
    xml
  );
}

async function removeButton() {
  const config = await xapi.Command.UserInterface.Extensions.List();
  if (config.Extensions && config.Extensions.Panel) {
    const buttonExist = config.Extensions.Panel.find(
      (panel) => panel.PanelId === PANEL_ID
    );
    if (!buttonExist) {
      console.debug('Teams Button does not exist');
      return;
    }
  }

  console.debug(`Removing Teams Button`);
  await xapi.Command.UserInterface.Extensions.Panel.Close();
  await xapi.Command.UserInterface.Extensions.Panel.Remove({
    PanelId: PANEL_ID,
  });
}

function showDialPad(){
  xapi.command('UserInterface Message TextInput Display', {
      InputType: 'Numeric'
    , Placeholder: 'Conference ID'
    , Title: 'Microsoft Teams Meeting'
    , Text: 'Enter the Video Conference ID'
    , SubmitText: 'Dial' 
    , FeedbackId: DIALPAD_ID
  }).catch((error) => { console.error(error); });
}

// Init function
function init() {
  // Process Button Status
  if (SHOW_BUTTON) {
    addButton();
  } else {
    removeButton();
  }

  // Monitor for Panel Click Events
  xapi.event.on('UserInterface Extensions Panel Clicked', async (event) => {
    if (event.PanelId == PANEL_ID) {
      showDialPad();
    }
  });

  // Monitor for TextInput Responses
  xapi.event.on('UserInterface Message TextInput Response', (event) => {
    switch (event.FeedbackId) {
      case DIALPAD_ID:
        xapi.command('dial', {
          Number: `${event.Text}.${VIMT_TENANT}${WEBEX_DOMAIN}`
        }).catch((error) => { console.error(error); });
        break;
    }
  });
}

// Initialize Function
init();import xapi from 'xapi';

