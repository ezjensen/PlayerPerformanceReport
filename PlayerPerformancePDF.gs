/**
 * This script automates the process of generating PDFs for each athlete,
 * logging both successful and failed attempts. It now uses timed triggers
 * to process players in chunks to prevent script timeouts and saves files
 * to a dedicated "PerformancePlayerReports" folder in the user's Drive.
 *
 * It also includes a custom menu to manually start or continue the process.
 */

/**
 * Creates a custom menu in the spreadsheet UI when the sheet is opened.
 * This is necessary for manual functions to be accessible to all users.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('PDF Reports')
      .addItem('Start PDF Generation (All Players)', 'generatePlayerReports')
      .addSeparator()
      .addItem('Test-5 (Random 5 Players)', 'generateTestReports')
      .addToUi();
}

/**
 * This function should be run once to initiate the process. It is called
 * from the custom menu.
 */
function generatePlayerReports() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const profilesSheet = spreadsheet.getSheetByName("Profiles");
  
  if (!profilesSheet) {
    SpreadsheetApp.getUi().alert("Error: The 'Profiles' sheet was not found.");
    return;
  }
  
  // Get the list of all active athletes and their teams.
  const profileData = profilesSheet.getRange("A2:B").getValues()
                                   .filter(row => row[0] !== "" && row[1] !== "");

  // Store the list of players and their teams in a single property
  // This helps avoid re-reading the sheet on every run
  const playersProperty = PropertiesService.getScriptProperties();
  playersProperty.setProperty('playersList', JSON.stringify(profileData));
  playersProperty.setProperty('lastProcessedIndex', '0');
  playersProperty.setProperty('totalPlayers', profileData.length.toString());
  
  // Find or create the destination folder
  const destinationFolderName = "PerformancePlayerReports";
  const folders = DriveApp.getFoldersByName(destinationFolderName);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(destinationFolderName);
  }
  playersProperty.setProperty('destinationFolderId', folder.getId());

  // Delete any existing triggers to start fresh
  deleteTriggers();
  
  // Set up a timed trigger to process the first chunk in a few seconds
  ScriptApp.newTrigger('processPlayersInChunks')
           .timeBased()
           .after(10000) // 10 seconds
           .create();
           
  SpreadsheetApp.getUi().alert("PDF generation has started and will run in the background. Please close the script editor and sheet. The process will complete automatically. Check the 'ScriptLog' sheet for progress, or use the 'PDF Reports' menu to manually process the next chunk.");
}

/**
 * This new function initiates a test run for a random subset of 5 players.
 * It follows the same logic as generatePlayerReports but with a smaller list.
 * This can be run to test the process without generating reports for all players.
 */
function generateTestReports() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const profilesSheet = spreadsheet.getSheetByName("Profiles");

  if (!profilesSheet) {
    SpreadsheetApp.getUi().alert("Error: The 'Profiles' sheet was not found.");
    return;
  }

  // Get the list of all active athletes.
  const allProfiles = profilesSheet.getRange("A2:B").getValues()
                                   .filter(row => row[0] !== "" && row[1] !== "");
  
  // Select a random sample of 5 players.
  const numPlayersToTest = 5;
  const testPlayers = [];
  while (testPlayers.length < numPlayersToTest && allProfiles.length > 0) {
    const randomIndex = Math.floor(Math.random() * allProfiles.length);
    testPlayers.push(allProfiles.splice(randomIndex, 1)[0]);
  }

  // Store the test list of players in a single property
  const playersProperty = PropertiesService.getScriptProperties();
  playersProperty.setProperty('playersList', JSON.stringify(testPlayers));
  playersProperty.setProperty('lastProcessedIndex', '0');
  playersProperty.setProperty('totalPlayers', testPlayers.length.toString());

  // Find or create the destination folder
  const destinationFolderName = "PerformancePlayerReports";
  const folders = DriveApp.getFoldersByName(destinationFolderName);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(destinationFolderName);
  }
  playersProperty.setProperty('destinationFolderId', folder.getId());

  // Delete any existing triggers to start fresh
  deleteTriggers();
  
  // Set up a timed trigger to process the first chunk in a few seconds
  ScriptApp.newTrigger('processPlayersInChunks')
           .timeBased()
           .after(10000) // 10 seconds
           .create();
           
  SpreadsheetApp.getUi().alert("Test PDF generation has started and will run in the background for 5 random athletes.");
}

/**
 * This function processes a chunk of players and is designed to be called by a timed trigger
 * or a manual menu item. It reads the last processed index, handles a limited number
 * of players, and then sets up the next trigger if needed.
 */
function processPlayersInChunks() {
  const SCRIPT_TIMEOUT = 300000; // 5 minutes in milliseconds
  const CHUNK_SIZE = 15; // Number of players to process in each chunk
  
  // This is a new line to ensure we only have one trigger at a time.
  deleteTriggers();

  const properties = PropertiesService.getScriptProperties();
  const playersList = JSON.parse(properties.getProperty('playersList'));
  let lastProcessedIndex = parseInt(properties.getProperty('lastProcessedIndex'));
  const totalPlayers = parseInt(properties.getProperty('totalPlayers'));
  let folderId = properties.getProperty('destinationFolderId');

  const startTime = new Date().getTime();
  
  // Get the active spreadsheet and the relevant sheets
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = spreadsheet.getSheetByName("AthleteDashboard");
  const overallScoreFormula = "=IFERROR(INDEX(ProfileData, MATCH($A$5,Profiles!$A:$A,0),MATCH($C8,ProfileHeaders,0)), \"\")";
  const dashboardSheetId = dashboardSheet.getSheetId();

  // New logic to handle folder permissions for collaborators
  let folder;
  const destinationFolderName = "PerformancePlayerReports";
  try {
    // Attempt to get the folder by the stored ID
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log(`Error accessing folder by ID: ${e.message}. Attempting to find or create folder by name.`);
    // If that fails, find the folder by name in the current user's Drive
    const folders = DriveApp.getFoldersByName(destinationFolderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      // If the folder doesn't exist, create it
      folder = DriveApp.createFolder(destinationFolderName);
    }
    // Update the property for the current user's runs
    properties.setProperty('destinationFolderId', folder.getId());
    folderId = folder.getId();
  }

  // Get the log sheet, or create it if it doesn't exist.
  let logSheet = spreadsheet.getSheetByName("ScriptLog");
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet("ScriptLog");
    // Add a new column for the PDF link
    logSheet.appendRow(["Timestamp", "Player_Name", "Team", "Status", "Message", "PDF Link"]);
  }

  // Loop through the chunk of players
  for (let i = lastProcessedIndex; i < lastProcessedIndex + CHUNK_SIZE && i < totalPlayers; i++) {
    // Check if the script is approaching the timeout limit
    if (new Date().getTime() - startTime > SCRIPT_TIMEOUT) {
      Logger.log("Approaching timeout, stopping and rescheduling.");
      // Stop and set up the next trigger to continue from this point
      properties.setProperty('lastProcessedIndex', i.toString());
      ScriptApp.newTrigger('processPlayersInChunks')
               .timeBased()
               .after(5 * 60 * 1000) // 5 minutes
               .create();
      return;
    }

    const [playerName, playerTeam] = playersList[i];
    const timestamp = new Date().toLocaleString();
    
    try {
      // Set the player's name and team to ensure the file name logic is correct
      dashboardSheet.getRange("A5").setValue(playerName);
      dashboardSheet.getRange("D3").setValue(playerTeam);
      
      // Increased sleep time to prevent "Too Many Requests" errors.
      Utilities.sleep(2000);
      
      dashboardSheet.getRange("D8").setFormula(overallScoreFormula);
      const overallScore = dashboardSheet.getRange("D8").getDisplayValue();
      const sanitizedScore = overallScore.toString().replace(/[^a-z0-9]/gi, '_');
      const pdfName = `${playerName}_${sanitizedScore}.pdf`;
      
      // Check if the file already exists in the folder
      const existingFiles = folder.getFilesByName(pdfName);
      if (existingFiles.hasNext()) {
        logSheet.appendRow([timestamp, playerName, playerTeam, "Skipped", "PDF already exists.", ""]);
        Logger.log(`Skipped generating report for ${playerName}: PDF already exists.`);
        continue;
      }
      
      // Increased sleep time to prevent "Too Many Requests" errors.
      Utilities.sleep(2000);
  
      // Create the PDF file using the URL Fetch method.
      const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?` +
                  `format=pdf&size=letter&portrait=false&fitw=true&fith=true&` +
                  `top_margin=0.75&bottom_margin=0.75&left_margin=0.7&right_margin=0.7&` +
                  `gid=${dashboardSheetId}&` +
                  `r1=0&r2=35&c1=0&c2=12`; // This corresponds to A1:M35
      
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(url, {
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });
  
      const blob = response.getBlob();
      const file = folder.createFile(blob); // Capture the created file object
      file.setName(pdfName);
      const pdfLink = file.getUrl();
      
      // Log the PDF link along with the other information
      logSheet.appendRow([timestamp, playerName, playerTeam, "Success", "PDF created successfully.", pdfLink]);
      Logger.log(`Generated report for ${playerName} from team ${playerTeam} with a score of ${overallScore}`);
      
    } catch (e) {
      logSheet.appendRow([timestamp, playerName, playerTeam, "Error", e.message, ""]);
      Logger.log(`Error generating report for ${playerName}: ${e.message}`);
    }
  }

  // Update the last processed index for the next run
  lastProcessedIndex += CHUNK_SIZE;
  properties.setProperty('lastProcessedIndex', lastProcessedIndex.toString());

  // Check if all players have been processed
  if (lastProcessedIndex >= totalPlayers) {
    // Log the completion message instead of using an alert
    logSheet.appendRow([new Date().toLocaleString(), "", "", "Complete", "All player reports have been processed.", ""]);
    Logger.log("All player reports have been processed. The process is complete.");
    deleteTriggers();
  } else {
    // Set up the next trigger
    ScriptApp.newTrigger('processPlayersInChunks')
             .timeBased()
             .after(1 * 60 * 1000) // 1 minute
             .create();
  }
}

/**
 * Deletes all triggers for the current project.
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
