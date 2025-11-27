/***********************
 *  CONFIG & GLOBALS   *
 ***********************/
const DEFAULT_ACCESS_TOKEN = '9SEFkXuQXEiLcf3VxTW4Pw%3D%3D';

// --- DRIVE CACHE CONFIGURATION ---
const CACHE_FOLDER_NAME = 'EDT_Cache'; // The name of the folder that will be created on Drive
const CACHE_FILE_NAME = 'sanitized_edt.xml'; // The name of the cache file
const CACHE_EXPIRATION_SECONDS_DRIVE = 1800; // 1800 seconds = 30 minutes

/**
 * Serves the main HTML page of the web app.
 * @param {Object} e The event parameter.
 * @returns {HtmlService.HtmlOutput} The HTML page to serve.
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.userEmail = getUserEmail();
  
  return template
    .evaluate()
    .setTitle('EDT Management')
    .setFaviconUrl('https://iam.pubblica.istruzione.it/iam-ssum/master/assets/img/favicon.ico')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Gets the email of the active user.
 * @returns {string} The user's email.
 */
function getUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    Logger.log(`User Email: ${email}`);
    return email;
  } catch (e) {
    Logger.log('Could not get user email (likely running as developer).');
    return '';
  }
}

/********************
 *  FETCH & CACHE   *
 ********************/

/**
 * Fetches the raw XML data from the source.
 * @param {string} accessToken The access token for the API.
 * @returns {{raw: string}} An object containing the raw XML text.
 */
function fetchEDTXml(accessToken = DEFAULT_ACCESS_TOKEN) {
  const url = 'https://potf010003edt.index-education.net/edt/api/DataSync/V1/ExportStd?AccessToken=' + accessToken;
  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  
  if (res.getResponseCode() !== 200) {
    throw new Error('HTTP ' + res.getResponseCode() + ': ' + res.getContentText());
  }
  
  const raw = res.getContentText();
  if (!raw || !raw.trim()) {
    throw new Error('Received empty XML content');
  }
  return { raw };
}

/*********************************
 *  CLIENT-ACCESSIBLE FUNCTIONS  *
 *********************************/

/**
 * Removes sensitive attributes from an XML element.
 * @param {XmlService.Element} element The element to clean.
 * @param {string[]} attributesToRemove An array of attribute names to remove.
 */
function removeSensitiveAttributes(element, attributesToRemove) {
  if (!element) return;
  
  attributesToRemove.forEach(attrName => {
    if (attrName.endsWith('X')) {
      const baseAttrName = attrName.slice(0, -1);
      const attributes = element.getAttributes();
      attributes.slice().forEach(attr => {
        if (attr.getName().startsWith(baseAttrName)) {
          element.removeAttribute(attr.getName());
        }
      });
    } else {
      if (element.getAttribute(attrName)) {
        element.removeAttribute(attrName);
      }
    }
  });
}

/**
 * Gets or creates the cache folder on Google Drive.
 * @return {DriveApp.Folder} The cache Folder object.
 */
function getOrCreateCacheFolder() {
  const folders = DriveApp.getFoldersByName(CACHE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CACHE_FOLDER_NAME);
}

/**
 * Provides sanitized XML data to the client, using a Drive file as a cache.
 * This function is exposed to the client-side JavaScript.
 * @return {string} The sanitized XML as a string.
 */
function getEdtXml() {
  try {
    const cacheFolder = getOrCreateCacheFolder();
    const files = cacheFolder.getFilesByName(CACHE_FILE_NAME);
    
    if (files.hasNext()) {
      const file = files.next();
      const lastUpdated = file.getLastUpdated();
      const now = new Date();
      const ageInSeconds = (now.getTime() - lastUpdated.getTime()) / 1000;

      if (ageInSeconds < CACHE_EXPIRATION_SECONDS_DRIVE) {
        Logger.log('Returning XML from Drive cache.');
        return file.getBlob().getDataAsString();
      }
    }

    Logger.log('Drive cache is stale or missing. Fetching and cleaning new data.');

    const { raw } = fetchEDTXml();

    // Remove large, sensitive nodes with regex for performance before parsing
    let sanitizedRaw = raw.replace(/<Eleves>[\s\S]*?<\/Eleves>/g, '<Eleves/>');
    sanitizedRaw = sanitizedRaw.replace(/<Personnels>[\s\S]*?<\/Personnels>/g, '<Personnels/>');
    
    const doc = XmlService.parse(sanitizedRaw);
    const root = doc.getRootElement();
    const ns = root.getNamespace();

    const attrsToRemove = ['DateNaissance', 'IDPN', 'AdresseX', 'CodePostal', 'TelPortable', 'Pays', 'TelFixe'];

    const professeursNode = root.getChild('Professeurs', ns);
    if (professeursNode) {
      professeursNode.getChildren('Professeur', ns).forEach(prof => removeSensitiveAttributes(prof, attrsToRemove));
    }

    const responsablesNode = root.getChild('Responsables', ns);
    if (responsablesNode) {
      responsablesNode.getChildren('Responsable', ns).forEach(resp => removeSensitiveAttributes(resp, attrsToRemove));
    }
    
    const outputter = XmlService.getPrettyFormat();
    const finalXml = outputter.format(doc);
    
    // Update or create the cache file on Drive
    const filesForUpdate = cacheFolder.getFilesByName(CACHE_FILE_NAME);
    if (filesForUpdate.hasNext()) {
      filesForUpdate.next().setContent(finalXml);
      Logger.log('Drive cache updated.');
    } else {
      cacheFolder.createFile(CACHE_FILE_NAME, finalXml);
      Logger.log('Drive cache file created.');
    }
    
    return finalXml;
    
  } catch (e) {
    Logger.log('Error in getEdtXml: ' + e.toString() + '\nStack: ' + e.stack);
    // Return a user-friendly error message to the client
    return 'Error: Could not retrieve or clean the XML file. ' + e.message;
  }
}
