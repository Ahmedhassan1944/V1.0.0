/**
 * =========================================================
 * ENTERPRISE HR MOBILIZATION SYSTEM - BACKEND CONTROLLER
 * =========================================================
 */

/**
 * Standard GET handler. Evaluates and serves the Index.html template.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Enterprise HR Mobilization')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Required if embedding
}

/**
 * Helper function to inject Styles.html and Script.html into Index.html.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * =========================================================
 * API LAYER -> To be called via google.script.run
 * =========================================================
 */

/**
 * Simulates fetching dashboard data from the underlying Google Sheet.
 */
function api_getDashboardData() {
  try {
    const res = api_getAllCandidates();
    if (!res.success) throw new Error(res.error);
    
    let activeCount = 0;
    let missingDocs = 0;
    let pendingValidation = 0;
    let mobilized = 0;
    
    res.data.forEach(cand => {
      const status = cand.CurrentStatus || '';
      
      if (status !== 'Closed') {
        activeCount++;
      }
      
      if (status === 'Documents Requested') {
        missingDocs++;
      } else if (status === 'Pending Validation') {
        pendingValidation++;
      } else if (status === 'Mobilized') {
        mobilized++;
      }
    });
    
    return {
      success: true,
      data: {
        activeCount: activeCount,
        missingDocs: missingDocs,
        pendingValidation: pendingValidation,
        mobilized: mobilized,
        msg: "Live data successfully retrieved from Google Sheets."
      }
    };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}
