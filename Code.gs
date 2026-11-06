// ==============================================================================
// YASH MARKETING - ENTERPRISE CRM ENGINE (APPS SCRIPT LOGIC)
// FULLY EXPANDED, UNCOMPRESSED, INCLUSIVE OF ALL FEATURES & GAMIFICATION
// ==============================================================================

/**
 * doGet is the entry point for the web app.
 * It loads the Index.html file and passes URL parameters if any.
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  // Check if Employee ID and Role are passed via URL (For SSO / Portal Bypass)
  if (e && e.parameter && e.parameter.empId) {
    template.empId = e.parameter.empId;
    template.role = e.parameter.role;
  } else {
    template.empId = ""; 
    template.role = "";
  }
  
  // Evaluate the HTML template and set meta tags for mobile responsiveness
  return template.evaluate()
    .setTitle('Yash Marketing - CRM Enterprise Tower')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * VERIFY LOGIN
 * Checks credentials against the "Staff Master" sheet.
 */
function verifyLogin(empId, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staffSheet = ss.getSheetByName("Staff Master");
  
  if (!staffSheet) {
    return { success: false, msg: "Staff Master sheet missing" };
  }
  
  const data = staffSheet.getDataRange().getValues();
  
  // Loop through the data to find a matching Employee ID and Password
  for (let i = 1; i < data.length; i++) {
    
    let currentEmpId = data[i][0];
    let currentPassword = data[i][4];
    let currentName = data[i][1];
    let currentRole = data[i][5];
    let currentHourlySalary = data[i][7] || 0;
    
    // Allow standard password OR SSO_BYPASS from master portal
    if (currentEmpId == empId && (currentPassword == password || password === "SSO_BYPASS")) {
      return { 
        success: true, 
        name: currentName, 
        role: currentRole, 
        hourlySalary: currentHourlySalary 
      };
    }
    
  }
  
  // If no match is found
  return { success: false, msg: "Invalid Access Credentials." };
}

// -------------------------------------------------------------
// NEW FEATURE: DYNAMIC INCENTIVE ENGINE
// Reads from "Incentive Master" to drive the agent gamification UI
// -------------------------------------------------------------
function getCRMIncentiveRules() {
  var rules = { callRate: 0, orderRate: 0 };
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Incentive Master");
    
    if(sheet) {
      const data = sheet.getDataRange().getValues();
      
      for(let i = 0; i < data.length; i++) {
        let key = String(data[i][0]).trim().toLowerCase();
        let val = data[i][1];
        
        if(key.includes("crm per call") || key.includes("number of calls")) {
          rules.callRate = parseFloat(val) || 0;
        }
        
        if(key.includes("crm per order") || key.includes("orders received by crm")) {
          rules.orderRate = parseFloat(val) || 0;
        }
      }
    }
  } catch(e) {
    // Graceful fallback if sheet isn't created yet
    Logger.log("Incentive Master Error: " + e.toString());
  }
  
  return rules;
}


/**
 * GET BEAT PLAN
 * Fetches the daily calling queue, follow-ups, and outstanding dues for a specific agent.
 */
function getBeatPlan(staffName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Dealer Database");
  const scriptSheet = ss.getSheetByName("Scripts Master");
  const outSheet = ss.getSheetByName("Outstanding"); 
  
  if (!dbSheet) {
    return { dealers: [], error: "Dealer Database sheet not found" };
  }
  
  // Load Dynamic Rules for the Gamification Engine
  const rules = getCRMIncentiveRules(); 

  // ---------------------------------------------------------
  // STEP 1: Fetch Outstanding Data and map it by Dealer Name
  // ---------------------------------------------------------
  let outstandingMap = {};
  
  if (outSheet) {
    const outData = outSheet.getDataRange().getValues();
    
    for (let i = 1; i < outData.length; i++) {
      
      let nameWithBill = outData[i][0] ? outData[i][0].toString().trim().toLowerCase() : null;
      let amtWithBill = outData[i][1] || 0;
      
      let nameWithoutBill = outData[i][3] ? outData[i][3].toString().trim().toLowerCase() : null;
      let amtWithoutBill = outData[i][4] || 0;
      
      // Store With Bill Outstanding
      if (nameWithBill) {
        let prevWithout = outstandingMap[nameWithBill] ? outstandingMap[nameWithBill].withoutBill : 0;
        outstandingMap[nameWithBill] = { 
          withBill: amtWithBill, 
          withoutBill: prevWithout 
        };
      }
      
      // Store Without Bill Outstanding
      if (nameWithoutBill) {
        let prevWith = outstandingMap[nameWithoutBill] ? outstandingMap[nameWithoutBill].withBill : 0;
        outstandingMap[nameWithoutBill] = { 
          withBill: prevWith, 
          withoutBill: amtWithoutBill 
        };
      }
      
    }
  }

  // ---------------------------------------------------------
  // STEP 2: Date and Day Calculations
  // ---------------------------------------------------------
  const days = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
  const todayDate = new Date();
  const todayDay = days[todayDate.getDay()];
  const todayStr = todayDate.toLocaleDateString('en-IN');
  
  const data = dbSheet.getDataRange().getValues();
  const cleanStaffName = staffName.toString().trim().toLowerCase();
  
  let resultQueue = [];
  let resultFollowUps = [];
  let allInteractions = {}; 
  
  let funnel = { assigned: 0, talked: 0, orders: 0 };
  let baseCommission = 0; 
  let biggestDeal = { agent: "", amount: 0 };

  // ---------------------------------------------------------
  // STEP 3: First Loop - Build History, Deal of the day & Funnel
  // ---------------------------------------------------------
  for (let i = 1; i < data.length; i++) {
    
    let row = data[i];
    let phone = row[0];
    let crm = row[3] ? row[3].toString().trim().toLowerCase() : "";
    let callingDay = row[4] ? row[4].toString().trim().toLowerCase() : "";
    let status = row[7] || "";
    let callDate = row[9] ? new Date(row[9]).toLocaleDateString('en-IN') : "";
    let amt = parseFloat(row[10]) || 0;
    let roi = parseFloat(row[12]) || 0;
    
    // Initialize Interaction History Array for this phone number
    if (!allInteractions[phone]) {
      allInteractions[phone] = { 
        last3: [], 
        lastOrderDate: null 
      };
    }
    
    // Populate History if status exists
    if (status) {
      allInteractions[phone].last3.unshift({ 
        status: status, 
        date: callDate, 
        amt: amt 
      });
      
      if (status === "Order Received") {
        allInteractions[phone].lastOrderDate = new Date(row[9]);
      }
    }
    
    // Global Deal of the day logic (Highest order amount today)
    if (callDate === todayStr && status === "Order Received" && amt > biggestDeal.amount) {
      biggestDeal = { 
        agent: row[1], // Storing Dealer name to display
        amount: amt 
      };
    }
    
    // Specific Agent Stats for Funnel and Commission
    if (crm === cleanStaffName) {
      
      // Base Hourly/ROI Commission
      if (callDate === todayStr && roi > 0) {
        baseCommission += roi;
      }
      
      // Assigned Calls count
      if (callingDay === todayDay) {
        funnel.assigned++;
      }
      
      // Talked and Orders count
      if (callDate === todayStr) {
        if (status) {
          funnel.talked++;
        }
        if (status === "Order Received") {
          funnel.orders++;
        }
      }
      
    }
  }

  // Inject Dynamic Incentives into the final Live Commission calculation
  let liveCommission = baseCommission + (funnel.talked * rules.callRate) + (funnel.orders * rules.orderRate);

  // ---------------------------------------------------------
  // STEP 4: Second Loop - Build the actual UI Cards Data
  // ---------------------------------------------------------
  for (let i = 1; i < data.length; i++) {
    
    let row = data[i];
    let phone = row[0];
    let dNameRaw = row[1] || "";
    let areaRaw = row[2] || "";
    let crm = row[3] ? row[3].toString().trim().toLowerCase() : "";
    let callingDay = row[4] ? row[4].toString().trim().toLowerCase() : "";
    let campaignRaw = row[5] || "";
    let vipFlagRaw = row[6] || "";
    let nextDateStr = row[11] ? new Date(row[11]).toLocaleDateString('en-IN') : "";
    
    // Calculate Health Score (Based on last order date)
    let health = "🔴 Dormant";
    let lastOrder = allInteractions[phone] ? allInteractions[phone].lastOrderDate : null;
    
    if (lastOrder) {
      let diffDays = Math.floor((todayDate - lastOrder) / (1000 * 60 * 60 * 24));
      
      if (diffDays <= 7) {
        health = "🟢 Hot";
      } else if (diffDays <= 30) {
        health = "🟡 Warm";
      }
    }

    // Retrieve Outstanding Data using normalized dealer name
    let rawNameLow = dNameRaw.toString().trim().toLowerCase();
    let outWB = outstandingMap[rawNameLow] ? outstandingMap[rawNameLow].withBill : 0;
    let outWOB = outstandingMap[rawNameLow] ? outstandingMap[rawNameLow].withoutBill : 0;
    
    // Get maximum 3 history records
    let histList = allInteractions[phone] ? allInteractions[phone].last3.slice(0, 3) : [];
    
    // Determine status: Only show status if it was updated TODAY. Otherwise blank.
    let displayStatus = "";
    if (row[9] && new Date(row[9]).toLocaleDateString('en-IN') === todayStr) {
      displayStatus = row[7];
    }

    // Create the Dealer Object
    let dealerObj = {
      rowNumber: i + 1, 
      phone: phone, 
      dealerName: dNameRaw, 
      area: areaRaw,
      campaign: campaignRaw, 
      vipFlag: vipFlagRaw, 
      status: displayStatus, 
      outWithBill: outWB,
      outWithoutBill: outWOB,
      health: health,
      history: histList
    };

    // Add to Standard Live Queue if the calling day is today
    if (crm === cleanStaffName && callingDay === todayDay) { 
      resultQueue.push(dealerObj); 
    }
    
    // Add to Follow-up Queue if the next date is today
    if (crm === cleanStaffName && nextDateStr === todayStr) { 
      resultFollowUps.push(dealerObj); 
    }
    
  }
  
  // ---------------------------------------------------------
  // STEP 5: Fetch Live Scripts
  // ---------------------------------------------------------
  let activeScripts = [];
  
  if (scriptSheet) {
    const sData = scriptSheet.getDataRange().getValues();
    
    for (let j = 1; j < sData.length; j++) {
      let campaignName = sData[j][0];
      let scriptGoal = sData[j][1] || "";
      let scriptText = sData[j][2] || "";
      let scriptColor = sData[j][3] || "primary";
      
      if (campaignName) {
        activeScripts.push({ 
          campaign: campaignName, 
          goal: scriptGoal, 
          text: scriptText, 
          color: scriptColor 
        });
      }
    }
  }
  
  return { 
    dealers: resultQueue, 
    followups: resultFollowUps, 
    scripts: activeScripts, 
    funnel: funnel, 
    commission: liveCommission, 
    biggestDeal: biggestDeal, 
    sprintTarget: 1000000 
  }; 
}


/**
 * UPDATE CALL DATA
 * Saves the execution locked by the agent into the database.
 * Pushes historical data to "Activity Logs" and orders to "Order Book".
 */
function updateCallData(rowNumber, status, audioBase64, fileName, orderAmt, orderItems, payMode, followUpDate, staffHourly, staffName, dealerName, phone, lostReason, durationSecs) {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Dealer Database");
  
  // Handle Audio File Upload to Google Drive
  let audioUrl = "";
  if (audioBase64) {
    try {
      const splitBase = audioBase64.split(',');
      const type = splitBase[0].split(';')[0].replace('data:', '');
      const blob = Utilities.newBlob(Utilities.base64Decode(splitBase[1]), type, fileName || 'recording.mp3');
      const file = DriveApp.createFile(blob); 
      
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      audioUrl = "https://drive.google.com/uc?export=view&id=" + file.getId(); 
      
    } catch (e) { 
      Logger.log("Audio Error: " + e.toString()); 
    }
  }
  
  const timestamp = new Date();
  let roi = 0;
  
  // ---------------------------------------------------------
  // 1. UPDATE MASTER STATE (Dealer Database)
  // ---------------------------------------------------------
  dbSheet.getRange(rowNumber, 8).setValue(status); // Column H: Status
  dbSheet.getRange(rowNumber, 10).setValue(timestamp); // Column J: Last Updated
  
  if (audioUrl !== "") dbSheet.getRange(rowNumber, 9).setValue(audioUrl); 
  if (orderAmt) dbSheet.getRange(rowNumber, 11).setValue(orderAmt); 
  if (followUpDate) dbSheet.getRange(rowNumber, 12).setValue(followUpDate); 
  if (lostReason) dbSheet.getRange(rowNumber, 14).setValue(lostReason); 
  if (durationSecs !== undefined && durationSecs !== null) dbSheet.getRange(rowNumber, 17).setValue(durationSecs); 
  
  if (orderAmt && staffHourly) {
    // 15% Gross Margin ROI Calculation mapped from framework ledger
    roi = (orderAmt * 0.15) - staffHourly; 
    dbSheet.getRange(rowNumber, 13).setValue(roi); 
  }

  // ---------------------------------------------------------
  // 2. LOG HISTORICAL ACTIVITY (Activity Logs)
  // ---------------------------------------------------------
  let logSheet = ss.getSheetByName("Activity Logs");
  if (!logSheet) {
    logSheet = ss.insertSheet("Activity Logs");
    logSheet.appendRow(["Timestamp", "Agent Name", "Dealer Name", "Phone", "Status", "Duration (Secs)", "Order Amount", "Items Ordered", "Pay Mode", "Next Follow-up Date", "Lost Reason", "Audio Link", "Commission Earned"]);
  }
  
  logSheet.appendRow([
    timestamp, 
    staffName, 
    dealerName, 
    phone, 
    status, 
    durationSecs || 0, 
    orderAmt || 0, 
    orderItems || "", 
    payMode || "", 
    followUpDate || "", 
    lostReason || "", 
    audioUrl, 
    roi
  ]);

  // ---------------------------------------------------------
  // 3. PUSH TO ORDER BOOK (If applicable)
  // Note: Restored original duplicated blocks precisely as you provided
  // ---------------------------------------------------------
  if (status === "Order Received" && orderAmt) {
    
    let orderSheet = ss.getSheetByName("Order Book");
    
    if (!orderSheet) {
      orderSheet = ss.insertSheet("Order Book");
      orderSheet.appendRow(["Date & Time", "OrderID", "ShopName", "DealerPhone", "OrderItems", "TotalAmount", "Payment Mode", "Additional Discount"]);
    }
    
    let autoOrderId = "YM-" + Math.floor(10000 + Math.random() * 90000);
    orderSheet.appendRow([timestamp, autoOrderId, dealerName, phone, orderItems, orderAmt, payMode, ""]);
  }
  
  // ---------------------------------------------------------
  // Generate Order inside "Order Book"
  // (Preserving this exact redundant block from your original source code)
  // ---------------------------------------------------------
  if (status === "Order Received" && orderAmt) {
    
    let orderSheet = ss.getSheetByName("Order Book");
    
    // Create sheet if it does not exist
    if (!orderSheet) {
      orderSheet = ss.insertSheet("Order Book");
      orderSheet.appendRow([
        "Date & Time", 
        "OrderID", 
        "ShopName", 
        "DealerPhone", 
        "OrderItems", 
        "TotalAmount", 
        "Payment Mode", 
        "Additional Discount"
      ]);
    }
    
    // Generate an automatic Order ID (e.g., YM-84920)
    let autoOrderId = "YM-" + Math.floor(10000 + Math.random() * 90000);
    
    // Append the exact array to match Order Book columns
    orderSheet.appendRow([
      timestamp, 
      autoOrderId, 
      dealerName, 
      phone, 
      orderItems, 
      orderAmt, 
      payMode, 
      "" // Additional discount left blank for accountant
    ]);
  }
  
  return {success: true, msg: "Execution Locked Successfully."};
}


/**
 * GET CEO DATA
 * Fetches macro analytics, agent performance, heatmaps, and live feeds based on date filters.
 */
function getCeoData(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("Dealer Database");
  
  if (!dbSheet) {
    return { error: "Database missing" };
  }
  
  const data = dbSheet.getDataRange().getValues();
  
  let totalCalls = 0;
  let totalOrders = 0;
  let revenue = 0;
  let recentActivity = [];
  
  let agentStats = {}; 
  let areaStats = {}; 
  let idleAlerts = [];
  
  // Setup Date Boundaries for Filtering
  let start = startDateStr ? new Date(startDateStr) : new Date();
  let end = endDateStr ? new Date(endDateStr) : new Date();
  
  // Normalize time to cover entire days
  start.setHours(0, 0, 0, 0); 
  end.setHours(23, 59, 59, 999);
  
  let latestAgentTime = {}; 

  // Loop through Dealer Database
  for (let i = 1; i < data.length; i++) {
    
    let row = data[i];
    let lastUpdated = row[9] ? new Date(row[9]) : null;
    let crmName = row[3] || "Unknown";
    
    // Track Last Interaction Time for Idle Alerts
    if (lastUpdated && (!latestAgentTime[crmName] || lastUpdated > latestAgentTime[crmName])) {
      latestAgentTime[crmName] = lastUpdated;
    }

    // If data falls within the Date Filter
    if (lastUpdated && lastUpdated >= start && lastUpdated <= end) {
      
      totalCalls++;
      
      let dealerName = row[1] || "Unknown";
      let area = row[2] || "Unknown";
      let status = row[7] || "";
      let audioLink = row[8] || "";
      let orderAmt = parseFloat(row[10]) || 0;
      let duration = parseInt(row[16]) || 0; // Column Q (Call Duration)
      
      // Agent Stats Logic
      if (!agentStats[crmName]) {
        agentStats[crmName] = { rev: 0, calls: 0, totalSecs: 0 };
      }
      
      agentStats[crmName].calls++;
      agentStats[crmName].totalSecs += duration;
      
      // Area Stats Logic (Heatmap)
      if (!areaStats[area]) {
        areaStats[area] = 0;
      }

      // Add to Orders & Revenue if status is Order Received
      if (status === "Order Received") {
        totalOrders++;
        revenue += orderAmt;
        agentStats[crmName].rev += orderAmt;
        areaStats[area] += orderAmt;
      }
      
      // Unshift to recent activity to show latest first
      recentActivity.unshift({
        rowId: i + 1, 
        time: lastUpdated.toLocaleString('en-IN'), 
        crm: crmName, 
        dealer: dealerName,
        status: status, 
        amt: orderAmt, 
        audio: audioLink
      });
      
    }
  }
  
  // Calculate Idle Time Alerts (Checking against current real time)
  let now = new Date();
  
  for (let agent in latestAgentTime) {
    let diffMins = Math.floor((now - latestAgentTime[agent]) / 60000);
    
    // If agent hasn't logged anything in 45 mins
    if (diffMins > 45) { 
      idleAlerts.push(`${agent} (${diffMins} mins idle)`); 
    }
  }

  // Generate Leaderboard Array formatting average time
  let leaderboard = Object.keys(agentStats).map(function(k) { 
    
    let currentAgent = agentStats[k];
    
    // Calculate Average Seconds
    let avgSecs = currentAgent.calls > 0 ? Math.floor(currentAgent.totalSecs / currentAgent.calls) : 0;
    
    // Convert to String Format (e.g., "2m 15s")
    let avgStr = Math.floor(avgSecs / 60) + "m " + (avgSecs % 60) + "s";
    
    return { 
      name: k, 
      rev: currentAgent.rev, 
      calls: currentAgent.calls, 
      avgTime: avgStr 
    }; 
    
  });
  
  // Sort Leaderboard descending by Revenue
  leaderboard.sort(function(a, b) { 
    return b.rev - a.rev; 
  });
  
  // Generate Heatmap Array
  let heatmap = Object.keys(areaStats).map(function(k) { 
    return { 
      area: k, 
      rev: areaStats[k] 
    }; 
  });
  
  // Sort Heatmap and take top 5
  heatmap.sort(function(a, b) { 
    return b.rev - a.rev; 
  }).slice(0, 5); 

  // Return the entire payload to UI
  return { 
    calls: totalCalls, 
    orders: totalOrders, 
    revenue: revenue, 
    feed: recentActivity.slice(0, 50), 
    leaderboard: leaderboard, 
    heatmap: heatmap, 
    idleAlerts: idleAlerts 
  };
}


/**
 * SUBMIT QA RATING
 * Allows CEO to rate the call and provide feedback from the dashboard.
 */
function submitQaRating(rowNumber, rating, feedback) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dealer Database");
  sheet.getRange(rowNumber, 15).setValue(rating); // Column O
  sheet.getRange(rowNumber, 16).setValue(feedback); // Column P
  return "Audit Saved";
}


/**
 * PLACEHOLDERS FOR FUTURE FEATURES
 * Left uncompressed to avoid UI breaking if called accidentally.
 */

// Cloud Telephony Initiation
function initiateCloudCall(dealerPhone, agentPhone) {
  Logger.log("Initiating API call to connect: " + dealerPhone);
  return "Dialing initiated for " + dealerPhone + ".";
}

// AI Call Summarization
function summarizeWithAI(textNotes) {
  return "AI Summary feature requires API key setup.";
}

// EOD PDF Report
function sendEODReport() {
  Logger.log("EOD Report generation triggered.");
}

// ==============================================================================
// SALES ENQUIRY ENGINE (PIPELINE: ENQUIRY RECEIVED -> ENQUIRY CLOSED)
// ==============================================================================

/**
 * HELPER: GET OR CREATE "Sales Enquiries" SHEET
 */
function getSalesEnquirySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Sales Enquiries");
  if (!sheet) {
    sheet = ss.insertSheet("Sales Enquiries");
    sheet.appendRow([
      "Enquiry ID",          // Col 1: A
      "Created At",          // Col 2: B
      "Dealer Name",         // Col 3: C
      "Phone",               // Col 4: D
      "Area",                // Col 5: E
      "Products Required",   // Col 6: F
      "Estimated Value",     // Col 7: G
      "Priority",            // Col 8: H (Hot, Warm, Cold)
      "Current Stage",       // Col 9: I (Enquiry Received, Quotation / Discussion, Follow-up Scheduled, Converted to Order, Enquiry Closed)
      "Assigned Agent",      // Col 10: J
      "Next Follow-up Date", // Col 11: K
      "Final Order Amount",  // Col 12: L
      "Linked Order ID",     // Col 13: M
      "Closure Reason",      // Col 14: N
      "Activity Notes",      // Col 15: O
      "Last Updated At"      // Col 16: P
    ]);
    sheet.getRange(1, 1, 1, 16).setFontWeight("bold").setBackground("#e8f0fe");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * ADD SALES ENQUIRY
 * Records a new dealer sales enquiry into "Sales Enquiries" sheet.
 */
function addSalesEnquiry(data) {
  try {
    const sheet = getSalesEnquirySheet_();
    const timestamp = new Date();
    const autoEnqId = "ENQ-" + Math.floor(10000 + Math.random() * 90000);
    
    const dealerName = (data.dealerName || "").toString().trim();
    const phone = (data.phone || "").toString().trim();
    const area = (data.area || "").toString().trim();
    const products = (data.products || "").toString().trim();
    const estValue = parseFloat(data.estValue) || 0;
    const priority = data.priority || "Warm";
    const stage = "Enquiry Received"; // Initial stage
    const agent = (data.agent || "").toString().trim();
    const followUpDate = data.followUpDate || "";
    const notes = (data.notes || "").toString().trim();
    
    const initialNoteHistory = `[${timestamp.toLocaleString('en-IN')}] Enquiry created by ${agent}: ${notes || 'Initial enquiry logged'}`;
    
    sheet.appendRow([
      autoEnqId,
      timestamp,
      dealerName,
      phone,
      area,
      products,
      estValue,
      priority,
      stage,
      agent,
      followUpDate,
      0,                  // Final Order Amount
      "",                 // Linked Order ID
      "",                 // Closure Reason
      initialNoteHistory, // Activity Notes
      timestamp           // Last Updated At
    ]);
    
    return {
      success: true,
      enquiryId: autoEnqId,
      msg: `Sales Enquiry ${autoEnqId} successfully created for ${dealerName}!`
    };
  } catch (err) {
    Logger.log("addSalesEnquiry Error: " + err.toString());
    return { success: false, msg: "Error adding enquiry: " + err.toString() };
  }
}

/**
 * GET SALES ENQUIRIES
 * Returns all active & historic sales enquiries with stage counts.
 */
function getSalesEnquiries(staffName, role) {
  try {
    const sheet = getSalesEnquirySheet_();
    const data = sheet.getDataRange().getValues();
    
    let list = [];
    let counts = {
      total: 0,
      received: 0,
      discussion: 0,
      followup: 0,
      converted: 0,
      closed: 0
    };
    
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let enqId = row[0];
      if (!enqId) continue;
      
      let createdDate = row[1] ? new Date(row[1]).toLocaleString('en-IN') : "";
      let dName = row[2] || "";
      let phone = row[3] || "";
      let area = row[4] || "";
      let products = row[5] || "";
      let estValue = parseFloat(row[6]) || 0;
      let priority = row[7] || "Warm";
      let stage = row[8] || "Enquiry Received";
      let assignedAgent = row[9] || "";
      let followUpDate = row[10] ? (row[10] instanceof Date ? row[10].toLocaleDateString('en-IN') : String(row[10])) : "";
      let finalAmt = parseFloat(row[11]) || 0;
      let orderId = row[12] || "";
      let closeReason = row[13] || "";
      let notes = row[14] || "";
      let lastUpdated = row[15] ? new Date(row[15]).toLocaleString('en-IN') : "";
      
      counts.total++;
      if (stage === "Enquiry Received") counts.received++;
      else if (stage === "Quotation / Discussion") counts.discussion++;
      else if (stage === "Follow-up Scheduled") counts.followup++;
      else if (stage === "Converted to Order") counts.converted++;
      else if (stage === "Enquiry Closed") counts.closed++;
      
      list.unshift({
        rowNumber: i + 1,
        enquiryId: enqId,
        createdAt: createdDate,
        dealerName: dName,
        phone: phone,
        area: area,
        products: products,
        estValue: estValue,
        priority: priority,
        stage: stage,
        assignedAgent: assignedAgent,
        followUpDate: followUpDate,
        finalOrderAmt: finalAmt,
        linkedOrderId: orderId,
        closureReason: closeReason,
        notes: notes,
        lastUpdated: lastUpdated
      });
    }
    
    return {
      success: true,
      enquiries: list,
      counts: counts
    };
  } catch (err) {
    Logger.log("getSalesEnquiries Error: " + err.toString());
    return { success: false, enquiries: [], counts: {}, error: err.toString() };
  }
}

/**
 * UPDATE SALES ENQUIRY STAGE
 * Moves enquiry through lifecycle stages:
 * Enquiry Received -> Quotation / Discussion -> Follow-up Scheduled -> Converted to Order -> Enquiry Closed
 */
function updateSalesEnquiryStage(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getSalesEnquirySheet_();
    const data = sheet.getDataRange().getValues();
    const timestamp = new Date();
    
    let targetRow = -1;
    let currentEnq = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === payload.enquiryId.toString().trim()) {
        targetRow = i + 1;
        currentEnq = data[i];
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, msg: "Enquiry ID not found: " + payload.enquiryId };
    }
    
    const newStage = payload.stage;
    const notesUpdate = payload.notes || "";
    const staffName = payload.staffName || "Staff";
    let linkedOrderId = currentEnq[12] || "";
    let finalOrderAmt = parseFloat(payload.orderAmt) || parseFloat(currentEnq[11]) || 0;
    let followUpDate = payload.followUpDate !== undefined ? payload.followUpDate : (currentEnq[10] || "");
    let closeReason = payload.closingReason !== undefined ? payload.closingReason : (currentEnq[13] || "");
    
    // Update Stage (Col I: 9)
    sheet.getRange(targetRow, 9).setValue(newStage);
    
    // Update Follow-up Date (Col K: 11)
    if (followUpDate) {
      sheet.getRange(targetRow, 11).setValue(followUpDate);
    }
    
    // If Converted to Order, push directly to "Order Book"
    if (newStage === "Converted to Order" && finalOrderAmt > 0) {
      if (!linkedOrderId) {
        linkedOrderId = "YM-" + Math.floor(10000 + Math.random() * 90000);
      }
      
      let orderSheet = ss.getSheetByName("Order Book");
      if (!orderSheet) {
        orderSheet = ss.insertSheet("Order Book");
        orderSheet.appendRow(["Date & Time", "OrderID", "ShopName", "DealerPhone", "OrderItems", "TotalAmount", "Payment Mode", "Additional Discount"]);
      }
      
      const orderItems = payload.orderItems || currentEnq[5] || "Enquiry Converted Order";
      const payMode = payload.payMode || "Credit / Bill";
      
      orderSheet.appendRow([
        timestamp,
        linkedOrderId,
        currentEnq[2], // Dealer Name
        currentEnq[3], // Phone
        orderItems,
        finalOrderAmt,
        payMode,
        "Enquiry: " + payload.enquiryId
      ]);
      
      sheet.getRange(targetRow, 12).setValue(finalOrderAmt); // Col L: Final Order Amount
      sheet.getRange(targetRow, 13).setValue(linkedOrderId);  // Col M: Linked Order ID
    }
    
    // If Enquiry Closed, store reason
    if (newStage === "Enquiry Closed" && closeReason) {
      sheet.getRange(targetRow, 14).setValue(closeReason); // Col N: Closure Reason
    }
    
    // Append to Activity Notes (Col O: 15)
    let prevNotes = currentEnq[14] || "";
    let newEntry = `[${timestamp.toLocaleString('en-IN')}] Stage changed to "${newStage}" by ${staffName}${notesUpdate ? ': ' + notesUpdate : ''}${linkedOrderId ? ' (Order: ' + linkedOrderId + ')' : ''}`;
    let combinedNotes = prevNotes ? (prevNotes + "\n" + newEntry) : newEntry;
    sheet.getRange(targetRow, 15).setValue(combinedNotes);
    
    // Last Updated At (Col P: 16)
    sheet.getRange(targetRow, 16).setValue(timestamp);
    
    return {
      success: true,
      msg: `Enquiry ${payload.enquiryId} updated to "${newStage}" successfully!`,
      orderId: linkedOrderId
    };
  } catch (err) {
    Logger.log("updateSalesEnquiryStage Error: " + err.toString());
    return { success: false, msg: "Error updating enquiry: " + err.toString() };
  }
}

/**
 * GET DEALERS FOR ENQUIRY DATALIST
 * Returns list of known dealers from Dealer Database for smart autocomplete.
 */
function getDealersForEnquiryDatalist() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName("Dealer Database");
    if (!dbSheet) return [];
    
    const data = dbSheet.getDataRange().getValues();
    let dealers = [];
    let seen = {};
    
    for (let i = 1; i < data.length; i++) {
      let phone = data[i][0] ? data[i][0].toString().trim() : "";
      let name = data[i][1] ? data[i][1].toString().trim() : "";
      let area = data[i][2] ? data[i][2].toString().trim() : "";
      
      if (name && !seen[name.toLowerCase()]) {
        seen[name.toLowerCase()] = true;
        dealers.push({ name: name, phone: phone, area: area });
      }
    }
    return dealers;
  } catch (err) {
    Logger.log("getDealersForEnquiryDatalist Error: " + err.toString());
    return [];
  }
}
