const SHEET_ID = "1dvf5f_dFWX61uMK_Tt5-01VdoTTil79SHK5iaz6jnRw";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Leave Management Dashboard")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initUserSession() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const data = ss.getSheetByName("Agents Database").getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase()) {
      return { 
        email: email, name: data[i][1], 
        role: String(data[i][3]).toLowerCase().trim(), 
        team: String(data[i][4]).toLowerCase().trim() 
      };
    }
  }
  return { email: email, name: "Unknown User", role: "employee", team: "none" };
}

function getCalendarDetails() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Leave Requests").getDataRange().getValues();
  let countMap = {};
  let namesMap = {};
  for (let i = 1; i < sh.length; i++) {
    const status = (sh[i][6] || "").toUpperCase();
    if (status === "REJECTED" || status === "DECLINED") continue;
    let s = new Date(sh[i][4]);
    let e = new Date(sh[i][5]);
    for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
      let key = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
      countMap[key] = (countMap[key] || 0) + 1;
      if (!namesMap[key]) namesMap[key] = [];
      namesMap[key].push(sh[i][2]);
    }
  }
  return { countMap: countMap, namesMap: namesMap };
}

function getLeaveCredits(user) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Leave Requests").getDataRange().getValues();
  let used = 0;
  for (let i = 1; i < sh.length; i++) {
    if (sh[i][1].toLowerCase() === user.email.toLowerCase() && (sh[i][6] || "").toUpperCase() === "APPROVED") {
      let s = new Date(sh[i][4]);
      let e = new Date(sh[i][5]);
      used += Math.round((e - s) / 86400000) + 1;
    }
  }
  let today = new Date();
  let total = Math.floor((today.getMonth() + 1) * 1.6667 * 100) / 100;
  return { total: total, used: used, remaining: Math.max(0, (total - used).toFixed(2)) };
}

function getManagerDashboardData(user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const agents = ss.getSheetByName("Agents Database").getDataRange().getValues();
  const requests = ss.getSheetByName("Leave Requests").getDataRange().getValues();
  let teamData = [];
  const today = new Date();
  const totalAccrued = Math.floor((today.getMonth() + 1) * 1.6667 * 100) / 100;

  for (let i = 1; i < agents.length; i++) {
    let empEmail = agents[i][0];
    let empTeam = String(agents[i][4]).toLowerCase().trim();
    
    if (user.role === "operations manager" || empTeam === user.team.toLowerCase()) {
      let empName = agents[i][1];
      let used = 0;
      let pendingList = [];
      let historyList = []; // Captures the exact days they are off
      
      for (let j = 1; j < requests.length; j++) {
        if (requests[j][1].toLowerCase() === empEmail.toLowerCase()) {
          let status = (requests[j][6] || "").toUpperCase();
          if (status === "APPROVED") {
            let s = new Date(requests[j][4]);
            let e = new Date(requests[j][5]);
            used += Math.round((e - s) / 86400000) + 1;
            let rangeStr = s.getTime() === e.getTime() ? 
               Utilities.formatDate(s, "GMT", "MMM dd") : 
               `${Utilities.formatDate(s, "GMT", "MMM dd")}-${Utilities.formatDate(e, "GMT", "MMM dd")}`;
            historyList.push(rangeStr);
          } else if (status === "PENDING") {
            let s = new Date(requests[j][4]);
            let e = new Date(requests[j][5]);
            let rangeStr = s.getTime() === e.getTime() ? 
               Utilities.formatDate(s, "GMT", "MMM dd") : 
               `${Utilities.formatDate(s, "GMT", "MMM dd")}-${Utilities.formatDate(e, "GMT", "MMM dd")}`;
               
            pendingList.push({ row: j + 1, name: empName, range: rangeStr });
          }
        }
      }
      teamData.push({ 
        name: empName, 
        team: empTeam, 
        remaining: Math.max(0, (totalAccrued - used).toFixed(2)), 
        pending: pendingList,
        history: historyList.length > 0 ? historyList.slice(-3).reverse() : ["No approved leaves"]
      });
    }
  }
  return teamData;
}

function updateLeaveStatus(row, status) {
  const ss = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Leave Requests");
  ss.getRange(row, 7).setValue(status);
  return { ok: true, msg: `Leave ${status}` };
}

function submitLeave(d) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const start = new Date(d.start + "T00:00:00");
  const end = new Date(d.end + "T00:00:00");
  
  const today = new Date(); today.setHours(0,0,0,0);
  const minLead = new Date(today); minLead.setDate(minLead.getDate() + 30);
  if (start < minLead) return { ok: false, msg: "Error: 30 days advance notice required." };

  ss.getSheetByName("Leave Requests").appendRow([new Date(), d.email, d.name, d.team, start, end, d.status, d.reason]);
  return { ok: true, msg: d.status === "APPROVED" ? "Leave Approved!" : "Insufficient credits. Submission is PENDING." };
}
