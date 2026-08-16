/**
 * ╔══════════════════════════════════════════════════════════════════════════════╗
 * ║  GITHUB ISSUE AUTO-LABELER v1.0                                              ║
 * ║  Built with Google Apps Script + GitHub API (free) + Keyword NLP             ║
 * ║                                                                              ║
 * ║  Problem: Maintainers waste hours manually triaging and labeling issues.     ║
 * ║  GitHub auto-labeling requires paid plans or complex Actions workflows.      ║
 * ║                                                                              ║
 * ║  How it works:                                                               ║
 * ║  1. Run setupAutoLabeler() once — creates config sheets + label rules       ║
 * ║  2. Paste your GitHub Personal Access Token (free) in Config tab            ║
 * ║  3. Define keyword → label rules in Label_Rules tab                         ║
 * ║  4. Set polling trigger or click "Run Auto-Labeler Now"                     ║
 * ║  5. Script fetches new open issues, reads title/body, matches keywords      ║
 * ║  6. Calls GitHub API to auto-apply labels + logs everything                 ║
 * ║                                                                              ║
 * ║  Setup: Create Sheet → Extensions → Apps Script → Paste code                ║
 * ║         → Run setupAutoLabeler() → Fill Config + Label_Rules → Done         ║
 * ╚══════════════════════════════════════════════════════════════════════════════╝
 */

// ════════════════════════════════════════════════════════════════════════════════
// CONFIG
// ════════════════════════════════════════════════════════════════════════════════
const CONFIG = {
  TAB_CONFIG: 'Config',
  TAB_RULES: 'Label_Rules',
  TAB_LOG: 'Activity_Log',
  TAB_INSTRUCTIONS: 'Instructions',
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 2000,
  GITHUB_API_BASE: 'https://api.github.com'
};

// ════════════════════════════════════════════════════════════════════════════════
// MENU
// ════════════════════════════════════════════════════════════════════════════════
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('GitHubLabeler')
      .addItem('One-Time Setup', 'setupAutoLabeler')
      .addSeparator()
      .addItem('Run Auto-Labeler Now', 'runAutoLabeler')
      .addItem('Test GitHub Connection', 'testGitHubConnection')
      .addItem('Fetch & Preview (No Apply)', 'previewLabeling')
      .addSeparator()
      .addItem('Setup Hourly Trigger', 'setupHourlyTrigger')
      .addItem('Delete All Triggers', 'deleteAllTriggers')
      .addSeparator()
      .addItem('Help & Troubleshoot', 'helpAndTroubleshoot')
      .addToUi();
  } catch (e) {
    console.error('onOpen failed:', e);
  }
}

/**
 * RUN THIS MANUALLY IF MENU DOES NOT APPEAR
 * Extensions → Apps Script → Select "forceMenu" → Run
 */
function forceMenu() {
  onOpen();
  SpreadsheetApp.getUi().alert('Menu refreshed. Check the top menu bar.');
}

// ════════════════════════════════════════════════════════════════════════════════
// ONE-TIME SETUP
// ════════════════════════════════════════════════════════════════════════════════
function setupAutoLabeler() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. Config
    let configSheet = ss.getSheetByName(CONFIG.TAB_CONFIG);
    if (!configSheet) {
      configSheet = ss.insertSheet(CONFIG.TAB_CONFIG);
      configSheet.appendRow(['Setting', 'Value']);
      configSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      configSheet.appendRow(['GitHub_Repo', 'owner/repo-name']);
      configSheet.appendRow(['GitHub_PAT', 'ghp_xxxxxxxxxxxxxxxxxxxx']);
      configSheet.appendRow(['Last_Checked_Issue_ID', '0']);
      configSheet.appendRow(['Apply_Labels', 'TRUE']);
      configSheet.appendRow(['Create_Missing_Labels', 'TRUE']);
      configSheet.appendRow(['Max_Issues_Per_Run', '30']);
      configSheet.appendRow(['Include_Body_Text', 'TRUE']);
      configSheet.appendRow(['Case_Sensitive_Matching', 'FALSE']);
      configSheet.appendRow(['Manager_Email', Session.getActiveUser().getEmail()]);
    }

    // 2. Label Rules
    let rulesSheet = ss.getSheetByName(CONFIG.TAB_RULES);
    if (!rulesSheet) {
      rulesSheet = ss.insertSheet(CONFIG.TAB_RULES);
      rulesSheet.appendRow(['Label_Name', 'Keywords', 'Match_Mode', 'Priority', 'Color', 'Active']);
      rulesSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      // Default rules — edit or add more
      rulesSheet.appendRow(['bug', 'bug,crash,error,broken,fail,exception,not working,does not work', 'any', '1', 'd73a4a', 'TRUE']);
      rulesSheet.appendRow(['enhancement', 'feature,request,add,suggestion,improve,would be nice,want to', 'any', '2', 'a2eeef', 'TRUE']);
      rulesSheet.appendRow(['documentation', 'docs,documentation,readme,typo,spell,grammar,wiki', 'any', '3', '0075ca', 'TRUE']);
      rulesSheet.appendRow(['good first issue', 'beginner,newbie,first time,getting started,easy,starter', 'any', '4', '7057ff', 'TRUE']);
      rulesSheet.appendRow(['security', 'security,vulnerability,cve,exploit,auth,login,breach,xss,injection', 'any', '5', 'ff0000', 'TRUE']);
      rulesSheet.appendRow(['help wanted', 'help,how to,question,confused,stuck,not sure', 'any', '6', '008672', 'TRUE']);
      rulesSheet.appendRow(['duplicate', 'duplicate,already exists,same as,closing in favor', 'any', '7', 'cfd3d7', 'TRUE']);
      rulesSheet.appendRow(['performance', 'slow,lag,performance,timeout,freeze,hang,loading,optimize', 'any', '8', 'ff7619', 'TRUE']);

      const matchRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['any', 'all'], true).build();
      rulesSheet.getRange('C2:C1000').setDataValidation(matchRule);

      const boolRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['TRUE', 'FALSE'], true).build();
      rulesSheet.getRange('F2:F1000').setDataValidation(boolRule);
    }

    // 3. Activity Log
    let logSheet = ss.getSheetByName(CONFIG.TAB_LOG);
    if (!logSheet) {
      logSheet = ss.insertSheet(CONFIG.TAB_LOG);
      logSheet.appendRow(['Timestamp', 'Issue_Number', 'Issue_Title', 'Action', 'Labels_Applied', 'Status', 'Details']);
      logSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }

    // 4. Instructions
    let helpSheet = ss.getSheetByName(CONFIG.TAB_INSTRUCTIONS);
    if (!helpSheet) {
      helpSheet = ss.insertSheet(CONFIG.TAB_INSTRUCTIONS);
      helpSheet.appendRow(['GITHUB ISSUE AUTO-LABELER — Quick Start Guide']);
      helpSheet.appendRow(['']);
      helpSheet.appendRow(['STEP 1: Get GitHub Personal Access Token (FREE)']);
      helpSheet.appendRow(['→ Go to https://github.com/settings/tokens']);
      helpSheet.appendRow(['→ Click "Generate new token (classic)"']);
      helpSheet.appendRow(['→ Scopes needed: repo (full control of private repos) OR public_repo (public only)']);
      helpSheet.appendRow(['→ Copy the token (starts with ghp_)']);
      helpSheet.appendRow(['→ Paste it in Config!B3 (GitHub_PAT)']);
      helpSheet.appendRow(['']);
      helpSheet.appendRow(['STEP 2: Configure Your Repo']);
      helpSheet.appendRow(['→ Set Config!B2 to your repo: "owner/repo-name" (e.g., "facebook/react")']);
      helpSheet.appendRow(['']);
      helpSheet.appendRow(['STEP 3: Customize Label Rules']);
      helpSheet.appendRow(['→ Edit Label_Rules tab: add/remove keywords, change labels, adjust priority']);
      helpSheet.appendRow(['→ Match_Mode: "any" = at least one keyword found | "all" = every keyword must be found']);
      helpSheet.appendRow(['→ Priority: lower number = checked first (first match wins for conflicting labels)']);
      helpSheet.appendRow(['']);
      helpSheet.appendRow(['STEP 4: Test']);
      helpSheet.appendRow(['→ Click "Test GitHub Connection" to verify PAT + repo']);
      helpSheet.appendRow(['→ Click "Fetch & Preview" to see what labels WOULD be applied (no actual changes)']);
      helpSheet.appendRow(['→ Click "Run Auto-Labeler Now" to execute']);
      helpSheet.appendRow(['']);
      helpSheet.appendRow(['STEP 5: Automate']);
      helpSheet.appendRow(['→ Click "Setup Hourly Trigger" to auto-run every hour']);
      helpSheet.appendRow(['']);
      helpSheet.appendRow(['TROUBLESHOOTING']);
      helpSheet.appendRow(['• 401 Error = Bad PAT. Regenerate token with correct scopes.']);
      helpSheet.appendRow(['• 403 Error = Rate limited. Wait 1 hour or use authenticated requests (PAT fixes this).']);
      helpSheet.appendRow(['• 404 Error = Repo not found or no access. Check owner/repo-name format.']);
      helpSheet.appendRow(['• 422 Error = Label does not exist on repo. Set Config!B5 = TRUE to auto-create labels.']);
      helpSheet.appendRow(['• No labels applied = No keywords matched. Add more keywords to Label_Rules.']);
    }

    ui.alert('Setup Complete',
      'All tabs created.\n\nNEXT STEPS:\n1. Generate GitHub PAT at github.com/settings/tokens\n2. Paste it in Config!B3\n3. Set your repo in Config!B2\n4. Click "Test GitHub Connection"',
      ui.ButtonSet.OK);

    logActivity('Setup', '', '', 'SUCCESS', 'All tabs created');

  } catch (error) {
    logActivity('Setup', '', '', 'FAILED', error.message);
    ui.alert('Setup Error: ' + error.message);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// MAIN: RUN AUTO-LABELER
// ════════════════════════════════════════════════════════════════════════════════
function runAutoLabeler() {
  const ui = SpreadsheetApp.getUi();

  try {
    const config = loadConfig();
    const rules = loadLabelRules();

    // Validate config
    if (!config.GitHub_PAT || config.GitHub_PAT === 'ghp_xxxxxxxxxxxxxxxxxxxx') {
      throw new Error('GitHub PAT not configured. Paste your token in Config!B3.');
    }
    if (!config.GitHub_Repo || config.GitHub_Repo === 'owner/repo-name') {
      throw new Error('GitHub Repo not configured. Set owner/repo-name in Config!B2.');
    }
    if (rules.length === 0) {
      throw new Error('No active label rules found. Check Label_Rules tab.');
    }

    // Fetch issues
    const issues = fetchOpenIssues(config);
    if (issues.length === 0) {
      ui.alert('No new unlabeled issues found.');
      return;
    }

    let labeledCount = 0;
    let skippedCount = 0;

    for (const issue of issues) {
      // Skip pull requests (GitHub API returns PRs as issues too)
      if (issue.pull_request) {
        skippedCount++;
        continue;
      }

      // Skip if already has labels (optional — remove this check if you want to re-label)
      if (issue.labels && issue.labels.length > 0) {
        logActivity('Skip', issue.number, issue.title, 'SKIPPED', '', 'Already has labels: ' + issue.labels.map(l => l.name).join(', '));
        skippedCount++;
        continue;
      }

      // Analyze
      const matchedLabels = analyzeIssue(issue, rules, config);

      if (matchedLabels.length === 0) {
        logActivity('Analyze', issue.number, issue.title, 'NO_MATCH', '', 'No keywords matched');
        continue;
      }

      // Apply labels
      if (config.Apply_Labels === 'TRUE') {
        const result = applyLabelsToIssue(config, issue.number, matchedLabels);

        if (result.status === 'SUCCESS') {
          labeledCount++;
          logActivity('Label', issue.number, issue.title, 'SUCCESS', matchedLabels.join(', '), result.message);
        } else if (result.status === 'LABEL_MISSING' && config.Create_Missing_Labels === 'TRUE') {
          // Try creating missing labels then re-apply
          let createdAny = false;
          for (const label of matchedLabels) {
            const createResult = createGitHubLabel(config, label, rules);
            if (createResult.status === 'SUCCESS') createdAny = true;
          }

          if (createdAny) {
            // Retry applying labels
            const retryResult = applyLabelsToIssue(config, issue.number, matchedLabels);
            if (retryResult.status === 'SUCCESS') {
              labeledCount++;
              logActivity('Label', issue.number, issue.title, 'SUCCESS', matchedLabels.join(', '), 'Created missing labels then applied');
            } else {
              logActivity('Label', issue.number, issue.title, 'FAILED', matchedLabels.join(', '), retryResult.message);
            }
          } else {
            logActivity('Label', issue.number, issue.title, 'FAILED', matchedLabels.join(', '), 'Could not create missing labels');
          }
        } else {
          logActivity('Label', issue.number, issue.title, 'FAILED', matchedLabels.join(', '), result.message);
        }
      } else {
        // Preview mode
        logActivity('Preview', issue.number, issue.title, 'PREVIEW', matchedLabels.join(', '), 'Would apply these labels');
      }
    }

    // Update last checked
    updateLastChecked(config, issues);

    const mode = config.Apply_Labels === 'TRUE' ? 'Applied' : 'Previewed';
    ui.alert(`${mode} Complete`, 
      `Issues scanned: ${issues.length}\nLabeled: ${labeledCount}\nSkipped (PRs/already labeled): ${skippedCount}`,
      ui.ButtonSet.OK);

  } catch (error) {
    logActivity('Run', '', '', 'FAILED', '', error.message);
    ui.alert('Error: ' + error.message);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// PREVIEW MODE — Shows what WOULD be labeled without applying
// ════════════════════════════════════════════════════════════════════════════════
function previewLabeling() {
  const ui = SpreadsheetApp.getUi();

  try {
    const config = loadConfig();
    const rules = loadLabelRules();

    if (!config.GitHub_PAT || config.GitHub_PAT === 'ghp_xxxxxxxxxxxxxxxxxxxx') {
      throw new Error('GitHub PAT not configured.');
    }
    if (!config.GitHub_Repo || config.GitHub_Repo === 'owner/repo-name') {
      throw new Error('GitHub Repo not configured.');
    }

    const issues = fetchOpenIssues(config);
    if (issues.length === 0) {
      ui.alert('No open issues found.');
      return;
    }

    let previewRows = [];
    for (const issue of issues) {
      if (issue.pull_request) continue;
      const matched = analyzeIssue(issue, rules, config);
      if (matched.length > 0) {
        previewRows.push(`#${issue.number}: "${truncate(issue.title, 50)}" → ${matched.join(', ')}`);
      }
    }

    const msg = previewRows.length > 0 
      ? previewRows.join('\n') 
      : 'No issues matched any label rules.';

    ui.alert('Preview Results', msg, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('Preview Error: ' + error.message);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// TEST CONNECTION
// ════════════════════════════════════════════════════════════════════════════════
function testGitHubConnection() {
  const ui = SpreadsheetApp.getUi();

  try {
    const config = loadConfig();
    if (!config.GitHub_PAT || config.GitHub_PAT === 'ghp_xxxxxxxxxxxxxxxxxxxx') {
      throw new Error('GitHub PAT not configured in Config!B3.');
    }
    if (!config.GitHub_Repo || config.GitHub_Repo === 'owner/repo-name') {
      throw new Error('GitHub Repo not configured in Config!B2.');
    }

    const url = `${CONFIG.GITHUB_API_BASE}/repos/${config.GitHub_Repo}`;
    const result = safeGitHubRequest('get', url, null, config.GitHub_PAT);

    if (result.status === 'SUCCESS') {
      const repo = JSON.parse(result.body);
      ui.alert('Connection Successful',
        `Repo: ${repo.full_name}\nStars: ${repo.stargazers_count}\nOpen Issues: ${repo.open_issues_count}\n\nYou are ready to auto-label!`,
        ui.ButtonSet.OK);
      logActivity('Test', '', '', 'SUCCESS', '', `Connected to ${repo.full_name}`);
    } else {
      throw new Error(result.message);
    }
  } catch (error) {
    logActivity('Test', '', '', 'FAILED', '', error.message);
    ui.alert('Connection Failed', error.message, ui.ButtonSet.OK);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// GITHUB API HELPERS
// ════════════════════════════════════════════════════════════════════════════════
function fetchOpenIssues(config) {
  const maxIssues = parseInt(config.Max_Issues_Per_Run) || 30;
  const url = `${CONFIG.GITHUB_API_BASE}/repos/${config.GitHub_Repo}/issues?state=open&per_page=${maxIssues}&sort=created&direction=asc`;

  const result = safeGitHubRequest('get', url, null, config.GitHub_PAT);

  if (result.status !== 'SUCCESS') {
    throw new Error(`Failed to fetch issues: ${result.message}`);
  }

  return JSON.parse(result.body);
}

function applyLabelsToIssue(config, issueNumber, labels) {
  const url = `${CONFIG.GITHUB_API_BASE}/repos/${config.GitHub_Repo}/issues/${issueNumber}/labels`;
  const payload = { labels: labels };

  const result = safeGitHubRequest('post', url, payload, config.GitHub_PAT);

  if (result.status === 'SUCCESS') {
    return { status: 'SUCCESS', message: 'Labels applied' };
  }

  // Check for specific errors
  if (result.code === 422) {
    return { status: 'LABEL_MISSING', message: 'One or more labels do not exist on this repo' };
  }
  if (result.code === 403) {
    return { status: 'FAILED', message: 'Permission denied. Check your PAT has "repo" scope.' };
  }
  if (result.code === 404) {
    return { status: 'FAILED', message: 'Issue or repo not found.' };
  }

  return { status: 'FAILED', message: result.message };
}

function createGitHubLabel(config, labelName, rules) {
  const url = `${CONFIG.GITHUB_API_BASE}/repos/${config.GitHub_Repo}/labels`;

  // Find color from rules
  const rule = rules.find(r => r.Label_Name === labelName);
  const color = rule ? rule.Color : 'cccccc';

  const payload = {
    name: labelName,
    color: color.replace('#', ''),
    description: `Auto-created by GitHubLabeler`
  };

  const result = safeGitHubRequest('post', url, payload, config.GitHub_PAT);

  if (result.status === 'SUCCESS' || (result.code === 422 && result.body.includes('already_exists'))) {
    return { status: 'SUCCESS', message: `Label "${labelName}" ready` };
  }

  return { status: 'FAILED', message: result.message };
}

// ════════════════════════════════════════════════════════════════════════════════
// SAFE GITHUB REQUEST — Retry + specific error handling
// ════════════════════════════════════════════════════════════════════════════════
function safeGitHubRequest(method, url, payload, pat) {
  let lastError = null;

  for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
    try {
      const options = {
        method: method,
        headers: {
          'Authorization': `token ${pat}`,
          'Accept': 'application/vnd.github.v3+json',
          'User-Agent': 'GitHubLabeler-GAS'
        },
        muteHttpExceptions: true
      };

      if (payload && (method === 'post' || method === 'put')) {
        options.contentType = 'application/json';
        options.payload = JSON.stringify(payload);
      }

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const body = response.getContentText();

      if (code >= 200 && code < 300) {
        return { status: 'SUCCESS', code: code, body: body };
      }

      // Specific error handling
      if (code === 401) {
        return { status: 'FAILED', code: code, body: body, message: '401 Unauthorized — Bad GitHub PAT. Go to github.com/settings/tokens and regenerate with "repo" scope.' };
      }
      if (code === 403) {
        const rateLimitRemaining = response.getHeaders()['X-RateLimit-Remaining'];
        if (rateLimitRemaining === '0') {
          return { status: 'FAILED', code: code, body: body, message: '403 Rate Limited — GitHub API quota exhausted. Authenticated users get 5,000 requests/hour. Wait 1 hour.' };
        }
        return { status: 'FAILED', code: code, body: body, message: '403 Forbidden — PAT may lack "repo" scope, or you do not have write access to this repo.' };
      }
      if (code === 404) {
        return { status: 'FAILED', code: code, body: body, message: `404 Not Found — Repo "${url.split('/repos/')[1].split('/')[0]}/${url.split('/repos/')[1].split('/')[1]}" does not exist or is private (and PAT lacks access).` };
      }
      if (code === 422) {
        return { status: 'FAILED', code: code, body: body, message: '422 Validation Failed — Label may not exist, or issue is a pull request.' };
      }

      throw new Error(`HTTP ${code}: ${body.substring(0, 200)}`);

    } catch (error) {
      lastError = error;
      if (attempt < CONFIG.MAX_RETRIES) {
        Utilities.sleep(CONFIG.RETRY_DELAY_MS * attempt);
      }
    }
  }

  return { status: 'FAILED', code: 0, body: '', message: `Failed after ${CONFIG.MAX_RETRIES} attempts: ${lastError ? lastError.message : 'Unknown'}` };
}

// ════════════════════════════════════════════════════════════════════════════════
// NLP: KEYWORD MATCHING ENGINE
// ════════════════════════════════════════════════════════════════════════════════
function analyzeIssue(issue, rules, config) {
  const text = buildSearchText(issue, config);
  const searchText = config.Case_Sensitive_Matching === 'TRUE' ? text : text.toLowerCase();
  const matchedLabels = [];

  // Sort by priority (lower = first)
  const sortedRules = [...rules].sort((a, b) => parseInt(a.Priority) - parseInt(b.Priority));

  for (const rule of sortedRules) {
    if (rule.Active !== 'TRUE') continue;

    const keywords = rule.Keywords.split(',').map(k => k.trim()).filter(k => k.length > 0);
    if (keywords.length === 0) continue;

    const compareText = config.Case_Sensitive_Matching === 'TRUE' ? searchText : searchText.toLowerCase();
    const compareKeywords = config.Case_Sensitive_Matching === 'TRUE' ? keywords : keywords.map(k => k.toLowerCase());

    let isMatch = false;

    if (rule.Match_Mode === 'all') {
      // ALL keywords must be found
      isMatch = compareKeywords.every(kw => compareText.includes(kw));
    } else {
      // ANY keyword found = match
      isMatch = compareKeywords.some(kw => compareText.includes(kw));
    }

    if (isMatch) {
      matchedLabels.push(rule.Label_Name);
    }
  }

  return matchedLabels;
}

function buildSearchText(issue, config) {
  let text = issue.title || '';
  if (config.Include_Body_Text === 'TRUE' && issue.body) {
    text += ' ' + issue.body;
  }
  // Also include labels already present (for context, though we skip already-labeled issues)
  if (issue.labels && issue.labels.length > 0) {
    text += ' ' + issue.labels.map(l => l.name).join(' ');
  }
  return text;
}

// ════════════════════════════════════════════════════════════════════════════════
// DATA HELPERS
// ════════════════════════════════════════════════════════════════════════════════
function loadConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TAB_CONFIG);
  if (!sheet) throw new Error('Config tab not found. Run Setup first.');

  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).replace(/\s+/g, '_');
    config[key] = data[i][1];
  }
  return config;
}

function loadLabelRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TAB_RULES);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const rules = [];

  for (const row of rows) {
    const obj = {};
    for (let i = 0; i < headers.length; i++) {
      obj[String(headers[i]).replace(/\s+/g, '_')] = row[i];
    }
    if (obj.Active === 'TRUE') {
      rules.push(obj);
    }
  }
  return rules;
}

function updateLastChecked(config, issues) {
  if (issues.length === 0) return;

  // Find highest issue number
  const maxNumber = Math.max(...issues.map(i => i.number));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TAB_CONFIG);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).replace(/\s+/g, '_') === 'Last_Checked_Issue_ID') {
      sheet.getRange(i + 1, 2).setValue(maxNumber);
      break;
    }
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// TRIGGER MANAGEMENT
// ════════════════════════════════════════════════════════════════════════════════
function setupHourlyTrigger() {
  const ui = SpreadsheetApp.getUi();

  try {
    deleteAllTriggers();

    ScriptApp.newTrigger('runAutoLabeler')
      .timeBased()
      .everyHours(1)
      .create();

    ui.alert('Hourly trigger set. Auto-Labeler will run every hour.');
    logActivity('Trigger', '', '', 'SUCCESS', '', 'Hourly trigger created');
  } catch (error) {
    ui.alert('Trigger Error: ' + error.message);
  }
}

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  logActivity('Trigger', '', '', 'SUCCESS', '', `${triggers.length} triggers removed`);
}

// ════════════════════════════════════════════════════════════════════════════════
// LOGGING
// ════════════════════════════════════════════════════════════════════════════════
function logActivity(action, issueNum, issueTitle, status, labels, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.TAB_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.TAB_LOG);
      sheet.appendRow(['Timestamp', 'Issue_Number', 'Issue_Title', 'Action', 'Labels_Applied', 'Status', 'Details']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }
    sheet.appendRow([
      new Date(),
      issueNum || '',
      issueTitle || '',
      action,
      labels || '',
      status,
      details || ''
    ]);
  } catch (e) {
    console.error('Log failed:', e);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// UTILITIES
// ════════════════════════════════════════════════════════════════════════════════
function truncate(str, maxLen) {
  if (!str) return '';
  return str.length > maxLen ? str.substring(0, maxLen) + '...' : str;
}

// ════════════════════════════════════════════════════════════════════════════════
// HELP
// ════════════════════════════════════════════════════════════════════════════════
function helpAndTroubleshoot() {
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif;padding:20px;max-width:650px;">
      <h2>GitHub Issue Auto-Labeler Help</h2>

      <h3>How It Works</h3>
      <ol>
        <li>Script fetches open issues from your GitHub repo via API</li>
        <li>Reads issue title + body text</li>
        <li>Matches against keyword rules in Label_Rules tab</li>
        <li>Applies matched labels via GitHub API</li>
        <li>Logs every action in Activity_Log tab</li>
      </ol>

      <h3>Getting Your GitHub PAT (Personal Access Token)</h3>
      <ol>
        <li>Go to <a href="https://github.com/settings/tokens" target="_blank">github.com/settings/tokens</a></li>
        <li>Click "Generate new token (classic)"</li>
        <li>Select scope: <strong>repo</strong> (for private repos) or <strong>public_repo</strong> (for public only)</li>
        <li>Copy the token (starts with <code>ghp_</code>)</li>
        <li>Paste into Config!B3</li>
      </ol>

      <h3>Common Errors & Fixes</h3>
      <p><strong>401 Unauthorized</strong><br>→ Your PAT is invalid or expired. Regenerate it.</p>
      <p><strong>403 Forbidden / Rate Limited</strong><br>→ Unauthenticated users get 60 requests/hour. With PAT you get 5,000/hour. If you hit this, wait 1 hour.</p>
      <p><strong>404 Not Found</strong><br>→ Repo name is wrong. Use exact format: <code>owner/repo-name</code> (e.g., <code>facebook/react</code>)</p>
      <p><strong>422 Validation Failed</strong><br>→ Label does not exist on the repo. Set Config!B5 (Create_Missing_Labels) to TRUE.</p>
      <p><strong>No labels applied</strong><br>→ No keywords matched. Add more keywords to Label_Rules or check case sensitivity setting.</p>

      <h3>Label Rule Format</h3>
      <table border="1" cellpadding="5" style="border-collapse:collapse;font-size:13px;">
        <tr style="background:#f0f0f0;"><th>Column</th><th>Meaning</th></tr>
        <tr><td>Label_Name</td><td>Exact label as it appears on GitHub</td></tr>
        <tr><td>Keywords</td><td>Comma-separated words to search for</td></tr>
        <tr><td>Match_Mode</td><td>"any" = one keyword found | "all" = every keyword found</td></tr>
        <tr><td>Priority</td><td>Lower number = checked first</td></tr>
        <tr><td>Color</td><td>Hex color for auto-created labels (no #)</td></tr>
        <tr><td>Active</td><td>TRUE = rule is live</td></tr>
      </table>

      <h3>Pro Tips</h3>
      <ul>
        <li>Use "Fetch & Preview" before "Run Auto-Labeler Now" to see what would happen</li>
        <li>Set "Apply_Labels" to FALSE in Config to run in dry-run mode</li>
        <li>Script skips issues that already have labels (to avoid overwriting)</li>
        <li>Pull requests are automatically skipped</li>
      </ul>

      <p style="color:#666;font-size:12px;margin-top:20px;">GitHubLabeler v1.0 | GitHub API (free) | Zero cost | Built with GAS</p>
    </div>`
  ).setWidth(700).setHeight(750);

  SpreadsheetApp.getUi().showModalDialog(html, 'GitHubLabeler — Help');
}
