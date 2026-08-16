Built a GitHub Issue Auto-Labeler using only Google Apps Script + free GitHub API.

No GitHub Actions. No paid bots. No servers.

What it does:
→ Fetches open issues from any repo via GitHub REST API
→ Reads title + body, runs keyword-based NLP rules
→ Auto-applies labels: bug, enhancement, documentation, security, etc.
→ Creates missing labels automatically if they don't exist
→ Full audit log of every labeled issue

The "NLP" is simple but brutal:
"crash" + "error" + "broken" → bug
"feature" + "request" + "add" → enhancement
"security" + "vulnerability" + "CVE" → security

Rules are fully editable in a Google Sheet. Add your own keywords, labels, priorities.

Why I built this:
Maintainers on small teams waste 2-3 hours/week triaging unlabeled issues. GitHub's built-in auto-labeling is paywalled. Actions YAML is overkill for a solo dev.

This runs on a free GitHub Personal Access Token + one GAS trigger. One file. Paste and run.

Code is 703 lines. One Sheet. One PAT. One button.
