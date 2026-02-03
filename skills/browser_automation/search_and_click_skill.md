# Browser Automation Skill: Search and Click

## Description
This skill opens a browser, performs a search, and clicks on a specific search result.

## Steps
1. Launch browser (preferably using Start-Process)
2. Identify search input field using State-Tool
3. Fill search query in the correct field
4. Submit search
5. Identify search results
6. Click on the specified result

## Improved Process Based on Actual Execution
1. Launch browser using PowerShell: `Start-Process msedge`
2. Use State-Tool to identify search input field
3. Fill search query in the correct field (typically the address bar or search box)
4. Submit search using Enter key
5. Use State-Tool to identify search results page
6. Visually identify the correct result (based on user requirements)
7. If needed, scroll down to find the specific result
8. Click on the specified result using its coordinates

## Key Insights from Execution
- Always use State-Tool to accurately identify UI elements before interacting with them
- Search results may require scrolling to find the desired result
- Visual identification of results is important - the "third result" may not be the third item in the HTML structure
- Coordinate-based clicking is precise but requires accurate element identification
- Browser back/forward navigation may be needed to return to search results if wrong link is clicked

## Common Issues and Solutions
- **Search Loop Issue**: If the search doesn't return expected results, verify the search query is correctly entered in the address bar
- **Search Bar Clearing**: Before entering a new search query, ensure the search bar is cleared properly
- **Focus Problems**: If the browser window loses focus, use State-Tool to verify the current window state
- **Window Switching**: If multiple browser windows are open, use appropriate tools to switch to the correct window
- **Process Restart**: If the browser gets stuck, it may be necessary to kill and restart the browser process
- **Empty Results**: If search returns no results, verify the query syntax and try alternative search approaches

## Analysis Summary
- Redundant steps: Initially trying App-Tool to launch Edge when PowerShell would work better
- Errors encountered: Msedge not found in start menu, confusion between address bar and search box, misidentifying the target search result
- Key insight: Always use State-Tool to accurately identify UI elements before interacting with them
- Success factor: Careful visual identification of the target result as it appears to a human user
- Recovery strategies: When stuck, restart the browser process and try a different approach