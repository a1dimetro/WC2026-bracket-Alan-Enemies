/**
 * Fetch leaderboard data from the generated JSON file and populate the table.
 */
async function populateLeaderboard() {
  const tbody = document.getElementById('leaderboard-body');
  tbody.innerHTML = '';
  try {
    const response = await fetch('leaderboard.json');
    if (!response.ok) {
      throw new Error('Failed to load leaderboard data');
    }
    const data = await response.json();
    // data is an array of objects: {bracket, total_points}
    data.forEach((item, index) => {
      const tr = document.createElement('tr');
      const rankTd = document.createElement('td');
      rankTd.textContent = index + 1;
      tr.appendChild(rankTd);
      const bracketTd = document.createElement('td');
      bracketTd.textContent = item.bracket;
      tr.appendChild(bracketTd);
      const totalTd = document.createElement('td');
      totalTd.textContent = item.total_points;
      tr.appendChild(totalTd);
      // Group and knockout breakdowns are not yet computed in the JSON; leave blank or use placeholders
      const groupTd = document.createElement('td');
      groupTd.textContent = '-';
      tr.appendChild(groupTd);
      const knockoutTd = document.createElement('td');
      knockoutTd.textContent = '-';
      tr.appendChild(knockoutTd);
      tbody.appendChild(tr);
    });
  } catch (err) {
    const tr = document.createElement('tr');
    const td = document.createElement('td');
    td.setAttribute('colspan', '5');
    td.textContent = 'Error loading leaderboard.';
    tr.appendChild(td);
    tbody.appendChild(tr);
    console.error(err);
  }
}

/**
 * Set the update time to the current local date and time.
 */
function setUpdateTime() {
  const now = new Date();
  const options = { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' };
  const formatted = now.toLocaleString(undefined, options);
  document.getElementById('update-time').textContent = formatted;
}

/**
 * Toggle dark/light theme.
 */
function setupThemeToggle() {
  const toggle = document.getElementById('themeToggle');
  // Initialize theme based on local storage or default to light
  const savedTheme = localStorage.getItem('theme');
  if (savedTheme === 'dark') {
    document.body.classList.add('dark');
  } else {
    document.body.classList.add('light');
  }
  toggle.addEventListener('click', () => {
    if (document.body.classList.contains('dark')) {
      document.body.classList.remove('dark');
      document.body.classList.add('light');
      localStorage.setItem('theme', 'light');
    } else {
      document.body.classList.remove('light');
      document.body.classList.add('dark');
      localStorage.setItem('theme', 'dark');
    }
  });
}

// Initialise the page
document.addEventListener('DOMContentLoaded', () => {
  populateLeaderboard();
  setUpdateTime();
  setupThemeToggle();
});