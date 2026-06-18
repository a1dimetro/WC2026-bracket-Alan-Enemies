// Data for the leaderboard
const leaderboardData = [
  { bracket: 'Alga Kazakhstan', week1: 20, week2: 2 },
  { bracket: 'Blue Lock', week1: 15, week2: 1 },
  { bracket: 'Ali', week1: 15, week2: 1 }
];

/**
 * Populate the leaderboard table based on the data above.
 */
function populateLeaderboard() {
  const tbody = document.getElementById('leaderboard-body');
  // sort data by total descending
  const sorted = leaderboardData.slice().sort((a, b) => (b.week1 + b.week2) - (a.week1 + a.week2));
  // build table rows
  sorted.forEach((item, index) => {
    const tr = document.createElement('tr');
    // rank
    const rankTd = document.createElement('td');
    rankTd.textContent = index + 1;
    tr.appendChild(rankTd);
    // bracket name
    const bracketTd = document.createElement('td');
    bracketTd.textContent = item.bracket;
    tr.appendChild(bracketTd);
    // total points
    const totalTd = document.createElement('td');
    totalTd.textContent = item.week1 + item.week2;
    tr.appendChild(totalTd);
    // week1
    const w1Td = document.createElement('td');
    w1Td.textContent = item.week1;
    tr.appendChild(w1Td);
    // week2
    const w2Td = document.createElement('td');
    w2Td.textContent = item.week2;
    tr.appendChild(w2Td);
    tbody.appendChild(tr);
  });
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