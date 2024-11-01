function calculateScores() {
  function convertTimeToMinutes(timeStr) {
    if (typeof timeStr === 'string') {
      const [hrs, mins, secs] = timeStr.split(':').map(Number);
      return hrs * 60 + mins + secs / 60;
    }
    return timeStr; // If already a number, return as is
  }
  function convertMinutesToTimeFormat(minutes) {
    const sign = minutes < 0 ? '-' : '';
    const absMinutes = Math.abs(minutes);
    const hrs = Math.floor(absMinutes / 60);
    const mins = Math.floor(absMinutes % 60);
    const secs = Math.round((absMinutes * 60) % 60);
    return `${sign}${hrs.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registerSheet = ss.getSheetByName('Register');
  const scoresheet = ss.getSheetByName('Scoresheet');
  const puzzleMultiplierSheet = ss.getSheetByName('Puzzle Multipliers');
  const wordLengthBonusSheet = ss.getSheetByName('Word Length Bonus');

  // Clear previous results from Scoresheet
  scoresheet.clearContents();

  // Get data from Register tab
  const registerData = registerSheet.getDataRange().getValues();

  // Create a structure to store team scores
  const teamScores = {};

  // Process Register data
  let stageStartTime = null;
  for (let i = 1; i < registerData.length; i++) {
    const [team, point, timestamp, input, correct] = registerData[i];
    const currentTime = new Date(timestamp);

    if (!teamScores[team]) {
      teamScores[team] = {
        time: 0,
        wordBonuses: [],
        puzzleMultipliers: [],
        allLettersBonus: 0,
        
        finishTime: 0,
        latePenalty: 0,
        totalPenalty: 0,
        totalBonus: 0,
        lettersUsed: new Set(),
        words: [],
        puzzlesSolved: 0,
        pendingPuzzleMultiplier: 1,
        progressiveWordChains: [],
        currentChain: [],
      };
    }

    if (point === 'SFP') {
      if (stageStartTime === null) {
        stageStartTime = currentTime;
      }
      teamScores[team].finishTime = (currentTime - stageStartTime) / (1000 * 60);
      // If the current chain is still active, finalize it
      if (teamScores[team].currentChain.length >= 2) {
        teamScores[team].progressiveWordChains.push(teamScores[team].currentChain);
      }
      teamScores[team].currentChain = [];
    }

    if (correct === 1) {
      if (point === 'WP' && input) {
        // Word submission
        const wordLength = input.length;
        const wordBonus = convertTimeToMinutes(getWordLengthBonus(wordLength));
        teamScores[team].wordBonuses.push(wordBonus);
        teamScores[team].puzzleMultipliers.push(teamScores[team].pendingPuzzleMultiplier); // Apply the pending multiplier to this word
        teamScores[team].words.push(input);

        // Check for progressive word chain starting from a 3-letter word and incrementing by 1
        if (
          (teamScores[team].currentChain.length === 0 && wordLength === 3) ||
          (teamScores[team].currentChain.length > 0 && wordLength === teamScores[team].currentChain[teamScores[team].currentChain.length - 1] + 1)
        ) {
          teamScores[team].currentChain.push(wordLength);
        } else {
          // End the current chain and start a new one if the word is 3 letters
          if (teamScores[team].currentChain.length >= 2) {
            teamScores[team].progressiveWordChains.push(teamScores[team].currentChain);
          }
          teamScores[team].currentChain = wordLength === 3 ? [wordLength] : [];
        }

        // Reset pending puzzle multiplier
        teamScores[team].pendingPuzzleMultiplier = 1;
      } else if (typeof point === 'string' && point.startsWith('PZ')) {
        // Puzzle multiplier submission
        const puzzleMultiplier = getPuzzleMultiplier(point);
        teamScores[team].puzzlesSolved += 1;
        // Set the multiplier to be applied to the next word
        teamScores[team].pendingPuzzleMultiplier = parseFloat(input) || puzzleMultiplier;
      }
    }

    const timeElapsed = stageStartTime ? (currentTime - stageStartTime) / (1000 * 60) : 0;
    teamScores[team].time = timeElapsed;
  }

  // Calculate bonuses and penalties
  Object.keys(teamScores).forEach(team => {
    const teamData = teamScores[team];

    // Calculate Word Bonus with Puzzle Multipliers
    teamData.wordBonus = teamData.wordBonuses.reduce((total, wordBonus, index) => {
      return total + (wordBonus * teamData.puzzleMultipliers[index]);
    }, 0);

    // Calculate All Letters Bonus
    const uniqueLettersUsed = new Set(teamData.words.join('').split(''));
    if (uniqueLettersUsed.size === 12) {
      teamData.allLettersBonus = 30;
    }

    // Calculate Spare Words Bonus
    const wordsUsed = teamData.words.length;
    if (wordsUsed < 5 && teamData.allLettersBonus > 0) {
      
    }

    // Calculate Progressive Word Chain Bonus
    teamData.progressiveWordChainBonus = teamData.progressiveWordChains.reduce((total, chain) => {
      if (chain.length >= 2) {
        const longestWordLength = Math.max(...chain);
        return total + longestWordLength * 5;
      }
      return total;
    }, 0);

    // Calculate total bonus and penalties
    teamData.totalBonus = teamData.wordBonus + teamData.allLettersBonus + teamData.progressiveWordChainBonus;
    teamData.totalPenalty = teamData.latePenalty;
  });

  // Sort teams by total bonus descending
  const sortedTeams = Object.keys(teamScores).sort((a, b) => a.localeCompare(b));

  const maxWords = Math.max(...Object.values(teamScores).map(team => team.words.length));
  let headers = ['Pos', 'Team Number', 'Final Time', 'Word Bonus', 'All Letters Bonus', 'Progressive Word Chain Bonus', 'Finish Time', 'Late Penalty', 'Total Penalty', 'Total Bonus'];
  for (let i = 1; i <= maxWords; i++) {
    headers.push('Word ' + i, 'Word ' + i + ' Bonus', 'Puzzle ' + i + ' Bonus');
  }
  scoresheet.appendRow(headers);

  // Append results to Scoresheet
  sortedTeams.forEach((team, index) => {
    const teamData = teamScores[team];
    const finalTime = convertMinutesToTimeFormat(teamData.finishTime + teamData.totalPenalty - teamData.totalBonus);
    const row = [
      index + 1,
      team,
      finalTime,
      convertMinutesToTimeFormat(teamData.wordBonus),
      convertMinutesToTimeFormat(teamData.allLettersBonus),
      
      convertMinutesToTimeFormat(teamData.progressiveWordChainBonus),
      convertMinutesToTimeFormat(teamData.finishTime),
      convertMinutesToTimeFormat(teamData.latePenalty),
      convertMinutesToTimeFormat(teamData.totalPenalty),
      convertMinutesToTimeFormat(teamData.totalBonus),
    ];

    // Add individual word and puzzle bonuses
    for (let i = 0; i < maxWords; i++) {
      if (i < teamData.words.length) {
        row.push(teamData.words[i]);
        const wordBonus = teamData.wordBonuses[i] * teamData.puzzleMultipliers[i];
        row.push(convertMinutesToTimeFormat(wordBonus));
        const puzzleBonus = teamData.puzzleMultipliers[i] > 1 ? teamData.puzzleMultipliers[i] : '';
        row.push(puzzleBonus !== '' ? puzzleBonus : '');
      } else {
        row.push('', '', '');
      }
    }
    scoresheet.appendRow(row);
  });
}

function getWordLengthBonus(wordLength) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wordLengthBonusSheet = ss.getSheetByName('Word Length Bonus');
  const bonuses = wordLengthBonusSheet.getDataRange().getValues();
  for (let i = 1; i < bonuses.length; i++) {
    if (bonuses[i][0] === wordLength) {
      return bonuses[i][1];
    }
  }
  return 0;
}

function getPuzzleMultiplier(puzzleId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const puzzleMultiplierSheet = ss.getSheetByName('Puzzle Multipliers');
  const multipliers = puzzleMultiplierSheet.getDataRange().getValues();
  for (let i = 1; i < multipliers.length; i++) {
    if (multipliers[i][0] === puzzleId) {
      return multipliers[i][1];
    }
  }
  return 1;
}
