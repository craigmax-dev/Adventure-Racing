function preprocessData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registerSheet = ss.getSheetByName('Register');
  const pointsMapSheet = ss.getSheetByName('Points Map');
  const squareSidesSheet = ss.getSheetByName('Square Sides');
  const puzzleMultiplierSheet = ss.getSheetByName('Puzzle Multipliers');

  // Get data from Register, Points Map, Square Sides, and Puzzle Multipliers tabs
  const registerData = registerSheet.getDataRange().getValues();
  const pointsMapData = pointsMapSheet.getDataRange().getValues();
  const squareSidesData = squareSidesSheet.getDataRange().getValues();
  const puzzleMultiplierData = puzzleMultiplierSheet.getDataRange().getValues();

  // Create a map of Orienteering Points to their corresponding type (LP, PP, WP)
  const pointsMap = {};
  for (let i = 1; i < pointsMapData.length; i++) {
    const point = pointsMapData[i][0];
    const lp = pointsMapData[i][1];
    const pp = pointsMapData[i][2];
    const wp = pointsMapData[i][3];
    pointsMap[point] = { lp, pp, wp };
  }

  // Create a map of letter sides from Square Sides tab
  const letterSides = {};
  for (let i = 0; i < squareSidesData.length; i++) {
    const sideLetters = squareSidesData[i].filter(letter => letter !== '');
    sideLetters.forEach(letter => {
      letterSides[letter] = i; // Map each letter to its side index
    });
  }

  // Create a map of puzzle multipliers from Puzzle Multipliers tab
  const puzzleMultipliers = {};
  for (let i = 1; i < puzzleMultiplierData.length; i++) {
    const puzzleCode = puzzleMultiplierData[i][0];
    const multiplier = puzzleMultiplierData[i][1];
    puzzleMultipliers[puzzleCode] = multiplier;
  }

  // Sort register data by Time and then by Team
  const sortedData = registerData.slice(1).sort((a, b) => {
    const timeA = new Date(`1970-01-01T${a[2]}Z`);
    const timeB = new Date(`1970-01-01T${b[2]}Z`);
    if (timeA - timeB !== 0) {
      return timeA - timeB;
    }
    return a[0].localeCompare(b[0]);
  });

  // Add headers back to sorted data (Action column already exists)
  sortedData.unshift(registerData[0]);

  // Preprocess the Input, Correct, and Action columns based on the type of point visited
  const teamLetters = {}; // To track letters collected by each team
  const teamLastSide = {}; // To track the last side used by each team
  const teamPendingPuzzles = {}; // To track pending puzzles for each team
  for (let i = 1; i < sortedData.length; i++) {
    const [team, point, time, input, correct] = sortedData[i];
    let action = sortedData[i][5] || '';

    if (!teamLetters[team]) {
      teamLetters[team] = '';
      teamLastSide[team] = null;
      teamPendingPuzzles[team] = [];
    }

    if (point === 'WP') {
      // WP visited: Input should contain the word entered, then reset to the last letter
      sortedData[i][3] = teamLetters[team];
      action = `Word Submitted - ${sortedData[i][3]}`;
      teamLetters[team] = teamLetters[team].slice(-1); // Keep only the last letter immediately after word is played
    } else if (pointsMap[point] && pointsMap[point].lp) {
      // LP visited: check if the letter can be added
      const letter = pointsMap[point].lp;
      const letterSide = letterSides[letter];
      if (teamLastSide[team] !== null && teamLastSide[team] === letterSide) {
        // Invalid move: same side as the previous letter
        sortedData[i][4] = 0;
        action = 'Invalid LP - Same side as previous letter';
      } else {
        // Valid move: add letter to teamLetters and update Input
        teamLetters[team] += letter;
        teamLastSide[team] = letterSide;
        sortedData[i][3] = teamLetters[team];
        sortedData[i][4] = 1;
        action = `LP Added - ${letter}`;
      }
    } else if (pointsMap[point] && pointsMap[point].pp) {
      // PP visited: record puzzle as unlocked (no multiplier relevant here)
      teamPendingPuzzles[team].push(point);
      action = 'PP Visited - Puzzle Unlocked';
    } else if (point.startsWith('PZ')) {
      // Point visited is a puzzle submission (e.g., PZXX): set Input to the multiplier
      const multiplier = puzzleMultipliers[point] || 'N/A';
      sortedData[i][3] = multiplier;
      action = `Puzzle Multiplier Applied - ${point} (Multiplier: ${multiplier})`;
    } else if (point === 'SFP') {
      // SFP visited to complete the stage
      sortedData[i][3] = ''; // No Input value when SFP is visited
      action = 'SFP Visited - Stage Completed';
    }

    sortedData[i][5] = action;
  }

  // Write the sorted and processed data back to the Register sheet
  registerSheet.clear();
  registerSheet.getRange(1, 1, sortedData.length, sortedData[0].length).setValues(sortedData);
}
