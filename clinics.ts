const TESTING = true;
const VERBOSE = false;
// TODO: constraint support
// Spread assignment - based on user pool per iteration
// TODO: Spanish general volunteer type
// TODO: datetime overlap & get sorting functionality back

//#region Objects [b]
interface User {
  Row: number
  Name: string
  MedYear: string
  Elective : boolean // Elective course (for 4th years only? - not enforced)
  ClinicsOfInterest: number[] // Clinics of interest for this user
  Type: UserType    // Default or Translator
  NumAssignments: number // Number of clninic dates this user has been assigned to
  ClinicRanks: string[] // first index is the highest preference
  DateIDsAssigned: Set<string> // Which dates has this user been assigned to (to prevent duplicate days)
}

enum UserType {
  Default,           // Default user type, goes through normal process
  SpanishVolunteer,  // TODO: A default user that can also serve as a translator (pool depends on the clinic)
  SpanishTranslator  // Special user type, assigned to a different pool
}

interface ClinicAssignment {
  Name: string
  ClinicIndex: number // Index corresponding to clinics list
  DateAvailColIndex: number // Index of availability col ex: "Select Date Availability for Agape Dermatology"
  AvailabilityDict: AvailabilityDict // Pools for each availability date ex "2/25"

  // Constraints:
  MaxSpanishUsers: number
  MaxDefaultUsers: number // Note Spanish pool is not affected by constraints except for overall Max
  MaxConstraints: number[] // Max of MS1, MS2, MS3, MS4 by const indices
  SharedMaxIndices: number[][] // Sets that should share the same bool 
  // ex MS1 & MS2 shared => [[0,1][] 
}

const MS0 = 0;
const MS1 = 1;
const MS2 = 2;
const MS3 = 3;
const MS4 = 4;

interface AvailabilityDict {
  [key: string]: AvailabilityDate;
}

interface AvailabilityDate {
  DateID: string // Simple month / date string
  DefaultUserPool: User[]
  SpanishTranslatorPool: User[] // Spanish translators are a separate pool with their own cap

  DefaultUserAssignments: Set<User>
  SpanishTranslatorAssignments: Set<User>

  NumDefaultByYr: number[] // number of MS0, MS1, MS2, MS3, MS4 by const indices, used for assignment algorithm
}

const SHORT_NAME: number = 0;
const LONG_NAME: number = 1;
const CLINICS: [string, string][] = [
  ["MD", "Agape MD Clinic"],
  ["Dermatology", "Agape Dermatology Clinic"],
  ["Smoking", "UGM Smoking Cessation Clinic"],
  ["Shelter", "UGM Shelter Clinic"],
  ["General", "BBHH General Clinic"],
  ["Women", "BBHH Women's Clinic"],
  ["Monday", "The Monday Clinic"],
  ["Diabetes","BBHH Diabetes Clinic"]
];
//#endregion

//#region Parsing Functions

/**
 * Parse a str cell input with a list of clinics and get an array of indices corresponding to CLINICS indexing
 */
function clinicMask(str: string): number[] {
  let ret: number[] = [];
  for (let i = 0; i < CLINICS.length; i++) {
    if (str.toLowerCase().includes(CLINICS[i][SHORT_NAME].toLowerCase())) {
      ret.push(i);
    }
  }
  return ret;
}

function translatorType(str: string): UserType {
  // if(str.toLowerCase().includes('but')) return UserType.SpanishVolunteer; // "I would like to be a general volunteer but can also speak Spanish."
  return str.toLowerCase().includes('y') ? UserType.SpanishTranslator : UserType.Default;
}

/**
 * Get list of clinic short names by rank.
 */
function getClinicRanks(str: string) {
  let ranks: string[] = [];
  let names_long = str.split(';').filter((x) => x !== undefined && x.length > 0);
  for (let name of names_long) {
    let i = CLINICS.findIndex((c) => name.toLowerCase().includes(c[SHORT_NAME].toLowerCase()));
    ranks.push(CLINICS[i][SHORT_NAME])
  }

  if (names_long.length !== ranks.length) {
    throw new Error("Unexpected ranking format in column.");
  }

  return ranks;
}


/**
 * Parse variables from a "Prompt" workbook / tab in the excel doc
 * A2 : The start time to query, no user rows will be included before this time
 */
function validateAndProcessPrompt(promptSheet: ExcelScript.Worksheet) {
  let values = promptSheet.getUsedRange().getValues();
  let valid = promptSheet.getRange("A20").getValue() === "Start time";

  let v = promptSheet.getRange("A21").getValue() as number;
  let promptStartDate: Date = new Date(Math.round((v - 25569) * 86400 * 1000));
  valid = valid && !isNaN(promptStartDate.getTime());

  console.log("Prompt: using dates only after " + promptStartDate.toDateString() + ' ' + promptStartDate.toTimeString());

  return { promptStartDate, valid }
}

function validatePromptCells(promptSheet: ExcelScript.Worksheet){
    // General prompt validation:
    let valid = promptSheet.getRange("A1").getValue() === "Clinic" 
      && promptSheet.getRange("A2").getValue() === "Max General Volunteers"
      && promptSheet.getRange("A3").getValue() === "Max Spanish Translators"
      && promptSheet.getRange("A4").getValue() === "Max MS0"
      && promptSheet.getRange("A5").getValue() === "Max MS1"
      && promptSheet.getRange("A6").getValue() === "Max MS2"
      && promptSheet.getRange("A7").getValue() === "Max MS3"
      && promptSheet.getRange("A8").getValue() === "Max MS4"
      && promptSheet.getRange("A9").getValue() === "Group 1"
      && promptSheet.getRange("A10").getValue() === "Group 2"
      && promptSheet.getRange("A11").getValue() === "Group 3"
      && promptSheet.getRange("A12").getValue() === "Group 4"
      && promptSheet.getRange("A13").getValue() === "Group 5";

    if (!valid) {
      throw new Error("Unexpected prompt ordering.");
    }
}

/**
 * Get constraints per clinic, s.t. when assigning users from a pool, we will not allow more assignments than the specified maximums.
 * For default users (non-translators), groups can have sub-constraints based on maximums.
 * NOTE 1: If clinic.MAX is not met because, for example, there aren't enough MS4's available. Then we will take above the max from other years in **an equal distribution** if possible.
 * Sub groups can have constraints when determining distributions, for example, group MS1s & MS2s and group MS3s and MS4s
 * NOTE 2: If MS1 max is 2 and MS2 max is 2, and the pool has no MS1's then priority will be given to MS2's given that MS3 and MS4 are already filled
 */
function getPromptConstraints(clinicIdx: number, promptSheet: ExcelScript.Worksheet) {
  let longName = CLINICS[clinicIdx][LONG_NAME];
  let shortName = CLINICS[clinicIdx][SHORT_NAME];
  let values = promptSheet.getUsedRange().getValues();

  // Each column corresponds to variables for a clinic, first col is headers
  let clinicCol: number = values[0].findIndex((clinicName) => clinicName.toString().toLowerCase().includes(shortName.toLowerCase()));
  if (values[0][clinicCol] !== longName) {
    throw new Error("Clinic name does not match header in prompt, expected: " + longName);
  }

  // TODO: add row validation of mapping
  let maxDefaultUsers = values[1][clinicCol] as number; // B2
  let maxSpanUsers = values[2][clinicCol] as number; // B3
  let maxMS0 = values[3][clinicCol] as number; // B4
  let maxMS1 = values[4][clinicCol] as number; // B5
  let maxMS2 = values[5][clinicCol] as number; // B6
  let maxMS3 = values[6][clinicCol] as number; // B7
  let maxMS4 = values[7][clinicCol] as number; // B8
  let maxConstraints = [maxMS0, maxMS1, maxMS2, maxMS3, maxMS4]; // Max for [MS1, MS2, MS3, MS4]

  // Get groups
  let groups: number[][] = []; // Check each cell for the MS1-4 numbers
  let foundIndices: number[] = [];
  let offset = 8; // B8 - this is an offset for r following:
  let NUM_GROUPS = 5;
  for (let r = 0; r < NUM_GROUPS; ++r) { // r corresponds to row and also year index ex: MS1 = 0
    let groupCell = values[r + offset][clinicCol].toString().toLowerCase();
    let subgroup: number[] = [];
    for (let y = 1; y <= 4; ++y) { // y is year number (1-indexed)
      if (groupCell.includes(y.toString())) {
        subgroup.push(y - 1); // convert to year index (ex: MS1 = 0)
        if (foundIndices.includes(y - 1)) {
          throw new Error("Column " + longName + " group constraints arenot valid.");
        }
        foundIndices.push(y - 1);
      }
    }
    if (subgroup.length > 1) {
      groups.push(subgroup);
    }
  }

  return { maxDefaultUsers, maxSpanUsers, maxConstraints, groups };
}

/**
 * Filter range based on a prompted start time (inclusive)
 * @param startTimeIdx the index of the column "Start time" to filter rows from
 */
function filterRows(range: ExcelScript.Range, promptStartDate: Date, startTimeIdx: number): (number | boolean | string)[][] {
  const rowCount = range.getRowCount();
  let ret: (number | boolean | string)[][] = [];
  let values = range.getValues();

  ret.push(values[0]); // Push header
  for (let r = 1; r < rowCount; r++) {
    let cellDateRaw = values[r][startTimeIdx] as number;
    let cellDate: Date = new Date(Math.round((cellDateRaw - 25569) * 86400 * 1000));
    if (cellDate >= promptStartDate) ret.push(values[r]); // Keep rows with date after promptStartDate
  }

  return ret;
}

/**
 * Convert from field "Undergraduate" or "Undergrad" to MS0 
 **/
function interpretYear(yr : string){
  if(yr.toLowerCase().includes("under")){ return "MS0"; }
  return yr;
}
//#endregion

const PROMPT_TAB = "Prompt"; // Name of tab with prompt / filter variables
const MAX_DEFAULT = 10; // Max users that can fit in a default user list
const MAX_TRANSLATOR_SPAN = 2; // Max spanish translator users needed for a clinic (separate pool from default)

// Process workbook
/**
 * Precondition: Form data in first tab "Form1", Prompt data in second tab "Prompt"
 */
function main(workbook: ExcelScript.Workbook) {
  //#region Variables
  const selectedSheet = workbook.getFirstWorksheet(); // IMPORTANT: Left-most tab is where form data is contained
  const range = selectedSheet.getUsedRange();
  const headerValues = range.getValues()[0];
  let values = range.getValues();

  // Headers:
  const nameIdx = headerValues.indexOf("First and Last Name");
  const yearIdx = headerValues.indexOf("What year are you in?");
  const clinicsOfInterestIdx = headerValues.indexOf("Please select any clinics you are interested in volunteering with");
  const translatorIdx = headerValues.indexOf("Are you interested in being a translator (Spanish) instead of a general volunteer?")
  const startTimeIdx = headerValues.indexOf("Start time");
  const rankIdx = headerValues.indexOf("Please rank your clinic preference");
  const electiveIdx = headerValues.indexOf("Are you currently in the fourth-year elective?");

  if(nameIdx < 0 || yearIdx < 0 || clinicsOfInterestIdx < 0 || translatorIdx < 0 || startTimeIdx < 0 || rankIdx < 0 || electiveIdx < 0){
    throw new Error("Invalid Headers - check that names have not been changed");
  }

  // FILTER values based on prompt start time (inclusive)
  let promptSheet = workbook.getWorksheet(PROMPT_TAB);
  validatePromptCells(promptSheet); // Are fields in the prompt tab where we expect them?
  const { promptStartDate, valid } = validateAndProcessPrompt(promptSheet)
  if (!valid) {
    throw new Error("Invalid prompts - Date was not found.");
  }
  values = filterRows(range, promptStartDate, startTimeIdx)// Update values with 'Start time' from prompt start range
  //#endregion

  //#region Process Rows
  // Create user objects ---------------------------
  let users: User[] = [];
  for (let r = 1; r < values.length/*numRows*/; r++) {
    let user: User = {
      Row: r,
      Name: values[r][nameIdx].toString(),
      MedYear: interpretYear(values[r][yearIdx].toString()),
      Elective : values[r][electiveIdx].toString().toLowerCase().includes('y'),
      ClinicsOfInterest: clinicMask(values[r][clinicsOfInterestIdx].toString()), // An array of clinic indices
      Type: translatorType(values[r][translatorIdx].toString()),
      NumAssignments: 0,
      DateIDsAssigned: new Set<string>(),
      ClinicRanks: getClinicRanks(values[r][rankIdx].toString())
    };
    users.push(user);
  }

  console.log("Found " + users.length + " users.");

  //  Assign to clinics ---------------------------
  // First get pools of users per clinic (w/o applying max constraints or duplicate constraints)
  let assignments = assignClinicPools(users, values, promptSheet);
  rankAndChooseUsers(users, assignments, values);
  console.log(assignments);

  //#endregion

  let resultSheet = workbook.getWorksheet("Results") || workbook.addWorksheet("Results");
  if (resultSheet.getUsedRange() !== undefined) resultSheet.getUsedRange().clear();
  var results: string[][] = [];//getExcelResults(assignments);

  let resultHeader: string[] = ["Clinic", "Date", "Volunteers", "Translators"];
  results.push(resultHeader);

  assignments.forEach((clinic) => {
    Object.values(clinic.AvailabilityDict).forEach((date) => {
      let row: string[] = [];
      row.push(clinic.Name);
      row.push(date.DateID);
      row.push(Object.values(Array.from(date.DefaultUserAssignments).map(u => u.Name)).join(","));
      row.push(Object.values(Array.from(date.SpanishTranslatorAssignments).map(u => u.Name)).join(","));
      results.push(row);
    })
    results.push([]);
  })
  results = fillRaggedArrays(results);
  console.log(results);

  resultSheet.getRangeByIndexes(0, 0, results.length, results[0].length).setValues(results);

  //console.log(users)
  //logPools(assignments);
  //console.log(range.getValues())
}

function fillRaggedArrays(arr: string[][]): string[][] {
  // Get the length of the longest ragged array from the input:
  let maxLength = arr.reduce((max, curr) => Math.max(max, curr.length), 0);
  // Assign to a new array
  let newArr: string[][] = Array.from({ length: arr.length }, () => Array.from({ length: maxLength }));
  for (let i = 0; i < arr.length; i++) {
    for (let j = 0; j < maxLength; j++) {
      if (j < arr[i].length) {
        newArr[i][j] = arr[i][j];
      } else {
        newArr[i][j] = ""; // undefined
      }
    }
  }

  return newArr;
}

const hasDuplicates = (array: string[]): boolean => {
  return new Set(array).size !== array.length;
};

const dateIDRegex = new RegExp("(1[0-2]|0?[1-9])\/(3[01]|[12][0-9]|0?[1-9])"); // Get match [1-12]/[1-31] including 01/02

function assignClinicPools(users: User[], values: (string | number | boolean)[][], promptSheet: ExcelScript.Worksheet): ClinicAssignment[] {
  // Get all date-specific columns:
  const headerValues = values[0];
  const allHeaderDateCols = headerValues.map((item, index) =>
    (item.toString().toLowerCase().includes("date") && item.toString().toLowerCase().includes("availability"))
      ? index : null
  ).filter(index => index !== null); // All availability date columns
  const allHeaderDateVals = headerValues.filter((_, i) => allHeaderDateCols.includes(i)); // All availability date names
  function getAvailabilityDates(clinicIdx: number): { dates: AvailabilityDict, headerDateCol: number } {

    //Get the column index of the "Select Date Availability" for this clinic at given index
    const shortClinicName = CLINICS[clinicIdx][SHORT_NAME];
    const headerDateCols = headerValues.map((header, idx) =>
      (header.toString().toLowerCase().includes("date") && header.toString().toLowerCase().includes("availability") && header.toString().toLowerCase().includes(shortClinicName.toLowerCase()))
        ? idx : null
    ).filter(index => index !== null); //Aavailability date columns
    if (headerDateCols.length > 1) {
      throw new Error("More than one column for clinic present.");
    }
    else if(headerDateCols.length ==0){
      throw new Error("No columns for clinic present.");
    }
    const headerDateCol: number = headerDateCols[0] as number;

    let dates: AvailabilityDict = {};
    let allDateIDs = values.map(row => row[headerDateCol]).slice(1);
    const flattenedArr: string[] = allDateIDs.reduce((acc, innerArr) => acc.concat(innerArr), []).map(x => x.toString());
    let uniqueDateIDStringsRaw = Array.from(new Set(flattenedArr));
    const uniqueDateIDs = Array.from(new Set(uniqueDateIDStringsRaw.join(';').split(';').map(s => s.trim())));
    //const finalIDs = uniqueDateIDs.filter((str) => dateIDRegex.test(str)).map(x => x.match(dateIDRegex)[0]); // Prev: 3/14
    const finalIDs = uniqueDateIDs.filter((str) => dateIDRegex.test(str)).map(x => x.trim()); // Filter based on regex

    if(finalIDs.length ===0){
      throw new Error("Clinic dates are not in a valid format");
    }
    else if(hasDuplicates(finalIDs)){
      throw new Error("Clinic "+ CLINICS[clinicIdx][LONG_NAME]+" has duplicate date IDs."); // Likely error with time regex
    }
    //console.log(finalIDs);

    finalIDs.forEach((id) => {
      let availability: AvailabilityDate = {
        DateID: id,
        DefaultUserPool: [],
        SpanishTranslatorPool: [],
        DefaultUserAssignments: new Set<User>(),
        SpanishTranslatorAssignments: new Set<User>(),
        NumDefaultByYr: [0, 0, 0, 0, 0] // MS0,1,2,3,4
      };
      dates[id] = availability;
    });

    return { dates, headerDateCol };
  }

  // Initialize clinic assignment objects:
  let clinics: ClinicAssignment[] = [];
  for (let c = 0; c < CLINICS.length; ++c) {
    let { maxDefaultUsers, maxSpanUsers, maxConstraints, groups } = getPromptConstraints(c, promptSheet);

    let { dates, headerDateCol } = getAvailabilityDates(c);
    let clinic: ClinicAssignment = {
      Name: CLINICS[c][LONG_NAME],
      ClinicIndex: c,
      DateAvailColIndex: headerDateCol,
      AvailabilityDict: dates,
      MaxSpanishUsers: maxSpanUsers,
      MaxDefaultUsers: maxDefaultUsers,
      MaxConstraints: maxConstraints,
      SharedMaxIndices: groups
    };
    clinics.push(clinic);
  }

  /**
   * Is the given user valid at a provided clinic?
   * Precondition: this clinic is already one of the users preferences.
   */
  function userIsValid(_user: User, clinicIdx: number): [boolean, string[]] {
    let clinic = clinics[clinicIdx];
    let availabilityStr = values[_user.Row][clinic.DateAvailColIndex].toString();

    // "No Dates Available / Not Interested;"
    if (availabilityStr.toLowerCase().includes('no')) { return [false, []]; }
    // Prev approach: 3/14 as an id
    //const ids = availabilityStr.split(';').filter((str) => dateIDRegex.test(str)).map(x => x.match(dateIDRegex)[0]);
    // Get full string as id:
    const ids = availabilityStr.split(';').filter((str) => dateIDRegex.test(str)).map(x => x.trim());

    return [true, ids];
  }

  // Assign users to each clinic pool
  // Note this does not prevent duplicate assignments across clinics on the same day and needs to be adjusted after filter, also does not apply prompt maxes per pool
  users.forEach(
    (user) => {
      let anyAssignment = false; // Error checking
      user.ClinicsOfInterest.forEach((clinicIdx) => {  // First filter by clinics of interest
        const [valid, ids] = userIsValid(user, clinicIdx);
        if (valid) {
          anyAssignment = true; // was user assigned to any pool?
          //clinics[clinicIdx].DefaultUserPool.push(user);
          for (let dateID of ids) {
            if (user.Type == UserType.Default) {
              clinics[clinicIdx].AvailabilityDict[dateID].DefaultUserPool.push(user);
            }
            else if (user.Type == UserType.SpanishTranslator) {
              clinics[clinicIdx].AvailabilityDict[dateID].SpanishTranslatorPool.push(user);
            }
          }
        }
      });
      // Volunteer was not assigned to any clinic pool - either clinic names are invalid or in TESTING mode
      if(!anyAssignment && !TESTING){
        console.log("Invalid user " + user.Name + " likely chose \"Not Interested\" on all clinics.");
        //throw new Error("Invalid volunteer " + user.Name);
      }
    }
  );

  return clinics;
  // TODO: warnings if a user is not in any pool or if they have no assignemtn (after this func)
}

/** 1 (highest preference) to infinity rank given a user and query clinic */
function rank(user: User, clinicIdx: number) {
  let i = user.ClinicRanks.findIndex((r) => r.toLowerCase().includes(CLINICS[clinicIdx][SHORT_NAME].toLowerCase()));
  if (i < 0) throw new Error("Rank not found for user.");
  return i + 1; // Convert to 1-indexed
}

/**
 *  Get first available user from pool by rank of preference order: undefined if none found.
 */
function firstByRank(clinic: ClinicAssignment, pool: User[], language: string, date: AvailabilityDate) {
  if (pool.length === 0) { return undefined; } // Empty pool

  // Check if date for this clinic is already at maximum capacity
  // Then filter the user pool based on what is not already in the assigned set of users to that date
  let filtered: User[] = [];
  // DEFAULT -------------
  if (language === "default") {
    // Hit max constraint
    if (date.DefaultUserAssignments.size >= clinic.MaxDefaultUsers) return undefined;
    filtered = pool.filter((x) => !x.DateIDsAssigned.has(date.DateID) && !date.DefaultUserAssignments.has(x));
    if (filtered.length === 0) { return undefined; } // Case: all assigned
  }
  // SPANISH ----------
  else if (language == "spanish") {
    // Hit max constraint
    if (date.SpanishTranslatorAssignments.size >= clinic.MaxSpanishUsers) return undefined;
    filtered = pool.filter((x) => !x.DateIDsAssigned.has(date.DateID) && !date.SpanishTranslatorAssignments.has(x));
    if (filtered.length === 0) { return undefined; } // Case: all assigned

  }
  else throw new Error("Unexpected language parsed.");

  // Filtered users (if already assigned), and ordered by rank preference, solving ties with random choice
  let ranks: { user: User, rank: number }[] = [];
  let users = filtered.sort((u) => rank(u, clinic.ClinicIndex));
  for (let u of users) {
    ranks.push({ user: u, rank: rank(u, clinic.ClinicIndex) }); // TODO optimize
  }

  if (ranks.length === 0) {
    throw new Error("Invalid ranking from user pool");
  }

  // Find first set with same rank
  let topRank = ranks[0].rank;
  let tie: User[] = [];
  for (let i = 0; i < ranks.length && ranks[i].rank == topRank; ++i) {
    tie.push(ranks[i].user);
  }

  if (tie.length === 1) {
    if(VERBOSE) console.log(clinic.Name + " " + language + ", rank " + topRank + " | " + tie[0].Name + " | " + date.DateID);
    return tie[0];
  }
  else {
    // Get tie of the users tied at the current rank:
    const randomIndex = Math.floor(Math.random() * tie.length);
    if(VERBOSE) console.log(clinic.Name + " tie, rank " + topRank + " | " + tie[randomIndex].Name + " | " + date.DateID);
    return tie[randomIndex];
  }
}

function rankAndChooseUsers(users: User[], assignments: ClinicAssignment[], values: (string | number | boolean)[][]) {


  // Get all dates to process across all clinics:
  // TODO: this assumes all the same date, and TODO: use start date year for new Date()
  let allDates = assignments.map((clinic) => Object.values(clinic.AvailabilityDict).map((date) => date.DateID)).reduce((acc, cur) => [...acc, ...cur], []);
  let uniqueDates = Array.from(new Set(allDates));
  let sortedDates = uniqueDates;
  // TODO: This was removed with the introduction of time ranges
  /*const sortedDates: string[] = uniqueDates.sort((a, b) => {
    const [aMonth, aDay] = a.split('/'), [bMonth, bDay] = b.split('/');
    return new Date(parseInt(aMonth) - 1, parseInt(aDay)).getTime() - new Date(parseInt(bMonth) - 1, parseInt(bDay)).getTime();
  });*/
  console.log("Parsing the following dates:", sortedDates);

  // VISIT each date separately for multi-assignment:
  sortedDates.forEach((dateQuery) => {
    let numAssigned = 0;
    do {
      numAssigned = 0;

      // Visit each clinic and assign one user at a time by preference & constraints
      // This prevents one clinic from consuming all of the user availability
      assignments.forEach((clinic) => {
        let date = clinic.AvailabilityDict[dateQuery];
        if (date !== undefined) { // Is this date in the clinic's list of dates for the month?
          // Default:
          let result_default = firstByRank(clinic, date.DefaultUserPool, "default", date);
          if (result_default !== undefined) {
            result_default.NumAssignments++;
            result_default.DateIDsAssigned.add(date.DateID);
            date.DefaultUserAssignments.add(result_default);
            numAssigned++;
          }
          // Spanish:
          let result_span = firstByRank(clinic, date.SpanishTranslatorPool, "spanish", date);
          if (result_span !== undefined) {
            result_span.NumAssignments++;
            result_span.DateIDsAssigned.add(date.DateID);
            date.SpanishTranslatorAssignments.add(result_span);
            numAssigned++;
          }
        }
      });
    } while (numAssigned != 0);
  });
}

//#region Utilities
function logPools(assignments: ClinicAssignment[]) {
  console.log("\n")
  assignments.forEach(
    (clinic) => {
      console.log(clinic.Name);
      for (let date in clinic.AvailabilityDict) {
        let obj = clinic.AvailabilityDict[date];
        if (obj.DefaultUserPool.length > 0) {
          console.log("Default " + obj.DateID + ": { " + obj.DefaultUserPool.map((u) => u.Name).join() + " }");
        }
        if (obj.SpanishTranslatorPool.length > 0) {
          console.log("Spanish " + obj.DateID + ": { " + obj.SpanishTranslatorPool.map((u) => u.Name).join() + "}");
        }
      }
      console.log("\n")
    });
}
//#endregion