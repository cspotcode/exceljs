const fs = require('fs');
const os = require('os');
const path = require('path');
const {execFile} = require('child_process');
const {promisify} = require('util');
const {v4: uuidv4} = require('uuid');

const execFileAsync = promisify(execFile);

const EXE_PATH = path.join(__dirname, 'excel-validator', 'bin', 'excel-validator.exe');
const LOG_SETTLE_MS = 1500;

function listErrorLogs(tempDir) {
  try {
    return fs
      .readdirSync(tempDir)
      .filter(name => /^error\d+_\d+\.xml$/i.test(name))
      .map(name => path.join(tempDir, name));
  } catch {
    return [];
  }
}

/**
 * Finds the %TEMP%\error<digits>_<NN>.xml file (if any) that Excel wrote while
 * opening `fileName`, by requiring both a recent mtime (created after
 * `sinceMs`) and that the log's content names `fileName`.
 */
function findMatchingLog(tempDir, fileName, sinceMs) {
  const candidates = listErrorLogs(tempDir)
    .filter(candidatePath => {
      try {
        return fs.statSync(candidatePath).mtimeMs >= sinceMs;
      } catch {
        return false;
      }
    })
    .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);

  for (const candidatePath of candidates) {
    let xml;
    try {
      xml = fs.readFileSync(candidatePath, 'utf8');
    } catch {
      continue;
    }
    if (xml.toLowerCase().includes(fileName.toLowerCase())) {
      return {path: candidatePath, xml};
    }
  }
  return null;
}

/**
 * Drives excel-validator.exe to check whether `buffer` (xlsx bytes) opens cleanly in
 * Excel Desktop.
 *
 * Returns `{valid, successfulMode, errors, log}`:
 * - `valid` - true only if Excel opened the file on the normal-load attempt.
 * - `successfulMode` - which CorruptLoad mode Excel required to open the file:
 *   'normal' | 'repair' | 'extractData' | null (null when every mode failed).
 * - `errors` - `{normal, repair, extractData}`, the COM error message for every mode
 *   that was attempted and failed
 * - `log` - matched repair-log file, only populated when
 *   `captureLog: true` and a matching log is found
 */
async function checkExcelValidity(buffer, {baseName = 'input.xlsx', captureLog = false} = {}) {
  const ext = path.extname(baseName) || '.xlsx';
  const stem = path.basename(baseName, ext);
  const tempName = `${stem}-${uuidv4()}${ext}`;
  const tempPath = path.join(os.tmpdir(), tempName);

  fs.writeFileSync(tempPath, buffer);

  try {
    const startMs = Date.now();

    const {stdout} = await execFileAsync(EXE_PATH, [tempPath]);
    const {successfulMode, errors} = JSON.parse(stdout);

    let log = null;
    if (captureLog) {
      await new Promise(resolve => setTimeout(resolve, LOG_SETTLE_MS));
      log = findMatchingLog(os.tmpdir(), tempName, startMs);
    }

    return {valid: successfulMode === 'normal', successfulMode, errors, log};
  } finally {
    try {
      fs.unlinkSync(tempPath);
    } catch {
      // best-effort cleanup
    }
  }
}

module.exports = {checkExcelValidity};
