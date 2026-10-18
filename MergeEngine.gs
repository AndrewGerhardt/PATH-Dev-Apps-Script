function generatePdfForMergedRow_(run, merged) {
  const scout = merged.scout || {};
  const isSecond = merged.targetBadgeSlot === 2;

  const badgeName = isSecond
    ? (merged.badge2?.BadgeName || 'Badge')
    : (merged.badge1?.BadgeName || 'Badge');

  const unitNumber = String(scout.UnitNumber || '').trim();
  const unitTypeBG = String(scout.UnitType || '').trim();
  const unitState = stateToAbbrev_(scout.UnitState);

  const unitFolderName = [unitNumber + unitTypeBG, unitState]
    .filter(Boolean)
    .join(' ');

  const folder = getOrCreateUnitFolderCached_(
    String(CFG.OUTPUT_FOLDER_ID).trim(),
    unitFolderName
  );

  const last = String(scout.LastName || 'Last').trim();
  const first = String(scout.FirstName || 'First').trim();

  const safe = (s) => String(s || '').replace(/[\\/:*?"<>|]/g, '').trim();

  const filename =
    `${safe(unitNumber)}${safe(unitTypeBG)} ${safe(last)}, ${safe(first)} - ` +
    `${safe(badgeName)} - S${safe(run.Session)} ${safe(run.Year)}.pdf`;

  const copy = Drive.Files.copy(
    {
      title: filename.replace(/\.pdf$/i, ''),
      parents: [{ id: folder.getId() }]
    },
    CFG.TEMPLATE_DOC_ID,
    { supportsAllDrives: true }
  );

  const tempDocId = copy.id;

  const tagMap = buildTemplateTagMap_(merged);

  console.log('PDF filename:', filename);
  console.log('Temp Doc ID:', tempDocId);

  replaceDocTags_(tempDocId, tagMap);

  // Give Google Docs time to fully save before export.
  Utilities.sleep(2000);

  const url = `https://www.googleapis.com/drive/v2/files/${tempDocId}/export?mimeType=${encodeURIComponent(MimeType.PDF)}`;

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: false
  });

  const pdfBlob = response.getBlob().setName(filename);

  const pdfFile = DriveApp.getFolderById(folder.getId())
    .createFile(pdfBlob);

  DriveApp.getFileById(tempDocId).setTrashed(true);

  return {
    id: pdfFile.getId(),
    url: pdfFile.getUrl(),
    name: pdfFile.getName()
  };
}

function replaceDocTags_(docId, tagMap) {
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  for (const [tag, value] of Object.entries(tagMap)) {
    body.replaceText(
      escapeForDocReplace_(tag),
      String(value ?? '')
    );
  }

  doc.saveAndClose();

  // Force a short pause after save before anything tries to export.
  Utilities.sleep(2000);
}

function buildTemplateTagMap_(merged) {
  const roster = merged.roster || {};
  const scout = merged.scout || {};
  const offering = merged.offering || {};

  const isSecond = merged.targetBadgeSlot === 2;

  const badgeName = isSecond
    ? (merged.badge2?.BadgeName || '')
    : (merged.badge1?.BadgeName || '');

  const earned = isSecond
    ? String(roster.Earned_2 || '').trim()
    : String(roster.Earned_1 || '').trim();

  const period = String(offering.Period || '').trim();

  const offeringId = String(
    roster.OfferingID || 
    offering.OfferingID || 
    ''
  ).trim();

  const isCombo = offeringId.includes('+');;

  const reqComplete = formatReqListForPdfExport_(
    roster.ReqComplete,
    badgeName,
    isCombo
  );

  const reqOptional = formatReqListForPdfExport_(
    roster.ReqOptional,
    badgeName,
    isCombo
  );

  const reqIncomplete = formatReqListForPdfExport_(
    roster.ReqIncomplete,
    badgeName,
    isCombo
  );

  return {
    '<<BadgeName>>': badgeName,
    '<<Period>>': period,

    '<<FirstName>>': String(scout.FirstName || ''),
    '<<LastName>>': String(scout.LastName || ''),
    '<<UnitNumber>>': String(scout.UnitNumber || ''),
    '<<UnitType>>': String(scout.UnitType || ''),

    '<<M>>': String(roster.M || ''),
    '<<T>>': String(roster.T || ''),
    '<<W>>': String(roster.W || ''),
    '<<TH>>': String(roster.TH || ''),
    '<<F>>': String(roster.F || ''),

    '<<Earned>>': earned,

    '<<ReqComplete>>': reqComplete,
    '<<ReqOptional>>': reqOptional,
    '<<ReqIncomplete>>': reqIncomplete,

    '<<Comments>>': String(roster.Comments || '')
  };
}

function formatReqListForPdfExport_(value, badgeName, isCombo) {
  const original = String(value || '').trim();
  if (!original) return '';

  const parts = original
    .split(/\s*,\s*/)
    .map(x => x.trim())
    .filter(Boolean);

  if (!isCombo) {
    return parts.join(', ');
  }

  const badge = String(badgeName || '').trim();
  if (!badge) return parts.join(', ');

  const prefixRegex = new RegExp(
    '^' + escapeRegex_(badge) + '\\s*[-–—:]\\s*',
    'i'
  );

  const matchingParts = parts
    .filter(x => prefixRegex.test(x))
    .map(x => x.replace(prefixRegex, '').trim())
    .filter(Boolean);

  return matchingParts.length
    ? matchingParts.join(', ')
    : parts.join(', ');
}

function stateToAbbrev_(state) {
  const s = String(state || '').trim();
  if (!s) return '';

  const upper = s.toUpperCase();
  if (/^[A-Z]{2}$/.test(upper)) return upper;

  const map = {
    'ALABAMA': 'AL',
    'ALASKA': 'AK',
    'ARIZONA': 'AZ',
    'ARKANSAS': 'AR',
    'CALIFORNIA': 'CA',
    'COLORADO': 'CO',
    'CONNECTICUT': 'CT',
    'DELAWARE': 'DE',
    'FLORIDA': 'FL',
    'GEORGIA': 'GA',
    'HAWAII': 'HI',
    'IDAHO': 'ID',
    'ILLINOIS': 'IL',
    'INDIANA': 'IN',
    'IOWA': 'IA',
    'KANSAS': 'KS',
    'KENTUCKY': 'KY',
    'LOUISIANA': 'LA',
    'MAINE': 'ME',
    'MARYLAND': 'MD',
    'MASSACHUSETTS': 'MA',
    'MICHIGAN': 'MI',
    'MINNESOTA': 'MN',
    'MISSISSIPPI': 'MS',
    'MISSOURI': 'MO',
    'MONTANA': 'MT',
    'NEBRASKA': 'NE',
    'NEVADA': 'NV',
    'NEW HAMPSHIRE': 'NH',
    'NEW JERSEY': 'NJ',
    'NEW MEXICO': 'NM',
    'NEW YORK': 'NY',
    'NORTH CAROLINA': 'NC',
    'NORTH DAKOTA': 'ND',
    'OHIO': 'OH',
    'OKLAHOMA': 'OK',
    'OREGON': 'OR',
    'PENNSYLVANIA': 'PA',
    'RHODE ISLAND': 'RI',
    'SOUTH CAROLINA': 'SC',
    'SOUTH DAKOTA': 'SD',
    'TENNESSEE': 'TN',
    'TEXAS': 'TX',
    'UTAH': 'UT',
    'VERMONT': 'VT',
    'VIRGINIA': 'VA',
    'WASHINGTON': 'WA',
    'WEST VIRGINIA': 'WV',
    'WISCONSIN': 'WI',
    'WYOMING': 'WY'
  };

  return map[upper] || s;
}

function escapeRegex_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeForDocReplace_(literal) {
  return String(literal).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function getOrCreateUnitFolder_(rootFolderId, unitFolderName) {
  const root = DriveApp.getFolderById(String(rootFolderId).trim());

  const it = root.getFoldersByName(unitFolderName);
  if (it.hasNext()) return it.next();

  return root.createFolder(unitFolderName);
}

const _unitFolderCache = new Map();

function getOrCreateUnitFolderCached_(rootFolderId, unitFolderName) {
  const key = `${rootFolderId}::${unitFolderName}`;
  if (_unitFolderCache.has(key)) return _unitFolderCache.get(key);

  const folder = getOrCreateUnitFolder_(rootFolderId, unitFolderName);
  _unitFolderCache.set(key, folder);
  return folder;
}

function generatePdfsForNextRunOnly() {
  const run = findNextPendingOrRunningRun_();
  if (!run) return;

  processRunChunkParallel_(0);
}