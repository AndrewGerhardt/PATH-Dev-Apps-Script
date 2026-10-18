function buildMergedRowData_(rosterObj, scoutMap, badgeMap, offeringMap) {
  const scoutId = String(rosterObj[CFG.COLS.ROSTER_SCOUTID] || '').trim();
  const offeringId = String(rosterObj[CFG.COLS.ROSTER_OFFERINGID] || '').trim();

  const scout = scoutMap.get(scoutId) || null;
  const offering = offeringMap.get(offeringId) || null;

  const badge1 = offering ? badgeMap.get(String(offering.BadgeID_1 || '').trim()) : null;
  const badge2 = offering ? badgeMap.get(String(offering.BadgeID_2 || '').trim()) : null;

  return {
    runYear: rosterObj[CFG.COLS.ROSTER_YEAR],
    runSession: rosterObj[CFG.COLS.ROSTER_SESSION],
    roster: rosterObj,
    scout,
    offering,
    badge1,
    badge2,
    ScoutName: scout ? `${scout.FirstName} ${scout.LastName}` : '',
    Unit: scout ? `${scout.UnitType} ${scout.UnitNumber}` : '',
    OfferingID: offeringId || '',
    BadgeName_1: badge1 ? badge1.BadgeName : '',
    BadgeName_2: badge2 ? badge2.BadgeName : '',
  };
}

// --- Lookups

function getScoutById_(ss, scoutId) {
  if (!scoutId) return null;
  return lookupByKey_(ss.getSheetByName(CFG.SHEETS.SCOUTS), CFG.COLS.SCOUTID, scoutId);
}

function getOfferingById_(ss, offeringId) {
  if (!offeringId) return null;
  return lookupByKey_(ss.getSheetByName(CFG.SHEETS.OFFERINGS), CFG.COLS.OFFERINGID, offeringId);
}

function getBadgeById_(ss, badgeId) {
  if (!badgeId) return null;
  return lookupByKey_(ss.getSheetByName(CFG.SHEETS.BADGES), CFG.COLS.BADGEID, badgeId);
}

function lookupByKey_(sheet, keyColName, keyVal) {
  const { headers, idx } = readHeaders_(sheet);
  const values = sheet.getDataRange().getValues();

  const keyIdx = idx[keyColName];
  if (keyIdx === undefined) throw new Error(`${sheet.getName()} missing key column ${keyColName}`);

  for (let r = CFG.HEADER_ROW; r < values.length; r++) {
    const row = values[r];
    if (String(row[keyIdx]).trim() === String(keyVal).trim()) {
      return rowToObject_(headers, row);
    }
  }
  return null;
}
