/**
 * Austin Lunch Club shared backend
 *
 * Setup:
 * 1. Create a Google Sheet and copy its ID from the URL.
 * 2. Paste the ID below.
 * 3. In that Sheet, open Extensions > Apps Script and replace Code.gs with this file.
 * 4. Deploy > New deployment > Web app.
 * 5. Execute as: Me. Who has access: Anyone.
 * 6. Copy the /exec URL into the Lunch Club website's "Connect shared data" dialog.
 * 7. Run importRestaurantsFromExistingSheet once to convert the original
 *    category-column layout into the structured Restaurants tab.
 * 8. Optional ratings: add GOOGLE_PLACES_API_KEY in Script Properties, run
 *    refreshGoogleRatings once, then run installRatingRefreshTrigger once.
 */

const SPREADSHEET_ID = "1n71XTZ6j4A4C_ahP0VTHMGrjEOOL1gPxl5i_F4tZ2-s";
const RESTAURANTS_SHEET = "Restaurants";
const LOG_SHEET = "Lunch Log";
const TEAMMATES_SHEET = "Teammates";
const RESTAURANT_HEADERS = [
  "ID", "Name", "Category", "Address", "Updated At", "Google Rating",
  "Google Rating Count", "Google Place ID", "Google Maps URL", "Rating Updated At", "Rating Status",
  "Google Address", "Google Latitude", "Google Longitude"
];
const LOG_HEADERS = ["ID", "Date", "Restaurant ID", "Restaurant Name", "Winning Picker", "Notes", "Created At", "Personal Rating"];
const TEAMMATE_HEADERS = ["ID", "Name", "Created At"];
const GOOGLE_API_KEY_PROPERTY = "GOOGLE_PLACES_API_KEY";
const RATING_REFRESH_DAYS = 21;
const MAX_RATING_REFRESH_PER_RUN = 100;
const AUSTIN_CENTER = { latitude: 30.2672, longitude: -97.7431 };

function doGet(event) {
  const lock = LockService.getScriptLock();
  try {
    if (event && event.parameter && event.parameter.payload) {
      lock.waitLock(30000);
      const request = JSON.parse(event.parameter.payload);
      const result = handleAction(request.action, request.data || {});
      return getResponse({ ok: true, result: result }, event);
    }
    return getResponse({
      ok: true,
      service: "Austin Lunch Club",
      restaurants: readRestaurants(),
      logs: readLogs(),
      teammates: readTeammates()
    }, event);
  } catch (error) {
    return getResponse({
      ok: false,
      error: error && error.message ? error.message : String(error)
    }, event);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function doPost(event) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const request = parseRequest(event);
    const result = handleAction(request.action, request.data || {});
    return jsonResponse({ ok: true, result: result });
  } catch (error) {
    return errorResponse(error);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function handleAction(action, data) {
  switch (action) {
    case "seedRestaurants":
      return seedRestaurants(data.restaurants || []);
    case "seedLogs":
      return seedLogs(data.logs || []);
    case "seedTeammates":
      return seedTeammates(data.teammates || []);
    case "upsertRestaurant":
      return upsertRestaurant(data.restaurant);
    case "deleteRestaurant":
      return deleteById(RESTAURANTS_SHEET, RESTAURANT_HEADERS, data.id);
    case "upsertLog":
      return upsertLog(data.log);
    case "deleteLog":
      return deleteById(LOG_SHEET, LOG_HEADERS, data.id);
    case "upsertTeammate":
      return upsertTeammate(data.teammate);
    default:
      throw new Error("Unsupported action: " + action);
  }
}

function parseRequest(event) {
  const raw = event && event.parameter && event.parameter.payload
    ? event.parameter.payload
    : event && event.postData
      ? event.postData.contents
      : "";
  if (!raw) throw new Error("Missing request payload.");
  return JSON.parse(raw);
}

function spreadsheet() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === "PASTE_GOOGLE_SHEET_ID_HERE") {
    throw new Error("Set SPREADSHEET_ID at the top of Code.gs before deploying.");
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(name, headers) {
  const book = spreadsheet();
  let sheet = book.getSheetByName(name);
  if (!sheet) sheet = book.insertSheet(name);
  const isNew = sheet.getLastRow() === 0;
  const currentHeaders = isNew ? [] : sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  if (isNew || currentHeaders.join("|") !== headers.join("|")) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  if (isNew) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#d8c8ee")
      .setFontColor("#3b2a24");
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function readRestaurants() {
  const rows = readRows(RESTAURANTS_SHEET, RESTAURANT_HEADERS);
  return rows.map(function(row) {
    return {
      id: row[0],
      name: row[1],
      category: row[2],
      address: row[3],
      rating: normalizeRating(row[5]),
      ratingCount: Number(row[6]) || 0,
      googlePlaceId: row[7] || "",
      googleMapsUrl: row[8] || "",
      ratingUpdatedAt: row[9] || "",
      ratingStatus: row[10] || "",
      googleAddress: row[11] || "",
      latitude: row[12] === "" ? null : Number(row[12]),
      longitude: row[13] === "" ? null : Number(row[13]),
      seeded: String(row[0]).indexOf("seed-") === 0
    };
  }).filter(function(item) {
    return item.id && item.name;
  });
}

function readLogs() {
  const rows = readRows(LOG_SHEET, LOG_HEADERS);
  return rows.map(function(row) {
    return {
      id: row[0],
      date: row[1],
      restaurantId: row[2],
      restaurantName: row[3],
      person: row[4],
      notes: row[5],
      createdAt: Number(row[6]) || Date.parse(row[6]) || Date.now(),
      rating: normalizeRating(row[7])
    };
  }).filter(function(item) {
    return item.id && item.restaurantName;
  });
}

function readTeammates() {
  const rows = readRows(TEAMMATES_SHEET, TEAMMATE_HEADERS);
  return rows.map(function(row) {
    return {
      id: row[0],
      name: row[1],
      createdAt: Number(row[2]) || Date.parse(row[2]) || Date.now()
    };
  }).filter(function(item) {
    return item.id && item.name;
  });
}

function readRows(name, headers) {
  const sheet = getSheet(name, headers);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
}

function seedRestaurants(restaurants) {
  if (!Array.isArray(restaurants)) throw new Error("restaurants must be an array.");
  const sheet = getSheet(RESTAURANTS_SHEET, RESTAURANT_HEADERS);
  const existing = readRestaurants();
  const ids = {};
  const names = {};
  existing.forEach(function(item) {
    ids[item.id] = true;
    names[normalize(item.name)] = true;
  });
  const rows = [];
  restaurants.slice(0, 500).forEach(function(item) {
    validateRestaurant(item);
    if (ids[item.id] || names[normalize(item.name)]) return;
    rows.push(restaurantRow(item));
    ids[item.id] = true;
    names[normalize(item.name)] = true;
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, RESTAURANT_HEADERS.length).setValues(rows);
  }
  return { added: rows.length };
}

/**
 * One-time helper for the original spreadsheet layout, where categories are
 * column headers and restaurant names are listed beneath them. Run this
 * directly from the Apps Script editor before the first rating refresh.
 */
function importRestaurantsFromExistingSheet() {
  const book = spreadsheet();
  const source = book.getSheets().find(function(sheet) {
    return sheet.getName() !== RESTAURANTS_SHEET &&
      sheet.getName() !== LOG_SHEET &&
      sheet.getName() !== TEAMMATES_SHEET;
  });
  if (!source) {
    throw new Error("No source tab was found. Keep the original restaurant-list tab in this spreadsheet.");
  }

  const values = source.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error("The source tab does not contain restaurant rows.");
  const headers = values[0];
  const restaurants = [];
  const seen = {};

  for (let column = 0; column < headers.length; column += 1) {
    const category = String(headers[column] || "Other").trim() || "Other";
    for (let row = 1; row < values.length; row += 1) {
      let name = String(values[row][column] || "").trim();
      name = name.replace(/^\d+\|\s*/, "").trim();
      if (!name || /^\d+\|?$/.test(name) || name === "-- 1 of 1 --") continue;
      const key = normalize(name);
      if (seen[key]) continue;
      seen[key] = true;
      restaurants.push({
        id: importedRestaurantId(name),
        name: name,
        category: category,
        address: "",
        rating: "",
        seeded: true
      });
    }
  }

  if (!restaurants.length) {
    throw new Error("No restaurant names were found beneath the source tab's category headers.");
  }
  const result = seedRestaurants(restaurants);
  result.sourceTab = source.getName();
  result.found = restaurants.length;
  console.log(JSON.stringify(result));
  return result;
}

function importedRestaurantId(name) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    normalize(name),
    Utilities.Charset.UTF_8
  );
  return "import-" + Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, "").slice(0, 20);
}

function seedLogs(logs) {
  if (!Array.isArray(logs)) throw new Error("logs must be an array.");
  let saved = 0;
  logs.slice(0, 1000).forEach(function(log) {
    upsertLog(log);
    saved += 1;
  });
  return { saved: saved };
}

function seedTeammates(teammates) {
  if (!Array.isArray(teammates)) throw new Error("teammates must be an array.");
  const existing = readTeammates();
  const ids = {};
  const names = {};
  existing.forEach(function(item) {
    ids[item.id] = true;
    names[normalize(item.name)] = true;
  });
  let added = 0;
  teammates.slice(0, 200).forEach(function(teammate) {
    validateTeammate(teammate);
    if (ids[teammate.id] || names[normalize(teammate.name)]) return;
    upsertTeammate(teammate);
    ids[teammate.id] = true;
    names[normalize(teammate.name)] = true;
    added += 1;
  });
  return { added: added };
}

function upsertRestaurant(restaurant) {
  validateRestaurant(restaurant);
  return upsertById(
    RESTAURANTS_SHEET,
    RESTAURANT_HEADERS,
    restaurant.id,
    restaurantRow(restaurant)
  );
}

function upsertLog(log) {
  validateLog(log);
  return upsertById(
    LOG_SHEET,
    LOG_HEADERS,
    log.id,
    [
      sheetSafe(log.id),
      sheetSafe(log.date),
      sheetSafe(log.restaurantId || ""),
      sheetSafe(log.restaurantName),
      sheetSafe(log.person),
      sheetSafe(log.notes || ""),
      String(Number(log.createdAt) || Date.now()),
      log.rating === "" || log.rating == null ? "" : String(normalizeRating(log.rating))
    ]
  );
}

function upsertTeammate(teammate) {
  validateTeammate(teammate);
  return upsertById(
    TEAMMATES_SHEET,
    TEAMMATE_HEADERS,
    teammate.id,
    [
      sheetSafe(teammate.id),
      sheetSafe(teammate.name),
      String(Number(teammate.createdAt) || Date.now())
    ]
  );
}

function upsertById(name, headers, id, row) {
  const sheet = getSheet(name, headers);
  const rowNumber = findRowById(sheet, id);
  if (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
    return { updated: true };
  }
  sheet.appendRow(row);
  return { added: true };
}

function deleteById(name, headers, id) {
  validateId(id);
  const sheet = getSheet(name, headers);
  const rowNumber = findRowById(sheet, id);
  if (!rowNumber) return { deleted: false };
  sheet.deleteRow(rowNumber);
  return { deleted: true };
}

function findRowById(sheet, id) {
  validateId(id);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const match = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(String(id))
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : 0;
}

function restaurantRow(restaurant) {
  return [
    sheetSafe(restaurant.id),
    sheetSafe(restaurant.name),
    sheetSafe(restaurant.category),
    sheetSafe(restaurant.address || ""),
    new Date().toISOString(),
    restaurant.rating === "" || restaurant.rating == null ? "" : String(Number(restaurant.rating)),
    String(Number(restaurant.ratingCount) || ""),
    sheetSafe(restaurant.googlePlaceId || ""),
    sheetSafe(restaurant.googleMapsUrl || ""),
    sheetSafe(restaurant.ratingUpdatedAt || ""),
    sheetSafe(restaurant.ratingStatus || ""),
    sheetSafe(restaurant.googleAddress || ""),
    restaurant.latitude == null || !isFinite(Number(restaurant.latitude)) ? "" : Number(restaurant.latitude),
    restaurant.longitude == null || !isFinite(Number(restaurant.longitude)) ? "" : Number(restaurant.longitude)
  ];
}

function validateRestaurant(restaurant) {
  if (!restaurant || typeof restaurant !== "object") throw new Error("Missing restaurant.");
  validateId(restaurant.id);
  validateText(restaurant.name, "Restaurant name", 100, true);
  validateText(restaurant.category, "Category", 50, true);
  validateText(restaurant.address || "", "Address", 180, false);
  if (restaurant.rating !== "" && restaurant.rating != null) {
    const rating = Number(restaurant.rating);
    if (!isFinite(rating) || rating < 1 || rating > 5) throw new Error("Google rating must be between 1 and 5.");
  }
}

function validateLog(log) {
  if (!log || typeof log !== "object") throw new Error("Missing log entry.");
  validateId(log.id);
  validateText(log.date, "Date", 10, true);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(log.date)) throw new Error("Date must use YYYY-MM-DD.");
  validateText(log.restaurantId || "", "Restaurant ID", 100, false);
  validateText(log.restaurantName, "Restaurant name", 100, true);
  validateText(log.person, "Winning picker", 50, true);
  validateText(log.notes || "", "Notes", 240, false);
  if (log.rating !== "" && log.rating != null && normalizeRating(log.rating) === "") {
    throw new Error("Rating must be between 0.5 and 5.");
  }
}

function validateTeammate(teammate) {
  if (!teammate || typeof teammate !== "object") throw new Error("Missing teammate.");
  validateId(teammate.id);
  validateText(teammate.name, "Teammate name", 50, true);
}

function validateId(id) {
  if (typeof id !== "string" || !/^[A-Za-z0-9_-]{1,100}$/.test(id)) {
    throw new Error("Invalid record ID.");
  }
}

function validateText(value, label, maxLength, required) {
  if (typeof value !== "string") throw new Error(label + " must be text.");
  if (required && !value.trim()) throw new Error(label + " is required.");
  if (value.length > maxLength) throw new Error(label + " is too long.");
}

/**
 * Run this function once from the Apps Script editor after setting the
 * GOOGLE_PLACES_API_KEY Script Property. It creates a daily trigger, but the
 * trigger only calls Google for rows whose rating is at least 21 days old.
 */
function installRatingRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "refreshGoogleRatings") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger("refreshGoogleRatings")
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();
  return "Installed daily stale-rating check.";
}

/**
 * Refreshes at most 100 stale restaurants per run. With the current list of
 * roughly 80 restaurants, a 21-day refresh interval averages about 116
 * Place Details calls per month after the initial place matching.
 */
function refreshGoogleRatings() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(GOOGLE_API_KEY_PROPERTY);
  if (!apiKey) {
    throw new Error("Add GOOGLE_PLACES_API_KEY under Apps Script Project Settings > Script properties.");
  }

  const sheet = getSheet(RESTAURANTS_SHEET, RESTAURANT_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { checked: 0, refreshed: 0, failed: 0 };

  const rows = sheet.getRange(2, 1, lastRow - 1, RESTAURANT_HEADERS.length).getDisplayValues();
  const now = Date.now();
  const staleAfterMs = RATING_REFRESH_DAYS * 24 * 60 * 60 * 1000;
  let checked = 0;
  let refreshed = 0;
  let failed = 0;

  for (let index = 0; index < rows.length && checked < MAX_RATING_REFRESH_PER_RUN; index += 1) {
    const row = rows[index];
    const name = row[1];
    if (!name) continue;
    const lastRatingUpdate = Date.parse(row[9]);
    const hasCoordinates = row[12] !== "" && row[13] !== "" &&
      isFinite(Number(row[12])) && isFinite(Number(row[13]));
    if (hasCoordinates && Number.isFinite(lastRatingUpdate) && now - lastRatingUpdate < staleAfterMs) continue;

    checked += 1;
    const sheetRow = index + 2;
    const timestamp = new Date().toISOString();
    try {
      const place = getGooglePlace(apiKey, {
        name: name,
        address: row[3],
        placeId: row[7]
      });
      if (!place) {
        sheet.getRange(sheetRow, 6, 1, 9).setValues([[
          row[5], row[6], row[7], row[8], timestamp, "No Google match",
          row[11], row[12], row[13]
        ]]);
      } else {
        const location = place.location || {};
        sheet.getRange(sheetRow, 6, 1, 9).setValues([[
          place.rating == null ? "" : place.rating,
          place.userRatingCount == null ? "" : place.userRatingCount,
          sheetSafe(place.id || row[7]),
          sheetSafe(place.googleMapsUri || row[8]),
          timestamp,
          "Google Places",
          sheetSafe(place.formattedAddress || row[11]),
          location.latitude == null ? row[12] : location.latitude,
          location.longitude == null ? row[13] : location.longitude
        ]]);
        refreshed += 1;
      }
    } catch (error) {
      failed += 1;
      const status = ("Error: " + (error.message || String(error))).slice(0, 180);
      sheet.getRange(sheetRow, 10, 1, 2).setValues([[timestamp, status]]);
      if (error.status === 401 || error.status === 403) throw error;
    }
    Utilities.sleep(75);
  }

  SpreadsheetApp.flush();
  const summary = { checked: checked, refreshed: refreshed, failed: failed };
  console.log(JSON.stringify(summary));
  return summary;
}

function getGooglePlace(apiKey, restaurant) {
  if (restaurant.placeId) {
    try {
      return fetchGooglePlaceDetails(apiKey, restaurant.placeId);
    } catch (error) {
      if (error.status !== 404) throw error;
    }
  }
  return searchGooglePlace(apiKey, restaurant.name, restaurant.address);
}

function fetchGooglePlaceDetails(apiKey, placeId) {
  const response = UrlFetchApp.fetch(
    "https://places.googleapis.com/v1/places/" + encodeURIComponent(placeId),
    {
      method: "get",
      headers: {
        "X-Goog-Api-Key": apiKey,
        "X-Goog-FieldMask": "id,formattedAddress,location,rating,userRatingCount,googleMapsUri"
      },
      muteHttpExceptions: true
    }
  );
  return parseGooglePlaceResponse(response);
}

function searchGooglePlace(apiKey, name, address) {
  const query = [name, address || "Austin, TX"].filter(Boolean).join(", ");
  const response = UrlFetchApp.fetch(
    "https://places.googleapis.com/v1/places:searchText",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        "X-Goog-Api-Key": apiKey,
        "X-Goog-FieldMask": "places.id,places.formattedAddress,places.location,places.rating,places.userRatingCount,places.googleMapsUri"
      },
      payload: JSON.stringify({
        textQuery: query,
        locationBias: {
          circle: {
            center: AUSTIN_CENTER,
            radius: 50000
          }
        }
      }),
      muteHttpExceptions: true
    }
  );
  const data = parseGooglePlaceResponse(response);
  return data && Array.isArray(data.places) ? (data.places[0] || null) : null;
}

function parseGooglePlaceResponse(response) {
  const status = response.getResponseCode();
  let data = {};
  try {
    data = JSON.parse(response.getContentText() || "{}");
  } catch (error) {
    data = {};
  }
  if (status < 200 || status >= 300) {
    const message = data.error && data.error.message ? data.error.message : "Google Places request failed.";
    const requestError = new Error(status + ": " + message);
    requestError.status = status;
    throw requestError;
  }
  return data;
}

function sheetSafe(value) {
  const text = String(value == null ? "" : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function normalize(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizeRating(value) {
  const rating = Number(value);
  return isFinite(rating) && rating >= 0.5 && rating <= 5 ? Math.round(rating * 10) / 10 : "";
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getResponse(data, event) {
  const callback = event && event.parameter ? event.parameter.callback : "";
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(data) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse(data);
}

function errorResponse(error) {
  return jsonResponse({
    ok: false,
    error: error && error.message ? error.message : String(error)
  });
}
