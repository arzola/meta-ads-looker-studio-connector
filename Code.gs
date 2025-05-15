/**
 * Looker Studio Connector for Meta Ads Data
 * This connector uses the Google Apps Script and Data Studio Connector framework
 * to pull data from Meta's Marketing API into Looker Studio dashboards.
 */

// Global configuration
var CONFIG = {
  API_VERSION: 'v18.0',
  AUTH_TYPES: ['OAUTH2', 'USER_TOKEN'],
  DATE_RANGE_TYPES: ['CUSTOM', 'LAST_30_DAYS', 'LAST_7_DAYS', 'YESTERDAY']
};

var cc = DataStudioApp.createCommunityConnector();

// https://developers.google.com/datastudio/connector/reference#getauthtype
function getAuthType() {
  var AuthTypes = cc.AuthType;
  Logger.log('getAuthType');
  return cc
    .newAuthTypeResponse()
    .setAuthType(AuthTypes.OAUTH2)
    .build();
}

/**
 * Returns the user configurable options for the connector.
 * @param {object} request The request object
 * @return {object} The user configurable options
 */
function getConfig(request) {
  Logger.log("getConfig called");
  Logger.log(JSON.stringify(request));
  return {
    configParams: [
      {
        type: 'INFO',
        name: 'connect',
        text: 'This connector allows you to import Meta Ads data for Spend, Total Conversion Value, and Campaign Name into Looker Studio. Enter your Ad Account ID below.'
      },
      {
        type: 'TEXTINPUT',
        name: 'accountId',
        displayName: 'Ad Account ID Number',
        helpText: 'Enter only the numerical part of your Meta Ad Account ID',
        placeholder: '1234567890'
      }
    ]
  };
}

/**
 * Returns the schema for the given request.
 * @param {object} request The request object
 * @return {object} The schema
 */
function getSchema(request) {
  var fields = [];
  // var configParams = request.configParams; // No longer needed for schema definition

  Logger.log('getSchema called - returning fixed schema');
  // Logger.log(JSON.stringify(request.configParams)); // configParams not used here anymore
  
  // Always include date dimension
  fields.push({
    name: 'date_start',
    label: 'Date',
    dataType: 'STRING',
    semantics: {
      conceptType: 'DIMENSION',
      semanticType: 'YEAR_MONTH_DAY'
    }
  });
  
  // Always include Campaign Name dimension
  fields.push({
    name: 'campaign_name',
    label: 'Campaign Name',
    dataType: 'STRING',
    semantics: {
      conceptType: 'DIMENSION'
    }
  });
  
  // Always include Spend metric
  fields.push({
    name: 'spend',
    label: 'Spend (Cost)',
    dataType: 'NUMBER',
    semantics: {
      conceptType: 'METRIC' // CURRENCY_USD was previously removed, keeping it that way
    }
  });
  
  // Always include Total Conversion Value metric
  fields.push({
    name: 'conversion_value_total',
    label: 'Total Conversion Value',
    dataType: 'NUMBER',
    semantics: {
      conceptType: 'METRIC' // CURRENCY_USD was previously removed, keeping it that way
    }
  });

  // Logic for dynamically adding dimensions and metrics based on configParams is removed.
  
  Logger.log('Fixed schema fields: ' + JSON.stringify(fields));
  
  return { schema: fields };
}

/**
 * Returns the tabular data for the given request.
 * @param {object} request The request object
 * @return {object} The tabular data
 */
function getData(request) {
  var requestedFields; // Declare here to be available in catch block
  try {
    Logger.log('getData called');
    Logger.log('Request: ' + JSON.stringify(request));

    var configParams = request.configParams;
    requestedFields = getRequestedFields(request); // Assign here

    // Validate required parameters using validateConfig
    var validation = validateConfig(request);
    if (!validation.isValid) {
      var errorMessages = validation.errors.map(function(err) { return err.message; }).join(', ');
      cc.newUserError()
        .setText('Configuration Error: ' + errorMessages)
        .setDebugText('Validation failed: ' + JSON.stringify(validation.errors))
        .throwException();
    }

    // Get the numerical account ID from config and prepend 'act_'
    var accountIdNumber = configParams.accountId.trim();
    var accountId = 'act_' + accountIdNumber;
    Logger.log('Using Account ID: ' + accountId); // Log the full ID being used

    // Check authentication
    if (!isAuthValid()) {
      cc.newUserError()
        .setText('Authentication is invalid or expired. Please re-authenticate the connector.')
        .setDebugText('isAuthValid() returned false.')
        .throwException();
    }

    // Set up date range - handle Looker Studio date range if present
    // Pass the request object to getDateRange
    var dateRange = getDateRange(configParams.dateRangeType, request);
    Logger.log('Using date range: ' + JSON.stringify(dateRange));


    // Construct fields for the API request
    var apiFields = constructApiFields(requestedFields);
    Logger.log('API fields: ' + JSON.stringify(apiFields));

    // Handle action_breakdowns for conversion metrics
    var hasConversionMetrics = hasMetrics(requestedFields, ['conversion_value_total', 'actions']);
    Logger.log('Has conversion metrics: ' + hasConversionMetrics);

    // Build breakdowns parameter
    var breakdowns = buildBreakdowns(configParams.dimensions);
    Logger.log('Breakdowns: ' + breakdowns);

    // *** BEGIN VALIDATION: Check for conflicting breakdowns and conversion metrics ***
    var conflictingBreakdowns = ['age', 'gender', 'country', 'device_platform'];
    var requestedBreakdowns = breakdowns ? breakdowns.split(',') : [];
    var hasConflictingBreakdown = requestedBreakdowns.some(function(b) {
      return conflictingBreakdowns.indexOf(b) !== -1;
    });

    if (hasConversionMetrics && hasConflictingBreakdown) {
      var conflictingList = requestedBreakdowns.filter(function(b) { return conflictingBreakdowns.indexOf(b) !== -1; }).join(', ');
      Logger.log('Validation Error: Conflicting breakdowns (' + conflictingList + ') requested with conversion metrics.');
      cc.newUserError()
        .setText('Invalid Configuration: Conversion metrics (like Total Conversion Value, Actions) cannot be combined with breakdowns like Age, Gender, Country, or Device Platform due to Meta API limitations. Please remove either the conversion metrics or these specific breakdowns.')
        .setDebugText('Conflict between action_breakdowns (implied by conversion metrics) and requested breakdowns: ' + conflictingList)
        .throwException();
    }
    // *** END VALIDATION ***

    // Make API request to Meta Ads Insights API
    Logger.log('Fetching insights from Meta API...');
    var allData = [];
    var nextUrl = null;

    // Initial fetch
    var response = fetchInsights(accountId, dateRange, apiFields, breakdowns, hasConversionMetrics, null); // Start with no nextUrl

    // Handle potential immediate API error
    if (response.error) {
      Logger.log('API returned error on initial fetch: ' + JSON.stringify(response.error));
      cc.newUserError()
        .setText('Meta API Error: ' + response.error.message)
        .setDebugText(JSON.stringify(response.error))
        .throwException();
    }

    if (response && response.data) {
      allData = allData.concat(response.data);
      nextUrl = (response.paging && response.paging.next) ? response.paging.next : null;
    } else {
      Logger.log('Initial API response format invalid or no data.');
    }

    // Handle pagination
    var pageCount = 1;
    while (nextUrl && pageCount < 100) { // Add a safety limit for pages
        Logger.log('Fetching next page: ' + pageCount);
        pageCount++;
        response = fetchInsights(accountId, dateRange, apiFields, breakdowns, hasConversionMetrics, nextUrl); // Fetch using next URL

        if (response.error) {
          Logger.log('API returned error on page ' + pageCount + ': ' + JSON.stringify(response.error));
          // Decide whether to stop or continue. For now, stop and report.
          cc.newUserError()
            .setText('Meta API Error during pagination: ' + response.error.message)
            .setDebugText('Error on page ' + pageCount + ': ' + JSON.stringify(response.error))
            .throwException();
        }

        if (response && response.data) {
            allData = allData.concat(response.data);
            nextUrl = (response.paging && response.paging.next) ? response.paging.next : null;
        } else {
            Logger.log('Invalid response or no data on page ' + pageCount);
            nextUrl = null; // Stop pagination
        }
    }
     if (pageCount >= 100) {
       Logger.log('Reached pagination limit (100 pages). Data might be incomplete.');
       // Optionally inform the user via an error or just log it.
     }


    Logger.log('Total records fetched after pagination: ' + allData.length);

    if (allData.length === 0) {
      Logger.log('No data returned from API after processing all pages.');
      return {
        schema: requestedFields.schema,
        rows: []
      };
    }

    // Process and format the response
    Logger.log('Processing ' + allData.length + ' rows from API response');
    var rows = processResponse({ data: allData }, requestedFields); // Pass the combined data and the dateRange used for fetching
    Logger.log('Processed ' + rows.length + ' rows for Looker Studio');

    return {
      schema: requestedFields.schema,
      rows: rows
    };
  } catch (e) {
    Logger.log('Error in getData: ' + e.toString());
    // Check if it's a user error thrown by throwException()
    if (e.constructor.name === 'UserError') {
      throw e; // Re-throw user errors directly
    }
    // Otherwise, create a new user error for unexpected issues
    cc.newUserError()
      .setText('An unexpected error occurred: ' + e.message)
      .setDebugText(e.stack || e.toString())
      .throwException();
  }
}

/**
 * Checks if requested fields contain specific metrics.
 * @param {object} requestedFields The requested fields
 * @param {Array} metricNames Array of metric names to check
 * @return {boolean} True if any of the metrics are requested
 */
function hasMetrics(requestedFields, metricNames) {
  if (!requestedFields || !requestedFields.schema) return false; // Add check
  var fieldNames = requestedFields.schema.map(function(field) {
    return field.name;
  });
  
  return metricNames.some(function(metric) {
    return fieldNames.indexOf(metric) >= 0;
  });
}

/**
 * Gets the requested fields from the request.
 * @param {object} request The request object
 * @return {object} The requested fields
 */
function getRequestedFields(request) {
  var schema = getSchema(request).schema;
  var requestedFieldIds = request.fields.map(function(field) {
    return field.name;
  });
  
  return {
    schema: schema.filter(function(field) {
      return requestedFieldIds.indexOf(field.name) >= 0;
    })
  };
}

/**
 * Gets the date range based on the selected type or Looker Studio request.
 * @param {string} dateRangeType The date range type from config
 * @param {object} request The getData request object containing potential dateRange
 * @return {object} The date range with start and end dates
 */
function getDateRange(dateRangeType, request) { // dateRangeType is configParams.dateRangeType
  Logger.log('getDateRange called. Date range type from config: ' + dateRangeType + '. LS request.dateRange: ' + JSON.stringify(request.dateRange));

  // Priority 1: Use date range from Looker Studio request if complete
  if (request && request.dateRange && request.dateRange.startDate && request.dateRange.endDate) {
    Logger.log('Using date range from Looker Studio request.');
    return {
      since: request.dateRange.startDate,
      until: request.dateRange.endDate
    };
  }

  Logger.log('Looker Studio date range not used or incomplete. Processing config dateRangeType: ' + dateRangeType);
  var today = new Date();
  var startDate, endDate;

  // Default endDate to yesterday
  endDate = new Date(today);
  endDate.setDate(today.getDate() - 1);

  // If dateRangeType from config is undefined/null/empty, default it to 'LAST_30_DAYS' for fallback logic.
  var effectiveDateRangeType = dateRangeType;
  if (effectiveDateRangeType === undefined || effectiveDateRangeType === null || String(effectiveDateRangeType).trim() === '') {
      Logger.log('Config dateRangeType is not set, defaulting to LAST_30_DAYS for fallback logic.');
      effectiveDateRangeType = 'LAST_30_DAYS';
  }

  switch (effectiveDateRangeType) {
    case 'LAST_30_DAYS':
      startDate = new Date(today);
      startDate.setDate(today.getDate() - 30);
      break;
    case 'LAST_7_DAYS':
      startDate = new Date(today);
      startDate.setDate(today.getDate() - 7);
      break;
    case 'YESTERDAY':
      startDate = new Date(today);
      startDate.setDate(today.getDate() - 1);
      break;
    case 'CUSTOM':
      Logger.log('Warning: Config dateRangeType is CUSTOM, but no date range provided by Looker Studio. Defaulting to LAST_30_DAYS.');
      startDate = new Date(today);
      startDate.setDate(today.getDate() - 30);
      break;
    default: // This handles genuinely unknown/invalid dateRangeType values from config
      Logger.log('Warning: Unknown date range type configured: "' + effectiveDateRangeType + '". Defaulting to LAST_30_DAYS.');
      startDate = new Date(today);
      startDate.setDate(today.getDate() - 30);
      break;
  }

  // Ensure start date is not after end date
  if (startDate > endDate) {
      startDate = new Date(endDate); 
      Logger.log('Adjusted start date to match end date as initial calculation was invalid.');
  }

  Logger.log('Fallback date range determined: since ' + formatDate(startDate) + ', until ' + formatDate(endDate));
  return {
    since: formatDate(startDate),
    until: formatDate(endDate)
  };
}

/**
 * Formats a date as YYYY-MM-DD.
 * @param {Date} date The date to format
 * @return {string} The formatted date
 */
function formatDate(date) {
  return date.toISOString().slice(0, 10);
}

/**
 * Constructs fields parameter for the API request.
 * @param {object} requestedFields The requested fields
 * @return {string} The fields parameter
 */
function constructApiFields(requestedFields) {
  var fields = ['date_start'];
  var actionFields = [];
  
  requestedFields.schema.forEach(function(field) {
    if (field.name !== 'date_start') {
      // Handle special fields that require action_breakdowns
      if (field.name === 'conversion_value_total') {
        fields.push('action_values');
        actionFields.push('action_values');
      } else if (field.name === 'actions') {
        fields.push('actions');
        actionFields.push('actions');
      } else if (field.name !== 'roas') {
        // Standard fields (excluding ROAS)
        fields.push(field.name);
      }
    }
  });
  
  return {
    regular: fields.join(','),
    actionFields: actionFields
  };
}

/**
 * Builds the breakdowns parameter for the API request.
 * @param {string} dimensions The requested dimensions
 * @return {string} The breakdowns parameter
 */
function buildBreakdowns(dimensions) {
  if (!dimensions) return '';
  
  var breakdownMap = {
    'age': 'age',
    'gender': 'gender',
    'country': 'country',
    'device_platform': 'device_platform'
  };
  
  var breakdowns = [];
  dimensions.split(',').forEach(function(dimension) {
    if (breakdownMap[dimension]) {
      breakdowns.push(breakdownMap[dimension]);
    }
  });
  
  return breakdowns.join(',');
}

/**
 * Fetches insights data from Meta Marketing API.
 * @param {string} accountId The ad account ID
 * @param {object} dateRange The date range
 * @param {object} fields The fields to fetch
 * @param {string} breakdowns The breakdowns to apply
 * @param {boolean} hasConversionMetrics Whether conversion metrics are requested
 * @param {string} [nextUrl] Optional URL for the next page of results
 * @return {object} The API response
 */
function fetchInsights(accountId, dateRange, fields, breakdowns, hasConversionMetrics, nextUrl) {

  Logger.log('fetchInsights called. Fetching URL: ' + (nextUrl || 'Initial Fetch'));

  var token = getOAuthService().getAccessToken();
  if (!token) {
      cc.newUserError()
        .setText('Unable to retrieve access token. Please re-authenticate.')
        .setDebugText('getOAuthService().getAccessToken() returned null or empty.')
        .throwException();
  }

  var endpoint;
  var params = { access_token: token };
  var urlToFetch;

  if (nextUrl) {
    // If nextUrl is provided, use it directly. It already contains parameters.
    urlToFetch = nextUrl;
  } else {
    // Construct the initial URL
    endpoint = 'https://graph.facebook.com/' + CONFIG.API_VERSION +
                 '/' + accountId + '/insights';

    params.time_range = JSON.stringify({
      since: dateRange.since,
      until: dateRange.until
    });
    params.fields = fields.regular;
    params.level = 'campaign'; // Changed from 'ad' to 'campaign'
    params.time_increment = 1; // Daily data
    params.limit = 100; // Request a reasonable limit per page

    if (breakdowns) {
      params.breakdowns = breakdowns;
    }

    // Add action breakdowns when conversion metrics are requested
    if (hasConversionMetrics) {
      // Common action types. Consider making this configurable or dynamic.
      params.action_breakdowns = 'action_type';
    }

    urlToFetch = endpoint + '?' + objectToQueryString(params);
  }

  Logger.log('Request URL: ' + urlToFetch);

  // *** ADD DETAILED LOGGING HERE ***
  Logger.log('Attempting to fetch URL: ' + urlToFetch);

  var options = {
    method: 'GET',
    muteHttpExceptions: true, // Important to check response code manually
    headers: {
      'Authorization': 'Bearer ' + token
    }
  };

  var response;
  var responseCode;
  var responseBody;
  try {
    response = UrlFetchApp.fetch(urlToFetch, options);
    responseCode = response.getResponseCode();
    responseBody = response.getContentText();
  } catch (fetchError) {
      Logger.log('Network error during API fetch: ' + fetchError.toString());
      cc.newUserError()
        .setText('Network error communicating with Meta API: ' + fetchError.message)
        .setDebugText(fetchError.stack || fetchError.toString())
        .throwException();
  }


  Logger.log('API Response Code: ' + responseCode);
  // Log cautiously for potential large responses, maybe just first 1000 chars
  // Logger.log('API Response Body: ' + (responseBody ? responseBody.substring(0, 1000) : 'EMPTY'));


  if (responseCode !== 200) {
    Logger.log('API Error Response (' + responseCode + '): ' + responseBody);
    var errorData = parseJsonSafely(responseBody);
    var errorMessage = 'API request failed with status code ' + responseCode + '.';
    if (errorData && errorData.error && errorData.error.message) {
        errorMessage = 'Meta API Error (' + responseCode + '): ' + errorData.error.message;
        // Check for specific error types like token expiration
        if (errorData.error.type === 'OAuthException' || (errorData.error.code && (errorData.error.code === 190 || errorData.error.code === 102))) {
           errorMessage += ' Your access token may have expired. Please re-authenticate.';
           // Optionally reset auth here if appropriate: resetAuth();
        }
    }
    cc.newUserError()
        .setText(errorMessage)
        .setDebugText('Response Code: ' + responseCode + ', Body: ' + responseBody)
        .throwException();
  }

  var jsonData = parseJsonSafely(responseBody);
   if (!jsonData) {
       Logger.log('Failed to parse JSON response: ' + responseBody);
       cc.newUserError()
           .setText('Failed to parse response from Meta API.')
           .setDebugText('Invalid JSON received: ' + responseBody)
           .throwException();
   }

   // Check for error structure even within a 200 response (though less common for insights)
   if (jsonData.error) {
     Logger.log('API returned error within 200 response: ' + JSON.stringify(jsonData.error));
     cc.newUserError()
       .setText('Meta API Error: ' + jsonData.error.message)
       .setDebugText(JSON.stringify(jsonData.error))
       .throwException();
   }


  return jsonData; // Return parsed JSON
}

/**
 * Safely parses a JSON string.
 * @param {string} jsonString The JSON string to parse.
 * @return {object|null} The parsed object or null if parsing fails.
 */
function parseJsonSafely(jsonString) {
    try {
        return JSON.parse(jsonString);
    } catch (e) {
        Logger.log('JSON parsing error: ' + e);
        return null;
    }
}

/**
 * Processes the API response and formats it for Looker Studio.
 * Each item in response.data is expected to be a daily record.
 * @param {object} response The API response containing { data: allData }
 * @param {object} requestedFields The requested fields schema
 * @return {Array} The formatted rows for Looker Studio
 */
function processResponse(response, requestedFields) { 
  if (!response || !response.data || response.data.length === 0) {
    return [];
  }

  var schemaMap = {};
  requestedFields.schema.forEach(function(field, index) {
    schemaMap[field.name] = index;
  });

  return response.data.map(function(item) { // item is a daily record from Meta API
    var row = new Array(requestedFields.schema.length).fill(null);

    requestedFields.schema.forEach(function(field) {
      var value = null;
      if (schemaMap[field.name] === undefined) return; // Should not happen if schema is well-formed

      switch (field.name) {
        case 'date_start':
          value = item.date_start ? item.date_start.replace(/-/g, '') : '';
          break;
        
        // Direct dimensions from API item
        case 'campaign_name':
          value = item.campaign_name || '';
          break;
        case 'adset_name': // Assuming adset_name might be in item if level='adset' or 'ad' was used
          value = item.adset_name || '';
          break;
        case 'ad_name':    // Assuming ad_name might be in item if level='ad' was used
          value = item.ad_name || ''; 
          break;
        case 'age':
          value = item.age || '';
          break;
        case 'gender':
          value = item.gender || '';
          break;
        case 'country':
          value = item.country || '';
          break;
        case 'device_platform':
          value = item.device_platform || '';
          break;

        // Base metrics from API item
        case 'spend':
        case 'impressions':
        case 'clicks':
        case 'reach':
          value = parseFloat(item[field.name]) || 0;
          break;

        // Conversion metrics specific to this item (daily record)
        case 'conversion_value_total':
          value = extractPurchaseConversionValue(item.action_values);
          break;
        case 'actions':
          value = extractPurchaseActionCount(item.actions);
          break;

        // Calculated metrics for this item (daily record)
        case 'cpc':
          var clicks_cpc = parseFloat(item.clicks) || 0;
          var spend_cpc = parseFloat(item.spend) || 0;
          value = clicks_cpc > 0 ? spend_cpc / clicks_cpc : 0;
          break;
        case 'cpm':
          var impressions_cpm = parseFloat(item.impressions) || 0;
          var spend_cpm = parseFloat(item.spend) || 0;
          value = impressions_cpm > 0 ? (spend_cpm / impressions_cpm) * 1000 : 0;
          break;
        case 'ctr':
          var impressions_ctr = parseFloat(item.impressions) || 0;
          var clicks_ctr = parseFloat(item.clicks) || 0;
          value = impressions_ctr > 0 ? (clicks_ctr / impressions_ctr) * 100 : 0;
          break;
        case 'frequency':
          var impressions_freq = parseFloat(item.impressions) || 0;
          var reach_freq = parseFloat(item.reach) || 0;
          value = reach_freq > 0 ? impressions_freq / reach_freq : 0;
          break;
        
        default:
          // Fallback for any other metric not explicitly handled above but present in item
          if (item.hasOwnProperty(field.name) && field.semantics && field.semantics.conceptType === 'METRIC') {
            value = parseFloat(item[field.name]) || 0;
          } else if (item.hasOwnProperty(field.name)) { // Other potential dimension
            value = item[field.name] || '';
          }
          break;
      }
      
      row[schemaMap[field.name]] = value;
    });

    return { values: row };
  });
}

/**
 * Extracts the total value for "purchase" conversions from an array of action objects.
 * Sums the 'value' property of objects where action_type is 'purchase'.
 * @param {Array<object>} actionValues Array of action objects (e.g., from action_values field).
 * @return {number} The total summed value for purchase conversions.
 */
function extractPurchaseConversionValue(actionValues) {
  if (!actionValues || !Array.isArray(actionValues)) return 0;
  var totalPurchaseValue = 0;
  actionValues.forEach(function(action) {
    if (action.action_type === 'purchase') { // Specifically look for 'purchase'
      totalPurchaseValue += parseFloat(action.value) || 0;
    }
  });
  return totalPurchaseValue;
}

/**
 * Extracts the total count for "purchase" actions from an array of action objects.
 * Sums the 'value' property (count) of objects where action_type is 'purchase'.
 * @param {Array<object>} actionsList Array of action objects (e.g., from actions field).
 * @return {number} The total summed count for purchase actions.
 */
function extractPurchaseActionCount(actionsList) {
  if (!actionsList || !Array.isArray(actionsList)) return 0;
  var totalPurchaseCount = 0;
  actionsList.forEach(function(action) {
    if (action.action_type === 'purchase') { // Specifically look for 'purchase'
      totalPurchaseCount += parseInt(action.value) || 0;
    }
  });
  return totalPurchaseCount;
}

/**
 * Creates and returns the OAuth2 service.
 * @return {OAuth2Service} The OAuth2 service
 */
function getOAuthService() {
  // IMPORTANT: Store your Client ID and Secret in Script Properties.
  // Go to File > Project properties > Script properties.
  // Add two properties: 'META_CLIENT_ID' and 'META_CLIENT_SECRET'.
  var scriptProperties = PropertiesService.getScriptProperties();
  var clientId = scriptProperties.getProperty('META_CLIENT_ID');
  var clientSecret = scriptProperties.getProperty('META_CLIENT_SECRET');

  if (!clientId || !clientSecret) {
      // Throw an error or provide guidance if properties are not set.
      // This error will appear during connector setup/authorization.
      throw new Error("OAuth2 Client ID or Secret not set in Script Properties. Please configure 'META_CLIENT_ID' and 'META_CLIENT_SECRET'.");
  }


  // This function would be implemented with your OAuth credentials
  // You would need to create a project in Google Cloud Console and set up OAuth credentials
  return OAuth2.createService('facebook')
    .setAuthorizationBaseUrl('https://www.facebook.com/' + CONFIG.API_VERSION + '/dialog/oauth')
    .setTokenUrl('https://graph.facebook.com/' + CONFIG.API_VERSION + '/oauth/access_token')
    .setClientId(clientId) // Use property value
    .setClientSecret(clientSecret) // Use property value
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('ads_read,read_insights'); // Added read_insights scope
}

/**
 * Handles the OAuth callback.
 * @param {object} request The request data
 * @return {HtmlOutput} The HTML output
 */
function authCallback(request) {
  var authorized = getOAuthService().handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Authorization denied.');
  }
}

/**
 * Resets the OAuth service.
 */
function resetAuth() {
  getOAuthService().reset();
}

/**
 * Required for Looker Studio connector.
 * @return {boolean} Always returns true
 */
function isAdminUser() {
  return true;
}

function isAuthValid() {
  var service = getOAuthService();
  return service.hasAccess();
}

function get3PAuthorizationUrls() {
  return getOAuthService().getAuthorizationUrl();
}

/**
 * Indicates if the connector supports data refresh.
 * @return {boolean} True if data refresh is supported
 */
function isDataRefreshable() {
  return true;
}

/**
 * This checks for parameters required by the connector.
 * @param {Object} request The request object.
 * @return {object} errors
 */
function validateConfig(request) {
  Logger.log('Validating config: ' + JSON.stringify(request.configParams));
  var configParams = request.configParams || {}; // Ensure configParams exists
  var errors = [];

  // Validate Account ID (now expecting only numbers)
  if (!configParams.accountId || configParams.accountId.trim() === '') {
    errors.push({
      errorCode: 'MISSING_ACCOUNT_ID',
      message: "Ad Account ID number is required."
    });
  } else if (!/^[0-9]+$/.test(configParams.accountId.trim())) { // Check if it's only digits
     errors.push({
         errorCode: 'INVALID_ACCOUNT_ID_FORMAT',
         message: "Ad Account ID must contain only numbers."
     });
  }

  // Metrics validation is removed as metrics are now fixed.

  // Add more validation as needed (e.g., date range type)

  Logger.log('Validation result: ' + (errors.length === 0) + ', Errors: ' + JSON.stringify(errors));

  return {
    isValid: errors.length === 0,
    errors: errors // Return the array of error objects
  };
}

// Add helper function for validation in getData
function getValidatedConfig(request) {
    var result = validateConfig(request);
    if (!result.isValid) {
        var errorMessages = result.errors.map(function(err) { return err.message; }).join(' ');
        cc.newUserError()
            .setText('Configuration Error: ' + errorMessages)
            .setDebugText(JSON.stringify(result.errors))
            .throwException();
    }
    return request.configParams;
}

/**
 * Converts a JavaScript object to a URL query string.
 * @param {object} obj The object to convert.
 * @return {string} The URL query string.
 */
function objectToQueryString(obj) {
  return Object.keys(obj).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(obj[key]);
  }).join('&');
}