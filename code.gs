// ============================================
// AXIOM GRC CROSS-SELL NAVIGATOR v2.3
// DATABASE-DRIVEN VERSION - All logic now in Excel
// ============================================

var SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ============================================
// WEB APP ENTRY
// ============================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Axiom Xplore')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================
// CONFIGURATION MANAGEMENT (NEW)
// ============================================

/**
 * Get configuration value from CONFIG sheet
 * Falls back to default if not found
 */
function getConfig(settingName, defaultValue) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'config_' + settingName;
    var cached = cache.get(cacheKey);
    
    if (cached !== null) {
      return cached;
    }
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('CONFIG');
    
    if (!sheet) {
      Logger.log('WARNING: CONFIG sheet not found, using default: ' + defaultValue);
      return defaultValue;
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === settingName) {
        var value = data[i][1];
        cache.put(cacheKey, String(value), 3600); // Cache for 1 hour
        return value;
      }
    }
    
    return defaultValue;
    
  } catch (error) {
    Logger.log('ERROR in getConfig: ' + error.toString());
    return defaultValue;
  }
}

/**
 * Get Cross-Sell Lead form URL from CONFIG sheet
 */
function getCrossSellLeadUrl() {
  return getConfig('CROSS_SELL_LEAD_URL', '');
}

// ============================================
// CORE DATA LOADING WITH CACHING & ERROR HANDLING
// ============================================

function getAllProducts() {
  try {
    var enableCaching = getConfig('ENABLE_CACHING', 'TRUE');
    var cache = CacheService.getScriptCache();
    var cached = null;
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      cached = cache.get('all_products');
      if (cached) {
        return JSON.parse(cached);
      }
    }
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('PRODUCTS');
    
    if (!sheet) {
      throw new Error('PRODUCTS sheet not found in database');
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      throw new Error('PRODUCTS sheet is empty');
    }
    
    // Get company logos for mapping
    var companyLogos = getCompanyLogoMap();
    
    var headers = data[0];
    var products = [];
    
    for (var i = 1; i < data.length; i++) {
      var product = {};
      for (var j = 0; j < headers.length; j++) {
        var value = data[i][j];
        var header = headers[j];
        
        // Normalize arrays (semicolon-separated fields)
        if (header === 'PrimaryIndustries') {
          // Industries now uses semicolons as separator
          product.Industries = normalizeArraySemicolon(value);
          product[header] = value;
        } else if (header === 'KeyStats') {
          product.KeySellingPoints = normalizeArraySemicolon(value);
          product[header] = value;
        } else {
          product[header] = value || '';
        }
      }
      
      // ====== ALL DERIVATION LOGIC REMOVED - NOW READ FROM DATABASE ======
      
      // TargetClientSize is read directly from database - no default set
      // Blank values remain blank
      
      // Read demo/learn more flags from database (or default to false)
      product.hasDemo = product.HasDemo === true || product.HasDemo === 'TRUE' || product.HasDemo === 'Yes';
      product.hasLearnMore = product.HasLearnMore === true || product.HasLearnMore === 'TRUE' || product.HasLearnMore === 'Yes';
      
      // Pricing bands from database
      product.PricingBand_SME = product.SME_Value || 'STILL NEEDED';
      product.PricingBand_Enterprise = product.Enterprise_Value || 'STILL NEEDED';
      
      // Add company logo URL
      product.CompanyLogoURL = companyLogos[product.CompanyName] || '';
      
      products.push(product);
    }
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      var cacheDuration = parseInt(getConfig('CACHE_DURATION_SECONDS', 600));
      cache.put('all_products', JSON.stringify(products), cacheDuration);
    }
    
    return products;
    
  } catch (error) {
    Logger.log('ERROR in getAllProducts: ' + error.toString());
    throw new Error('Failed to load products: ' + error.message);
  }
}

/**
 * Get company logo mapping
 */
function getCompanyLogoMap() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('COMPANY_INFO');
    
    if (!sheet) {
      return {};
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var logoMap = {};
    
    // Find column indices by name instead of hardcoded positions
    var companyNameIndex = headers.indexOf('CompanyName');
    var logoURLIndex = headers.indexOf('LogoURL');
    
    // If columns not found, return empty map
    if (companyNameIndex === -1 || logoURLIndex === -1) {
      Logger.log('WARNING: CompanyName or LogoURL column not found in COMPANY_INFO sheet');
      return {};
    }
    
    for (var i = 1; i < data.length; i++) {
      var companyName = data[i][companyNameIndex];
      var logoURL = data[i][logoURLIndex];
      
      if (companyName && logoURL) {
        logoMap[companyName] = convertDriveLinkToDirectURL(logoURL);
      }
    }
    
    return logoMap;
    
  } catch (error) {
    Logger.log('ERROR in getCompanyLogoMap: ' + error.toString());
    return {};
  }
}

function getProductDetails() {
  try {
    var enableCaching = getConfig('ENABLE_CACHING', 'TRUE');
    var cache = CacheService.getScriptCache();
    var cached = null;
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      cached = cache.get('product_details');
      if (cached) {
        return JSON.parse(cached);
      }
    }
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('PRODUCT_DETAILS');
    
    if (!sheet) {
      Logger.log('WARNING: PRODUCT_DETAILS sheet not found');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      Logger.log('WARNING: PRODUCT_DETAILS sheet is empty');
      return [];
    }
    
    var headers = data[0];
    var details = [];
    
    for (var i = 1; i < data.length; i++) {
      var detail = {};
      for (var j = 0; j < headers.length; j++) {
        detail[headers[j]] = data[i][j] || '';
      }
      details.push(detail);
    }
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      var cacheDuration = parseInt(getConfig('CACHE_DURATION_SECONDS', 600));
      cache.put('product_details', JSON.stringify(details), cacheDuration);
    }
    
    return details;
    
  } catch (error) {
    Logger.log('ERROR in getProductDetails: ' + error.toString());
    return [];
  }
}

function getCompanies() {
  try {
    var enableCaching = getConfig('ENABLE_CACHING', 'TRUE');
    var cache = CacheService.getScriptCache();
    var cached = null;
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      cached = cache.get('companies');
      if (cached) {
        return JSON.parse(cached);
      }
    }
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('COMPANY_INFO');
    
    if (!sheet) {
      Logger.log('WARNING: COMPANY_INFO sheet not found');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      Logger.log('WARNING: COMPANY_INFO sheet is empty');
      return [];
    }
    
    var headers = data[0];
    var companies = [];
    
    for (var i = 1; i < data.length; i++) {
      var company = {};
      for (var j = 0; j < headers.length; j++) {
        company[headers[j]] = data[i][j] || '';
      }
      companies.push(company);
    }
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      var cacheDuration = parseInt(getConfig('CACHE_DURATION_SECONDS', 600));
      cache.put('companies', JSON.stringify(companies), cacheDuration);
    }
    
    return companies;
    
  } catch (error) {
    Logger.log('ERROR in getCompanies: ' + error.toString());
    return [];
  }
}

function getCrossSellData() {
  try {
    var enableCaching = getConfig('ENABLE_CACHING', 'TRUE');
    var cache = CacheService.getScriptCache();
    var cached = null;
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      cached = cache.get('rules');
      if (cached) {
        return JSON.parse(cached);
      }
    }
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('CROSSSELL_DATA');
    
    if (!sheet) {
      Logger.log('WARNING: CROSSSELL_DATA sheet not found');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      Logger.log('WARNING: CROSSSELL_DATA sheet is empty');
      return [];
    }
    
    var headers = data[0];
    var rules = [];
    
    for (var i = 1; i < data.length; i++) {
      var rule = {};
      for (var j = 0; j < headers.length; j++) {
        rule[headers[j]] = data[i][j] || '';
      }
      rules.push(rule);
    }
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      var cacheDuration = parseInt(getConfig('CACHE_DURATION_SECONDS', 600));
      cache.put('rules', JSON.stringify(rules), cacheDuration);
    }
    
    return rules;
    
  } catch (error) {
    Logger.log('ERROR in getCrossSellData: ' + error.toString());
    return [];
  }
}

/**
 * Get unique industries from PRODUCTS sheet
 */
function getUniqueIndustries() {
  try {
    var products = getAllProducts();
    var industriesSet = new Set();
    
    products.forEach(function(product) {
      if (product.Industries && Array.isArray(product.Industries)) {
        product.Industries.forEach(function(industry) {
          if (industry && industry !== 'all') {
            // Capitalize first letter
            var formatted = industry.charAt(0).toUpperCase() + industry.slice(1);
            industriesSet.add(formatted);
          }
        });
      }
    });
    
    return Array.from(industriesSet).sort();
    
  } catch (error) {
    Logger.log('ERROR in getUniqueIndustries: ' + error.toString());
    return [];
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function normalizeArray(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value;
  
  return String(value).split(',').map(function(item) {
    return item.trim();
  }).filter(function(item) {
    return item.length > 0;
  });
}

function normalizeArraySemicolon(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value;
  
  return String(value).split(';').map(function(item) {
    return item.trim();
  }).filter(function(item) {
    return item.length > 0;
  });
}

// ============================================
// SCORING AND RECOMMENDATION ENGINE
// ============================================

function generateRecommendations(clientData) {
  try {
    var products = getAllProducts();
    var rules = getCrossSellData();
    
    // Create a map of ProductID to Product for quick lookup
    var productMap = {};
    products.forEach(function(product) {
      productMap[product.ProductID] = product;
    });
    
    var currentProductIDs = clientData.currentProducts || [];
    var selectedSize = clientData.companySize || 'Both';
    var selectedIndustries = clientData.industries || [];
    
    var opportunities = [];
    var processedProductIDs = {}; // Track which products we've already added
    
    // Helper function to check if direction is bidirectional (case-insensitive, dash optional)
    function isBidirectional(direction) {
      if (!direction) return false;
      var normalized = String(direction).toLowerCase().replace(/-/g, '');
      return normalized === 'bidirectional' || normalized === 'bi directional';
    }
    
    // Helper function to get data source priority (lower number = higher priority)
    function getDataSourcePriority(dataSource) {
      if (!dataSource) return 4; // Default to lowest priority
      var ds = String(dataSource).toLowerCase().trim();
      // Priority 1: "real,expert" or "real and expert"
      if (ds.includes('real') && ds.includes('expert')) {
        return 1;
      }
      // Priority 2: "real" only (not containing "expert")
      if (ds.includes('real') && !ds.includes('expert')) {
        return 2;
      }
      // Priority 3: "expert" only (not containing "real")
      if (ds.includes('expert') && !ds.includes('real')) {
        return 3;
      }
      // Priority 4: "inferred" or "ai-inferred"
      if (ds.includes('inferred')) {
        return 4;
      }
      // Default priority for other data sources
      return 4;
    }
    
    // Find all matching rules based on current products
    rules.forEach(function(rule) {
      var isBidir = isBidirectional(rule.Direction);
      
      // Check A -> B direction: if client has ProductA, recommend ProductB
      currentProductIDs.forEach(function(currentProductID) {
        if (rule.ProductA === currentProductID) {
          // Client has ProductA, recommend ProductB
          var recommendedProduct = productMap[rule.ProductB];
          if (!recommendedProduct) {
            return; // Product not found in products list
          }
          
          // Skip if client already owns this product
          if (currentProductIDs.indexOf(rule.ProductB) !== -1) {
            return;
          }
          
          // Check size match (SME_Enterprise column)
          var sizeMatch = false;
          if (rule.SME_Enterprise === 'Both') {
            sizeMatch = true;
          } else if (rule.SME_Enterprise === 'SME' && (selectedSize === 'SME' || selectedSize === 'Both')) {
            sizeMatch = true;
          } else if (rule.SME_Enterprise === 'Enterprise' && (selectedSize === 'Enterprise' || selectedSize === 'Both')) {
            sizeMatch = true;
          }
          
          // Also check product's TargetClientSize
          if (!sizeMatch && recommendedProduct.TargetClientSize !== 'Both') {
            if (recommendedProduct.TargetClientSize === selectedSize || selectedSize === 'Both') {
              sizeMatch = true;
            }
          }
          
          if (!sizeMatch && selectedSize !== 'Both') {
            return; // Size doesn't match
          }
      
          // Extract key benefits for AB direction
          var keyBenefits = [];
          if (rule.KeyBenefit1) keyBenefits.push(rule.KeyBenefit1);
          if (rule.KeyBenefit2) keyBenefits.push(rule.KeyBenefit2);
          if (rule.KeyBenefit3) keyBenefits.push(rule.KeyBenefit3);
          
          // Get rationale and insights source for AB direction
          var rationale = rule['Rationale AB'] || '';
          var insightsSource = rule['InsightsSource AB'] || '';
          
          // If we already have this product in opportunities, use the one with better data source priority
          if (processedProductIDs[rule.ProductB]) {
            var existingOpp = opportunities.find(function(opp) {
              return opp.product.ProductID === rule.ProductB;
            });
            if (existingOpp) {
              var existingPriority = getDataSourcePriority(existingOpp.dataSource);
              var newPriority = getDataSourcePriority(rule.DataSource);
              var existingCount = parseInt(existingOpp.clientCount) || 0;
              var newCount = parseInt(rule.ClientCount) || 0;
              
              // Update if new rule has better data source priority, or same priority but higher client count
              if (newPriority < existingPriority || 
                  (newPriority === existingPriority && newCount > existingCount)) {
                existingOpp.rationale = rationale || existingOpp.rationale;
                existingOpp.keyBenefits = keyBenefits.length > 0 ? keyBenefits : existingOpp.keyBenefits;
                existingOpp.insightsSource = insightsSource || existingOpp.insightsSource;
                existingOpp.dataSource = rule.DataSource || existingOpp.dataSource || '';
                existingOpp.clientCount = rule.ClientCount || existingOpp.clientCount || 0;
                existingOpp.rule = rule; // Store the rule for reference
              }
            }
            return;
          }
          
          // Create opportunity
          var opportunity = {
            product: recommendedProduct,
            rationale: rationale,
            keyBenefits: keyBenefits,
            insightsSource: insightsSource,
            dataSource: rule.DataSource || '',
            clientCount: rule.ClientCount || 0,
            rule: rule // Store full rule for reference
          };
          
          opportunities.push(opportunity);
          processedProductIDs[rule.ProductB] = true;
        }
      });
      
      // Check B -> A direction: if bidirectional and client has ProductB, recommend ProductA
      if (isBidir) {
        currentProductIDs.forEach(function(currentProductID) {
          if (rule.ProductB === currentProductID) {
            // Client has ProductB, recommend ProductA
            var recommendedProduct = productMap[rule.ProductA];
            if (!recommendedProduct) {
              return; // Product not found in products list
            }
            
            // Skip if client already owns this product
            if (currentProductIDs.indexOf(rule.ProductA) !== -1) {
              return;
            }
            
            // Check size match (SME_Enterprise column)
            var sizeMatch = false;
            if (rule.SME_Enterprise === 'Both') {
              sizeMatch = true;
            } else if (rule.SME_Enterprise === 'SME' && (selectedSize === 'SME' || selectedSize === 'Both')) {
              sizeMatch = true;
            } else if (rule.SME_Enterprise === 'Enterprise' && (selectedSize === 'Enterprise' || selectedSize === 'Both')) {
              sizeMatch = true;
            }
            
            // Also check product's TargetClientSize
            if (!sizeMatch && recommendedProduct.TargetClientSize !== 'Both') {
              if (recommendedProduct.TargetClientSize === selectedSize || selectedSize === 'Both') {
                sizeMatch = true;
              }
            }
            
            if (!sizeMatch && selectedSize !== 'Both') {
              return; // Size doesn't match
            }
            
            // Extract key benefits for BA direction
            var keyBenefits = [];
            if (rule['KeyBenefit BA1']) keyBenefits.push(rule['KeyBenefit BA1']);
            if (rule['KeyBenefit BA2']) keyBenefits.push(rule['KeyBenefit BA2']);
            if (rule['KeyBenefit BA3']) keyBenefits.push(rule['KeyBenefit BA3']);
            
            // Get rationale and insights source for BA direction
            var rationale = rule['Rationale BA'] || '';
            var insightsSource = rule['InsightsSource BA'] || '';
            
            // If we already have this product in opportunities, use the one with better data source priority
            if (processedProductIDs[rule.ProductA]) {
              var existingOpp = opportunities.find(function(opp) {
                return opp.product.ProductID === rule.ProductA;
              });
              if (existingOpp) {
                var existingPriority = getDataSourcePriority(existingOpp.dataSource);
                var newPriority = getDataSourcePriority(rule.DataSource);
                var existingCount = parseInt(existingOpp.clientCount) || 0;
                var newCount = parseInt(rule.ClientCount) || 0;
                
                // Update if new rule has better data source priority, or same priority but higher client count
                if (newPriority < existingPriority || 
                    (newPriority === existingPriority && newCount > existingCount)) {
                  existingOpp.rationale = rationale || existingOpp.rationale;
                  existingOpp.keyBenefits = keyBenefits.length > 0 ? keyBenefits : existingOpp.keyBenefits;
                  existingOpp.insightsSource = insightsSource || existingOpp.insightsSource;
                  existingOpp.dataSource = rule.DataSource || existingOpp.dataSource || '';
                  existingOpp.clientCount = rule.ClientCount || existingOpp.clientCount || 0;
                  existingOpp.rule = rule; // Store the rule for reference
                }
              }
              return;
            }
            
            // Create opportunity
            var opportunity = {
              product: recommendedProduct,
              rationale: rationale,
              keyBenefits: keyBenefits,
              insightsSource: insightsSource,
              dataSource: rule.DataSource || '',
              clientCount: rule.ClientCount || 0,
              rule: rule // Store full rule for reference
            };
            
            opportunities.push(opportunity);
            processedProductIDs[rule.ProductA] = true;
          }
        });
      }
    });
    
    // Sort: Prioritize by data source (real,expert > real > expert > inferred),
    // then by clientCount descending (highest to lowest)
    opportunities.sort(function(a, b) {
      var aPriority = getDataSourcePriority(a.dataSource);
      var bPriority = getDataSourcePriority(b.dataSource);
      
      // Sort by data source priority first
      if (aPriority !== bPriority) {
        return aPriority - bPriority; // Lower priority number = higher priority
      }
      
      // If same data source priority, sort by clientCount descending
      var aCount = parseInt(a.clientCount) || 0;
      var bCount = parseInt(b.clientCount) || 0;
      
      return bCount - aCount;
    });
    
    // Return top 5 maximum
    return opportunities.slice(0, 5);
    
  } catch (error) {
    Logger.log('ERROR in generateRecommendations: ' + error.toString());
    throw new Error('Failed to generate recommendations: ' + error.message);
  }
}

// ============================================
// UPDATE REQUEST SUBMISSION
// ============================================

function submitUpdateRequest(requestData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('UpdateRequests');
    
    if (!sheet) {
      throw new Error('UpdateRequests sheet not found. Please create it first.');
    }
    
    // Generate request ID
    var timestamp = new Date();
    var requestId = 'REQ' + timestamp.getTime().toString().slice(-8);
    
    // Prepare row data
    var rowData = [
      requestId,                                          // RequestID
      timestamp,                                          // Timestamp
      requestData.submitterName,                          // SubmitterName
      requestData.submitterEmail,                         // SubmitterEmail
      requestData.updateType,                             // RequestType
      getTargetSheetName(requestData.updateType),         // TargetSheet
      requestData.updateAction,                           // Action
      '',                                                 // RecordID (empty for new)
      JSON.stringify(requestData.additionalFields),       // FieldsToUpdate
      requestData.description,                            // NewValues
      requestData.justification,                          // Justification
      'Pending',                                          // Status
      '',                                                 // ReviewedBy
      '',                                                 // ReviewedDate
      ''                                                  // ReviewNotes
    ];
    
    // Append to sheet
    sheet.appendRow(rowData);
    
    // Send email notification to stakeholders
    try {
      sendUpdateRequestNotification(requestId, requestData, timestamp);
    } catch (emailError) {
      // Log email error but don't fail the request submission
      Logger.log('WARNING: Failed to send email notification: ' + emailError.toString());
    }
    
    return {
      success: true,
      requestId: requestId,
      message: 'Request submitted successfully'
    };
    
  } catch (error) {
    Logger.log('ERROR in submitUpdateRequest: ' + error.toString());
    throw new Error('Failed to submit request: ' + error.message);
  }
}

function getTargetSheetName(updateType) {
  var mapping = {
    'Product': 'PRODUCTS',
    'Contact': 'Contacts',
    'Resource': 'Resources',
    'CaseStudy': 'CaseStudies',
    'Other': 'N/A'
  };
  
  return mapping[updateType] || 'N/A';
}

function getPendingUpdateRequests() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('UpdateRequests');
    
    if (!sheet) {
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var requests = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      if (row[11] === 'Pending') {
        var request = {
          requestId: row[0],
          timestamp: row[1],
          submitterName: row[2],
          submitterEmail: row[3],
          requestType: row[4],
          targetSheet: row[5],
          action: row[6],
          fieldsToUpdate: row[8],
          newValues: row[9],
          justification: row[10],
          status: row[11]
        };
        
        requests.push(request);
      }
    }
    
    return requests;
    
  } catch (error) {
    Logger.log('ERROR in getPendingUpdateRequests: ' + error.toString());
    return [];
  }
}

/**
 * Get stakeholder email addresses from CONFIG sheet
 * Returns array of email addresses
 */
function getStakeholderEmails() {
  try {
    var stakeholderEmails = getConfig('UPDATE_REQUEST_NOTIFICATION_EMAILS', '');
    
    if (!stakeholderEmails || stakeholderEmails.trim() === '') {
      Logger.log('WARNING: No stakeholder emails configured in CONFIG sheet (UPDATE_REQUEST_NOTIFICATION_EMAILS)');
      return [];
    }
    
    // Split by semicolon and clean up
    var emails = stakeholderEmails.split(';').map(function(email) {
      return email.trim();
    }).filter(function(email) {
      return email.length > 0 && email.indexOf('@') !== -1;
    });
    
    return emails;
    
  } catch (error) {
    Logger.log('ERROR in getStakeholderEmails: ' + error.toString());
    return [];
  }
}

/**
 * Send email notification to stakeholders when an update request is submitted
 */
function sendUpdateRequestNotification(requestId, requestData, timestamp) {
  try {
    var stakeholderEmails = getStakeholderEmails();
    
    if (stakeholderEmails.length === 0) {
      Logger.log('No stakeholder emails configured, skipping email notification');
      return;
    }
    
    // Format timestamp
    var formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'MMMM dd, yyyy HH:mm:ss');
    
    // Build email subject
    var subject = 'New Update Request: ' + requestData.updateType + ' - ' + requestData.updateAction;
    
    // Build email body
    var htmlBody = '<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">';
    htmlBody += '<h2 style="color: #1f2937;">New Update Request Submitted</h2>';
    htmlBody += '<p>A new update request has been submitted to the Cross-Sell Hub database.</p>';
    htmlBody += '<hr style="border: none; border-top: 1px solid #e5e7eb; margin: 20px 0;">';
    
    htmlBody += '<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">';
    htmlBody += '<tr><td style="padding: 8px; font-weight: bold; width: 200px; vertical-align: top;">Request ID:</td><td style="padding: 8px;">' + requestId + '</td></tr>';
    htmlBody += '<tr><td style="padding: 8px; font-weight: bold; vertical-align: top;">Submitted:</td><td style="padding: 8px;">' + formattedDate + '</td></tr>';
    htmlBody += '<tr><td style="padding: 8px; font-weight: bold; vertical-align: top;">Submitter Name:</td><td style="padding: 8px;">' + escapeHtml(requestData.submitterName) + '</td></tr>';
    htmlBody += '<tr><td style="padding: 8px; font-weight: bold; vertical-align: top;">Submitter Email:</td><td style="padding: 8px;">' + escapeHtml(requestData.submitterEmail) + '</td></tr>';
    htmlBody += '<tr><td style="padding: 8px; font-weight: bold; vertical-align: top;">Update Type:</td><td style="padding: 8px;">' + escapeHtml(requestData.updateType) + '</td></tr>';
    htmlBody += '<tr><td style="padding: 8px; font-weight: bold; vertical-align: top;">Action:</td><td style="padding: 8px;">' + escapeHtml(requestData.updateAction) + '</td></tr>';
    htmlBody += '<tr><td style="padding: 8px; font-weight: bold; vertical-align: top;">Target Sheet:</td><td style="padding: 8px;">' + escapeHtml(getTargetSheetName(requestData.updateType)) + '</td></tr>';
    
    // Add additional fields if present
    if (requestData.additionalFields && Object.keys(requestData.additionalFields).length > 0) {
      htmlBody += '<tr><td style="padding: 8px; font-weight: bold; vertical-align: top;">Additional Fields:</td><td style="padding: 8px;">';
      for (var key in requestData.additionalFields) {
        htmlBody += '<strong>' + escapeHtml(key) + ':</strong> ' + escapeHtml(String(requestData.additionalFields[key])) + '<br>';
      }
      htmlBody += '</td></tr>';
    }
    
    htmlBody += '<tr><td style="padding: 8px; font-weight: bold; vertical-align: top;">Description:</td><td style="padding: 8px; white-space: pre-wrap;">' + escapeHtml(requestData.description) + '</td></tr>';
    
    if (requestData.justification && requestData.justification.trim() !== '') {
      htmlBody += '<tr><td style="padding: 8px; font-weight: bold; vertical-align: top;">Justification:</td><td style="padding: 8px; white-space: pre-wrap;">' + escapeHtml(requestData.justification) + '</td></tr>';
    }
    
    htmlBody += '</table>';
    
    htmlBody += '<hr style="border: none; border-top: 1px solid #e5e7eb; margin: 20px 0;">';
    htmlBody += '<p style="color: #6b7280; font-size: 0.9em;">This is an automated notification from the Axiom GRC Cross-Sell Hub.</p>';
    htmlBody += '</div>';
    
    // Plain text version for email clients that don't support HTML
    var plainBody = 'New Update Request Submitted\n\n';
    plainBody += 'Request ID: ' + requestId + '\n';
    plainBody += 'Submitted: ' + formattedDate + '\n';
    plainBody += 'Submitter Name: ' + requestData.submitterName + '\n';
    plainBody += 'Submitter Email: ' + requestData.submitterEmail + '\n';
    plainBody += 'Update Type: ' + requestData.updateType + '\n';
    plainBody += 'Action: ' + requestData.updateAction + '\n';
    plainBody += 'Target Sheet: ' + getTargetSheetName(requestData.updateType) + '\n';
    if (requestData.additionalFields && Object.keys(requestData.additionalFields).length > 0) {
      plainBody += '\nAdditional Fields:\n';
      for (var key in requestData.additionalFields) {
        plainBody += '  ' + key + ': ' + String(requestData.additionalFields[key]) + '\n';
      }
    }
    plainBody += '\nDescription:\n' + requestData.description + '\n';
    if (requestData.justification && requestData.justification.trim() !== '') {
      plainBody += '\nJustification:\n' + requestData.justification + '\n';
    }
    
    // Send email to all stakeholders
    MailApp.sendEmail({
      to: stakeholderEmails.join(','),
      subject: subject,
      htmlBody: htmlBody,
      body: plainBody,
      name: 'Axiom Xplore Update Request',  // Custom sender display name
      replyTo: requestData.submitterEmail   // Replies go to the submitter
    });
    
    Logger.log('Email notification sent to ' + stakeholderEmails.length + ' stakeholder(s) for request ' + requestId);
    
  } catch (error) {
    Logger.log('ERROR in sendUpdateRequestNotification: ' + error.toString());
    throw error;
  }
}

/**
 * Helper function to escape HTML special characters
 */
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function clearCache() {
  var cache = CacheService.getScriptCache();
  cache.remove('all_products');
  cache.remove('product_details');
  cache.remove('companies');
  cache.remove('rules');
  cache.remove('glossary');
  
  // Also clear any config caches
  cache.remove('config_CACHE_DURATION_SECONDS');
  cache.remove('config_ENABLE_CACHING');
  cache.remove('config_CACHE_KEY_PREFIX');
  cache.remove('config_UPDATE_REQUEST_NOTIFICATION_EMAILS');
  
  return 'Cache cleared!';
}

/**
 * Get company information for About Us page
 */
function getCompanyInfo() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('COMPANY_INFO');
    
    if (!sheet) {
      Logger.log('WARNING: COMPANY_INFO sheet not found');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var companies = [];
    for (var i = 1; i < data.length; i++) {
      var company = {};
      for (var j = 0; j < headers.length; j++) {
        var value = data[i][j];
        
        // Auto-convert LogoURL if it's a raw Drive link
        if (headers[j] === 'LogoURL' && value) {
          company[headers[j]] = convertDriveLinkToDirectURL(value);
        } 
        // Auto-convert VideoURL if it's a Google Drive link
        else if (headers[j] === 'VideoURL' && value) {
          company[headers[j]] = convertDriveVideoToEmbedURL(value);
        } 
        else {
          company[headers[j]] = value;
        }
      }
      companies.push(company);
    }
    
    // Keep original spreadsheet row order - no sorting
    // This ensures both the about page and slides appear in database order
    Logger.log('Using original spreadsheet row order for companies');
    
    return companies;
    
  } catch (error) {
    Logger.log('ERROR in getCompanyInfo: ' + error.toString());
    return [];
  }
}

/**
 * Convert any Google Drive link format to direct image URL
 * Updated to use thumbnail format which works better with CORS
 */
function convertDriveLinkToDirectURL(url) {
  if (!url || typeof url !== 'string') return url;
  
  // Already in direct format
  if (url.includes('drive.google.com/uc?export=view&id=') || 
      url.includes('drive.google.com/thumbnail?id=')) {
    return url;
  }
  
  // Extract file ID from various Drive URL formats
  var fileId = null;
  
  // Format: https://drive.google.com/file/d/FILE_ID/view
  var match = url.match(/\/file\/d\/([^\/\?]+)/);
  if (match) {
    fileId = match[1];
  }
  
  // Format: https://drive.google.com/open?id=FILE_ID
  if (!fileId) {
    match = url.match(/[?&]id=([^&]+)/);
    if (match) {
      fileId = match[1];
    }
  }
  
  // If we found a file ID, convert to thumbnail URL (better for web embedding)
  if (fileId) {
    return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';
  }
  
  // Return original URL if we couldn't parse it
  return url;
}

/**
 * Convert Google Drive video link to embeddable preview URL
 * Only converts Drive links, leaves YouTube/Vimeo/etc. unchanged
 */
function convertDriveVideoToEmbedURL(url) {
  if (!url || typeof url !== 'string') return url;
  
  // Only process Google Drive URLs
  if (!url.includes('drive.google.com')) {
    return url; // Return YouTube, Vimeo, etc. as-is
  }
  
  // Already in preview format
  if (url.includes('/preview')) {
    return url;
  }
  
  // Extract file ID from various Drive URL formats
  var fileId = null;
  
  // Format: https://drive.google.com/file/d/FILE_ID/view?usp=drive_link
  // Format: https://drive.google.com/file/d/FILE_ID/view
  var match = url.match(/\/file\/d\/([^\/\?]+)/);
  if (match) {
    fileId = match[1];
  }
  
  // Format: https://drive.google.com/open?id=FILE_ID
  if (!fileId) {
    match = url.match(/[?&]id=([^&]+)/);
    if (match) {
      fileId = match[1];
    }
  }
  
  // If we found a file ID, convert to preview URL (embeddable format)
  if (fileId) {
    return 'https://drive.google.com/file/d/' + fileId + '/preview?usp=sharing';
  }
  
  // Return original URL if we couldn't parse it
  return url;
}

/**
 * Get contacts (with optional ProductID filter)
 */
function getContacts(productId) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Contacts');
    
    if (!sheet) {
      Logger.log('WARNING: Contacts sheet not found');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      // Empty or only headers
      return [];
    }
    
    var headers = data[0];
    var contacts = [];
    
    for (var i = 1; i < data.length; i++) {
      var contact = {};
      for (var j = 0; j < headers.length; j++) {
        contact[headers[j]] = data[i][j];
      }
      
      // If productId specified, filter by it; otherwise return all
      if (!productId || contact.ProductID === productId) {
        contacts.push(contact);
      }
    }
    
    return contacts;
    
  } catch (error) {
    Logger.log('ERROR in getContacts: ' + error.toString());
    return [];
  }
}

/**
 * Get resources (with optional Category filter)
 */
function getResources(category) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Resources');
    
    if (!sheet) {
      Logger.log('WARNING: Resources sheet not found');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      // Empty or only headers
      return [];
    }
    
    var headers = data[0];
    var resources = [];
    
    for (var i = 1; i < data.length; i++) {
      var resource = {};
      for (var j = 0; j < headers.length; j++) {
        resource[headers[j]] = data[i][j];
      }
      
      // If category specified, filter by it; otherwise return all
      if (!category || resource.Category === category) {
        resources.push(resource);
      }
    }
    
    return resources;
    
  } catch (error) {
    Logger.log('ERROR in getResources: ' + error.toString());
    return [];
  }
}

/**
 * Get glossary terms
 */
function getGlossary() {
  try {
    var enableCaching = getConfig('ENABLE_CACHING', 'TRUE');
    var cache = CacheService.getScriptCache();
    var cached = null;
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      cached = cache.get('glossary');
      if (cached) {
        return JSON.parse(cached);
      }
    }
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Glossary');
    
    if (!sheet) {
      Logger.log('WARNING: Glossary sheet not found');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    
    var headers = data[0];
    var glossary = [];
    
    for (var i = 1; i < data.length; i++) {
      var term = {};
      for (var j = 0; j < headers.length; j++) {
        term[headers[j]] = data[i][j] || '';
      }
      glossary.push(term);
    }
    
    if (enableCaching === 'TRUE' || enableCaching === true) {
      var cacheDuration = parseInt(getConfig('CACHE_DURATION_SECONDS', 600));
      cache.put('glossary', JSON.stringify(glossary), cacheDuration);
    }
    
    return glossary;
    
  } catch (error) {
    Logger.log('ERROR in getGlossary: ' + error.toString());
    return [];
  }
}

/**
 * Find cross-sell opportunities for a client
 * Returns opportunities sorted by data source priority (real,expert > real > expert > inferred), then by client count
 */
function findCrossSellOpportunities(clientData) {
  try {
    var opportunities = generateRecommendations(clientData);
    
    // Return whatever opportunities were found (no minimum requirement)
    return opportunities;
    
  } catch (error) {
    Logger.log('ERROR in findCrossSellOpportunities: ' + error.toString());
    throw error;
  }
}

// ============================================
// EXPORT TO GOOGLE SLIDES FUNCTIONALITY
// ============================================

/**
 * Export to Google Slides
 * @param {Object} options - Export options
 * @param {boolean} options.includeOverview - Include company overviews
 * @param {boolean} options.includeProducts - Include product library
 * @returns {Object} Result with success status and presentation URL
 */
function exportToSlides(options) {
  try {
    // Get template ID from CONFIG sheet, fallback to default
    var templateId = getConfig('SLIDES_TEMPLATE_ID', '16Rwh50bq1lWbvtMCSYaE03rQBu6gigLHX7u5w01jLH0');
    
    if (!templateId || templateId === 'YOUR_TEMPLATE_PRESENTATION_ID_HERE') {
      throw new Error('Slides template ID not configured. Please set SLIDES_TEMPLATE_ID in CONFIG sheet.');
    }
    
    // Validate template ID format (Google Slides IDs are typically 44 characters)
    templateId = String(templateId).trim();
    if (templateId.length < 20) {
      throw new Error('Invalid template ID format. Please check your SLIDES_TEMPLATE_ID in CONFIG sheet.');
    }
    
    // Try to access the template file first
    var templateFile;
    try {
      templateFile = DriveApp.getFileById(templateId);
      // Verify it's actually a Google Slides file
      if (templateFile.getMimeType() !== 'application/vnd.google-apps.presentation') {
        throw new Error('The template ID does not point to a Google Slides presentation. Please verify your SLIDES_TEMPLATE_ID.');
      }
    } catch (driveError) {
      Logger.log('DriveApp error: ' + driveError.toString());
      // If DriveApp fails, try using SlidesApp directly to verify access
      try {
        var testPresentation = SlidesApp.openById(templateId);
        Logger.log('Template accessible via SlidesApp');
        // If SlidesApp works but DriveApp doesn't, we can still proceed
        // Let's try to get file via DriveApp again with better error message
        throw new Error('Cannot access template via DriveApp. Please ensure: 1) Drive permissions are granted, 2) The template file exists, 3) You have access to the file. Try running a function that uses DriveApp first to trigger permissions. Original error: ' + driveError.toString());
      } catch (slidesError) {
        throw new Error('Cannot access template file. Please verify: 1) The SLIDES_TEMPLATE_ID is correct, 2) The template file exists, 3) You have access to the file, 4) Drive permissions are granted. Original error: ' + driveError.toString());
      }
    }
    
    // Copy the template using DriveApp
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var newPresentation;
    try {
      newPresentation = templateFile.makeCopy('Axiom GRC Overview - ' + timestamp);
    } catch (copyError) {
      throw new Error('Failed to copy template. Please ensure you have permission to copy the file. Error: ' + copyError.toString());
    }
    
    var presentationId = newPresentation.getId();
    var presentation = SlidesApp.openById(presentationId);
    
    var slides = presentation.getSlides();
    
    if (slides.length < 3) {
      throw new Error('Template must have at least 3 slides (title, company template, product template, and optionally opportunity template)');
    }
    
    // Update title slide (Slide 0)
    replaceTextInSlide(slides[0], '{{COMPANY_NAME}}', 'Axiom GRC');
    
    var currentSlideIndex = 1; // Track where we are after removing templates
    
    // ====================================
    // ADD COMPANY OVERVIEWS
    // ====================================
    if (options.includeOverview) {
      var companies = getCompanyInfo();
      if (companies.length > 0) { // Only add if there's content
        // Reverse the array since duplicate() adds slides right after template, which reverses order
        companies.reverse();
        var companyTemplateSlide = slides[currentSlideIndex];
        
        for (var i = 0; i < companies.length; i++) {
          var company = companies[i];
          
          // Duplicate the company template slide (automatically adds after the original)
          var newSlide = companyTemplateSlide.duplicate();
          
          // Replace text placeholders
          replaceTextInSlide(newSlide, '{{COMPANY_NAME}}', company.CompanyName || '');
          replaceTextInSlide(newSlide, '{{ABOUT_TEXT}}', company.AboutText || 'Company information coming soon...');
          
          // Add company logo
          if (company.LogoURL) {
            try {
              addImageToSlide(newSlide, company.LogoURL, '{{COMPANY_LOGO}}');
            } catch (e) {
              Logger.log('Error adding logo for ' + company.CompanyName + ': ' + e);
              replaceTextInSlide(newSlide, '{{COMPANY_LOGO}}', '');
            }
          } else {
            replaceTextInSlide(newSlide, '{{COMPANY_LOGO}}', '');
          }
          
          // Add video
          if (company.VideoURL) {
            try {
              addVideoToSlide(newSlide, company.VideoURL, '{{COMPANY_VIDEO}}');
            } catch (e) {
              Logger.log('Error adding video for ' + company.CompanyName + ': ' + e);
              replaceTextInSlide(newSlide, '{{COMPANY_VIDEO}}', '');
            }
          } else {
            replaceTextInSlide(newSlide, '{{COMPANY_VIDEO}}', '');
          }
        }
        
        // Remove the company template slide only if we added content
        companyTemplateSlide.remove();
      } else {
        // No companies, remove template
        slides[currentSlideIndex].remove();
      }
    } else {
      // Remove company template if not including overviews
      slides[currentSlideIndex].remove();
    }
    
    // ====================================
    // ADD PRODUCT LIBRARY
    // ====================================
    if (options.includeProducts) {
      var products = getAllProducts();
      var productDetails = getProductDetails();
      var companyLogos = getCompanyLogoMap();
      
      // Get the product template slide
      var currentSlides = presentation.getSlides();
      var productTemplateSlide = currentSlides[currentSlideIndex];
      
      if (!productTemplateSlide) {
        throw new Error('Product template slide not found');
      }
      
      // Group products by company
      var productsByCompany = {};
      for (var i = 0; i < products.length; i++) {
        var product = products[i];
        if (!productsByCompany[product.CompanyName]) {
          productsByCompany[product.CompanyName] = [];
        }
        productsByCompany[product.CompanyName].push(product);
      }
      
      // Filter by selected companies if provided
      var companyNames = Object.keys(productsByCompany).sort();
      if (options.selectedCompanies && options.selectedCompanies.length > 0) {
        // Filter to only include selected companies
        var selectedCompaniesSet = {};
        for (var s = 0; s < options.selectedCompanies.length; s++) {
          selectedCompaniesSet[options.selectedCompanies[s]] = true;
        }
        companyNames = companyNames.filter(function(companyName) {
          return selectedCompaniesSet[companyName] === true;
        });
      }
      
      // Filter products to only those from selected companies
      var filteredProducts = [];
      for (var c = 0; c < companyNames.length; c++) {
        var companyName = companyNames[c];
        var companyProducts = productsByCompany[companyName];
        for (var p = 0; p < companyProducts.length; p++) {
          filteredProducts.push(companyProducts[p]);
        }
      }
      
      if (filteredProducts.length > 0) { // Only add if there's content
        // Add products for each company
        for (var c = 0; c < companyNames.length; c++) {
          var companyName = companyNames[c];
          var companyProducts = productsByCompany[companyName];
          
          for (var p = 0; p < companyProducts.length; p++) {
            var product = companyProducts[p];
            
            // Duplicate product template (automatically adds after the template slide)
            var newSlide = productTemplateSlide.duplicate();
            
            // Get full description from product_details if available
            var detailRecord = null;
            for (var d = 0; d < productDetails.length; d++) {
              if (productDetails[d].ProductID === product.ProductID) {
                detailRecord = productDetails[d];
                break;
              }
            }
            
            var fullDescription = (detailRecord && detailRecord.FullDescription) 
              ? detailRecord.FullDescription 
              : (product.ShortDescription || 'No description available');
            
            // Collect key features
            var keyFeatures = [];
            for (var f = 1; f <= 10; f++) {
              var feature = product['KeyFeature' + f] || (detailRecord && detailRecord['KeyFeature' + f]);
              if (feature) {
                keyFeatures.push(feature);
              }
            }
            var keyFeaturesText = keyFeatures.slice(0, 5).join('\n'); // Max 5 features for slide
            
            // Replace placeholders
            replaceTextInSlide(newSlide, '{{PRODUCT_NAME}}', product.ProductName || '');
            replaceTextInSlide(newSlide, '{{COMPANY_NAME}}', product.CompanyName || '');
            replaceTextInSlide(newSlide, '{{PRODUCT_DESCRIPTION}}', fullDescription);
            replaceTextInSlide(newSlide, '{{KEY_FEATURES}}', keyFeaturesText || 'Contact for more information');
            
            // Add company logo (small)
            var logoURL = companyLogos[product.CompanyName] || product.CompanyLogoURL || '';
            if (logoURL) {
              try {
                addImageToSlide(newSlide, logoURL, '{{COMPANY_LOGO}}');
              } catch (e) {
                Logger.log('Error adding product company logo: ' + e);
                replaceTextInSlide(newSlide, '{{COMPANY_LOGO}}', '');
              }
            } else {
              replaceTextInSlide(newSlide, '{{COMPANY_LOGO}}', '');
            }
          }
        }
        
        // Remove the product template slide only if we added content
        productTemplateSlide.remove();
      } else {
        // No products, remove template
        productTemplateSlide.remove();
      }
    } else {
      // Remove product template if not including products
      var currentSlidesAfterCheck = presentation.getSlides();
      if (currentSlidesAfterCheck.length > currentSlideIndex) {
        currentSlidesAfterCheck[currentSlideIndex].remove();
      }
    }
    
    // ====================================
    // ADD OPPORTUNITIES (NEW)
    // ====================================
    if (options.includeOpportunities && options.opportunities && options.opportunities.length > 0) {
      var currentSlidesAfterProducts = presentation.getSlides();
      var companyLogos = getCompanyLogoMap();
      
      // Find opportunity template slide by looking for placeholder text
      var opportunityTemplateSlide = null;
      var opportunityTemplateIndex = -1;
      
      for (var i = 0; i < currentSlidesAfterProducts.length; i++) {
        var slide = currentSlidesAfterProducts[i];
        var shapes = slide.getShapes();
        for (var j = 0; j < shapes.length; j++) {
          try {
            var text = shapes[j].getText().asString();
            if (text.indexOf('{{OPPORTUNITY_PRODUCT_NAME}}') !== -1 || 
                text.indexOf('{{OPPORTUNITY_RANK}}') !== -1) {
              opportunityTemplateSlide = slide;
              opportunityTemplateIndex = i;
              break;
            }
          } catch (e) {
            // Continue
          }
        }
        if (opportunityTemplateSlide) break;
      }
      
      // Also check page elements if not found in shapes
      if (!opportunityTemplateSlide) {
        for (var i = 0; i < currentSlidesAfterProducts.length; i++) {
          try {
            var pageElements = currentSlidesAfterProducts[i].getPageElements();
            for (var p = 0; p < pageElements.length; p++) {
              try {
                var element = pageElements[p];
                if (element.getPageElementType() === 'SHAPE') {
                  var shape = element.asShape();
                  var text = shape.getText().asString();
                  if (text.indexOf('{{OPPORTUNITY_PRODUCT_NAME}}') !== -1 || 
                      text.indexOf('{{OPPORTUNITY_RANK}}') !== -1) {
                    opportunityTemplateSlide = currentSlidesAfterProducts[i];
                    opportunityTemplateIndex = i;
                    break;
                  }
                }
              } catch (e) {
                // Continue
              }
            }
            if (opportunityTemplateSlide) break;
          } catch (e) {
            // Continue
          }
        }
      }
      
      if (opportunityTemplateSlide) {
        // Reverse the loop to fix rank ordering (duplicate() adds slides right after template, reversing order)
        for (var i = options.opportunities.length - 1; i >= 0; i--) {
          var opp = options.opportunities[i];
          var product = opp.product;
          var newSlide = opportunityTemplateSlide.duplicate();
          
          // Replace basic placeholders - rank is based on original index (i+1), but we're looping backwards
          var rank = i + 1;
          replaceTextInSlide(newSlide, '{{OPPORTUNITY_RANK}}', rank.toString());
          
          // PRODUCT_NAME should show current products (selected products), not the opportunity product
          var currentProductsText = '';
          if (options.currentProducts && options.currentProducts.length > 0) {
            // Get product names from product IDs
            var allProducts = getAllProducts();
            var productNames = [];
            for (var cp = 0; cp < options.currentProducts.length; cp++) {
              for (var ap = 0; ap < allProducts.length; ap++) {
                if (allProducts[ap].ProductID === options.currentProducts[cp]) {
                  productNames.push(allProducts[ap].ProductName);
                  break;
                }
              }
            }
            currentProductsText = productNames.join(', ');
          }
          // Only replace if there are current products, otherwise remove placeholder
          if (currentProductsText) {
            replaceTextInSlide(newSlide, '{{PRODUCT_NAME}}', currentProductsText);
          } else {
            replaceTextInSlide(newSlide, '{{PRODUCT_NAME}}', '');
          }
          replaceTextInSlide(newSlide, '{{OPPORTUNITY_PRODUCT_NAME}}', product.ProductName || '');
          replaceTextInSlide(newSlide, '{{OPPORTUNITY_COMPANY_NAME}}', product.CompanyName || '');
          replaceTextInSlide(newSlide, '{{OPPORTUNITY_DESCRIPTION}}', product.ShortDescription || '');
          
          // DataSource Badge - always check dataSource separately (expert, real, etc.)
          var dataSourceBadges = [];
          var dataSources = String(opp.dataSource || '').split(/[,\s]+/).filter(function(ds) { 
            return ds.trim(); 
          });
          
          dataSources.forEach(function(ds) {
            var upperDs = ds.trim().toUpperCase();
            if (upperDs === 'EXPERT') {
              dataSourceBadges.push('Recommended by our cross-sell experts');
            } else if (upperDs === 'REAL' && opp.clientCount && parseInt(opp.clientCount) > 0) {
              var count = parseInt(opp.clientCount);
              var clientWord = count === 1 ? 'client' : 'clients';
              dataSourceBadges.push(count + ' ' + clientWord + ' use both products!');
            } else if (upperDs === 'INFERRED') {
              dataSourceBadges.push('AI-Inferred pairing');
            }
          });
          
          // Set dataSource badge placeholder (separate from insights)
          if (dataSourceBadges.length > 0) {
            replaceTextInSlide(newSlide, '{{OPPORTUNITY_DATASOURCE_BADGE}}', dataSourceBadges.join(', '));
          } else {
            replaceTextInSlide(newSlide, '{{OPPORTUNITY_DATASOURCE_BADGE}}', '');
          }
          
          // Insights Source - only use insightsSource field, separate from dataSource
          var insightsSourceText = '';
          var hasInsightsSource = false;
          
          if (opp.insightsSource && opp.insightsSource.trim()) {
            insightsSourceText = opp.insightsSource;
            hasInsightsSource = true;
          }
          
          // If insights source exists, format as "Why this recommendation? [insightsSource]"
          if (hasInsightsSource && insightsSourceText) {
            replaceTextInSlide(newSlide, '{{OPPORTUNITY_INSIGHTSSOURCE}}', 'Why this recommendation? ' + insightsSourceText);
          } else {
            // Remove if no insights source
            replaceTextInSlide(newSlide, '{{OPPORTUNITY_INSIGHTSSOURCE}}', '');
          }
          
          // Rationale - only show if it exists
          if (opp.rationale && opp.rationale.trim()) {
            replaceTextInSlide(newSlide, '{{OPPORTUNITY_RATIONALE}}', opp.rationale);
            // Also replace "Why this recommendation?" text placeholder if it exists
            replaceTextInSlide(newSlide, '{{WHY_RECOMMENDATION}}', 'Why this recommendation?');
          } else {
            // Remove rationale section if no rationale
            replaceTextInSlide(newSlide, '{{WHY_RECOMMENDATION}}', '');
            replaceTextInSlide(newSlide, 'Why this recommendation?', '');
            replaceTextInSlide(newSlide, '{{OPPORTUNITY_RATIONALE}}', '');
          }
          
          // Key benefits - only show if they exist, as bulleted list
          if (opp.keyBenefits && opp.keyBenefits.length > 0) {
            var keyBenefitsText = opp.keyBenefits
              .filter(function(b) { return b && b.trim(); })
              .map(function(b) { return '• ' + b; })
              .join('\n');
            replaceTextInSlide(newSlide, '{{OPPORTUNITY_KEY_BENEFITS}}', keyBenefitsText);
            // Also replace "Key benefits:" text placeholder if it exists
            replaceTextInSlide(newSlide, '{{KEY_BENEFITS_LABEL}}', 'Key benefits:');
          } else {
            // Remove key benefits section if none exist
            replaceTextInSlide(newSlide, '{{KEY_BENEFITS_LABEL}}', '');
            replaceTextInSlide(newSlide, 'Key benefits:', '');
            replaceTextInSlide(newSlide, '{{OPPORTUNITY_KEY_BENEFITS}}', '');
          }
          
          // Client count - format as "X client/clients have both products"
          var clientCountText = '';
          if (opp.clientCount && parseInt(opp.clientCount) > 0) {
            var count = parseInt(opp.clientCount);
            var clientWord = count === 1 ? 'client' : 'clients';
            clientCountText = count + ' ' + clientWord + ' have both products';
          }
          replaceTextInSlide(newSlide, '{{OPPORTUNITY_CLIENT_COUNT}}', clientCountText);
          
          // Add company logo if available
          var logoURL = companyLogos[product.CompanyName] || product.CompanyLogoURL || '';
          if (logoURL) {
            try {
              addImageToSlide(newSlide, logoURL, '{{OPPORTUNITY_COMPANY_LOGO}}');
            } catch (e) {
              Logger.log('Error adding opportunity logo: ' + e);
              replaceTextInSlide(newSlide, '{{OPPORTUNITY_COMPANY_LOGO}}', '');
            }
          } else {
            replaceTextInSlide(newSlide, '{{OPPORTUNITY_COMPANY_LOGO}}', '');
          }
          
          // ====================================
          // ADD LINKS (URLs included in text since programmatic links don't work)
          // ====================================
          
          // Demo links - add "See Demo" with URL in text
          if (product.DemoURL1) {
            replaceTextInSlide(newSlide, '{{DEMO1_PLACEHOLDER}}', 'See Demo (' + product.DemoURL1 + ')');
            
            // If there are multiple demos, add them too
            if (product.DemoURL2) {
              replaceTextInSlide(newSlide, '{{DEMO2_PLACEHOLDER}}', 'See Demo 2 (' + product.DemoURL2 + ')');
            }
            if (product.DemoURL3) {
              replaceTextInSlide(newSlide, '{{DEMO3_PLACEHOLDER}}', 'See Demo 3 (' + product.DemoURL3 + ')');
            }
          } else if (product.DemoURL2) {
            // If DemoURL1 doesn't exist but DemoURL2 does, use that as primary
            replaceTextInSlide(newSlide, '{{DEMO1_PLACEHOLDER}}', 'See Demo (' + product.DemoURL2 + ')');
            
            if (product.DemoURL3) {
              replaceTextInSlide(newSlide, '{{DEMO2_PLACEHOLDER}}', 'See Demo 2 (' + product.DemoURL3 + ')');
            }
          } else if (product.DemoURL3) {
            replaceTextInSlide(newSlide, '{{DEMO1_PLACEHOLDER}}', 'See Demo (' + product.DemoURL3 + ')');
          } else {
            // No demos - remove placeholder
            replaceTextInSlide(newSlide, '{{DEMO1_PLACEHOLDER}}', '');
            replaceTextInSlide(newSlide, '{{DEMO2_PLACEHOLDER}}', '');
            replaceTextInSlide(newSlide, '{{DEMO3_PLACEHOLDER}}', '');
          }
          
          // Learn More link
          if (product.LearnMoreURL) {
            replaceTextInSlide(newSlide, '{{LEARN_MORE_PLACEHOLDER}}', 'Learn More (' + product.LearnMoreURL + ')');
          } else {
            // No Learn More - remove placeholder
            replaceTextInSlide(newSlide, '{{LEARN_MORE_PLACEHOLDER}}', '');
          }
          
          // Submit Cross-Sell Referral link
          // Get company-specific URL only (no general fallback)
          var crossSellLeadUrl = '';
          try {
            var companies = getCompanies();
            for (var c = 0; c < companies.length; c++) {
              if (companies[c].CompanyName === product.CompanyName) {
                crossSellLeadUrl = companies[c].CrossSellLeadUrl || companies[c].crossSellLeadUrl || companies[c].CrossSellLeadURL || '';
                break;
              }
            }
          } catch (e) {
            Logger.log('Error getting cross-sell lead URL: ' + e);
            crossSellLeadUrl = '';
          }
          
          if (crossSellLeadUrl && crossSellLeadUrl.trim() !== '') {
            replaceTextInSlide(newSlide, '{{SUBMIT_REFERRAL_PLACEHOLDER}}', 'Submit cross-sell referral (' + crossSellLeadUrl + ')');
          } else {
            // No referral URL - remove placeholder
            replaceTextInSlide(newSlide, '{{SUBMIT_REFERRAL_PLACEHOLDER}}', '');
          }
        }
        
        // Remove template
        opportunityTemplateSlide.remove();
      } else {
        Logger.log('WARNING: Opportunity template slide not found in presentation');
      }
    } else if (options.includeOpportunities) {
      // Opportunities requested but none found - try to remove template if it exists
      var currentSlidesFinal = presentation.getSlides();
      for (var i = 0; i < currentSlidesFinal.length; i++) {
        var slide = currentSlidesFinal[i];
        var shapes = slide.getShapes();
        for (var j = 0; j < shapes.length; j++) {
          try {
            var text = shapes[j].getText().asString();
            if (text.indexOf('{{OPPORTUNITY_') !== -1) {
              slide.remove();
              break;
            }
          } catch (e) {
            // Continue
          }
        }
      }
    }
    
    // Automatically share presentation with "Anyone with the link can edit"
    try {
      var presentationFile = DriveApp.getFileById(presentationId);
      presentationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      Logger.log('Presentation automatically shared with "Anyone with the link can edit"');
    } catch (shareError) {
      Logger.log('Warning: Could not set sharing permissions: ' + shareError.toString());
      // Continue even if sharing fails - presentation is still created
    }
    
    return {
      success: true,
      url: 'https://docs.google.com/presentation/d/' + presentationId,
      id: presentationId
    };
    
  } catch (error) {
    Logger.log('Error in exportToSlides: ' + error);
    throw new Error('Failed to create presentation: ' + error.message);
  }
}

// ====================================
// HELPER FUNCTIONS FOR SLIDES
// ====================================

function replaceTextInSlide(slide, placeholder, value) {
  try {
    var replaced = false;
    
    // Replace in all text elements - check shapes more thoroughly
    var shapes = slide.getShapes();
    for (var i = 0; i < shapes.length; i++) {
      try {
        var shape = shapes[i];
        var textRange = shape.getText();
        var textContent = textRange.asString();
        
        if (textContent.indexOf(placeholder) !== -1) {
          // Replace all occurrences
          textRange.replaceAllText(placeholder, value || '');
          replaced = true;
          Logger.log('Replaced ' + placeholder + ' in shape at index ' + i);
        }
      } catch (e) {
        // Shape might not have text or might be grouped - try to check if it's a group
        try {
          // Try to get grouped shapes
          var shape = shapes[i];
          if (shape.getShapeType && shape.getShapeType() === 'GROUP') {
            // Group shapes don't support getText directly - skip for now
            // Individual shapes in the group will be processed if they're also in the shapes array
          }
        } catch (groupError) {
          // Not a group or other error - continue
        }
      }
    }
    
    // Also check page elements (like text boxes that might be separate)
    try {
      var pageElements = slide.getPageElements();
      for (var p = 0; p < pageElements.length; p++) {
        try {
          var element = pageElements[p];
          if (element.getPageElementType() === 'SHAPE') {
            var shape = element.asShape();
            var textRange = shape.getText();
            var textContent = textRange.asString();
            if (textContent.indexOf(placeholder) !== -1) {
              textRange.replaceAllText(placeholder, value || '');
              replaced = true;
              Logger.log('Replaced ' + placeholder + ' in page element at index ' + p);
            }
          }
        } catch (e) {
          // Element might not have text
        }
      }
    } catch (e) {
      Logger.log('Error checking page elements: ' + e);
    }
    
    // Also check tables
    var tables = slide.getTables();
    for (var t = 0; t < tables.length; t++) {
      var table = tables[t];
      var numRows = table.getNumRows();
      var numCols = table.getNumColumns();
      for (var i = 0; i < numRows; i++) {
        for (var j = 0; j < numCols; j++) {
          try {
            var cell = table.getCell(i, j);
            var textRange = cell.getText();
            var textContent = textRange.asString();
            if (textContent.indexOf(placeholder) !== -1) {
              textRange.replaceAllText(placeholder, value || '');
              replaced = true;
              Logger.log('Replaced ' + placeholder + ' in table cell [' + i + ',' + j + ']');
            }
          } catch (e) {
            // Cell might not exist
          }
        }
      }
    }
    
    if (!replaced) {
      Logger.log('WARNING: Placeholder ' + placeholder + ' not found in any shape, page element, or table');
    }
  } catch (e) {
    Logger.log('Error in replaceTextInSlide for placeholder ' + placeholder + ': ' + e);
  }
}

function addImageToSlide(slide, imageUrl, placeholder) {
  if (!imageUrl || !placeholder) return;
  
  // Find placeholder shape
  var shapes = slide.getShapes();
  var placeholderShape = null;
  
  for (var i = 0; i < shapes.length; i++) {
    try {
      var text = shapes[i].getText().asString();
      if (text.indexOf(placeholder) !== -1) {
        placeholderShape = shapes[i];
        break;
      }
    } catch (e) {
      // Shape might not have text
    }
  }
  
  if (!placeholderShape) {
    Logger.log('Image placeholder not found: ' + placeholder);
    return;
  }
  
  var left = placeholderShape.getLeft();
  var top = placeholderShape.getTop();
  var width = placeholderShape.getWidth();
  var height = placeholderShape.getHeight();
  
  // Remove placeholder shape
  placeholderShape.remove();
  
  // Add image
  try {
    // Convert Google Drive link to direct URL if needed
    var directUrl = convertDriveLinkToDirectURL(imageUrl);
    var response = UrlFetchApp.fetch(directUrl);
    var blob = response.getBlob();
    slide.insertImage(blob, left, top, width, height);
  } catch (e) {
    Logger.log('Error fetching/inserting image from ' + imageUrl + ': ' + e);
  }
}

/**
 * Enhanced helper function to create clickable links in slides
 * Uses Google Slides API - links are added to text runs within shapes
 */
function addLinkToTextInSlide(slide, searchText, url) {
  if (!searchText || !url) return false;
  
  try {
    var linked = false;
    
    // Function to add link to text in a shape
    // Note: Google Slides API link support may require Advanced Slides Service
    // This function tries multiple approaches
    function addLinkInShape(shape, searchText, url) {
      try {
        var textRange = shape.getText();
        var fullText = textRange.asString();
        var index = fullText.indexOf(searchText);
        
        if (index === -1) return false;
        
        // Get the range of the text we want to link
        var range = textRange.getRange(index, index + searchText.length);
        
        // Try multiple methods to add link
        // Method 1: Try getLink() then setUrl() - this is the documented API
        try {
          // Check if getLink exists and is a function
          if (range.getLink && typeof range.getLink === 'function') {
            var link = range.getLink();
            if (link && typeof link.setUrl === 'function') {
              link.setUrl(url);
              Logger.log('Successfully added link using getLink().setUrl()');
              return true;
            }
          }
        } catch (e1) {
          Logger.log('getLink() method failed: ' + e1.toString());
        }
        
        // Method 2: Try setLinkUrl() directly - some API versions might support this
        try {
          if (range.setLinkUrl && typeof range.setLinkUrl === 'function') {
            range.setLinkUrl(url);
            Logger.log('Successfully added link using setLinkUrl()');
            return true;
          }
        } catch (e2) {
          Logger.log('setLinkUrl() method failed: ' + e2.toString());
        }
        
        // Method 3: Workaround - delete and re-insert with link formatting
        // This might work if we can insert text with link properties
        try {
          var textToLink = range.asString();
          range.clear();
          var insertPos = textRange.getRange(index, index);
          insertPos.insertText(textToLink);
          
          // Try to get link on newly inserted text
          var newRange = textRange.getRange(index, index + textToLink.length);
          if (newRange.getLink && typeof newRange.getLink === 'function') {
            var newLink = newRange.getLink();
            if (newLink && typeof newLink.setUrl === 'function') {
              newLink.setUrl(url);
              Logger.log('Successfully added link after re-inserting text');
              return true;
            }
          }
        } catch (e3) {
          Logger.log('Re-insert method failed: ' + e3.toString());
        }
        
        Logger.log('All link methods failed for text: ' + searchText);
        return false;
      } catch (e) {
        Logger.log('Error in addLinkInShape: ' + e.toString());
        return false;
      }
    }
    
    // Check shapes
    var shapes = slide.getShapes();
    for (var i = 0; i < shapes.length; i++) {
      try {
        if (addLinkInShape(shapes[i], searchText, url)) {
          linked = true;
          Logger.log('Added link to: ' + searchText + ' in shape ' + i);
          break; // Only link first occurrence
        }
      } catch (e) {
        Logger.log('Error processing shape ' + i + ': ' + e.toString());
      }
    }
    
    // Also check page elements
    if (!linked) {
      try {
        var pageElements = slide.getPageElements();
        for (var p = 0; p < pageElements.length; p++) {
          try {
            var element = pageElements[p];
            if (element.getPageElementType() === 'SHAPE') {
              var shape = element.asShape();
              if (addLinkInShape(shape, searchText, url)) {
                linked = true;
                Logger.log('Added link to page element: ' + searchText);
                break;
              }
            }
          } catch (e) {
            // Continue
          }
        }
      } catch (e) {
        Logger.log('Error checking page elements for links: ' + e);
      }
    }
    
    // Also check tables
    if (!linked) {
      try {
        var tables = slide.getTables();
        for (var t = 0; t < tables.length; t++) {
          var table = tables[t];
          var numRows = table.getNumRows();
          var numCols = table.getNumColumns();
          for (var r = 0; r < numRows; r++) {
            for (var c = 0; c < numCols; c++) {
              try {
                var cell = table.getCell(r, c);
                var cellTextRange = cell.getText();
                var cellText = cellTextRange.asString();
                
                if (cellText.indexOf(searchText) !== -1) {
                  var index = cellText.indexOf(searchText);
                  var range = cellTextRange.getRange(index, index + searchText.length);
                  try {
                    var link = range.getLink();
                    link.setUrl(url);
                    linked = true;
                    break;
                  } catch (e) {
                    // Try alternative
                    try {
                      range.setLinkUrl(url);
                      linked = true;
                      break;
                    } catch (e2) {
                      // Skip
                    }
                  }
                }
              } catch (e) {
                // Cell might not support links
              }
            }
            if (linked) break;
          }
          if (linked) break;
        }
      } catch (e) {
        Logger.log('Error checking tables for links: ' + e);
      }
    }
    
    return linked;
  } catch (e) {
    Logger.log('Error in addLinkToTextInSlide: ' + e);
    return false;
  }
}

function addVideoToSlide(slide, videoUrl, placeholder) {
  if (!videoUrl || !placeholder) return;
  
  Logger.log('Looking for video placeholder: ' + placeholder);
  
  // Find placeholder shape/element FIRST - before doing anything else
  var shapes = slide.getShapes();
  var placeholderShape = null;
  var placeholderElement = null;
  
  // First pass: look for shapes containing the placeholder text
  for (var i = 0; i < shapes.length; i++) {
    try {
      var text = shapes[i].getText().asString();
      if (text.indexOf(placeholder) !== -1) {
        placeholderShape = shapes[i];
        Logger.log('Found video placeholder in shape at index ' + i + ': "' + text + '"');
        break;
      }
    } catch (e) {
      // Shape might not have text
    }
  }
  
  // Second pass: if not found in shapes, try page elements
  if (!placeholderShape) {
    try {
      var pageElements = slide.getPageElements();
      Logger.log('Searching ' + pageElements.length + ' page elements for video placeholder');
      for (var i = 0; i < pageElements.length; i++) {
        try {
          var element = pageElements[i];
          if (element.getPageElementType() === 'SHAPE') {
            var shape = element.asShape();
            var text = shape.getText().asString();
            if (text.indexOf(placeholder) !== -1) {
              placeholderElement = element;
              Logger.log('Found video placeholder in page element at index ' + i + ': "' + text + '"');
              break;
            }
          }
        } catch (e) {
          // Continue searching
        }
      }
    } catch (e) {
      Logger.log('Error searching page elements: ' + e);
    }
  }
  
  // Get position and dimensions BEFORE removing the placeholder or doing anything else
  var left, top, width, height;
  
  if (placeholderElement) {
    // Use page element position - get these values NOW
    left = placeholderElement.getLeft();
    top = placeholderElement.getTop();
    width = placeholderElement.getWidth();
    height = placeholderElement.getHeight();
    
    Logger.log('Video placeholder found in page element - Position: (' + left + ',' + top + '), Size: ' + width + 'x' + height);
    
    // Extract video ID
    var videoId = extractVideoId(videoUrl);
    if (!videoId) {
      Logger.log('Could not extract video ID from: ' + videoUrl);
      placeholderElement.remove();
      return;
    }
    
    Logger.log('Extracted video ID: ' + videoId);
    
    // Insert video BEFORE removing placeholder
    try {
      insertVideoAtPosition(slide, videoUrl, videoId, left, top, width, height);
      Logger.log('Video inserted successfully');
    } catch (e) {
      Logger.log('Error inserting video: ' + e);
    }
    
    // Remove the placeholder element AFTER video is inserted and positioned
    placeholderElement.remove();
    return;
  }
  
  if (!placeholderShape) {
    Logger.log('WARNING: Video placeholder not found: ' + placeholder);
    Logger.log('Available shapes: ' + shapes.length);
    // Try to list what text is in shapes for debugging
    for (var i = 0; i < shapes.length; i++) {
      try {
        var text = shapes[i].getText().asString();
        if (text.length > 0) {
          Logger.log('Shape ' + i + ' text: "' + text.substring(0, 100) + '"');
        }
      } catch (e) {
        // Skip
      }
    }
    return;
  }
  
  // Get position from shape - do this BEFORE anything else
  left = placeholderShape.getLeft();
  top = placeholderShape.getTop();
  width = placeholderShape.getWidth();
  height = placeholderShape.getHeight();
  
  Logger.log('Video placeholder found in shape - Position: (' + left + ',' + top + '), Size: ' + width + 'x' + height);
  Logger.log('Placeholder shape details: left=' + left + ', top=' + top + ', width=' + width + ', height=' + height);
  
  // Extract video ID from URL
  var videoId = extractVideoId(videoUrl);
  
  if (!videoId) {
    Logger.log('Could not extract video ID from: ' + videoUrl);
    placeholderShape.remove();
    return;
  }
  
  Logger.log('Extracted video ID: ' + videoId);
  
  // Insert video within placeholder bounds BEFORE removing placeholder
  try {
    insertVideoAtPosition(slide, videoUrl, videoId, left, top, width, height);
    Logger.log('Video inserted successfully');
  } catch (e) {
    Logger.log('Error inserting video: ' + e);
  }
  
  // Remove placeholder shape AFTER video is inserted and positioned
  placeholderShape.remove();
}

// Helper function to insert video at specific position with proper sizing
function insertVideoAtPosition(slide, videoUrl, videoId, left, top, width, height) {
  try {
    Logger.log('Inserting video at position: (' + left + ',' + top + ') size: ' + width + 'x' + height);
    
    var videoElement;
    
    if (videoUrl.indexOf('youtube.com') !== -1 || videoUrl.indexOf('youtu.be') !== -1) {
      // YouTube video - insertVideo returns a PageElement
      videoElement = slide.insertVideo('https://www.youtube.com/watch?v=' + videoId);
      Logger.log('YouTube video inserted');
    } else if (videoUrl.indexOf('drive.google.com') !== -1) {
      // Google Drive video - try using original URL first, then fallback to thumbnail
      // Google Slides insertVideo() may not support Drive URLs directly
      try {
        // Try with original URL (preserving query parameters)
        videoElement = slide.insertVideo(videoUrl);
        Logger.log('Google Drive video inserted with original URL: ' + videoUrl);
      } catch (e) {
        Logger.log('insertVideo() failed for Drive URL, trying thumbnail approach: ' + e);
        // Fallback: insert thumbnail image with link (similar to Vimeo)
        addDriveVideoThumbnail(slide, videoUrl, videoId, left, top, width, height);
        return; // Thumbnail function handles positioning
      }
    } else if (videoUrl.indexOf('vimeo.com') !== -1) {
      // For Vimeo, add a linked thumbnail
      addVimeoThumbnail(slide, videoUrl, videoId, left, top, width, height);
      return; // Vimeo function handles positioning
    } else {
      Logger.log('Unsupported video platform: ' + videoUrl);
      return;
    }
    
    // insertVideo returns a PageElement - use it directly to set position
    // Use the getObjectId() method to verify we have the right element
    try {
      Logger.log('Video element type: ' + videoElement.getPageElementType());
      Logger.log('Setting video position to: left=' + left + ', top=' + top);
      Logger.log('Setting video size to: width=' + width + ', height=' + height);
      
      // Use PageElement methods to position the video
      videoElement.setLeft(left);
      videoElement.setTop(top);
      videoElement.setWidth(width);
      videoElement.setHeight(height);
      
      Logger.log('Video successfully positioned');
      
      // Verify the position was set correctly
      var actualLeft = videoElement.getLeft();
      var actualTop = videoElement.getTop();
      var actualWidth = videoElement.getWidth();
      var actualHeight = videoElement.getHeight();
      
      Logger.log('Video actual position: (' + actualLeft + ',' + actualTop + ') size: ' + actualWidth + 'x' + actualHeight);
      
      // Check if position matches (with small tolerance for rounding)
      if (Math.abs(actualLeft - left) > 1 || Math.abs(actualTop - top) > 1) {
        Logger.log('WARNING: Video position does not match expected position!');
        Logger.log('Expected: (' + left + ',' + top + ')');
        Logger.log('Actual: (' + actualLeft + ',' + actualTop + ')');
      }
    } catch (e) {
      Logger.log('Error setting video position/size: ' + e);
      Logger.log('Error details: ' + e.toString());
      throw e;
    }
  } catch (e) {
    Logger.log('Error in insertVideoAtPosition: ' + e);
    throw e;
  }
}

function extractVideoId(url) {
  if (!url) return null;
  
  // YouTube patterns
  var match = url.match(/(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/embed\/)([^&\?\/]+)/);
  if (match) return match[1];
  
  // Google Drive pattern
  match = url.match(/\/file\/d\/([^\/]+)/);
  if (match) return match[1];
  
  // Vimeo pattern
  match = url.match(/vimeo\.com\/(\d+)/);
  if (match) return match[1];
  
  // Already an ID (11 characters for YouTube)
  if (url.length === 11 && /^[a-zA-Z0-9_-]+$/.test(url)) return url;
  
  return null;
}

function addVimeoThumbnail(slide, vimeoUrl, vimeoId, left, top, width, height) {
  try {
    // Get Vimeo thumbnail
    var apiUrl = 'https://vimeo.com/api/v2/video/' + vimeoId + '.json';
    var response = UrlFetchApp.fetch(apiUrl);
    var data = JSON.parse(response.getContentText());
    
    if (data && data[0] && data[0].thumbnail_large) {
      var thumbnailUrl = data[0].thumbnail_large;
      
      // Insert thumbnail image
      var imgResponse = UrlFetchApp.fetch(thumbnailUrl);
      var blob = imgResponse.getBlob();
      var image = slide.insertImage(blob, left, top, width, height);
      
      // Add link to the video
      image.setLinkUrl(vimeoUrl);
    }
  } catch (e) {
    Logger.log('Error adding Vimeo thumbnail: ' + e);
  }
}

function addDriveVideoThumbnail(slide, driveUrl, driveFileId, left, top, width, height) {
  try {
    // Use Google Drive thumbnail API
    // Try larger size for better quality
    var thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + driveFileId + '&sz=w1920';
    
    Logger.log('Fetching Drive thumbnail from: ' + thumbnailUrl);
    
    // Insert thumbnail image
    var imgResponse = UrlFetchApp.fetch(thumbnailUrl);
    var blob = imgResponse.getBlob();
    var image = slide.insertImage(blob, left, top, width, height);
    
    // Add link to the video (use preview URL for better playback)
    var previewUrl = 'https://drive.google.com/file/d/' + driveFileId + '/preview';
    image.setLinkUrl(previewUrl);
    
    Logger.log('Drive video thumbnail inserted with link: ' + previewUrl);
  } catch (e) {
    Logger.log('Error adding Drive video thumbnail: ' + e);
    Logger.log('Error details: ' + e.toString());
  }
}

// ============================================
// TEST FUNCTION FOR TEMPLATE ACCESS
// Run this function to test if template is accessible
// ============================================

function testTemplateAccess() {
  try {
    var templateId = getConfig('SLIDES_TEMPLATE_ID', '16Rwh50bq1lWbvtMCSYaE03rQBu6gigLHX7u5w01jLH0');
    Logger.log('Template ID: ' + templateId);
    
    if (!templateId || templateId === 'YOUR_TEMPLATE_PRESENTATION_ID_HERE') {
      Logger.log('ERROR: Template ID not configured in CONFIG sheet');
      return {
        success: false,
        error: 'Template ID not configured. Please set SLIDES_TEMPLATE_ID in CONFIG sheet.'
      };
    }
    
    templateId = String(templateId).trim();
    Logger.log('Template ID length: ' + templateId.length);
    
    try {
      var file = DriveApp.getFileById(templateId);
      Logger.log('✓ File found via DriveApp');
      Logger.log('File name: ' + file.getName());
      Logger.log('MIME type: ' + file.getMimeType());
      
      if (file.getMimeType() !== 'application/vnd.google-apps.presentation') {
        Logger.log('WARNING: File is not a Google Slides presentation');
        return {
          success: false,
          error: 'File is not a Google Slides presentation. MIME type: ' + file.getMimeType()
        };
      }
      
      // Try to open with SlidesApp
      try {
        var testPresentation = SlidesApp.openById(templateId);
        var slides = testPresentation.getSlides();
        Logger.log('✓ Template accessible via SlidesApp');
        Logger.log('Number of slides: ' + slides.length);
        
        return {
          success: true,
          fileName: file.getName(),
          mimeType: file.getMimeType(),
          slideCount: slides.length,
          message: 'Template is accessible and ready to use!'
        };
      } catch (slidesError) {
        Logger.log('ERROR accessing via SlidesApp: ' + slidesError.toString());
        return {
          success: false,
          error: 'Cannot access via SlidesApp: ' + slidesError.toString()
        };
      }
    } catch (driveError) {
      Logger.log('ERROR accessing via DriveApp: ' + driveError.toString());
      return {
        success: false,
        error: 'Cannot access via DriveApp: ' + driveError.toString(),
        suggestion: 'Please ensure: 1) Drive permissions are granted, 2) The template file exists, 3) You have access to the file'
      };
    }
  } catch (error) {
    Logger.log('ERROR in testTemplateAccess: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ============================================
// TEST FUNCTION TO FIND PLACEHOLDERS IN TEMPLATE
// Run this to see if placeholders are being detected
// ============================================

function testFindPlaceholders() {
  try {
    var templateId = '16Rwh50bq1lWbvtMCSYaE03rQBu6gigLHX7u5w01jLH0';
    var presentation = SlidesApp.openById(templateId);
    var slides = presentation.getSlides();
    
    Logger.log('=== TEMPLATE DIAGNOSTICS ===');
    Logger.log('Total slides: ' + slides.length);
    Logger.log('');
    
    // Test each slide
    for (var s = 0; s < slides.length; s++) {
      var slide = slides[s];
      Logger.log('--- Slide ' + (s + 1) + ' ---');
      
      // Check shapes
      var shapes = slide.getShapes();
      Logger.log('Shapes found: ' + shapes.length);
      
      for (var i = 0; i < shapes.length; i++) {
        try {
          var shape = shapes[i];
          var text = shape.getText().asString();
          
          // Check for placeholders
          if (text.indexOf('{{') !== -1) {
            Logger.log('  Shape ' + i + ': "' + text.substring(0, 100) + '"');
            
            // List all placeholders found
            var placeholders = text.match(/\{\{[^}]+\}\}/g);
            if (placeholders) {
              Logger.log('    Placeholders: ' + placeholders.join(', '));
            }
          }
        } catch (e) {
          Logger.log('  Shape ' + i + ': (no text or error: ' + e + ')');
        }
      }
      
      // Check page elements (might be different from shapes)
      try {
        var pageElements = slide.getPageElements();
        Logger.log('Page elements found: ' + pageElements.length);
        
        for (var p = 0; p < pageElements.length; p++) {
          try {
            var element = pageElements[p];
            if (element.getPageElementType() === 'SHAPE') {
              var shape = element.asShape();
              var text = shape.getText().asString();
              
              if (text.indexOf('{{') !== -1) {
                Logger.log('  Page Element ' + p + ': "' + text.substring(0, 100) + '"');
                
                var placeholders = text.match(/\{\{[^}]+\}\}/g);
                if (placeholders) {
                  Logger.log('    Placeholders: ' + placeholders.join(', '));
                }
              }
            }
          } catch (e) {
            // Continue
          }
        }
      } catch (e) {
        Logger.log('Error checking page elements: ' + e);
      }
      
      Logger.log('');
    }
    
    Logger.log('=== END DIAGNOSTICS ===');
    return 'Check the logs for placeholder detection results';
    
  } catch (error) {
    Logger.log('ERROR in testFindPlaceholders: ' + error.toString());
    return 'Error: ' + error.toString();
  }
}

// ============================================
// LEADERSHIP TEAM FUNCTIONS
// ============================================

/**
 * Get leadership team members for a specific company
 */
function getLeadershipTeam(companyId) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('LEADERSHIP_TEAM');
    
    if (!sheet) {
      Logger.log('WARNING: LEADERSHIP_TEAM sheet not found');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    
    var headers = data[0];
    var teamMembers = [];
    
    for (var i = 1; i < data.length; i++) {
      var member = {};
      for (var j = 0; j < headers.length; j++) {
        member[headers[j]] = data[i][j] || '';
      }
      
      // Filter by CompanyID if provided
      if (!companyId || member.CompanyID === companyId) {
        teamMembers.push(member);
      }
    }
    
    // Sort by DisplayOrder if available
    teamMembers.sort(function(a, b) {
      var orderA = parseInt(a.DisplayOrder) || 999;
      var orderB = parseInt(b.DisplayOrder) || 999;
      return orderA - orderB;
    });
    
    return teamMembers;
    
  } catch (error) {
    Logger.log('ERROR in getLeadershipTeam: ' + error.toString());
    return [];
  }
}

/**
 * Get all leadership teams grouped by company
 */
function getAllLeadershipTeams() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('LEADERSHIP_TEAM');
    
    if (!sheet) {
      Logger.log('WARNING: LEADERSHIP_TEAM sheet not found');
      return {};
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return {};
    }
    
    var headers = data[0];
    var teamsByCompany = {};
    
    for (var i = 1; i < data.length; i++) {
      var member = {};
      for (var j = 0; j < headers.length; j++) {
        member[headers[j]] = data[i][j] || '';
      }
      
      var companyId = member.CompanyID || '';
      if (!teamsByCompany[companyId]) {
        teamsByCompany[companyId] = [];
      }
      
      teamsByCompany[companyId].push(member);
    }
    
    // Sort each company's team by DisplayOrder
    Object.keys(teamsByCompany).forEach(function(companyId) {
      teamsByCompany[companyId].sort(function(a, b) {
        var orderA = parseInt(a.DisplayOrder) || 999;
        var orderB = parseInt(b.DisplayOrder) || 999;
        return orderA - orderB;
      });
    });
    
    return teamsByCompany;
    
  } catch (error) {
    Logger.log('ERROR in getAllLeadershipTeams: ' + error.toString());
    return {};
  }
}
// ============================================
// TEST FUNCTION FOR EMAIL PERMISSIONS
// Run this function once to authorize email sending
// ============================================

function testEmailPermissions() {
  try {
    Logger.log('Testing email permissions...');
    
    // Get your email from CONFIG or use the active user
    var testEmail = getConfig('UPDATE_REQUEST_NOTIFICATION_EMAILS', Session.getActiveUser().getEmail());
    var emails = testEmail.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e.indexOf('@') !== -1; });
    
    if (emails.length === 0) {
      emails = [Session.getActiveUser().getEmail()];
    }
    
    // This will trigger the permission request if not already granted
    MailApp.sendEmail({
      to: emails[0], // Send to first email in config, or yourself
      subject: 'Test Email - Permission Check',
      body: 'If you received this email, MailApp permissions are working correctly!'
    });
    
    Logger.log('Email sent successfully! Permissions are granted.');
    return 'Email sent successfully! Check your inbox at: ' + emails[0];
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return 'Error: ' + error.toString();
  }
}

// ============================================
// USAGE TRACKING FUNCTIONS
// ============================================

/**
 * Log usage events to the USAGE_LOG sheet
 * Called from client-side via google.script.run
 */
function logUsage(action, page, details) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var logSheet = ss.getSheetByName('USAGE_LOG');
    
    // Create sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet('USAGE_LOG');
      logSheet.appendRow([
        'Timestamp', 
        'UserID', 
        'SessionID', 
        'Action', 
        'Page', 
        'Details',
        'IsDeveloper',
        'Browser',
        'Timezone'
      ]);
      
      // Format header row
      var headerRange = logSheet.getRange(1, 1, 1, 9);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
    }
    
    var timestamp = new Date();
    var parsedDetails = typeof details === 'string' ? JSON.parse(details) : details;
    
    // Extract key fields from details
    var userId = parsedDetails.userId || 'unknown';
    var sessionId = parsedDetails.sessionId || 'unknown';
    var isDeveloper = parsedDetails.isDeveloper || false;
    var browser = '';
    var timezone = '';
    
    if (parsedDetails.fingerprint) {
      browser = parsedDetails.fingerprint.userAgent || '';
      timezone = parsedDetails.fingerprint.timezone || '';
    }
    
    // Append the log entry
    logSheet.appendRow([
      timestamp,
      userId,
      sessionId,
      action,
      page,
      JSON.stringify(parsedDetails),
      isDeveloper,
      browser,
      timezone
    ]);
    
    return { success: true };
    
  } catch(e) {
    Logger.log('ERROR in logUsage: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Get usage statistics
 * Returns summary data about app usage
 */
function getUsageStats(excludeDeveloper) {
  try {
    if (excludeDeveloper === undefined) {
      excludeDeveloper = true;
    }
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var logSheet = ss.getSheetByName('USAGE_LOG');
    
    if (!logSheet) {
      return {
        success: false,
        error: 'USAGE_LOG sheet not found'
      };
    }
    
    var data = logSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        totalUsers: 0,
        totalSessions: 0,
        totalActions: 0,
        topPages: [],
        topActions: [],
        actionsByDay: []
      };
    }
    
    // Process data
    var uniqueUsers = new Set();
    var uniqueSessions = new Set();
    var pageCount = {};
    var actionCount = {};
    var actionsByDay = {};
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var userId = row[1];
      var sessionId = row[2];
      var action = row[3];
      var page = row[4];
      var isDev = row[6];
      var timestamp = row[0];
      
      // Skip developer actions if requested
      if (excludeDeveloper && isDev) {
        continue;
      }
      
      uniqueUsers.add(userId);
      uniqueSessions.add(sessionId);
      
      // Count pages
      pageCount[page] = (pageCount[page] || 0) + 1;
      
      // Count actions
      actionCount[action] = (actionCount[action] || 0) + 1;
      
      // Count by day
      if (timestamp instanceof Date) {
        var dateKey = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        actionsByDay[dateKey] = (actionsByDay[dateKey] || 0) + 1;
      }
    }
    
    // Convert to sorted arrays
    var topPages = Object.keys(pageCount).map(function(key) {
      return { page: key, count: pageCount[key] };
    }).sort(function(a, b) {
      return b.count - a.count;
    });
    
    var topActions = Object.keys(actionCount).map(function(key) {
      return { action: key, count: actionCount[key] };
    }).sort(function(a, b) {
      return b.count - a.count;
    });
    
    var actionsByDayArray = Object.keys(actionsByDay).map(function(key) {
      return { date: key, count: actionsByDay[key] };
    }).sort(function(a, b) {
      return a.date.localeCompare(b.date);
    });
    
    return {
      success: true,
      totalUsers: uniqueUsers.size,
      totalSessions: uniqueSessions.size,
      totalActions: data.length - 1,
      topPages: topPages,
      topActions: topActions,
      actionsByDay: actionsByDayArray
    };
    
  } catch(e) {
    Logger.log('ERROR in getUsageStats: ' + e.toString());
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Get session details
 * Returns detailed breakdown by session
 */
function getSessionStats(excludeDeveloper) {
  try {
    if (excludeDeveloper === undefined) {
      excludeDeveloper = true;
    }
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var logSheet = ss.getSheetByName('USAGE_LOG');
    
    if (!logSheet) {
      return {
        success: false,
        error: 'USAGE_LOG sheet not found'
      };
    }
    
    var data = logSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        sessions: []
      };
    }
    
    var sessions = {};
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var timestamp = row[0];
      var userId = row[1];
      var sessionId = row[2];
      var action = row[3];
      var page = row[4];
      var isDev = row[6];
      
      // Skip developer sessions if requested
      if (excludeDeveloper && isDev) {
        continue;
      }
      
      if (!sessions[sessionId]) {
        sessions[sessionId] = {
          sessionId: sessionId,
          userId: userId,
          actions: 0,
          pages: [],
          isDeveloper: isDev,
          firstSeen: timestamp,
          lastSeen: timestamp
        };
      }
      
      sessions[sessionId].actions++;
      
      if (sessions[sessionId].pages.indexOf(page) === -1) {
        sessions[sessionId].pages.push(page);
      }
      
      if (timestamp < sessions[sessionId].firstSeen) {
        sessions[sessionId].firstSeen = timestamp;
      }
      
      if (timestamp > sessions[sessionId].lastSeen) {
        sessions[sessionId].lastSeen = timestamp;
      }
    }
    
    // Convert to array and calculate duration
    var sessionArray = Object.keys(sessions).map(function(key) {
      var session = sessions[key];
      var duration = 0;
      
      if (session.firstSeen instanceof Date && session.lastSeen instanceof Date) {
        duration = Math.round((session.lastSeen - session.firstSeen) / 1000 / 60); // minutes
      }
      
      return {
        sessionId: session.sessionId,
        userId: session.userId,
        actions: session.actions,
        uniquePages: session.pages.length,
        pages: session.pages.join(', '),
        isDeveloper: session.isDeveloper,
        firstSeen: session.firstSeen,
        lastSeen: session.lastSeen,
        durationMinutes: duration
      };
    }).sort(function(a, b) {
      return b.firstSeen - a.firstSeen; // Most recent first
    });
    
    return {
      success: true,
      sessions: sessionArray
    };
    
  } catch(e) {
    Logger.log('ERROR in getSessionStats: ' + e.toString());
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * Clear usage logs older than specified days
 * Use with caution!
 */
function clearOldUsageLogs(daysToKeep) {
  try {
    if (!daysToKeep || daysToKeep < 7) {
      throw new Error('Must keep at least 7 days of logs');
    }
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var logSheet = ss.getSheetByName('USAGE_LOG');
    
    if (!logSheet) {
      return { success: false, error: 'USAGE_LOG sheet not found' };
    }
    
    var data = logSheet.getDataRange().getValues();
    var cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    
    var rowsToDelete = [];
    
    for (var i = data.length - 1; i >= 1; i--) {
      var timestamp = data[i][0];
      
      if (timestamp instanceof Date && timestamp < cutoffDate) {
        rowsToDelete.push(i + 1); // +1 for 1-indexed
      }
    }
    
    // Delete rows in reverse order to maintain indices
    rowsToDelete.forEach(function(rowNum) {
      logSheet.deleteRow(rowNum);
    });
    
    return {
      success: true,
      rowsDeleted: rowsToDelete.length,
      message: 'Deleted ' + rowsToDelete.length + ' rows older than ' + daysToKeep + ' days'
    };
    
  } catch(e) {
    Logger.log('ERROR in clearOldUsageLogs: ' + e.toString());
    return {
      success: false,
      error: e.toString()
    };
  }
}

// ============================================
// HUMAN-READABLE ANALYTICS DASHBOARD
// ============================================

/**
 * Generate human-readable analytics from usage logs
 * Run this manually or set up a time-based trigger to run hourly/daily
 */
function generateAnalyticsDashboard() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var logSheet = ss.getSheetByName('USAGE_LOG');
    var dashSheet = ss.getSheetByName('Analytics Dashboard');
    
    if (!logSheet) {
      Logger.log('No USAGE_LOG sheet found');
      return { success: false, error: 'USAGE_LOG sheet not found. Start using the app to generate logs first.' };
    }
    
    // Create dashboard sheet if it doesn't exist
    if (!dashSheet) {
      dashSheet = ss.insertSheet('Analytics Dashboard');
    }
    
    // Clear existing content
    dashSheet.clear();
    
    // Get all log data (excluding header)
    var data = logSheet.getDataRange().getValues();
    if (data.length <= 1) {
      dashSheet.getRange(1, 1).setValue('No usage data yet. Start using the app to generate analytics!');
      return { success: true, message: 'No data to analyze yet' };
    }
    
    var headers = data.shift(); // Remove header row
    
    // Filter out developer activity
    var userActivity = data.filter(function(row) {
      var isDev = row[6]; // IsDeveloper column (index 6)
      return isDev === false || isDev === 'FALSE' || isDev === '' || isDev === null;
    });
    
    if (userActivity.length === 0) {
      dashSheet.getRange(1, 1).setValue('No user activity found (all activity is from developers).');
      return { success: true, message: 'No user activity to analyze' };
    }
    
    // Group by session
    var sessions = {};
    userActivity.forEach(function(row) {
      var sessionId = row[2]; // SessionID column
      if (!sessions[sessionId]) {
        sessions[sessionId] = [];
      }
      sessions[sessionId].push({
        timestamp: row[0],
        userId: row[1],
        action: row[3],
        page: row[4],
        details: row[5]
      });
    });
    
    // Build dashboard output
    var output = [];
    
    // Title
    output.push(['USER ACTIVITY SUMMARY']);
    output.push(['Generated: ' + new Date().toLocaleString()]);
    output.push(['']);
    
    // Summary stats
    var totalSessions = Object.keys(sessions).length;
    var uniqueUsers = [...new Set(userActivity.map(function(row) { return row[1]; }))].length;
    var totalActions = userActivity.length;
    
    output.push(['OVERVIEW']);
    output.push(['Total Users', uniqueUsers]);
    output.push(['Total Sessions', totalSessions]);
    output.push(['Total Actions', totalActions]);
    output.push(['Avg Actions per Session', Math.round((totalActions / totalSessions) * 10) / 10 || 0]);
    output.push(['']);
    
    // Page popularity
    var pageCounts = {};
    userActivity.forEach(function(row) {
      var page = row[4] || 'unknown';
      pageCounts[page] = (pageCounts[page] || 0) + 1;
    });
    
    var topPages = Object.keys(pageCounts).map(function(key) {
      return { page: key, count: pageCounts[key] };
    }).sort(function(a, b) {
      return b.count - a.count;
    });
    
    if (topPages.length > 0) {
      output.push(['MOST VISITED PAGES']);
      topPages.slice(0, 5).forEach(function(item) {
        output.push([translatePageName(item.page), item.count + ' visits']);
      });
      output.push(['']);
    }
    
    // Action popularity
    var actionCounts = {};
    userActivity.forEach(function(row) {
      var action = row[3] || 'unknown';
      actionCounts[action] = (actionCounts[action] || 0) + 1;
    });
    
    var topActions = Object.keys(actionCounts).map(function(key) {
      return { action: key, count: actionCounts[key] };
    }).sort(function(a, b) {
      return b.count - a.count;
    });
    
    if (topActions.length > 0) {
      output.push(['MOST COMMON ACTIONS']);
      topActions.slice(0, 5).forEach(function(item) {
        output.push([translateActionName(item.action), item.count + ' times']);
      });
      output.push(['']);
    }
    
    // Session narratives
    output.push(['USER JOURNEY SESSIONS']);
    output.push(['']);
    
    var sessionNum = 0;
    var sortedSessionIds = Object.keys(sessions).sort(function(a, b) {
      var timeA = sessions[a][0].timestamp;
      var timeB = sessions[b][0].timestamp;
      return new Date(timeA) - new Date(timeB); // Oldest first
    });
    
    for (var i = 0; i < sortedSessionIds.length; i++) {
      var sessionId = sortedSessionIds[i];
      sessionNum++;
      var events = sessions[sessionId].sort(function(a, b) {
        return new Date(a.timestamp) - new Date(b.timestamp);
      });
      
      var startTime = new Date(events[0].timestamp);
      var endTime = new Date(events[events.length - 1].timestamp);
      var duration = Math.round((endTime - startTime) / 1000 / 60); // minutes
      
      output.push(['Session #' + sessionNum]);
      output.push(['When', startTime.toLocaleString()]);
      output.push(['Duration', duration + ' minute' + (duration !== 1 ? 's' : '') + ' (' + events.length + ' action' + (events.length !== 1 ? 's' : '') + ')']);
      output.push(['']);
      output.push(['What They Did:']);
      
      // Convert events to narrative
      events.forEach(function(event, idx) {
        var narrative = translateEventToNarrative(event, idx === 0);
        if (narrative) {
          output.push(['  ' + (idx + 1) + '. ' + narrative]);
        }
      });
      
      output.push(['']);
      output.push(['---']);
      output.push(['']);
    }
    
    // Write to sheet
    if (output.length > 0) {
      var numRows = output.length;
      var numCols = 2;
      
      // Ensure we have enough rows/columns
      var maxRows = dashSheet.getMaxRows();
      if (numRows > maxRows) {
        dashSheet.insertRowsAfter(maxRows, numRows - maxRows);
      }
      var maxCols = dashSheet.getMaxColumns();
      if (numCols > maxCols) {
        dashSheet.insertColumnsAfter(maxCols, numCols - maxCols);
      }
      
      // Normalize output to always have 2 columns
      var normalized = output.map(function(row) {
        if (!Array.isArray(row)) return [row, ''];
        if (row.length === 1) return [row[0], ''];
        return [row[0], row[1]];
      });
      
      dashSheet.getRange(1, 1, numRows, numCols).setValues(normalized);
      
      // Format the sheet
      dashSheet.setColumnWidth(1, 250);
      dashSheet.setColumnWidth(2, 400);
      
      // Bold section headers and format
      for (var i = 1; i <= numRows; i++) {
        var cell = dashSheet.getRange(i, 1);
        var value = cell.getValue();
        if (value && (
          value === 'OVERVIEW' || 
          value === 'MOST VISITED PAGES' ||
          value === 'MOST COMMON ACTIONS' ||
          value === 'USER JOURNEY SESSIONS' ||
          value === 'USER ACTIVITY SUMMARY' ||
          value.toString().startsWith('Session #')
        )) {
          cell.setFontWeight('bold');
          cell.setFontSize(12);
          cell.setBackground('#f3f4f6');
        }
      }
      
      // Format "Generated" row
      var generatedRow = dashSheet.getRange(2, 1);
      generatedRow.setFontStyle('italic');
      generatedRow.setFontSize(10);
    }
    
    Logger.log('Analytics dashboard generated successfully');
    return { success: true, sessions: totalSessions };
    
  } catch(e) {
    Logger.log('ERROR in generateAnalyticsDashboard: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Translate raw events into human-readable narratives
 */
function translateEventToNarrative(event, isFirst) {
  var action = event.action;
  var page = event.page;
  var details = {};
  
  try {
    if (typeof event.details === 'string') {
      details = JSON.parse(event.details);
    } else {
      details = event.details || {};
    }
  } catch(e) {
    details = {};
  }
  
  // Map actions to narratives
  switch(action) {
    case 'app_loaded':
      return isFirst ? 'Opened the app' : 'Refreshed the page';
      
    case 'navigate':
      var toPage = details.to || page;
      var fromPage = details.from || '';
      if (fromPage && toPage && fromPage === toPage) {
        return null;
      }
      if (fromPage && fromPage !== 'none' && fromPage !== 'unknown') {
        return 'Navigated from ' + translatePageName(fromPage) + ' to ' + translatePageName(toPage);
      }
      return 'Visited the ' + translatePageName(toPage) + ' page';
      
    case 'search':
      var query = details.query || details.searchQuery || 'products';
      return 'Searched for "' + query + '"';
      
    case 'product_view':
      var productName = details.productName || details.product || 'a product';
      return 'Viewed product: ' + productName;
      
    case 'company_view':
      var companyName = details.companyName || details.company || 'a company';
      return 'Viewed company profile: ' + companyName;

    case 'company_products_view':
      return 'Opened products for company: ' + (details.companyName || 'Unknown');
      
    case 'glossary_tooltip_view':
      var abbreviation = details.abbreviation || 'term';
      var fullName = details.fullName || '';
      if (fullName) {
        return 'Looked up glossary term: ' + abbreviation + ' (' + fullName + ')';
      }
      return 'Looked up glossary term: ' + abbreviation;
      
    case 'crosssell_generate':
      if (details.products) {
        return 'Selected products for cross-sell: ' + details.products;
      }
      var count = details.selectedProductsCount || details.count || details.opportunitiesCount || '?';
      return 'Generated cross-sell opportunities (selected ' + count + ' product' + (count !== 1 ? 's' : '') + ')';

    case 'crosssell_results':
      if (details.recommendations) {
        return 'Cross-sell recommendations shown: ' + details.recommendations;
      }
      return 'Cross-sell recommendations shown';
      
    case 'export_slides':
      var productCount = details.productCount || details.opportunitiesCount || 0;
      var hasOverview = details.includeOverview ? 'company overviews, ' : '';
      var hasProducts = details.includeProducts ? 'product catalogue' : '';
      var hasOpps = details.includeOpportunities ? 'opportunities' : '';
      var parts = [hasOverview, hasProducts, hasOpps].filter(function(p) { return p; });
      return 'Exported presentation with ' + parts.join(' and ');
      
    case 'filters_applied':
      var brand = details.brand && details.brand !== 'all' ? details.brand : '';
      var industry = details.industry && details.industry !== 'all' ? details.industry : '';
      var size = details.size && details.size !== 'all' ? details.size : '';
      var filters = [brand, industry, size].filter(function(f) { return f; });
      if (filters.length > 0) {
        return 'Applied filters: ' + filters.join(', ');
      }
      return 'Applied filters to narrow results';
      
    case 'form_submit':
      var formType = details.formType || 'a form';
      var updateType = details.updateType ? ' (' + details.updateType + ')' : '';
      return 'Submitted ' + formType + updateType;
      
    case 'app_unload':
      return 'Left the app';
      
    default:
      // Generic fallback
      if (page && page !== 'unknown') {
        return 'Interacted with ' + translatePageName(page) + ' (' + action.replace(/_/g, ' ') + ')';
      }
      return action.replace(/_/g, ' ').replace(/\b\w/g, function(l) { return l.toUpperCase(); });
  }
}

/**
 * Translate technical page names to friendly names
 */
function translatePageName(pageName) {
  if (!pageName || pageName === 'unknown') return 'Unknown Page';
  
  var pageMap = {
    'home': 'Home',
    'about': 'About Us',
    'knowledge': 'Knowledge Base',
    'crosssell': 'Cross-Sell Finder',
    'resources': 'Resources',
    'submit': 'Submit Update Request'
  };
  
  return pageMap[pageName] || pageName.replace(/-/g, ' ').replace(/\b\w/g, function(l) { return l.toUpperCase(); });
}

/**
 * Translate action names to friendly names
 */
function translateActionName(actionName) {
  var actionMap = {
    'app_loaded': 'App Opened',
    'navigate': 'Page Navigation',
    'search': 'Search',
    'product_view': 'Product View',
    'company_view': 'Company View',
    'company_products_view': 'Viewed Company Products',
    'glossary_tooltip_view': 'Glossary Lookup',
    'crosssell_generate': 'Cross-Sell Generated',
    'crosssell_results': 'Cross-Sell Recommendations',
    'export_slides': 'Export to Slides',
    'filters_applied': 'Filters Applied',
    'form_submit': 'Form Submitted',
    'app_unload': 'App Closed'
  };
  
  return actionMap[actionName] || actionName.replace(/_/g, ' ').replace(/\b\w/g, function(l) { return l.toUpperCase(); });
}
