# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script project that creates a Looker Studio connector for Meta (Facebook) Ads data. The connector uses OAuth2 authentication to fetch performance metrics from the Meta Marketing API and presents them in Looker Studio dashboards.

## Architecture

- **Single File Structure**: The entire connector logic is contained in `Code.gs` - a Google Apps Script file
- **OAuth2 Authentication**: Uses the OAuth2 library (ID: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`) for Meta API authentication
- **Fixed Schema**: Provides a simplified, fixed set of metrics rather than dynamic configuration:
  - Dimensions: `Date`, `Campaign Name`
  - Metrics: `Spend (Cost)`, `Total Conversion Value`

## Key Components

### Core Functions
- `getConfig()`: Defines user configuration (Ad Account ID input)
- `getSchema()`: Returns fixed schema of available fields
- `getData()`: Main data fetching function with pagination support
- `fetchInsights()`: Handles Meta Marketing API requests
- `processResponse()`: Transforms API data for Looker Studio

### Authentication Flow
- `getAuthType()`: Specifies OAuth2 authentication
- `getOAuthService()`: Configures OAuth2 service with Meta endpoints
- `authCallback()`: Handles OAuth callback
- `get3PAuthorizationUrls()`: Generates authorization URLs

### Data Processing
- `extractPurchaseConversionValue()`: Extracts purchase conversion values from action_values
- `extractPurchaseActionCount()`: Extracts purchase action counts
- Date range handling with Looker Studio integration support

## Configuration Requirements

### Script Properties (Required)
Set these in Google Apps Script Project Settings → Script Properties:
- `META_CLIENT_ID`: Meta App ID
- `META_CLIENT_SECRET`: Meta App Secret

### Meta API Configuration
- API Version: v18.0 (defined in CONFIG.API_VERSION)
- Required Scopes: `ads_read`, `read_insights`
- Uses campaign-level insights with daily time increment

## Development Notes

### No Build/Test Commands
This is a Google Apps Script project - there are no npm scripts or traditional build commands. Development and testing happens directly in the Google Apps Script editor.

### Deployment Process
1. Configure Script Properties with Meta credentials
2. Deploy as web app through Google Apps Script interface
3. Use deployment ID (not web app URL) in Looker Studio

### Data Limitations
- Maximum 100 pages of API results to prevent timeouts
- Purchase-specific conversion tracking (action_type = 'purchase')
- Fixed schema prevents dynamic field selection but ensures reliability

### Error Handling
- Comprehensive API error checking with user-friendly messages
- OAuth token validation and refresh handling
- Configuration validation for account ID format (numbers only)

## API Integration Details

- **Endpoint**: `https://graph.facebook.com/v18.0/{account_id}/insights`
- **Level**: Campaign-level data
- **Time Increment**: Daily (time_increment = 1)
- **Pagination**: Automatic handling up to 100 pages
- **Date Ranges**: Supports both Looker Studio date ranges and fallback configs