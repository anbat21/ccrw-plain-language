# Azure Application Insights Telemetry Setup Guide

## Overview

This guide explains how to set up and use Azure Application Insights with OpenTelemetry for the CCRW Plain Language Annotator add-in. The telemetry system captures usage signals and performance metrics without collecting personally identifiable information (PII) or business data.

## Prerequisites

- Azure subscription with access to Application Insights
- Node.js and npm installed locally for development
- Git repository initialized

## Setup Steps

### 1. Create Azure Application Insights Resource

1. Go to [Azure Portal](https://portal.azure.com)
2. Click **Create a resource** and search for **Application Insights**
3. Create a new Application Insights resource:
   - **Name**: `ccrw-plain-language-insights` (or similar)
   - **Resource Group**: Create new or select existing
   - **Region**: Select your preferred region
   - **Resource Mode**: Workspace-based
4. Once created, go to the resource and note the **Instrumentation Key**
   - Located under: **Overview > Connection Strings**
   - Format: `00000000-0000-0000-0000-000000000000`

### 2. Configure Environment Variables

#### Local Development

1. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```

2. Update `.env` with your Instrumentation Key:
   ```
   REACT_APP_INSTRUMENTATION_KEY=your-actual-instrumentation-key-here
   ```

#### Azure Deployment

When deploying to Azure (App Service/Container Instances):
1. Go to your app's configuration settings
2. Add a new application setting:
   - **Name**: `REACT_APP_INSTRUMENTATION_KEY`
   - **Value**: Your Instrumentation Key

### 3. Install Dependencies

```bash
cd CCRW-Annotator
npm install
```

This installs:
- `@microsoft/applicationinsights-web` - Client-side SDK
- `@azure/monitor-opentelemetry` - Server-side OpenTelemetry
- `@opentelemetry/api` - OpenTelemetry core API

### 4. Verify Installation

#### Start Development Server

```bash
npm start
```

#### Check Console Output

Look for these messages in the browser console:
```
[Telemetry] Azure Application Insights initialized successfully
```

And in the server logs:
```
[Telemetry] Azure Monitor OpenTelemetry initialized
```

## Tracked Events

### Client-Side Events (App.tsx)

The add-in tracks the following user events:

| Event | Properties | Trigger |
|-------|-----------|---------|
| `AddInInitialized` | `sessionId`, `timestamp` | Add-in loads |
| `AudienceSetupCompleted` | `primaryAudience`, `audienceType`, `communicationType` | User completes audience setup and clicks "Start Analysis" |
| `AnalysisStarted` | `audience`, `audienceType` | User clicks "Analyze Selected Text" |
| `AnalysisCompleted` | `findingsCount`, `plainessScore` | Analysis finishes successfully |
| `AnalysisError` | `reason`, `detail` | Analysis fails (no text, invalid plan, etc.) |
| `TipsApplyStarted` | `selectedCount`, `totalCount` | User clicks "Apply Plan" |
| `TipsApplyCompleted` | `appliedCount`, `skippedCount`, `remainingCount` | Tips are applied successfully |
| `TipsApplyError` | `context` (from exception) | Error while applying tips |

### Server-Side Events (server.js)

The Express server automatically tracks:
- HTTP request/response metrics
- Duration and performance timing
- Errors and exceptions
- Bot token generation requests

## Viewing Telemetry Data

### In Azure Portal

1. **Live Metrics Stream**:
   - Azure Portal > Application Insights > Live Metrics
   - See real-time user activity and performance

2. **Custom Events**:
   - Azure Portal > Application Insights > Events
   - View all custom tracked events with properties

3. **Analytics Queries**:
   - Azure Portal > Application Insights > Logs
   - Run KQL (Kusto Query Language) queries

#### Example Queries

**Recent Analysis Events**:
```kusto
customEvents
| where name == "AnalysisCompleted"
| order by timestamp desc
| take 50
```

**Tips Application Summary**:
```kusto
customEvents
| where name startswith "TipsApply"
| summarize count() by name
```

**Error Tracking**:
```kusto
exceptions
| where timestamp > ago(24h)
| summarize count() by outerType
```

**Adoption Metrics**:
```kusto
customEvents
| where name == "AddInInitialized"
| where timestamp > ago(7d)
| summarize UniqueUsers = dcount(tostring(customDimensions["sessionId"]))
```

## Privacy and Data Handling

### What We Track
- Feature usage (analysis, tips applied)
- User demographics (audience type, communication type)
- Performance metrics (analysis count, findings count)
- Error types and exceptions
- Session IDs for activity correlation

### What We Don't Track
- Document content or user text
- User names or email addresses
- IP addresses (can be disabled in settings)
- Query parameters or sensitive URLs
- Authentication tokens

### Data Retention

- Application Insights default retention: 90 days
- Can be extended up to 730 days in settings
- Data can be exported for longer-term analysis

### GDPR Compliance

Data is stored in:
- **US**: East US, South Central US
- **Europe**: North Europe, West Europe

Select the appropriate region when creating Application Insights to comply with local data residency requirements.

## Troubleshooting

### Telemetry Not Appearing

1. **Check Environment Variable**:
   ```bash
   echo $REACT_APP_INSTRUMENTATION_KEY
   ```
   Should return your key (not empty)

2. **Check Console for Initialization**:
   - Open browser DevTools > Console
   - Look for telemetry initialization messages

3. **Verify Network Requests**:
   - DevTools > Network
   - Look for requests to `dc.applicationinsights.azure.com`

### Events Appearing in Portal but Not Completing

- May be a sampling issue (set sampling to 100% for development)
- Data can take 5-10 minutes to appear in portal

### Permission Errors

- Verify Application Insights resource exists
- Confirm Instrumentation Key is correct
- Check Azure Portal > Application Insights > Access Control
- Ensure your account has "Contributor" or "Application Insights API Read" role

## Advanced Configuration

### Sampling

To reduce costs, you can enable sampling:

In [telemetry.ts](../src/services/telemetry.ts):
```typescript
samplingPercentage: 50, // Sample 50% of requests
```

### Custom Dimensions

Add context to all events by modifying the telemetry service:

```typescript
// In App.tsx or any component
telemetry.trackEvent('CustomEvent', {
  environment: process.env.NODE_ENV,
  version: '1.0.0',
  feature: 'myFeature',
});
```

### Performance Monitoring

Application Insights automatically captures:
- Page load times
- AJAX call durations
- Custom timing measurements

## Next Steps

1. **Deploy to Azure**:
   - Add `REACT_APP_INSTRUMENTATION_KEY` to Azure App Service config
   - Deploy and monitor real-world usage

2. **Set Up Alerts**:
   - Azure Portal > Application Insights > Alerts
   - Create alerts for high error rates or performance degradation

3. **Create Dashboards**:
   - Azure Portal > Application Insights > Pinned Metrics
   - Custom dashboards for KPI tracking

4. **Analyze Trends**:
   - Use Analytics to identify adoption patterns
   - Track feature adoption over time
   - Monitor error patterns and improvements

## Support

For issues or questions:
- Check [Application Insights Documentation](https://docs.microsoft.com/en-us/azure/azure-monitor/app/app-insights-overview)
- Review [OpenTelemetry Documentation](https://opentelemetry.io/docs/)
- Contact your Azure support team
