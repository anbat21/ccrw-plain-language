/**
 * Telemetry Service for Azure Application Insights
 * Manages Application Insights initialization and custom event logging
 */

import { ApplicationInsights } from '@microsoft/applicationinsights-web';

interface TelemetryConfig {
  instrumentationKey: string;
  enableAutoRouteTracking?: boolean;
  enableUnhandledPromiseRejectionTracking?: boolean;
  autoTrackPageVisitTime?: boolean;
}

class TelemetryService {
  private appInsights: ApplicationInsights | null = null;
  private isInitialized = false;
  private sessionId = this.generateSessionId();

  /**
   * Initialize Application Insights
   */
  public initialize(config: TelemetryConfig): void {
    if (this.isInitialized) {
      console.warn('Telemetry service already initialized');
      return;
    }

    if (!config.instrumentationKey) {
      console.warn('Application Insights Instrumentation Key not provided. Telemetry disabled.');
      return;
    }

    try {
      const appInsights = new ApplicationInsights({
        config: {
          instrumentationKey: config.instrumentationKey,
          enableAutoRouteTracking: config.enableAutoRouteTracking ?? true,
          enableUnhandledPromiseRejectionTracking: config.enableUnhandledPromiseRejectionTracking ?? true,
          autoTrackPageVisitTime: config.autoTrackPageVisitTime ?? true,
          enableDebug: false,
          disableExceptionTracking: false,
          disableAjaxTracking: false,
          disableFetchTracking: false,
          // Privacy: Don't track URLs, user agent details
          disableCorrelationHeaders: false,
          samplingPercentage: 100,
        },
      });

      appInsights.loadAppInsights();
      this.appInsights = appInsights;
      this.isInitialized = true;

      // Log initialization event
      this.trackEvent('AddInInitialized', {
        sessionId: this.sessionId,
        timestamp: new Date().toISOString(),
      });

      console.log('[Telemetry] Azure Application Insights initialized successfully');
    } catch (error) {
      console.error('[Telemetry] Failed to initialize Application Insights', error);
    }
  }

  /**
   * Track custom event
   */
  public trackEvent(
    eventName: string,
    properties?: { [key: string]: string | number | boolean | undefined },
    measurements?: { [key: string]: number }
  ): void {
    if (!this.isInitialized || !this.appInsights) {
      console.debug(`[Telemetry] Event not tracked (not initialized): ${eventName}`);
      return;
    }

    try {
      const enrichedProperties: { [key: string]: string } = {
        sessionId: this.sessionId,
      };

      // Convert all properties to strings (Application Insights requirement)
      if (properties) {
        Object.entries(properties).forEach(([key, value]) => {
          if (value !== undefined && value !== null) {
            enrichedProperties[key] = String(value);
          }
        });
      }

      this.appInsights.trackEvent({
        name: eventName,
        properties: enrichedProperties,
        measurements,
      });
    } catch (error) {
      console.error(`[Telemetry] Failed to track event: ${eventName}`, error);
    }
  }

  /**
   * Track exception/error
   */
  public trackException(error: Error, context?: string): void {
    if (!this.isInitialized || !this.appInsights) {
      console.debug(`[Telemetry] Exception not tracked (not initialized): ${error.message}`);
      return;
    }

    try {
      this.appInsights.trackException({
        exception: error,
        severityLevel: 2, // Warning level for non-critical errors
        properties: {
          context: context || 'Unknown',
          sessionId: this.sessionId,
        },
      });
    } catch (err) {
      console.error('[Telemetry] Failed to track exception', err);
    }
  }

  /**
   * Track page view (for add-in state changes)
   */
  public trackPageView(pageName: string, properties?: { [key: string]: string }): void {
    if (!this.isInitialized || !this.appInsights) {
      console.debug(`[Telemetry] Page view not tracked (not initialized): ${pageName}`);
      return;
    }

    try {
      this.appInsights.trackPageView({
        name: pageName,
        properties: {
          ...properties,
          sessionId: this.sessionId,
        },
      });
    } catch (error) {
      console.error(`[Telemetry] Failed to track page view: ${pageName}`, error);
    }
  }

  /**
   * Track metric/measurement (as a custom event with measurements)
   */
  public trackMetric(
    name: string,
    value: number,
    properties?: { [key: string]: string }
  ): void {
    if (!this.isInitialized || !this.appInsights) {
      console.debug(`[Telemetry] Metric not tracked (not initialized): ${name}`);
      return;
    }

    try {
      this.appInsights.trackEvent({
        name: `Metric_${name}`,
        measurements: { value },
        properties: {
          ...properties,
          sessionId: this.sessionId,
        },
      });
    } catch (error) {
      console.error(`[Telemetry] Failed to track metric: ${name}`, error);
    }
  }

  /**
   * Flush telemetry data
   */
  public flush(): void {
    if (this.appInsights) {
      this.appInsights.flush();
    }
  }

  /**
   * Generate unique session ID
   */
  private generateSessionId(): string {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Get current session ID
   */
  public getSessionId(): string {
    return this.sessionId;
  }
}

// Export singleton instance
export const telemetry = new TelemetryService();

// Export for easier access
export default telemetry;
