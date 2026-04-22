// Vercel Speed Insights integration
// This module imports and initializes Speed Insights for the static site

import { injectSpeedInsights } from '@vercel/speed-insights';

// Initialize Speed Insights
// This will automatically track web vitals and performance metrics
injectSpeedInsights({
  // Framework identifier for analytics (optional)
  framework: 'vanilla',
  // Debug mode in development (optional)
  debug: false
});
