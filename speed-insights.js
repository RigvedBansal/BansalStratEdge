// Vercel Speed Insights integration
// This module imports and initializes Speed Insights for the static site

import { injectSpeedInsights } from './dist/speed-insights.mjs';

// Initialize Speed Insights
injectSpeedInsights({
  // Framework identifier for analytics
  framework: 'vanilla'
});
