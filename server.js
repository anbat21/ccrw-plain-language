const path = require('path');

// Load .env only in local development (Azure uses portal environment variables)
try {
    require('dotenv').config({ path: path.join(__dirname, '.env') });
} catch (e) {
    console.log('dotenv not available, using system environment variables');
}

const express = require('express');
const fetch = require('node-fetch');

const app = express();
const port = process.env.PORT || 3000;

console.log(`[${new Date().toISOString()}] Starting server...`);
console.log(`PORT=${port}`);
console.log(`NODE_ENV=${process.env.NODE_ENV}`);
console.log(`Has REACT_APP_DIRECT_LINE_SECRET: ${!!process.env.REACT_APP_DIRECT_LINE_SECRET}`);

// Cache control headers
app.use((req, res, next) => {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    next();
});

// CORS middleware
app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    if (req.method === 'OPTIONS') {
        res.sendStatus(200);
    } else {
        next();
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.status(200).json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Token generation API
app.get('/api/get-token', async (req, res) => {
    try {
        const secret = process.env.REACT_APP_DIRECT_LINE_SECRET;
        
        if (!secret) {
            console.error('Missing REACT_APP_DIRECT_LINE_SECRET in Environment Variables');
            return res.status(500).json({ error: "Server Configuration Error: Missing Secret Key" });
        }

        const response = await fetch('https://directline.botframework.com/v3/directline/tokens/generate', {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${secret}` }
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error(`Bot Framework error: ${response.status} ${errorText}`);
            throw new Error(`Bot Framework error: ${response.status}`);
        }

        const data = await response.json();
        res.json({ token: data.token });
    } catch (error) {
        console.error('Error generating token:', error);
        res.status(500).json({ error: "Failed to generate token", details: error.message });
    }
});

// Serve static files from dist
const distPath = path.join(__dirname, 'dist');
console.log(`Serving static files from: ${distPath}`);
app.use(express.static(distPath));

// Fallback to taskpane.html for SPA routing
app.get('*', (req, res) => {
    res.sendFile(path.join(distPath, 'taskpane.html'));
});

// Start server
app.listen(port, () => {
    console.log(`[${new Date().toISOString()}] Server is running on port ${port}`);
    console.log(`[${new Date().toISOString()}] Visit: http://localhost:${port}`);
}).on('error', (err) => {
    console.error(`[${new Date().toISOString()}] Server error:`, err);
    process.exit(1);
});
