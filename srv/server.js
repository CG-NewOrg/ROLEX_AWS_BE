const cds = require('@sap/cds');
const express = require('express');
const xssProtection = require('./middleware/xssProtection');
require('dotenv').config();

cds.on('bootstrap', app => {
  const customApp = express();

  customApp.use((req, res, next) => {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-XSS-Protection', '1; mode=block');
    res.setHeader('Content-Security-Policy', "default-src 'self'");
    res.setHeader('X-Frame-Options', 'DENY');
    next();
  });

  customApp.use(express.json({ limit: '20mb' }));
  customApp.use(express.urlencoded({ limit: '20mb', extended: true }));
  customApp.use(xssProtection);
                                               
  customApp.use((req, res, next) => {
    const encodeEntities = (obj) => {
      if (typeof obj === 'string') {
        return obj.replace(/[<>"'&]/g, (char) => {
          const map = {
            '<':  '&lt;',
            '>':  '&gt;',
            '"':  '&quot;',
            "'":  '&#x27;',
            '&':  '&amp;'
          };
          return map[char];
        });
      }
      if (Array.isArray(obj)) return obj.map(encodeEntities);
      if (typeof obj === 'object' && obj !== null) {
        const out = {};
        for (const key in obj) out[key] = encodeEntities(obj[key]);
        return out;
      }
      return obj;
    };

    if (req.body)   req.body   = encodeEntities(req.body);
    if (req.query)  req.query  = encodeEntities(req.query);
    if (req.params) req.params = encodeEntities(req.params);
    next();
  });

  cds.serve('all').in(customApp);
  app.use(customApp);
});

module.exports = cds.server;