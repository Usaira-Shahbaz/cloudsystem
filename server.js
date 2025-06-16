require('dotenv').config();
const express = require('express');
const multer = require('multer');
const { DefaultAzureCredential } = require('@azure/identity');
const { BlobServiceClient } = require('@azure/storage-blob');
const { TableClient } = require('@azure/data-tables');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const mime = require('mime-types');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Middleware
app.use(cors());
app.use(express.static('public'));
app.use(session({
  secret: process.env.SESSION_SECRET,
  resave: false,
  saveUninitialized: false
}));
app.use(passport.initialize());
app.use(passport.session());

// Azure Blob Storage Configuration
const accountName = process.env.AZURE_STORAGE_ACCOUNT_NAME;
const containerName = process.env.AZURE_CONTAINER_NAME;
const credential = new DefaultAzureCredential();
const blobServiceClient = new BlobServiceClient(
  `https://${accountName}.blob.core.windows.net`,
  credential
);
const containerClient = blobServiceClient.getContainerClient(containerName);

// Azure Table Storage Configuration
const tableName = process.env.AZURE_TABLE_NAME;
const tableClient = new TableClient(
  `https://${accountName}.table.core.windows.net`,
  tableName,
  credential
);

// Ensure table exists
async function initializeTable() {
  try {
    await tableClient.createTable();
    console.log(`Table ${tableName} created or already exists.`);
  } catch (err) {
    console.error('Error creating table:', err.message);
  }
}
initializeTable();

// Logging Function
async function logAction(userId, action, fileName, details = {}) {
  try {
    const entity = {
      partitionKey: userId,
      rowKey: `${Date.now()}-${Math.random().toString(36).substring(2, 15)}`,
      action,
      fileName,
      timestamp: new Date().toISOString(),
      details: JSON.stringify(details)
    };
    await tableClient.createEntity(entity);
    console.log(`Logged ${action} for file ${fileName} by user ${userId}`);
  } catch (err) {
    console.error('Logging error:', err.message);
  }
}

// Passport Azure AD Configuration
passport.use(new OIDCStrategy({
  identityMetadata: `https://login.microsoftonline.com/${process.env.AZURE_AD_TENANT_ID}/v2.0/.well-known/openid-configuration`,
  clientID: process.env.AZURE_AD_CLIENT_ID,
  clientSecret: process.env.AZURE_AD_CLIENT_SECRET,
  responseType: 'code',
  responseMode: 'query',
  redirectUrl: `https://${process.env.WEBSITE_HOSTNAME}/auth/redirect`, // Use HTTPS for Azure
  allowHttpForRedirectUrl: false,
  scope: ['profile', 'email', 'openid']
}, function (iss, sub, profile, accessToken, refreshToken, done) {
  if (!profile.oid) return done(new Error("No OID found"), null);
  return done(null, profile);
}));

passport.serializeUser((user, done) => done(null, user));
passport.deserializeUser((user, done) => done(null, user));

// Authentication Middleware
function ensureAuthenticated(req, res, next) {
  if (req.isAuthenticated()) return next();
  res.redirect('/login');
}

// Auth Routes
app.get('/login', passport.authenticate('azuread-openidconnect'));
app.get('/auth/redirect',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
  (req, res) => res.redirect('/')
);
app.get('/logout', (req, res) => {
  req.logout(err => {
    const redirectUrl = `https://${process.env.WEBSITE_HOSTNAME}/logout`;
    res.redirect(`https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=${redirectUrl}`);
  });
});
app.get('/me', (req, res) => {
  if (!req.isAuthenticated()) return res.status(401).send('Not logged in');
  res.send(req.user);
});

// Upload Multiple Files
app.post('/upload', ensureAuthenticated, upload.array('files'), async (req, res) => {
  try {
    const uploadResults = [];

    for (const file of req.files) {
      const blobName = file.originalname;
      const blockBlobClient = containerClient.getBlockBlobClient(blobName);
      await blockBlobClient.uploadFile(file.path);
      fs.unlinkSync(file.path);
      uploadResults.push(blobName);
      // Log upload action
      await logAction(req.user.oid, 'upload', blobName, { fileSize: file.size });
    }

    res.send({ message: `Uploaded ${uploadResults.length} files.`, files: uploadResults });
  } catch (err) {
    console.error('Upload error:', err.message);
    res.status(500).send({ error: err.message });
  }
});

// List Files
app.get('/files', ensureAuthenticated, async (req, res) => {
  try {
    const files = [];
    for await (const blob of containerClient.listBlobsFlat()) {
      files.push(blob.name);
    }
    res.send(files);
  } catch (err) {
    console.error('List error:', err.message);
    res.status(500).send({ error: err.message });
  }
});

// Download File
app.get('/download/:filename', ensureAuthenticated, async (req, res) => {
  try {
    const blobClient = containerClient.getBlobClient(req.params.filename);
    const downloadBlockBlobResponse = await blobClient.download();
    res.setHeader('Content-Disposition', `attachment; filename=${req.params.filename}`);
    downloadBlockBlobResponse.readableStreamBody.pipe(res);
    // Log download action
    await logAction(req.user.oid, 'download', req.params.filename);
  } catch (err) {
    console.error('Download error:', err.message);
    res.status(500).send({ error: err.message });
  }
});

// Delete File
app.delete('/delete/:filename', ensureAuthenticated, async (req, res) => {
  try {
    const blobClient = containerClient.getBlobClient(req.params.filename);
    await blobClient.deleteIfExists();
    // Log delete action
    await logAction(req.user.oid, 'delete', req.params.filename);
    res.send({ message: 'Deleted' });
  } catch (err) {
    console.error('Delete error:', err.message);
    res.status(500).send({ error: err.message });
  }
});

// Update File (Re-upload)
app.post('/update', ensureAuthenticated, upload.single('file'), async (req, res) => {
  try {
    const blobName = req.file.originalname;
    const blockBlobClient = containerClient.getBlockBlobClient(blobName);
    await blockBlobClient.uploadFile(req.file.path, { overwrite: true });
    fs.unlinkSync(req.file.path);
    // Log update action
    await logAction(req.user.oid, 'update', blobName, { fileSize: req.file.size });
    res.send({ message: 'Updated' });
  } catch (err) {
    console.error('Update error:', err.message);
    res.status(500).send({ error: err.message });
  }
});

// File Preview
app.get('/preview/:filename', ensureAuthenticated, async (req, res) => {
  try {
    const blobClient = containerClient.getBlobClient(req.params.filename);
    const downloadResponse = await blobClient.download();
    const contentType = blobClient.name.match(/\.(jpg|jpeg|png|gif)$/i)
      ? 'image/' + blobClient.name.split('.').pop()
      : blobClient.name.match(/\.(txt|md|log|csv|json|js|html|css)$/i)
        ? 'text/plain'
        : 'application/octet-stream';

    res.setHeader('Content-Type', contentType);
    downloadResponse.readableStreamBody.pipe(res);
    // Log preview action
    await logAction(req.user.oid, 'preview', req.params.filename);
  } catch (err) {
    console.error('Preview error:', err.message);
    res.status(500).send({ error: err.message });
  }
});

app.get('/logs', ensureAuthenticated, async (req, res) => {
  try {
    const entities = [];
    for await (const entity of tableClient.listEntities({ queryOptions: { filter: `PartitionKey eq '${req.user.oid}'` } })) {
      entities.push(entity);
    }
    res.json(entities);
  } catch (err) {
    console.error('Log fetch error:', err.message);
    res.status(500).send({ error: err.message });
  }
});

// Analytics
app.get('/analytics', ensureAuthenticated, async (req, res) => {
  try {
    const fileCategories = { Images: 0, Text: 0, Others: 0 };
    let totalSize = 0;
    let fileCount = 0;

    for await (const blob of containerClient.listBlobsFlat()) {
      fileCount++;
      totalSize += blob.properties.contentLength || 0;

      const ext = path.extname(blob.name).toLowerCase();

      if (ext.match(/\.(jpg|jpeg|png|gif)$/i)) {
        fileCategories.Images++;
      } else if (ext.match(/\.(txt|md|log|csv|json|js|html|css)$/i)) {
        fileCategories.Text++;
      } else {
        fileCategories.Others++;
      }
    }

    res.json({
      totalFiles: fileCount,
      totalSizeMB: (totalSize / (1024 * 1024)).toFixed(2),
      fileCategories
    });
  } catch (err) {
    console.error('Analytics error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Start server
const port = process.env.PORT || 3000;
app.listen(port, '0.0.0.0', () => {
  console.log(`Server running on port ${port}`);
});
