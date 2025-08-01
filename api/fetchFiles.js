// Converted to CommonJS require syntax
const { initializeApp } = require('firebase/app');
const { getFirestore, doc, setDoc, getDoc } = require('firebase/firestore');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');

// Initialize Firebase
const firebaseConfig = {
  apiKey: process.env.FIREBASE_API_KEY || "mock-key-for-dev",
  authDomain: process.env.FIREBASE_AUTH_DOMAIN,
  projectId: process.env.FIREBASE_PROJECT_ID,
  storageBucket: process.env.FIREBASE_STORAGE_BUCKET || "",
  messagingSenderId: process.env.FIREBASE_MESSAGING_SENDER_ID || "",
  appId: process.env.FIREBASE_APP_ID
};

// Initialize Firebase only if config is valid
let db;
try {
  if (!firebaseConfig.apiKey || !firebaseConfig.projectId || !firebaseConfig.appId) {
    throw new Error("Missing required Firebase config");
  }
  const app = initializeApp(firebaseConfig);
  db = getFirestore(app);
} catch (firebaseError) {
  console.error("Firebase initialization failed:", firebaseError);
}

module.exports = async function handler(req, res) {
  // Early return if Firebase failed to initialize
  if (!db) {
    return res.status(500).json({
      error: 'Server configuration error',
      details: 'Firebase not initialized'
    });
  }

  try {
    // 1. Check Firestore cache first
    const cacheRef = doc(db, 'cache', 'files');
    const cacheSnap = await getDoc(cacheRef);
    
    if (cacheSnap.exists()) {
      const cacheData = cacheSnap.data();
      if (cacheData.expiry > Date.now()) {
        return res.status(200).json({
          files: cacheData.files,
          source: 'cache'
        });
      }
    }

    // 2. Initialize Microsoft Graph Client
    if (!process.env.AZURE_TENANT_ID || !process.env.AZURE_CLIENT_ID || !process.env.AZURE_CLIENT_SECRET) {
      throw new Error("Missing Azure AD credentials");
    }

    const credential = new ClientSecretCredential(
      process.env.AZURE_TENANT_ID,
      process.env.AZURE_CLIENT_ID,
      process.env.AZURE_CLIENT_SECRET
    );

    const client = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          const token = await credential.getToken("https://graph.microsoft.com/.default");
          return token.token;
        }
      }
    });

    // 3. Fetch files from OneDrive
    const response = await client.api('/me/drive/root/children')
      .select('id,name,size,lastModifiedDateTime,@microsoft.graph.downloadUrl')
      .get();

    const files = response.value.map((item) => ({
      id: item.id,
      name: item.name,
      '@microsoft.graph.downloadUrl': item['@microsoft.graph.downloadUrl'],
      lastModifiedDateTime: item.lastModifiedDateTime,
      size: item.size
    }));

    // 4. Cache in Firestore (1-hour expiry)
    await setDoc(cacheRef, {
      files,
      expiry: Date.now() + 3600000 // 1 hour
    });

    // 5. Return files to frontend
    res.status(200).json({
      files,
      source: 'graph-api'
    });

  } catch (error) {
    console.error('Error fetching files:', error);
    res.status(500).json({
      error: 'Failed to fetch files',
      details: error instanceof Error ? error.message : String(error),
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
};
