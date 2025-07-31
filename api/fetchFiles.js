import { initializeApp } from 'firebase/app';
import { getFirestore, doc, setDoc, getDoc } from 'firebase/firestore';
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';

// Type definitions
interface DriveItem {
  id: string;
  name: string;
  '@microsoft.graph.downloadUrl'?: string;
  lastModifiedDateTime: string;
  size: number;
}

interface CachedFiles {
  files: DriveItem[];
  expiry: number;
}

// Initialize Firebase
const firebaseConfig = {
  apiKey: process.env.FIREBASE_API_KEY,
  authDomain: process.env.FIREBASE_AUTH_DOMAIN,
  projectId: process.env.FIREBASE_PROJECT_ID,
  storageBucket: process.env.FIREBASE_STORAGE_BUCKET,
  messagingSenderId: process.env.FIREBASE_MESSAGING_SENDER_ID,
  appId: process.env.FIREBASE_APP_ID
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

export default async function handler(req, res) {
  try {
    // 1. Check Firestore cache first
    const cacheRef = doc(db, 'cache', 'files');
    const cacheSnap = await getDoc(cacheRef);
    
    if (cacheSnap.exists()) {
      const cacheData = cacheSnap.data() as CachedFiles;
      if (cacheData.expiry > Date.now()) {
        return res.status(200).json({
          files: cacheData.files,
          source: 'cache'
        });
      }
    }

    // 2. Initialize Microsoft Graph Client
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

    const files: DriveItem[] = response.value.map((item: any) => ({
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
    } as CachedFiles);

    // 5. Return files to frontend
    res.status(200).json({
      files,
      source: 'graph-api'
    });

  } catch (error) {
    console.error('Error fetching files:', error);
    res.status(500).json({
      error: 'Failed to fetch files',
      details: error instanceof Error ? error.message : String(error)
    });
  }
}