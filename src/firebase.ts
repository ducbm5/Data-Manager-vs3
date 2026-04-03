import { initializeApp } from 'firebase/app';
import { getAuth } from 'firebase/auth';
import { getFirestore } from 'firebase/firestore';
import defaultConfig from '../firebase-applet-config.json';

// Support dynamic configuration from localStorage
const getInitialConfig = () => {
  if (typeof window !== 'undefined') {
    const saved = localStorage.getItem('custom_firebase_config');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error("Failed to parse saved config", e);
      }
    }
  }
  return defaultConfig;
};

export const firebaseConfig = getInitialConfig();
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app, firebaseConfig.firestoreDatabaseId);
export const auth = getAuth(app);
