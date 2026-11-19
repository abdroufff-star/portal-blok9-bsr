import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

// GANTI dengan konfigurasi dari Firebase Console (Project Settings > Your apps)
const firebaseConfig = {
  apiKey: "AIzaSyBSOXwf3NvkAu7CzJ6bcqKsevBaAdjmUzc",
  authDomain: "portal-blok-9bsr.firebaseapp.com",
  projectId: "portal-blok-9bsr",
  storageBucket: "portal-blok-9bsr.firebasestorage.app",
  messagingSenderId: "862617676728",
  appId: "1:862617676728:web:7afdfe86aa1699a7d3fd4d",
  measurementId: "G-SVS8BJE7FM"
};


// App Utama
const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);

// Secondary App (Khusus agar Admin tidak ter-logout saat membuat akun warga)
const secondaryApp = initializeApp(firebaseConfig, "SecondaryApp");
export const secondaryAuth = getAuth(secondaryApp);