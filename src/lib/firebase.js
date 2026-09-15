// Firebase Realtime Database — single source of truth for all İSTREGISTER data.
import { initializeApp } from 'firebase/app';
import { getDatabase, ref, onValue, set, update, remove, push } from 'firebase/database';

const firebaseConfig = {
  apiKey: "AIzaSyBphO5XhddqFwaxelZnCShcJ6ZOops990Y",
  authDomain: "ogstorage-13ca0.firebaseapp.com",
  databaseURL: "https://ogstorage-13ca0-default-rtdb.firebaseio.com",
  projectId: "ogstorage-13ca0",
  storageBucket: "ogstorage-13ca0.firebasestorage.app",
  messagingSenderId: "422108490754",
  appId: "1:422108490754:web:03e6c19125ff40ef3c0b1b"
};

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

const RECORDS_PATH = 'istregister/records';

export function recordsRef() {
  return ref(db, RECORDS_PATH);
}

export function singleRecordRef(id) {
  return ref(db, `${RECORDS_PATH}/${id}`);
}

export { db, onValue, set, update, remove, push, ref };
