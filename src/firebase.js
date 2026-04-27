import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyC63Gt1ygbCmygAkEtg2QKfW2IBoJUVSDU",
  authDomain: "khalid-realestate.firebaseapp.com",
  projectId: "khalid-realestate",
  storageBucket: "khalid-realestate.firebasestorage.app",
  messagingSenderId: "818523254211",
  appId: "1:818523254211:web:9bc328ed83803ebfd6e176"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);