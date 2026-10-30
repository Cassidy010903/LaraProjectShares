importScripts("https://www.gstatic.com/firebasejs/9.19.1/firebase-app-compat.js");
importScripts("https://www.gstatic.com/firebasejs/9.19.1/firebase-messaging-compat.js");

firebase.initializeApp({
    apiKey: "AIzaSyBZKCvH6uwREo42lgjZ_9uUeXjphibFT1c",
    authDomain: "vleissentraal-98271.firebaseapp.com",
    projectId: "vleissentraal-98271",
    storageBucket: "vleissentraal-98271.firebasestorage.app",
    messagingSenderId: "777330577028",
    appId: "1:777330577028:web:3195b1a4e7c2a8dea4bcd1",
    measurementId: "G-GECNL7HPG0"
});
// Necessary to receive background messages:
const messaging = firebase.messaging();
messaging.onBackgroundMessage(async (payload) => {
  console.log('Received background message:', payload);

  const notificationTitle = payload.notification.title;
  const notificationOptions = {
    body: payload.notification.body,
    icon: '/icons/Icon-192.png',
    badge: '/icons/Icon-192.png',
    data: payload.data,
    click_action: payload.notification.click_action,
  };

  self.registration.showNotification(notificationTitle, notificationOptions);
});