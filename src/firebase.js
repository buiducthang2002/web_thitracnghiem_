// src/firebase.js
// Config Firebase lấy từ file .env (Firebase Console → Project Settings → Your apps → SDK setup)

import { initializeApp } from 'firebase/app'
import { initializeFirestore } from 'firebase/firestore'

const firebaseConfig = {
  apiKey:            import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain:        import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId:         import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket:     import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId:             import.meta.env.VITE_FIREBASE_APP_ID,
}

// Thiếu config (quên .env khi build) sẽ khiến app chạy nhưng không bao giờ lưu được
// dữ liệu — báo lỗi ngay thay vì để hỏng âm thầm.
export const missingConfig = Object.entries(firebaseConfig)
  .filter(([, v]) => !v)
  .map(([k]) => k)

if (missingConfig.length) {
  console.error('[Firebase] Thiếu config:', missingConfig.join(', '),
    '— kiểm tra file .env rồi chạy lại (npm run dev) / build lại (npm run build).')
}

const app = initializeApp(firebaseConfig)

// Firestore mặc định dùng WebChannel (streaming). Nhiều mạng nội bộ / proxy / firewall
// chặn kiểu kết nối này: app vẫn chạy, đọc được cache, nhưng lệnh ghi nằm mãi trong
// hàng đợi và mất khi reload. autoDetectLongPolling tự chuyển sang long-polling
// (HTTP thường) khi phát hiện WebChannel không thông.
export const db = initializeFirestore(app, {
  experimentalAutoDetectLongPolling: true,
})

export const projectId = firebaseConfig.projectId
