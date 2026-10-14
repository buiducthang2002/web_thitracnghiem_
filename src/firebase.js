// src/firebase.js
// Cấu hình lấy từ file .env (xem .env.example)

import { initializeApp, getApps } from 'firebase/app'
import { getFirestore } from 'firebase/firestore'
import { getAuth, createUserWithEmailAndPassword, signOut } from 'firebase/auth'

const firebaseConfig = {
  apiKey:            import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain:        import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId:         import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket:     import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId:             import.meta.env.VITE_FIREBASE_APP_ID,
}

// Kiểm tra config bị thiếu
const requiredKeys = ['apiKey', 'authDomain', 'projectId', 'storageBucket', 'messagingSenderId', 'appId']
export const missingConfig = requiredKeys.filter(key => !firebaseConfig[key])
export const projectId = firebaseConfig.projectId

const app = initializeApp(firebaseConfig)
export const db = getFirestore(app)
export const auth = getAuth(app)

// Tên đăng nhập → email nội bộ cho Firebase Auth (người dùng không bao giờ thấy email này).
// Tên có ký tự lạ (dấu, khoảng trắng...) được mã hoá hex để luôn là email hợp lệ.
const AUTH_DOMAIN_SUFFIX = '@quiz.local'
export const toEmail = account => {
  const a = String(account).trim().toLowerCase()
  if (/^[a-z0-9._-]+$/.test(a)) return a + AUTH_DOMAIN_SUFFIX
  const hex = Array.from(new TextEncoder().encode(a)).map(b => b.toString(16).padStart(2, '0')).join('')
  return 'x' + hex + AUTH_DOMAIN_SUFFIX
}

// Mật khẩu ngẫu nhiên, bỏ ký tự dễ nhầm (0/O, 1/l/I)
export const genPassword = (len = 8) => {
  const chars = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'
  const buf = crypto.getRandomValues(new Uint32Array(len))
  return Array.from(buf, n => chars[n % chars.length]).join('')
}

// Tạo tài khoản đăng nhập cho thí sinh mà KHÔNG đăng xuất admin: dùng một app Firebase
// thứ hai chỉ để gọi createUser (createUser tự đăng nhập tài khoản vừa tạo vào app đó).
export const createLoginAccount = async (account, password) => {
  const secondary = getApps().find(a => a.name === 'secondary') || initializeApp(firebaseConfig, 'secondary')
  const secAuth = getAuth(secondary)
  const cred = await createUserWithEmailAndPassword(secAuth, toEmail(account), password)
  const uid = cred.user.uid
  await signOut(secAuth)
  return uid
}
