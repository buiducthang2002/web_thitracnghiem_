// scripts/adminTools.mjs — công cụ quản trị chạy trên máy admin (dùng firebase-admin, bỏ qua rules)
//
// Chuẩn bị 1 lần: Firebase Console → ⚙️ Project settings → Service accounts → Generate new private key
//   → lưu file thành serviceAccountKey.json ở thư mục gốc dự án (file này đã nằm trong .gitignore, KHÔNG commit).
//
// Lệnh:
//   node scripts/adminTools.mjs create-admin <tài khoản> <mật khẩu>   tạo admin đầu tiên (mật khẩu >= 6 ký tự)
//   node scripts/adminTools.mjs reset <tài khoản> [mật khẩu]          đặt lại mật khẩu (bỏ trống = tự sinh ngẫu nhiên)

import { readFileSync } from 'fs'
import { randomInt } from 'crypto'
import { initializeApp, cert } from 'firebase-admin/app'
import { getAuth } from 'firebase-admin/auth'
import { getFirestore } from 'firebase-admin/firestore'

// PHẢI giống hàm toEmail trong src/firebase.js
const toEmail = account => {
  const a = String(account).trim().toLowerCase()
  if (/^[a-z0-9._-]+$/.test(a)) return a + '@quiz.local'
  const hex = Array.from(new TextEncoder().encode(a)).map(b => b.toString(16).padStart(2, '0')).join('')
  return 'x' + hex + '@quiz.local'
}
const genPassword = (len = 8) => {
  const chars = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'
  return Array.from({ length: len }, () => chars[randomInt(chars.length)]).join('')
}

const keyPath = process.env.GOOGLE_APPLICATION_CREDENTIALS || 'serviceAccountKey.json'
let key
try { key = JSON.parse(readFileSync(keyPath, 'utf8')) }
catch { console.error(`❌ Không đọc được ${keyPath}. Xem hướng dẫn đầu file.`); process.exit(1) }

initializeApp({ credential: cert(key) })
const auth = getAuth()
const db = getFirestore()

const [cmd, account, pw] = process.argv.slice(2)

try {
  if (cmd === 'create-admin') {
    if (!account || !pw) throw new Error('Cú pháp: create-admin <tài khoản> <mật khẩu>')
    if (pw.length < 6) throw new Error('Mật khẩu phải có ít nhất 6 ký tự')
    const user = await auth.createUser({ email: toEmail(account), password: pw })
    await db.doc(`users/${user.uid}`).set({ role: 'admin', account })
    console.log(`✅ Đã tạo admin "${account}". Đăng nhập vào web bằng tài khoản này.`)
  } else if (cmd === 'reset') {
    if (!account) throw new Error('Cú pháp: reset <tài khoản> [mật khẩu]')
    const newPw = pw || genPassword()
    const user = await auth.getUserByEmail(toEmail(account))
    await auth.updateUser(user.uid, { password: newPw })
    console.log(`✅ Mật khẩu mới của "${account}": ${newPw}`)
  } else {
    throw new Error('Lệnh không hợp lệ. Dùng: create-admin | reset')
  }
  process.exit(0)
} catch (err) {
  console.error('❌', err.message)
  process.exit(1)
}
