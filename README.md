# 🏥 Bệnh viện Quân y 4 — Hệ thống thi trắc nghiệm

Hướng dẫn cài đặt và deploy để mọi người cùng sử dụng.

---

## 📋 Yêu cầu

- Máy tính có kết nối Internet
- Tài khoản Google (để tạo Firebase)
- Tài khoản GitHub (để deploy Vercel)

---

## BƯỚC 1 — Tạo Firebase (database miễn phí)

### 1.1 Tạo project Firebase
1. Truy cập [https://console.firebase.google.com](https://console.firebase.google.com)
2. Nhấn **"Add project"** → đặt tên (VD: `bvqy4-quiz`) → Next → Next → Create project
3. Chờ tạo xong → nhấn **Continue**

### 1.2 Tạo Firestore Database
1. Trong Firebase Console → menu trái chọn **"Firestore Database"**
2. Nhấn **"Create database"**
3. Chọn **"Start in test mode"** → Next
4. Chọn location **`asia-southeast1 (Singapore)`** → **Enable**

### 1.3 Lấy thông tin cấu hình
1. Nhấn ⚙️ (bánh răng) → **Project Settings**
2. Kéo xuống mục **"Your apps"** → nhấn biểu tượng **`</>`** (Web)
3. Đặt tên app (VD: `quiz-web`) → **Register app**
4. Sao chép đoạn `firebaseConfig` — trông như thế này:
```js
const firebaseConfig = {
  apiKey: "AIza...",
  authDomain: "bvqy4-quiz.firebaseapp.com",
  projectId: "bvqy4-quiz",
  storageBucket: "bvqy4-quiz.appspot.com",
  messagingSenderId: "123456789",
  appId: "1:123456789:web:abc123"
};
```

---

## BƯỚC 2 — Cài đặt project trên máy

### 2.1 Cài Node.js (nếu chưa có)
- Tải tại [https://nodejs.org](https://nodejs.org) → chọn **LTS** → cài đặt

### 2.2 Giải nén project
- Giải nén file `quizpro.zip` vào một thư mục (VD: `Desktop/quizpro`)

### 2.3 Tạo file cấu hình
1. Trong thư mục project, sao chép file `.env.example` → đổi tên thành `.env`
2. Mở `.env` bằng Notepad/VSCode, điền thông tin Firebase:

```
VITE_FIREBASE_API_KEY=AIza...
VITE_FIREBASE_AUTH_DOMAIN=bvqy4-quiz.firebaseapp.com
VITE_FIREBASE_PROJECT_ID=bvqy4-quiz
VITE_FIREBASE_STORAGE_BUCKET=bvqy4-quiz.appspot.com
VITE_FIREBASE_MESSAGING_SENDER_ID=123456789
VITE_FIREBASE_APP_ID=1:123456789:web:abc123
```

### 2.4 Cài dependencies và chạy thử
Mở **Terminal / Command Prompt** trong thư mục project:

```bash
npm install
npm run dev
```

Mở trình duyệt vào `http://localhost:5173` — app chạy là thành công ✅

---

## BƯỚC 3 — Deploy lên Vercel (hosting miễn phí)

### 3.1 Đưa code lên GitHub
1. Tạo tài khoản [https://github.com](https://github.com)
2. Tạo **New repository** → đặt tên `bvqy4-quiz` → Private → Create
3. Trong Terminal:
```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/TEN_BAN/bvqy4-quiz.git
git push -u origin main
```

### 3.2 Deploy trên Vercel
1. Truy cập [https://vercel.com](https://vercel.com) → **Sign up with GitHub**
2. Nhấn **"Add New Project"** → chọn repository `bvqy4-quiz`
3. Mở mục **"Environment Variables"** → thêm từng biến từ file `.env`:

| Name | Value |
|------|-------|
| VITE_FIREBASE_API_KEY | AIza... |
| VITE_FIREBASE_AUTH_DOMAIN | ... |
| VITE_FIREBASE_PROJECT_ID | ... |
| VITE_FIREBASE_STORAGE_BUCKET | ... |
| VITE_FIREBASE_MESSAGING_SENDER_ID | ... |
| VITE_FIREBASE_APP_ID | ... |

4. Nhấn **Deploy** → chờ 1-2 phút

5. Vercel sẽ cấp cho bạn link dạng: `https://bvqy4-quiz.vercel.app`  
   → Chia sẻ link này cho thí sinh là xong! 🎉

---

## BƯỚC 4 — Thêm dữ liệu ban đầu

Sau khi deploy xong, đăng nhập bằng tài khoản Admin:
- **Tài khoản:** `admin`
- **Mật khẩu:** `123`

Vào các mục để thêm:
- **Thí sinh** → Thêm thủ công hoặc Import Excel
- **Câu hỏi** → Thêm thủ công hoặc Import Word
- **Đề thi** → Tạo đề thi từ ngân hàng câu hỏi

---

## 🔒 Bảo mật Firestore (quan trọng!)

Sau khi test xong, vào **Firestore → Rules** và thay bằng:

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if true;  // Tạm thời - cần thêm auth sau
    }
  }
}
```

> ⚠️ **Lưu ý:** Rule trên cho phép mọi người đọc/ghi. Nếu cần bảo mật cao hơn, hãy tích hợp Firebase Authentication.

---

## 📞 Hỗ trợ

Nếu gặp lỗi trong quá trình cài đặt, chụp màn hình lỗi và liên hệ bộ phận IT.

---

## 📁 Cấu trúc thư mục

```
quizpro/
├── src/
│   ├── App.jsx          # Toàn bộ ứng dụng
│   ├── firebase.js      # Kết nối Firebase
│   ├── main.jsx         # Entry point
│   └── index.css        # Styles
├── .env.example         # Mẫu biến môi trường
├── .env                 # ⚠️ KHÔNG commit file này lên GitHub
├── index.html
├── package.json
├── vite.config.js
└── README.md
```
