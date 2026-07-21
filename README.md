# 🏥 Cục hậu cần kỹ thuật quân khu 4 — Hệ thống thi trắc nghiệm

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

## 🗄️ Cấu trúc dữ liệu trên Firestore

App dùng **4 collection**, mỗi document có `id` trùng với tên document (dạng chuỗi):

| Collection | Mô tả | Các trường |
|---|---|---|
| `employees` | Thí sinh | `id`, `name`, `dept` |
| `questions` | Ngân hàng câu hỏi | `id`, `topic`, `level` (`Dễ`/`TB`/`Khó`), `text`, `opts[]`, `ans` (index đáp án đúng, tính từ 0) |
| `exams` | Đề thi | `id`, `title`, `topic`, `qIds[]`, `time` (phút), `pass` (% điểm đạt), `desc` |
| `results` | Kết quả bài làm | `id`, `empId`, `examId`, `score` (%), `correct`, `timeTaken` (giây), `date`, `answers[]` |

Không cần tạo collection thủ công — app tự tạo khi ghi document đầu tiên.

### Nạp dữ liệu mẫu (tuỳ chọn)

Sau khi đã điền `.env`, chạy:

```bash
npm run seed
```

> ⚠️ Lệnh này **xoá sạch** dữ liệu cũ trong cả 4 collection rồi ghi lại dữ liệu mẫu trong `seed.mjs`. Chỉ chạy khi mới cài đặt.

---

## 🔒 Bảo mật Firestore (quan trọng!)

Rules đã được viết sẵn trong file **`firestore.rules`**. Có 2 cách áp dụng:

**Cách 1 — Dán thủ công (đơn giản nhất)**
1. Mở file `firestore.rules` trong project
2. Copy toàn bộ nội dung
3. Vào **Firebase Console → Firestore Database → Rules** → dán đè lên → **Publish**

**Cách 2 — Deploy bằng Firebase CLI**
```bash
npm install -g firebase-tools
firebase login
firebase use --add          # chọn project vừa tạo
firebase deploy --only firestore:rules
```

> ⚠️ **Lưu ý:** Rules hiện tại cho phép đọc/ghi không cần đăng nhập (vì app chưa dùng Firebase Authentication), nhưng đã chặn mọi collection lạ và kiểm tra kiểu dữ liệu. Nếu cần bảo mật cao hơn, xem phần ghi chú ở cuối file `firestore.rules`.

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
├── firebase.json        # Cấu hình Firebase CLI
├── firestore.rules      # Quy tắc bảo mật Firestore
├── firestore.indexes.json
├── seed.mjs             # Script nạp dữ liệu mẫu (npm run seed)
├── index.html
├── package.json
├── vite.config.js
└── README.md
```
