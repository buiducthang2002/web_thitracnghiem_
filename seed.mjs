// seed.mjs — Chạy 1 lần để đẩy dữ liệu ban đầu lên Firebase
// Cách chạy: node seed.mjs

import { initializeApp } from "firebase/app";
import { getFirestore, doc, setDoc, collection, getDocs, deleteDoc } from "firebase/firestore";
import { readFileSync } from "fs";

// Đọc .env thủ công
const env = {};
try {
  const envFile = readFileSync(".env", "utf-8");
  envFile.split("\n").forEach(line => {
    const [k, ...v] = line.split("=");
    if (k && v.length) env[k.trim()] = v.join("=").trim();
  });
} catch {
  console.error("❌ Không tìm thấy file .env");
  process.exit(1);
}

const firebaseConfig = {
  apiKey:            env.VITE_FIREBASE_API_KEY,
  authDomain:        env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId:         env.VITE_FIREBASE_PROJECT_ID,
  storageBucket:     env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId:             env.VITE_FIREBASE_APP_ID,
};

const app = initializeApp(firebaseConfig);
const db  = getFirestore(app);

// ── Dữ liệu mẫu ──────────────────────────────────────────

const EMPLOYEES = [
  {id:1,name:"Nguyễn Văn An",dept:"Kinh doanh"},
  {id:2,name:"Trần Thị Bình",dept:"Kế toán"},
  {id:3,name:"Lê Văn Cường",dept:"Kỹ thuật"},
  {id:4,name:"Phạm Thị Dung",dept:"Nhân sự"},
  {id:5,name:"Hoàng Văn Em",dept:"Kinh doanh"},
];

const QUESTIONS = [
  {id:1,topic:"An toàn lao động",level:"TB",text:"Khi phát hiện sự cố về điện, việc đầu tiên cần làm là?",opts:["Tự sửa chữa","Ngắt nguồn và báo kỹ thuật","Bỏ qua nếu nhỏ","Chờ sự cố tự hết"],ans:1},
  {id:2,topic:"An toàn lao động",level:"Dễ",text:"Lối thoát hiểm trong trường hợp khẩn cấp phải như thế nào?",opts:["Khóa để bảo mật","Luôn thông thoáng, không bị chặn","Chỉ mở khi có sự cố","Để vật dụng tạm thời"],ans:1},
  {id:3,topic:"An toàn lao động",level:"Dễ",text:"Khi xảy ra hỏa hoạn trong tòa nhà, thí sinh phải?",opts:["Dùng thang máy cho nhanh","Dùng cầu thang bộ, không dùng thang máy","Chờ hỏa hoạn tắt","Gọi điện hỏi ý kiến"],ans:1},
  {id:4,topic:"An toàn lao động",level:"TB",text:"Bình chữa cháy cần được kiểm tra định kỳ bao lâu một lần?",opts:["1 năm/lần","3 năm/lần","Khi nào hỏng thì kiểm tra","5 năm/lần"],ans:0},
  {id:5,topic:"An toàn lao động",level:"Khó",text:"Khi đồng nghiệp bị ngã, bất tỉnh, cần làm gì đầu tiên?",opts:["Tự xử lý","Lay mạnh để người đó tỉnh","Gọi cấp cứu và sơ cứu an toàn","Bỏ qua nếu còn thở"],ans:2},
];

const EXAMS = [
  {id:1,title:"Kiểm tra An toàn lao động",topic:"An toàn lao động",qIds:[1,2,3,4,5],time:15,pass:80,desc:"Bài kiểm tra an toàn lao động bắt buộc"},
];

const RESULTS = [];

// ── Seed function ─────────────────────────────────────────

async function clearCol(name) {
  const snap = await getDocs(collection(db, name));
  await Promise.all(snap.docs.map(d => deleteDoc(d.ref)));
  console.log(`  🗑  Đã xóa ${snap.docs.length} docs cũ trong "${name}"`);
}

async function seedCol(name, items) {
  await clearCol(name);
  await Promise.all(items.map(item => setDoc(doc(db, name, String(item.id)), item)));
  console.log(`  ✅ Đã thêm ${items.length} docs vào "${name}"`);
}

async function main() {
  console.log("\n🚀 Bắt đầu seed dữ liệu lên Firebase...\n");
  try {
    await seedCol("employees", EMPLOYEES);
    await seedCol("questions", QUESTIONS);
    await seedCol("exams",     EXAMS);
    await seedCol("results",   RESULTS);
    console.log("\n🎉 Seed xong! Mọi dữ liệu đã lên Firebase.");
    console.log("   Reload app tại http://localhost:5173 để xem.\n");
  } catch (err) {
    console.error("❌ Lỗi:", err.message);
  }
  process.exit(0);
}

main();
