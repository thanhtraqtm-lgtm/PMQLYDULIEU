import React, { createContext, useContext, useState, useEffect } from "react";
import { 
  getAuth, 
  onAuthStateChanged, 
  signInWithEmailAndPassword, 
  createUserWithEmailAndPassword, 
  signOut, 
  User 
} from "firebase/auth";
import { initializeApp, getApps, initializeApp as initFirebaseApp } from "firebase/app";
import { getFirestore, doc, getDoc, setDoc } from "firebase/firestore";

// ==========================================
// CẤU HÌNH FIREBASE (FIREBASE CONFIGURATION)
// ==========================================
// THAY THẾ CÁC THÔNG TIN DƯỚI ĐÂY BẰNG CẤU HÌNH THỰC TẾ TỪ FIREBASE CONSOLE CỦA BẠN:
// REPLACE THESE VALUES WITH YOUR REAL FIREBASE PROJECTS CONFIG FROM FIREBASE CONSOLE:
export const firebaseConfig = {
  apiKey: "YOUR_FIREBASE_API_KEY",
  authDomain: "YOUR_PROJECT_ID.firebaseapp.com",
  projectId: "YOUR_PROJECT_ID",
  storageBucket: "YOUR_PROJECT_ID.appspot.com",
  messagingSenderId: "YOUR_SENDER_ID",
  appId: "YOUR_APP_ID"
};

// Khởi tạo Firebase
let firebaseApp: any = null;
let auth: any = null;
let db: any = null;
let isFirebaseInitialized = false;

try {
  // Chỉ khởi tạo nếu cấu hình hợp lệ và không chứa các giá trị giữ chỗ mặc định
  if (firebaseConfig.apiKey && !firebaseConfig.apiKey.includes("YOUR_")) {
    if (getApps().length === 0) {
      firebaseApp = initFirebaseApp(firebaseConfig);
    } else {
      firebaseApp = getApps()[0];
    }
    auth = getAuth(firebaseApp);
    db = getFirestore(firebaseApp);
    isFirebaseInitialized = true;
    console.log("Firebase initialized successfully inside AuthContext.");
  } else {
    console.warn(
      "Firebase Config chưa được cấu hình thực tế. Đang chuyển sang chế độ Mô Phỏng (Mock Mode) để chạy mượt mà ngay lập tức."
    );
  }
} catch (error) {
  console.error("Lỗi khởi tạo Firebase:", error);
}

// Định nghĩa kiểu dữ liệu cho User Session của hệ thống
export interface SystemUser {
  uid: string;
  email: string | null;
  unitID: string;
  role: "admin" | "user";
  displayName: string;
  isMock?: boolean;
}

interface AuthContextType {
  user: SystemUser | null;
  loading: boolean;
  isFirebaseMode: boolean;
  login: (email: string, password: string, mockUnitID?: string) => Promise<void>;
  register: (email: string, password: string, unitID: string, displayName: string, role: "admin" | "user") => Promise<void>;
  logout: () => Promise<void>;
}

const AuthContext = createContext<AuthContextType | null>(null);

// Hàm tự động xác định unitID và Role từ email
export function determineUnitInfo(email: string | null, uid: string): { unitID: string; role: "admin" | "user"; displayName: string } {
  if (!email) {
    return { unitID: `unit_${uid.slice(0, 5)}`, role: "user", displayName: "Người dùng Khách" };
  }

  const normalized = email.toLowerCase().trim();
  
  // 1. Nếu là tài khoản Admin
  if (normalized.startsWith("admin") || normalized === "thanhtraqtm@gmail.com") {
    return { unitID: "admin_central", role: "admin", displayName: "Quản trị viên Trung ương" };
  }

  // 2. Nếu email có định dạng donvi_xxx@domain.com -> unitID là xxx
  const donviMatch = normalized.match(/^donvi_([a-zA-Z0-9_-]+)@/);
  if (donviMatch && donviMatch[1]) {
    return { 
      unitID: donviMatch[1], 
      role: "user", 
      displayName: `Đơn vị ${donviMatch[1].toUpperCase()}` 
    };
  }

  // 3. Phân tách theo phần tên email trước ký tự @ làm unitID mặc định
  const username = normalized.split("@")[0];
  return { 
    unitID: username, 
    role: "user", 
    displayName: `Đơn vị ${username.toUpperCase()}` 
  };
}

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<SystemUser | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  // Đọc phiên mock user đã lưu từ localStorage khi khởi động
  useEffect(() => {
    if (isFirebaseInitialized && auth) {
      const unsubscribe = onAuthStateChanged(auth, (firebaseUser: User | null) => {
        if (firebaseUser) {
          const info = determineUnitInfo(firebaseUser.email, firebaseUser.uid);
          setUser({
            uid: firebaseUser.uid,
            email: firebaseUser.email,
            unitID: info.unitID,
            role: info.role,
            displayName: firebaseUser.displayName || info.displayName,
            isMock: false
          });
        } else {
          setUser(null);
        }
        setLoading(false);
      });
      return unsubscribe;
    } else {
      // Mock mode: khôi phục từ localStorage nếu có
      const stored = localStorage.getItem("system_auth_user");
      if (stored) {
        try {
          setUser(JSON.parse(stored));
        } catch {
          setUser(null);
        }
      }
      setLoading(false);
    }
  }, []);

  const login = async (email: string, password: string, mockUnitID?: string) => {
    setLoading(true);
    try {
      if (isFirebaseInitialized && auth) {
        // Đăng nhập Firebase thực tế
        const userCredential = await signInWithEmailAndPassword(auth, email, password);
        const fbUser = userCredential.user;
        const info = determineUnitInfo(fbUser.email, fbUser.uid);
        setUser({
          uid: fbUser.uid,
          email: fbUser.email,
          unitID: info.unitID,
          role: info.role,
          displayName: fbUser.displayName || info.displayName,
          isMock: false
        });
      } else {
        // Đăng nhập mô phỏng nếu chưa cấu hình Firebase
        const cleanEmail = email.toLowerCase().trim();
        const info = determineUnitInfo(cleanEmail, "mock_uid_" + Date.now());
        
        // Cho phép ghi đè unitID khi đăng nhập mô phỏng để dễ test các phòng ban khác nhau
        const finalUnitID = mockUnitID ? mockUnitID.trim().toLowerCase() : info.unitID;
        const finalDisplayName = mockUnitID ? `Đơn vị ${mockUnitID.toUpperCase()}` : info.displayName;
        const finalRole = (cleanEmail.startsWith("admin") || mockUnitID === "admin") ? "admin" : info.role;

        const mockUser: SystemUser = {
          uid: "mock_uid_" + Math.random().toString(36).substr(2, 9),
          email: cleanEmail,
          unitID: finalUnitID,
          role: finalRole as any,
          displayName: finalDisplayName,
          isMock: true
        };
        setUser(mockUser);
        localStorage.setItem("system_auth_user", JSON.stringify(mockUser));
      }
    } catch (error: any) {
      console.error("Lỗi đăng nhập:", error);
      throw error;
    } finally {
      setLoading(false);
    }
  };

  const register = async (
    email: string, 
    password: string, 
    unitID: string, 
    displayName: string, 
    role: "admin" | "user"
  ) => {
    setLoading(true);
    try {
      if (isFirebaseInitialized && auth) {
        // Tạo tài khoản Firebase thực tế
        const userCredential = await createUserWithEmailAndPassword(auth, email, password);
        const fbUser = userCredential.user;
        setUser({
          uid: fbUser.uid,
          email: fbUser.email,
          unitID: unitID.trim().toLowerCase(),
          role: role,
          displayName: displayName,
          isMock: false
        });
      } else {
        // Mô phỏng đăng ký
        const mockUser: SystemUser = {
          uid: "mock_uid_" + Math.random().toString(36).substr(2, 9),
          email: email.toLowerCase().trim(),
          unitID: unitID.trim().toLowerCase(),
          role: role,
          displayName: displayName,
          isMock: true
        };
        setUser(mockUser);
        localStorage.setItem("system_auth_user", JSON.stringify(mockUser));
      }
    } catch (error: any) {
      console.error("Lỗi đăng ký:", error);
      throw error;
    } finally {
      setLoading(false);
    }
  };

  const logout = async () => {
    setLoading(true);
    try {
      if (isFirebaseInitialized && auth) {
        await signOut(auth);
      }
      setUser(null);
      localStorage.removeItem("system_auth_user");
    } catch (error: any) {
      console.error("Lỗi đăng xuất:", error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <AuthContext.Provider
      value={{
        user,
        loading,
        isFirebaseMode: isFirebaseInitialized,
        login,
        register,
        logout
      }}
    >
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error("useAuth phải được bọc trong AuthProvider");
  }
  return context;
};

// Export các đối tượng Firebase để tái sử dụng ở các file khác
export { db, auth, isFirebaseInitialized };
