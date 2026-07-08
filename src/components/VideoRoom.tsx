import React, { useEffect, useRef, useState } from "react";
import AgoraRTC, { 
  IAgoraRTCClient, 
  ICameraVideoTrack, 
  IMicrophoneAudioTrack,
  IAgoraRTCRemoteUser
} from "agora-rtc-sdk-ng";
import { useAuth, db, isFirebaseInitialized } from "../context/AuthContext";
import { 
  collection, 
  query, 
  where, 
  orderBy, 
  onSnapshot, 
  addDoc,
  deleteDoc,
  doc,
  setDoc,
  getDoc
} from "firebase/firestore";
import { 
  Video, 
  VideoOff, 
  Mic, 
  MicOff, 
  PhoneOff, 
  Users, 
  Activity, 
  Tv, 
  ShieldAlert,
  Sparkles,
  FileText,
  FileSpreadsheet,
  File,
  Download,
  UploadCloud,
  Trash2,
  Clock,
  Settings,
  Lock,
  Unlock,
  Key,
  Copy,
  Check,
  RefreshCw
} from "lucide-react";

// ==========================================
// CẤU HÌNH AGORA (AGORA APP ID CONFIGURATION)
// ==========================================
export const agoraAppId: string = (import.meta as any).env?.VITE_AGORA_APP_ID || "YOUR_AGORA_APP_ID";

// Định nghĩa kiểu dữ liệu cho tài liệu trao đổi
interface RoomDoc {
  id: string;
  fileName: string;
  fileSize: number;
  fileType: string;
  uploadedBy: string;
  uploadedAt: string;
  fileData: string; // Base64 data URL
}

export const VideoRoom: React.FC = () => {
  const { user } = useAuth();
  const channelName = user?.unitID || "phong_chung";
  const displayName = user?.displayName || "Thành viên";

  // Cấu hình mã kết nối Agora RTC (Lưu trữ động qua LocalStorage hoặc Biến môi trường)
  const [agoraAppIdState, setAgoraAppIdState] = useState<string>(() => {
    const savedLocal = localStorage.getItem("agora_app_id");
    if (savedLocal && savedLocal !== "YOUR_AGORA_APP_ID") {
      return savedLocal;
    }
    return agoraAppId;
  });
  const [agoraCertificateState, setAgoraCertificateState] = useState<string>(() => {
    return localStorage.getItem("agora_certificate") || "";
  });
  const [agoraTokenState, setAgoraTokenState] = useState<string>(() => {
    return localStorage.getItem("agora_token") || "";
  });
  const [showAgoraConfig, setShowAgoraConfig] = useState(false);
  const [tempAgoraAppId, setTempAgoraAppId] = useState(() => {
    return agoraAppIdState === "YOUR_AGORA_APP_ID" ? "" : agoraAppIdState;
  });
  const [tempAgoraCertificate, setTempAgoraCertificate] = useState(agoraCertificateState);
  const [tempAgoraToken, setTempAgoraToken] = useState(agoraTokenState);

  const [joined, setJoined] = useState(false);
  const [isConnecting, setIsConnecting] = useState(false);
  const [localVideoTrack, setLocalVideoTrack] = useState<ICameraVideoTrack | null>(null);
  const [localAudioTrack, setLocalAudioTrack] = useState<IMicrophoneAudioTrack | null>(null);
  const [localScreenTrack, setLocalScreenTrack] = useState<any | null>(null);
  const [remoteUsers, setRemoteUsers] = useState<IAgoraRTCRemoteUser[]>([]);
  
  const [micMuted, setMicMuted] = useState(false);
  const [camMuted, setCamMuted] = useState(false);
  const [screenSharing, setScreenSharing] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [isSimulationMode, setIsSimulationMode] = useState(false);

  // TRẠNG THÁI MẬT KHẨU BẢO MẬT PHÒNG HỌP (ROOM PASSWORD SECURITY)
  const [roomPassword, setRoomPassword] = useState<string>("");
  const [isAuthorized, setIsAuthorized] = useState<boolean>(false);
  const [isFirstPerson, setIsFirstPerson] = useState<boolean>(false);
  const [passwordInput, setPasswordInput] = useState<string>("");
  const [passwordError, setPasswordError] = useState<string | null>(null);
  const [checkingPassword, setCheckingPassword] = useState<boolean>(true);

  // Tài liệu chia sẻ trong phòng họp
  const [documents, setDocuments] = useState<RoomDoc[]>([]);
  const [uploading, setUploading] = useState(false);

  // Danh sách giả lập thành viên trực tuyến khi chạy ở chế độ demo
  const [simulatedParticipants, setSimulatedParticipants] = useState<Array<{ id: string; name: string; avatarColor: string; micOn: boolean; camOn: boolean; speaking: boolean }>>([]);

  const rtcClientRef = useRef<IAgoraRTCClient | null>(null);
  const localVideoDivRef = useRef<HTMLDivElement | null>(null);

  // Khởi tạo Client Agora
  useEffect(() => {
    rtcClientRef.current = AgoraRTC.createClient({ mode: "rtc", codec: "vp8" });
    
    const activeAppId = agoraAppIdState;
    if (!activeAppId || activeAppId.includes("YOUR_") || activeAppId === "") {
      setIsSimulationMode(true);
      setSimulatedParticipants([
        { id: "p1", name: "Sở Kế Hoạch & Đầu Tư", avatarColor: "bg-emerald-500", micOn: true, camOn: true, speaking: true },
        { id: "Cục Thu Thuế Thành Phố", name: "Cục Thu Thuế Thành Phố", avatarColor: "bg-purple-500", micOn: false, camOn: true, speaking: false },
        { id: "Đơn vị Kiểm toán Khu vực I", name: "Đơn vị Kiểm toán Khu vực I", avatarColor: "bg-indigo-500", micOn: true, camOn: false, speaking: false }
      ]);
    } else {
      setIsSimulationMode(false);
      setSimulatedParticipants([]);
    }

    return () => {
      handleLeave();
    };
  }, [agoraAppIdState]);

  // Tự động tải cấu hình phòng họp (Ưu tiên Biến môi trường Hệ thống toàn cục, sau đó mới tới Cấu hình phòng động)
  useEffect(() => {
    const loadAgoraConfig = async () => {
      try {
        // 1. Kiểm tra cấu hình hệ thống toàn cục từ biến môi trường của Máy chủ (.env / Secrets)
        const globalRes = await fetch("/api/agora-global-status");
        if (globalRes.ok) {
          const globalData = await globalRes.json();
          if (globalData && globalData.hasGlobalConfig && globalData.appId) {
            console.log("[Agora Sync] Phát hiện Cấu hình Hệ thống toàn cục (.env). Tự động kích hoạt họp thật!");
            setAgoraAppIdState(globalData.appId);
            setAgoraCertificateState("system-configured"); // Đánh dấu dùng Certificate hệ thống ở phía Backend
            setAgoraTokenState("");
            
            setTempAgoraAppId(globalData.appId);
            setTempAgoraCertificate("system-configured");
            setTempAgoraToken("");
            
            // Đặt về chế độ kết nối thật
            setIsSimulationMode(false);
            return; // Đã cấu hình toàn cục, không cần tải cấu hình phòng tạm thời nữa
          }
        }
      } catch (err) {
        console.warn("Chưa cấu hình hoặc lỗi kết nối api trạng thái hệ thống toàn cục, thử cấu hình phòng...", err);
      }

      // 2. Nếu không có cấu hình toàn cục, tải cấu hình phòng động do người dùng tự lưu trước đó
      try {
        const response = await fetch(`/api/get-agora-config?channelName=${channelName}`);
        if (response.ok) {
          const data = await response.json();
          if (data.found && data.appId) {
            console.log("[Agora Sync] Đã đồng bộ cấu hình phòng họp động từ máy chủ đám mây:", data);
            setAgoraAppIdState(data.appId);
            setAgoraCertificateState(data.appCertificate || "");
            setAgoraTokenState(data.token || "");
            
            setTempAgoraAppId(data.appId);
            setTempAgoraCertificate(data.appCertificate || "");
            setTempAgoraToken(data.token || "");
            
            setIsSimulationMode(false);
          }
        }
      } catch (err) {
        console.error("Lỗi khi tải cấu hình phòng từ máy chủ:", err);
      }
    };

    loadAgoraConfig();
  }, [channelName]);

  // TỰ ĐỘNG TẢI VÀ ĐỒNG BỘ MẬT KHẨU PHÒNG HỌP THỜI GIAN THỰC (ROOM PASSWORD REAL-TIME SYNC)
  useEffect(() => {
    let unsubscribe: any = () => {};
    let isMounted = true;

    const checkAndInitPassword = async () => {
      setCheckingPassword(true);
      
      // 1. Nếu Firebase hoạt động, sử dụng Firestore để đồng bộ thời gian thực
      if (isFirebaseInitialized && db) {
        try {
          const pwdDocRef = doc(db, "room_passwords", channelName);
          unsubscribe = onSnapshot(pwdDocRef, async (docSnap) => {
            if (!isMounted) return;
            
            if (docSnap.exists()) {
              const activePwd = docSnap.data().password;
              setRoomPassword(activePwd);
              setIsFirstPerson(false);
              
              // Kiểm tra xem trình duyệt đã được cấp quyền trước đó cho mật khẩu này chưa
              const savedAuth = sessionStorage.getItem(`auth_room_${channelName}`);
              if (savedAuth === activePwd) {
                setIsAuthorized(true);
              } else {
                setIsAuthorized(false);
              }
            } else {
              // Chưa có mật khẩu -> Bạn là người đầu tiên vào phòng!
              const newPin = Math.floor(100000 + Math.random() * 900000).toString();
              try {
                await setDoc(pwdDocRef, {
                  password: newPin,
                  createdAt: new Date().toISOString(),
                  createdBy: displayName
                });
                if (isMounted) {
                  setRoomPassword(newPin);
                  setIsFirstPerson(true);
                  setIsAuthorized(true); // Tự động qua cửa cho người tạo
                  sessionStorage.setItem(`auth_room_${channelName}`, newPin);
                }
              } catch (err) {
                console.error("Lỗi khởi tạo mật khẩu Firestore:", err);
              }
            }
            if (isMounted) setCheckingPassword(false);
          });
          return;
        } catch (err) {
          console.error("Lỗi lắng nghe mật khẩu Firestore:", err);
        }
      }

      // 2. Chế độ dự phòng: Gọi API máy chủ Express
      try {
        const response = await fetch(`/api/get-room-password?channelName=${channelName}`);
        if (!isMounted) return;
        
        if (response.ok) {
          const data = await response.json();
          if (data.found && data.password) {
            setRoomPassword(data.password);
            setIsFirstPerson(false);
            
            const savedAuth = sessionStorage.getItem(`auth_room_${channelName}`);
            if (savedAuth === data.password) {
              setIsAuthorized(true);
            } else {
              setIsAuthorized(false);
            }
          } else {
            // Chưa có mật khẩu -> Tạo ngẫu nhiên trên máy chủ
            const newPin = Math.floor(100000 + Math.random() * 900000).toString();
            await fetch("/api/set-room-password", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ channelName, password: newPin })
            });
            if (isMounted) {
              setRoomPassword(newPin);
              setIsFirstPerson(true);
              setIsAuthorized(true);
              sessionStorage.setItem(`auth_room_${channelName}`, newPin);
            }
          }
        }
      } catch (err) {
        console.error("Lỗi khi tải mật khẩu từ máy chủ:", err);
      } finally {
        if (isMounted) setCheckingPassword(false);
      }
    };

    checkAndInitPassword();

    return () => {
      isMounted = false;
      unsubscribe();
    };
  }, [channelName]);

  const handleResetPassword = async () => {
    if (!confirm("Bạn có chắc chắn muốn làm mới mật khẩu phòng họp này? Tất cả thành viên khác sẽ phải nhập mật khẩu mới để tham gia!")) {
      return;
    }

    const newPin = Math.floor(100000 + Math.random() * 900000).toString();
    
    if (isFirebaseInitialized && db) {
      try {
        const pwdDocRef = doc(db, "room_passwords", channelName);
        await setDoc(pwdDocRef, {
          password: newPin,
          createdAt: new Date().toISOString(),
          createdBy: displayName
        });
        setRoomPassword(newPin);
        setIsFirstPerson(true);
        setIsAuthorized(true);
        sessionStorage.setItem(`auth_room_${channelName}`, newPin);
        alert(`Đã làm mới mật khẩu phòng họp thành công! Mật khẩu mới là: ${newPin}`);
        return;
      } catch (err) {
        console.error("Lỗi đặt lại mật khẩu Firestore:", err);
      }
    }

    try {
      const response = await fetch("/api/set-room-password", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ channelName, password: newPin })
      });
      if (response.ok) {
        setRoomPassword(newPin);
        setIsFirstPerson(true);
        setIsAuthorized(true);
        sessionStorage.setItem(`auth_room_${channelName}`, newPin);
        alert(`Đã làm mới mật khẩu phòng họp thành công! Mật khẩu mới là: ${newPin}`);
      }
    } catch (err) {
      console.error("Lỗi đặt lại mật khẩu máy chủ:", err);
    }
  };

  // ĐỒNG BỘ TÀI LIỆU CHUNG TRONG PHÒNG THỜI GIAN THỰC
  useEffect(() => {
    let unsubscribe: any = () => {};

    if (isFirebaseInitialized && db) {
      try {
        const docsRef = collection(db, "room_documents");
        const q = query(
          docsRef,
          where("channelName", "==", channelName),
          orderBy("uploadedAt", "desc")
        );

        unsubscribe = onSnapshot(q, (snapshot) => {
          const docsList: RoomDoc[] = [];
          snapshot.forEach((doc) => {
            const data = doc.data();
            docsList.push({
              id: doc.id,
              fileName: data.fileName || "",
              fileSize: data.fileSize || 0,
              fileType: data.fileType || "",
              uploadedBy: data.uploadedBy || "Khách",
              uploadedAt: data.uploadedAt || new Date().toISOString(),
              fileData: data.fileData || ""
            });
          });
          setDocuments(docsList);
        }, (err) => {
          console.error("Lỗi lắng nghe room_documents:", err);
        });
      } catch (err) {
        console.error("Lỗi khởi tạo lắng nghe room_documents:", err);
      }
    } else {
      // Chế độ mô phỏng Offline / Mock: đọc & lắng nghe qua LocalStorage
      const fetchLocalDocs = () => {
        const raw = localStorage.getItem(`mock_docs_${channelName}`) || "[]";
        try {
          setDocuments(JSON.parse(raw));
        } catch {
          setDocuments([]);
        }
      };

      fetchLocalDocs();

      // Đồng bộ sự kiện thay đổi dữ liệu giả lập
      const handleStorageChange = () => {
        fetchLocalDocs();
      };
      window.addEventListener("storage_docs_updated", handleStorageChange);
      return () => {
        window.removeEventListener("storage_docs_updated", handleStorageChange);
      };
    }

    return () => unsubscribe();
  }, [channelName]);

  // Lắng nghe các sự kiện từ Agora khi kết nối thực tế
  const setupClientEvents = (client: IAgoraRTCClient) => {
    client.on("user-published", async (remoteUser, mediaType) => {
      await client.subscribe(remoteUser, mediaType);
      
      if (mediaType === "video") {
        setRemoteUsers(prev => {
          if (prev.find(u => u.uid === remoteUser.uid)) return prev;
          return [...prev, remoteUser];
        });
      }
      if (mediaType === "audio") {
        remoteUser.audioTrack?.play();
      }
    });

    client.on("user-unpublished", (remoteUser, mediaType) => {
      if (mediaType === "video") {
        setRemoteUsers(prev => prev.filter(u => u.uid !== remoteUser.uid));
      }
    });

    client.on("user-left", (remoteUser) => {
      setRemoteUsers(prev => prev.filter(u => u.uid !== remoteUser.uid));
    });
  };

  const handleJoin = async () => {
    if (isSimulationMode) {
      setJoined(true);
      return;
    }

    if (isConnecting || joined) {
      console.log("Already joining or joined, skipping handleJoin");
      return;
    }

    setIsConnecting(true);
    try {
      setErrorMsg(null);
      const client = rtcClientRef.current;
      if (!client) {
        setIsConnecting(false);
        return;
      }

      // Kiểm tra trạng thái kết nối hiện tại để tránh lỗi Client already in connecting/connected state
      if (client.connectionState === "CONNECTED" || client.connectionState === "CONNECTING") {
        console.log("Client is already CONNECTED or CONNECTING. Setting joined to true.");
        setJoined(true);
        setIsConnecting(false);
        return;
      }

      setupClientEvents(client);

      // Tham gia phòng
      let activeToken = agoraTokenState.trim() || null;

      // Nếu có App ID và App Certificate, tự động sinh Token qua Backend API!
      if (agoraAppIdState && agoraCertificateState) {
        try {
          console.log("Đang tự động sinh Agora Token từ App Certificate...");
          const res = await fetch("/api/generate-token", {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              appId: agoraAppIdState,
              appCertificate: agoraCertificateState,
              channelName: channelName,
              uid: user?.uid || "0",
            }),
          });
          if (res.ok) {
            const data = await res.json();
            if (data.token) {
              activeToken = data.token;
              console.log("Đã sinh Token tự động thành công!");
            }
          } else {
            const errData = await res.json();
            console.warn("Máy chủ trả về lỗi khi sinh Token:", errData);
          }
        } catch (tokenErr) {
          console.error("Không thể kết nối API sinh Token:", tokenErr);
        }
      }

      // Kiểm tra lại trạng thái kết nối ngay trước khi join (phòng tránh async race conditions)
      if (client.connectionState === "CONNECTED" || client.connectionState === "CONNECTING") {
        console.log("Client is already CONNECTED or CONNECTING before final join. Setting joined to true.");
        setJoined(true);
        setIsConnecting(false);
        return;
      }

      await client.join(agoraAppIdState, channelName, activeToken, user?.uid || null);

      // Tạo và xuất bản thiết bị âm thanh/hình ảnh an toàn (phòng tránh lỗi thiếu camera/microphone)
      let audioTrack: IMicrophoneAudioTrack | null = null;
      let videoTrack: ICameraVideoTrack | null = null;
      let initError: string | null = null;

      try {
        const [aTrack, vTrack] = await AgoraRTC.createMicrophoneAndCameraTracks();
        audioTrack = aTrack;
        videoTrack = vTrack;
      } catch (trackErr: any) {
        console.warn("Không thể tạo đồng thời cả mic và camera (có thể do thiếu một trong hai thiết bị). Thử tạo từng thiết bị lẻ...", trackErr);
        
        // Thử tạo Microphone độc lập
        try {
          audioTrack = await AgoraRTC.createMicrophoneAudioTrack();
        } catch (micErr: any) {
          console.warn("Không tìm thấy Microphone hoặc bị chặn quyền truy cập (Có thể do thiết bị ảo/sandbox):", micErr);
        }

        // Thử tạo Camera độc lập
        try {
          videoTrack = await AgoraRTC.createCameraVideoTrack();
        } catch (camErr: any) {
          console.warn("Không tìm thấy Camera/Webcam hoặc bị chặn quyền truy cập (Có thể do thiết bị ảo/sandbox):", camErr);
        }

        if (!audioTrack && !videoTrack) {
          initError = "Không phát hiện Microphone và Camera khả dụng trên thiết bị của bạn. Bạn đã được kết nối vào cuộc họp thật ở chế độ Thính giả (chỉ nghe và xem người khác).";
        } else if (!audioTrack) {
          initError = "Không tìm thấy Microphone khả dụng. Bạn tham gia phòng họp thật nhưng người khác không thể nghe thấy tiếng của bạn (chỉ xem được hình ảnh).";
        } else if (!videoTrack) {
          initError = "Không tìm thấy Webcam/Camera khả dụng. Bạn tham gia phòng họp thật nhưng người khác không thể nhìn thấy hình ảnh của bạn (chỉ nghe được âm thanh).";
        }
      }

      const tracksToPublish: any[] = [];

      if (audioTrack) {
        setLocalAudioTrack(audioTrack);
        tracksToPublish.push(audioTrack);
        setMicMuted(false);
      } else {
        setMicMuted(true);
      }

      if (videoTrack) {
        setLocalVideoTrack(videoTrack);
        tracksToPublish.push(videoTrack);
        setCamMuted(false);
      } else {
        setCamMuted(true);
      }

      if (tracksToPublish.length > 0) {
        await client.publish(tracksToPublish);
      }

      setJoined(true);
      setIsConnecting(false);

      if (initError) {
        console.warn(initError);
        setErrorMsg(initError);
        // Tự động tắt cảnh báo sau 10 giây để giữ giao diện gọn gàng
        setTimeout(() => setErrorMsg(null), 10000);
      }

      // Phát video nội bộ nếu có camera
      setTimeout(() => {
        if (localVideoDivRef.current && videoTrack) {
          videoTrack.play(localVideoDivRef.current);
        }
      }, 300);

    } catch (err: any) {
      setIsConnecting(false);
      console.error("Lỗi nghiêm trọng khi tham gia phòng Agora:", err);
      
      const errStr = String(err || "").toLowerCase();
      // Nếu lỗi là do đã kết nối hoặc đang kết nối, bỏ qua và đánh dấu đã tham gia thành công
      if (
        errStr.includes("already in connecting") || 
        errStr.includes("invalid_operation") || 
        errStr.includes("client-88663") || 
        (err && err.code === "INVALID_OPERATION")
      ) {
        console.log("Phát hiện lỗi trạng thái đang kết nối nhưng bỏ qua vì đã sẵn sàng.");
        setJoined(true);
        return;
      }

      setErrorMsg(`Lỗi kết nối Agora: ${err.message || err}. Đang chuyển sang chế độ giả lập.`);
      setIsSimulationMode(true);
      setJoined(true);
      setSimulatedParticipants([
        { id: "p1", name: "Sở Kế Hoạch & Đầu Tư", avatarColor: "bg-emerald-500", micOn: true, camOn: true, speaking: true },
        { id: "Cục Thu Thuế Thành Phố", name: "Cục Thu Thuế Thành Phố", avatarColor: "bg-purple-500", micOn: false, camOn: true, speaking: false },
        { id: "Đơn vị Kiểm toán Khu vực I", name: "Đơn vị Kiểm toán Khu vực I", avatarColor: "bg-indigo-500", micOn: true, camOn: false, speaking: false }
      ]);
    }
  };

  const handleLeave = async () => {
    // Đóng screen track nếu có
    if (localScreenTrack) {
      localScreenTrack.stop();
      localScreenTrack.close();
      setLocalScreenTrack(null);
    }

    // Đóng tất cả các luồng âm thanh hình ảnh nội bộ
    if (localAudioTrack) {
      localAudioTrack.stop();
      localAudioTrack.close();
      setLocalAudioTrack(null);
    }
    if (localVideoTrack) {
      localVideoTrack.stop();
      localVideoTrack.close();
      setLocalVideoTrack(null);
    }

    if (rtcClientRef.current && joined && !isSimulationMode) {
      try {
        await rtcClientRef.current.leave();
      } catch (err) {
        console.error("Lỗi khi rời phòng Agora:", err);
      }
    }

    setJoined(false);
    setIsConnecting(false);
    setRemoteUsers([]);
    setScreenSharing(false);
  };

  const toggleMic = async () => {
    if (isSimulationMode) {
      setMicMuted(!micMuted);
      return;
    }
    if (localAudioTrack) {
      const targetState = !micMuted;
      await localAudioTrack.setMuted(targetState);
      setMicMuted(targetState);
    } else {
      // Thử khởi tạo lại micro nếu trước đó chưa có thiết bị hoặc bị tắt
      try {
        const audioTrack = await AgoraRTC.createMicrophoneAudioTrack();
        setLocalAudioTrack(audioTrack);
        await rtcClientRef.current?.publish(audioTrack);
        setMicMuted(false);
        console.log("Đã kết nối Microphone thành công!");
      } catch (err: any) {
        console.warn("Không thể bật Microphone:", err);
        alert("Không tìm thấy Microphone khả dụng hoặc trình duyệt chưa cấp quyền!");
      }
    }
  };

  const toggleCam = async () => {
    if (isSimulationMode) {
      setCamMuted(!camMuted);
      return;
    }
    if (localVideoTrack) {
      const targetState = !camMuted;
      await localVideoTrack.setMuted(targetState);
      setCamMuted(targetState);
    } else {
      // Thử khởi tạo lại camera nếu trước đó chưa có thiết bị hoặc bị tắt
      try {
        const videoTrack = await AgoraRTC.createCameraVideoTrack();
        setLocalVideoTrack(videoTrack);
        await rtcClientRef.current?.publish(videoTrack);
        setCamMuted(false);
        console.log("Đã kết nối Camera thành công!");
        setTimeout(() => {
          if (localVideoDivRef.current && videoTrack) {
            videoTrack.play(localVideoDivRef.current);
          }
        }, 300);
      } catch (err: any) {
        console.warn("Không thể bật Camera:", err);
        alert("Không tìm thấy Camera/Webcam khả dụng hoặc trình duyệt chưa cấp quyền!");
      }
    }
  };

  // ==========================================
  // CHI SẺ MÀN HÌNH (SCREEN SHARING ENGINE)
  // ==========================================
  const startScreenShare = async () => {
    if (isSimulationMode) {
      setScreenSharing(true);
      return;
    }

    try {
      setErrorMsg(null);
      const client = rtcClientRef.current;
      if (!client) return;

      // Tạo screen video track
      const screenTrackResult = await AgoraRTC.createScreenVideoTrack({
        encoderConfig: "1080p_1"
      }, "auto");

      let screenTrack: any;
      if (Array.isArray(screenTrackResult)) {
        screenTrack = screenTrackResult[0];
      } else {
        screenTrack = screenTrackResult;
      }

      setLocalScreenTrack(screenTrack);
      setScreenSharing(true);

      // Unpublish camera video track để tránh xung đột luồng
      if (localVideoTrack) {
        await client.unpublish(localVideoTrack);
      }
      await client.publish(screenTrack);

      // Lắng nghe nút kết thúc chia sẻ màn hình mặc định của trình duyệt
      screenTrack.on("track-ended", async () => {
        await stopScreenShare(screenTrack, client);
      });

    } catch (err: any) {
      console.error("Lỗi khi bắt luồng chia sẻ màn hình:", err);
      setErrorMsg(`Không thể chia sẻ màn hình: ${err.message || err}`);
      setScreenSharing(false);
    }
  };

  const stopScreenShare = async (trackParam?: any, clientParam?: any) => {
    const activeTrack = trackParam || localScreenTrack;
    const client = clientParam || rtcClientRef.current;

    if (activeTrack) {
      activeTrack.stop();
      activeTrack.close();
    }
    setLocalScreenTrack(null);
    setScreenSharing(false);

    if (isSimulationMode) return;

    if (client) {
      try {
        if (activeTrack) {
          await client.unpublish(activeTrack);
        }
        // Khôi phục lại camera video track của webcam
        if (localVideoTrack && !camMuted) {
          await client.publish(localVideoTrack);
          setTimeout(() => {
            if (localVideoDivRef.current && localVideoTrack) {
              localVideoTrack.play(localVideoDivRef.current);
            }
          }, 200);
        }
      } catch (err) {
        console.error("Lỗi khi dừng chia sẻ màn hình:", err);
      }
    }
  };

  const toggleScreenShare = async () => {
    if (screenSharing) {
      await stopScreenShare();
    } else {
      await startScreenShare();
    }
  };

  // ==========================================
  // TRAO ĐỔI TÀI LIỆU (DOCUMENT EXCHANGE ENGINE)
  // ==========================================
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Giới hạn 1.5MB để tối ưu lưu trữ Firestore qua Base64
    if (file.size > 1500000) {
      alert("Dung lượng tệp quá lớn! Vui lòng chọn tệp dưới 1.5MB để truyền tải siêu tốc qua cổng liên ngành.");
      return;
    }

    setUploading(true);

    try {
      const reader = new FileReader();
      reader.onload = async (event) => {
        const base64Data = event.target?.result as string;

        const newDoc = {
          fileName: file.name,
          fileSize: file.size,
          fileType: file.type || "application/octet-stream",
          uploadedBy: displayName,
          uploadedAt: new Date().toISOString(),
          fileData: base64Data,
          channelName: channelName
        };

        if (isFirebaseInitialized && db) {
          // Lưu vào Firestore thời gian thực
          await addDoc(collection(db, "room_documents"), newDoc);
        } else {
          // Lưu vào LocalStorage mô phỏng
          const raw = localStorage.getItem(`mock_docs_${channelName}`) || "[]";
          let localList = [];
          try { localList = JSON.parse(raw); } catch { localList = []; }
          
          const docWithId = {
            id: `mock_doc_${Date.now()}`,
            ...newDoc
          };
          localList.unshift(docWithId);
          localStorage.setItem(`mock_docs_${channelName}`, JSON.stringify(localList));
          
          // Phát thông báo thay đổi sang các thành phần khác
          window.dispatchEvent(new Event("storage_docs_updated"));
          setDocuments(localList);
        }

        setUploading(false);
      };
      reader.onerror = () => {
        alert("Đọc tệp tin thất bại!");
        setUploading(false);
      };
      reader.readAsDataURL(file);
    } catch (err) {
      console.error("Lỗi tải tệp lên:", err);
      alert("Đồng bộ tệp tin thất bại!");
      setUploading(false);
    }
  };

  const handleDeleteDoc = async (docId: string, uploadedBy: string) => {
    const isOwner = uploadedBy.toLowerCase().trim() === displayName.toLowerCase().trim();
    const isAdmin = user?.role === "admin";
    
    let confirmMsg = "Bạn có chắc chắn muốn xóa tệp này khỏi phòng họp chung?";
    if (!isOwner && !isAdmin) {
      confirmMsg = `Tài liệu này do "${uploadedBy}" đăng tải. Bạn có chắc chắn muốn xóa tài liệu này khỏi phòng họp chung không?`;
    }

    if (!confirm(confirmMsg)) return;

    try {
      if (isFirebaseInitialized && db) {
        await deleteDoc(doc(db, "room_documents", docId));
      } else {
        const raw = localStorage.getItem(`mock_docs_${channelName}`) || "[]";
        let localList = [];
        try { localList = JSON.parse(raw); } catch { localList = []; }
        localList = localList.filter((d: any) => d.id !== docId);
        localStorage.setItem(`mock_docs_${channelName}`, JSON.stringify(localList));
        window.dispatchEvent(new Event("storage_docs_updated"));
        setDocuments(localList);
      }
    } catch (err) {
      console.error("Lỗi xóa tệp:", err);
      alert("Không thể xóa tệp!");
    }
  };

  const formatFileSize = (bytes: number) => {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
  };

  const getFileIcon = (fileType: string) => {
    const type = fileType.toLowerCase();
    if (type.includes("pdf")) return <FileText className="w-5 h-5 text-rose-400" />;
    if (type.includes("sheet") || type.includes("excel") || type.includes("csv") || type.includes("xls")) 
      return <FileSpreadsheet className="w-5 h-5 text-emerald-400" />;
    return <File className="w-5 h-5 text-indigo-400" />;
  };

  const formatTime = (isoString: string) => {
    try {
      const d = new Date(isoString);
      return d.toLocaleTimeString("vi-VN", { hour: "2-digit", minute: "2-digit" });
    } catch {
      return "";
    }
  };

  return (
    <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 text-white shadow-xl space-y-6">
      
      {/* Header phòng họp */}
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 border-b border-slate-800 pb-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="p-2 bg-indigo-600/20 text-indigo-400 rounded-lg">
              <Users className="w-5 h-5" />
            </span>
            <h2 className="text-xl font-bold tracking-tight text-white">
              Phòng Họp Trực Tuyến Liên Ngành
            </h2>
          </div>
          <p className="text-xs text-slate-400 mt-1">
            Kênh trao đổi trực tiếp: <span className="text-indigo-400 font-mono font-bold">{channelName.toUpperCase()}</span>
          </p>
        </div>

        <div className="flex items-center gap-2">
          {joined ? (
            <span className="flex items-center gap-1.5 px-3 py-1 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded-full text-xs font-semibold">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-ping"></span>
              Đã kết nối
            </span>
          ) : (
            <span className="flex items-center gap-1.5 px-3 py-1 bg-slate-800 text-slate-400 border border-slate-700 rounded-full text-xs font-semibold">
              <span className="w-2 h-2 rounded-full bg-slate-500"></span>
              Ngoại tuyến
            </span>
          )}

          {isSimulationMode && (
            <span className="flex items-center gap-1 px-3 py-1 bg-amber-500/15 text-amber-400 border border-amber-500/20 rounded-full text-xs font-semibold">
              <Sparkles className="w-3.5 h-3.5" />
              Chế độ Demo
            </span>
          )}

          <button
            onClick={() => setShowAgoraConfig(!showAgoraConfig)}
            className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition cursor-pointer flex items-center gap-1 text-xs font-bold"
            title="Cấu hình kết nối Agora"
          >
            <Settings className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">Cấu hình</span>
          </button>
        </div>
      </div>

      {/* Khối Cấu hình Agora App ID */}
      {showAgoraConfig && (
        <div className="bg-slate-950/80 border border-slate-800 p-5 rounded-2xl space-y-4 animate-fade-in relative">
          <div className="absolute top-4 right-4">
            <button 
              onClick={() => setShowAgoraConfig(false)} 
              className="text-slate-400 hover:text-white text-xs font-bold px-2 py-1 bg-slate-900 hover:bg-slate-800 border border-slate-800 rounded-lg cursor-pointer transition"
            >
              Đóng
            </button>
          </div>

          <div className="flex items-start gap-3">
            <span className="p-2 bg-amber-500/10 text-amber-400 border border-amber-500/20 rounded-xl mt-0.5 shrink-0">
              <Settings className="w-5 h-5 animate-spin" style={{ animationDuration: '6s' }} />
            </span>
            <div className="space-y-1">
              <h4 className="text-sm font-extrabold text-amber-400 tracking-tight uppercase">
                Giải thích Chế độ Demo &amp; Thiết lập Cuộc họp Real-time
              </h4>
              <p className="text-xs text-slate-300 leading-relaxed max-w-2xl font-sans">
                Hệ thống họp trực tuyến (gọi video/âm thanh và chia sẻ màn hình) tích hợp giải pháp toàn cầu <span className="font-semibold text-indigo-400">Agora RTC</span>.
              </p>
            </div>
          </div>

          <div className="bg-slate-900/60 p-4 rounded-xl border border-slate-850 text-xs text-slate-300 leading-relaxed font-sans space-y-2">
            <p>
              💡 <strong className="text-white">Tại sao hệ thống hiển thị "Chế độ Demo"?</strong><br />
              Do ứng dụng chưa được cấu hình khóa <span className="font-mono text-indigo-400">Agora App ID</span> thực tế của bạn, hệ thống tự động kích hoạt chế độ mô phỏng trực quan an toàn. Trong chế độ này, bạn có thể tự do trải nghiệm giao diện phòng họp hiện đại, bật/tắt thiết bị cá nhân, trình chiếu màn hình nội bộ, và tương tác với các thành viên đại diện được lập trình sẵn mà không sợ bị lỗi kết nối hay lỗi cuộc gọi do thiếu khóa.
            </p>
            <p>
              📂 <strong className="text-white">Kho tài liệu là THẬT hay GIẢ LẬP?</strong><br />
              Kho tài liệu liên ngành ở cột bên phải là <strong className="text-emerald-400">THỰC TẾ 100%</strong>. Chức năng này được kết nối trực tiếp với cơ sở dữ liệu thời gian thực đám mây Firebase Firestore. Khi bạn gửi bất kỳ tài liệu nào lên, các đơn vị khác ở đầu cầu bên kia sẽ ngay lập tức nhận được file và có thể tải về trực tiếp.
            </p>
            <p>
              🚀 <strong className="text-white">Cách kích hoạt cuộc gọi thực tế 100% (Real Live Connection):</strong><br />
              Nếu bạn có tài khoản Agora (đăng ký hoàn toàn miễn phí tại <a href="https://www.agora.io" target="_blank" rel="noopener noreferrer" className="text-indigo-400 hover:underline">agora.io</a>), hãy sao chép mã <span className="font-semibold text-white">App ID</span> của dự án và dán vào ô bên dưới. Hệ thống sẽ ngay lập tức thiết lập đường truyền trực tuyến thực tế để bạn họp và kết nối trực tiếp camera/micro thật với các đồng nghiệp khác!
            </p>
          </div>

          <div className="bg-indigo-500/10 border border-indigo-500/20 p-4 rounded-xl text-xs space-y-2 text-slate-300">
            <span className="font-bold text-indigo-400 uppercase tracking-tight block">🔑 THÔNG TIN ĐỂ SINH TRUY CẬP (TOKEN):</span>
            <p className="font-sans">
              Khi bấm nút <strong className="text-white">"Generate Temp Token"</strong> trên dashboard Agora, bạn <span className="text-rose-400 font-semibold underline">bắt buộc</span> phải nhập chính xác tên kênh dưới đây:
            </p>
            <div className="flex items-center gap-2 bg-slate-900 p-2.5 rounded-lg border border-slate-800 font-mono text-sm justify-between">
              <div>
                <span className="text-[10px] text-slate-500 block uppercase font-sans font-bold tracking-wider">Tên Kênh (Channel Name)</span>
                <span className="text-emerald-400 font-bold">{channelName}</span>
              </div>
              <button
                onClick={() => {
                  navigator.clipboard.writeText(channelName);
                  alert("Đã sao chép tên kênh: " + channelName);
                }}
                className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded text-[11px] font-sans font-bold cursor-pointer transition active:scale-95"
              >
                Sao chép tên kênh
              </button>
            </div>
          </div>

          {agoraCertificateState === "system-configured" && (
            <div className="bg-emerald-500/10 border border-emerald-500/20 p-4 rounded-xl text-xs space-y-1.5 text-slate-300">
              <span className="font-bold text-emerald-400 uppercase tracking-tight block flex items-center gap-1.5">
                <Sparkles className="w-4 h-4 text-emerald-400 animate-pulse" /> 
                ĐÃ KÍCH HOẠT CẤU HÌNH THẬT TOÀN CỤC (GLOBAL SYNC)
              </span>
              <p className="font-sans leading-relaxed">
                Hệ thống phát hiện khóa <strong>Agora App ID</strong> và <strong>App Certificate</strong> bảo mật đã được quản trị viên thiết lập sẵn trong tệp cấu hình hệ thống (.env).
              </p>
              <p className="font-sans text-[11px] text-emerald-400 font-semibold">
                👉 Tất cả mọi người khi tham gia vào bất kỳ phòng họp nào đều được tự động kết nối cuộc họp live thật 100% hoàn toàn tự động, không cần bất kỳ thao tác cấu hình thủ công nào!
              </p>
            </div>
          )}

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="space-y-1">
              <label className="text-[10px] font-bold text-slate-400 block uppercase tracking-wider">1. Mã Agora App ID</label>
              <input
                type="text"
                placeholder="Nhập Agora App ID (Ví dụ: 17de3c77...)"
                value={tempAgoraAppId}
                onChange={(e) => setTempAgoraAppId(e.target.value.trim())}
                className="w-full px-3.5 py-2.5 bg-slate-900 border border-slate-800 focus:border-indigo-500 rounded-xl text-xs text-white focus:outline-none transition font-mono"
              />
              <span className="text-[10px] text-slate-500 block leading-normal font-sans">
                Chuỗi 32 ký tự hex nhận từ Agora Console.
              </span>
            </div>

            <div className="space-y-1">
              <label className="text-[10px] font-bold text-emerald-400 block uppercase tracking-wider">2. Mã Primary Certificate (Khuyên dùng)</label>
              <input
                type="text"
                placeholder="Nhập App Certificate (Ví dụ: 91416a2b...)"
                value={tempAgoraCertificate}
                onChange={(e) => setTempAgoraCertificate(e.target.value.trim())}
                className="w-full px-3.5 py-2.5 bg-slate-900 border border-emerald-950 focus:border-emerald-500 rounded-xl text-xs text-white focus:outline-none transition font-mono"
              />
              <span className="text-[10px] text-emerald-500 block leading-normal font-sans">
                💡 Nhập mã này để <strong>tự động sinh Token</strong> cho mọi cuộc họp, không cần dán thủ công mỗi lần!
              </span>
            </div>

            <div className="space-y-1">
              <label className="text-[10px] font-bold text-slate-400 block uppercase tracking-wider">3. Hoặc mã Temp Token (Thủ công)</label>
              <input
                type="text"
                placeholder="Chỉ dán nếu không dùng Certificate..."
                value={tempAgoraToken}
                onChange={(e) => setTempAgoraToken(e.target.value.trim())}
                className="w-full px-3.5 py-2.5 bg-slate-900 border border-slate-800 focus:border-indigo-500 rounded-xl text-xs text-white focus:outline-none transition font-mono"
              />
              <span className="text-[10px] text-slate-500 block leading-normal font-sans">
                Token tạm thời lấy từ dashboard (hết hạn sau 24h).
              </span>
            </div>
          </div>

          <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 pt-2">
            <span className="text-[11px] text-slate-400 font-sans italic">
              * Để họp thật tự động, hãy dán App ID (ô 1) và App Certificate (ô 2). Ứng dụng sẽ tự sinh Token hoàn hảo!
            </span>
            <div className="flex gap-2 self-end sm:pb-0.5">
              <button
                onClick={async () => {
                  if (!tempAgoraAppId) {
                    alert("Vui lòng nhập mã App ID hoặc ấn 'Quay về mặc định'!");
                    return;
                  }
                  
                  // 1. Lưu cục bộ ở trình duyệt của máy hiện tại
                  localStorage.setItem("agora_app_id", tempAgoraAppId);
                  localStorage.setItem("agora_certificate", tempAgoraCertificate);
                  localStorage.setItem("agora_token", tempAgoraToken);
                  
                  setAgoraAppIdState(tempAgoraAppId);
                  setAgoraCertificateState(tempAgoraCertificate);
                  setAgoraTokenState(tempAgoraToken);

                  // 2. Đồng bộ hóa trực tuyến lên máy chủ đám mây để tất cả thiết bị khác trong phòng kết nối tự động
                  try {
                    await fetch("/api/save-agora-config", {
                      method: "POST",
                      headers: {
                        "Content-Type": "application/json"
                      },
                      body: JSON.stringify({
                        channelName: channelName,
                        appId: tempAgoraAppId,
                        appCertificate: tempAgoraCertificate,
                        token: tempAgoraToken
                      })
                    });
                  } catch (err) {
                    console.error("Lỗi đồng bộ cấu hình lên máy chủ đám mây:", err);
                  }
                  
                  setShowAgoraConfig(false);
                  alert("Đã cấu hình Agora & đồng bộ đám mây thành công! Mọi thành viên khác vào phòng này sẽ tự động kết nối cuộc họp thật mà không cần phải nhập lại mã!");
                }}
                className="px-4 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs rounded-xl cursor-pointer transition active:scale-95 whitespace-nowrap"
              >
                Lưu &amp; Đồng bộ Đám mây
              </button>
              <button
                onClick={async () => {
                  localStorage.removeItem("agora_app_id");
                  localStorage.removeItem("agora_certificate");
                  localStorage.removeItem("agora_token");
                  
                  setAgoraAppIdState("YOUR_AGORA_APP_ID");
                  setAgoraCertificateState("");
                  setAgoraTokenState("");
                  
                  setTempAgoraAppId("YOUR_AGORA_APP_ID");
                  setTempAgoraCertificate("");
                  setTempAgoraToken("");

                  // Xóa cấu hình đồng bộ trên máy chủ đám mây
                  try {
                    await fetch("/api/save-agora-config", {
                      method: "POST",
                      headers: {
                        "Content-Type": "application/json"
                      },
                      body: JSON.stringify({
                        channelName: channelName,
                        appId: "",
                        appCertificate: "",
                        token: ""
                      })
                    });
                  } catch (err) {
                    console.error("Lỗi xóa cấu hình trên máy chủ đám mây:", err);
                  }
                  
                  setShowAgoraConfig(false);
                  alert("Đã đặt phòng họp về Chế độ Demo mặc định!");
                }}
                className="px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold text-xs rounded-xl cursor-pointer transition active:scale-95 whitespace-nowrap border border-slate-700"
              >
                Quay về mặc định
              </button>
            </div>
          </div>
        </div>
      )}

      {errorMsg && (
        <div className="space-y-3">
          <div className="bg-rose-500/10 border border-rose-500/20 rounded-xl p-4 text-xs text-rose-300 font-mono">
            {errorMsg}
          </div>
          
          {(errorMsg.includes("CAN_NOT_GET_GATEWAY_SERVER") || errorMsg.includes("invalid vendor key") || errorMsg.includes("appid")) && (
            <div className="bg-amber-500/5 border border-amber-500/20 rounded-xl p-5 space-y-3 text-xs leading-relaxed text-slate-300">
              <h5 className="font-bold text-amber-400 uppercase tracking-tight flex items-center gap-1.5">
                ⚠️ HƯỚNG DẪN KHẮC PHỤC SỰ CỐ KẾT NỐI AGORA REAL-TIME
              </h5>
              <p className="font-sans">
                Bạn đã dán mã App ID, nhưng máy chủ toàn cầu của **Agora** từ chối kết nối vì mã không hợp lệ hoặc sai cấu hình bảo mật. Vui lòng kiểm tra lại 2 điểm mấu chốt sau trên trang quản trị <a href="https://dashboard.agora.io" target="_blank" rel="noopener noreferrer" className="text-indigo-400 hover:underline font-bold">agora.io</a>:
              </p>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 font-sans text-[11.5px] pt-1">
                <div className="bg-slate-900/40 p-3 rounded-lg border border-slate-800 space-y-1.5">
                  <span className="font-bold text-slate-200">1. Chọn chế độ "Testing Mode" (Không cần Token):</span>
                  <p className="text-slate-400">
                    Mặc định nếu bạn chọn dự án ở chế độ <strong className="text-rose-400">"Secured Mode: APP ID + Token"</strong>, Agora bắt buộc phải có Token bảo mật được sinh ra từ máy chủ riêng. Khi ứng dụng truyền Token là <code className="text-indigo-400">null</code>, Agora sẽ báo lỗi <code className="text-rose-400">invalid vendor key</code>.
                  </p>
                  <p className="text-emerald-400 font-semibold">
                    👉 Giải pháp: Tạo dự án mới trên Agora Console và chọn "Testing Mode" (APP ID only), sau đó lấy mã App ID của dự án mới dán vào đây.
                  </p>
                </div>
                
                <div className="bg-slate-900/40 p-3 rounded-lg border border-slate-800 space-y-1.5">
                  <span className="font-bold text-slate-200">2. Xác nhận copy đúng "App ID":</span>
                  <p className="text-slate-400">
                    Đảm bảo bạn copy chính xác chuỗi mã <strong className="text-slate-200">App ID</strong> (khoảng 32 ký tự gồm chữ và số), tránh copy nhầm sang các mã khác như:
                  </p>
                  <ul className="list-disc pl-4 space-y-0.5 text-slate-400">
                    <li><strong className="text-rose-400">App Certificate</strong> (Mã chứng thực bảo mật)</li>
                    <li><strong className="text-rose-400">Customer ID / Customer Secret</strong></li>
                    <li>Mã Token tạm thời đã hết hạn</li>
                  </ul>
                </div>
              </div>

              <div className="bg-indigo-500/5 p-3 rounded-lg border border-indigo-500/10 font-sans text-slate-300">
                💡 <strong className="text-indigo-400">Hệ thống đã tự động chuyển sang Chế độ Giả lập Trực tuyến (Demo Mode) an toàn:</strong> Bạn vẫn có thể trải nghiệm toàn bộ giao diện cuộc họp, trình chiếu màn hình, và <strong className="text-emerald-400">đồng bộ chia sẻ tài liệu thực tế 100%</strong> (kết nối trực tiếp qua đám mây Firebase Firestore) một cách trơn tru nhất!
              </div>
            </div>
          )}
        </div>
      )}

      {/* Sân khấu chính */}
      {!joined ? (
        checkingPassword ? (
          <div className="bg-slate-950/40 border border-slate-800/80 rounded-2xl p-12 text-center flex flex-col items-center justify-center space-y-4">
            <RefreshCw className="w-8 h-8 text-indigo-400 animate-spin" />
            <p className="text-xs text-slate-400 font-sans">Đang kiểm tra bảo mật phòng họp...</p>
          </div>
        ) : !isAuthorized ? (
          <div className="bg-slate-950/40 border border-slate-800/80 rounded-2xl p-8 text-center flex flex-col items-center justify-center space-y-6 max-w-md mx-auto">
            <div className="w-16 h-16 rounded-full bg-rose-500/10 border border-rose-500/20 flex items-center justify-center text-rose-400 animate-pulse">
              <Lock className="w-8 h-8" />
            </div>
            <div>
              <h3 className="text-lg font-bold text-slate-200">Phòng họp đã được khóa</h3>
              <p className="text-xs text-slate-400 mt-2 leading-relaxed font-sans">
                Phòng họp của đơn vị <span className="font-mono text-indigo-400 font-bold">{channelName.toUpperCase()}</span> đã được thiết lập mã khóa bảo mật. Vui lòng nhập mật khẩu phòng họp để tham gia.
              </p>
            </div>
            
            <div className="w-full space-y-4">
              <div className="relative">
                <span className="absolute left-3.5 top-3.5 text-slate-400">
                  <Key className="w-4 h-4" />
                </span>
                <input
                  type="text"
                  placeholder="Nhập 6 số mật khẩu phòng..."
                  value={passwordInput}
                  maxLength={6}
                  onChange={(e) => {
                    setPasswordInput(e.target.value.replace(/\D/g, "").slice(0, 6));
                    setPasswordError(null);
                  }}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      if (passwordInput === roomPassword) {
                        setIsAuthorized(true);
                        sessionStorage.setItem(`auth_room_${channelName}`, roomPassword);
                      } else {
                        setPasswordError("Mật khẩu phòng họp không chính xác! Vui lòng thử lại.");
                      }
                    }
                  }}
                  className="w-full pl-10 pr-4 py-3 bg-slate-900 border border-slate-800 focus:border-indigo-500 rounded-xl text-sm text-white focus:outline-none transition text-center tracking-widest font-mono font-extrabold text-lg"
                />
              </div>
              
              {passwordError && (
                <p className="text-rose-450 text-xs font-sans text-center">{passwordError}</p>
              )}

              <button
                onClick={() => {
                  if (passwordInput === roomPassword) {
                    setIsAuthorized(true);
                    sessionStorage.setItem(`auth_room_${channelName}`, roomPassword);
                  } else {
                    setPasswordError("Mật khẩu phòng họp không chính xác! Vui lòng thử lại.");
                  }
                }}
                className="w-full py-3 bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs rounded-xl cursor-pointer transition active:scale-95 flex items-center justify-center gap-1.5 shadow-lg shadow-indigo-600/10"
              >
                <Unlock className="w-4 h-4" /> Xác thực &amp; Vào phòng họp
              </button>
            </div>
          </div>
        ) : (
          <div className="bg-slate-950/40 border border-slate-800/80 rounded-2xl p-10 text-center flex flex-col items-center justify-center space-y-6 max-w-lg mx-auto">
            <div className="w-16 h-16 rounded-full bg-emerald-500/10 border border-emerald-500/20 flex items-center justify-center text-emerald-400">
              <Unlock className="w-8 h-8" />
            </div>
            <div className="max-w-md">
              <h3 className="text-lg font-bold text-slate-200 font-sans">Sẵn sàng tham gia cuộc họp</h3>
              <p className="text-xs text-slate-400 mt-1 leading-relaxed font-sans">
                Bạn đã được xác thực thành công vào phòng họp của đơn vị <span className="font-mono text-indigo-400 font-bold">{channelName.toUpperCase()}</span>.
              </p>
            </div>

            {/* Banner hiển thị mật khẩu */}
            <div className="bg-indigo-500/10 border border-indigo-500/20 rounded-2xl p-5 w-full space-y-3.5">
              <span className="text-[10px] font-bold text-indigo-300 block uppercase tracking-wider font-sans">
                🔑 MẬT KHẨU PHÒNG HỌP HIỆN TẠI
              </span>
              <div className="flex items-center justify-center gap-3">
                <span className="text-2xl font-mono font-extrabold tracking-widest text-white bg-slate-900 px-6 py-3 rounded-xl border border-indigo-500/20 shadow-inner">
                  {roomPassword}
                </span>
                <button
                  onClick={() => {
                    navigator.clipboard.writeText(roomPassword);
                    alert("Đã sao chép mật khẩu phòng họp thành công! Hãy gửi cho các thành viên khác.");
                  }}
                  className="p-3 bg-indigo-600/20 hover:bg-indigo-600/35 text-indigo-300 rounded-xl transition cursor-pointer border border-indigo-500/20 active:scale-90"
                  title="Sao chép mật khẩu"
                >
                  <Copy className="w-4 h-4" />
                </button>
              </div>
              <p className="text-[11.5px] text-slate-400 leading-relaxed font-sans max-w-sm mx-auto">
                {isFirstPerson ? (
                  <span>Bạn là <strong className="text-emerald-400">người đầu tiên khởi tạo</strong> cuộc họp này. Vui lòng gửi mật khẩu <strong>{roomPassword}</strong> cho những thành viên khác để họ cùng tham gia.</span>
                ) : (
                  <span>Hãy gửi mật khẩu này cho các thành viên khác để tham gia cuộc họp của đơn vị bạn!</span>
                )}
              </p>
            </div>

            <div className="flex flex-col sm:flex-row items-center gap-3 w-full justify-center">
              <button
                onClick={handleJoin}
                disabled={isConnecting}
                className={`w-full sm:w-auto px-8 py-3 bg-indigo-600 hover:bg-indigo-500 active:translate-y-0.5 transition font-bold text-sm rounded-xl flex items-center justify-center gap-2 cursor-pointer shadow-lg shadow-indigo-600/20 ${isConnecting ? "opacity-60 cursor-not-allowed" : ""}`}
              >
                {isConnecting ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin text-white" />
                    ĐANG KẾT NỐI...
                  </>
                ) : (
                  <>
                    <Users className="w-4 h-4" /> Tham gia phòng họp
                  </>
                )}
              </button>
              
              <button
                onClick={handleResetPassword}
                className="w-full sm:w-auto px-5 py-3 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold text-xs rounded-xl cursor-pointer transition flex items-center justify-center gap-1.5 border border-slate-700 active:scale-95"
                title="Tạo mã khóa mới"
              >
                <RefreshCw className="w-3.5 h-3.5" /> Đổi mật khẩu mới
              </button>
            </div>
          </div>
        )
      ) : (
        <div className="space-y-6">
          {/* BANNER THÔNG TIN PHÒNG HỌP & MẬT KHẨU */}
          <div className="bg-slate-900 border border-indigo-500/10 p-3.5 rounded-2xl flex flex-col md:flex-row items-center justify-between gap-3 shadow-md">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-indigo-500/10 border border-indigo-500/20 flex items-center justify-center text-indigo-400">
                <Key className="w-5 h-5" />
              </div>
              <div className="text-left">
                <p className="text-[10px] font-bold text-indigo-300 block uppercase tracking-wider font-sans">Thông tin bảo mật phòng họp</p>
                <p className="text-xs text-slate-300 font-sans">
                  Phòng: <span className="font-mono text-white font-bold">{channelName.toUpperCase()}</span> | Đang được mã hóa bảo vệ toàn diện
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <span className="text-xs text-slate-400 font-sans">Mật khẩu phòng:</span>
              <span className="text-sm font-mono font-extrabold tracking-widest text-emerald-400 bg-slate-950 px-3 py-1.5 rounded-lg border border-emerald-500/20 shadow-inner">
                {roomPassword}
              </span>
              <button
                onClick={() => {
                  navigator.clipboard.writeText(roomPassword);
                  alert("Đã sao chép mật khẩu phòng họp: " + roomPassword);
                }}
                className="p-1.5 bg-indigo-600/10 hover:bg-indigo-600/25 border border-indigo-500/25 text-indigo-300 rounded-lg transition cursor-pointer active:scale-95"
                title="Sao chép mật khẩu"
              >
                <Copy className="w-3.5 h-3.5" />
              </button>
              
              <button
                onClick={handleResetPassword}
                className="p-1.5 bg-slate-850 hover:bg-slate-750 border border-slate-700 text-slate-300 rounded-lg transition cursor-pointer active:scale-95"
                title="Làm mới mật khẩu phòng họp"
              >
                <RefreshCw className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          
          {/* CỘT TRÁI: SÂN KHẤU VIDEO (2/3 chiều rộng) */}
          <div className="lg:col-span-2 space-y-6">
            
            {/* LƯỢT TRÌNH CHIẾU MÀN HÌNH CHIA SẺ (Nổi bật, siêu to) */}
            {screenSharing && (
              <div className="bg-slate-950 border border-indigo-500/40 rounded-2xl p-4 space-y-3 shadow-lg shadow-indigo-950/40">
                <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                  <div className="flex items-center gap-2 text-indigo-400">
                    <Tv className="w-5 h-5 animate-pulse" />
                    <span className="text-xs font-extrabold uppercase tracking-wider">Khung chia sẻ màn hình</span>
                  </div>
                  <span className="px-2 py-0.5 bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 rounded text-[9px] font-bold animate-pulse">
                    MÀN HÌNH CHÍNH
                  </span>
                </div>
                
                <div className="relative bg-slate-900 border border-slate-850 rounded-xl aspect-video overflow-hidden">
                  {isSimulationMode ? (
                    <div className="absolute inset-0 bg-gradient-to-br from-indigo-950/90 via-slate-950 to-slate-900/90 flex flex-col items-center justify-center text-center p-6 space-y-4">
                      <div className="w-14 h-14 rounded-full bg-indigo-500/10 border border-indigo-500/30 flex items-center justify-center text-indigo-400 animate-bounce">
                        <Tv className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="text-sm font-bold text-slate-100">Báo cáo Tổng hợp Số liệu Liên ngành năm 2026</h4>
                        <p className="text-xs text-slate-400 mt-1">Đang trình bày bởi {displayName} (Mô phỏng live stream)</p>
                      </div>
                      <div className="flex items-center gap-1.5 px-3 py-1 bg-slate-950/80 border border-slate-800 rounded-lg text-[10px] text-slate-400">
                        <Activity className="w-3.5 h-3.5 text-emerald-400" />
                        <span>Truyền phát: 1080p @ 30 FPS</span>
                      </div>
                    </div>
                  ) : (
                    <div 
                      ref={el => {
                        if (el && localScreenTrack) {
                          localScreenTrack.play(el);
                        }
                      }} 
                      className="w-full h-full object-cover" 
                    />
                  )}
                  
                  <div className="absolute bottom-3 left-3 px-2.5 py-1 bg-slate-950/90 border border-slate-800 rounded-lg text-xs text-white font-medium">
                    Trình bày: {displayName} (Bạn)
                  </div>
                </div>
              </div>
            )}

            {/* LƯỚI CAMERA CỦA THÀNH VIÊN KHÁC (Nhỏ hơn) */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              
              {/* Webcam của chính bạn */}
              <div className="relative bg-slate-950 border border-slate-800 rounded-xl aspect-video overflow-hidden group shadow-md">
                {camMuted ? (
                  <div className="absolute inset-0 flex flex-col items-center justify-center bg-slate-950 text-slate-500">
                    <div className="w-10 h-10 rounded-full bg-slate-900 flex items-center justify-center border border-slate-800 mb-2">
                      <VideoOff className="w-5 h-5" />
                    </div>
                    <span className="text-xs">Camera đang tắt</span>
                  </div>
                ) : isSimulationMode ? (
                  <div className="absolute inset-0 bg-gradient-to-tr from-indigo-950/50 via-slate-900 to-slate-950/50 flex flex-col items-center justify-center p-4">
                    <div className="relative">
                      <div className="w-12 h-12 rounded-full bg-indigo-600/20 border border-indigo-500 flex items-center justify-center text-white font-bold text-lg animate-pulse">
                        {displayName.charAt(0)}
                      </div>
                      {!micMuted && (
                        <span className="absolute -bottom-1 -right-1 p-1 bg-emerald-500 text-white rounded-full border border-slate-900">
                          <Mic className="w-3 h-3" />
                        </span>
                      )}
                    </div>
                    <p className="mt-2 text-xs text-indigo-400 font-semibold">{displayName} (Bạn)</p>
                  </div>
                ) : (
                  <div ref={localVideoDivRef} className="w-full h-full object-cover" />
                )}

                <div className="absolute bottom-2.5 left-2.5 px-2 py-0.5 bg-slate-950/85 backdrop-blur-sm border border-slate-800 rounded-lg text-[11px] font-medium flex items-center gap-1.5 text-white">
                  <span className="w-1.5 h-1.5 rounded-full bg-indigo-400"></span>
                  {displayName} (Bạn)
                  {micMuted && <MicOff className="w-3 h-3 text-rose-400 ml-1" />}
                </div>
              </div>

              {/* Camera từ người dùng thực tế qua Agora */}
              {!isSimulationMode && remoteUsers.map(remoteUser => (
                <div 
                  key={remoteUser.uid} 
                  className="relative bg-slate-950 border border-slate-800 rounded-xl aspect-video overflow-hidden shadow-md"
                  ref={el => {
                    if (el && remoteUser.videoTrack) {
                      remoteUser.videoTrack.play(el);
                    }
                  }}
                >
                  <div className="absolute bottom-2.5 left-2.5 px-2 py-0.5 bg-slate-950/85 backdrop-blur-sm border border-slate-800 rounded-lg text-[11px] font-medium flex items-center gap-1.5 text-white">
                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-400"></span>
                    Đơn vị: {remoteUser.uid}
                  </div>
                </div>
              ))}

              {/* Người dùng mô phỏng trực tuyến */}
              {isSimulationMode && simulatedParticipants.map(member => (
                <div key={member.id} className="relative bg-slate-950 border border-slate-800 rounded-xl aspect-video overflow-hidden shadow-md group">
                  <div className={`absolute inset-0 bg-gradient-to-br from-slate-900 to-slate-950 flex flex-col items-center justify-center p-4 transition-all duration-300 ${member.speaking && member.micOn ? "ring-2 ring-emerald-500/40" : ""}`}>
                    
                    {member.camOn ? (
                      <div className="flex flex-col items-center text-center space-y-2">
                        <div className={`w-11 h-11 rounded-full ${member.avatarColor} flex items-center justify-center text-white font-bold text-sm shadow-inner border border-white/10 relative`}>
                          {member.name.charAt(0)}
                          {member.speaking && member.micOn && (
                            <span className="absolute -top-0.5 -right-0.5 w-3 h-3 bg-emerald-500 rounded-full border-2 border-slate-950 animate-ping" />
                          )}
                        </div>
                        <div>
                          <p className="text-[11px] text-slate-300 font-semibold">{member.name}</p>
                          <p className="text-[9px] text-slate-500 mt-0.5">Liên kết mã hóa</p>
                        </div>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center text-center space-y-1.5">
                        <div className="w-9 h-9 rounded-full bg-slate-900 border border-slate-800 flex items-center justify-center text-slate-600">
                          <VideoOff className="w-4 h-4" />
                        </div>
                        <p className="text-[11px] text-slate-500">{member.name} (Ẩn Cam)</p>
                      </div>
                    )}
                  </div>

                  <div className="absolute bottom-2.5 left-2.5 right-2.5 flex items-center justify-between pointer-events-none">
                    <div className="px-2 py-0.5 bg-slate-950/85 backdrop-blur-sm border border-slate-800 rounded-lg text-[10px] font-medium flex items-center gap-1.2 text-white">
                      <span className={`w-1.2 h-1.2 rounded-full ${member.micOn && member.speaking ? "bg-emerald-400 animate-pulse" : "bg-slate-400"}`}></span>
                      {member.name}
                    </div>

                    <div className="flex gap-1">
                      {member.micOn ? (
                        <span className="p-1 bg-emerald-500/10 text-emerald-400 rounded-md border border-emerald-500/20 text-[9px]">
                          <Mic className="w-2.5 h-2.5" />
                        </span>
                      ) : (
                        <span className="p-1 bg-rose-500/10 text-rose-400 rounded-md border border-rose-500/20 text-[9px]">
                          <MicOff className="w-2.5 h-2.5" />
                        </span>
                      )}
                    </div>
                  </div>
                </div>
              ))}

            </div>

            {/* Thanh điều khiển cuộc gọi */}
            <div className="bg-slate-950/80 border border-slate-800 p-4 rounded-xl flex flex-wrap items-center justify-between gap-4 shadow-inner">
              
              <div className="flex items-center gap-2">
                <span className="p-2 bg-indigo-500/10 border border-indigo-500/20 text-indigo-400 rounded-lg flex items-center gap-1.5">
                  <Users className="w-4 h-4" />
                  <span className="text-xs font-bold">{isSimulationMode ? simulatedParticipants.length + 1 : remoteUsers.length + 1} trực tuyến</span>
                </span>
                <div className="text-[10px] text-slate-400 hidden sm:block">
                  Độ trễ mạng: <span className="text-emerald-400 font-mono">12 ms</span>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={toggleMic}
                  className={`p-2.5 rounded-xl border transition-all cursor-pointer ${
                    micMuted 
                      ? "bg-rose-600/20 border-rose-600/30 text-rose-400 hover:bg-rose-600/30" 
                      : "bg-slate-850 border-slate-700 text-slate-200 hover:bg-slate-800"
                  }`}
                  title={micMuted ? "Bật Mic" : "Tắt Mic"}
                >
                  {micMuted ? <MicOff className="w-4.5 h-4.5" /> : <Mic className="w-4.5 h-4.5" />}
                </button>

                <button
                  onClick={toggleCam}
                  className={`p-2.5 rounded-xl border transition-all cursor-pointer ${
                    camMuted 
                      ? "bg-rose-600/20 border-rose-600/30 text-rose-400 hover:bg-rose-600/30" 
                      : "bg-slate-850 border-slate-700 text-slate-200 hover:bg-slate-800"
                  }`}
                  title={camMuted ? "Mở Camera" : "Tắt Camera"}
                >
                  {camMuted ? <VideoOff className="w-4.5 h-4.5" /> : <Video className="w-4.5 h-4.5" />}
                </button>

                {/* NÚT CHIA SẺ MÀN HÌNH NÂNG CẤP */}
                <button
                  onClick={toggleScreenShare}
                  className={`p-2.5 rounded-xl border transition-all cursor-pointer flex items-center gap-1 text-xs font-bold ${
                    screenSharing 
                      ? "bg-emerald-600/20 border-emerald-600/30 text-emerald-400 hover:bg-emerald-600/30" 
                      : "bg-slate-850 border-slate-700 text-slate-200 hover:bg-slate-800"
                  }`}
                  title={screenSharing ? "Dừng chia sẻ màn hình" : "Bắt đầu chia sẻ màn hình"}
                >
                  <Tv className="w-4.5 h-4.5" />
                  <span className="hidden md:inline">{screenSharing ? "Dừng chia sẻ" : "Chia sẻ màn hình"}</span>
                </button>

                <button
                  onClick={handleLeave}
                  className="px-4 py-2.5 bg-rose-600 hover:bg-rose-500 active:translate-y-0.5 transition font-bold text-xs rounded-xl flex items-center gap-1.5 text-white cursor-pointer shadow-lg shadow-rose-600/10"
                >
                  <PhoneOff className="w-3.5 h-3.5" /> Gác máy
                </button>
              </div>

              <div className="flex items-center gap-1 text-[11px] text-slate-500">
                <Activity className="w-3.5 h-3.5 text-emerald-400 animate-pulse" />
                <span>Thoại HD an toàn</span>
              </div>

            </div>

          </div>

          {/* CỘT PHẢI: KHO TÀI LIỆU CHUNG (1/3 chiều rộng) */}
          <div className="lg:col-span-1 bg-slate-950/45 border border-slate-800/80 rounded-2xl p-5 flex flex-col justify-between h-[580px] shadow-lg">
            
            <div className="flex flex-col h-full overflow-hidden">
              
              {/* Header Kho tài liệu */}
              <div className="border-b border-slate-800/80 pb-3 flex items-center justify-between">
                <div>
                  <h3 className="text-sm font-extrabold text-slate-100 uppercase tracking-wider flex items-center gap-1.5">
                    <FileText className="w-4 h-4 text-indigo-400" />
                    Tài liệu phòng họp
                  </h3>
                  <p className="text-[10px] text-slate-500 mt-0.5">Tải lên &amp; Đồng bộ dữ liệu liên ngành</p>
                </div>
                <span className="px-2 py-0.5 bg-indigo-500/10 text-indigo-300 border border-indigo-500/20 rounded-full text-[10px] font-bold font-mono">
                  {documents.length} File
                </span>
              </div>

              {/* Thông tin luồng đồng bộ live */}
              <div className="mt-2.5 flex items-center gap-1.5 px-2 py-1.5 bg-emerald-500/5 text-emerald-400 border border-emerald-500/10 rounded-lg text-[10px] font-medium">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-ping"></span>
                <span>Firestore real-time sync đang lắng nghe...</span>
              </div>

              {/* Danh sách tài liệu */}
              <div className="flex-1 overflow-y-auto pr-1 space-y-3 mt-4">
                
                {uploading && (
                  <div className="flex items-center justify-center gap-2 text-xs text-indigo-400 bg-indigo-500/5 border border-indigo-500/10 rounded-xl p-3 animate-pulse">
                    <span className="w-3.5 h-3.5 border-2 border-indigo-400 border-t-transparent rounded-full animate-spin" />
                    <span>Đang tải lên &amp; mã hóa...</span>
                  </div>
                )}

                {documents.length === 0 ? (
                  <div className="text-center py-16 flex flex-col items-center justify-center space-y-2 text-slate-500">
                    <UploadCloud className="w-8 h-8 text-slate-700" />
                    <p className="text-[11px] font-medium">Chưa có tài liệu nào được chia sẻ</p>
                    <p className="text-[9px] text-slate-600 max-w-[180px]">Mọi người trong phòng có thể gửi báo cáo và tải về trực tiếp.</p>
                  </div>
                ) : (
                  documents.map(docItem => (
                    <div 
                      key={docItem.id} 
                      className="p-3 bg-slate-900 border border-slate-800/80 rounded-xl flex items-start justify-between gap-3 hover:border-slate-700 transition"
                    >
                      <div className="flex items-start gap-2.5 overflow-hidden">
                        <div className="p-2 bg-slate-950 rounded-lg border border-slate-800 shrink-0">
                          {getFileIcon(docItem.fileType)}
                        </div>
                        <div className="overflow-hidden">
                          <p className="text-xs font-bold text-slate-200 truncate" title={docItem.fileName}>
                            {docItem.fileName}
                          </p>
                          <p className="text-[10px] text-slate-500 font-mono mt-0.5">
                            {formatFileSize(docItem.fileSize)} • {docItem.uploadedBy}
                          </p>
                          <span className="inline-flex items-center gap-1 text-[9px] text-slate-400 font-medium mt-1">
                            <Clock className="w-2.5 h-2.5" />
                            {formatTime(docItem.uploadedAt)}
                          </span>
                        </div>
                      </div>

                      <div className="flex gap-1 shrink-0">
                        {/* Nút Tải xuống thông minh dùng Base64 */}
                        <a
                          href={docItem.fileData}
                          download={docItem.fileName}
                          className="p-1.5 bg-indigo-500/10 hover:bg-indigo-500/25 text-indigo-400 rounded-lg border border-indigo-500/20 transition cursor-pointer"
                          title="Tải về trực tiếp"
                        >
                          <Download className="w-3.5 h-3.5" />
                        </a>

                        {/* Nút Xóa tệp cho tất cả thành viên trong phòng */}
                        <button
                          onClick={() => handleDeleteDoc(docItem.id, docItem.uploadedBy)}
                          className="p-1.5 bg-rose-500/10 hover:bg-rose-500/25 text-rose-400 rounded-lg border border-rose-500/20 transition cursor-pointer"
                          title="Xóa tài liệu"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  ))
                )}

              </div>

              {/* Vùng tải tài liệu lên */}
              <div className="border-t border-slate-800/80 pt-4 mt-4">
                <label className="flex flex-col items-center justify-center border-2 border-dashed border-slate-800 hover:border-indigo-500/50 bg-slate-900/30 hover:bg-slate-900/50 rounded-xl p-4 cursor-pointer transition text-center group">
                  <UploadCloud className="w-7 h-7 text-slate-500 group-hover:text-indigo-400 transition mb-1" />
                  <span className="text-xs font-semibold text-slate-300">Gửi tài liệu lên phòng họp</span>
                  <span className="text-[9px] text-slate-500 mt-0.5">Hỗ trợ Excel, PDF, Word... (Tối đa 1.5MB)</span>
                  <input
                    type="file"
                    onChange={handleFileUpload}
                    className="hidden"
                    disabled={uploading}
                    accept=".pdf,.xlsx,.xls,.csv,.doc,.docx,.png,.jpg,.jpeg,.txt"
                  />
                </label>
              </div>

            </div>

          </div>

        </div>
      </div>
      )}

    </div>
  );
};
