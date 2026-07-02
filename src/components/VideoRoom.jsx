import React, { useEffect, useRef } from 'react';
import AgoraRTC from 'agora-rtc-sdk-ng';

const client = AgoraRTC.createClient({ mode: "rtc", codec: "vp8" });

const VideoRoom = ({ appId, channel }) => {
  const localPlayerRef = useRef(null);

  useEffect(() => {
    const init = async () => {
      await client.join(appId, channel, null, null);
      const tracks = await AgoraRTC.createMicrophoneAndCameraTracks();
      await client.publish(tracks);
      
      // Hiển thị video của chính mình
      tracks[1].play(localPlayerRef.current);
    };
    init();
    return () => { client.leave(); };
  }, [appId, channel]);

  return <div ref={localPlayerRef} style={{ width: '100%', height: '500px', backgroundColor: '#000' }} />;
};

export default VideoRoom;
