import React, { useEffect, useRef } from 'react';
import AgoraRTC from 'agora-rtc-sdk-ng';

const client = AgoraRTC.createClient({ mode: "rtc", codec: "vp8" });

export default function VideoRoom({ appId, channel }) {
  const localRef = useRef(null);

  useEffect(() => {
    let localTracks = [];
    const init = async () => {
      await client.join(appId, channel, null, null);
      localTracks = await AgoraRTC.createMicrophoneAndCameraTracks();
      await client.publish(localTracks);
      localTracks[1].play(localRef.current);
    };
    init();
    return () => { 
      localTracks.forEach(t => t.close());
      client.leave(); 
    };
  }, [appId, channel]);

  return <div ref={localRef} style={{ width: '100%', height: '500px', background: '#000' }} />;
}
