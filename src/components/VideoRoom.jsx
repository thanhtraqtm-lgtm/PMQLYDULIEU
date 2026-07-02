import React, { useEffect } from 'react';
import AgoraRTC from 'agora-rtc-sdk-ng';

const client = AgoraRTC.createClient({ mode: "rtc", codec: "vp8" });

const VideoRoom = ({ appId, channel }) => {
  useEffect(() => {
    const init = async () => {
      await client.join(appId, channel, null, null);
      const localTracks = await AgoraRTC.createMicrophoneAndCameraTracks();
      await client.publish(localTracks);
      localTracks[1].play("local-player"); // Gắn video vào div có id là local-player
    };
    init();
    return () => { client.leave(); };
  }, [appId, channel]);

  return <div id="local-player" style={{ width: '100%', height: '400px' }} />;
};

export default VideoRoom;
