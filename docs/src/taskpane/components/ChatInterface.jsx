import React, { useState, useEffect } from 'react';

function ChatInterface({ websocket }) {
    const [message, setMessage] = useState('');
    const [isWebSocketConnected, setIsWebSocketConnected] = useState(false);

    useEffect(() => {
        if (websocket) {
            setIsWebSocketConnected(websocket.readyState === WebSocket.OPEN);

            websocket.onopen = () => setIsWebSocketConnected(true);
            websocket.onclose = () => setIsWebSocketConnected(false);
        }
    }, [websocket]);

    const sendMessage = () => {
        if (message && websocket && isWebSocketConnected) {
            websocket.send(message);
            setMessage('');
        }
    };

    return (
        <div>
            <input
                type="text"
                value={message}
                onChange={(e) => setMessage(e.target.value)}
                placeholder="Type your question here..."
            />
            <button onClick={sendMessage} disabled={!isWebSocketConnected}>
                Send
            </button>
        </div>
    );
}

export default ChatInterface;
