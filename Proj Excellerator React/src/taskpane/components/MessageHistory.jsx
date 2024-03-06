import React from 'react';

function MessageHistory({ messages }) {
    return (
        <div className="message-history">
            {messages.map((msg, index) => (
                <div key={index} className={`message ${msg.type}`}>
                    {msg.content}
                </div>
            ))}
        </div>
    );
}

export default MessageHistory;
