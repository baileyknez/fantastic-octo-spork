import React, { useState, useEffect } from "react";
import PropTypes from "prop-types";
import ChatInterface from "./ChatInterface";
import MessageHistory from "./MessageHistory";
import { makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props) => {
  const styles = useStyles();
  const [messages, setMessages] = useState([]); // State for storing messages
  const [websocket, setWebsocket] = useState(null); // State for WebSocket connection

  useEffect(() => {
    const ws = new WebSocket('wss://intellisync.ai/plugin_chat'); // Replace with actual URL
    setWebsocket(ws);

    ws.onopen = () => {
      console.log("WebSocket connection established.");
    };

    ws.onmessage = (event) => {
      // Handle incoming messages
      const receivedMessage = { content: event.data, type: 'received' };
      setMessages((prevMessages) => [...prevMessages, receivedMessage]);
    };

    ws.onerror = (event) => {
      console.error("WebSocket error observed:", event);
    };

    ws.onclose = (event) => {
      console.log("WebSocket connection closed:", event.reason);
      // Optional: implement reconnection logic here
    };

    return () => {
      if (ws) {
        ws.close();
      }
    };
  }, []);


  const sendMessage = (messageContent) => {
    // Function to send messages and update state
    if (messageContent && websocket && websocket.readyState === WebSocket.OPEN) {
      websocket.send(messageContent);
      const sentMessage = { content: messageContent, type: 'sent' };
      setMessages((prevMessages) => [...prevMessages, sentMessage]);
    }
  };

  return (
    <div className={styles.root}>
      <MessageHistory messages={messages} />
      <ChatInterface websocket={websocket} onSendMessage={sendMessage} />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string, // You might want to remove this if the title is no longer used
};

export default App;
