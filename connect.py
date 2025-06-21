import asyncio
import websockets
from flask import Flask, request, jsonify
from threading import Thread
from myfunctions import sendMessage

app = Flask(__name__)
websocket_connection = None
websocket_event_loop = None

async def connect_to_websocket():
    
    global websocket_connection,websocket_event_loop
    async def process(message):
        print(message)
    url = "wss://devgadbadr.com:3005"
    websocket_connection = await websockets.connect(url)
    websocket_event_loop = asyncio.get_event_loop()
    print("Connected to WebSocket server")
    async for message in websocket_connection:
        await process(message)
        
def start_websocket():
    asyncio.run(connect_to_websocket())

@app.route('/send', methods=['POST'])
def send_message():
    global websocket_connection
    data = request.json
    message = data.get("message", "")
    print(data)
    if websocket_connection:
        asyncio.run_coroutine_threadsafe(
            websocket_connection.send(message),
            websocket_event_loop
        )
        return jsonify({"status": "Message sent", "message": message})
    else:
        return jsonify({"error": "WebSocket connection not established"}), 500

def start_flask():
   app.run(host='0.0.0.0', port=3006, ssl_context=('../../SSL/devgadbadr.com_full.crt', '../../SSL/devgadbadr.com_key.txt'))
   
def connectAppToWebsocket():
    websocket_thread = Thread(target=start_websocket)
    websocket_thread.start()
    flaskThread = Thread(target=start_flask)
    flaskThread.start()

