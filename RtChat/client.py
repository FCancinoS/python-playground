import socket
import threading
from datetime import datetime
import os
import random
import string

SERVER_IP = '192.168.1.72'  # Cambia esto por la IP de tu laptop
SERVER_PORT = 12345

def receive_messages(sock):
    while True:
        try:
            message = sock.recv(1024).decode()
            print(f"\n {message}\n> ", end="")
        except:
            print("\n❌ Conexión cerrada.")
            break

def generate_fake_code():
    lorem = [
        "Lorem ipsum dolor sit amet \n psum has been the industry's standard \n dummy text ever since the 1500s, \n when an unknown printer took a \n galley of type and ,", "consectetur adipiscing elit, \n psum has been the industry's standard \n dummy text ever since the 1500s, \n when an unknown printer took a \n galley of type and ", 
        "sed do eiusmod  \n psum has been the industry's standard \n dummy text ever since the 1500s, \n when an unknown printer took a \n galley of type and tempor incididunt ut labore et dolore magna aliqua.",
        "Ut enim ad mi \n psum has been the industry's standard \n dummy text ever since the 1500s, \n when an unknown printer took a \n galley of type and nim veniam,", "quis nostrud exercitation ullamco l \n psum has been the industry's standard \n dummy text ever since the 1500s, \n when an unknown printer took a \n galley of type and aboris nisi ut aliquip ex ea commodo consequat."
    ]
    noise = ''.join(random.choices(string.ascii_letters + string.digits, k=20))
    return f"{random.choice(lorem)} // {noise}"

def send_messages(sock, username):
    while True:
        msg = input("> ")

        if msg.lower() == "clear":
            os.system("cls" if os.name == "nt" else "clear")
            continue

        if msg.lower() == "code":

            msg = generate_fake_code()
            print(msg)

        timestamp = datetime.now().strftime("%H:%M")
        full_msg = f"{timestamp} {username}: {msg}"
        sock.send(full_msg.encode())

def start_client():
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.connect((SERVER_IP, SERVER_PORT))
        print(f" Conectado a {SERVER_IP}:{SERVER_PORT}")
    except Exception as e:
        print(f"❌ Error de conexión: {e}")
        return

    username = input("Ingresa tu alias: ")

    threading.Thread(target=receive_messages, args=(sock,), daemon=True).start()
    send_messages(sock, username)

if __name__ == "__main__":
    start_client()
