import os, requests
host = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434").rstrip("/")
url = f"{host}/api/chat"
print(f"→ Requesting: {url}")  # 关键：确认输出是否含重复路径

try:
    resp = requests.post(
        url,
        json={"model": "llama3", "messages": [{"role": "user", "content": "hi"}], "stream": False},
        timeout=5
    )
    print(f"Status: {resp.status_code}\nResponse: {resp.text}")
except Exception as e:
    print(f"Error: {e}")