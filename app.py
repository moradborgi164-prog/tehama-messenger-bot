import os
import json
import logging
import requests
import openpyxl
from datetime import datetime
from flask import Flask, request
from openai import OpenAI
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

app = Flask(__name__)

# =========================
# ENV VARIABLES (Render)
# =========================
PAGE_ACCESS_TOKEN = os.environ.get("PAGE_ACCESS_TOKEN", "")
VERIFY_TOKEN = os.environ.get("VERIFY_TOKEN", "")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")

RECIPIENTS = os.environ.get("RECIPIENTS", "").split(",")

FB_MESSAGES_URL = "https://graph.facebook.com/v19.0/me/messages"

# =========================
# OPENAI CLIENT
# =========================
client = OpenAI(api_key=OPENAI_API_KEY)

# =========================
# SYSTEM PROMPT
# =========================
SYSTEM_PROMPT = """
أنت مساعد ذكي لشركة تهامة التجارية في اليمن (CAT).
أجب بشكل JSON فقط:

{
 "intent": "find_rep | human_agent | greeting | unclear",
 "product": "قطع الغيار | معدات | مولدات | null",
 "governorate": "صنعاء | الحديدة | عدن | تعز | المكلا | null",
 "reply": "رد مختصر ومهني بالعربية",
 "missing": "product | governorate | both | none"
}

إذا لم تفهم، اجعل intent = unclear
"""

# =========================
# AI FUNCTION
# =========================
def ask_ai(message):
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": message}
            ]
        )

        raw = response.choices[0].message.content

        log.info(f"AI RAW: {raw}")

        # نحاول نفهم إذا رجع JSON
        try:
            return json.loads(raw)
        except:
            # لو ما كان JSON، نخليه رد عادي
            return {
                "intent": "greeting",
                "product": None,
                "governorate": None,
                "reply": raw,
                "missing": "none"
            }

    except Exception as e:
        log.error(f"AI ERROR: {e}")
        return {
            "intent": "unclear",
            "product": None,
            "governorate": None,
            "reply": "عذرًا، حدث خطأ في الذكاء الاصطناعي.",
            "missing": "both"
        }
# =========================
# SEND MESSAGE
# =========================
def send_message(uid, text):
    requests.post(
        FB_MESSAGES_URL,
        params={"access_token": PAGE_ACCESS_TOKEN},
        json={
            "recipient": {"id": uid},
            "message": {"text": text}
        }
    )

# =========================
# WEBHOOK VERIFY
# =========================
@app.route("/webhook", methods=["GET"])
def verify():
    if request.args.get("hub.verify_token") == VERIFY_TOKEN:
        return request.args.get("hub.challenge")
    return "error", 403

# =========================
# WEBHOOK RECEIVE
# =========================
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json()

    for entry in data.get("entry", []):
        for event in entry.get("messaging", []):

            uid = event["sender"]["id"]

            if "message" in event and "text" in event["message"]:
                text = event["message"]["text"]

                result = ask_ai(text)

                send_message(uid, result.get("reply", ""))
    return "ok", 200

# =========================
@app.route("/", methods=["GET"])
def home():
    return "Tehama Bot Running", 200

if __name__ == "__main__":
    app.run()
