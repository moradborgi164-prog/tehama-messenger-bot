import os
import json
import logging
from datetime import datetime

import requests
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from flask import Flask, request
from openai import OpenAI

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)
app = Flask(__name__)

# =========================
# Environment Variables
# =========================
PAGE_ACCESS_TOKEN = os.getenv("PAGE_ACCESS_TOKEN", "")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")

RECIPIENTS_RAW = os.getenv("RECIPIENTS", "")
RECIPIENTS = [x.strip() for x in RECIPIENTS_RAW.split(",") if x.strip()]

BOT_ENABLED = os.getenv("BOT_ENABLED", "true").lower() == "true"
LEADS_FILE = os.getenv("LEADS_FILE", "messenger_leads.xlsx")

FB_MESSAGES_URL = "https://graph.facebook.com/v19.0/me/messages"

# =========================
# OpenAI client
# =========================
client = OpenAI(api_key=OPENAI_API_KEY)

# =========================
# Sales contacts
# =========================
CONTACTS = {
    "صنعاء": {
        "قطع الغيار": {"name": "مراد قاسم",  "phones": ["733715113", "772814400"]},
        "معدات":      {"name": "بسام عقلان", "phones": ["730099119"]},
        "مولدات":     {"name": "مروان فضل",  "phones": ["730512588"]},
    },
    "الحديدة": {
        "قطع الغيار": {"name": "عرفات عقيل", "phones": ["737522352"]},
        "معدات":      {"name": "محمد بورجي", "phones": ["737889508"]},
        "مولدات":     {"name": "علي حسين",   "phones": ["737522313"]},
    },
    "تعز": {
        "قطع الغيار": {"name": "عبدالرقيب", "phones": ["737522362"]},
        "معدات":      {"name": "حبيب",      "phones": ["737889509"]},
        "مولدات":     {"name": "حبيب",      "phones": ["737889509"]},
    },
    "عدن": {
        "قطع الغيار": {"name": "عبدالناصر", "phones": ["739177443"]},
        "معدات":      {"name": "توفيق",     "phones": ["737522398"]},
        "مولدات":     {"name": "توفيق",     "phones": ["737522398"]},
    },
    "المكلا": {
        "قطع الغيار": {"name": "طارق",      "phones": ["737522393"]},
        "معدات":      {"name": "كمال",      "phones": ["737522333"]},
        "مولدات":     {"name": "كمال",      "phones": ["737522333"]},
    },
}

# جلسات بسيطة داخل الذاكرة
user_sessions = {}

SYSTEM_PROMPT = """
أنت مساعد ذكي لشركة تهامة التجارية في اليمن.
اسم الصفحة: CAT-Yemen-Tehama Trading Co
الشركة وكيل معتمد لـ CAT وتبيع: مولدات، معدات ثقيلة، قطع غيار.
فروع الشركة: صنعاء، الحديدة، عدن، تعز، المكلا.

مهمتك: فهم رسالة العميل واستخراج المعلومات التالية بصيغة JSON فقط:

{
  "intent": "find_rep | human_agent | general_question | greeting | unclear",
  "product": "قطع الغيار | معدات | مولدات | null",
  "governorate": "صنعاء | الحديدة | عدن | تعز | المكلا | null",
  "reply": "رد طبيعي ومناسب للعميل بالعربي والانجليزي",
  "missing": "product | governorate | both | none"
}

قواعد:
- إذا ذكر العميل CAT أو كات أو مولد أو generator = مولدات
- إذا ذكر حفار أو رافعة أو معدة = معدات
- إذا ذكر قطعة أو سبيرات أو spare parts = قطع الغيار
- إذا ذكر العاصمة أو أمانة = صنعاء
- إذا أراد موظف أو إنسان = human_agent
- reply يجب أن يكون ودياً ومختصراً
- لا تكتب أي شيء خارج JSON
"""

# =========================
# Helpers
# =========================
def ask_ai(user_message, session):
    """استخدام OpenAI لفهم رسالة العميل وإرجاع JSON فقط"""
    try:
        history = session.get("history", [])
        history.append({"role": "user", "content": user_message})

        messages = [{"role": "system", "content": SYSTEM_PROMPT}]
        messages.extend(history[-10:])

        response = client.chat.completions.create(
            model="gpt-4.1-mini",
            temperature=0.2,
            messages=messages
        )

        raw = response.choices[0].message.content.strip()
        raw = raw.replace("```json", "").replace("```", "").strip()
        result = json.loads(raw)

        history.append({"role": "assistant", "content": raw})
        session["history"] = history[-10:]

        return result

    except Exception as e:
        log.exception(f"AI error: {e}")
        return {
            "intent": "unclear",
            "product": None,
            "governorate": None,
            "reply": (
                "عذراً، لم أفهم سؤالك. هل تبحث عن قطع غيار أو معدات أو مولدات؟\n"
                "Sorry, could you clarify? Are you looking for spare parts, equipment, or generators?"
            ),
            "missing": "both"
        }


def send_fb_request(payload):
    try:
        r = requests.post(
            FB_MESSAGES_URL,
            params={"access_token": PAGE_ACCESS_TOKEN},
            json=payload,
            timeout=30
        )
        if r.status_code >= 400:
            log.error(f"Facebook API error {r.status_code}: {r.text}")
        return r
    except Exception as e:
        log.exception(f"Facebook request failed: {e}")
        return None


def send_text(uid, text):
    payload = {
        "recipient": {"id": uid},
        "message": {"text": text}
    }
    send_fb_request(payload)


def send_quick_replies(uid, text, options):
    qr = [
        {
            "content_type": "text",
            "title": o["title"][:20],   # فيسبوك يفضل العنوان قصير
            "payload": o["payload"]
        }
        for o in options
    ]
    payload = {
        "recipient": {"id": uid},
        "message": {"text": text, "quick_replies": qr}
    }
    send_fb_request(payload)


def save_lead(name, phone, gov, product):
    try:
        try:
            wb = openpyxl.load_workbook(LEADS_FILE)
            ws = wb.active
        except Exception:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "العملاء"
            ws.append(["التاريخ", "الاسم", "الهاتف", "المحافظة", "المنتج"])

        ws.append([datetime.now().strftime("%Y-%m-%d %H:%M"), name, phone, gov, product])
        wb.save(LEADS_FILE)
    except Exception as e:
        log.exception(f"Excel error: {e}")


def send_email(subject, body):
    if not SMTP_USER or not SMTP_PASSWORD or not RECIPIENTS:
        log.warning("SMTP settings or recipients are missing; email skipped.")
        return

    try:
        msg = MIMEMultipart()
        msg["Subject"] = subject
        msg["From"] = SMTP_USER
        msg["To"] = ", ".join(RECIPIENTS)
        msg.attach(MIMEText(body, "plain", "utf-8"))

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(SMTP_USER, SMTP_PASSWORD)
            smtp.sendmail(SMTP_USER, RECIPIENTS, msg.as_string())
    except Exception as e:
        log.exception(f"Email error: {e}")


def send_email_lead(name, phone, gov, product, rep_name, rep_phones):
    rep_str = " / ".join(rep_phones)
    body = f"""
━━━━━━━━━━━━━━━━━━━━━━━━
🔔 عميل جديد — ماسنجر AI
━━━━━━━━━━━━━━━━━━━━━━━━
👤 الاسم    : {name}
📞 الهاتف   : {phone}
📍 المحافظة : {gov}
🛒 المنتج   : {product}
👨 المندوب  : {rep_name} | {rep_str}
🕐 الوقت    : {datetime.now().strftime("%Y-%m-%d %H:%M")}
━━━━━━━━━━━━━━━━━━━━━━━━
"""
    send_email(
        subject=f"🔔 عميل جديد AI | {name} | {product}",
        body=body
    )


def send_email_human(uid, name, phone):
    body = f"""
━━━━━━━━━━━━━━━━━━━━━━━━
🆘 عاجل! عميل يطلب موظف
━━━━━━━━━━━━━━━━━━━━━━━━
👤 الاسم  : {name or 'غير محدد'}
📞 الهاتف : {phone or 'غير محدد'}
🕐 الوقت  : {datetime.now().strftime("%Y-%m-%d %H:%M")}
🔗 المحادثة: https://www.facebook.com/messages/t/{uid}
⚠️ الرجاء الرد اليدوي على العميل
━━━━━━━━━━━━━━━━━━━━━━━━
"""
    send_email(
        subject=f"🆘 عاجل! موظف مطلوب — {name or 'عميل'}",
        body=body
    )


def send_contact_card(uid, gov, product, session):
    contact = CONTACTS.get(gov, {}).get(product)
    name = session.get("name", "")
    phone = session.get("phone", "")

    if contact:
        phones_str = "\n".join(f"📞 {p}" for p in contact["phones"])
        send_text(
            uid,
            f"✅ وجدنا المندوب المناسب!\n"
            f"✅ Found the right sales rep!\n\n"
            f"📍 المحافظة: {gov}\n"
            f"🛒 المنتج: {product}\n"
            f"👤 الاسم: {contact['name']}\n"
            f"{phones_str}\n\n"
            f"تواصل معه مباشرة 🙏"
        )
        save_lead(name, phone, gov, product)
        send_email_lead(name, phone, gov, product, contact["name"], contact["phones"])
    else:
        send_text(uid, "⚠️ لا يوجد مندوب متاح حالياً.\nNo rep available now.")

    send_quick_replies(
        uid,
        "هل تحتاج شيئاً آخر؟ | Anything else?",
        [
            {"title": "استفسار جديد", "payload": "RESTART"},
            {"title": "تحدث مع موظف", "payload": "HUMAN_AGENT"},
        ]
    )
    session["step"] = "done"


# =========================
# Conversation Logic
# =========================
def process_ai_message(uid, text):
    session = user_sessions.get(uid, {"step": "ai_mode", "history": []})
    user_sessions[uid] = session
    step = session.get("step", "ai_mode")

    # إذا البوت موقوف مؤقتاً
    if not BOT_ENABLED:
        send_text(
            uid,
            "شكراً لتواصلكم مع شركة تهامة.\n"
            "خدمة الرد الآلي متوقفة مؤقتاً، وسيتم الرد عليكم من فريق خدمة العملاء في أقرب وقت."
        )
        send_email_human(uid, session.get("name"), session.get("phone"))
        return

    if step == "awaiting_name_ai":
        session["name"] = text.strip()
        session["step"] = "awaiting_phone_ai"
        send_text(uid, "ممتاز 👤\nيرجى كتابة رقم هاتفك:\nPlease enter your phone number:")
        return

    if step == "awaiting_phone_ai":
        session["phone"] = text.strip()
        gov = session.get("gov")
        product = session.get("product")

        if gov and product:
            session["step"] = "done"
            send_contact_card(uid, gov, product, session)
        else:
            session["step"] = "ai_mode"
            send_text(uid, "شكراً! كيف يمكنني مساعدتك؟\nThank you! How can I help you?")
        return

    if step == "human_mode":
        log.info(f"Human mode active for {uid}; ignoring bot reply.")
        return

    result = ask_ai(text, session)
    intent = result.get("intent", "unclear")
    product = result.get("product")
    governorate = result.get("governorate")
    reply = result.get("reply", "")
    missing = result.get("missing", "both")

    if intent == "human_agent":
        session["step"] = "human_mode"
        send_text(
            uid,
            "👨‍💼 تم تحويل طلبك إلى موظف خدمة العملاء.\n"
            "You'll be contacted by a customer service representative shortly.\n\n"
            "🕐 أوقات الدوام: 8 صباحاً — 5 مساءً"
        )
        send_email_human(uid, session.get("name"), session.get("phone"))
        return

    if intent == "find_rep" and product and governorate:
        session["product"] = product
        session["gov"] = governorate
        send_text(uid, reply)

        if not session.get("name"):
            session["step"] = "awaiting_name_ai"
            send_text(
                uid,
                "📝 يرجى كتابة اسمك الكامل لتسجيل طلبك:\n"
                "Please enter your full name to register your inquiry:"
            )
        else:
            send_contact_card(uid, governorate, product, session)
        return

    if intent == "find_rep" and product and not governorate:
        session["product"] = product
        session["step"] = "ai_mode"
        send_quick_replies(
            uid,
            f"{reply}\n\n📍 ما أقرب محافظة إليك؟\nNearest governorate?",
            [
                {"title": "صنعاء", "payload": "GOV_AI_صنعاء"},
                {"title": "الحديدة", "payload": "GOV_AI_الحديدة"},
                {"title": "عدن", "payload": "GOV_AI_عدن"},
                {"title": "تعز", "payload": "GOV_AI_تعز"},
                {"title": "المكلا", "payload": "GOV_AI_المكلا"},
            ]
        )
        return

    if intent == "find_rep" and governorate and not product:
        session["gov"] = governorate
        session["step"] = "ai_mode"
        send_quick_replies(
            uid,
            f"{reply}\n\n🛒 ما نوع المنتج؟\nWhat product?",
            [
                {"title": "قطع الغيار", "payload": "PRODUCT_AI_قطع الغيار"},
                {"title": "معدات", "payload": "PRODUCT_AI_معدات"},
                {"title": "مولدات", "payload": "PRODUCT_AI_مولدات"},
            ]
        )
        return

    send_text(uid, reply)
    if intent in ("greeting", "unclear", "general_question"):
        send_quick_replies(
            uid,
            "كيف يمكنني مساعدتك؟ | How can I help you?",
            [
                {"title": "قطع الغيار", "payload": "PRODUCT_قطع الغيار"},
                {"title": "معدات", "payload": "PRODUCT_معدات"},
                {"title": "مولدات", "payload": "PRODUCT_مولدات"},
                {"title": "موظف", "payload": "HUMAN_AGENT"},
            ]
        )


def handle_payload(uid, payload):
    session = user_sessions.get(uid, {})

    if payload in ("GET_STARTED", "RESTART"):
        user_sessions[uid] = {"step": "ai_mode", "history": []}
        send_text(
            uid,
            "👋 أهلاً وسهلاً!\n"
            "شركة تهامة — CAT معدات ومولدات وقطع غيار\n"
            "Tehama Co. — CAT Equipment, Generators & Spare Parts\n\n"
            "💬 اكتب سؤالك بحرية وسأساعدك."
        )
        return

    if payload == "HUMAN_AGENT":
        session["step"] = "human_mode"
        user_sessions[uid] = session
        send_text(
            uid,
            "👨‍💼 تم تحويلك إلى موظف خدمة العملاء.\n"
            "⏳ سيتواصل معك أحد موظفينا خلال وقت قصير.\n"
            "🕐 أوقات الدوام: 8 صباحاً — 5 مساءً"
        )
        send_email_human(uid, session.get("name"), session.get("phone"))
        return

    if payload.startswith("GOV_AI_"):
        gov = payload[len("GOV_AI_"):]
        session["gov"] = gov
        user_sessions[uid] = session
        product = session.get("product")

        if product:
            if not session.get("name"):
                session["step"] = "awaiting_name_ai"
                send_text(uid, "📝 يرجى كتابة اسمك الكامل:\nPlease enter your full name:")
            else:
                send_contact_card(uid, gov, product, session)
        else:
            send_quick_replies(
                uid,
                "ما نوع المنتج؟ | What product?",
                [
                    {"title": "قطع الغيار", "payload": "PRODUCT_AI_قطع الغيار"},
                    {"title": "معدات", "payload": "PRODUCT_AI_معدات"},
                    {"title": "مولدات", "payload": "PRODUCT_AI_مولدات"},
                ]
            )
        return

    if payload.startswith("PRODUCT_AI_"):
        product = payload[len("PRODUCT_AI_"):]
        session["product"] = product
        user_sessions[uid] = session
        gov = session.get("gov")

        if gov:
            if not session.get("name"):
                session["step"] = "awaiting_name_ai"
                send_text(uid, "📝 يرجى كتابة اسمك الكامل:\nPlease enter your full name:")
            else:
                send_contact_card(uid, gov, product, session)
        else:
            send_quick_replies(
                uid,
                "ما أقرب محافظة؟ | Nearest governorate?",
                [
                    {"title": "صنعاء", "payload": "GOV_AI_صنعاء"},
                    {"title": "الحديدة", "payload": "GOV_AI_الحديدة"},
                    {"title": "عدن", "payload": "GOV_AI_عدن"},
                    {"title": "تعز", "payload": "GOV_AI_تعز"},
                    {"title": "المكلا", "payload": "GOV_AI_المكلا"},
                ]
            )
        return

    if payload.startswith("PRODUCT_"):
        product = payload[len("PRODUCT_"):]
        session["product"] = product
        session["step"] = "ai_mode"
        user_sessions[uid] = session
        send_quick_replies(
            uid,
            f"ممتاز، استفسارك عن {product}.\nما أقرب محافظة؟",
            [
                {"title": "صنعاء", "payload": "GOV_AI_صنعاء"},
                {"title": "الحديدة", "payload": "GOV_AI_الحديدة"},
                {"title": "عدن", "payload": "GOV_AI_عدن"},
                {"title": "تعز", "payload": "GOV_AI_تعز"},
                {"title": "المكلا", "payload": "GOV_AI_المكلا"},
            ]
        )
        return


# =========================
# Routes
# =========================
@app.route("/webhook", methods=["GET"])
def verify():
    mode = request.args.get("hub.mode")
    token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if mode == "subscribe" and token == VERIFY_TOKEN:
        return challenge, 200
    return "Verification failed", 403


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json(silent=True) or {}
    log.info(f"Incoming webhook: {json.dumps(data, ensure_ascii=False)[:1500]}")

    try:
        for entry in data.get("entry", []):
            for event in entry.get("messaging", []):
                uid = event.get("sender", {}).get("id")
                if not uid:
                    continue

                if "postback" in event:
                    handle_payload(uid, event["postback"].get("payload", ""))

                elif "message" in event:
                    msg = event["message"]

                    if msg.get("is_echo"):
                        continue

                    if "quick_reply" in msg:
                        handle_payload(uid, msg["quick_reply"].get("payload", ""))

                    elif "text" in msg:
                        process_ai_message(uid, msg["text"])

    except Exception as e:
        log.exception(f"Webhook error: {e}")

    return "ok", 200


@app.route("/resume/<uid>", methods=["GET"])
def resume_bot(uid):
    if uid in user_sessions:
        user_sessions[uid] = {"step": "ai_mode", "history": []}
        send_text(uid, "✅ شكراً لتواصلك مع تهامة!\nيمكنك طرح أي سؤال آخر.")
        return "Bot resumed", 200
    return "Not found", 404


@app.route("/", methods=["GET"])
def health():
    return "Messenger AI Bot Running", 200
