# -*- coding: utf-8 -*-
"""
कार्य-वृत्त — फॉर्म विन्यास (यहीं से प्रश्न बदलें)
====================================================
संरचना:
SECTIONS = [
  { "part": "प्रथम भाग",            # दो भाग: प्रथम भाग / द्वितीय भाग
    "title": "खंड का शीर्षक",
    "items": [
      { "grp": "क.1",
        "label": "प्रश्न का पूरा वाक्य",
        "fields": [
          { "id": "p1_kha1_cur", "label": "पूरा नाम (Excel/PDF हेडर)",
            "short": "संक्षिप्त नाम (डैशबोर्ड)", "type": "number",
            "required": True, "placeholder": "0" },
        ] } ] }
]

विशेष फ़ील्ड:
- "calc": ["id1", "id2"]  → योग अपने आप (id1 + id2), readonly
- "show_if"/"show_if_value" → सशर्त फ़ील्ड (जैसे "नहीं" चुनने पर ही दिखे)
- "type": "select" + "options" → dropdown सूची
"""

import calendar
from datetime import datetime, timezone, timedelta

# ============ माह / वर्ष — हर माह अपने आप बदलेगा (IST) ============
# ⚠️ HINDI_MONTHS को कभी न बदलें — अंदरूनी कैलेंडर इसी से चलता है।
HINDI_MONTHS = [
    "जनवरी", "फरवरी", "मार्च", "अप्रैल", "मई", "जून",
    "जुलाई", "अगस्त", "सितंबर", "अक्टूबर", "नवंबर", "दिसंबर",
]

# रिपोर्टिंग माह dropdown — यही form में दिखेगा (सूची यहाँ बदलें)
REPORTING_MONTHS = [
    "अगस्त, 2026", "सितंबर, 2026", "अक्टूबर, 2026", "नवंबर, 2026",
    "दिसंबर, 2026", "जनवरी, 2027", "फरवरी, 2027", "मार्च, 2027",
]


def _now_ist():
    """भारत का वर्तमान समय (Asia/Kolkata)।"""
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Asia/Kolkata"))
    except Exception:
        return datetime.now(timezone(timedelta(hours=5, minutes=30)))


def current_month_year():
    """('सितंबर', 2026) जैसा — हर request पर ताज़ा।"""
    n = _now_ist()
    return HINDI_MONTHS[n.month - 1], n.year


def get_month_label():
    """'सितंबर, 2026' — टाइटल, Excel, PDF में यही दिखेगा।"""
    m, y = current_month_year()
    return f"{m}, {y}"


def get_form_title():
    return f"कार्य-वृत्त — {get_month_label()}"


def get_instruction():
    """जमा करने की अंतिम तिथि = चालू माह का अंतिम दिन (अपने आप)।"""
    n = _now_ist()
    last_day = calendar.monthrange(n.year, n.month)[1]
    m, y = HINDI_MONTHS[n.month - 1], n.year
    return f"आप इस फॉर्मेट को भरकर {last_day} {m}, {y} तक भेजने का कष्ट करेंगे।"


# ============ संगठन की जानकारी (हेडर में दिखती है) ============
ORG_NAME = "वनवासी रक्षा परिवार फॉउन्डेशन"
ORG_TAGLINE = "वन एवं नगरीय वंचित समाज के उत्थान को समर्पित"
ORG_SUB = "राष्ट्र-रक्षा-यज्ञ अनुष्ठान"
# लोगो: static/logo.png फ़ाइल रखें (अगर फ़ाइल न हो तो केवल नाम दिखेगा)

# ============ दो भागों की जानकारी ============
PART_ORDER = ["प्रथम भाग", "द्वितीय भाग"]
PART_SUBTITLES = {
    "प्रथम भाग": "संस्कार परिवार योजना एवं राष्ट्र रक्षा यज्ञ",
    "द्वितीय भाग": "प्रतिभा विकास केन्द्र संबंधी विवरण",
}
PART_SHORT = {"प्रथम भाग": "1", "द्वितीय भाग": "2"}  # डैशबोर्ड टैग हेतु

# प्रांत dropdown में अंतिम विकल्प — चुनने पर नीचे textbox खुलता है
OTHER_PROV_LABEL = "अन्य (नीचे लिखें)"

# ============ अपेक्षित प्रांतों की आधिकारिक सूची (उत्तर-सूची ट्रैकर) ============
EXPECTED_PROVINCES = [
    "उ. असम",
    "द. असम",
    "उ. बंग",
    "द. बंग",
    "मं. बंग",
    "पू. ओडिशा",
    "प. ओडिशा",
    "झारखण्ड",
    "द. बिहार",
    "छत्तीसगढ़",
    "ब्रज",
    "मालवा",
    "चित्तौड़",
    "जयपुर",
    "द. गुजरात",
    "दिल्ली उत्तरी संभाग",
    "दिल्ली दक्षिणी संभाग",
    "दिल्ली पूर्वी संभाग",
    "उ. तमिलनाडु",
    "द. तमिलनाडु",
]

FORM_DESC = "कृपया नीचे दिए गए सभी प्रश्नों के उत्तर भरें। सभी संख्याएँ अंकों में लिखें।"
THANK_YOU_MSG = "धन्यवाद! आपका कार्य-वृत्त विवरण सफलतापूर्वक दर्ज हो गया। 🙏"
SITE_TITLE = "कार्य-वृत्त"

SECTIONS = [
    # ==================== प्रथम भाग ====================
    {
        "part": "प्रथम भाग",
        "title": "क. परिचयात्मक विवरण",
        "items": [
            {
                "grp": "",
                "label": "रिपोर्टिंग माह",
                "fields": [
                    {"id": "p1_month", "label": "रिपोर्टिंग माह (किस माह का विवरण भेज रहे हैं)",
                     "short": "रिपोर्ट माह", "type": "select", "options": REPORTING_MONTHS,
                     "default": "__current__", "required": True, "full": True},
                ],
            },
            {
                "grp": "",
                "label": "प्रांत का नाम",
                "fields": [
                    {"id": "p1_prov", "label": "प्रांत का नाम", "short": "प्रांत",
                     "type": "select", "options": EXPECTED_PROVINCES + [OTHER_PROV_LABEL],
                     "required": True, "full": True},
                    {"id": "p1_prov_other", "label": "अपने प्रांत का नाम लिखें",
                     "short": "अन्य प्रांत", "type": "text",
                     "required": True, "placeholder": "प्रांत का नाम...",
                     "full": True, "show_if": "p1_prov", "show_if_value": OTHER_PROV_LABEL},
                ],
            },
            {
                "grp": "",
                "label": "प्रांत अध्यक्ष/प्रमुख का नाम",
                "fields": [
                    {"id": "p1_head", "label": "प्रांत अध्यक्ष / प्रमुख का नाम", "short": "अध्यक्ष/प्रमुख",
                     "type": "text", "required": True,
                     "placeholder": "श्री / श्रीमती ......................",
                     "full": True},
                ],
            },
            {
                "grp": "",
                "label": "महामंत्री का नाम",
                "fields": [
                    {"id": "p1_secy", "label": "महामंत्री का नाम", "short": "महामंत्री",
                     "type": "text", "required": True,
                     "placeholder": "श्री / श्रीमती ......................",
                     "full": True},
                ],
            },
        ],
    },

    {
        "part": "प्रथम भाग",
        "title": "ख. कार्य का व्याप एवं विस्तार (प्रतिभा विकास केन्द्र योजना सहित)",
        "items": [
            {
                "grp": "ख.1",
                "label": "संस्कार परिवार युक्त जिले",
                "fields": [
                    {"id": "p1_kha1_cur", "label": "संस्कार परिवार युक्त जिलों की संख्या",
                     "short": "जिले — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_kha1_tgt", "label": "जिलों हेतु आगामी माह का नवीन लक्ष्य",
                     "short": "जिले — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ख.2",
                "label": "संस्कार परिवार युक्त प्रखण्ड",
                "fields": [
                    {"id": "p1_kha2_cur", "label": "संस्कार परिवार युक्त प्रखण्डों की संख्या",
                     "short": "प्रखण्ड — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_kha2_tgt", "label": "प्रखण्डों हेतु आगामी माह का नवीन लक्ष्य",
                     "short": "प्रखण्ड — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ख.3",
                "label": "संस्कार परिवार युक्त नगर",
                "fields": [
                    {"id": "p1_kha3_cur", "label": "संस्कार परिवार युक्त नगरों की संख्या",
                     "short": "नगर — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_kha3_tgt", "label": "नगरों हेतु आगामी माह का नवीन लक्ष्य",
                     "short": "नगर — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ख.4",
                "label": "संस्कार परिवार युक्त ग्राम",
                "fields": [
                    {"id": "p1_kha4_cur", "label": "संस्कार परिवार युक्त ग्रामों की संख्या",
                     "short": "ग्राम — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_kha4_tgt", "label": "ग्रामों हेतु आगामी माह का नवीन लक्ष्य",
                     "short": "ग्राम — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
        ],
    },

    {
        "part": "प्रथम भाग",
        "title": "ग. संस्कार परिवार योजना विवरण (प्रतिभा विकास केन्द्र योजना सहित)",
        "items": [
            {
                "grp": "ग.1",
                "label": "दैनिक संस्कार परिवार विवरण",
                "fields": [
                    {"id": "p1_g1_cur", "label": "दैनिक परिवार — वर्तमान संख्या",
                     "short": "दैनिक — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_g1_tgt", "label": "दैनिक परिवार — आगामी नवीन लक्ष्य",
                     "short": "दैनिक — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_g1_tot", "label": "दैनिक परिवार — योग (स्वतः)",
                     "short": "दैनिक — योग", "type": "number", "required": False,
                     "calc": ["p1_g1_cur", "p1_g1_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ग.2",
                "label": "मासिक संस्कार परिवार विवरण",
                "fields": [
                    {"id": "p1_g2_cur", "label": "मासिक परिवार — वर्तमान संख्या",
                     "short": "मासिक — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_g2_tgt", "label": "मासिक परिवार — आगामी नवीन लक्ष्य",
                     "short": "मासिक — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_g2_tot", "label": "मासिक परिवार — योग (स्वतः)",
                     "short": "मासिक — योग", "type": "number", "required": False,
                     "calc": ["p1_g2_cur", "p1_g2_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
        ],
    },

    {
        "part": "प्रथम भाग",
        "title": "घ. राष्ट्र रक्षा यज्ञ आयोजन समितियों का गठन की स्थिति का विवरण (प्रतिभा विकास केन्द्र योजना सहित)",
        "items": [
            {
                "grp": "घ.1",
                "label": "प्रखण्डों में समिति गठन की स्थिति",
                "fields": [
                    {"id": "p1_gh1_lak", "label": "कितने प्रखण्डों में गठन का लक्ष्य था?",
                     "short": "प्रखण्ड लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_gh1_done", "label": "कितने प्रखण्डों में गठन हुआ?",
                     "short": "प्रखण्ड गठित", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_gh1_by", "label": "शेष प्रखण्डों में कब तक गठन होगा?",
                     "short": "शेष कब तक", "type": "text", "required": True,
                     "placeholder": "जैसे: 15 अक्टूबर तक", "full": True},
                ],
            },
            {
                "grp": "घ.2",
                "label": "नगरों में समिति गठन की स्थिति",
                "fields": [
                    {"id": "p1_gh2_lak", "label": "कितने नगरों में गठन का लक्ष्य था?",
                     "short": "नगर लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_gh2_done", "label": "कितने नगरों में गठन हुआ?",
                     "short": "नगर गठित", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_gh2_by", "label": "शेष नगरों में गठन कब तक होगा?",
                     "short": "शेष कब तक", "type": "text", "required": True,
                     "placeholder": "जैसे: 15 अक्टूबर तक", "full": True},
                ],
            },
            {
                "grp": "घ.3",
                "label": "जिला समिति गठन की स्थिति",
                "fields": [
                    {"id": "p1_gh3_lak", "label": "कितने जिलों में गठन का लक्ष्य था?",
                     "short": "जिले लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_gh3_done", "label": "कितने जिलों की समिति गठित हुई?",
                     "short": "जिले गठित", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_gh3_by", "label": "शेष जिलों की समिति कब तक गठित होगी?",
                     "short": "शेष कब तक", "type": "text", "required": True,
                     "placeholder": "जैसे: 20 अक्टूबर तक", "full": True},
                ],
            },
            {
                "grp": "घ.4",
                "label": "प्रांत की राष्ट्र रक्षा यज्ञ आयोजन समिति",
                "fields": [
                    {"id": "p1_gh4_formed", "label": "क्या प्रांत की राष्ट्र रक्षा यज्ञ आयोजन समिति का गठन हुआ?",
                     "short": "समिति गठित?", "type": "radio", "required": True,
                     "options": ["हाँ", "नहीं"], "full": True},
                    {"id": "p1_gh4_mem", "label": "यदि हाँ, तो सदस्य संख्या लिखें",
                     "short": "सदस्य संख्या", "type": "number", "required": True,
                     "placeholder": "0", "full": True,
                     "show_if": "p1_gh4_formed", "show_if_value": "हाँ"},
                    {"id": "p1_gh4_by", "label": "यदि नहीं, तो कब तक गठित होगी? (दिनांक लिखें)",
                     "short": "नहीं तो कब?", "type": "text", "required": True,
                     "placeholder": "जैसे: 30 अक्टूबर तक", "full": True,
                     "show_if": "p1_gh4_formed", "show_if_value": "नहीं"},
                ],
            },
        ],
    },

    {
        "part": "प्रथम भाग",
        "title": "ड. राष्ट्र रक्षा यज्ञ अनुष्ठान आयोजन लक्ष्य, वर्तमान स्थिति एवं शेष (प्रतिभा विकास केन्द्र की योजना सहित)",
        "items": [
            {
                "grp": "ड.1",
                "label": "समितियों द्वारा यज्ञ अनुष्ठान का कुल लक्ष्य",
                "fields": [
                    {"id": "p1_d1_tgt", "label": "विभिन्न स्तर की समितियों द्वारा यज्ञ अनुष्ठान का कुल लक्ष्य",
                     "short": "यज्ञ — कुल लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ड.2",
                "label": "अब तक सम्पन्न यज्ञ अनुष्ठान",
                "fields": [
                    {"id": "p1_d2_done", "label": "अब तक सम्पन्न यज्ञ अनुष्ठान संख्या",
                     "short": "यज्ञ — सम्पन्न", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ड.3",
                "label": "प्रतिभा विकास केन्द्रों द्वारा कुल लक्ष्य",
                "fields": [
                    {"id": "p1_d3_tgt", "label": "प्रतिभा विकास केन्द्रों द्वारा यज्ञ अनुष्ठान का कुल लक्ष्य",
                     "short": "केन्द्र — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ड.4",
                "label": "केन्द्रों द्वारा सम्पन्न यज्ञ संख्या",
                "fields": [
                    {"id": "p1_d4_done", "label": "केन्द्रों द्वारा अब तक सम्पन्न यज्ञ संख्या",
                     "short": "केन्द्र — सम्पन्न", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ड.5",
                "label": "कुल राष्ट्र रक्षा यज्ञ अनुष्ठान का लक्ष्य",
                "fields": [
                    {"id": "p1_d5_tot", "label": "कुल लक्ष्य (स्वतः: समिति + केन्द्र)",
                     "short": "कुल लक्ष्य", "type": "number", "required": False,
                     "calc": ["p1_d1_tgt", "p1_d3_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ड.6",
                "label": "अब तक कुल सम्पन्न यज्ञ संख्या",
                "fields": [
                    {"id": "p1_d6_tot", "label": "कुल सम्पन्न (स्वतः: समिति + केन्द्र)",
                     "short": "कुल सम्पन्न", "type": "number", "required": False,
                     "calc": ["p1_d2_done", "p1_d4_done"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ड.7",
                "label": "ग्रामों में यज्ञ अनुष्ठान का लक्ष्य",
                "fields": [
                    {"id": "p1_d7_vill", "label": "कुल कितने ग्रामों में यज्ञ अनुष्ठान का लक्ष्य है?",
                     "short": "ग्राम लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ड.8",
                "label": "नगरों में यज्ञ अनुष्ठान का लक्ष्य",
                "fields": [
                    {"id": "p1_d8_town", "label": "कुल कितने नगरों में यज्ञ अनुष्ठान का लक्ष्य है?",
                     "short": "नगर लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ड.9",
                "label": "आमंत्रण लक्ष्य — ग्राम व बस्तियाँ",
                "fields": [
                    {"id": "p1_d9_vill", "label": "कितने ग्रामों के लोगों को आमंत्रित करने का लक्ष्य है?",
                     "short": "आमंत्रण — ग्राम", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p1_d9_basti", "label": "कितनी बस्तियों के लोगों को आमंत्रित करने का लक्ष्य है?",
                     "short": "आमंत्रण — बस्ती", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
        ],
    },

    # ==================== द्वितीय भाग ====================
    {
        "part": "द्वितीय भाग",
        "title": "क. परिचयात्मक विवरण",
        "items": [
            {
                "grp": "",
                "label": "प्रांत का नाम",
                "fields": [
                    {"id": "p2_prov", "label": "प्रांत का नाम", "short": "प्रांत",
                     "type": "text", "required": True,
                     "placeholder": "प्रांत का नाम...",
                     "full": True},
                ],
            },
            {
                "grp": "",
                "label": "प्रांत संयोजक",
                "fields": [
                    {"id": "p2_coord", "label": "प्रांत संयोजक", "short": "संयोजक",
                     "type": "text", "required": True,
                     "placeholder": "श्री / श्रीमती ......................",
                     "full": True},
                ],
            },
            {
                "grp": "",
                "label": "सेवाव्रती संख्या",
                "fields": [
                    {"id": "p2_seva", "label": "प्रांत में कुल सेवाव्रती संख्या (प्रांत संयोजक तथा जिस वरिष्ठ कार्यकर्ता का केन्द्र आपके प्रांत में है, उन्हें जोड़कर लिखें)",
                     "short": "सेवाव्रती संख्या", "type": "number", "required": True,
                     "placeholder": "0", "full": True},
                ],
            },
        ],
    },

    {
        "part": "द्वितीय भाग",
        "title": "ख. प्रतिभा विकास केन्द्र का व्याप",
        "items": [
            {
                "grp": "ख.1",
                "label": "केन्द्र युक्त जिले — वर्तमान संख्या, नवीन लक्ष्य एवं योग",
                "fields": [
                    {"id": "p2_kh1_cur", "label": "केन्द्र युक्त जिले — वर्तमान संख्या",
                     "short": "जिले — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_kh1_tgt", "label": "केन्द्र युक्त जिले — नवीन लक्ष्य",
                     "short": "जिले — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_kh1_tot", "label": "केन्द्र युक्त जिले — योग (स्वतः)",
                     "short": "जिले — योग", "type": "number", "required": False,
                     "calc": ["p2_kh1_cur", "p2_kh1_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ख.2",
                "label": "केन्द्र युक्त प्रखण्ड — वर्तमान संख्या, नवीन लक्ष्य एवं योग",
                "fields": [
                    {"id": "p2_kh2_cur", "label": "केन्द्र युक्त प्रखण्ड — वर्तमान संख्या",
                     "short": "प्रखण्ड — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_kh2_tgt", "label": "केन्द्र युक्त प्रखण्ड — नवीन लक्ष्य",
                     "short": "प्रखण्ड — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_kh2_tot", "label": "केन्द्र युक्त प्रखण्ड — योग (स्वतः)",
                     "short": "प्रखण्ड — योग", "type": "number", "required": False,
                     "calc": ["p2_kh2_cur", "p2_kh2_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ख.3",
                "label": "केन्द्र युक्त नगर — वर्तमान संख्या, नवीन लक्ष्य एवं योग",
                "fields": [
                    {"id": "p2_kh3_cur", "label": "केन्द्र युक्त नगर — वर्तमान संख्या",
                     "short": "नगर — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_kh3_tgt", "label": "केन्द्र युक्त नगर — नवीन लक्ष्य",
                     "short": "नगर — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_kh3_tot", "label": "केन्द्र युक्त नगर — योग (स्वतः)",
                     "short": "नगर — योग", "type": "number", "required": False,
                     "calc": ["p2_kh3_cur", "p2_kh3_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ख.4",
                "label": "केन्द्र युक्त ग्राम — वर्तमान संख्या, नवीन लक्ष्य एवं योग",
                "fields": [
                    {"id": "p2_kh4_cur", "label": "केन्द्र युक्त ग्राम — वर्तमान संख्या",
                     "short": "ग्राम — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_kh4_tgt", "label": "केन्द्र युक्त ग्राम — नवीन लक्ष्य",
                     "short": "ग्राम — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_kh4_tot", "label": "केन्द्र युक्त ग्राम — योग (स्वतः)",
                     "short": "ग्राम — योग", "type": "number", "required": False,
                     "calc": ["p2_kh4_cur", "p2_kh4_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
        ],
    },

    {
        "part": "द्वितीय भाग",
        "title": "ग. प्रतिभा विकास केन्द्र का संख्यात्मक विवरण",
        "items": [
            {
                "grp": "ग.1",
                "label": "केन्द्र की वर्तमान संख्या, नवीन लक्ष्य एवं योग",
                "fields": [
                    {"id": "p2_g1_cur", "label": "केन्द्र — वर्तमान संख्या",
                     "short": "केन्द्र — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g1_tgt", "label": "केन्द्र — नवीन लक्ष्य",
                     "short": "केन्द्र — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g1_tot", "label": "केन्द्र — योग (स्वतः)",
                     "short": "केन्द्र — योग", "type": "number", "required": False,
                     "calc": ["p2_g1_cur", "p2_g1_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ग.2",
                "label": "छात्रों की वर्तमान संख्या, नवीन लक्ष्य एवं योग",
                "fields": [
                    {"id": "p2_g2_cur", "label": "छात्र — वर्तमान संख्या",
                     "short": "छात्र — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g2_tgt", "label": "छात्र — नवीन लक्ष्य",
                     "short": "छात्र — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g2_tot", "label": "छात्र — योग (स्वतः)",
                     "short": "छात्र — योग", "type": "number", "required": False,
                     "calc": ["p2_g2_cur", "p2_g2_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ग.3",
                "label": "अन्य छात्रों की वर्तमान संख्या, नवीन लक्ष्य एवं योग",
                "fields": [
                    {"id": "p2_g3_cur", "label": "अन्य छात्र — वर्तमान संख्या",
                     "short": "अन्य छात्र — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g3_tgt", "label": "अन्य छात्र — नवीन लक्ष्य",
                     "short": "अन्य छात्र — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g3_tot", "label": "अन्य छात्र — योग (स्वतः)",
                     "short": "अन्य छात्र — योग", "type": "number", "required": False,
                     "calc": ["p2_g3_cur", "p2_g3_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ग.4",
                "label": "अभिभावकों की वर्तमान संख्या, नवीन लक्ष्य एवं योग",
                "fields": [
                    {"id": "p2_g4_cur", "label": "अभिभावक — वर्तमान संख्या",
                     "short": "अभिभावक — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g4_tgt", "label": "अभिभावक — नवीन लक्ष्य",
                     "short": "अभिभावक — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g4_tot", "label": "अभिभावक — योग (स्वतः)",
                     "short": "अभिभावक — योग", "type": "number", "required": False,
                     "calc": ["p2_g4_cur", "p2_g4_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ग.5",
                "label": "अन्य परिवारों की वर्तमान संख्या, नवीन लक्ष्य एवं योग",
                "fields": [
                    {"id": "p2_g5_cur", "label": "अन्य परिवार — वर्तमान संख्या",
                     "short": "अन्य परिवार — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g5_tgt", "label": "अन्य परिवार — नवीन लक्ष्य",
                     "short": "अन्य परिवार — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g5_tot", "label": "अन्य परिवार — योग (स्वतः)",
                     "short": "अन्य परिवार — योग", "type": "number", "required": False,
                     "calc": ["p2_g5_cur", "p2_g5_tgt"], "readonly": True, "placeholder": "स्वतः"},
                ],
            },
            {
                "grp": "ग.6",
                "label": "11 सदस्यीय केन्द्र संचालन समिति गठन की वर्तमान स्थिति",
                "fields": [
                    {"id": "p2_g6_cur", "label": "वर्तमान समिति संख्या",
                     "short": "समिति — वर्तमान", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g6_mem", "label": "सदस्यों की कुल संख्या",
                     "short": "कुल सदस्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_g6_tgt", "label": "शेष समितियों का गठन लक्ष्य",
                     "short": "शेष लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
        ],
    },

    {
        "part": "द्वितीय भाग",
        "title": "घ. मासिक परिवार मिलन का संख्यात्मक विवरण",
        "items": [
            {
                "grp": "घ.1",
                "label": "मासिक परिवार मिलन",
                "fields": [
                    {"id": "p2_d1_exp", "label": "मिलन — कुल अपेक्षित संख्या",
                     "short": "मिलन — अपेक्षित", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_d1_done", "label": "सम्पन्न केन्द्रों की संख्या",
                     "short": "मिलन — सम्पन्न", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "घ.2",
                "label": "छात्र उपस्थिति विवरण",
                "fields": [
                    {"id": "p2_d2_stu", "label": "उपस्थित छात्र संख्या",
                     "short": "छात्र उपस्थित", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_d2_stu_tgt", "label": "छात्र — आगामी माह का लक्ष्य",
                     "short": "छात्र — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_d2_ostu", "label": "उपस्थित अन्य छात्र संख्या",
                     "short": "अन्य छात्र उपस्थित", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_d2_ostu_tgt", "label": "अन्य छात्र — आगामी माह का लक्ष्य",
                     "short": "अन्य छात्र — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "घ.3",
                "label": "अभिभावक एवं अन्य परिवार उपस्थिति विवरण",
                "fields": [
                    {"id": "p2_d3_par", "label": "उपस्थित अभिभावक संख्या",
                     "short": "अभिभावक उपस्थित", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_d3_par_tgt", "label": "अभिभावक — आगामी माह का लक्ष्य",
                     "short": "अभिभावक — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_d3_ofam", "label": "उपस्थित अन्य परिवारों की संख्या",
                     "short": "अन्य परिवार उपस्थित", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_d3_ofam_tgt", "label": "अन्य परिवार — आगामी माह का लक्ष्य",
                     "short": "अन्य परिवार — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "घ.4",
                "label": "संचालन समिति उपस्थिति विवरण",
                "fields": [
                    {"id": "p2_d4_mem", "label": "उपस्थित सदस्य संख्या",
                     "short": "सदस्य उपस्थित", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "p2_d4_mem_tgt", "label": "सदस्य — आगामी माह का लक्ष्य",
                     "short": "सदस्य — लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
        ],
    },
]

# ============ एडमिन सेटिंग्स ============
ADMIN_PASSWORD = "admin123"  # बदलना न भूलें! (उत्पादन में परिवेश-चर ADMIN_PASSWORD का उपयोग होगा)
