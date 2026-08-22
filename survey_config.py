# -*- coding: utf-8 -*-
"""
कार्य-प्रगति वृत्त — फॉर्म विन्यास (यहीं से प्रश्न बदलें)
==========================================================
संरचना:
SECTIONS = [
  { "title": "खंड का शीर्षक", "items": [
      { "grp": "क.1",            # खंड अक्षर + क्रमांक (Excel/तालिका में भी यही दिखेगा)
        "label": "प्रश्न का पूरा वाक्य",
        "fields": [               # एक प्रश्न में कई उप-क्षेत्र हो सकते हैं
          { "id": "k1_cur", "label": "उप-क्षेत्र का नाम",
            "type": "number",    # number | text | textarea | radio
            "required": True, "placeholder": "0" },
        ] } ] }
"""

FORM_TITLE = "कार्य-प्रगति वृत्त — अगस्त, 2026"
FORM_DESC = "कृपया नीचे दिए गए सभी प्रश्नों के उत्तर भरें। सभी संख्याएँ अंकों में लिखें।"
INSTRUCTION = "आप इस फॉर्मेट को भरकर 31 अगस्त, 2026 तक भेजने का कष्ट करेंगे।"
THANK_YOU_MSG = "धन्यवाद! आपका कार्य-प्रगति विवरण सफलतापूर्वक दर्ज हो गया। 🙏"
SITE_TITLE = "कार्य-प्रगति वृत्त"

SECTIONS = [
    # ---------- प्रारंभिक जानकारी (बिना क्रमांक) ----------
    {
        "title": "प्रारंभिक जानकारी",
        "items": [
            {
                "grp": "",
                "label": "प्रांत का नाम",
                "fields": [
                    {"id": "prov", "label": "प्रांत का नाम", "type": "text",
                     "required": True, "placeholder": "जैसे: उत्तर प्रदेश, मध्य प्रदेश...",
                     "full": True},
                ],
            },
            {
                "grp": "",
                "label": "प्रांत अध्यक्ष/प्रमुख",
                "fields": [
                    {"id": "head", "label": "प्रांत अध्यक्ष / प्रमुख", "type": "text",
                     "required": True, "placeholder": "श्री / श्रीमती ......................",
                     "full": True},
                ],
            },
        ],
    },

    # ---------- क. संस्कार परिवार योजना का व्याप ----------
    {
        "title": "क. संस्कार परिवार योजना का व्याप",
        "items": [
            {
                "grp": "क.1",
                "label": "संस्कार परिवार युक्त जिलों की वर्तमान संख्या तथा सितंबर माह का विस्तार लक्ष्य",
                "fields": [
                    {"id": "k1_cur", "label": "जिलों की वर्तमान संख्या", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "k1_tgt", "label": "सितंबर माह का विस्तार लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "क.2",
                "label": "संस्कार परिवार युक्त प्रखण्डों की वर्तमान संख्या तथा सितंबर माह का लक्ष्य",
                "fields": [
                    {"id": "k2_cur", "label": "प्रखण्डों की वर्तमान संख्या", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "k2_tgt", "label": "सितंबर माह का लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "क.3",
                "label": "संस्कार परिवार योजना युक्त नगरों की संख्या",
                "fields": [
                    {"id": "k3", "label": "योजना युक्त नगरों की संख्या", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
        ],
    },

    # ---------- ख. संस्कार परिवारों का विवरण ----------
    {
        "title": "ख. संस्कार परिवारों का विवरण",
        "items": [
            {
                "grp": "ख.1",
                "label": "दैनिक संस्कार परिवारों की वर्तमान संख्या तथा सितंबर माह का विस्तार लक्ष्य",
                "fields": [
                    {"id": "kh1_cur", "label": "दैनिक परिवार — वर्तमान संख्या", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "kh1_tgt", "label": "दैनिक परिवार — सितंबर विस्तार लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ख.2",
                "label": "मासिक संस्कार परिवारों की वर्तमान संख्या तथा सितंबर माह का विस्तार लक्ष्य",
                "fields": [
                    {"id": "kh2_cur", "label": "मासिक परिवार — वर्तमान संख्या", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "kh2_tgt", "label": "मासिक परिवार — सितंबर विस्तार लक्ष्य", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
        ],
    },

    # ---------- ग. राष्ट्र रक्षा यज्ञ आयोजन समितियों का गठन ----------
    {
        "title": "ग. राष्ट्र रक्षा यज्ञ आयोजन समितियों का गठन",
        "items": [
            {
                "grp": "ग.1",
                "label": "प्रखण्डों/नगरों में आयोजन समिति के गठन की स्थिति",
                "fields": [
                    {"id": "g1_lak", "label": "कितने प्रखण्डों/नगरों में समिति बनाने का लक्ष्य था?", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "g1_done", "label": "कितने प्रखण्डों/नगरों की 15 सदस्यीय समिति बन गई?", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "g1_by", "label": "शेष प्रखण्डों/नगरों की समिति कब तक बन जाएगी?", "type": "text", "required": True, "placeholder": "जैसे: 15 सितंबर तक", "full": True},
                ],
            },
            {
                "grp": "ग.2",
                "label": "जिलों में आयोजन समिति के गठन की स्थिति",
                "fields": [
                    {"id": "g2_lak", "label": "कितने जिलों में समिति बनाने का लक्ष्य था?", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "g2_done", "label": "कितने जिलों की 25 सदस्यीय समिति बन गई?", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "g2_by", "label": "शेष जिलों की समिति कब तक बनेगी?", "type": "text", "required": True, "placeholder": "जैसे: 20 सितंबर तक", "full": True},
                ],
            },
            {
                "grp": "ग.3",
                "label": "51 सदस्यीय यज्ञ आयोजन समिति के गठन की स्थिति",
                "fields": [
                    {"id": "g3_formed", "label": "क्या आपके प्रांत की 51 सदस्यीय यज्ञ आयोजन समिति बन गई?", "type": "radio",
                     "required": True, "options": ["हाँ", "नहीं"], "full": True},
                    # सशर्त फ़ील्ड: "नहीं" चुनने पर ही दिखेगा और ज़रूरी होगा
                    {"id": "g3_by", "label": "यदि नहीं तो समिति कब तक बन जाएगी?", "type": "text",
                     "required": True, "placeholder": "जैसे: 30 सितंबर तक", "full": True,
                     "show_if": "g3_formed", "show_if_value": "नहीं"},
                ],
            },
        ],
    },

    # ---------- घ. राष्ट्र रक्षा यज्ञ अनुष्ठान आयोजन कार्यक्रम ----------
    {
        "title": "घ. राष्ट्र रक्षा यज्ञ अनुष्ठान आयोजन कार्यक्रम",
        "items": [
            {
                "grp": "घ.1",
                "label": "सितंबर, 2026 में राष्ट्र रक्षा यज्ञ आयोजन",
                "fields": [
                    {"id": "gh1", "label": "कितने नगरों/प्रखण्डों में राष्ट्र रक्षा यज्ञ सम्पन्न होंगे?", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "घ.2",
                "label": "जिला समितियों द्वारा यज्ञ अनुष्ठान (जिला केन्द्र)",
                "fields": [
                    {"id": "gh2", "label": "कितने जिला केन्द्रों पर यज्ञ अनुष्ठान होना तय है?", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "घ.3",
                "label": "प्रांत स्तरीय यज्ञ अनुष्ठान की तिथि",
                "fields": [
                    {"id": "gh3_prov", "label": "किस प्रांत में?", "type": "text", "required": True, "placeholder": "प्रांत का नाम"},
                    {"id": "gh3_month", "label": "किस माह में?", "type": "text", "required": True, "placeholder": "जैसे: अक्टूबर"},
                    {"id": "gh3_when", "label": "कब (तिथि)?", "type": "text", "required": True, "placeholder": "जैसे: 10 अक्टूबर, 2026"},
                ],
            },
        ],
    },

    # ---------- ड. अगस्त, 2026 के मासिक परिवार मिलन का विवरण ----------
    {
        "title": "ड. अगस्त, 2026 के मासिक परिवार मिलन का विवरण",
        "items": [
            {
                "grp": "ड.1",
                "label": "मासिक परिवार मिलन — कुल स्थान",
                "fields": [
                    {"id": "d1_exp", "label": "कुल कितने स्थानों पर मिलन अपेक्षित था?", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "d1_done", "label": "कुल कितने स्थानों पर सम्पन्न हुआ?", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ड.2",
                "label": "मिलन में उपस्थिति (व्यक्ति)",
                "fields": [
                    {"id": "d2_exp", "label": "कुल अपेक्षित संख्या", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "d2_pres", "label": "उपस्थित संख्या", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
            {
                "grp": "ड.3",
                "label": "मिलन में उपस्थिति (परिवार)",
                "fields": [
                    {"id": "d3_exp", "label": "कुल अपेक्षित परिवार संख्या", "type": "number", "required": True, "placeholder": "0"},
                    {"id": "d3_pres", "label": "उपस्थित परिवार संख्या", "type": "number", "required": True, "placeholder": "0"},
                ],
            },
        ],
    },
]

# ============ एडमिन सेटिंग्स ============
ADMIN_PASSWORD = "admin123"  # बदलना न भूलें! (उत्पादन में परिवेश-चर ADMIN_PASSWORD का उपयोग होगा)
