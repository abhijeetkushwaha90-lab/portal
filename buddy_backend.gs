/* ====================================================================
   STUDY BUDDIES — AI CHATBOT BACKEND ADD-ON
   Paste this ENTIRE block at the END of your existing Code.gs file.
   Do not delete anything from your existing file.
   ==================================================================== */

/* ----------------------------------------------------------------
   1) BOT CONFIG SHEET SETUP
   ---------------------------------------------------------------- */
var BOT_CONFIG_HEADERS = ["BotID","Name","Subject","AvatarColor","AvatarInitials","SystemPrompt","Active"];
var BUDDY_CHATS_HEADERS = ["Timestamp","Email","BotID","Role","Message"];

var DEFAULT_BOTS = [
  { id:"aditi",  name:"Aditi Mam",  subject:"English",
    color:"#e879f9", initials:"AM",
    prompt:"Tum Aditi Mam ho, ek friendly aur experienced English teacher. Tum students ke English doubts (grammar, vocabulary, comprehension, writing skills) clear karti ho. Hindi-English mix (Hinglish) me simple examples ke saath samjhao. Hamesha encouraging aur patient tone rakho. Answer chhota aur clear rakho, zarurat ho to hi lamba likho." },
  { id:"neha",   name:"Neha Mam",   subject:"Mathematics",
    color:"#60a5fa", initials:"NM",
    prompt:"Tum Neha Mam ho, ek Mathematics teacher jo students ko step-by-step formula aur concepts samjhati ho. Hinglish me simple language use karo, har step clearly likho. Calculation mistakes na ho iska khayal rakho." },
  { id:"anaya",  name:"Anaya Mam",  subject:"General Studies",
    color:"#34d399", initials:"AnM",
    prompt:"Tum Anaya Mam ho, General Studies (GS) ki teacher — History, Polity, Economy, Science basics cover karti ho. Hinglish me clear aur factually accurate jawab do. Agar exact date/fact certain na ho to 'approximately' bolo, galat fact mat do." },
  { id:"rohan",  name:"Rohan Sir",  subject:"Reasoning",
    color:"#fbbf24", initials:"RS",
    prompt:"Tum Rohan Sir ho, Logical Reasoning aur Analytical Ability ke teacher. Puzzles, series, coding-decoding jaise topics step-by-step solve karke samjhao, Hinglish me." },
  { id:"priya",  name:"Priya Mam",  subject:"Current Affairs",
    color:"#f472b6", initials:"PM",
    prompt:"Tum Priya Mam ho, Current Affairs teacher. Recent events ke baare me tumhe exact real-time data nahi pata, isliye agar koi bahut recent event (last few months) puchhe to politely bolo ki tumhe latest update nahi pata, general concepts aur purane established facts explain karo." },
  { id:"vikram", name:"Vikram Sir", subject:"History",
    color:"#a78bfa", initials:"VS",
    prompt:"Tum Vikram Sir ho, History teacher (Ancient, Medieval, Modern Indian History). Dates aur events Hinglish me clearly samjhao, story-telling style use karo taaki yaad rahe." },
  { id:"sanya",  name:"Sanya Mam",  subject:"Geography",
    color:"#22d3ee", initials:"SM",
    prompt:"Tum Sanya Mam ho, Geography teacher (Physical, Indian, World Geography). Maps, climate, rivers jaise topics simple Hinglish me samjhao." },
  { id:"aarav",  name:"Aarav Sir",  subject:"Polity",
    color:"#fb923c", initials:"AS",
    prompt:"Tum Aarav Sir ho, Indian Polity aur Civics ke teacher. Constitution, articles, governance structure Hinglish me clearly explain karo, examples ke saath." },
  { id:"ishita", name:"Ishita Mam", subject:"Science",
    color:"#4ade80", initials:"IM",
    prompt:"Tum Ishita Mam ho, Science teacher (Physics, Chemistry, Biology basics). Concepts ko simple real-life examples se Hinglish me samjhao." },
  { id:"kabir",  name:"Kabir Sir",  subject:"SSB & Personality",
    color:"#f87171", initials:"KS",
    prompt:"Tum Kabir Sir ho, SSB interview aur Personality Development expert (Operation Baaz theme follow karte ho — OLQs, interview tips, confidence building). Students ko motivate karo aur practical SSB tips do, Hinglish me." }
];

function ensureBuddyTables_(ss) {
  ensureSheet_(ss, "BotConfig", BOT_CONFIG_HEADERS);
  ensureSheet_(ss, "BuddyChats", BUDDY_CHATS_HEADERS);
  var cfgSheet = ss.getSheetByName("BotConfig");
  if (cfgSheet.getLastRow() <= 1) {
    DEFAULT_BOTS.forEach(function(b) {
      cfgSheet.appendRow([b.id, b.name, b.subject, b.color, b.initials, b.prompt, "TRUE"]);
    });
  }
}

/* ----------------------------------------------------------------
   2) GET HANDLERS — bot list + chat history
   ---------------------------------------------------------------- */
function handleBuddyListAction_(e, ss) {
  var action = (e.parameter && e.parameter.action) ? e.parameter.action : "";
  ensureBuddyTables_(ss);

  if (action === "get_bots") {
    var cfgSheet = ss.getSheetByName("BotConfig");
    var d = cfgSheet.getDataRange().getValues();
    var out = [];
    for (var i=1; i<d.length; i++) {
      if (!d[i][0]) continue;
      if (d[i][6].toString().toUpperCase() !== "TRUE") continue;
      out.push({
        botId: d[i][0].toString(),
        name: d[i][1].toString(),
        subject: d[i][2].toString(),
        avatarColor: d[i][3].toString(),
        avatarInitials: d[i][4].toString()
      });
    }
    return json_(out);
  }

  if (action === "get_bot_history") {
    var email = (e.parameter.email || "").toLowerCase();
    var botId = e.parameter.botId || "";
    var chatSheet = ss.getSheetByName("BuddyChats");
    var d = chatSheet.getDataRange().getValues();
    var out = [];
    for (var i=1; i<d.length; i++) {
      if (!d[i][1]) continue;
      if (d[i][1].toString().toLowerCase() === email && d[i][2].toString() === botId) {
        out.push({ timestamp:d[i][0], role:d[i][3].toString(), message:d[i][4].toString() });
      }
    }
    return json_(out);
  }

  return null; // not a buddy action, let the rest of doGet handle it
}

/* ----------------------------------------------------------------
   3) POST HANDLER — send message to bot, save history, admin edit
   ---------------------------------------------------------------- */
function handleBuddyChatAction_(payload, ss) {
  ensureBuddyTables_(ss);

  if (payload.action === "chat_with_bot") {
    var email   = (payload.email || "anonymous").toLowerCase();
    var botId   = payload.botId || "";
    var message = (payload.message || "").toString().trim().substring(0, 1000);
    if (!message) return json_({ status:"error", message:"Khali message bhej nahi sakte." });

    var cfgSheet = ss.getSheetByName("BotConfig");
    var cfg = cfgSheet.getDataRange().getValues();
    var bot = null;
    for (var i=1; i<cfg.length; i++) {
      if (cfg[i][0].toString() === botId) { bot = { name:cfg[i][1], prompt:cfg[i][5] }; break; }
    }
    if (!bot) return json_({ status:"error", message:"Bot nahi mila." });

    var chatSheet = ss.getSheetByName("BuddyChats");
    chatSheet.appendRow([new Date(), email, botId, "user", message]);

    // last 10 messages for context
    var allChats = chatSheet.getDataRange().getValues();
    var history = [];
    for (var j=1; j<allChats.length; j++) {
      if (allChats[j][1].toString().toLowerCase() === email && allChats[j][2].toString() === botId) {
        history.push({ role: allChats[j][3].toString(), text: allChats[j][4].toString() });
      }
    }
    history = history.slice(-10);

    var replyText = callGeminiBot_(bot.prompt, history);

    chatSheet.appendRow([new Date(), email, botId, "bot", replyText]);

    return json_({ status:"success", reply: replyText });
  }

  /* ---- ADMIN: update bot config ---- */
  if (payload.action === "update_bot_config") {
    var cfgSheet = ss.getSheetByName("BotConfig");
    var d = cfgSheet.getDataRange().getValues();
    for (var i=1; i<d.length; i++) {
      if (d[i][0].toString() === payload.botId) {
        if (payload.name !== undefined) cfgSheet.getRange(i+1,2).setValue(payload.name);
        if (payload.subject !== undefined) cfgSheet.getRange(i+1,3).setValue(payload.subject);
        if (payload.avatarColor !== undefined) cfgSheet.getRange(i+1,4).setValue(payload.avatarColor);
        if (payload.avatarInitials !== undefined) cfgSheet.getRange(i+1,5).setValue(payload.avatarInitials);
        if (payload.systemPrompt !== undefined) cfgSheet.getRange(i+1,6).setValue(payload.systemPrompt);
        if (payload.active !== undefined) cfgSheet.getRange(i+1,7).setValue(payload.active ? "TRUE" : "FALSE");
        return json_({ status:"success", message:"Bot updated." });
      }
    }
    return json_({ status:"error", message:"Bot not found." });
  }

  /* ---- ADMIN: get full bot config (for editing) ---- */
  if (payload.action === "get_bot_config_admin") {
    var cfgSheet = ss.getSheetByName("BotConfig");
    var d = cfgSheet.getDataRange().getValues();
    var out = [];
    for (var i=1; i<d.length; i++) {
      if (!d[i][0]) continue;
      out.push({
        botId:d[i][0], name:d[i][1], subject:d[i][2], avatarColor:d[i][3],
        avatarInitials:d[i][4], systemPrompt:d[i][5], active:d[i][6]
      });
    }
    return json_(out);
  }

  return null; // not a buddy action
}

/* ----------------------------------------------------------------
   4) GEMINI API CALL (uses one shared key from Script Properties)
   ---------------------------------------------------------------- */
function callGeminiBot_(systemPrompt, history) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) return "Abhi bot setup ho raha hai, thodi der baad try karo. (API key missing)";

  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=" + apiKey;

  var contents = [];
  history.forEach(function(h) {
    contents.push({
      role: h.role === "bot" ? "model" : "user",
      parts: [{ text: h.text }]
    });
  });

  var body = {
    system_instruction: { parts: [{ text: systemPrompt }] },
    contents: contents,
    generationConfig: { temperature: 0.7, maxOutputTokens: 500 }
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  };

  try {
    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    var json = JSON.parse(resp.getContentText());

    if (code !== 200) {
      Logger.log("Gemini error: " + resp.getContentText());
      return "Sorry, abhi thoda busy hoon (server limit). Thodi der baad try karo.";
    }
    var text = json.candidates && json.candidates[0] && json.candidates[0].content
                && json.candidates[0].content.parts && json.candidates[0].content.parts[0]
                && json.candidates[0].content.parts[0].text;
    return text ? text.trim() : "Sorry, jawab nahi mil paya. Phir se try karo.";
  } catch (err) {
    Logger.log("Gemini exception: " + err);
    return "Sorry, kuch problem ho gayi. Thodi der baad try karo.";
  }
}
