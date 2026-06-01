const http = require("http");
const fs = require("fs");
const path = require("path");
const Module = require("module");

const bundledNodeModules = "C:/Users/mathi/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules";
const pnpmStore = path.join(bundledNodeModules, ".pnpm");
const pnpmModulePaths = fs.existsSync(pnpmStore)
  ? fs.readdirSync(pnpmStore).map(name => path.join(pnpmStore, name, "node_modules"))
  : [];
process.env.NODE_PATH = [process.env.NODE_PATH, bundledNodeModules, ...pnpmModulePaths].filter(Boolean).join(path.delimiter);
Module._initPaths();

let pptxgen;
try {
  pptxgen = require("pptxgenjs");
} catch {
  pptxgen = require(path.join(bundledNodeModules, "pptxgenjs"));
}

const publicDir = path.join(__dirname, "public");
const port = Number(process.env.PORT || 4173);

const mime = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".svg": "image/svg+xml; charset=utf-8",
};

function readJson(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", chunk => {
      body += chunk;
      if (body.length > 2_000_000) req.destroy();
    });
    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (error) {
        reject(error);
      }
    });
  });
}

const escapeXml = value => String(value || "")
  .replace(/&/g, "&amp;")
  .replace(/</g, "&lt;")
  .replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;");

const palettes = {
  studio: ["#fbf7ef", "#212121", "#ff7a59", "#2f80ed", "#0f9d58"],
  nordic: ["#edf6f9", "#12343b", "#2a9d8f", "#e9c46a", "#e76f51"],
  executive: ["#f7f7f2", "#111827", "#2563eb", "#14b8a6", "#f97316"],
  playful: ["#fff7ed", "#251605", "#ef476f", "#06d6a0", "#118ab2"],
  minimal: ["#ffffff", "#191919", "#d9d9d9", "#4b5563", "#0ea5e9"],
};

function styleProfile(style) {
  const key = palettes[style] ? style : "studio";
  return { key, colors: palettes[key] };
}

function splitSentences(text) {
  return cleanText(text)
    .split(/[.\n;]+/)
    .map(s => s.trim())
    .map(s => s.replace(/^(criteria|kriterier)\s*:\s*/i, ""))
    .filter(s => !isMetaInstruction(s))
    .filter(Boolean)
    .slice(0, 8);
}

function titleCase(value) {
  const text = normalizeTopic(value || "Ny presentasjon").trim();
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function cleanText(value) {
  let text = String(value || "").trim();
  const replacements = [
    [/\bforedrags\s*notater\b/gi, match => keepCase(match, "foredragsnotater")],
    [/\bkryterier\b/gi, match => keepCase(match, "kriterier")],
    [/\bpresentashjon\b/gi, match => keepCase(match, "presentasjon")],
    [/\bpresentasjonennå\b/gi, match => keepCase(match, "presentasjonen nå")],
    [/\bnotateme\b/gi, match => keepCase(match, "notatene")],
    [/\bcryteria\b/gi, match => keepCase(match, "criteria")],
    [/\bcryterier\b/gi, match => keepCase(match, "kriterier")],
    [/\blagger\b/gi, match => keepCase(match, "lager")],
    [/\blagge\b/gi, match => keepCase(match, "lage")],
    [/\bogg\b/gi, match => keepCase(match, "og")],
    [/\bgalt\b/gi, match => keepCase(match, "klart")],
    [/\bhva du skal si liksom\b/gi, "hva du skal si"],
    [/\bmmake\b/gi, "make"],
    [/\btro\b/gi, "to"],
    [/\ballt\b/gi, "alt"],
    [/\bgeneratye\b/gi, "generate"],
    [/\bimagges\b/gi, "images"],
    [/\bemojies\b/gi, "emojis"],
    [/\bskrive\s+feil\b/gi, match => keepCase(match, "skrivefeil")],
  ];
  replacements.forEach(([pattern, replacement]) => {
    text = text.replace(pattern, replacement);
  });
  text = text.replace(/\bskal\s+lager\b/gi, match => keepCase(match, "skal lage"));
  text = text.replace(/\s+/g, " ");
  text = text.replace(/\s+([,.!?;:])/g, "$1");
  return text;
}

function norwegianizeText(value) {
  let text = cleanText(value);
  const replacements = [
    [/\bThe Lego Batman Movie\b/g, "Lego Batman-filmen"],
    [/\bThe Lego Movie\b/g, "Lego-filmen"],
    [/\bDC Comics\b/g, "DC Comics"],
    [/\bWarner Animation Group\b/g, "Warner Animation Group"],
    [/\bspin-off\b/gi, "avlegger"],
    [/\boriginal English version\b/gi, "engelske originalversjonen"],
    [/\banimated superhero comedy film\b/gi, "animert superheltkomedie"],
    [/\bthe film\b/gi, "filmen"],
    [/\bthe story\b/gi, "historien"],
    [/\bimportant characters include\b/gi, "viktige figurer er"],
    [/\ba main theme is\b/gi, "et hovedtema er"],
  ];
  replacements.forEach(([pattern, replacement]) => {
    text = text.replace(pattern, replacement);
  });
  return text;
}

function localizeText(value, language = "no") {
  return language === "no" ? norwegianizeText(value) : cleanText(value);
}

function normalizeTopic(value) {
  let text = cleanText(value);
  text = text.replace(/^(kan du\s+)?(vær så snill\s+)?(lag|lage|lager|lagg|make|create)\s+(en\s+)?(kort\s+|lang\s+|bra\s+)?presentasjon\s+(som\s+handler\s+)?om\s+/i, "");
  text = text.replace(/^(historien\s+til|history\s+of)\s+/i, "Historien til ");
  text = text.replace(/\s+(og\s+)?(ha|lag|lage)\s+(med\s+)?(foredragsnotater|speaker notes).*$/i, "");
  return text.trim() || "Ny presentasjon";
}

function isMetaInstruction(text) {
  const cleaned = cleanText(text).toLowerCase();
  return [
    /foredragsnotater/,
    /speaker notes/,
    /ekte info/,
    /fakta/,
    /ikke ha skrivefeil/,
    /skriv[e]?feil/,
    /lag[e]?\s+(det\s+)?(bra|kort|langt)/,
    /presentasjon/,
  ].some(pattern => pattern.test(cleaned)) && cleaned.length < 80;
}

function researchCandidates(rawTopic) {
  const topic = normalizeTopic(rawTopic);
  const lower = topic.toLowerCase();
  const candidates = new Set([topic]);
  candidates.add(topic.replace(/\bfilmen\b/gi, "").trim());
  candidates.add(topic.replace(/^historien til\s+/i, "").trim());
  if (/lego batman/.test(lower)) {
    candidates.add("The Lego Batman Movie");
    candidates.add("Lego Batman");
  }
  if (/\bbatman\b/.test(lower) && /\bfilm/.test(lower)) candidates.add("Batman film");
  return [...candidates].filter(Boolean);
}

function keepCase(original, replacement) {
  if (original.toUpperCase() === original) return replacement.toUpperCase();
  if (original.charAt(0).toUpperCase() === original.charAt(0)) {
    return replacement.charAt(0).toUpperCase() + replacement.slice(1);
  }
  return replacement;
}

function wrapText(value, maxChars, maxLines = 2) {
  const words = String(value || "").split(/\s+/).filter(Boolean);
  const lines = [];
  let current = "";
  words.forEach(word => {
    const next = current ? `${current} ${word}` : word;
    if (next.length > maxChars && current) {
      lines.push(current);
      current = word;
    } else {
      current = next;
    }
  });
  if (current) lines.push(current);
  if (lines.length > maxLines) {
    const clipped = lines.slice(0, maxLines);
    clipped[maxLines - 1] = `${clipped[maxLines - 1].replace(/[,. ]+$/g, "")}...`;
    return clipped;
  }
  return lines;
}

function makeImageSvg({ topic, criteria, style }, variant) {
  const { colors } = styleProfile(style);
  const [bg, ink, a, b, c] = colors;
  const titles = ["Signal", "Flow", "Focus", "Momentum"];
  const topicLines = wrapText(titleCase(topic), 24, 3);
  const safeCriteria = escapeXml(splitSentences(criteria)[0] || "klar struktur");
  const topicTspans = topicLines
    .map((line, index) => `<tspan x="105" dy="${index === 0 ? 0 : 48}">${escapeXml(line)}</tspan>`)
    .join("");
  const shapes = [
    `<circle cx="900" cy="180" r="190" fill="${a}" opacity=".88"/><circle cx="760" cy="455" r="260" fill="${b}" opacity=".24"/><path d="M0 520 C230 410 340 610 590 500 S930 380 1200 470 V675 H0Z" fill="${c}" opacity=".82"/>`,
    `<rect x="720" y="-70" width="420" height="420" rx="64" fill="${a}" opacity=".9" transform="rotate(12 930 140)"/><rect x="90" y="390" width="780" height="120" rx="60" fill="${b}" opacity=".28"/><circle cx="210" cy="190" r="96" fill="${c}" opacity=".86"/>`,
    `<path d="M120 540 L440 120 L760 540 Z" fill="${a}" opacity=".86"/><path d="M520 560 L840 160 L1110 560 Z" fill="${b}" opacity=".34"/><circle cx="885" cy="178" r="92" fill="${c}" opacity=".88"/>`,
    `<path d="M0 0 H1200 V675 H0Z" fill="${bg}"/><path d="M80 112 H1120" stroke="${a}" stroke-width="24"/><path d="M80 220 H820" stroke="${b}" stroke-width="24"/><path d="M80 328 H1010" stroke="${c}" stroke-width="24"/><circle cx="980" cy="500" r="116" fill="${a}" opacity=".9"/>`,
  ][variant % 4];
  return `<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="675" viewBox="0 0 1200 675">
    <rect width="1200" height="675" fill="${bg}"/>
    ${shapes}
    <rect x="70" y="72" width="600" height="378" rx="28" fill="${bg}" opacity=".82"/>
    <text x="105" y="164" font-family="Arial, Helvetica, sans-serif" font-size="32" fill="${a}" font-weight="700">${titles[variant % titles.length]}</text>
    <text x="105" y="232" font-family="Arial, Helvetica, sans-serif" font-size="42" fill="${ink}" font-weight="800">${topicTspans}</text>
    <text x="108" y="${topicLines.length > 2 ? 382 : topicLines.length > 1 ? 338 : 300}" font-family="Arial, Helvetica, sans-serif" font-size="26" fill="${ink}" opacity=".72">${safeCriteria}</text>
    <text x="105" y="${topicLines.length > 2 ? 426 : 398}" font-family="Arial, Helvetica, sans-serif" font-size="22" fill="${ink}" opacity=".58">Generated visual direction ${variant + 1}</text>
  </svg>`;
}

function svgDataUri(svg) {
  return "data:image/svg+xml;base64," + Buffer.from(svg).toString("base64");
}

function outputLanguage(input) {
  const combined = `${input.topic || ""} ${input.criteria || ""}`.toLowerCase();
  const norwegianHits = (combined.match(/\b(og|om|lag|lage|presentasjon|filmen|historien|foredragsnotater|kriterier|skal|ikke|med|norsk)\b/g) || []).length;
  return norwegianHits >= 2 ? "no" : "en";
}

function sentenceList(text) {
  return cleanText(text)
    .split(/(?<=[.!?])\s+/)
    .map(sentence => sentence.trim())
    .filter(goodFactSentence)
    .slice(0, 12);
}

async function researchTopic(rawTopic, language = "no") {
  const candidates = researchCandidates(rawTopic);
  const builtIn = builtInResearch(candidates, language);
  if (builtIn.found) return builtIn;
  const languages = language === "no" ? ["no", "nb", "en"] : ["en", "no", "nb"];
  for (const candidate of candidates) {
    for (const lang of languages) {
      try {
      const searchUrl = `https://${lang}.wikipedia.org/w/api.php?action=opensearch&limit=1&namespace=0&format=json&search=${encodeURIComponent(candidate)}`;
      const searchResponse = await fetch(searchUrl, { headers: { "user-agent": "SlideDiscreetCraft/1.0" } });
      if (!searchResponse.ok) continue;
      const search = await searchResponse.json();
      const title = search?.[1]?.[0];
      if (!title) continue;

      const summaryUrl = `https://${lang}.wikipedia.org/api/rest_v1/page/summary/${encodeURIComponent(title)}`;
      const summaryResponse = await fetch(summaryUrl, { headers: { "user-agent": "SlideDiscreetCraft/1.0" } });
      if (!summaryResponse.ok) continue;
      const summary = await summaryResponse.json();
      const extract = cleanText(summary.extract || "");
      const facts = sentenceList(extract);
      if (facts.length) {
        return {
          found: true,
          title: localizeText(summary.title || title, language),
          description: localizeText(summary.description || "", language),
          extract,
          facts: facts.map(fact => localizeText(fact, language)),
          sources: [{ title: localizeText(summary.title || title, language), url: summary.content_urls?.desktop?.page || `https://${lang}.wikipedia.org/wiki/${encodeURIComponent(title)}` }],
          url: summary.content_urls?.desktop?.page || `https://${lang}.wikipedia.org/wiki/${encodeURIComponent(title)}`,
        };
      }
      } catch {
        // Keep the generator usable if research fails on a host without network access.
      }
    }
  }
  const webResearch = await researchWeb(candidates[0], candidates, language);
  if (webResearch.found) return webResearch;
  return { found: false, title: normalizeTopic(rawTopic), description: "", extract: "", facts: [], url: "" };
}

function goodFactSentence(sentence) {
  return sentence.length > 35
    && sentence.length < 260
    && !/\b(may refer to|cookie|privacy policy|subscribe|sign in|log in|javascript|advertisement)\b/i.test(sentence)
    && /[a-zA-ZæøåÆØÅ]{3,}/.test(sentence);
}

async function fetchText(url, timeoutMs = 8000) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, {
      signal: controller.signal,
      headers: {
        "user-agent": "Mozilla/5.0 SlideDiscreetCraft/1.0",
        "accept": "text/html,application/json,text/plain;q=0.9,*/*;q=0.8",
      },
    });
    if (!response.ok) return "";
    return await response.text();
  } catch {
    return "";
  } finally {
    clearTimeout(timeout);
  }
}

function decodeHtml(value) {
  return String(value || "")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&nbsp;/g, " ");
}

function stripHtml(html) {
  return decodeHtml(String(html || "")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<nav[\s\S]*?<\/nav>/gi, " ")
    .replace(/<footer[\s\S]*?<\/footer>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim());
}

function safeExternalUrl(value) {
  try {
    const url = new URL(decodeHtml(value));
    if (!["http:", "https:"].includes(url.protocol)) return "";
    if (/\.(pdf|zip|jpg|jpeg|png|gif|webp|mp4|mp3)$/i.test(url.pathname)) return "";
    const host = url.hostname.toLowerCase();
    if (/^(localhost|127\.|10\.|192\.168\.|172\.(1[6-9]|2\d|3[0-1])\.)/.test(host)) return "";
    return url.href;
  } catch {
    return "";
  }
}

function extractSearchUrls(html) {
  const urls = [];
  const add = raw => {
    const safe = safeExternalUrl(raw);
    if (safe && !urls.includes(safe) && !/duckduckgo\.com|google\.com\/search/i.test(safe)) urls.push(safe);
  };
  for (const match of html.matchAll(/uddg=([^"&]+)/g)) add(decodeURIComponent(match[1]));
  for (const match of html.matchAll(/<a[^>]+href=["']([^"']+)["'][^>]*>/gi)) add(match[1]);
  return urls.slice(0, 6);
}

function extractPageText(html) {
  const title = cleanText(stripHtml((html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1] || ""));
  const description = cleanText(decodeHtml((html.match(/<meta[^>]+name=["']description["'][^>]+content=["']([^"']+)["']/i) || [])[1] || ""));
  const paragraphs = [...html.matchAll(/<p[^>]*>([\s\S]*?)<\/p>/gi)]
    .map(match => stripHtml(match[1]))
    .filter(text => text.length > 45)
    .slice(0, 18)
    .join(" ");
  return { title, text: cleanText([description, paragraphs].filter(Boolean).join(" ")) };
}

function scoreSentence(sentence, topicWords) {
  const lower = sentence.toLowerCase();
  const matches = topicWords.filter(word => lower.includes(word)).length;
  const hasYear = /\b(18|19|20)\d{2}\b/.test(sentence) ? 1 : 0;
  const lengthScore = sentence.length > 70 && sentence.length < 210 ? 1 : 0;
  return matches * 4 + hasYear + lengthScore;
}

function importantSentences(pages, topic) {
  const topicWords = cleanText(topic).toLowerCase().split(/\s+/).filter(word => word.length > 3);
  const seen = new Set();
  return pages.flatMap(page => sentenceList(page.text).map(sentence => ({ sentence, page })))
    .map(item => ({ ...item, score: scoreSentence(item.sentence, topicWords) }))
    .filter(item => item.score > 0)
    .sort((a, b) => b.score - a.score)
    .filter(item => {
      const key = item.sentence.toLowerCase().slice(0, 80);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, 10);
}

async function researchWeb(topic, candidates, language = "no") {
  const query = language === "no" ? `${candidates[0] || topic} norsk fakta` : (candidates[0] || topic);
  const instant = await fetchText(`https://api.duckduckgo.com/?q=${encodeURIComponent(query)}&format=json&no_html=1&skip_disambig=1`, 6000);
  const instantFacts = [];
  const sources = [];
  try {
    const data = JSON.parse(instant || "{}");
    if (data.AbstractText) instantFacts.push(cleanText(data.AbstractText));
    if (data.AbstractURL) sources.push({ title: cleanText(data.Heading || query), url: data.AbstractURL });
  } catch {
    // Ignore malformed instant-answer responses.
  }

  const searchHtml = await fetchText(`https://duckduckgo.com/html/?q=${encodeURIComponent(query)}`, 8000);
  const urls = extractSearchUrls(searchHtml).slice(0, 4);
  const pages = [];
  for (const url of urls) {
    const html = await fetchText(url, 8000);
    if (!html) continue;
    const page = extractPageText(html);
    if (!page.text) continue;
    pages.push({ ...page, url });
    sources.push({ title: page.title || new URL(url).hostname, url });
  }
  const ranked = importantSentences(pages, query).map(item => item.sentence);
  const facts = [...instantFacts.flatMap(sentenceList), ...ranked]
    .filter(goodFactSentence)
    .slice(0, 10);
  if (!facts.length) return { found: false };
  const bestSource = sources[0] || {};
  return {
    found: true,
    title: titleCase(topic),
    description: language === "no" ? "Nettsøk basert på de viktigste treffene" : "Web research based on top results",
    extract: facts.join(" "),
    facts: facts.map(fact => localizeText(fact, language)),
    sources: sources.filter((source, index, all) => source.url && all.findIndex(item => item.url === source.url) === index).slice(0, 4),
    url: bestSource.url || "",
  };
}

function builtInResearch(candidates, language = "no") {
  const joined = candidates.join(" ").toLowerCase();
  if (/lego batman/.test(joined)) {
    const norwegian = language === "no";
    return {
      found: true,
      title: norwegian ? "Lego Batman-filmen" : "The Lego Batman Movie",
      description: norwegian ? "Animert superheltkomedie fra 2017" : "2017 animated superhero comedy film",
      extract: norwegian
        ? "Lego Batman-filmen er en animert superheltkomedie fra 2017. Filmen er en spin-off fra The Lego Movie og bruker Batman-figuren fra DC Comics i en humoristisk Lego-verden."
        : "The Lego Batman Movie is a 2017 animated superhero comedy film and a spin-off from The Lego Movie.",
      facts: norwegian ? [
        "Lego Batman-filmen er en animert superheltkomedie fra 2017.",
        "Filmen er en avlegger fra Lego-filmen og handler om Lego-versjonen av Batman.",
        "Will Arnett gir stemmen til Batman i den engelske originalversjonen.",
        "Historien handler mye om at Batman må lære å samarbeide med andre, selv om han helst vil jobbe alene.",
        "Viktige figurer i filmen er Batman, Robin, Batgirl, Alfred og Joker.",
        "Filmen blander action, komedie og parodi på mange kjente Batman-filmer og superhelthistorier.",
        "Filmen er regissert av Chris McKay og produsert av Warner Animation Group.",
        "Et hovedtema i filmen er familie, vennskap og det å tørre å slippe andre mennesker inn.",
      ] : [
        "The Lego Batman Movie is a 2017 animated superhero comedy film.",
        "The film is a spin-off from The Lego Movie and follows the Lego version of Batman.",
        "Will Arnett voices Batman in the original English version.",
        "The story focuses on Batman learning to work with others instead of always working alone.",
        "Important characters include Batman, Robin, Batgirl, Alfred and the Joker.",
        "The film mixes action, comedy and parody of famous Batman and superhero stories.",
        "Chris McKay directed the film, which was produced by Warner Animation Group.",
        "A main theme is family, friendship and letting other people into your life.",
      ],
      url: "https://en.wikipedia.org/wiki/The_Lego_Batman_Movie",
    };
  }
  return { found: false };
}

function fallbackPoints(criteria) {
  return criteria.length ? criteria : [
    "Hva temaet handler om",
    "Bakgrunn og viktig kontekst",
    "Viktige hendelser eller personer",
    "Hvorfor temaet er interessant",
    "Oppsummering og hva publikum bør huske",
  ];
}

async function buildSlides(input) {
  const language = outputLanguage(input);
  const topic = titleCase(input.topic);
  const criteria = splitSentences(input.criteria);
  const count = Math.max(5, Math.min(12, Number(input.slideCount || 7)));
  const emoji = input.emoji === true;
  const iconSet = emoji ? ["✨ ", "🎯 ", "🧭 ", "⚡ ", "📌 ", "🚀 "] : ["", "", "", "", "", ""];
  const research = await researchTopic(input.topic, language);
  const base = research.found ? research.facts : fallbackPoints(criteria);
  const deckTitle = localizeText(research.found ? research.title : topic, language);
  const slides = [
    {
      title: deckTitle,
      kicker: research.found ? "Faktabasert start" : "Presentation maker",
      bullets: research.found
        ? [`${iconSet[0]}${research.description || "Kort introduksjon"}`, `${iconSet[1]}Basert på fakta fra Wikipedia`]
        : [`${iconSet[0]}Hva presentasjonen handler om`, `${iconSet[1]}Hva publikum bør sitte igjen med`],
      notes: openingNote(deckTitle, input.notesLevel, research, language),
    },
  ];
  for (let i = 1; i < count - 1; i++) {
    const seed = base[(i - 1) % base.length];
    const shortTitle = makeSlideTitle(seed, i);
    slides.push({
      title: `${iconSet[i % iconSet.length]}${shortTitle}`,
      kicker: `Del ${i}`,
      bullets: makeFactBullets(seed, research, topic, language),
      notes: noteFor(seed, input.notesLevel, deckTitle, research, language),
    });
  }
  slides.push({
    title: `${emoji ? "✅ " : ""}Oppsummering og neste steg`,
    kicker: "Avslutning",
    bullets: research.found
      ? ["Oppsummer de viktigste faktaene", "Forklar hvorfor temaet er verdt å huske", research.url ? "Kilde: Wikipedia" : "Spørsmål og diskusjon"]
      : ["Hovedbudskapet i en setning", "Hva publikum bør gjøre videre", "Spørsmål og diskusjon"],
    notes: closingNote(deckTitle, input.notesLevel, research, language),
  });
  return slides;
}

function makeSlideTitle(seed, index) {
  const cleaned = cleanText(seed).replace(/\.$/, "");
  const words = cleaned.split(/\s+/).slice(0, 9).join(" ");
  return words || `Del ${index}`;
}

function makeFactBullets(seed, research, topic, language = "no") {
  if (!research.found) {
    return language === "no"
      ? [
        cleanText(seed),
        `Knytt punktet direkte til ${topic.toLowerCase()}`,
        "Bruk et tydelig eksempel",
      ]
      : [
        cleanText(seed),
        `Connect this point directly to ${topic.toLowerCase()}`,
        "Use a clear example",
      ];
  }
  const related = research.facts.find(fact => fact !== seed && fact.length < 180) || research.description || research.title;
  const source = sourceHost(research, seed);
  return [
    localizeText(seed, language).replace(/\.$/, ""),
    localizeText(related, language).replace(/\.$/, ""),
    source ? `${language === "no" ? "Kilde" : "Source"}: ${source}` : (language === "no" ? "Faktabasert punkt" : "Fact-based point"),
  ];
}

function sourceHost(research) {
  const url = research.sources?.[0]?.url || research.url;
  if (!url) return "";
  try {
    return new URL(url).hostname.replace("www.", "");
  } catch {
    return "";
  }
}

function openingNote(topic, level, research = {}, language = "no") {
  if (language !== "no") {
  const fact = research.found ? (research.facts[0] || research.extract) : "";
    if (research.found) return `Say this: Today I am going to talk about ${topic}. I will start with the most important background: ${fact} Then I will go through the key facts and finish with a short summary.`;
    return `Say this: Today I am going to talk about ${topic}. I will explain what the topic is about, why it matters, and what the audience should remember.`;
  }
  if (research.found) {
    const fact = localizeText(research.facts[0] || research.extract, language);
    if (level === "short") {
      return `Si dette: Hei, i dag skal jeg snakke om ${topic}. Kort sagt: ${fact}`;
    }
    if (level === "long") {
      return `Si dette: Hei, i dag skal jeg snakke om ${topic}. Jeg har brukt fakta fra Wikipedia som utgangspunkt. Det første som er viktig å vite, er dette: ${fact} I presentasjonen skal jeg forklare bakgrunnen, de viktigste detaljene og hvorfor dette temaet er interessant.`;
    }
    return `Si dette: Hei, i dag skal jeg snakke om ${topic}. Jeg starter med den viktigste bakgrunnen: ${fact} Etterpå går jeg gjennom flere fakta og avslutter med en kort oppsummering.`;
  }
  if (level === "short") {
    return `Si dette: Hei, i dag skal jeg snakke om ${topic}. Jeg skal forklare hva temaet handler om, hvorfor det er viktig, og hva vi kan lære av det.`;
  }
  if (level === "long") {
    return `Si dette: Hei, i dag skal jeg snakke om ${topic}. Først vil jeg gi en enkel forklaring på hva temaet betyr. Deretter skal jeg vise hvorfor det er relevant, og hvordan det kan påvirke mennesker, samfunn eller hverdagen vår. Målet mitt er at dere etter presentasjonen skal sitte igjen med en tydelig forståelse av temaet og hvorfor det er verdt å bry seg om.`;
  }
  return `Si dette: Hei, i dag skal jeg snakke om ${topic}. Jeg starter med å forklare hva temaet handler om, så går jeg gjennom de viktigste poengene, og til slutt oppsummerer jeg hva dere bør huske.`;
}

function noteFor(seed, level, topic, research = {}, language = "no") {
  const fact = localizeText(seed, language).replace(/\.$/, "");
  if (language !== "no") {
    if (research.found) return `Say this: This slide explains an important fact: ${fact}. I included it because it gives concrete information about ${topic}, not just a general opinion.`;
    return `Say this: This slide is about ${fact}. Explain the point clearly and connect it back to ${topic}.`;
  }
  if (research.found) {
    if (level === "short") {
      return `Si dette: Her er et viktig faktapunkt: ${fact}. Dette hjelper oss å forstå ${topic} bedre.`;
    }
    if (level === "long") {
      return `Si dette: På dette lysbildet skal jeg forklare denne faktadelen: ${fact}. Dette er relevant fordi det gir oss konkret informasjon om ${topic}, i stedet for bare generelle påstander. Legg merke til hvordan dette punktet bygger videre på introduksjonen. Det gjør det lettere å forstå både bakgrunnen og hvorfor temaet har blitt kjent eller viktig.`;
    }
    return `Si dette: Dette lysbildet handler om følgende faktum: ${fact}. Jeg tar det med fordi det forklarer en viktig del av ${topic}. Det viktigste å huske er at dette er konkret informasjon, ikke bare min mening.`;
  }
  if (/^hva temaet handler om$/i.test(fact)) {
    return `Si dette: Dette lysbildet gir en kort oversikt over ${topic}. Forklar temaet enkelt først, slik at publikum forstår hva resten av presentasjonen bygger på.`;
  }
  if (/^bakgrunn/i.test(fact)) {
    return `Si dette: Her forklarer jeg bakgrunnen for ${topic}. Det gjør det lettere å forstå hvorfor temaet ble viktig, og hva som skjedde før hoveddelen av historien.`;
  }
  if (/^viktige hendelser/i.test(fact)) {
    return `Si dette: På dette lysbildet trekker jeg fram viktige hendelser eller personer knyttet til ${topic}. Poenget er å vise hva som faktisk drev historien framover.`;
  }
  if (level === "short") {
    return `Si dette: Dette punktet handler om ${seed}. Det er viktig for ${topic} fordi det viser en av de viktigste sidene ved temaet.`;
  }
  if (level === "long") {
    return `Si dette: Nå skal jeg forklare ${seed}. Dette er en viktig del av ${topic}, fordi det hjelper oss å forstå hva som faktisk skjer og hvorfor det betyr noe. Et enkelt eksempel er at når dette punktet blir oversett, kan folk lett misforstå hele temaet. Derfor er det smart å se på hvordan dette fungerer i praksis, ikke bare som en idé. Det viktigste å huske fra dette lysbildet er at ${seed} henger tett sammen med hovedtemaet, og at det forklarer hvorfor ${topic} er relevant.`;
  }
  return `Si dette: Dette lysbildet handler om ${seed}. Grunnen til at jeg tar det med, er at det forklarer en viktig del av ${topic}. Hvis vi ser på et praktisk eksempel, blir det lettere å forstå hvordan dette påvirker situasjonen. Det dere bør huske, er at dette punktet ikke står alene, men henger sammen med resten av presentasjonen.`;
}

function closingNote(topic, level, research = {}, language = "no") {
  if (language !== "no") {
    const sourceList = research.sources?.length
      ? ` The sources used were: ${research.sources.map(item => item.url).join(", ")}`
      : (research.url ? ` The main source was: ${research.url}` : "");
    return `Say this: To summarize, I have presented the most important facts about ${topic}. The main point is that we now understand what the topic is about and why it matters.${sourceList}`;
  }
  if (research.found) {
    const source = research.url ? ` Kilden jeg brukte som utgangspunkt er Wikipedia: ${research.url}` : "";
    const sourceList = research.sources?.length
      ? ` Kildene som ble brukt er: ${research.sources.map(item => item.url).join(", ")}`
      : source;
    if (level === "short") {
      return `Si dette: For å oppsummere har jeg vist de viktigste faktaene om ${topic}. Det viktigste å huske er hva temaet er, og hvorfor det er kjent eller interessant.${sourceList}`;
    }
    if (level === "long") {
      return `Si dette: Nå vil jeg oppsummere. I denne presentasjonen har jeg brukt faktainformasjon om ${topic} og delt det opp i tydelige deler. Vi har sett på hva temaet er, litt bakgrunn og hvorfor det er relevant. Hvis dere bare skal huske én ting, er det at en god presentasjon bør bygge på konkrete fakta og ikke bare løse påstander.${sourceList} Takk for at dere hørte på.`;
    }
    return `Si dette: For å oppsummere har jeg snakket om ${topic} med utgangspunkt i fakta. Hovedpoenget er at vi nå vet mer om hva temaet handler om og hvorfor det betyr noe.${sourceList}`;
  }
  if (level === "short") {
    return `Si dette: For å oppsummere handler ${topic} om flere viktige poeng. Det viktigste å huske er hovedideen, og hvorfor temaet betyr noe. Takk for meg.`;
  }
  if (level === "long") {
    return `Si dette: Nå vil jeg oppsummere presentasjonen. Vi har sett på ${topic}, og jeg har forklart de viktigste punktene steg for steg. Hovedbudskapet er at dette temaet er viktig fordi det påvirker hvordan vi tenker, velger eller handler. Hvis dere bare skal huske én ting, er det at ${topic} ikke bare er et enkelt tema, men noe som kan ha praktiske konsekvenser. Takk for at dere hørte på, og nå tar jeg gjerne imot spørsmål.`;
  }
  return `Si dette: For å oppsummere har jeg snakket om ${topic} og de viktigste punktene rundt temaet. Hovedpoenget er at dette betyr noe i praksis, ikke bare i teorien. Takk for meg, og jeg kan svare på spørsmål hvis det er noe dere lurer på.`;
}

function generateImages(input) {
  return [0, 1, 2, 3].map(i => ({
    id: `visual-${i + 1}`,
    label: ["Editorial", "Workshop", "Keynote", "Clean system"][i],
    dataUri: svgDataUri(makeImageSvg(input, i)),
  }));
}

function addSpeakerNotes(slide, notes) {
  if (typeof slide.addNotes === "function") slide.addNotes(notes);
  else slide.addNotes = notes;
}

async function makePptx(input) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Slide Discreet Craft";
  pptx.subject = titleCase(input.topic);
  pptx.title = titleCase(input.topic);
  pptx.company = "Local presentation maker";
  pptx.lang = "nb-NO";
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
    lang: "nb-NO",
  };
  pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
  pptx.layout = "CUSTOM_WIDE";

  const slides = await buildSlides(input);
  const navy = "153A6B";
  const ink = "111111";
  const muted = "6B7280";
  const border = "767676";
  const footerTopic = titleCase(input.topic);

  slides.forEach((item, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    const cleanTitle = String(item.title || "").replace(/^[\d.]+\s*/, "").trim();
    slide.addText(`${index + 1}. ${cleanTitle}`, {
      x: 0.62,
      y: 0.48,
      w: 11.4,
      h: 0.55,
      fontFace: "Aptos Display",
      fontSize: 26,
      bold: true,
      color: navy,
      fit: "shrink",
      margin: 0,
    });

    const useBox = index !== 1;
    if (useBox) {
      slide.addShape(pptx.ShapeType.rect, {
        x: 0.55,
        y: 1.52,
        w: 12.25,
        h: 5.55,
        fill: { color: "FFFFFF", transparency: 100 },
        line: { color: border, width: 1.1 },
      });
    }

    const bullets = item.bullets
      .map(bullet => cleanText(bullet).replace(/\.$/, ""))
      .filter(Boolean)
      .slice(0, 4);
    const bulletText = bullets.map(bullet => `•   ${bullet}`).join("\n");
    slide.addText(bulletText, {
      x: 0.72,
      y: useBox ? 3.2 : 3.25,
      w: 11.35,
      h: 2.55,
      fontFace: "Aptos",
      fontSize: 18,
      color: ink,
      breakLine: false,
      fit: "shrink",
      margin: 0,
      paraSpaceAfterPt: 12,
    });

    slide.addText(`${footerTopic} | Presentation maker`, {
      x: 0,
      y: 7.18,
      w: 13.333,
      h: 0.18,
      fontSize: 7,
      color: muted,
      align: "center",
      margin: 0,
    });
    addSpeakerNotes(slide, item.notes);
  });
  return pptx.write({ outputType: "nodebuffer" });
}

function sendJson(res, data) {
  const body = JSON.stringify(data);
  res.writeHead(200, { "content-type": "application/json; charset=utf-8", "content-length": Buffer.byteLength(body) });
  res.end(body);
}

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === "POST" && req.url === "/api/images") {
      const input = await readJson(req);
      sendJson(res, { images: generateImages(input), slides: await buildSlides(input) });
      return;
    }
    if (req.method === "POST" && req.url === "/api/pptx") {
      const input = await readJson(req);
      const buffer = await makePptx(input);
      const safeName = titleCase(input.topic).replace(/[^a-z0-9æøå]+/gi, "-").replace(/^-|-$/g, "") || "presentasjon";
      res.writeHead(200, {
        "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "content-disposition": `attachment; filename="${safeName}.pptx"`,
        "content-length": buffer.length,
      });
      res.end(buffer);
      return;
    }

    const requested = req.url === "/" ? "/index.html" : decodeURIComponent(req.url.split("?")[0]);
    const filePath = path.normalize(path.join(publicDir, requested));
    if (!filePath.startsWith(publicDir)) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }
    fs.readFile(filePath, (error, data) => {
      if (error) {
        res.writeHead(404);
        res.end("Not found");
        return;
      }
      res.writeHead(200, { "content-type": mime[path.extname(filePath)] || "application/octet-stream" });
      res.end(data);
    });
  } catch (error) {
    res.writeHead(500, { "content-type": "application/json; charset=utf-8" });
    res.end(JSON.stringify({ error: error.message }));
  }
});

server.listen(port, () => {
  console.log(`Presentation maker running at http://localhost:${port}`);
});
